/**
 * SGW Chat Two-Way Bridge (`https://sgw.seastar.work` / Cloudflare D1 API Parity)
 *
 * - Authenticates via official `POST /api/auth/login` using employee company credentials.
 * - Persists ONLY the SGW Bearer JWT & user metadata encrypted at rest (AES-256-GCM)
 *   in `<dataDir>/sgw-chat-session.enc` using a local 256-bit key in `<dataDir>/.sgw-chat.key`.
 * - Raw passwords are NEVER stored, cached, or logged; tokens NEVER leave the server side.
 * - Enforces SGW server payload bounds (`content <= 10,000`, `fileB64 <= 7MB` ~ 5MB binary).
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { MailItem } from '../types/mail';

export const SGW_DEFAULT_BASE_URL = 'https://sgw.seastar.work';
export const SGW_MAX_CONTENT_CHARS = 10_000;
export const SGW_MAX_RAW_ATTACHMENT_BYTES = 5 * 1024 * 1024;
export const SGW_MAX_FILE_BASE64_CHARS = 7 * 1024 * 1024;
const REQUEST_TIMEOUT_MS = 15_000;
const EMPLOYEE_CACHE_TTL_MS = 60_000;

export interface SgwChatUser {
  sub: string;
  name: string;
  email: string;
  role: string;
  department: string;
  position: string;
}

export interface SgwEmployeeDirectoryItem {
  id: string;
  name: string;
  email: string;
  role: string;
  department: string;
  position: string;
  status: string;
  phone?: string;
}

export interface SgwChatRoom {
  id: string;
  participants: string[];
  lastMessage: string;
  lastMessageTime: number;
  lastMessageSenderId?: string;
  rev?: number;
  unreadCount?: number;
  participantProfiles?: SgwEmployeeDirectoryItem[];
}

export interface SgwChatMessage {
  id: string;
  roomId: string;
  senderId: string;
  content: string;
  timestamp: number;
  readBy: string[];
  hasFile: boolean;
  fileName?: string;
  fileData?: string;
  revision?: number;
  isEdited?: boolean;
  isDeleted?: boolean;
  replyToId?: string;
  mentions?: string[];
  reactions?: Record<string, string[]>;
}

export interface SgwDirectMessage {
  id: string;
  senderId: string;
  senderName: string;
  receiverId: string;
  receiverName: string;
  content: string;
  timestamp: number;
  isRead: boolean;
  hasFile: boolean;
  fileName?: string;
  fileData?: string;
  revision?: number;
  isEdited?: boolean;
  isDeleted?: boolean;
  replyToId?: string;
  reactions?: Record<string, string[]>;
}

interface StoredSessionPayload {
  sgwBaseUrl: string;
  token?: string;
  cookieHeader?: string;
  user: SgwChatUser;
  savedAt: string;
}

interface EncryptedEnvelope {
  v: 1;
  iv: string;
  tag: string;
  data: string;
}

export class SgwBridgeError extends Error {
  status: number;
  constructor(message: string, status = 500) {
    super(message);
    this.name = 'SgwBridgeError';
    this.status = status;
  }
}

export function normalizeSgwBaseUrl(raw?: string | null): string {
  const candidate = String(raw || '').trim() || SGW_DEFAULT_BASE_URL;
  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    throw new SgwBridgeError('유효한 SGW 서버 주소(URL)가 아닙니다.', 400);
  }
  const isHttps = parsed.protocol === 'https:';
  const isLocalHttp =
    parsed.protocol === 'http:' &&
    (parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost');
  if (!isHttps && !isLocalHttp) {
    throw new SgwBridgeError('SGW 연결은 HTTPS(또는 로컬 테스트 127.0.0.1)만 허용됩니다.', 400);
  }
  return parsed.origin;
}

export function normalizeBase64Payload(raw?: string | null): string | undefined {
  if (!raw) return undefined;
  const trimmed = raw.trim();
  if (!trimmed) return undefined;
  const dataUrlIdx = trimmed.indexOf(';base64,');
  const cleanB64 = dataUrlIdx !== -1 ? trimmed.slice(dataUrlIdx + ';base64,'.length) : trimmed;
  if (cleanB64.length > SGW_MAX_FILE_BASE64_CHARS) {
    throw new SgwBridgeError('첨부파일 크기가 SGW 허용 상한(5MB / Base64 7MB)을 초과했습니다.', 413);
  }
  return cleanB64;
}

export function buildMailShareText(
  mail: Partial<MailItem> & { id: string; subject: string },
  options?: { note?: string; includeDeepLink?: boolean }
): string {
  const senderLabel = mail.fromName
    ? mail.fromEmail
      ? `${mail.fromName} <${mail.fromEmail}>`
      : mail.fromName
    : mail.fromEmail || '발신자 미상';
  const dateLabel = [mail.date, mail.timeLabel].filter(Boolean).join(' ').trim() || mail.timestamp || '';
  const bodyPreviewRaw = String(mail.snippet || mail.body || '')
    .replace(/\r?\n+/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .trim();
  const bodyPreview =
    bodyPreviewRaw.length > 500 ? `${bodyPreviewRaw.slice(0, 497)}...` : bodyPreviewRaw;

  const lines = [
    '[Bmail 메일 공유]',
    `제목: ${String(mail.subject || '(제목 없음)').slice(0, 220)}`,
    `발신: ${senderLabel}`,
  ];
  if (dateLabel) lines.push(`일시: ${dateLabel}`);
  if (mail.attachments && mail.attachments.length > 0) {
    const attNames = mail.attachments
      .slice(0, 4)
      .map((a) => a.filename)
      .join(', ');
    const more = mail.attachments.length > 4 ? ` 외 ${mail.attachments.length - 4}개` : '';
    lines.push(`첨부: ${attNames}${more}`);
  }
  if (bodyPreview) {
    lines.push(`요약: ${bodyPreview}`);
  }
  if (options?.note && options.note.trim()) {
    lines.push(`메모: ${options.note.trim().slice(0, 1200)}`);
  }
  if (options?.includeDeepLink !== false && mail.id) {
    lines.push(`원문 열기: bmail://open?id=${encodeURIComponent(mail.id)}`);
  }

  const text = lines.join('\n');
  return text.length > SGW_MAX_CONTENT_CHARS - 200
    ? `${text.slice(0, SGW_MAX_CONTENT_CHARS - 203)}...`
    : text;
}

export class SgwChatBridge {
  private readonly dataDir: string;
  private readonly keyPath: string;
  private readonly sessionPath: string;
  private readonly defaultBaseUrl: string;
  private cachedSession: StoredSessionPayload | null = null;
  private cachedEmployees: { at: number; items: SgwEmployeeDirectoryItem[] } | null = null;

  constructor(dataDir: string, defaultBaseUrl?: string) {
    this.dataDir = path.resolve(dataDir);
    this.keyPath = path.join(this.dataDir, '.sgw-chat.key');
    this.sessionPath = path.join(this.dataDir, 'sgw-chat-session.enc');
    this.defaultBaseUrl = normalizeSgwBaseUrl(
      defaultBaseUrl || process.env.SGW_BASE_URL || SGW_DEFAULT_BASE_URL
    );
  }

  private getOrCreateKey(): Buffer {
    fs.mkdirSync(this.dataDir, { recursive: true });
    if (fs.existsSync(this.keyPath)) {
      const hex = fs.readFileSync(this.keyPath, 'utf8').trim();
      if (/^[a-f0-9]{64}$/i.test(hex)) {
        return Buffer.from(hex, 'hex');
      }
    }
    const key = crypto.randomBytes(32);
    fs.writeFileSync(this.keyPath, key.toString('hex'), { encoding: 'utf8', mode: 0o600 });
    return key;
  }

  private saveSession(payload: StoredSessionPayload): void {
    const key = this.getOrCreateKey();
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
    const encrypted = Buffer.concat([
      cipher.update(JSON.stringify(payload), 'utf8'),
      cipher.final(),
    ]);
    const tag = cipher.getAuthTag();
    const env: EncryptedEnvelope = {
      v: 1,
      iv: iv.toString('base64'),
      tag: tag.toString('base64'),
      data: encrypted.toString('base64'),
    };
    fs.writeFileSync(this.sessionPath, JSON.stringify(env), { encoding: 'utf8', mode: 0o600 });
    this.cachedSession = payload;
  }

  private loadSession(): StoredSessionPayload | null {
    if (this.cachedSession) return this.cachedSession;
    if (!fs.existsSync(this.sessionPath) || !fs.existsSync(this.keyPath)) {
      return null;
    }
    try {
      const key = this.getOrCreateKey();
      const raw = JSON.parse(fs.readFileSync(this.sessionPath, 'utf8')) as EncryptedEnvelope;
      if (!raw || raw.v !== 1 || !raw.iv || !raw.tag || !raw.data) return null;
      const decipher = crypto.createDecipheriv(
        'aes-256-gcm',
        key,
        Buffer.from(raw.iv, 'base64')
      );
      decipher.setAuthTag(Buffer.from(raw.tag, 'base64'));
      const plain = Buffer.concat([
        decipher.update(Buffer.from(raw.data, 'base64')),
        decipher.final(),
      ]).toString('utf8');
      const parsed = JSON.parse(plain) as StoredSessionPayload;
      if (!parsed || (!parsed.token && !parsed.cookieHeader) || !parsed.user?.sub) return null;
      this.cachedSession = {
        ...parsed,
        sgwBaseUrl: normalizeSgwBaseUrl(parsed.sgwBaseUrl),
      };
      return this.cachedSession;
    } catch {
      return null;
    }
  }

  clearSession(): void {
    this.cachedSession = null;
    this.cachedEmployees = null;
    try {
      if (fs.existsSync(this.sessionPath)) fs.unlinkSync(this.sessionPath);
    } catch {
      // ignore
    }
  }

  getStatus(): {
    connected: boolean;
    sgwBaseUrl: string;
    user: SgwChatUser | null;
    savedAt?: string;
  } {
    const sess = this.loadSession();
    if (!sess) {
      return {
        connected: false,
        sgwBaseUrl: this.defaultBaseUrl,
        user: null,
      };
    }
    return {
      connected: true,
      sgwBaseUrl: sess.sgwBaseUrl,
      user: sess.user,
      savedAt: sess.savedAt,
    };
  }

  private async requestJson<T>(
    method: 'GET' | 'POST',
    endpointPath: string,
    options?: {
      body?: unknown;
      requireAuth?: boolean;
      baseUrlOverride?: string;
      extraHeaders?: Record<string, string>;
    }
  ): Promise<T> {
    const requireAuth = options?.requireAuth !== false;
    const sess = this.loadSession();
    if (requireAuth && !sess) {
      throw new SgwBridgeError('SGW 채팅 계정에 먼저 로그인해 주세요.', 401);
    }
    const baseUrl = normalizeSgwBaseUrl(options?.baseUrlOverride || sess?.sgwBaseUrl || this.defaultBaseUrl);
    const headers: Record<string, string> = {
      Accept: 'application/json',
      ...(options?.extraHeaders || {}),
    };
    if (options?.body !== undefined) {
      headers['Content-Type'] = 'application/json; charset=utf-8';
    }
    if (requireAuth && sess) {
      if (sess.token && !headers['Authorization']) {
        headers['Authorization'] = `Bearer ${sess.token}`;
      }
      if (sess.cookieHeader && !headers['Cookie']) {
        headers['Cookie'] = sess.cookieHeader;
      }
    }

    let res: Response;
    try {
      res = await fetch(`${baseUrl}${endpointPath}`, {
        method,
        headers,
        body: options?.body !== undefined ? JSON.stringify(options.body) : undefined,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '네트워크 오류';
      throw new SgwBridgeError(`SGW 서버(${baseUrl}) 연결 실패: ${msg}`, 502);
    }

    let data: Record<string, unknown> = {};
    try {
      data = (await res.json()) as Record<string, unknown>;
    } catch {
      data = {};
    }

    if (res.status === 401 && requireAuth) {
      this.clearSession();
      throw new SgwBridgeError('SGW 로그인 세션이 만료되었습니다. 다시 로그인해 주세요.', 401);
    }
    if (!res.ok || data.ok === false) {
      const errCode = typeof data.error === 'string' ? data.error : `HTTP ${res.status}`;
      const friendly =
        errCode === 'invalid_credentials'
          ? '사원 이메일(또는 아이디)과 비밀번호가 일치하지 않습니다.'
          : errCode === 'rate_limited'
            ? '로그인 시도가 많아 잠시 제한되었습니다. 1분 후 다시 시도해 주세요.'
            : `SGW API 요청 실패 (${errCode})`;
      throw new SgwBridgeError(friendly, res.status || 400);
    }

    return data as T;
  }

  async login(params: {
    email: string;
    password: string;
    sgwBaseUrl?: string;
  }): Promise<{
    connected: boolean;
    sgwBaseUrl: string;
    user: SgwChatUser | null;
    savedAt?: string;
  }> {
    const email = String(params.email || '').trim();
    const password = String(params.password || '');
    if (!email || !password) {
      throw new SgwBridgeError('이메일(또는 사번/이름)과 비밀번호를 입력해 주세요.', 400);
    }
    const baseUrl = normalizeSgwBaseUrl(params.sgwBaseUrl || this.defaultBaseUrl);
    const data = await this.requestJson<{
      ok?: boolean;
      token?: string;
      user?: Record<string, unknown>;
    }>('POST', '/api/auth/login', {
      requireAuth: false,
      baseUrlOverride: baseUrl,
      body: { email, password },
    });

    if (!data.token || !data.user) {
      throw new SgwBridgeError('SGW 서버가 유효한 세션 토큰을 반환하지 않았습니다.', 502);
    }

    const u = data.user;
    const normalizedUser: SgwChatUser = {
      sub: String(u.sub ?? u.id ?? ''),
      name: String(u.name ?? email),
      email: String(u.email ?? email),
      role: String(u.role ?? 'user'),
      department: String(u.department ?? ''),
      position: String(u.position ?? ''),
    };
    if (!normalizedUser.sub) {
      throw new SgwBridgeError('SGW 사용자 식별자(sub)가 비어 있습니다.', 502);
    }

    this.saveSession({
      sgwBaseUrl: baseUrl,
      token: String(data.token),
      user: normalizedUser,
      savedAt: new Date().toISOString(),
    });
    this.cachedEmployees = null;
    return this.getStatus();
  }

  async connectWithSession(params: {
    sgwBaseUrl?: string;
    token?: string;
    cookieHeader?: string;
    user?: Partial<SgwChatUser>;
  }): Promise<{
    connected: boolean;
    sgwBaseUrl: string;
    user: SgwChatUser | null;
    savedAt?: string;
  }> {
    const token = String(params.token || '').trim().replace(/^Bearer\s+/i, '');
    const cookieHeader = String(params.cookieHeader || '').trim();
    if (!token && !cookieHeader) {
      throw new SgwBridgeError('SGW 인증 세션 토큰 또는 쿠키 정보가 없습니다.', 400);
    }

    const baseUrl = normalizeSgwBaseUrl(params.sgwBaseUrl || this.defaultBaseUrl);
    const extraHeaders: Record<string, string> = {};
    if (token) extraHeaders['Authorization'] = `Bearer ${token}`;
    if (cookieHeader) extraHeaders['Cookie'] = cookieHeader;

    const me = await this.requestJson<Record<string, unknown>>('GET', '/api/auth/me', {
      requireAuth: false,
      baseUrlOverride: baseUrl,
      extraHeaders,
    });

    const rawUser =
      me.user && typeof me.user === 'object' && !Array.isArray(me.user)
        ? (me.user as Record<string, unknown>)
        : me;

    const normalizedUser: SgwChatUser = {
      sub: String(rawUser.sub ?? rawUser.id ?? params.user?.sub ?? ''),
      name: String(rawUser.name ?? params.user?.name ?? rawUser.email ?? '사원'),
      email: String(rawUser.email ?? params.user?.email ?? ''),
      role: String(rawUser.role ?? params.user?.role ?? 'user'),
      department: String(rawUser.department ?? params.user?.department ?? ''),
      position: String(rawUser.position ?? params.user?.position ?? ''),
    };

    if (!normalizedUser.sub) {
      throw new SgwBridgeError('SGW /api/auth/me 응답에 사용자 식별자(sub)가 없습니다.', 401);
    }

    this.saveSession({
      sgwBaseUrl: baseUrl,
      token: token || undefined,
      cookieHeader: cookieHeader || undefined,
      user: normalizedUser,
      savedAt: new Date().toISOString(),
    });
    this.cachedEmployees = null;
    return this.getStatus();
  }

  async listEmployees(forceRefresh = false): Promise<SgwEmployeeDirectoryItem[]> {
    if (
      !forceRefresh &&
      this.cachedEmployees &&
      Date.now() - this.cachedEmployees.at < EMPLOYEE_CACHE_TTL_MS
    ) {
      return this.cachedEmployees.items;
    }
    const data = await this.requestJson<{ employees?: SgwEmployeeDirectoryItem[] }>(
      'GET',
      '/api/employees'
    );
    const items = Array.isArray(data.employees) ? data.employees : [];
    this.cachedEmployees = { at: Date.now(), items };
    return items;
  }

  async listRooms(): Promise<{ rooms: SgwChatRoom[]; unreadTotal: number }> {
    const sess = this.loadSession();
    if (!sess) {
      throw new SgwBridgeError('SGW 채팅 계정에 먼저 로그인해 주세요.', 401);
    }
    const [chatData, employees] = await Promise.all([
      this.requestJson<{ rooms?: SgwChatRoom[] }>(
        'GET',
        `/api/chat?userId=${encodeURIComponent(sess.user.sub)}`
      ),
      this.listEmployees().catch(() => [] as SgwEmployeeDirectoryItem[]),
    ]);
    const empById = new Map(employees.map((e) => [String(e.id), e]));
    const rooms = (Array.isArray(chatData.rooms) ? chatData.rooms : []).map((room) => ({
      ...room,
      participantProfiles: (room.participants || [])
        .map((pid) => empById.get(String(pid)))
        .filter((p): p is SgwEmployeeDirectoryItem => Boolean(p)),
    }));
    const unreadTotal = rooms.reduce((sum, r) => sum + Number(r.unreadCount || 0), 0);
    return { rooms, unreadTotal };
  }

  async getRoomMessages(
    roomId: string,
    msgId?: string
  ): Promise<{ messages: SgwChatMessage[]; notices?: unknown[] }> {
    if (!roomId.trim()) {
      throw new SgwBridgeError('roomId가 필요합니다.', 400);
    }
    const qs = new URLSearchParams({ roomId: roomId.trim() });
    if (msgId && msgId.trim()) qs.set('msgId', msgId.trim());
    const data = await this.requestJson<{ messages?: SgwChatMessage[]; notices?: unknown[] }>(
      'GET',
      `/api/chat?${qs.toString()}`
    );
    return {
      messages: Array.isArray(data.messages) ? data.messages : [],
      ...(Array.isArray(data.notices) ? { notices: data.notices } : {}),
    };
  }

  async createOrOpenRoom(params: {
    participants: string[];
    roomId?: string;
    lastMessage?: string;
  }): Promise<{ ok: boolean; id: string }> {
    const sess = this.loadSession();
    if (!sess) {
      throw new SgwBridgeError('SGW 채팅 계정에 먼저 로그인해 주세요.', 401);
    }
    const merged = Array.from(
      new Set([sess.user.sub, ...(params.participants || []).map(String).filter(Boolean)])
    );
    if (merged.length < 1) {
      throw new SgwBridgeError('최소 1명 이상의 참여자가 필요합니다.', 400);
    }
    return this.requestJson<{ ok: boolean; id: string }>('POST', '/api/chat?action=room', {
      body: {
        ...(params.roomId ? { id: params.roomId } : {}),
        participants: merged,
        ...(params.lastMessage ? { lastMessage: params.lastMessage.slice(0, SGW_MAX_CONTENT_CHARS) } : {}),
      },
    });
  }

  async sendRoomMessage(
    roomId: string,
    payload: {
      content: string;
      fileName?: string;
      fileData?: string;
      replyToId?: string;
      mentions?: string[];
    }
  ): Promise<{ ok: boolean; id: string }> {
    const cleanRoomId = String(roomId || '').trim();
    if (!cleanRoomId) {
      throw new SgwBridgeError('대화방 ID(roomId)가 필요합니다.', 400);
    }
    const content = String(payload.content || '').trim();
    const cleanB64 = normalizeBase64Payload(payload.fileData);
    if (!content && !cleanB64) {
      throw new SgwBridgeError('메시지 내용 또는 첨부파일이 필요합니다.', 400);
    }
    if (content.length > SGW_MAX_CONTENT_CHARS) {
      throw new SgwBridgeError(`메시지는 최대 ${SGW_MAX_CONTENT_CHARS.toLocaleString()}자까지 전송 가능합니다.`, 400);
    }
    const clientMsgId = `bm_chat_${crypto.randomUUID()}`;
    return this.requestJson<{ ok: boolean; id: string }>('POST', '/api/chat', {
      body: {
        id: clientMsgId,
        roomId: cleanRoomId,
        content,
        ...(payload.fileName ? { fileName: String(payload.fileName).slice(0, 240) } : {}),
        ...(cleanB64 ? { fileData: cleanB64 } : {}),
        ...(payload.replyToId ? { replyToId: String(payload.replyToId) } : {}),
        ...(Array.isArray(payload.mentions) && payload.mentions.length > 0
          ? { mentions: payload.mentions.map(String) }
          : {}),
      },
    });
  }

  async markRoomRead(roomId: string, lastReadMessageId: string): Promise<{ ok: boolean }> {
    return this.requestJson<{ ok: boolean }>('POST', '/api/chat?action=read', {
      body: { roomId, lastReadMessageId },
    });
  }

  async toggleRoomReaction(
    roomId: string,
    id: string,
    emoji: string,
    on = true
  ): Promise<{ ok: boolean; id: string }> {
    return this.requestJson<{ ok: boolean; id: string }>('POST', '/api/chat?action=react', {
      body: { id, roomId, emoji, on },
    });
  }

  async listDirectMessages(options?: {
    since?: number;
    id?: string;
  }): Promise<{ messages: SgwDirectMessage[]; serverNow?: number }> {
    const qs = new URLSearchParams();
    if (options?.id) qs.set('id', options.id);
    else if (typeof options?.since === 'number' && Number.isFinite(options.since)) {
      qs.set('since', String(options.since));
    }
    const query = qs.toString();
    const data = await this.requestJson<{ messages?: SgwDirectMessage[]; serverNow?: number }>(
      'GET',
      `/api/messages${query ? `?${query}` : ''}`
    );
    return {
      messages: Array.isArray(data.messages) ? data.messages : [],
      serverNow: data.serverNow,
    };
  }

  async sendDirectMessage(payload: {
    receiverId: string;
    receiverName?: string;
    content: string;
    fileName?: string;
    fileData?: string;
    replyToId?: string;
  }): Promise<{ ok: boolean; id: string }> {
    const receiverId = String(payload.receiverId || '').trim();
    if (!receiverId) {
      throw new SgwBridgeError('수신자 사원 ID(receiverId)가 필요합니다.', 400);
    }
    const content = String(payload.content || '').trim();
    const cleanB64 = normalizeBase64Payload(payload.fileData);
    if (!content && !cleanB64) {
      throw new SgwBridgeError('쪽지 내용 또는 첨부파일이 필요합니다.', 400);
    }
    if (content.length > SGW_MAX_CONTENT_CHARS) {
      throw new SgwBridgeError(`쪽지는 최대 ${SGW_MAX_CONTENT_CHARS.toLocaleString()}자까지 전송 가능합니다.`, 400);
    }
    const msgId = `bm_dm_${crypto.randomUUID()}`;
    return this.requestJson<{ ok: boolean; id: string }>('POST', '/api/messages', {
      body: {
        id: msgId,
        receiverId,
        receiverName: payload.receiverName ? String(payload.receiverName).slice(0, 100) : '',
        content,
        ...(payload.fileName ? { fileName: String(payload.fileName).slice(0, 240) } : {}),
        ...(cleanB64 ? { fileData: cleanB64 } : {}),
        ...(payload.replyToId ? { replyToId: String(payload.replyToId) } : {}),
      },
    });
  }

  async markDmRead(otherId: string): Promise<{ ok: boolean }> {
    if (!otherId.trim()) {
      throw new SgwBridgeError('otherId가 필요합니다.', 400);
    }
    return this.requestJson<{ ok: boolean }>(
      'POST',
      `/api/messages?action=read&otherId=${encodeURIComponent(otherId.trim())}`,
      { body: {} }
    );
  }

  async getLedgerDashboard(filters?: {
    project?: string;
    direction?: string;
    q?: string;
  }): Promise<Record<string, unknown>> {
    const qs = new URLSearchParams();
    if (filters?.project) qs.set('project', filters.project);
    if (filters?.direction) qs.set('direction', filters.direction);
    if (filters?.q) qs.set('q', filters.q);
    const query = qs.toString();
    return this.requestJson<Record<string, unknown>>(
      'GET',
      `/api/bmail-ledger${query ? `?${query}` : ''}`
    );
  }

  async syncEmailsToLedger(
    emails: MailItem[],
    options?: { sourceSystem?: string }
  ): Promise<{ ok: boolean; inserted: number; touchedDates: string[] }> {
    if (!Array.isArray(emails) || emails.length === 0) {
      throw new SgwBridgeError('SGW 원장으로 동기화할 메일이 없습니다.', 400);
    }
    const batch = emails.slice(0, 100).map((m) => {
      const mailDate = String(m.date || m.timestamp || new Date().toISOString())
        .slice(0, 10);
      const direction = m.folder === 'sent' ? 'sent' : 'recv';
      const sender = m.fromEmail
        ? m.fromName
          ? `${m.fromName} <${m.fromEmail}>`
          : m.fromEmail
        : m.fromName || 'unknown@seastar.work';
      const recipients = Array.isArray(m.to) && m.to.length > 0 ? m.to.join(', ') : '';
      return {
        bmail_uid: m.id,
        direction,
        mail_date: mailDate,
        mail_timestamp: m.timestamp || `${mailDate}T09:00:00+09:00`,
        sender,
        recipients,
        subject: String(m.subject || '(제목 없음)').slice(0, 300),
        body_snippet: String(m.snippet || m.body || '').slice(0, 600),
        attachments: (m.attachments || []).map((a) => ({
          filename: a.filename,
          size_bytes: Number(a.sizeBytes || 0),
        })),
        action_status: direction === 'recv' && !m.isRead ? 'need_reply' : 'done',
        source_system: options?.sourceSystem || 'BMAIL_DESKTOP_V12',
      };
    });

    const data = await this.requestJson<{
      ok?: boolean;
      inserted?: number;
      touchedDates?: string[];
    }>('POST', '/api/bmail-ledger', {
      body: { action: 'ingest', items: batch },
    });
    return {
      ok: Boolean(data.ok),
      inserted: Number(data.inserted || 0),
      touchedDates: Array.isArray(data.touchedDates) ? data.touchedDates : [],
    };
  }

  async pushDailyLogToWeekly(logDate?: string): Promise<Record<string, unknown>> {
    const cleanDate =
      (logDate && logDate.trim().slice(0, 10)) || new Date().toISOString().slice(0, 10);
    return this.requestJson<Record<string, unknown>>('POST', '/api/bmail-ledger', {
      body: { action: 'push-to-weekly', log_date: cleanDate },
    });
  }
}
