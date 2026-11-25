import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';

export type EmvScope = 'global' | 'employee';

export const ALLOWED_EMV_KEYS = [
  'GEMINI_API_KEY',
  'GEMINI_MODEL',
  'GROQ_API_KEY',
  'GROQ_API_KEY_2',
  'GROQ_API_KEY_3',
  'GROQ_API_KEY_4',
  'GROQ_API_KEY_5',
  'ANTHROPIC_API_KEY',
  'OPENAI_API_KEY',
  'AGENT_BUILDER_WORKER_URL',
  'AGENT_BUILDER_WORKER_TOKEN',
  'MICROSOFT_CLIENT_ID',
  'MICROSOFT_CLIENT_SECRET',
] as const;

export type EmvKeyName = (typeof ALLOWED_EMV_KEYS)[number];

export interface EmvKeyDefinition {
  key: EmvKeyName;
  label: string;
  category: 'ai-primary' | 'ai-rag' | 'ai-fallback' | 'oauth';
  sensitive: boolean;
  description: string;
  placeholder: string;
}

export const EMV_KEY_CATALOG: EmvKeyDefinition[] = [
  {
    key: 'GEMINI_API_KEY',
    label: 'Gemini 4 Hydrogen API 키 (1순위)',
    category: 'ai-primary',
    sensitive: true,
    description: '실시간 메일 요약·업무일지·블로그 생성 1순위 엔진 키 (13.43만 원 크레딧)',
    placeholder: 'AIza... (Google AI Studio 유료 키)',
  },
  {
    key: 'GEMINI_MODEL',
    label: 'Gemini 기본 모델명',
    category: 'ai-primary',
    sensitive: false,
    description: '미설정 시 gemini-4-hydrogen 자동 적용',
    placeholder: 'gemini-4-hydrogen',
  },
  {
    key: 'AGENT_BUILDER_WORKER_URL',
    label: 'Agent Builder / Cloudflare Worker URL',
    category: 'ai-rag',
    sensitive: false,
    description: '과거 누적 메일 RAG 검색용 중앙 Worker 엔드포인트 (https://...)',
    placeholder: 'https://bmail-ai-gateway.company.workers.dev',
  },
  {
    key: 'AGENT_BUILDER_WORKER_TOKEN',
    label: 'Agent Builder Worker 인증 토큰',
    category: 'ai-rag',
    sensitive: true,
    description: 'Cloudflare Worker 내부 호출 보호용 Bearer 토큰',
    placeholder: 'worker-shared-token',
  },
  {
    key: 'GROQ_API_KEY',
    label: 'Groq API 키 #1 (기본 5-키 순환)',
    category: 'ai-primary',
    sensitive: true,
    description: '기본 LLM(openai/gpt-oss-120b) 5-키 라운드로빈 1번 슬롯 및 Whisper STT 키',
    placeholder: 'gsk_...',
  },
  {
    key: 'GROQ_API_KEY_2',
    label: 'Groq API 키 #2 (순환 풀)',
    category: 'ai-primary',
    sensitive: true,
    description: '기본 LLM 5-키 라운드로빈 2번 슬롯',
    placeholder: 'gsk_...',
  },
  {
    key: 'GROQ_API_KEY_3',
    label: 'Groq API 키 #3 (순환 풀)',
    category: 'ai-primary',
    sensitive: true,
    description: '기본 LLM 5-키 라운드로빈 3번 슬롯',
    placeholder: 'gsk_...',
  },
  {
    key: 'GROQ_API_KEY_4',
    label: 'Groq API 키 #4 (순환 풀)',
    category: 'ai-primary',
    sensitive: true,
    description: '기본 LLM 5-키 라운드로빈 4번 슬롯',
    placeholder: 'gsk_...',
  },
  {
    key: 'GROQ_API_KEY_5',
    label: 'Groq API 키 #5 (순환 풀)',
    category: 'ai-primary',
    sensitive: true,
    description: '기본 LLM 5-키 라운드로빈 5번 슬롯',
    placeholder: 'gsk_...',
  },
  {
    key: 'ANTHROPIC_API_KEY',
    label: 'Claude API 키 (고성능·장애 에스컬레이션)',
    category: 'ai-fallback',
    sensitive: true,
    description: 'Groq 장애 시 자동 폴백 및 고성능 정밀 분석·감리 에스컬레이션용 Anthropic 키',
    placeholder: 'sk-ant-...',
  },
  {
    key: 'OPENAI_API_KEY',
    label: 'OpenAI ChatGPT API 키 (2순위 폴백)',
    category: 'ai-fallback',
    sensitive: true,
    description: '문서·스레드 분석 및 음성 전사 폴백용 OpenAI 키',
    placeholder: 'sk-...',
  },
  {
    key: 'MICROSOFT_CLIENT_ID',
    label: 'Microsoft Entra ID Client ID',
    category: 'oauth',
    sensitive: false,
    description: 'Microsoft 365 Outlook Graph OAuth 앱 클라이언트 ID',
    placeholder: '00000000-0000-0000-0000-000000000000',
  },
  {
    key: 'MICROSOFT_CLIENT_SECRET',
    label: 'Microsoft Entra ID Client Secret',
    category: 'oauth',
    sensitive: true,
    description: 'Microsoft 365 Outlook Graph OAuth 클라이언트 시크릿',
    placeholder: 'client-secret-value',
  },
];

const ALLOWED_EMV_KEY_SET = new Set<string>(ALLOWED_EMV_KEYS);
const SENSITIVE_KEY_SET = new Set<string>(
  EMV_KEY_CATALOG.filter((item) => item.sensitive).map((item) => item.key)
);

interface EncryptedBlob {
  iv: string;
  tag: string;
  ciphertext: string;
}

interface StoredEmvRecord {
  key: EmvKeyName;
  scope: EmvScope;
  employeeId?: string;
  encrypted: EncryptedBlob;
  maskedPreview: string;
  fingerprint: string;
  updatedAt: string;
  updatedBy: string;
}

interface AdminClaimHash {
  salt: string;
  hash: string;
  claimedAt: string;
}

interface VaultFileEnvelope {
  version: 1;
  salt: string;
  adminClaim?: AdminClaimHash;
  globalEntries: Record<string, StoredEmvRecord>;
  employeeEntries: Record<string, { employeeId: string; entries: Record<string, StoredEmvRecord> }>;
}

export interface PublicEmvEntryView {
  key: EmvKeyName;
  scope: EmvScope;
  employeeId?: string;
  maskedPreview: string;
  fingerprint: string;
  updatedAt: string;
  updatedBy: string;
}

export interface EmvResolutionSourceItem {
  key: EmvKeyName;
  configured: boolean;
  source: 'employee-emv' | 'global-emv' | 'server-env' | 'unconfigured';
  maskedPreview: string;
}

export interface EmvAdminSession {
  token: string;
  createdAt: number;
  expiresAt: number;
  label: string;
}

const ADMIN_SESSION_TTL_MS = 2 * 60 * 60 * 1000; // 2 hours
const MAX_EMPLOYEES = 500;
const MAX_VALUE_LENGTH = 2048;

export function isAllowedEmvKey(key: string): key is EmvKeyName {
  return ALLOWED_EMV_KEY_SET.has(key);
}

/**
 * Normalizes and validates an employee identifier (corporate email or safe employee ID).
 * Blocks path traversal, prototype pollution names, control chars, and whitespace.
 */
export function normalizeEmployeeId(raw?: string | null): string {
  const trimmed = String(raw || '').trim().toLowerCase();
  if (!trimmed || trimmed.length > 160) return '';
  if (/[\r\n\0\s/\\]/.test(trimmed)) return '';
  if (trimmed === '__proto__' || trimmed === 'constructor' || trimmed === 'prototype') return '';
  const isEmail = /^[a-z0-9._+-]+@[a-z0-9.-]+\.[a-z]{2,}$/i.test(trimmed);
  const isStaffCode = /^[a-z0-9][a-z0-9._-]{1,63}$/i.test(trimmed);
  return isEmail || isStaffCode ? trimmed : '';
}

/**
 * Produces a safe masked display preview for UI display (never leaks full secret values).
 */
export function maskEmvValue(key: EmvKeyName, rawValue: string): string {
  const val = rawValue.trim();
  if (!val) return '';
  if (!SENSITIVE_KEY_SET.has(key)) {
    return val.length > 64 ? `${val.slice(0, 61)}...` : val;
  }
  if (val.length <= 8) return '••••••••';
  return `${val.slice(0, 4)}••••${val.slice(-4)}`;
}

export function computeSecretFingerprint(rawValue: string): string {
  return crypto.createHash('sha256').update(rawValue.trim(), 'utf8').digest('hex').slice(0, 12);
}

function hashPasscodeScrypt(passcode: string, saltHex: string): string {
  return crypto.scryptSync(passcode, saltHex, 32).toString('hex');
}

function timingSafeStringEqual(a: string, b: string): boolean {
  const bufA = Buffer.from(a, 'utf8');
  const bufB = Buffer.from(b, 'utf8');
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

export interface EmvVaultOptions {
  /** Base directory for confined vault file storage (defaults to <cwd>/data) */
  baseDir?: string;
  /** Optional filename inside baseDir (defaults to emv-vault.enc.json) */
  vaultFileName?: string;
  /** Server environment source (defaults to process.env) */
  env?: NodeJS.ProcessEnv;
}

/**
 * Server-side encrypted EMV (Environment Variable) vault supporting:
 * - Org-wide (global) API keys shared across all staff installations
 * - Per-employee isolated API keys keyed by HMAC-SHA256 of verified employee identity
 * - AES-256-GCM encryption at rest when EMV_MASTER_KEY or CONFIG_ADMIN_PASSWORD is configured
 * - Zero plaintext secret writes to disk; ephemeral encrypted memory fallback if no master env key is set
 * - Atomic first-run super-admin passcode claim when CONFIG_ADMIN_PASSWORD is not configured in .env
 */
export class EmvVaultService {
  private readonly baseDir: string;
  private readonly vaultFilePath: string;
  private readonly envSource: NodeJS.ProcessEnv;
  private readonly ephemeralMasterKey: Buffer;
  private readonly adminSessions = new Map<string, EmvAdminSession>();
  private state: VaultFileEnvelope;
  private readonly persistenceEnabled: boolean;

  constructor(options: EmvVaultOptions = {}) {
    this.envSource = options.env ?? process.env;
    this.baseDir = path.resolve(options.baseDir ?? path.resolve(process.cwd(), 'data'));
    const safeFileName = path.basename(options.vaultFileName || 'emv-vault.enc.json');
    const resolvedPath = path.resolve(this.baseDir, safeFileName);
    if (!resolvedPath.startsWith(this.baseDir + path.sep)) {
      throw new Error('유효하지 않은 EMV 볼트 저장 경로입니다.');
    }
    this.vaultFilePath = resolvedPath;
    this.ephemeralMasterKey = crypto.randomBytes(32);
    this.persistenceEnabled = Boolean(this.getConfiguredMasterSecret());

    this.state = this.loadOrInitEnvelope();
  }

  private getConfiguredMasterSecret(): string {
    const master = (this.envSource.EMV_MASTER_KEY || '').trim();
    if (master.length >= 8) return master;
    const adminPw = (this.envSource.CONFIG_ADMIN_PASSWORD || '').trim();
    if (adminPw.length >= 6) return adminPw;
    return '';
  }

  private deriveAesKey(saltHex: string): Buffer {
    const configuredSecret = this.getConfiguredMasterSecret();
    if (!configuredSecret) {
      return this.ephemeralMasterKey;
    }
    return crypto.scryptSync(configuredSecret, `bmail-emv-v1:${saltHex}`, 32);
  }

  private computeEmployeePartitionKey(normalizedEmployeeId: string): string {
    const aesKey = this.deriveAesKey(this.state.salt);
    return crypto
      .createHmac('sha256', aesKey)
      .update(`emp:${normalizedEmployeeId}`, 'utf8')
      .digest('hex');
  }

  private encryptValue(plaintext: string): EncryptedBlob {
    const key = this.deriveAesKey(this.state.salt);
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
    const encrypted = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
    const tag = cipher.getAuthTag();
    return {
      iv: iv.toString('base64'),
      tag: tag.toString('base64'),
      ciphertext: encrypted.toString('base64'),
    };
  }

  private decryptValue(blob: EncryptedBlob): string | null {
    try {
      const key = this.deriveAesKey(this.state.salt);
      const iv = Buffer.from(blob.iv, 'base64');
      const tag = Buffer.from(blob.tag, 'base64');
      const ciphertext = Buffer.from(blob.ciphertext, 'base64');
      const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
      decipher.setAuthTag(tag);
      const decrypted = Buffer.concat([decipher.update(ciphertext), decipher.final()]);
      return decrypted.toString('utf8');
    } catch {
      return null;
    }
  }

  private loadOrInitEnvelope(): VaultFileEnvelope {
    if (this.persistenceEnabled && fs.existsSync(this.vaultFilePath)) {
      try {
        const raw = JSON.parse(fs.readFileSync(this.vaultFilePath, 'utf8')) as Partial<VaultFileEnvelope>;
        if (raw && raw.version === 1 && typeof raw.salt === 'string') {
          return {
            version: 1,
            salt: raw.salt,
            adminClaim: raw.adminClaim,
            globalEntries:
              raw.globalEntries && typeof raw.globalEntries === 'object' ? raw.globalEntries : {},
            employeeEntries:
              raw.employeeEntries && typeof raw.employeeEntries === 'object' ? raw.employeeEntries : {},
          };
        }
      } catch {
        // Fall through to clean initialized envelope if vault file is corrupted
      }
    }
    return {
      version: 1,
      salt: crypto.randomBytes(16).toString('hex'),
      globalEntries: {},
      employeeEntries: {},
    };
  }

  private persistEnvelope(): void {
    if (!this.persistenceEnabled) return;
    if (!fs.existsSync(this.baseDir)) {
      fs.mkdirSync(this.baseDir, { recursive: true, mode: 0o700 });
    }
    const tmpPath = `${this.vaultFilePath}.${crypto.randomBytes(6).toString('hex')}.tmp`;
    if (!tmpPath.startsWith(this.baseDir + path.sep)) {
      throw new Error('볼트 임시 파일 경로 검증에 실패했습니다.');
    }
    fs.writeFileSync(tmpPath, `${JSON.stringify(this.state, null, 2)}\n`, {
      encoding: 'utf8',
      mode: 0o600,
    });
    fs.renameSync(tmpPath, this.vaultFilePath);
  }

  // --- Super-Admin Auth (Env Passcode > Atomic First-Run Claimed Scrypt Hash) ---

  getAuthStatus(bearerToken?: string | null) {
    const envPasscodeSet = Boolean((this.envSource.CONFIG_ADMIN_PASSWORD || '').trim());
    const claimed = Boolean(this.state.adminClaim);
    const session = this.verifyAdminSession(bearerToken);
    return {
      envAdminPasswordConfigured: envPasscodeSet,
      adminClaimed: claimed,
      requiresInitialClaim: !envPasscodeSet && !claimed,
      authenticated: Boolean(session),
      persistenceMode: this.persistenceEnabled ? ('encrypted-disk' as const) : ('encrypted-memory' as const),
      sessionExpiresAt: session?.expiresAt ?? null,
    };
  }

  /**
   * One-time atomic admin passcode bootstrap when CONFIG_ADMIN_PASSWORD is unset in .env.
   * Locks itself immediately after claiming.
   */
  claimInitialAdminPassword(passcode: string): EmvAdminSession {
    if ((this.envSource.CONFIG_ADMIN_PASSWORD || '').trim()) {
      throw new Error('서버 .env에 CONFIG_ADMIN_PASSWORD가 이미 설정되어 있어 초기 등록이 잠겨 있습니다.');
    }
    if (this.state.adminClaim) {
      throw new Error('종합 관리자 비밀번호가 이미 등록되어 있습니다. 기존 비밀번호로 로그인해 주세요.');
    }
    const clean = String(passcode || '').trim();
    if (clean.length < 8 || clean.length > 128) {
      throw new Error('종합 관리자 비밀번호는 8자 이상 128자 이하로 설정해 주세요.');
    }
    const salt = crypto.randomBytes(16).toString('hex');
    this.state.adminClaim = {
      salt,
      hash: hashPasscodeScrypt(clean, salt),
      claimedAt: new Date().toISOString(),
    };
    this.persistEnvelope();
    return this.createAdminSession('초기 등록 종합 관리자');
  }

  authenticateAdmin(passcode: string): EmvAdminSession {
    const clean = String(passcode || '').trim();
    if (!clean) {
      throw new Error('관리자 비밀번호를 입력해 주세요.');
    }
    const envPw = (this.envSource.CONFIG_ADMIN_PASSWORD || '').trim();
    if (envPw) {
      if (!timingSafeStringEqual(clean, envPw)) {
        throw new Error('서버 .env의 CONFIG_ADMIN_PASSWORD와 일치하지 않습니다.');
      }
      return this.createAdminSession('.env 종합 관리자');
    }
    if (!this.state.adminClaim) {
      throw new Error('먼저 종합 관리자 초기 비밀번호를 설정해 주세요.');
    }
    const computed = hashPasscodeScrypt(clean, this.state.adminClaim.salt);
    if (!timingSafeStringEqual(computed, this.state.adminClaim.hash)) {
      throw new Error('종합 관리자 비밀번호가 일치하지 않습니다.');
    }
    return this.createAdminSession('종합 관리자');
  }

  private createAdminSession(label: string): EmvAdminSession {
    this.purgeExpiredSessions();
    const now = Date.now();
    const session: EmvAdminSession = {
      token: crypto.randomBytes(32).toString('hex'),
      createdAt: now,
      expiresAt: now + ADMIN_SESSION_TTL_MS,
      label,
    };
    this.adminSessions.set(session.token, session);
    return session;
  }

  verifyAdminSession(token?: string | null): EmvAdminSession | null {
    if (!token) return null;
    const session = this.adminSessions.get(token);
    if (!session) return null;
    if (Date.now() > session.expiresAt) {
      this.adminSessions.delete(token);
      return null;
    }
    return session;
  }

  revokeAdminSession(token?: string | null): void {
    if (token) this.adminSessions.delete(token);
  }

  private purgeExpiredSessions(): void {
    const now = Date.now();
    for (const [k, v] of this.adminSessions.entries()) {
      if (now > v.expiresAt) this.adminSessions.delete(k);
    }
  }

  // --- EMV CRUD (Global & Per-Employee) ---

  setEntry(params: {
    scope: EmvScope;
    employeeId?: string;
    key: string;
    value: string;
    updatedBy?: string;
  }): PublicEmvEntryView {
    if (params.scope !== 'global' && params.scope !== 'employee') {
      throw new Error('유효하지 않은 EMV 적용 범위입니다 (global 또는 employee).');
    }
    if (!isAllowedEmvKey(params.key)) {
      throw new Error(`허용되지 않은 EMV 키 이름입니다: ${params.key}`);
    }
    const cleanValue = String(params.value || '').trim();
    if (!cleanValue || cleanValue.length > MAX_VALUE_LENGTH || /[\r\n\0]/.test(cleanValue)) {
      throw new Error('EMV 값은 비어 있거나 줄바꿈/제어 문자를 포함할 수 없습니다.');
    }
    if (params.key === 'AGENT_BUILDER_WORKER_URL') {
      let parsedUrl: URL;
      try {
        parsedUrl = new URL(cleanValue);
      } catch {
        throw new Error('AGENT_BUILDER_WORKER_URL은 유효한 https:// 주소여야 합니다.');
      }
      if (parsedUrl.protocol !== 'https:') {
        throw new Error('AGENT_BUILDER_WORKER_URL은 반드시 https:// 프로토콜을 사용해야 합니다.');
      }
    }

    const updatedAt = new Date().toISOString();
    const updatedBy = String(params.updatedBy || '종합 관리자').slice(0, 80);
    const encrypted = this.encryptValue(cleanValue);
    const maskedPreview = maskEmvValue(params.key, cleanValue);
    const fingerprint = computeSecretFingerprint(cleanValue);

    if (params.scope === 'global') {
      const record: StoredEmvRecord = {
        key: params.key,
        scope: 'global',
        encrypted,
        maskedPreview,
        fingerprint,
        updatedAt,
        updatedBy,
      };
      this.state.globalEntries[params.key] = record;
      this.persistEnvelope();
      return this.toPublicView(record);
    }

    const normalizedEmp = normalizeEmployeeId(params.employeeId);
    if (!normalizedEmp) {
      throw new Error('유효한 사원 이메일 또는 사원 ID를 입력해 주세요. (예: hong@company.co.kr)');
    }
    const partKey = this.computeEmployeePartitionKey(normalizedEmp);
    if (!this.state.employeeEntries[partKey]) {
      if (Object.keys(this.state.employeeEntries).length >= MAX_EMPLOYEES) {
        throw new Error(`최대 등록 가능 사원 수(${MAX_EMPLOYEES}명)를 초과했습니다.`);
      }
      this.state.employeeEntries[partKey] = {
        employeeId: normalizedEmp,
        entries: {},
      };
    }

    const record: StoredEmvRecord = {
      key: params.key,
      scope: 'employee',
      employeeId: normalizedEmp,
      encrypted,
      maskedPreview,
      fingerprint,
      updatedAt,
      updatedBy,
    };
    this.state.employeeEntries[partKey].entries[params.key] = record;
    this.persistEnvelope();
    return this.toPublicView(record);
  }

  deleteEntry(params: { scope: EmvScope; employeeId?: string; key: string }): boolean {
    if (!isAllowedEmvKey(params.key)) {
      throw new Error(`허용되지 않은 EMV 키 이름입니다: ${params.key}`);
    }
    if (params.scope === 'global') {
      if (!this.state.globalEntries[params.key]) return false;
      delete this.state.globalEntries[params.key];
      this.persistEnvelope();
      return true;
    }

    const normalizedEmp = normalizeEmployeeId(params.employeeId);
    if (!normalizedEmp) {
      throw new Error('삭제할 사원 이메일 또는 사원 ID가 올바르지 않습니다.');
    }
    const partKey = this.computeEmployeePartitionKey(normalizedEmp);
    const bucket = this.state.employeeEntries[partKey];
    if (!bucket || !bucket.entries[params.key]) return false;
    delete bucket.entries[params.key];
    if (Object.keys(bucket.entries).length === 0) {
      delete this.state.employeeEntries[partKey];
    }
    this.persistEnvelope();
    return true;
  }

  private toPublicView(rec: StoredEmvRecord): PublicEmvEntryView {
    return {
      key: rec.key,
      scope: rec.scope,
      ...(rec.employeeId ? { employeeId: rec.employeeId } : {}),
      maskedPreview: rec.maskedPreview,
      fingerprint: rec.fingerprint,
      updatedAt: rec.updatedAt,
      updatedBy: rec.updatedBy,
    };
  }

  /**
   * Lists all stored EMV entries with masked previews ONLY (never raw secrets).
   */
  listPublicEntries(): {
    globalEntries: PublicEmvEntryView[];
    employees: Array<{ employeeId: string; entries: PublicEmvEntryView[] }>;
  } {
    const globalEntries = Object.values(this.state.globalEntries)
      .map((rec) => this.toPublicView(rec))
      .sort((a, b) => a.key.localeCompare(b.key));

    const employees = Object.values(this.state.employeeEntries)
      .map((bucket) => ({
        employeeId: bucket.employeeId,
        entries: Object.values(bucket.entries)
          .map((rec) => this.toPublicView(rec))
          .sort((a, b) => a.key.localeCompare(b.key)),
      }))
      .sort((a, b) => a.employeeId.localeCompare(b.employeeId));

    return { globalEntries, employees };
  }

  /**
   * Server-side runtime key resolver (INTERNAL ONLY — NEVER send raw output to clients!).
   * Resolution precedence:
   *   1) Verified employee-specific EMV entry (ONLY when verifiedEmployeeId is authenticated server-side)
   *   2) Org-wide (global) EMV entry saved by super-admin
   *   3) Server process.env fallback
   */
  resolveRuntimeEnv(verifiedEmployeeId?: string | null): Partial<Record<EmvKeyName, string>> {
    const resolved: Partial<Record<EmvKeyName, string>> = {};
    const normalizedEmp = normalizeEmployeeId(verifiedEmployeeId);
    const empBucket = normalizedEmp
      ? this.state.employeeEntries[this.computeEmployeePartitionKey(normalizedEmp)]
      : undefined;

    for (const key of ALLOWED_EMV_KEYS) {
      // 1. Per-employee verified override
      const empRec = empBucket?.entries[key];
      if (empRec) {
        const plain = this.decryptValue(empRec.encrypted);
        if (plain && plain.trim()) {
          resolved[key] = plain.trim();
          continue;
        }
      }
      // 2. Org-wide global EMV entry
      const globalRec = this.state.globalEntries[key];
      if (globalRec) {
        const plain = this.decryptValue(globalRec.encrypted);
        if (plain && plain.trim()) {
          resolved[key] = plain.trim();
          continue;
        }
      }
      // 3. Server process.env
      const envVal = (this.envSource[key] || '').trim();
      if (envVal && envVal !== 'MY_GEMINI_API_KEY') {
        resolved[key] = envVal;
      }
    }

    return resolved;
  }

  /**
   * Safe public preview showing which layer (employee-emv > global-emv > server-env) supplies each key
   * for a given employee, returning ONLY masked previews and source labels.
   */
  getResolutionPreview(employeeId?: string | null): {
    employeeId: string | null;
    items: EmvResolutionSourceItem[];
  } {
    const normalizedEmp = normalizeEmployeeId(employeeId);
    const empBucket = normalizedEmp
      ? this.state.employeeEntries[this.computeEmployeePartitionKey(normalizedEmp)]
      : undefined;

    const items: EmvResolutionSourceItem[] = ALLOWED_EMV_KEYS.map((key) => {
      const empRec = empBucket?.entries[key];
      if (empRec) {
        return {
          key,
          configured: true,
          source: 'employee-emv',
          maskedPreview: empRec.maskedPreview,
        };
      }
      const globalRec = this.state.globalEntries[key];
      if (globalRec) {
        return {
          key,
          configured: true,
          source: 'global-emv',
          maskedPreview: globalRec.maskedPreview,
        };
      }
      const envVal = (this.envSource[key] || '').trim();
      if (envVal && envVal !== 'MY_GEMINI_API_KEY') {
        return {
          key,
          configured: true,
          source: 'server-env',
          maskedPreview: maskEmvValue(key, envVal),
        };
      }
      return {
        key,
        configured: false,
        source: 'unconfigured',
        maskedPreview: '',
      };
    });

    return {
      employeeId: normalizedEmp || null,
      items,
    };
  }
}
