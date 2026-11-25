import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type {
  DomainGlossaryEntry,
  LegacyInternalMailRow,
  SgwAuditLog,
  SgwMailAttachment,
  SgwMailboxFolder,
  SgwMailLimits,
  SgwMailRecipient,
  SgwMailRule,
  SgwMailViewItem,
  SgwUserAiPreferences,
  SgwUserIdentity,
} from '../types/sgwSedms';
import { signAttachmentDownloadToken } from '../utils/sgwAuth';

export const DEFAULT_SGW_MAIL_LIMITS: SgwMailLimits = {
  maxSubjectLength: 200,
  maxBodyLength: 100_000,
  maxAttachmentBytes: 10 * 1024 * 1024,
  allowedAttachmentExtensions: [
    '.pdf',
    '.xlsx',
    '.xls',
    '.csv',
    '.docx',
    '.doc',
    '.pptx',
    '.txt',
    '.png',
    '.jpg',
    '.jpeg',
    '.dwg',
    '.dxf',
    '.zip',
  ],
};

const BLOCKED_DANGEROUS_EXTENSIONS = new Set([
  '.exe',
  '.bat',
  '.cmd',
  '.ps1',
  '.sh',
  '.msi',
  '.vbs',
  '.js',
  '.scr',
  '.com',
  '.dll',
  '.jar',
]);

export class SgwAccessError extends Error {
  readonly statusCode: number;

  constructor(statusCode: number, message: string) {
    super(message);
    this.name = 'SgwAccessError';
    this.statusCode = statusCode;
  }
}

export function validateAttachmentPolicy(
  filename: string,
  sizeBytes: number,
  limits: SgwMailLimits = DEFAULT_SGW_MAIL_LIMITS
): { ok: boolean; extension: string; error?: string } {
  const cleanName = path.basename(String(filename || '').trim());
  const ext = path.extname(cleanName).toLowerCase();
  if (!cleanName || !ext) {
    return { ok: false, extension: ext, error: '유효한 파일명과 확장자가 필요합니다.' };
  }
  if (!Number.isFinite(sizeBytes) || sizeBytes <= 0) {
    return { ok: false, extension: ext, error: '0바이트 빈 첨부파일은 업로드할 수 없습니다.' };
  }
  if (sizeBytes > limits.maxAttachmentBytes) {
    return {
      ok: false,
      extension: ext,
      error: `첨부파일 크기(${sizeBytes}B)가 허용 한도(${limits.maxAttachmentBytes}B)를 초과했습니다.`,
    };
  }
  if (BLOCKED_DANGEROUS_EXTENSIONS.has(ext) || !limits.allowedAttachmentExtensions.includes(ext)) {
    return {
      ok: false,
      extension: ext,
      error: `보안 정책상 허용되지 않는 첨부파일 확장자(${ext})입니다.`,
    };
  }
  return { ok: true, extension: ext };
}

interface StoredSgwMailRecord {
  id: string;
  threadId: string;
  senderSub: string;
  senderName: string;
  senderDeletedAt: string | null;
  senderPurged: boolean;
  recipients: SgwMailRecipient[];
  subject: string;
  body: string;
  isDraft: boolean;
  createdAt: string;
  updatedAt: string;
  attachments: SgwMailAttachment[];
  legacyInternalMailId?: string;
}

interface PersistedSgwState {
  mails: StoredSgwMailRecord[];
  preferences: Record<string, SgwUserAiPreferences>;
  glossary: DomainGlossaryEntry[];
  rules: SgwMailRule[];
  auditLogs: SgwAuditLog[];
}

export interface SendOrSaveDraftInput {
  to?: Array<{ sub: string; name?: string }>;
  cc?: Array<{ sub: string; name?: string }>;
  bcc?: Array<{ sub: string; name?: string }>;
  subject?: string;
  body?: string;
  isDraft?: boolean;
  threadId?: string;
  attachments?: Array<Omit<SgwMailAttachment, 'id'> & { id?: string }>;
  // Untrusted fields potentially spoofed by clients — always ignored by server
  senderSub?: string;
  senderId?: string;
}

export class SgwMailStore {
  private readonly persistPath?: string;
  private readonly signingSecret: string;
  private readonly limits: SgwMailLimits;
  private mails: StoredSgwMailRecord[] = [];
  private preferences: Record<string, SgwUserAiPreferences> = {};
  private glossary: DomainGlossaryEntry[] = [];
  private rules: SgwMailRule[] = [];
  private auditLogs: SgwAuditLog[] = [];

  constructor(options?: {
    persistPath?: string;
    signingSecret?: string;
    limits?: SgwMailLimits;
  }) {
    this.persistPath = options?.persistPath;
    this.signingSecret = options?.signingSecret ?? process.env.SGW_SSO_SECRET?.trim() ?? '';
    this.limits = options?.limits ?? DEFAULT_SGW_MAIL_LIMITS;
    this.loadFromDisk();
  }

  getLimits(): SgwMailLimits {
    return {
      ...this.limits,
      allowedAttachmentExtensions: [...this.limits.allowedAttachmentExtensions],
    };
  }

  private loadFromDisk(): void {
    if (!this.persistPath || !fs.existsSync(this.persistPath)) return;
    try {
      const raw = JSON.parse(fs.readFileSync(this.persistPath, 'utf8')) as Partial<PersistedSgwState>;
      this.mails = Array.isArray(raw.mails) ? raw.mails : [];
      this.preferences = raw.preferences && typeof raw.preferences === 'object' ? raw.preferences : {};
      this.glossary = Array.isArray(raw.glossary) ? raw.glossary : [];
      this.rules = Array.isArray(raw.rules) ? raw.rules : [];
      this.auditLogs = Array.isArray(raw.auditLogs) ? raw.auditLogs : [];
    } catch {
      // Ignore corrupt persistence cleanly and start empty
    }
  }

  private persist(): void {
    if (!this.persistPath) return;
    try {
      const dir = path.dirname(this.persistPath);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      const payload: PersistedSgwState = {
        mails: this.mails,
        preferences: this.preferences,
        glossary: this.glossary,
        rules: this.rules,
        auditLogs: this.auditLogs.slice(-500),
      };
      fs.writeFileSync(this.persistPath, JSON.stringify(payload, null, 2), 'utf8');
    } catch {
      // Avoid crashing request path on non-fatal disk write error
    }
  }

  private getPartyRole(mail: StoredSgwMailRecord, sub: string): {
    isSender: boolean;
    recipientEntry?: SgwMailRecipient;
  } {
    const isSender = mail.senderSub === sub;
    const recipientEntry = mail.recipients.find((r) => r.sub === sub);
    return { isSender, recipientEntry };
  }

  private toUserView(mail: StoredSgwMailRecord, sub: string, preferredFolder?: SgwMailboxFolder): SgwMailViewItem {
    const { isSender, recipientEntry } = this.getPartyRole(mail, sub);

    // BCC privacy: sender sees all; recipient sees non-BCC plus themselves if they are BCC
    const visibleRecipients = mail.recipients
      .filter((r) => {
        if (isSender) return true;
        if (r.type !== 'bcc') return true;
        return r.sub === sub;
      })
      .map((r) => ({ ...r }));

    let folder: SgwMailboxFolder = 'inbox';
    let isRead = true;

    if (mail.isDraft && isSender) {
      folder = mail.senderDeletedAt ? 'trash' : 'drafts';
      isRead = true;
    } else if (preferredFolder === 'sent' && isSender) {
      folder = mail.senderDeletedAt ? 'trash' : 'sent';
      isRead = true;
    } else if (preferredFolder === 'inbox' && recipientEntry) {
      folder = recipientEntry.deletedAt ? 'trash' : 'inbox';
      isRead = recipientEntry.isRead;
    } else if (preferredFolder === 'trash') {
      folder = 'trash';
      isRead = recipientEntry ? recipientEntry.isRead : true;
    } else if (recipientEntry && !recipientEntry.purged) {
      folder = recipientEntry.deletedAt ? 'trash' : 'inbox';
      isRead = recipientEntry.isRead;
    } else if (isSender && !mail.senderPurged) {
      folder = mail.senderDeletedAt ? 'trash' : 'sent';
      isRead = true;
    }

    const nowSec = Math.floor(Date.now() / 1000);
    const attachments: SgwMailAttachment[] = mail.attachments.map((att) => ({
      ...att,
      ...(this.signingSecret
        ? {
            downloadToken: signAttachmentDownloadToken({
              attachmentId: att.id,
              ownerSub: sub,
              expiresAtSec: nowSec + 300,
              secret: this.signingSecret,
            }),
          }
        : {}),
    }));

    return {
      id: mail.id,
      threadId: mail.threadId,
      senderSub: mail.senderSub,
      senderName: mail.senderName,
      recipients: visibleRecipients,
      subject: mail.subject,
      body: mail.body,
      folder,
      isDraft: mail.isDraft,
      isRead,
      sentAt: mail.createdAt,
      createdAt: mail.createdAt,
      updatedAt: mail.updatedAt,
      attachments,
      legacyInternalMailId: mail.legacyInternalMailId,
    };
  }

  sendOrSaveDraft(user: SgwUserIdentity, input: SendOrSaveDraftInput): SgwMailViewItem {
    const subject = String(input.subject || '').trim();
    const body = String(input.body || '');
    const isDraft = Boolean(input.isDraft);

    if (subject.length > this.limits.maxSubjectLength) {
      throw new SgwAccessError(400, `제목 길이가 최대(${this.limits.maxSubjectLength}자)를 초과했습니다.`);
    }
    if (body.length > this.limits.maxBodyLength) {
      throw new SgwAccessError(400, `본문 길이가 최대(${this.limits.maxBodyLength}자)를 초과했습니다.`);
    }

    const recipients: SgwMailRecipient[] = [];
    const pushList = (list: Array<{ sub: string; name?: string }> | undefined, type: 'to' | 'cc' | 'bcc') => {
      if (!Array.isArray(list)) return;
      for (const entry of list) {
        const sub = String(entry?.sub || '').trim();
        if (!sub) continue;
        recipients.push({
          sub,
          name: String(entry?.name || sub).trim(),
          type,
          isRead: false,
          deletedAt: null,
          purged: false,
        });
      }
    };
    pushList(input.to, 'to');
    pushList(input.cc, 'cc');
    pushList(input.bcc, 'bcc');

    if (!isDraft && recipients.length === 0) {
      throw new SgwAccessError(400, '최소 1명의 수신자(To/Cc/Bcc)를 지정해 주세요.');
    }

    const attachments: SgwMailAttachment[] = [];
    if (Array.isArray(input.attachments)) {
      for (const rawAtt of input.attachments) {
        const policy = validateAttachmentPolicy(rawAtt.filename, rawAtt.sizeBytes, this.limits);
        if (!policy.ok) {
          throw new SgwAccessError(400, policy.error || '첨부파일 정책 위반');
        }
        attachments.push({
          id: rawAtt.id || `att_${crypto.randomUUID().slice(0, 12)}`,
          filename: path.basename(rawAtt.filename),
          mimeType: rawAtt.mimeType || 'application/octet-stream',
          sizeBytes: rawAtt.sizeBytes,
          sha256: rawAtt.sha256,
          storageKey: rawAtt.storageKey,
          scanStatus: rawAtt.scanStatus ?? 'clean',
        });
      }
    }

    const now = new Date().toISOString();
    const record: StoredSgwMailRecord = {
      id: `sgw_${crypto.randomUUID().slice(0, 12)}`,
      threadId: String(input.threadId || '').trim() || `thr_${crypto.randomUUID().slice(0, 10)}`,
      // Always bind sender strictly to authenticated server identity (never trust request body)
      senderSub: user.sub,
      senderName: user.name,
      senderDeletedAt: null,
      senderPurged: false,
      recipients,
      subject: subject || '(제목 없음)',
      body,
      isDraft,
      createdAt: now,
      updatedAt: now,
      attachments,
    };

    this.mails.unshift(record);
    this.recordAudit(user.sub, isDraft ? 'mail.draft_save' : 'mail.send', record.id);
    this.persist();
    return this.toUserView(record, user.sub, isDraft ? 'drafts' : 'sent');
  }

  listMail(
    user: SgwUserIdentity,
    filter: {
      box?: 'all' | SgwMailboxFolder;
      unreadOnly?: boolean;
      query?: string;
      threadId?: string;
    } = {}
  ): SgwMailViewItem[] {
    const box = filter.box ?? 'all';
    const q = (filter.query ?? '').trim().toLowerCase();
    const results: SgwMailViewItem[] = [];

    for (const mail of this.mails) {
      if (filter.threadId && mail.threadId !== filter.threadId) continue;
      const { isSender, recipientEntry } = this.getPartyRole(mail, user.sub);
      if (!isSender && !recipientEntry) continue;

      if (mail.isDraft && !isSender) continue;

      const senderVisible = isSender && !mail.senderPurged;
      const recipientVisible = Boolean(recipientEntry && !recipientEntry.purged);

      let matchesBox = false;
      let viewFolder: SgwMailboxFolder | undefined;

      if (box === 'inbox') {
        matchesBox = !mail.isDraft && recipientVisible && !recipientEntry?.deletedAt;
        viewFolder = 'inbox';
      } else if (box === 'sent') {
        matchesBox = !mail.isDraft && senderVisible && !mail.senderDeletedAt;
        viewFolder = 'sent';
      } else if (box === 'drafts') {
        matchesBox = mail.isDraft && senderVisible && !mail.senderDeletedAt;
        viewFolder = 'drafts';
      } else if (box === 'trash') {
        const senderInTrash = senderVisible && Boolean(mail.senderDeletedAt);
        const recInTrash = recipientVisible && Boolean(recipientEntry?.deletedAt);
        matchesBox = senderInTrash || recInTrash;
        viewFolder = 'trash';
      } else {
        matchesBox = senderVisible || recipientVisible;
      }

      if (!matchesBox) continue;

      const view = this.toUserView(mail, user.sub, viewFolder);
      if (filter.unreadOnly && view.isRead) continue;
      if (
        q &&
        !`${view.subject} ${view.body} ${view.senderName}`.toLowerCase().includes(q)
      ) {
        continue;
      }
      results.push(view);
    }

    return results.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
  }

  getMailById(user: SgwUserIdentity, mailId: string): SgwMailViewItem {
    const mail = this.mails.find((m) => m.id === mailId);
    if (!mail) {
      throw new SgwAccessError(404, '요청한 메일을 찾을 수 없습니다.');
    }
    const { isSender, recipientEntry } = this.getPartyRole(mail, user.sub);
    if (!isSender && !recipientEntry) {
      throw new SgwAccessError(403, '본인이 발신자 또는 수신자인 메일만 열람할 수 있습니다.');
    }
    if (mail.isDraft && !isSender) {
      throw new SgwAccessError(403, '임시보관 메일은 작성자 본인만 열람할 수 있습니다.');
    }
    const senderPurgedOnly = isSender && mail.senderPurged && (!recipientEntry || recipientEntry.purged);
    const recPurgedOnly = !isSender && Boolean(recipientEntry?.purged);
    if (senderPurgedOnly || recPurgedOnly) {
      throw new SgwAccessError(404, '영구 삭제된 메일입니다.');
    }
    return this.toUserView(mail, user.sub);
  }

  setReadStatus(user: SgwUserIdentity, mailId: string, isRead: boolean): SgwMailViewItem {
    const mail = this.mails.find((m) => m.id === mailId);
    if (!mail) {
      throw new SgwAccessError(404, '요청한 메일을 찾을 수 없습니다.');
    }
    const recipientEntry = mail.recipients.find((r) => r.sub === user.sub && !r.purged);
    if (!recipientEntry) {
      throw new SgwAccessError(403, '수신자 본인만 읽음 상태를 변경할 수 있습니다.');
    }
    recipientEntry.isRead = Boolean(isRead);
    recipientEntry.readAt = isRead ? new Date().toISOString() : undefined;
    mail.updatedAt = new Date().toISOString();
    this.recordAudit(user.sub, isRead ? 'mail.mark_read' : 'mail.mark_unread', mail.id);
    this.persist();
    return this.toUserView(mail, user.sub, 'inbox');
  }

  deleteMail(
    user: SgwUserIdentity,
    mailId: string,
    options?: { permanent?: boolean; restore?: boolean }
  ): { ok: true; state: 'trashed' | 'restored' | 'purged'; mailId: string } {
    const mail = this.mails.find((m) => m.id === mailId);
    if (!mail) {
      throw new SgwAccessError(404, '요청한 메일을 찾을 수 없습니다.');
    }
    const { isSender, recipientEntry } = this.getPartyRole(mail, user.sub);
    if (!isSender && !recipientEntry) {
      throw new SgwAccessError(403, '발신자 또는 수신자 당사자만 메일을 삭제할 수 있습니다.');
    }

    const now = new Date().toISOString();
    const isTargetRecipient = Boolean(recipientEntry && !recipientEntry.purged);
    const isTargetSender = isSender && !mail.senderPurged;

    if (options?.restore) {
      if (isTargetRecipient && recipientEntry) recipientEntry.deletedAt = null;
      if (isTargetSender) mail.senderDeletedAt = null;
      mail.updatedAt = now;
      this.recordAudit(user.sub, 'mail.restore', mail.id);
      this.persist();
      return { ok: true, state: 'restored', mailId: mail.id };
    }

    if (options?.permanent) {
      const inTrash =
        (isTargetRecipient && Boolean(recipientEntry?.deletedAt)) ||
        (isTargetSender && Boolean(mail.senderDeletedAt));
      if (!inTrash) {
        throw new SgwAccessError(
          400,
          '2단계 삭제 정책: 휴지통으로 먼저 이동한 메일만 영구 삭제할 수 있습니다.'
        );
      }
      if (isTargetRecipient && recipientEntry && recipientEntry.deletedAt) {
        recipientEntry.purged = true;
      }
      if (isTargetSender && mail.senderDeletedAt) {
        mail.senderPurged = true;
      }
      mail.updatedAt = now;
      this.recordAudit(user.sub, 'mail.purge', mail.id);
      this.persist();
      return { ok: true, state: 'purged', mailId: mail.id };
    }

    // Stage 1: Move to trash for the acting party only
    if (isTargetRecipient && recipientEntry && !recipientEntry.deletedAt) {
      recipientEntry.deletedAt = now;
    } else if (isTargetSender && !mail.senderDeletedAt) {
      mail.senderDeletedAt = now;
    }
    mail.updatedAt = now;
    this.recordAudit(user.sub, 'mail.trash', mail.id);
    this.persist();
    return { ok: true, state: 'trashed', mailId: mail.id };
  }

  getAttachmentForUser(user: SgwUserIdentity, attachmentId: string): SgwMailAttachment {
    for (const mail of this.mails) {
      const att = mail.attachments.find((a) => a.id === attachmentId);
      if (!att) continue;
      const { isSender, recipientEntry } = this.getPartyRole(mail, user.sub);
      if (!isSender && !recipientEntry) {
        throw new SgwAccessError(403, '메일 당사자 본인만 첨부파일을 다운로드할 수 있습니다.');
      }
      return { ...att };
    }
    throw new SgwAccessError(404, '첨부파일을 찾을 수 없습니다.');
  }

  getUserAiPreferences(sub: string): SgwUserAiPreferences {
    const existing = this.preferences[sub];
    if (existing) {
      return { ...existing, features: { ...existing.features } };
    }
    return {
      sub,
      aiEnabled: false,
      features: {
        summarize: false,
        replyDraft: false,
        classifyAndRules: false,
      },
      tone: 'formal',
      length: 'medium',
      language: 'ko',
      signature: '',
      allowTraining: false,
      updatedAt: new Date().toISOString(),
    };
  }

  updateUserAiPreferences(
    user: SgwUserIdentity,
    patch: Partial<SgwUserAiPreferences>
  ): SgwUserAiPreferences {
    const current = this.getUserAiPreferences(user.sub);
    const updated: SgwUserAiPreferences = {
      ...current,
      aiEnabled: typeof patch.aiEnabled === 'boolean' ? patch.aiEnabled : current.aiEnabled,
      features: {
        summarize:
          typeof patch.features?.summarize === 'boolean'
            ? patch.features.summarize
            : current.features.summarize,
        replyDraft:
          typeof patch.features?.replyDraft === 'boolean'
            ? patch.features.replyDraft
            : current.features.replyDraft,
        classifyAndRules:
          typeof patch.features?.classifyAndRules === 'boolean'
            ? patch.features.classifyAndRules
            : current.features.classifyAndRules,
      },
      tone:
        patch.tone === 'concise' || patch.tone === 'friendly' || patch.tone === 'formal'
          ? patch.tone
          : current.tone,
      length:
        patch.length === 'short' || patch.length === 'detailed' || patch.length === 'medium'
          ? patch.length
          : current.length,
      language:
        patch.language === 'en' || patch.language === 'auto' || patch.language === 'ko'
          ? patch.language
          : current.language,
      signature: typeof patch.signature === 'string' ? patch.signature.slice(0, 500) : current.signature,
      allowTraining:
        typeof patch.allowTraining === 'boolean' ? patch.allowTraining : current.allowTraining,
      updatedAt: new Date().toISOString(),
    };
    this.preferences[user.sub] = updated;
    this.recordAudit(user.sub, 'preferences.update', user.sub);
    this.persist();
    return { ...updated, features: { ...updated.features } };
  }

  listGlossary(user: SgwUserIdentity): DomainGlossaryEntry[] {
    return this.glossary.filter((g) => g.ownerSub === user.sub).map((g) => ({ ...g }));
  }

  upsertGlossaryEntry(
    user: SgwUserIdentity,
    input: { id?: string; term?: string; definition?: string; category?: string }
  ): DomainGlossaryEntry {
    const term = String(input.term || '').trim();
    const definition = String(input.definition || '').trim();
    if (!term || !definition) {
      throw new SgwAccessError(400, '용어(term)와 정의(definition)는 필수입니다.');
    }
    const existing = input.id
      ? this.glossary.find((g) => g.id === input.id && g.ownerSub === user.sub)
      : undefined;
    if (existing) {
      existing.term = term;
      existing.definition = definition;
      existing.category = String(input.category || existing.category || '일반').trim();
      this.persist();
      return { ...existing };
    }
    const created: DomainGlossaryEntry = {
      id: `glo_${crypto.randomUUID().slice(0, 10)}`,
      ownerSub: user.sub,
      term,
      definition,
      category: String(input.category || '일반').trim(),
      createdAt: new Date().toISOString(),
    };
    this.glossary.unshift(created);
    this.persist();
    return { ...created };
  }

  deleteGlossaryEntry(user: SgwUserIdentity, id: string): boolean {
    const before = this.glossary.length;
    this.glossary = this.glossary.filter((g) => !(g.id === id && g.ownerSub === user.sub));
    const deleted = this.glossary.length < before;
    if (deleted) this.persist();
    return deleted;
  }

  listRules(user: SgwUserIdentity): SgwMailRule[] {
    return this.rules.filter((r) => r.ownerSub === user.sub).map((r) => ({ ...r }));
  }

  addRule(
    user: SgwUserIdentity,
    input: { name?: string; matchKeyword?: string; targetCategory?: string }
  ): SgwMailRule {
    const name = String(input.name || '').trim();
    const matchKeyword = String(input.matchKeyword || '').trim();
    if (!name || !matchKeyword) {
      throw new SgwAccessError(400, '규칙 이름과 매칭 키워드는 필수입니다.');
    }
    const created: SgwMailRule = {
      id: `rule_${crypto.randomUUID().slice(0, 10)}`,
      ownerSub: user.sub,
      name,
      matchKeyword,
      targetCategory: String(input.targetCategory || '자동분류').trim(),
      enabled: true,
      createdAt: new Date().toISOString(),
    };
    this.rules.unshift(created);
    this.persist();
    return { ...created };
  }

  deleteRule(user: SgwUserIdentity, id: string): boolean {
    const before = this.rules.length;
    this.rules = this.rules.filter((r) => !(r.id === id && r.ownerSub === user.sub));
    const deleted = this.rules.length < before;
    if (deleted) this.persist();
    return deleted;
  }

  recordAudit(actorSub: string, action: string, targetId: string, detail?: string): void {
    this.auditLogs.unshift({
      id: `aud_${crypto.randomUUID().slice(0, 10)}`,
      actorSub,
      action,
      targetId,
      createdAt: new Date().toISOString(),
      detail,
    });
    if (this.auditLogs.length > 500) {
      this.auditLogs = this.auditLogs.slice(0, 500);
    }
  }

  listAuditLogs(user: SgwUserIdentity): SgwAuditLog[] {
    return this.auditLogs.filter((l) => l.actorSub === user.sub).slice(0, 100);
  }

  getMigratedLegacyIds(): Set<string> {
    const ids = new Set<string>();
    for (const m of this.mails) {
      if (m.legacyInternalMailId) ids.add(m.legacyInternalMailId);
    }
    return ids;
  }

  insertMigratedRow(row: LegacyInternalMailRow): SgwMailViewItem {
    const legacyId = String(row.id);
    const createdAt = row.createdAt || new Date().toISOString();
    const isRead = row.isRead === true || row.isRead === 1;
    const record: StoredSgwMailRecord = {
      id: `sgw_mig_${legacyId}_${crypto.randomUUID().slice(0, 6)}`,
      threadId: `thr_mig_${legacyId}`,
      senderSub: String(row.senderId).trim(),
      senderName: String(row.senderName || row.senderId).trim(),
      senderDeletedAt: null,
      senderPurged: false,
      recipients: [
        {
          sub: String(row.recipientId).trim(),
          name: String(row.recipientName || row.recipientId).trim(),
          type: 'to',
          isRead,
          readAt: isRead ? createdAt : undefined,
          deletedAt: null,
          purged: false,
        },
      ],
      subject: String(row.subject || '(제목 없음)'),
      body: String(row.body || ''),
      isDraft: false,
      createdAt,
      updatedAt: createdAt,
      attachments: [],
      legacyInternalMailId: legacyId,
    };
    this.mails.push(record);
    this.persist();
    return this.toUserView(record, record.recipients[0].sub, 'inbox');
  }

  rollbackMigratedRows(actorSub: string): { removedMessages: number } {
    const before = this.mails.length;
    this.mails = this.mails.filter((m) => !m.legacyInternalMailId);
    const removedMessages = before - this.mails.length;
    this.recordAudit(actorSub, 'migration.rollback', `removed:${removedMessages}`);
    this.persist();
    return { removedMessages };
  }
}
