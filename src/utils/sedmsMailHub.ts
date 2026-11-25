import crypto from 'node:crypto';
import type {
  SedmsIngestMessagePayload,
  SedmsMailLink,
  SedmsMatchCandidate,
  SedmsOutboundDispatchRef,
  SedmsProjectCatalogItem,
  SedmsReviewQueueItem,
  SedmsStoredAttachment,
  SedmsStoredMessage,
} from '../types/sgwSedms';

export function normalizeAlphaNumeric(input: string): string {
  return String(input || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

export function canonDrawingNo(input: string): string {
  return normalizeAlphaNumeric(input);
}

export function createFlexibleDrawingPattern(drawingNo: string): RegExp | null {
  const canon = canonDrawingNo(drawingNo);
  if (canon.length < 4) return null;
  const flex = canon.split('').join('[^A-Z0-9]*');
  return new RegExp(`(?<![A-Z0-9])${flex}(?![A-Z0-9])`, 'i');
}

function normalizeMessageId(raw?: string): string {
  return String(raw || '')
    .trim()
    .replace(/^<+|>+$/g, '')
    .toLowerCase();
}

function extractEmailDomain(addr: string): { email: string; domain: string } {
  const match = String(addr || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  const email = (match ? match[0] : String(addr || '')).trim().toLowerCase();
  const domain = email.includes('@') ? email.split('@')[1] : '';
  return { email, domain };
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function matchesProjectToken(haystack: string, token: string): boolean {
  const trimmed = String(token || '').trim();
  if (trimmed.length < 2) return false;
  const escaped = escapeRegExp(trimmed);
  const re = new RegExp(`(?<![A-Z0-9])${escaped}(?![A-Z0-9])`, 'i');
  return re.test(haystack);
}

export interface SedmsDeterministicMatchOutcome {
  decision: 'auto_link' | 'review_queue';
  reason: string;
  links: SedmsMailLink[];
  candidates: SedmsMatchCandidate[];
}

/**
 * Deterministic 5-tier SEDMS mail matcher (zero LLM, strict tenant isolation).
 * Rule 1: In-Reply-To / References -> outbound dispatch tracking
 * Rule 2: Canonical drawingNoKey in subject/body/attachment filenames
 * Rule 3: projectCode / vesselAliases DB token match
 * Rule 4: Company domain / email match -> ALWAYS review_queue
 * Rule 5: Unclassified -> honest empty review_queue
 * Any cross-rule or multi-project conflict -> review_queue (never blindly auto-link).
 */
export function matchMailDeterministically(params: {
  message: SedmsIngestMessagePayload;
  projects: SedmsProjectCatalogItem[];
  outboundDispatches: SedmsOutboundDispatchRef[];
}): SedmsDeterministicMatchOutcome {
  const { message } = params;
  const tenantProjects = params.projects.filter((p) => p.tenantId === message.tenantId);
  const tenantDispatches = params.outboundDispatches.filter((d) => d.tenantId === message.tenantId);
  const now = new Date().toISOString();

  // Rule 1: In-Reply-To / References matching outbound dispatch Message-ID
  const replyIds = new Set<string>();
  if (message.inReplyTo) {
    const norm = normalizeMessageId(message.inReplyTo);
    if (norm) replyIds.add(norm);
  }
  for (const ref of message.references ?? []) {
    const norm = normalizeMessageId(ref);
    if (norm) replyIds.add(norm);
  }

  const rule1Matches = tenantDispatches.filter((d) => replyIds.has(normalizeMessageId(d.outboundMessageId)));
  if (rule1Matches.length > 0) {
    const uniqueProjects = Array.from(new Set(rule1Matches.map((d) => d.projectId)));
    const candidates: SedmsMatchCandidate[] = rule1Matches.map((d) => ({
      projectId: d.projectId,
      rule: 1,
      score: 100,
      evidence: `Outbound dispatch ${d.distributionId}`,
      drawingNoKey: d.drawingNoKey,
      distributionId: d.distributionId,
    }));
    if (uniqueProjects.length === 1) {
      const hit = rule1Matches[0];
      return {
        decision: 'auto_link',
        reason: 'Rule 1: Outbound dispatch Message-ID reply match',
        links: [
          {
            id: `lnk_${crypto.randomUUID().slice(0, 10)}`,
            tenantId: message.tenantId,
            messageId: message.messageId,
            projectId: hit.projectId,
            rule: 1,
            drawingNoKey: hit.drawingNoKey ?? null,
            distributionId: hit.distributionId,
            confidence: 1.0,
            createdAt: now,
          },
        ],
        candidates,
      };
    }
    return {
      decision: 'review_queue',
      reason: 'Rule 1 conflict across multiple projects',
      links: [],
      candidates,
    };
  }

  const attachmentText = (message.attachments ?? []).map((a) => a.filename).join(' ');
  const combinedText = `${message.subject || ''}\n${message.bodyText || ''}\n${attachmentText}`;

  // Rule 2: Drawing number flexible match
  const rule2Candidates: SedmsMatchCandidate[] = [];
  const rule2ProjectSet = new Set<string>();
  for (const proj of tenantProjects) {
    for (const dwg of proj.drawingNos ?? []) {
      const pat = createFlexibleDrawingPattern(dwg);
      if (pat && pat.test(combinedText)) {
        const key = canonDrawingNo(dwg);
        rule2ProjectSet.add(proj.projectId);
        if (!rule2Candidates.some((c) => c.projectId === proj.projectId && c.drawingNoKey === key)) {
          rule2Candidates.push({
            projectId: proj.projectId,
            rule: 2,
            score: 90,
            evidence: `Drawing ${dwg}`,
            drawingNoKey: key,
          });
        }
      }
    }
  }

  // Rule 3: Project code / vessel alias DB token match
  const rule3Candidates: SedmsMatchCandidate[] = [];
  const rule3ProjectSet = new Set<string>();
  for (const proj of tenantProjects) {
    const tokens = [proj.projectCode, ...(proj.vesselAliases ?? [])].filter(Boolean);
    for (const token of tokens) {
      if (matchesProjectToken(combinedText, token)) {
        rule3ProjectSet.add(proj.projectId);
        if (!rule3Candidates.some((c) => c.projectId === proj.projectId)) {
          rule3Candidates.push({
            projectId: proj.projectId,
            rule: 3,
            score: 75,
            evidence: `Project/Vessel token ${token}`,
          });
        }
        break;
      }
    }
  }

  // Check Rule 2 & Rule 3 resolution and cross-rule conflict guard
  if (rule2ProjectSet.size > 0) {
    const allCandidates = [...rule2Candidates, ...rule3Candidates];
    if (rule2ProjectSet.size > 1) {
      return {
        decision: 'review_queue',
        reason: 'Rule 2 ambiguity: drawing matched multiple projects',
        links: [],
        candidates: allCandidates,
      };
    }
    const [rule2ProjectId] = Array.from(rule2ProjectSet);
    const hasRule3Conflict = Array.from(rule3ProjectSet).some((pid) => pid !== rule2ProjectId);
    if (hasRule3Conflict) {
      return {
        decision: 'review_queue',
        reason: 'Conflict between Rule 2 drawing project and Rule 3 project/vessel token',
        links: [],
        candidates: allCandidates,
      };
    }
    const primary = rule2Candidates[0];
    return {
      decision: 'auto_link',
      reason: 'Rule 2: Unique project drawing number match',
      links: [
        {
          id: `lnk_${crypto.randomUUID().slice(0, 10)}`,
          tenantId: message.tenantId,
          messageId: message.messageId,
          projectId: rule2ProjectId,
          rule: 2,
          drawingNoKey: primary?.drawingNoKey ?? null,
          confidence: 0.95,
          createdAt: now,
        },
      ],
      candidates: allCandidates,
    };
  }

  if (rule3ProjectSet.size > 0) {
    if (rule3ProjectSet.size === 1) {
      const [projectId] = Array.from(rule3ProjectSet);
      return {
        decision: 'auto_link',
        reason: 'Rule 3: Unique projectCode/vesselAlias token match',
        links: [
          {
            id: `lnk_${crypto.randomUUID().slice(0, 10)}`,
            tenantId: message.tenantId,
            messageId: message.messageId,
            projectId,
            drawingNoKey: null,
            rule: 3,
            confidence: 0.85,
            createdAt: now,
          },
        ],
        candidates: rule3Candidates,
      };
    }
    return {
      decision: 'review_queue',
      reason: 'Rule 3 ambiguity: multiple projects referenced',
      links: [],
      candidates: rule3Candidates,
    };
  }

  // Rule 4: Company domain / email match -> ALWAYS review_queue
  const participants = [message.from, ...(message.to ?? []), ...(message.cc ?? [])].map(extractEmailDomain);
  const rule4Candidates: SedmsMatchCandidate[] = [];
  for (const proj of tenantProjects) {
    const domains = new Set((proj.companyDomains ?? []).map((d) => d.trim().toLowerCase()));
    const emails = new Set((proj.companyEmails ?? []).map((e) => e.trim().toLowerCase()));
    const matched = participants.find(
      (p) => (p.email && emails.has(p.email)) || (p.domain && domains.has(p.domain))
    );
    if (matched) {
      rule4Candidates.push({
        projectId: proj.projectId,
        rule: 4,
        score: 50,
        evidence: `Company contact ${matched.email || matched.domain}`,
      });
    }
  }

  if (rule4Candidates.length > 0) {
    return {
      decision: 'review_queue',
      reason: 'Rule 4: Company domain/email candidate requires human review',
      links: [],
      candidates: rule4Candidates,
    };
  }

  // Rule 5: Unclassified -> honest empty review_queue
  return {
    decision: 'review_queue',
    reason: 'Rule 5: Unclassified message queued for human review',
    links: [],
    candidates: [],
  };
}

export function buildNasMailStorageKey(params: {
  projectRoot: string;
  sentAt: string;
  messageId: string;
  sha256: string;
  filename: string;
}): string {
  const safeRoot = String(params.projectRoot || 'UNASSIGNED').replace(/[^A-Za-z0-9._-]/g, '_');
  const yyyymmMatch = String(params.sentAt || '').match(/^(\d{4})-(\d{2})/);
  const yyyymm = yyyymmMatch ? `${yyyymmMatch[1]}${yyyymmMatch[2]}` : '000000';
  const safeMsgId = String(params.messageId || 'msg').replace(/[^A-Za-z0-9._-]/g, '_');
  const sha12 = String(params.sha256 || '000000000000').slice(0, 12).toLowerCase();
  const safeFilename = String(params.filename || 'file.bin').replace(/[^A-Za-z0-9._-]/g, '_');
  return `${safeRoot}/_mail/${yyyymm}/${safeMsgId}/${sha12}_${safeFilename}`;
}

export function signSedmsIngestPayload(params: {
  tenantId: string;
  timestampSec: number;
  rawBody: string;
  tenantSecret: string;
}): string {
  const signingInput = `${params.tenantId}.${params.timestampSec}.${params.rawBody}`;
  return crypto.createHmac('sha256', params.tenantSecret).update(signingInput).digest('hex');
}

export function verifySedmsIngestSignature(params: {
  tenantId?: string;
  timestampHeader?: string;
  signatureHeader?: string;
  rawBody: string;
  tenantSecrets: Record<string, string>;
  nowSec?: number;
  maxSkewSec?: number;
}): { ok: boolean; statusCode: 200 | 401 | 403; error?: string } {
  const tenantId = (params.tenantId ?? '').trim();
  const tsRaw = (params.timestampHeader ?? '').trim();
  const sig = (params.signatureHeader ?? '').trim();

  if (!tenantId || !tsRaw || !sig) {
    return { ok: false, statusCode: 401, error: '필수 HMAC 인증 헤더(X-SEDMS-Tenant/Timestamp/Signature)가 누락되었습니다.' };
  }

  const tsSec = Number.parseInt(tsRaw, 10);
  const nowSec = params.nowSec ?? Math.floor(Date.now() / 1000);
  const maxSkewSec = params.maxSkewSec ?? 300;
  if (!Number.isFinite(tsSec) || Math.abs(nowSec - tsSec) > maxSkewSec) {
    return { ok: false, statusCode: 401, error: 'HMAC 타임스탬프가 허용 오차 범위를 초과했습니다.' };
  }

  const secret = params.tenantSecrets[tenantId];
  if (!secret) {
    return { ok: false, statusCode: 403, error: '등록되지 않은 테넌트 시크릿입니다.' };
  }

  const expected = signSedmsIngestPayload({
    tenantId,
    timestampSec: tsSec,
    rawBody: params.rawBody,
    tenantSecret: secret,
  });

  const a = Buffer.from(sig, 'utf8');
  const b = Buffer.from(expected, 'utf8');
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    return { ok: false, statusCode: 403, error: 'SEDMS HMAC 서명이 일치하지 않습니다.' };
  }

  return { ok: true, statusCode: 200 };
}

export function buildSedmsOutboundDispatchHeaders(params: {
  tenantId: string;
  projectId: string;
  distributionId: string;
  domain?: string;
}): {
  messageId: string;
  headers: {
    'Message-ID': string;
    'X-SEDMS-Tenant': string;
    'X-SEDMS-Project': string;
    'X-SEDMS-Distribution': string;
  };
} {
  const domain = (params.domain || 'sedms.local').replace(/[^A-Za-z0-9.-]/g, '');
  const safeDist = params.distributionId.replace(/[^A-Za-z0-9._-]/g, '-');
  const messageId = `<sedms-${safeDist}-${crypto.randomUUID().slice(0, 8)}@${domain}>`;
  return {
    messageId,
    headers: {
      'Message-ID': messageId,
      'X-SEDMS-Tenant': params.tenantId,
      'X-SEDMS-Project': params.projectId,
      'X-SEDMS-Distribution': params.distributionId,
    },
  };
}

export class SedmsMailHubStore {
  private projects: SedmsProjectCatalogItem[] = [];
  private outboundDispatches: SedmsOutboundDispatchRef[] = [];
  private messages: SedmsStoredMessage[] = [];
  private reviewQueue: SedmsReviewQueueItem[] = [];
  private attachmentShaIndex = new Map<string, string>();

  setCatalog(projects: SedmsProjectCatalogItem[], outboundDispatches: SedmsOutboundDispatchRef[]): void {
    this.projects = projects.map((p) => ({
      ...p,
      vesselAliases: [...(p.vesselAliases ?? [])],
      drawingNos: [...(p.drawingNos ?? [])],
      companyDomains: [...(p.companyDomains ?? [])],
      companyEmails: [...(p.companyEmails ?? [])],
      sealedDrawings: { ...(p.sealedDrawings ?? {}) },
    }));
    this.outboundDispatches = outboundDispatches.map((d) => ({ ...d }));
  }

  getProjects(tenantId: string): SedmsProjectCatalogItem[] {
    return this.projects.filter((p) => p.tenantId === tenantId);
  }

  registerOutboundDispatch(ref: SedmsOutboundDispatchRef): void {
    this.outboundDispatches.push({ ...ref });
  }

  ingestMessage(payload: SedmsIngestMessagePayload): {
    duplicate: boolean;
    dedupedAttachmentCount: number;
    stored: SedmsStoredMessage;
    reviewQueueItem?: SedmsReviewQueueItem;
  } {
    const existing = this.messages.find(
      (m) => m.tenantId === payload.tenantId && m.messageId === payload.messageId
    );
    if (existing) {
      const existingQueue = this.reviewQueue.find(
        (q) => q.tenantId === payload.tenantId && q.messageId === payload.messageId
      );
      return {
        duplicate: true,
        dedupedAttachmentCount: 0,
        stored: existing,
        reviewQueueItem: existingQueue,
      };
    }

    const match = matchMailDeterministically({
      message: payload,
      projects: this.projects,
      outboundDispatches: this.outboundDispatches,
    });

    const projectRoot = match.links[0]?.projectId || 'UNASSIGNED';
    let dedupedAttachmentCount = 0;

    const attachments: SedmsStoredAttachment[] = (payload.attachments ?? []).map((att, idx) => {
      const dedupKey = `${payload.tenantId}::${projectRoot}::${att.sha256.toLowerCase()}`;
      const existingStorage = this.attachmentShaIndex.get(dedupKey);
      let storageKey: string;
      if (existingStorage) {
        storageKey = existingStorage;
        dedupedAttachmentCount += 1;
      } else {
        storageKey = buildNasMailStorageKey({
          projectRoot,
          sentAt: payload.sentAt,
          messageId: payload.messageId,
          sha256: att.sha256,
          filename: att.filename,
        });
        this.attachmentShaIndex.set(dedupKey, storageKey);
      }
      return {
        id: `satt_${crypto.randomUUID().slice(0, 10)}_${idx}`,
        filename: att.filename,
        mime: att.mime,
        size: att.size,
        sha256: att.sha256,
        storageKey,
      };
    });

    const stored: SedmsStoredMessage = {
      id: `smsg_${crypto.randomUUID().slice(0, 10)}`,
      tenantId: payload.tenantId,
      accountId: payload.accountId,
      messageId: payload.messageId,
      direction: payload.direction,
      from: payload.from,
      to: [...(payload.to ?? [])],
      cc: [...(payload.cc ?? [])],
      subject: payload.subject,
      bodyText: payload.bodyText,
      sentAt: payload.sentAt,
      inReplyTo: payload.inReplyTo,
      references: [...(payload.references ?? [])],
      attachments,
      links: match.links,
      ingestedAt: new Date().toISOString(),
    };

    this.messages.unshift(stored);

    let reviewQueueItem: SedmsReviewQueueItem | undefined;
    if (match.decision === 'review_queue') {
      reviewQueueItem = {
        id: `rq_${crypto.randomUUID().slice(0, 10)}`,
        tenantId: payload.tenantId,
        messageId: payload.messageId,
        subject: payload.subject,
        from: payload.from,
        sentAt: payload.sentAt,
        reason: match.reason,
        candidates: match.candidates,
        state: 'open',
      };
      this.reviewQueue.unshift(reviewQueueItem);
    }

    return {
      duplicate: false,
      dedupedAttachmentCount,
      stored,
      reviewQueueItem,
    };
  }

  listProjectMail(tenantId: string, projectId: string): SedmsStoredMessage[] {
    return this.messages.filter(
      (m) => m.tenantId === tenantId && m.links.some((l) => l.projectId === projectId)
    );
  }

  getMessageById(tenantId: string, id: string): SedmsStoredMessage | null {
    return (
      this.messages.find((m) => m.tenantId === tenantId && (m.id === id || m.messageId === id)) ?? null
    );
  }

  listReviewQueue(tenantId: string, state: 'open' | 'all' = 'open'): SedmsReviewQueueItem[] {
    return this.reviewQueue.filter(
      (q) => q.tenantId === tenantId && (state === 'all' || q.state === 'open')
    );
  }

  resolveReviewItem(params: {
    tenantId: string;
    messageId: string;
    actorSub: string;
    action: 'confirmed' | 'changed' | 'rejected';
    projectId?: string;
    drawingNoKey?: string | null;
  }): { ok: boolean; statusCode: number; error?: string; item?: SedmsReviewQueueItem } {
    const item = this.reviewQueue.find(
      (q) => q.tenantId === params.tenantId && q.messageId === params.messageId
    );
    if (!item) {
      return { ok: false, statusCode: 404, error: '대기열 항목을 찾을 수 없습니다.' };
    }

    item.state = params.action;
    item.resolvedBy = params.actorSub;
    item.resolvedAt = new Date().toISOString();

    if ((params.action === 'confirmed' || params.action === 'changed') && params.projectId) {
      const msg = this.messages.find(
        (m) => m.tenantId === params.tenantId && m.messageId === params.messageId
      );
      if (msg) {
        msg.links.push({
          id: `lnk_${crypto.randomUUID().slice(0, 10)}`,
          tenantId: params.tenantId,
          messageId: params.messageId,
          projectId: params.projectId,
          rule: 5,
          drawingNoKey: params.drawingNoKey ?? null,
          confidence: 1.0,
          createdAt: new Date().toISOString(),
          confirmedBy: params.actorSub,
        });
      }
    }

    return { ok: true, statusCode: 200, item };
  }

  attachMailFileToDrawing(params: {
    tenantId: string;
    messageId: string;
    attachmentId: string;
    projectId: string;
    drawingNo: string;
    proposedRevision: number;
    actorSub: string;
  }): { ok: boolean; statusCode: number; acceptedRevision?: number; error?: string } {
    const msg = this.messages.find(
      (m) => m.tenantId === params.tenantId && (m.messageId === params.messageId || m.id === params.messageId)
    );
    if (!msg) {
      return { ok: false, statusCode: 404, error: '해당 메일을 찾을 수 없습니다.' };
    }
    const att = msg.attachments.find((a) => a.id === params.attachmentId);
    if (!att) {
      return { ok: false, statusCode: 404, error: '해당 첨부파일을 찾을 수 없습니다.' };
    }
    const proj = this.projects.find(
      (p) => p.tenantId === params.tenantId && p.projectId === params.projectId
    );
    if (!proj) {
      return { ok: false, statusCode: 404, error: '해당 프로젝트를 찾을 수 없습니다.' };
    }

    const canonKey = canonDrawingNo(params.drawingNo);
    const sealed = proj.sealedDrawings ?? {};
    const currentSealed = sealed[canonKey] ?? -1;

    if (!Number.isFinite(params.proposedRevision) || params.proposedRevision <= currentSealed) {
      return {
        ok: false,
        statusCode: 409,
        error: `봉인 역행 차단(S6): 도면 ${params.drawingNo}는 Rev.${currentSealed}까지 봉인되어 있어 Rev.${params.proposedRevision}로 덮어쓸 수 없습니다.`,
      };
    }

    proj.sealedDrawings = { ...sealed, [canonKey]: params.proposedRevision };
    return {
      ok: true,
      statusCode: 200,
      acceptedRevision: params.proposedRevision,
    };
  }
}
