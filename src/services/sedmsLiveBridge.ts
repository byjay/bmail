/**
 * SEDMS Live Integration Bridge (`https://sedms.app` / `E:\code-project\SEDMS_G`)
 *
 * - Probes SEDMS live health (`GET /api/health`), syncs verified Outlook contacts
 *   (`POST /api/recipients/registry/import/preview` & `commit`), and runs local
 *   deterministic 5-tier drawing/project classification across Bmail emails.
 * - Optional SEDMS bearer token is stored encrypted at rest (AES-256-GCM) in
 *   `<dataDir>/sedms-bridge.enc` using `<dataDir>/.sedms-bridge.key` and NEVER leaves the server.
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { MailItem } from '../types/mail';
import type { SedmsProjectCatalogItem } from '../types/sgwSedms';
import type { OutlookContactRecord } from './outlookLocalBridge';
import { extractShipbuildingDrawingRefs } from '../utils/outlookWorkspaceModel';
import { canonDrawingNo, matchMailDeterministically } from '../utils/sedmsMailHub';

export const SEDMS_DEFAULT_BASE_URL = 'https://sedms.app';
const SEDMS_TIMEOUT_MS = 10_000;

export interface SedmsConfigPayload {
  sedmsBaseUrl: string;
  bearerToken?: string;
  tenantId: string;
  savedAt: string;
}

interface EncryptedEnvelope {
  v: 1;
  iv: string;
  tag: string;
  data: string;
}

export class SedmsBridgeError extends Error {
  status: number;
  constructor(message: string, status = 500) {
    super(message);
    this.name = 'SedmsBridgeError';
    this.status = status;
  }
}

export function normalizeSedmsBaseUrl(raw?: string | null): string {
  const candidate = String(raw || '').trim() || SEDMS_DEFAULT_BASE_URL;
  let parsed: URL;
  try {
    parsed = new URL(candidate);
  } catch {
    throw new SedmsBridgeError('유효한 SEDMS 서버 URL이 아닙니다.', 400);
  }
  const isHttps = parsed.protocol === 'https:';
  const isLocalHttp =
    parsed.protocol === 'http:' &&
    (parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost');
  if (!isHttps && !isLocalHttp) {
    throw new SedmsBridgeError('SEDMS 연결은 HTTPS(또는 로컬 127.0.0.1)만 허용됩니다.', 400);
  }
  return parsed.origin;
}

export interface SedmsLocalScanItem {
  emailId: string;
  date: string;
  direction: 'in' | 'out';
  decision: 'auto_link' | 'review_queue';
  matchedProject: string | null;
  rule: number | null;
  reason: string;
  drawingRefs: string[];
  canonKeys: string[];
  attachmentCount: number;
}

export interface SedmsLocalScanReport {
  scannedAt: string;
  tenantId: string;
  totalScanned: number;
  drawingMatchedCount: number;
  autoLinkedCount: number;
  reviewQueueCount: number;
  topProjects: Array<{ projectCode: string; mailCount: number; drawingCount: number }>;
  items: SedmsLocalScanItem[];
}

const PROJECT_CODE_TOKEN_RE = /\b(22K[-_]?FD|6\.?1K[-_]?MPC|KSIT|[HSN]\d{4})\b/gi;

export class SedmsLiveBridge {
  private readonly dataDir: string;
  private readonly keyPath: string;
  private readonly configPath: string;
  private readonly defaultBaseUrl: string;
  private cachedConfig: SedmsConfigPayload | null = null;

  constructor(dataDir: string, defaultBaseUrl?: string) {
    this.dataDir = path.resolve(dataDir);
    this.keyPath = path.join(this.dataDir, '.sedms-bridge.key');
    this.configPath = path.join(this.dataDir, 'sedms-bridge.enc');
    this.defaultBaseUrl = normalizeSedmsBaseUrl(
      defaultBaseUrl || process.env.SEDMS_BASE_URL || SEDMS_DEFAULT_BASE_URL
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

  private loadConfig(): SedmsConfigPayload {
    if (this.cachedConfig) return this.cachedConfig;
    const fallback: SedmsConfigPayload = {
      sedmsBaseUrl: this.defaultBaseUrl,
      tenantId: 'TENANT-DEFAULT',
      savedAt: '',
    };
    if (!fs.existsSync(this.configPath) || !fs.existsSync(this.keyPath)) {
      return fallback;
    }
    try {
      const key = this.getOrCreateKey();
      const raw = JSON.parse(fs.readFileSync(this.configPath, 'utf8')) as EncryptedEnvelope;
      if (!raw || raw.v !== 1 || !raw.iv || !raw.tag || !raw.data) return fallback;
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
      const parsed = JSON.parse(plain) as SedmsConfigPayload;
      this.cachedConfig = {
        sedmsBaseUrl: normalizeSedmsBaseUrl(parsed.sedmsBaseUrl || this.defaultBaseUrl),
        bearerToken: parsed.bearerToken?.trim() || undefined,
        tenantId: parsed.tenantId?.trim() || 'TENANT-DEFAULT',
        savedAt: parsed.savedAt || '',
      };
      return this.cachedConfig;
    } catch {
      return fallback;
    }
  }

  saveConfig(params: {
    sedmsBaseUrl?: string;
    bearerToken?: string;
    tenantId?: string;
  }): ReturnType<SedmsLiveBridge['getStatus']> {
    const prev = this.loadConfig();
    const nextPayload: SedmsConfigPayload = {
      sedmsBaseUrl: normalizeSedmsBaseUrl(params.sedmsBaseUrl || prev.sedmsBaseUrl),
      bearerToken:
        params.bearerToken !== undefined
          ? params.bearerToken.trim() || undefined
          : prev.bearerToken,
      tenantId: (params.tenantId || prev.tenantId || 'TENANT-DEFAULT').trim(),
      savedAt: new Date().toISOString(),
    };

    const key = this.getOrCreateKey();
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
    const encrypted = Buffer.concat([
      cipher.update(JSON.stringify(nextPayload), 'utf8'),
      cipher.final(),
    ]);
    const tag = cipher.getAuthTag();
    const env: EncryptedEnvelope = {
      v: 1,
      iv: iv.toString('base64'),
      tag: tag.toString('base64'),
      data: encrypted.toString('base64'),
    };
    fs.writeFileSync(this.configPath, JSON.stringify(env), { encoding: 'utf8', mode: 0o600 });
    this.cachedConfig = nextPayload;
    return this.getStatus();
  }

  clearToken(): ReturnType<SedmsLiveBridge['getStatus']> {
    const prev = this.loadConfig();
    return this.saveConfig({
      sedmsBaseUrl: prev.sedmsBaseUrl,
      tenantId: prev.tenantId,
      bearerToken: '',
    });
  }

  getStatus(): {
    sedmsBaseUrl: string;
    configured: boolean;
    hasToken: boolean;
    tenantId: string;
    savedAt?: string;
  } {
    const cfg = this.loadConfig();
    return {
      sedmsBaseUrl: cfg.sedmsBaseUrl,
      configured: Boolean(cfg.sedmsBaseUrl),
      hasToken: Boolean(cfg.bearerToken),
      tenantId: cfg.tenantId,
      ...(cfg.savedAt ? { savedAt: cfg.savedAt } : {}),
    };
  }

  async checkHealth(): Promise<{
    ok: boolean;
    status: number;
    latencyMs: number;
    sedmsBaseUrl: string;
    error?: string;
  }> {
    const cfg = this.loadConfig();
    const started = Date.now();
    try {
      const res = await fetch(`${cfg.sedmsBaseUrl}/api/health`, {
        headers: { Accept: 'application/json' },
        signal: AbortSignal.timeout(6000),
      });
      return {
        ok: res.ok,
        status: res.status,
        latencyMs: Date.now() - started,
        sedmsBaseUrl: cfg.sedmsBaseUrl,
      };
    } catch (err) {
      return {
        ok: false,
        status: 0,
        latencyMs: Date.now() - started,
        sedmsBaseUrl: cfg.sedmsBaseUrl,
        error: err instanceof Error ? err.message : 'SEDMS 연결 확인 실패',
      };
    }
  }

  private async postJson<T>(endpointPath: string, body: unknown): Promise<T> {
    const cfg = this.loadConfig();
    if (!cfg.bearerToken) {
      throw new SedmsBridgeError(
        'SEDMS API 인증 토큰이 필요합니다. SEDMS 관리 토큰을 먼저 등록해 주세요.',
        401
      );
    }
    let res: Response;
    try {
      res = await fetch(`${cfg.sedmsBaseUrl}${endpointPath}`, {
        method: 'POST',
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json; charset=utf-8',
          Authorization: `Bearer ${cfg.bearerToken}`,
        },
        body: JSON.stringify(body),
        signal: AbortSignal.timeout(SEDMS_TIMEOUT_MS),
      });
    } catch (err) {
      throw new SedmsBridgeError(
        `SEDMS 서버(${cfg.sedmsBaseUrl}) 연결 실패: ${err instanceof Error ? err.message : '네트워크 오류'}`,
        502
      );
    }

    let data: Record<string, unknown> = {};
    try {
      data = (await res.json()) as Record<string, unknown>;
    } catch {
      data = {};
    }
    if (!res.ok || data.ok === false) {
      const reason =
        typeof data.reason === 'string'
          ? data.reason
          : typeof data.error === 'string'
            ? data.error
            : `HTTP ${res.status}`;
      throw new SedmsBridgeError(`SEDMS 요청 실패 (${reason})`, res.status || 400);
    }
    return data as T;
  }

  async previewContactsImport(contacts: OutlookContactRecord[]): Promise<Record<string, unknown>> {
    if (!Array.isArray(contacts) || contacts.length === 0) {
      throw new SedmsBridgeError('SEDMS로 보낼 연락처가 없습니다.', 400);
    }
    const rows = contacts.slice(0, 500).map((c) => ({
      email: c.email,
      ...(c.fullName ? { name: c.fullName } : {}),
      ...(c.company ? { company: c.company } : {}),
      ...(c.phone ? { phone: c.phone } : {}),
      source: c.source || 'bmail-outlook',
    }));
    return this.postJson<Record<string, unknown>>('/api/recipients/registry/import/preview', { rows });
  }

  async commitContactsImport(contacts: OutlookContactRecord[]): Promise<Record<string, unknown>> {
    if (!Array.isArray(contacts) || contacts.length === 0) {
      throw new SedmsBridgeError('SEDMS에 반영할 연락처가 없습니다.', 400);
    }
    const rows = contacts.slice(0, 500).map((c) => ({
      email: c.email,
      ...(c.fullName ? { name: c.fullName } : {}),
      ...(c.company ? { company: c.company } : {}),
      ...(c.phone ? { phone: c.phone } : {}),
      source: c.source || 'bmail-outlook',
    }));
    return this.postJson<Record<string, unknown>>('/api/recipients/registry/import/commit', { rows });
  }

  scanLocalEmailsForSedms(
    emails: MailItem[],
    tenantIdOverride?: string,
    projectCatalog?: SedmsProjectCatalogItem[]
  ): SedmsLocalScanReport {
    const cfg = this.loadConfig();
    const tenantId = (tenantIdOverride || cfg.tenantId || 'TENANT-DEFAULT').trim();

    const inferredCodes = new Set<string>();
    const inferredDrawingsByCode = new Map<string, Set<string>>();

    for (const m of emails) {
      const hay = `${m.subject || ''} ${m.folderName || ''} ${m.snippet || ''}`;
      const tokens = hay.match(PROJECT_CODE_TOKEN_RE) || [];
      const folderToken =
        m.folderName && !/(받은\s*편지함|보낸\s*편지함|지운|임시|inbox|sent|trash|drafts)/i.test(m.folderName)
          ? m.folderName.split(/[\\/]/).pop()?.trim() || ''
          : '';
      const primaryCode = tokens[0]
        ? tokens[0].toUpperCase().replace(/\s+/g, '')
        : folderToken.length >= 2 && folderToken.length <= 24
          ? folderToken
          : '';
      if (primaryCode) {
        inferredCodes.add(primaryCode);
        if (!inferredDrawingsByCode.has(primaryCode)) {
          inferredDrawingsByCode.set(primaryCode, new Set());
        }
        const refs = extractShipbuildingDrawingRefs(
          `${m.subject} ${m.attachments?.map((a) => a.filename).join(' ') || ''}`
        );
        for (const r of refs) {
          inferredDrawingsByCode.get(primaryCode)!.add(canonDrawingNo(r.raw));
        }
      }
    }

    const catalog: SedmsProjectCatalogItem[] =
      projectCatalog && projectCatalog.length > 0
        ? projectCatalog
        : Array.from(inferredCodes).map((code) => ({
            tenantId,
            projectId: code,
            projectCode: code,
            vesselAliases: [code],
            drawingNos: Array.from(inferredDrawingsByCode.get(code) || []),
            companyDomains: [],
            companyEmails: [],
          }));

    const projectStats = new Map<string, { mailCount: number; drawings: Set<string> }>();
    const items: SedmsLocalScanItem[] = [];
    let drawingMatchedCount = 0;
    let autoLinkedCount = 0;
    let reviewQueueCount = 0;

    for (const m of emails) {
      const attNames = (m.attachments || []).map((a) => a.filename).join(' ');
      const refs = extractShipbuildingDrawingRefs(
        `${m.subject || ''} ${attNames} ${m.snippet || ''}`
      );
      const drawingRefs = refs.map((r) => r.raw);
      const canonKeys = Array.from(new Set(refs.map((r) => r.canonKey)));
      if (drawingRefs.length > 0) drawingMatchedCount += 1;

      const outcome = matchMailDeterministically({
        message: {
          tenantId,
          accountId: 'local-pst',
          messageId: m.id,
          direction: m.folder === 'sent' ? 'out' : 'in',
          from: m.fromEmail || m.fromName || '',
          to: Array.isArray(m.to) ? m.to : [],
          subject: m.subject || '',
          bodyText: `${m.snippet || ''}\n${(m.body || '').slice(0, 1500)}`,
          sentAt: m.timestamp || m.date || '',
          attachments: (m.attachments || []).map((a) => ({
            filename: a.filename,
            mime: a.mimeType || 'application/octet-stream',
            size: Number(a.sizeBytes || 0),
            sha256: '',
          })),
        },
        projects: catalog,
        outboundDispatches: [],
      });

      const primaryLink = outcome.links[0];
      const topCandidate = outcome.candidates[0];
      const matchedProject = primaryLink?.projectId || topCandidate?.projectId || null;
      const rule = primaryLink?.rule ?? topCandidate?.rule ?? null;

      if (outcome.decision === 'auto_link') autoLinkedCount += 1;
      else reviewQueueCount += 1;

      if (matchedProject) {
        if (!projectStats.has(matchedProject)) {
          projectStats.set(matchedProject, { mailCount: 0, drawings: new Set() });
        }
        const entry = projectStats.get(matchedProject)!;
        entry.mailCount += 1;
        for (const k of canonKeys) entry.drawings.add(k);
      }

      if (matchedProject || drawingRefs.length > 0) {
        items.push({
          emailId: m.id,
          date: m.date,
          direction: m.folder === 'sent' ? 'out' : 'in',
          decision: outcome.decision,
          matchedProject,
          rule,
          reason: outcome.reason,
          drawingRefs,
          canonKeys,
          attachmentCount: m.attachments?.length || 0,
        });
      }
    }

    const topProjects = Array.from(projectStats.entries())
      .map(([projectCode, s]) => ({
        projectCode,
        mailCount: s.mailCount,
        drawingCount: s.drawings.size,
      }))
      .sort((a, b) => b.mailCount - a.mailCount)
      .slice(0, 15);

    return {
      scannedAt: new Date().toISOString(),
      tenantId,
      totalScanned: emails.length,
      drawingMatchedCount,
      autoLinkedCount,
      reviewQueueCount,
      topProjects,
      items: items.slice(0, 50),
    };
  }
}
