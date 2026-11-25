import dotenv from 'dotenv';
import express from 'express';
import multer from 'multer';
import { execFile, spawn, type ChildProcess } from 'node:child_process';
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { runAutonomousMailAgent, sanitizeEmailsForOffline } from './src/services/agentService';
import { analyzeEmailsWithProvider, getProviderStatuses } from './src/services/aiService';
import { BackgroundMailMonitor } from './src/services/backgroundMailMonitor';
import { summarizeNewMailsForNotification } from './src/services/newMailNotifier';
import { buildUnreadPriorityBriefing } from './src/services/unreadBriefingEngine';
import { parseUploadedDocument } from './src/services/documentParser';
import { probeExternalEngine } from './src/services/externalEngineBridge';
import {
  analyzeDocumentsWithAi,
  analyzeMeetingMinutesWithAi,
  transcribeAudioFile,
} from './src/services/multimodalAiService';
import { parseUploadedMailFiles } from './src/services/pstParser';
import {
  connectMailWithAutoHeal,
  detectCompanyPop3Config,
  detectServerEgressInfo,
  fetchImapMessages,
  fetchPop3Messages,
  MAIL_SERVICE_PRESETS,
  MailAuthError,
  type MailReceiveProtocol,
  type MailServicePresetId,
  Pop3SessionStore,
  sanitizeMailError,
  sendSmtpMessage,
  toPublicPop3AccountInfo,
} from './src/services/pop3AutoConnect';
import {
  buildWindowsOutlookInspectorScript,
  parseWindowsOutlookProfileJson,
} from './src/services/outlookLocalBridge';
import { LocalMailDb } from './src/services/localMailDb';
import {
  discoverAndApplyAllOutlookSettings,
  inspectPstFileStates,
  pollLiveOutlookInboxAsync,
  readPstAttachmentBytes,
  readPstMessageRichView,
  recomputeProjectsFromLocalDb,
  sanitizeEmailHtml,
  syncIncrementalLocalOutlook,
  syncLiveOutlookViaCom,
  syncOutlookSettingsOnly,
  type LocalPstFileState,
  type MigrationProgressSummary,
} from './src/services/localOutlookMigrator';
import {
  buildAttachmentExportPlan,
  executeAttachmentExportPlan,
  type AttachmentExportFilterOptions,
} from './src/services/attachmentBatchExporter';
import {
  buildExecutiveDashboardReport,
  generateAiExecutiveNarrative,
  saveExecutiveReportSnapshot,
  type ExecutivePeriodType,
} from './src/services/executiveReportEngine';
import {
  invokeClaudeMessagesApi,
  selectClaudeModelByComplexity,
} from './src/services/claudeModelRouter';
import {
  getGroqKeyPoolStatus,
  invokeGroqStreamingCompletion,
  shouldEscalateToClaude,
} from './src/services/groqKeyRotator';
import {
  configureLlmUsageLedgerPath,
  getLlmUsageSummary,
  resetLlmUsageLedger,
} from './src/services/llmUsageTracker';
import { LocalBackupScheduler } from './src/services/localBackupScheduler';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  extractShipbuildingDrawingRefs,
  type ProjectMailOverrides,
} from './src/utils/outlookWorkspaceModel';
import {
  EMV_KEY_CATALOG,
  EmvVaultService,
  type EmvScope,
} from './src/services/emvVaultService';
import { getHybridCreditStatus } from './src/services/hybridCreditRouter';
import {
  buildSafeSourceArchive,
  buildSafeSourceZipArchive,
} from './src/services/sourceArchiveBuilder';
import type {
  AiProviderId,
  AnalysisOptions,
  AppConfig,
  BlogTone,
  MailItem,
  ParsedDocumentSummary,
} from './src/types/mail';
import {
  buildEnvSecretStatuses,
  buildEnvTemplateText,
  DEFAULT_APP_CONFIG,
  sanitizeAppConfig,
} from './src/utils/appConfig';
import { buildDeterministicWorkLog, exportEmailsToMbox, type ProjectDefinition } from './src/utils/deterministicAnalyzer';
import { createIconPngBuffer } from './src/utils/iconPng';
import {
  buildCommandsHtml,
  buildOutlookManifestXml,
  buildUnifiedManifestJson,
  buildWindowsRegistryCatalog,
  normalizeBaseUrl,
} from './src/utils/outlookManifest';
import { createSgwSedmsRouter } from './src/server/sgwSedmsRouter';
import {
  buildMailShareText,
  SGW_MAX_RAW_ATTACHMENT_BYTES,
  SgwBridgeError,
  SgwChatBridge,
} from './src/services/sgwChatBridge';
import { SedmsBridgeError, SedmsLiveBridge } from './src/services/sedmsLiveBridge';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const APP_ROOT = process.env.BMAIL_APP_ROOT?.trim()
  ? path.resolve(process.env.BMAIL_APP_ROOT.trim())
  : fs.existsSync(path.resolve(__dirname, 'package.json'))
    ? __dirname
    : path.resolve(__dirname, '..');
const DATA_DIR = process.env.BMAIL_DATA_DIR?.trim()
  ? path.resolve(process.env.BMAIL_DATA_DIR.trim())
  : path.resolve(APP_ROOT, 'data');

const rootEnvPath = path.resolve(APP_ROOT, '.env');
if (!process.env.BMAIL_SKIP_ROOT_ENV && fs.existsSync(rootEnvPath)) {
  dotenv.config({ path: rootEnvPath, override: true });
}
const dataEnvPath = path.resolve(DATA_DIR, '.env');
if (dataEnvPath !== rootEnvPath && fs.existsSync(dataEnvPath)) {
  dotenv.config({ path: dataEnvPath, override: true });
}
const explicitEnvPath = process.env.BMAIL_ENV_FILE?.trim();
if (explicitEnvPath && fs.existsSync(explicitEnvPath)) {
  dotenv.config({ path: explicitEnvPath, override: true });
}

const PORT = Number(process.env.BMAIL_DESKTOP_PORT || process.env.PORT || 3000);
const BIND_HOST = process.env.BMAIL_HOST?.trim() || '0.0.0.0';
const isProduction =
  process.env.NODE_ENV === 'production' ||
  (path.basename(__filename) === 'server.mjs' &&
    fs.existsSync(path.resolve(APP_ROOT, 'dist', 'index.html')));
const DEFAULT_CONFIG_FILE_PATH = path.resolve(APP_ROOT, 'config.json');
const CONFIG_FILE_PATH = process.env.BMAIL_DATA_DIR?.trim()
  ? path.resolve(DATA_DIR, 'config.json')
  : DEFAULT_CONFIG_FILE_PATH;

interface MicrosoftSession {
  accessToken: string;
  refreshToken?: string;
  expiresAt: number;
  sessionExpiresAt: number;
  userEmail: string;
  displayName: string;
}

const MS_SESSION_TTL_MS = 30 * 24 * 60 * 60 * 1000;
const MS_SESSION_FILE = path.resolve(DATA_DIR, 'ms-sessions.enc.json');
const msSessions = new Map<string, MicrosoftSession>();
const pop3Sessions = new Pop3SessionStore({ baseDir: DATA_DIR });

function deriveMsSessionKey(): Buffer {
  const secret =
    process.env.EMV_MASTER_KEY?.trim() ||
    process.env.BMAIL_SESSION_KEY?.trim() ||
    `bmail-local-session:${os.hostname()}:${os.userInfo().username}`;
  return crypto.scryptSync(secret, 'bmail-ms-session-salt-v1', 32);
}

function saveMsSessionsToDisk(): void {
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    const payload = JSON.stringify(Array.from(msSessions.entries()));
    const iv = crypto.randomBytes(12);
    const cipher = crypto.createCipheriv('aes-256-gcm', deriveMsSessionKey(), iv);
    const encrypted = Buffer.concat([cipher.update(payload, 'utf8'), cipher.final()]);
    const tag = cipher.getAuthTag();
    fs.writeFileSync(
      MS_SESSION_FILE,
      JSON.stringify({
        version: 1,
        iv: iv.toString('hex'),
        tag: tag.toString('hex'),
        ciphertext: encrypted.toString('hex'),
      }),
      'utf8'
    );
  } catch {
    // Best-effort local persistence
  }
}

function loadMsSessionsFromDisk(): void {
  if (!fs.existsSync(MS_SESSION_FILE)) return;
  try {
    const raw = JSON.parse(fs.readFileSync(MS_SESSION_FILE, 'utf8')) as {
      version?: number;
      iv?: string;
      tag?: string;
      ciphertext?: string;
    };
    if (raw.version !== 1 || !raw.iv || !raw.tag || !raw.ciphertext) return;
    const decipher = crypto.createDecipheriv(
      'aes-256-gcm',
      deriveMsSessionKey(),
      Buffer.from(raw.iv, 'hex')
    );
    decipher.setAuthTag(Buffer.from(raw.tag, 'hex'));
    const plain = Buffer.concat([
      decipher.update(Buffer.from(raw.ciphertext, 'hex')),
      decipher.final(),
    ]).toString('utf8');
    const entries = JSON.parse(plain);
    if (!Array.isArray(entries)) return;
    const now = Date.now();
    for (const pair of entries) {
      if (Array.isArray(pair) && typeof pair[0] === 'string' && pair[1] && typeof pair[1] === 'object') {
        const sess = pair[1] as MicrosoftSession;
        const limit = Number(sess.sessionExpiresAt || sess.expiresAt || 0);
        if (limit > now) {
          msSessions.set(pair[0], { ...sess, sessionExpiresAt: limit });
        }
      }
    }
  } catch {
    // Ignore invalid/different-host encrypted file
  }
}

loadMsSessionsFromDisk();

function isLoopbackRequest(req: express.Request): boolean {
  const host = String(req.hostname || req.headers.host || '').toLowerCase();
  const ip = String(req.ip || req.socket?.remoteAddress || '').toLowerCase();
  return (
    host.startsWith('127.0.0.1') ||
    host.startsWith('localhost') ||
    ip === '127.0.0.1' ||
    ip === '::1' ||
    ip === '::ffff:127.0.0.1'
  );
}
const emvVault = new EmvVaultService({ baseDir: DATA_DIR });
const localDb = new LocalMailDb(DATA_DIR);
configureLlmUsageLedgerPath(path.resolve(DATA_DIR, 'llm-usage-ledger.json'));
const localBackupScheduler = new LocalBackupScheduler(DATA_DIR, localDb, {
  maxRetention: 7,
  intervalHours: 24,
});
localBackupScheduler.startDailySchedule();
const sgwChatBridge = new SgwChatBridge(DATA_DIR);
const sedmsLiveBridge = new SedmsLiveBridge(DATA_DIR);
const backgroundMonitor = new BackgroundMailMonitor({
  localDb,
  pop3Sessions,
  sgwChatBridge,
});
const iconCache = new Map<number, Buffer>();
let migrationProgress: MigrationProgressSummary = {
  status: 'idle',
  startedAt: null,
  completedAt: null,
  storesProcessed: 0,
  totalEmailsIndexed: 0,
  totalContactsIndexed: 0,
  totalFoldersIndexed: 0,
  totalSignaturesIndexed: 0,
  discoveredProjectsCount: 0,
  error: null,
};

function readAppConfig(): AppConfig {
  try {
    const targetPath = fs.existsSync(CONFIG_FILE_PATH)
      ? CONFIG_FILE_PATH
      : DEFAULT_CONFIG_FILE_PATH;
    if (fs.existsSync(targetPath)) {
      const raw = JSON.parse(fs.readFileSync(targetPath, 'utf-8'));
      return sanitizeAppConfig(raw);
    }
  } catch {
    // Fall back to sanitized defaults if config.json is malformed
  }
  return DEFAULT_APP_CONFIG;
}

function writeAppConfig(nextConfig: AppConfig): AppConfig {
  const sanitized = sanitizeAppConfig(nextConfig);
  fs.mkdirSync(path.dirname(CONFIG_FILE_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_FILE_PATH, `${JSON.stringify(sanitized, null, 2)}\n`, 'utf-8');
  return sanitized;
}

function getAppBaseUrl(req: express.Request): string {
  const fromQuery = typeof req.query.baseUrl === 'string' && req.query.baseUrl.trim() ? req.query.baseUrl.trim() : '';
  if (fromQuery) {
    return normalizeBaseUrl(fromQuery);
  }
  const envUrl = process.env.APP_URL?.trim();
  if (envUrl && envUrl !== 'MY_APP_URL') {
    return normalizeBaseUrl(envUrl);
  }
  const proto = req.headers['x-forwarded-proto']?.toString().split(',')[0] || req.protocol || 'http';
  const host = req.headers['x-forwarded-host']?.toString().split(',')[0] || req.headers.host || `localhost:${PORT}`;
  return normalizeBaseUrl(`${proto}://${host}`);
}

function parseCookies(cookieHeader?: string): Record<string, string> {
  const out: Record<string, string> = {};
  if (!cookieHeader) return out;
  for (const part of cookieHeader.split(';')) {
    const idx = part.indexOf('=');
    if (idx === -1) continue;
    const k = part.slice(0, idx).trim();
    const v = part.slice(idx + 1).trim();
    out[k] = decodeURIComponent(v);
  }
  return out;
}

function resolveMicrosoftSessionEntry(
  req: express.Request
): { sid: string; session: MicrosoftSession } | null {
  const now = Date.now();
  const cookies = parseCookies(req.headers.cookie);
  const sid = cookies.ms_session_id;
  if (sid) {
    const session = msSessions.get(sid);
    if (session) {
      if (now <= (session.sessionExpiresAt || session.expiresAt) || Boolean(session.refreshToken)) {
        return { sid, session };
      }
      msSessions.delete(sid);
      saveMsSessionsToDisk();
    }
  }
  // Auto-restore persisted local desktop login across port changes / restarts
  if (isLoopbackRequest(req)) {
    let best: { sid: string; session: MicrosoftSession } | null = null;
    for (const [entrySid, sess] of msSessions.entries()) {
      const ttlLimit = sess.sessionExpiresAt || sess.expiresAt;
      if (now > ttlLimit && !sess.refreshToken) {
        msSessions.delete(entrySid);
        continue;
      }
      if (!best || sess.expiresAt > best.session.expiresAt) {
        best = { sid: entrySid, session: sess };
      }
    }
    if (best) return best;
  }
  return null;
}

function getMicrosoftSession(req: express.Request): MicrosoftSession | null {
  const entry = resolveMicrosoftSessionEntry(req);
  if (!entry) return null;
  if (!entry.session.refreshToken && Date.now() > entry.session.expiresAt) {
    return null;
  }
  return entry.session;
}

async function ensureFreshMicrosoftSession(req: express.Request): Promise<MicrosoftSession | null> {
  const entry = resolveMicrosoftSessionEntry(req);
  if (!entry) return null;
  const { sid, session } = entry;
  if (Date.now() <= session.expiresAt - 120_000) {
    return session;
  }
  if (!session.refreshToken) {
    return Date.now() <= session.expiresAt ? session : null;
  }
  const clientId = process.env.MICROSOFT_CLIENT_ID?.trim();
  const clientSecret = process.env.MICROSOFT_CLIENT_SECRET?.trim();
  const tenant = process.env.MICROSOFT_TENANT_ID?.trim() || 'common';
  if (!clientId || !clientSecret) return session;

  try {
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          refresh_token: session.refreshToken,
          grant_type: 'refresh_token',
        }),
      }
    );
    if (!tokenRes.ok) {
      return Date.now() <= session.expiresAt ? session : null;
    }
    const tokenData = (await tokenRes.json()) as {
      access_token?: string;
      refresh_token?: string;
      expires_in?: number;
    };
    if (tokenData.access_token) {
      session.accessToken = tokenData.access_token;
      if (tokenData.refresh_token) session.refreshToken = tokenData.refresh_token;
      session.expiresAt = Date.now() + (tokenData.expires_in || 3600) * 1000;
      session.sessionExpiresAt = Date.now() + MS_SESSION_TTL_MS;
      msSessions.set(sid, session);
      saveMsSessionsToDisk();
    }
    return session;
  } catch {
    return Date.now() <= session.expiresAt ? session : null;
  }
}

function extractBearerToken(req: express.Request): string {
  const auth = req.headers.authorization || '';
  const m = auth.match(/^Bearer\s+([A-Za-z0-9_-]{32,128})$/i);
  return m ? m[1] : '';
}

/**
 * Resolves the caller's employee identity strictly from server-verified sessions
 * (Microsoft OAuth session cookie or POP3 Bearer session token).
 * Never trusts spoofable raw client-supplied user IDs to prevent cross-employee key access.
 */
function resolveVerifiedEmployeeId(req: express.Request): string | null {
  const msSession = getMicrosoftSession(req);
  if (msSession?.userEmail) return msSession.userEmail;
  const bearer = extractBearerToken(req);
  const pop3Session = pop3Sessions.get(bearer);
  if (pop3Session?.email) return pop3Session.email;
  return null;
}

function stripHtmlToPlainText(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

async function startServer() {
  const app = express();
  app.set('trust proxy', 1);
  app.use(express.json({ limit: '10mb' }));

  // Allow Outlook Desktop WebView2 / OWA / SGW iframes and CORS manifest loading
  app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader(
      'Access-Control-Allow-Headers',
      'Content-Type, Authorization, X-Config-Password, X-SEDMS-Tenant, X-SEDMS-Timestamp, X-SEDMS-Signature'
    );
    const frameAncestors = process.env.SGW_FRAME_ANCESTORS?.trim();
    if (frameAncestors) {
      res.setHeader('Content-Security-Policy', `frame-ancestors 'self' ${frameAncestors}`);
    }
    if (req.method === 'OPTIONS') {
      res.status(204).end();
      return;
    }
    next();
  });

  // --- Outlook Add-in Real Assets & Manifests ---
  app.get('/assets/icon-:size.png', (req, res) => {
    const sizeNum = Number(req.params.size);
    const allowed = [16, 32, 64, 80, 128];
    const targetSize = allowed.includes(sizeNum) ? sizeNum : 64;
    let png = iconCache.get(targetSize);
    if (!png) {
      png = createIconPngBuffer(targetSize);
      iconCache.set(targetSize, png);
    }
    res.setHeader('Content-Type', 'image/png');
    res.setHeader('Cache-Control', 'public, max-age=86400');
    res.send(png);
  });

  app.get('/manifest.xml', (req, res) => {
    const cfg = readAppConfig();
    const baseUrl = getAppBaseUrl(req);
    const displayName =
      typeof req.query.name === 'string' && req.query.name.trim()
        ? req.query.name.trim()
        : cfg.plugin.displayName;
    const xml = buildOutlookManifestXml({
      baseUrl,
      displayName,
      supportsPinning: cfg.plugin.supportsPinning,
      supportsSharedFolders: cfg.plugin.supportsSharedFolders,
    });
    res.setHeader('Content-Type', 'application/xml; charset=utf-8');
    if (req.query.download === '1') {
      res.setHeader('Content-Disposition', 'attachment; filename="maillog-outlook-addin-manifest.xml"');
    }
    res.send(xml);
  });

  app.get('/manifest.json', (req, res) => {
    const cfg = readAppConfig();
    const baseUrl = getAppBaseUrl(req);
    const displayName =
      typeof req.query.name === 'string' && req.query.name.trim()
        ? req.query.name.trim()
        : cfg.plugin.displayName;
    const json = buildUnifiedManifestJson({
      baseUrl,
      displayName,
      supportsPinning: cfg.plugin.supportsPinning,
      supportsSharedFolders: cfg.plugin.supportsSharedFolders,
    });
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    if (req.query.download === '1') {
      res.setHeader('Content-Disposition', 'attachment; filename="maillog-m365-unified-manifest.json"');
    }
    res.send(json);
  });

  app.get('/api/outlook/trusted-catalog.reg', (req, res) => {
    const cfg = readAppConfig();
    const folderPath =
      typeof req.query.path === 'string' && req.query.path.trim()
        ? req.query.path.trim()
        : cfg.plugin.trustedCatalogPath;
    const reg = buildWindowsRegistryCatalog(folderPath);
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="maillog-trusted-catalog.reg"');
    res.send(reg);
  });

  app.get('/commands.html', (_req, res) => {
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.send(buildCommandsHtml());
  });

  // --- Separated config.json & .env Secret Readiness Endpoints ---
  app.get('/api/config', (req, res) => {
    const cfg = readAppConfig();
    const baseUrl = getAppBaseUrl(req);
    const verifiedEmp = resolveVerifiedEmployeeId(req);
    const runtimeEnv = emvVault.resolveRuntimeEnv(verifiedEmp);
    const providers = getProviderStatuses(cfg, runtimeEnv);
    const envSecrets = buildEnvSecretStatuses({ ...process.env, ...runtimeEnv });
    const session = getMicrosoftSession(req);
    const msConfigured = Boolean(
      (runtimeEnv.MICROSOFT_CLIENT_ID || process.env.MICROSOFT_CLIENT_ID)?.trim() &&
      (runtimeEnv.MICROSOFT_CLIENT_SECRET || process.env.MICROSOFT_CLIENT_SECRET)?.trim()
    );

    res.json({
      appUrl: baseUrl,
      manifestXmlUrl: `${baseUrl}/manifest.xml`,
      manifestJsonUrl: `${baseUrl}/manifest.json`,
      taskpaneUrl: `${baseUrl}/taskpane`,
      oauthCallbackUrl: `${baseUrl}/auth/microsoft/callback`,
      config: cfg,
      envSecrets,
      configPasswordRequired: Boolean(process.env.CONFIG_ADMIN_PASSWORD?.trim()),
      providers,
      microsoftAuth: {
        configured: msConfigured,
        connected: Boolean(session),
        userEmail: session?.userEmail || null,
        displayName: session?.displayName || null,
      },
    });
  });

  app.put('/api/config', (req, res) => {
    try {
      const adminPassword = process.env.CONFIG_ADMIN_PASSWORD?.trim();
      if (adminPassword) {
        const provided = String(req.headers['x-config-password'] || '');
        if (provided !== adminPassword) {
          res.status(401).json({
            error: '.env의 CONFIG_ADMIN_PASSWORD가 일치하지 않아 config.json을 수정할 수 없습니다.',
          });
          return;
        }
      }

      const saved = writeAppConfig(req.body as AppConfig);
      res.json({
        config: saved,
        providers: getProviderStatuses(saved),
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : 'config.json 저장 중 오류가 발생했습니다.';
      res.status(400).json({ error: message });
    }
  });

  app.get('/api/config/download', (_req, res) => {
    const cfg = readAppConfig();
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="config.json"');
    res.send(`${JSON.stringify(cfg, null, 2)}\n`);
  });

  app.get('/api/env-template/download', (req, res) => {
    const baseUrl = getAppBaseUrl(req);
    const template = buildEnvTemplateText(baseUrl);
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename=".env.example"');
    res.send(template);
  });

  // --- Autonomous AI Agent SSE Streaming Pipeline ---
  app.post('/api/agent/run', async (req, res) => {
    const cfg = readAppConfig();
    const body = req.body as {
      emails?: MailItem[];
      options?: AnalysisOptions;
      agentGoal?: string;
      autoPickLatestDateIfEmpty?: boolean;
    };

    res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache, no-transform');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders?.();

    const sendEvent = (event: string, payload: unknown) => {
      res.write(`event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`);
    };

    try {
      const emails = Array.isArray(body.emails) ? body.emails : [];
      const fallbackOptions: AnalysisOptions = {
        dateMode: cfg.agent.defaultDateMode,
        singleDate: emails[0]?.date || new Date().toISOString().slice(0, 10),
        startDate: emails[emails.length - 1]?.date || new Date().toISOString().slice(0, 10),
        endDate: emails[0]?.date || new Date().toISOString().slice(0, 10),
        scope: cfg.agent.defaultScope,
        selectedIds: [],
        provider: cfg.ai.defaultProvider,
        blogTone: cfg.agent.defaultBlogTone,
        maskSensitiveInfo: cfg.agent.autoMaskSensitiveInfo,
        includeActionItems: cfg.agent.includeActionItems,
        customFocus: '',
      };

      const options: AnalysisOptions = body.options
        ? { ...fallbackOptions, ...body.options }
        : fallbackOptions;

      const verifiedEmp = resolveVerifiedEmployeeId(req);
      const runtimeEnv = emvVault.resolveRuntimeEnv(verifiedEmp);

      const localProjects = localDb.isReady()
        ? localDb.getState<ProjectDefinition[]>('projects') || []
        : [];

      const result = await runAutonomousMailAgent({
        emails,
        options,
        config: cfg,
        projects: localProjects,
        agentGoal: typeof body.agentGoal === 'string' ? body.agentGoal.slice(0, 400) : '',
        autoPickLatestDateIfEmpty: body.autoPickLatestDateIfEmpty !== false,
        envOverrides: runtimeEnv,
        onStepUpdate: (steps) => {
          sendEvent('steps', { steps });
        },
      });

      sendEvent('result', result);
      res.end();
    } catch (error) {
      const message =
        error instanceof Error ? error.message : 'AI 에이전트 실행 중 오류가 발생했습니다.';
      sendEvent('error', { error: message });
      res.end();
    }
  });

  // --- Direct Multi-Provider AI Email Analysis Endpoint ---
  app.post('/api/analyze', async (req, res) => {
    const cfg = readAppConfig();
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      emails?: MailItem[];
      provider?: AiProviderId;
      periodLabel?: string;
      scopeLabel?: string;
      blogTone?: BlogTone;
      maskSensitiveInfo?: boolean;
      includeActionItems?: boolean;
      customFocus?: string;
      strictCloudOnly?: boolean;
    };

    try {
      if (!Array.isArray(body.emails) || body.emails.length === 0) {
        res.status(400).json({ error: '분석할 실제 메일이 최소 1건 이상 필요합니다.' });
        return;
      }

      const validProviders: AiProviderId[] = ['groq', 'claude', 'openai', 'builtin'];
      const provider: AiProviderId =
        body.provider && validProviders.includes(body.provider) ? body.provider : cfg.ai.defaultProvider;

      const validTones: BlogTone[] = ['insight', 'tech-dev', 'essay', 'newsletter'];
      const blogTone: BlogTone =
        body.blogTone && validTones.includes(body.blogTone) ? body.blogTone : cfg.agent.defaultBlogTone;

      const verifiedEmp = resolveVerifiedEmployeeId(req);
      const runtimeEnv = emvVault.resolveRuntimeEnv(verifiedEmp);

      const result = await analyzeEmailsWithProvider({
        emails: body.emails.slice(0, cfg.ai.maxEmailsPerRun),
        provider,
        periodLabel: body.periodLabel || '지정 기간',
        scopeLabel: body.scopeLabel || '선택 메일',
        blogTone,
        maskSensitiveInfo: body.maskSensitiveInfo !== false,
        includeActionItems: body.includeActionItems !== false,
        customFocus: typeof body.customFocus === 'string' ? body.customFocus.slice(0, 500) : '',
        modelOverrides: cfg.ai.models,
        envOverrides: runtimeEnv,
      });

      res.json(result);
    } catch (error) {
      if (
        cfg.ai.fallbackToBuiltIn !== false &&
        body.strictCloudOnly !== true &&
        Array.isArray(body.emails) &&
        body.emails.length > 0
      ) {
        const localProjects = localDb.isReady()
          ? localDb.getState<ProjectDefinition[]>('projects') || []
          : [];
        const safeEmails = sanitizeEmailsForOffline(
          body.emails.slice(0, cfg.ai.maxEmailsPerRun),
          body.maskSensitiveInfo !== false
        );
        res.json(
          buildDeterministicWorkLog({
            emails: safeEmails,
            projects: localProjects,
            periodLabel: body.periodLabel || '지정 기간',
            templateName: 'standard',
          })
        );
        return;
      }
      const message = error instanceof Error ? error.message : '메일 분석 중 알 수 없는 오류가 발생했습니다.';
      const isConfigError = message.includes('설정되지 않았습니다');
      res.status(isConfigError ? 400 : 500).json({
        error: message,
        code: isConfigError ? 'PROVIDER_NOT_CONFIGURED' : 'ANALYSIS_FAILED',
      });
    }
  });

  // --- Outlook .pst / .ost / .eml / .mbox File Import (MailCraft-style direct archive ingestion) ---
  const upload = multer({
    dest: path.join(os.tmpdir(), 'maillog-pst-uploads'),
    limits: {
      fileSize: 200 * 1024 * 1024,
      files: 30,
    },
  });

  app.post('/api/pst/import', upload.array('files', 30), (req, res) => {
    const uploaded = (req.files as Express.Multer.File[] | undefined) ?? [];
    if (uploaded.length === 0) {
      res.status(400).json({ error: '업로드된 .pst / .ost / .eml / .mbox / .msg 파일이 없습니다.' });
      return;
    }

    try {
      const result = parseUploadedMailFiles(
        uploaded.map((f) => ({ path: f.path, originalname: f.originalname }))
      );
      res.json(result);
    } catch (error) {
      const message =
        error instanceof Error
          ? error.message
          : 'PST 또는 메일 파일을 읽는 중 오류가 발생했습니다. 암호화된 PST 파일인지 확인해 주세요.';
      res.status(400).json({ error: message });
    } finally {
      for (const file of uploaded) {
        try {
          if (fs.existsSync(file.path)) {
            fs.unlinkSync(file.path);
          }
        } catch {
          // ignore cleanup failure
        }
      }
    }
  });

  // --- External PDF / Excel Project Folder Probe ---
  app.post('/api/external-engines/probe', (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      pdfEngineDir?: unknown;
      pdfEngineCommand?: unknown;
      excelEngineDir?: unknown;
      excelEngineCommand?: unknown;
    };
    const cfg = readAppConfig();
    const pdfDir = typeof body.pdfEngineDir === 'string' ? body.pdfEngineDir : cfg.externalEngines.pdfEngineDir;
    const pdfCmd =
      typeof body.pdfEngineCommand === 'string' ? body.pdfEngineCommand : cfg.externalEngines.pdfEngineCommand;
    const excelDir =
      typeof body.excelEngineDir === 'string' ? body.excelEngineDir : cfg.externalEngines.excelEngineDir;
    const excelCmd =
      typeof body.excelEngineCommand === 'string'
        ? body.excelEngineCommand
        : cfg.externalEngines.excelEngineCommand;

    const pdfProbe = probeExternalEngine('pdf', pdfDir, pdfCmd);
    const excelProbe = probeExternalEngine('excel', excelDir, excelCmd);
    res.json({ pdf: pdfProbe, excel: excelProbe });
  });

  // --- Real Document Parsing (PDF / Excel / Word / Text + External Folder Delegation) ---
  const docUpload = multer({
    dest: path.join(os.tmpdir(), 'maillog-doc-uploads'),
    limits: {
      fileSize: 50 * 1024 * 1024,
      files: 15,
    },
  });

  const cleanupFiles = (files: Express.Multer.File[]) => {
    for (const file of files) {
      try {
        if (fs.existsSync(file.path)) fs.unlinkSync(file.path);
      } catch (err) {
        const reason = err instanceof Error ? err.message : String(err);
        console.warn(`[cleanupFiles] Failed to remove temp upload (${path.basename(file.path)}): ${reason}`);
      }
    }
  };

  app.post('/api/documents/parse', docUpload.array('files', 15), async (req, res) => {
    const uploaded = (req.files as Express.Multer.File[] | undefined) ?? [];
    if (uploaded.length === 0) {
      res.status(400).json({ error: '파싱할 PDF, Excel(.xlsx/.csv), Word(.docx) 파일을 선택해 주세요.' });
      return;
    }

    const cfg = readAppConfig();
    try {
      const documents: ParsedDocumentSummary[] = [];
      const errors: Array<{ filename: string; error: string }> = [];

      for (const file of uploaded) {
        try {
          const parsed = await parseUploadedDocument(
            {
              path: file.path,
              originalname: file.originalname,
              mimetype: file.mimetype,
              size: file.size,
            },
            cfg.externalEngines
          );
          documents.push(parsed);
        } catch (err) {
          errors.push({
            filename: file.originalname,
            error: err instanceof Error ? err.message : '문서 파싱 실패',
          });
        }
      }

      if (documents.length === 0 && errors.length > 0) {
        res.status(400).json({ error: errors.map((e) => `${e.filename}: ${e.error}`).join(' / '), errors });
        return;
      }

      res.json({ documents, errors });
    } finally {
      cleanupFiles(uploaded);
    }
  });

  // --- AI Document & Attachment Analysis (Cross-Referenced with Outlook Emails) ---
  app.post('/api/analyze/documents', async (req, res) => {
    const cfg = readAppConfig();
    const body = req.body as {
      documents?: ParsedDocumentSummary[];
      emails?: MailItem[];
      provider?: AiProviderId;
      maskSensitiveInfo?: boolean;
      customFocus?: string;
    };

    try {
      const validProviders: AiProviderId[] = ['groq', 'claude', 'openai', 'builtin'];
      const provider: AiProviderId =
        body.provider && validProviders.includes(body.provider) ? body.provider : cfg.ai.defaultProvider;

      const result = await analyzeDocumentsWithAi({
        documents: Array.isArray(body.documents) ? body.documents : [],
        emails: Array.isArray(body.emails) ? body.emails : [],
        provider,
        config: cfg,
        maskSensitiveInfo: body.maskSensitiveInfo !== false,
        customFocus: typeof body.customFocus === 'string' ? body.customFocus.slice(0, 500) : '',
      });
      res.json(result);
    } catch (error) {
      const message = error instanceof Error ? error.message : '문서 분석 중 오류가 발생했습니다.';
      res.status(400).json({ error: message });
    }
  });

  // --- Audio File Transcription + Structured Meeting Minutes AI Synthesis ---
  app.post(
    '/api/analyze/meeting',
    docUpload.fields([
      { name: 'audio', maxCount: 2 },
      { name: 'docs', maxCount: 10 },
    ]),
    async (req, res) => {
      const fileMap = (req.files as Record<string, Express.Multer.File[]> | undefined) ?? {};
      const audioFiles = fileMap.audio ?? [];
      const docFiles = fileMap.docs ?? [];
      const cfg = readAppConfig();

      try {
        const validProviders: AiProviderId[] = ['groq', 'claude', 'openai', 'builtin'];
        const requestedProvider = String(req.body?.provider || '') as AiProviderId;
        const provider: AiProviderId = validProviders.includes(requestedProvider)
          ? requestedProvider
          : cfg.ai.defaultProvider;

        const titleHint = typeof req.body?.title === 'string' ? req.body.title.trim() : '';
        const meetingDate =
          typeof req.body?.meetingDate === 'string' && req.body.meetingDate.trim()
            ? req.body.meetingDate.trim()
            : new Date().toISOString().slice(0, 10);
        const notesText = typeof req.body?.notesText === 'string' ? req.body.notesText : '';
        const customFocus = typeof req.body?.customFocus === 'string' ? req.body.customFocus : '';
        const maskSensitiveInfo = String(req.body?.maskSensitiveInfo ?? 'true') !== 'false';

        if (audioFiles.length === 0 && docFiles.length === 0 && !notesText.trim()) {
          res.status(400).json({
            error: '분석할 음성 파일, 참고 문서, 또는 회의 메모 중 최소 1개 이상이 필요합니다.',
          });
          return;
        }

        let linkedEmails: MailItem[] = [];
        if (typeof req.body?.emailsJson === 'string' && req.body.emailsJson.trim()) {
          try {
            const parsed: unknown = JSON.parse(req.body.emailsJson);
            if (!Array.isArray(parsed)) {
              res.status(400).json({ error: 'emailsJson 필드는 JSON 배열 형식이어야 합니다.' });
              return;
            }
            linkedEmails = parsed as MailItem[];
          } catch (parseErr) {
            const reason = parseErr instanceof Error ? parseErr.message : 'JSON 구문 오류';
            res.status(400).json({ error: `emailsJson 파싱 실패: ${reason}` });
            return;
          }
        }

        // 1. Transcribe uploaded real audio files (gemini-3.5-transcribe / Groq Whisper / OpenAI Whisper)
        const transcriptParts: string[] = [];
        let transcriptionEngine: string | undefined;
        for (const af of audioFiles) {
          const stt = await transcribeAudioFile({
            filePath: af.path,
            originalName: af.originalname,
            mimeType: af.mimetype,
            preferredProvider: provider,
          });
          transcriptionEngine = stt.engine;
          transcriptParts.push(`[음성 파일: ${af.originalname}]\n${stt.transcript}`);
        }

        // 2. Parse optional reference PDF/Excel/Word documents attached to the meeting
        const parsedDocs: ParsedDocumentSummary[] = [];
        for (const df of docFiles) {
          const parsed = await parseUploadedDocument(
            {
              path: df.path,
              originalname: df.originalname,
              mimetype: df.mimetype,
              size: df.size,
            },
            cfg.externalEngines
          );
          parsedDocs.push(parsed);
        }

        // 3. Generate structured Meeting Minutes
        const result = await analyzeMeetingMinutesWithAi({
          titleHint,
          meetingDate,
          transcript: transcriptParts.join('\n\n'),
          transcriptionEngine,
          notesText,
          documents: parsedDocs,
          emails: linkedEmails,
          provider,
          config: cfg,
          maskSensitiveInfo,
          customFocus,
        });

        res.json(result);
      } catch (error) {
        const message =
          error instanceof Error ? error.message : '회의록 및 음성 파일 분석 중 오류가 발생했습니다.';
        res.status(400).json({ error: message });
      } finally {
        cleanupFiles([...audioFiles, ...docFiles]);
      }
    }
  );

  // --- RFC 4155 .mbox Archive Export (from loaded Outlook / Webmail / PST items) ---
  app.post('/api/export/mbox', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        emails?: MailItem[];
        filename?: string;
      };
      const emails = Array.isArray(body.emails)
        ? body.emails.filter((item) => Boolean(item && typeof item === 'object'))
        : [];
      if (emails.length === 0) {
        res.status(400).json({ error: '아카이브로 내보낼 메일이 1건 이상 필요합니다.' });
        return;
      }
      const mbox = exportEmailsToMbox(emails);
      const safeName = (
        typeof body.filename === 'string' && body.filename.trim()
          ? body.filename.trim()
          : `maillog-archive-${new Date().toISOString().slice(0, 10)}.mbox`
      ).replace(/[^a-zA-Z0-9._-]/g, '_');
      res.setHeader('Content-Type', 'application/mbox; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="${safeName}"`);
      res.send(mbox);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : '.mbox 아카이브 생성 중 오류가 발생했습니다.';
      res.status(400).json({ error: message });
    }
  });

  // --- Microsoft Graph OAuth (Popup Flow for Live Date-Range Outlook Sync) ---
  app.get('/api/auth/microsoft/url', (req, res) => {
    const clientId = process.env.MICROSOFT_CLIENT_ID?.trim();
    const tenant = process.env.MICROSOFT_TENANT_ID?.trim() || 'common';
    if (!clientId) {
      res.status(400).json({
        error: 'MICROSOFT_CLIENT_ID 및 MICROSOFT_CLIENT_SECRET이 .env(Secrets)에 설정되어 있지 않습니다.',
      });
      return;
    }

    const baseUrl = getAppBaseUrl(req);
    const redirectUri = `${baseUrl}/auth/microsoft/callback`;
    const params = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: 'openid profile email offline_access User.Read Mail.Read',
      prompt: 'select_account',
    });

    const authUrl = `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/authorize?${params.toString()}`;
    res.json({ url: authUrl, redirectUri });
  });

  const handleMicrosoftCallback: express.RequestHandler = async (req, res) => {
    const code = typeof req.query.code === 'string' ? req.query.code : '';
    const errorDesc = typeof req.query.error_description === 'string' ? req.query.error_description : '';
    if (!code) {
      res.status(400).send(`
        <html lang="ko"><body style="font-family:sans-serif;padding:24px;">
          <h3>Microsoft 로그인 실패</h3>
          <p>${errorDesc || '인증 코드가 전달되지 않았습니다.'}</p>
        </body></html>
      `);
      return;
    }

    const clientId = process.env.MICROSOFT_CLIENT_ID?.trim();
    const clientSecret = process.env.MICROSOFT_CLIENT_SECRET?.trim();
    const tenant = process.env.MICROSOFT_TENANT_ID?.trim() || 'common';
    const baseUrl = getAppBaseUrl(req);
    const redirectUri = `${baseUrl}/auth/microsoft/callback`;

    if (!clientId || !clientSecret) {
      res.status(500).send('Microsoft OAuth 환경변수가 누락되었습니다.');
      return;
    }

    try {
      const tokenRes = await fetch(
        `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({
            client_id: clientId,
            client_secret: clientSecret,
            code,
            redirect_uri: redirectUri,
            grant_type: 'authorization_code',
          }),
        }
      );

      if (!tokenRes.ok) {
        const errBody = await tokenRes.text();
        throw new Error(`토큰 교환 실패: ${errBody.slice(0, 200)}`);
      }

      const tokenData = (await tokenRes.json()) as {
        access_token: string;
        refresh_token?: string;
        expires_in?: number;
      };

      const meRes = await fetch('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName', {
        headers: { Authorization: `Bearer ${tokenData.access_token}` },
      });
      const meData = meRes.ok
        ? ((await meRes.json()) as { displayName?: string; mail?: string; userPrincipalName?: string })
        : {};

      const sid = crypto.randomBytes(24).toString('hex');
      const nowMs = Date.now();
      msSessions.set(sid, {
        accessToken: tokenData.access_token,
        refreshToken: tokenData.refresh_token,
        expiresAt: nowMs + (tokenData.expires_in || 3600) * 1000,
        sessionExpiresAt: nowMs + MS_SESSION_TTL_MS,
        userEmail: meData.mail || meData.userPrincipalName || 'outlook-user',
        displayName: meData.displayName || 'Outlook 사용자',
      });
      saveMsSessionsToDisk();

      const isHttps = req.protocol === 'https' || req.headers['x-forwarded-proto'] === 'https';
      res.setHeader(
        'Set-Cookie',
        isHttps
          ? `ms_session_id=${encodeURIComponent(sid)}; Path=/; HttpOnly; Secure; SameSite=None; Max-Age=2592000`
          : `ms_session_id=${encodeURIComponent(sid)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000`
      );
      res.send(`<!DOCTYPE html>
<html lang="ko">
  <body style="font-family:sans-serif;padding:24px;color:#0f172a;">
    <script>
      if (window.opener) {
        window.opener.postMessage({ type: 'OAUTH_AUTH_SUCCESS', provider: 'microsoft' }, '*');
        window.close();
      } else {
        window.location.href = '/';
      }
    </script>
    <p>Microsoft Outlook 계정 연동이 완료되었습니다. 창이 자동으로 닫히지 않으면 닫아 주세요.</p>
  </body>
</html>`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '알 수 없는 오류';
      res.status(500).send(`<html><body><h3>연동 오류</h3><p>${msg}</p></body></html>`);
    }
  };

  app.get(['/auth/microsoft/callback', '/auth/microsoft/callback/'], handleMicrosoftCallback);

  app.post('/api/auth/microsoft/logout', (req, res) => {
    const entry = resolveMicrosoftSessionEntry(req);
    if (entry) {
      msSessions.delete(entry.sid);
      saveMsSessionsToDisk();
    }
    res.setHeader('Set-Cookie', 'ms_session_id=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0');
    res.json({ ok: true });
  });

  app.post('/api/outlook/messages', async (req, res) => {
    const session = await ensureFreshMicrosoftSession(req);
    if (!session) {
      res.status(401).json({ error: 'Microsoft Outlook 계정이 연결되어 있지 않습니다.' });
      return;
    }

    const { startDate, endDate, folders } = req.body as {
      startDate?: string;
      endDate?: string;
      folders?: Array<'inbox' | 'sent'>;
    };

    const targetFolders: Array<'inbox' | 'sent'> =
      Array.isArray(folders) && folders.length > 0 ? folders : ['inbox', 'sent'];
    const startIso = startDate ? `${startDate}T00:00:00Z` : new Date(Date.now() - 7 * 86400000).toISOString();
    const endIso = endDate ? `${endDate}T23:59:59Z` : new Date().toISOString();

    try {
      const allMails: MailItem[] = [];
      for (const folder of targetFolders) {
        const graphFolder = folder === 'inbox' ? 'inbox' : 'sentitems';
        const dateField = folder === 'inbox' ? 'receivedDateTime' : 'sentDateTime';
        const filter = encodeURIComponent(`${dateField} ge ${startIso} and ${dateField} le ${endIso}`);
        const url = `https://graph.microsoft.com/v1.0/me/mailFolders/${graphFolder}/messages?$filter=${filter}&$top=50&$orderby=${dateField} desc&$select=id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview,uniqueBody,body,hasAttachments,importance,isRead`;

        const graphRes = await fetch(url, {
          headers: {
            Authorization: `Bearer ${session.accessToken}`,
            Prefer: 'outlook.body-content-type="text"',
          },
        });
        if (!graphRes.ok) {
          const text = await graphRes.text();
          throw new Error(`Microsoft Graph 메일 조회 실패 (${graphRes.status}): ${text.slice(0, 180)}`);
        }

        const data = (await graphRes.json()) as {
          value?: Array<{
            id: string;
            subject?: string;
            from?: { emailAddress?: { name?: string; address?: string } };
            toRecipients?: Array<{ emailAddress?: { address?: string } }>;
            receivedDateTime?: string;
            sentDateTime?: string;
            bodyPreview?: string;
            uniqueBody?: { content?: string; contentType?: string };
            body?: { content?: string; contentType?: string };
            hasAttachments?: boolean;
            importance?: string;
            isRead?: boolean;
          }>;
        };

        for (const item of data.value ?? []) {
          const rawTs = item.receivedDateTime || item.sentDateTime || new Date().toISOString();
          const dt = new Date(rawTs);
          const dateStr = dt.toISOString().slice(0, 10);
          const timeLabel = dt.toTimeString().slice(0, 5);
          const preferred = item.uniqueBody?.content ? item.uniqueBody : item.body;
          const bodyText =
            preferred?.contentType === 'html'
              ? stripHtmlToPlainText(preferred.content || '')
              : preferred?.content || item.bodyPreview || '';

          allMails.push({
            id: item.id,
            folder,
            subject: item.subject || '(제목 없음)',
            fromName: item.from?.emailAddress?.name || item.from?.emailAddress?.address || '알 수 없음',
            fromEmail: item.from?.emailAddress?.address || '',
            to: (item.toRecipients ?? [])
              .map((r) => r.emailAddress?.address || '')
              .filter(Boolean),
            date: dateStr,
            timestamp: rawTs,
            timeLabel,
            snippet: item.bodyPreview || bodyText.slice(0, 120),
            body: bodyText,
            hasAttachment: Boolean(item.hasAttachments),
            importance: item.importance === 'high' ? 'high' : 'normal',
            isRead: Boolean(item.isRead),
            source: 'graph-live',
          });
        }
      }

      res.json({ emails: allMails, userEmail: session.userEmail });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Graph API 조회 실패';
      res.status(500).json({ error: message });
    }
  });

  // Microsoft Graph /$batch (20 items per batch) for Outlook Multi-Select REST IDs
  app.post('/api/outlook/batch-messages', async (req, res) => {
    const session = await ensureFreshMicrosoftSession(req);
    if (!session) {
      res.status(401).json({ error: 'Microsoft Outlook 계정이 연결되어 있지 않습니다.' });
      return;
    }

    const { restIds } = req.body as { restIds?: string[] };
    if (!Array.isArray(restIds) || restIds.length === 0) {
      res.status(400).json({ error: '조회할 선택 메일 ID가 없습니다.' });
      return;
    }

    try {
      const results: MailItem[] = [];
      for (let i = 0; i < restIds.length; i += 20) {
        const chunk = restIds.slice(i, i + 20);
        const batchPayload = {
          requests: chunk.map((restId, idx) => ({
            id: String(idx + 1),
            method: 'GET',
            url: `/me/messages/${encodeURIComponent(restId)}?$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,uniqueBody,body,hasAttachments,importance,isRead`,
            headers: { Prefer: 'outlook.body-content-type="text"' },
          })),
        };

        const batchRes = await fetch('https://graph.microsoft.com/v1.0/$batch', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${session.accessToken}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(batchPayload),
        });
        if (!batchRes.ok) {
          throw new Error(`Graph $batch 실패 (${batchRes.status})`);
        }
        const batchData = (await batchRes.json()) as {
          responses?: Array<{
            status: number;
            body?: {
              id: string;
              subject?: string;
              from?: { emailAddress?: { name?: string; address?: string } };
              toRecipients?: Array<{ emailAddress?: { address?: string } }>;
              receivedDateTime?: string;
              bodyPreview?: string;
              uniqueBody?: { content?: string; contentType?: string };
              body?: { content?: string; contentType?: string };
              hasAttachments?: boolean;
              importance?: string;
              isRead?: boolean;
            };
          }>;
        };

        for (const r of batchData.responses ?? []) {
          if (r.status === 200 && r.body) {
            const m = r.body;
            const rawTs = m.receivedDateTime || new Date().toISOString();
            const dt = new Date(rawTs);
            const fromAddr = m.from?.emailAddress?.address || '';
            const isSent = fromAddr.toLowerCase() === session.userEmail.toLowerCase();
            const preferred = m.uniqueBody?.content ? m.uniqueBody : m.body;
            const bodyText =
              preferred?.contentType === 'html'
                ? stripHtmlToPlainText(preferred.content || '')
                : preferred?.content || m.bodyPreview || '';

            results.push({
              id: m.id,
              folder: isSent ? 'sent' : 'inbox',
              subject: m.subject || '(제목 없음)',
              fromName: m.from?.emailAddress?.name || fromAddr || '알 수 없음',
              fromEmail: fromAddr,
              to: (m.toRecipients ?? []).map((t) => t.emailAddress?.address || '').filter(Boolean),
              date: dt.toISOString().slice(0, 10),
              timestamp: rawTs,
              timeLabel: dt.toTimeString().slice(0, 5),
              snippet: m.bodyPreview || bodyText.slice(0, 120),
              body: bodyText,
              hasAttachment: Boolean(m.hasAttachments),
              importance: m.importance === 'high' ? 'high' : 'normal',
              isRead: Boolean(m.isRead),
              source: 'graph-live',
            });
          }
        }
      }

      res.json({ emails: results });
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Graph $batch 조회 실패';
      res.status(500).json({ error: message });
    }
  });

  // --- Company POP3 / SMTP Auto-Discovery & Live Mail Connect (Mailplug, Hiworks, Custom Domain, etc.) ---
  const extractPop3Bearer = (req: express.Request): string => {
    const auth = req.headers.authorization || '';
    const m = auth.match(/^Bearer\s+([A-Za-z0-9_-]{32,128})$/i);
    return m ? m[1] : '';
  };

  app.use('/api/pop3', (req, res, next) => {
    // Do not expose wildcard CORS on credentialed/authenticated mail endpoints
    const origin = req.headers.origin;
    const host = req.headers['x-forwarded-host']?.toString().split(',')[0] || req.headers.host || '';
    if (origin && host && origin.includes(host)) {
      res.setHeader('Access-Control-Allow-Origin', origin);
    } else {
      res.removeHeader('Access-Control-Allow-Origin');
    }
    res.setHeader('Cache-Control', 'no-store');
    next();
  });

  app.get('/api/pop3/presets', (_req, res) => {
    res.json({
      presets: MAIL_SERVICE_PRESETS.map((p) => ({
        id: p.id,
        label: p.label,
        vendorHint: p.vendorHint,
        pop3Port: p.pop3Port,
        pop3Tls: p.pop3Tls,
        smtpPort: p.smtpPort,
        smtpTls: p.smtpTls,
        note: p.note,
      })),
    });
  });

  app.post('/api/pop3/autodiscover', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        email?: string;
        providerPreset?: MailServicePresetId;
      };
      const email = String(body.email || '').trim();
      if (!email) {
        res.status(400).json({ error: '회사 이메일 주소를 입력해 주세요.' });
        return;
      }
      const [discovered, serverEgress] = await Promise.all([
        detectCompanyPop3Config({
          email,
          providerPreset: body.providerPreset,
        }),
        detectServerEgressInfo(),
      ]);
      res.json({ config: discovered, serverEgress });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'POP3 설정 자동 조회에 실패했습니다.';
      res.status(400).json({ error: sanitizeMailError(msg) });
    }
  });

  app.post('/api/pop3/connect', async (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      email?: string;
      password?: string;
      providerPreset?: MailServicePresetId;
      protocol?: MailReceiveProtocol;
      pop3Host?: string;
      pop3Port?: number;
      pop3Tls?: boolean;
      imapHost?: string;
      imapPort?: number;
      imapTls?: boolean;
      smtpHost?: string;
      smtpPort?: number;
      smtpTls?: boolean;
      maxMessages?: number;
    };
    const email = String(body.email || '').trim();
    const password = String(body.password || '');
    if (!email || !password) {
      res.status(400).json({ error: '이메일 주소와 비밀번호를 모두 입력해 주세요.' });
      return;
    }

    try {
      const discovered = await detectCompanyPop3Config({
        email,
        providerPreset: body.providerPreset,
      });

      const requestedProtocol: MailReceiveProtocol =
        body.protocol === 'imap' || body.protocol === 'pop3'
          ? body.protocol
          : discovered.recommendedProtocol;

      const pop3Host = String(body.pop3Host || '').trim() || discovered.pop3Host;
      const pop3Port = Number(body.pop3Port) || discovered.pop3Port;
      const pop3Tls = typeof body.pop3Tls === 'boolean' ? body.pop3Tls : discovered.pop3Tls;
      const imapHost = String(body.imapHost || '').trim() || discovered.imapHost;
      const imapPort = Number(body.imapPort) || discovered.imapPort;
      const imapTls = typeof body.imapTls === 'boolean' ? body.imapTls : discovered.imapTls;
      const smtpHost = String(body.smtpHost || '').trim() || discovered.smtpHost;
      const smtpPort = Number(body.smtpPort) || discovered.smtpPort;
      const smtpTls = typeof body.smtpTls === 'boolean' ? body.smtpTls : discovered.smtpTls;
      const maxMessages = Math.max(1, Math.min(50, Number(body.maxMessages) || 30));

      const healed = await connectMailWithAutoHeal({
        email,
        password,
        discovered,
        requestedProtocol,
        pop3Host,
        pop3Port,
        pop3Tls,
        imapHost,
        imapPort,
        imapTls,
        maxMessages,
      });

      const session = pop3Sessions.create({
        email,
        password: healed.effectivePassword,
        providerId: discovered.detectedProviderId,
        providerName: discovered.detectedProviderName,
        protocol: healed.effectiveProtocol,
        usernameFormat: discovered.usernameFormat,
        pop3Host: healed.effectiveProtocol === 'pop3' ? healed.effectiveHost : pop3Host,
        pop3Port: healed.effectiveProtocol === 'pop3' ? healed.effectivePort : pop3Port,
        pop3Tls: healed.effectiveProtocol === 'pop3' ? healed.effectiveTls : pop3Tls,
        imapHost: healed.effectiveProtocol === 'imap' ? healed.effectiveHost : imapHost,
        imapPort: healed.effectiveProtocol === 'imap' ? healed.effectivePort : imapPort,
        imapTls: healed.effectiveProtocol === 'imap' ? healed.effectiveTls : imapTls,
        smtpHost,
        smtpPort,
        smtpTls,
      });

      try {
        if (localDb.isReady() && healed.fetched.emails.length > 0) {
          localDb.upsertEmailsBatch(
            healed.fetched.emails.map((m) => ({
              ...m,
              storeName: session.email,
            }))
          );
        }
      } catch {
        // Best-effort local SQLite sync
      }

      res.json({
        sessionToken: session.token,
        account: toPublicPop3AccountInfo(session),
        discovered,
        protocolUsed: healed.effectiveProtocol,
        healedBy: healed.healedBy,
        autoHealAttempts: healed.attempts,
        totalMailboxCount: healed.fetched.totalMailboxCount,
        fetchedCount: healed.fetched.fetchedCount,
        emails: healed.fetched.emails,
      });
    } catch (err) {
      if (err instanceof MailAuthError) {
        res.status(400).json({
          error: sanitizeMailError(err.message, password),
          diagnosis: err.diagnosis,
        });
        return;
      }
      const rawMsg = err instanceof Error ? err.message : '메일 서버 연결에 실패했습니다.';
      res.status(400).json({ error: sanitizeMailError(rawMsg, password) });
    }
  });

  app.post('/api/pop3/sync', async (req, res) => {
    const token = extractPop3Bearer(req);
    const session = pop3Sessions.get(token);
    if (!session) {
      res.status(401).json({ error: '메일 연결 세션이 만료되었거나 연결되어 있지 않습니다. 다시 로그인해 주세요.' });
      return;
    }
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        maxMessages?: number;
      };
      const maxMessages = Math.max(1, Math.min(50, Number(body.maxMessages) || 30));
      const fetched =
        session.protocol === 'imap' && session.imapHost
          ? await fetchImapMessages({
              email: session.email,
              password: session.password,
              imapHost: session.imapHost,
              imapPort: session.imapPort,
              imapTls: session.imapTls,
              providerId: session.providerId,
              usernameFormat: session.usernameFormat,
              maxMessages,
            })
          : await fetchPop3Messages({
              email: session.email,
              password: session.password,
              pop3Host: session.pop3Host,
              pop3Port: session.pop3Port,
              pop3Tls: session.pop3Tls,
              providerId: session.providerId,
              usernameFormat: session.usernameFormat,
              maxMessages,
            });
      try {
        if (localDb.isReady() && fetched.emails.length > 0) {
          const existingBefore = localDb.getExistingEmailIds();
          const freshForMonitor = fetched.emails.filter(
            (m) => m?.id && !existingBefore.has(m.id)
          );
          localDb.upsertEmailsBatch(
            fetched.emails.map((m) => ({
              ...m,
              storeName: session.email,
            }))
          );
          if (freshForMonitor.length > 0) {
            void backgroundMonitor.ingestNewEmails(freshForMonitor);
          }
        }
      } catch {
        // Best-effort local SQLite sync
      }
      res.json({
        account: toPublicPop3AccountInfo(session),
        protocolUsed: fetched.protocolUsed,
        totalMailboxCount: fetched.totalMailboxCount,
        fetchedCount: fetched.fetchedCount,
        emails: fetched.emails,
      });
    } catch (err) {
      if (err instanceof MailAuthError) {
        res.status(400).json({
          error: sanitizeMailError(err.message, session.password),
          diagnosis: err.diagnosis,
        });
        return;
      }
      const rawMsg = err instanceof Error ? err.message : '메일 새로고침 실패';
      res.status(400).json({ error: sanitizeMailError(rawMsg, session.password) });
    }
  });

  app.post('/api/pop3/send', async (req, res) => {
    const token = extractPop3Bearer(req);
    const session = pop3Sessions.get(token);
    if (!session) {
      res.status(401).json({ error: '연결된 메일 계정이 없습니다. 먼저 POP3/SMTP 계정을 연결해 주세요.' });
      return;
    }
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        to?: string | string[];
        cc?: string | string[];
        bcc?: string | string[];
        subject?: string;
        body?: string;
        htmlBody?: string;
        importance?: 'low' | 'normal' | 'high';
        requestReadReceipt?: boolean;
        attachments?: Array<{ filename?: string; contentType?: string; contentBase64?: string }>;
      };
      const toList = Array.isArray(body.to)
        ? body.to
        : String(body.to || '')
            .split(/[;,]/)
            .map((s) => s.trim())
            .filter(Boolean);
      const ccList = Array.isArray(body.cc)
        ? body.cc
        : String(body.cc || '')
            .split(/[;,]/)
            .map((s) => s.trim())
            .filter(Boolean);
      const bccList = Array.isArray(body.bcc)
        ? body.bcc
        : String(body.bcc || '')
            .split(/[;,]/)
            .map((s) => s.trim())
            .filter(Boolean);

      const safeAttachments: Array<{ filename: string; contentType?: string; contentBase64: string }> = [];
      let totalAttBytes = 0;
      if (Array.isArray(body.attachments)) {
        for (const item of body.attachments.slice(0, 10)) {
          const fname = String(item?.filename || '').replace(/[\r\n\0]/g, '').trim();
          const b64 = String(item?.contentBase64 || '').replace(/\s+/g, '');
          if (!fname || !b64) continue;
          const approxBytes = Math.floor((b64.length * 3) / 4);
          if (totalAttBytes + approxBytes > 20 * 1024 * 1024) {
            throw new Error('첨부 파일 합계 용량은 20MB 이하만 발송 가능합니다.');
          }
          totalAttBytes += approxBytes;
          safeAttachments.push({
            filename: fname,
            contentType: typeof item?.contentType === 'string' ? item.contentType : undefined,
            contentBase64: b64,
          });
        }
      }

      const subjectText = String(body.subject || '(제목 없음)');
      const bodyText = String(body.body || '');
      const htmlBody = typeof body.htmlBody === 'string' && body.htmlBody.trim() ? body.htmlBody : undefined;
      const sent = await sendSmtpMessage({
        email: session.email,
        password: session.password,
        smtpHost: session.smtpHost,
        smtpPort: session.smtpPort,
        smtpTls: session.smtpTls,
        to: toList,
        cc: ccList,
        bcc: bccList,
        subject: subjectText,
        body: bodyText,
        htmlBody,
        importance:
          body.importance === 'high' || body.importance === 'low' ? body.importance : 'normal',
        requestReadReceipt: Boolean(body.requestReadReceipt),
        attachments: safeAttachments,
      });
      try {
        if (localDb.isReady()) {
          const now = new Date();
          const sentId = `smtp-sent-${now.getTime()}`;
          localDb.upsertEmailsBatch([
            {
              id: sentId,
              storeName: session.email,
              folder: 'sent',
              folderName: '보낸 편지함',
              subject: subjectText,
              fromName: session.email,
              fromEmail: session.email,
              to: sent.acceptedTo.length > 0 ? sent.acceptedTo : toList,
              cc: ccList.length > 0 ? ccList : undefined,
              date: now.toISOString().slice(0, 10),
              timestamp: now.toISOString(),
              timeLabel: now.toTimeString().slice(0, 5),
              snippet: bodyText.slice(0, 120),
              body: bodyText,
              isRead: true,
              source: 'outlook-live',
            },
          ]);
          if (htmlBody) {
            localDb.setEmailRichCache(sentId, sanitizeEmailHtml(htmlBody), []);
          }
        }
      } catch {
        // Best-effort local SQLite persistence
      }
      res.json({ ok: true, acceptedTo: sent.acceptedTo, fromEmail: session.email });
    } catch (err) {
      const rawMsg = err instanceof Error ? err.message : 'SMTP 메일 발송 실패';
      res.status(400).json({ error: sanitizeMailError(rawMsg, session.password) });
    }
  });

  app.post('/api/local/outlook/send-via-com', async (req, res) => {
    if (process.platform !== 'win32') {
      res.status(400).json({ error: 'Outlook COM 발송은 Windows 환경에서만 지원됩니다.' });
      return;
    }
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      to?: string;
      cc?: string;
      bcc?: string;
      subject?: string;
      body?: string;
      htmlBody?: string;
      action?: 'send' | 'display';
      importance?: 'low' | 'normal' | 'high';
      requestReadReceipt?: boolean;
      sourceEmailId?: string;
      forwardAttachmentIndices?: number[];
      userAttachments?: Array<{ filename?: string; contentType?: string; contentBase64?: string }>;
    };
    const toStr = String(body.to || '').trim();
    const ccStr = String(body.cc || '').trim();
    const bccStr = String(body.bcc || '').trim();
    const subjectStr = String(body.subject || '(제목 없음)').trim();
    const action = body.action === 'display' ? 'display' : 'send';
    if (action === 'send' && !toStr && !ccStr && !bccStr) {
      res.status(400).json({ error: '받는 사람 이메일 주소를 입력해 주세요.' });
      return;
    }

    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-com-send-'));
    try {
      const inlineAtts: Array<{ cid: string; filePath: string }> = [];
      let imgIdx = 0;
      const rawHtml =
        typeof body.htmlBody === 'string' && body.htmlBody.trim()
          ? body.htmlBody
          : `<div style="white-space:pre-wrap;font-family:'Malgun Gothic',sans-serif;">${String(body.body || '')
              .replace(/&/g, '&amp;')
              .replace(/</g, '&lt;')
              .replace(/>/g, '&gt;')}</div>`;

      const cidHtml = rawHtml.replace(
        /(src\s*=\s*)(["'])(data:image\/([a-z0-9+.-]+);base64,([A-Za-z0-9+/=\s]+))\2/gi,
        (_full, pre: string, q: string, _uri: string, subtype: string, rawB64: string) => {
          imgIdx += 1;
          const ext = subtype === 'jpeg' ? 'jpg' : subtype === 'svg+xml' ? 'svg' : subtype.replace(/[^a-z0-9]/gi, '') || 'png';
          const cid = `bmail_sig_${imgIdx}@bmail.local`;
          const filePath = path.join(tempDir, `inline-${imgIdx}.${ext}`);
          fs.writeFileSync(filePath, Buffer.from(rawB64.replace(/\s+/g, ''), 'base64'));
          inlineAtts.push({ cid, filePath });
          return `${pre}${q}cid:${cid}${q}`;
        }
      );

      const fileAtts: string[] = [];
      if (body.sourceEmailId && Array.isArray(body.forwardAttachmentIndices) && localDb.isReady()) {
        const rec = localDb.getEmailById(body.sourceEmailId);
        if (rec?.pstPath && rec.descriptorNodeId !== null && localDb.isAllowedPstPath(rec.pstPath)) {
          for (const attIdx of body.forwardAttachmentIndices.slice(0, 10)) {
            try {
              const extracted = readPstAttachmentBytes(rec.pstPath, rec.descriptorNodeId, Number(attIdx));
              const safeName = extracted.filename.replace(/[<>:"/\\|?*]/g, '_');
              const outFile = path.join(tempDir, `fwd-${attIdx}-${safeName}`);
              fs.writeFileSync(outFile, extracted.buffer);
              fileAtts.push(outFile);
            } catch {
              // Skip unreadable PST attachment cleanly
            }
          }
        }
      }

      if (Array.isArray(body.userAttachments)) {
        let userTotal = 0;
        let uIdx = 0;
        for (const ua of body.userAttachments.slice(0, 10)) {
          const safeName = path.basename(String(ua?.filename || '')).replace(/[<>:"/\\|?*\r\n\0]/g, '_').trim();
          const b64 = String(ua?.contentBase64 || '').replace(/\s+/g, '');
          if (!safeName || !b64) continue;
          const buf = Buffer.from(b64, 'base64');
          if (userTotal + buf.length > 20 * 1024 * 1024) {
            throw new Error('첨부 파일 합계 용량은 20MB 이하만 지원합니다.');
          }
          userTotal += buf.length;
          uIdx += 1;
          const outFile = path.join(tempDir, `u-${uIdx}-${safeName}`);
          fs.writeFileSync(outFile, buf);
          fileAtts.push(outFile);
        }
      }

      const payloadPath = path.join(tempDir, 'payload.json');
      const importanceNum = body.importance === 'high' ? 2 : body.importance === 'low' ? 0 : 1;
      fs.writeFileSync(
        payloadPath,
        JSON.stringify({
          to: toStr,
          cc: ccStr,
          bcc: bccStr,
          subject: subjectStr,
          htmlBody: cidHtml,
          action,
          importance: importanceNum,
          requestReadReceipt: Boolean(body.requestReadReceipt),
          inlineAtts,
          fileAtts,
        }),
        'utf8'
      );

      const psScript = `
$ErrorActionPreference = 'Stop'
$payload = Get-Content -LiteralPath '${payloadPath.replace(/'/g, "''")}' -Raw -Encoding UTF8 | ConvertFrom-Json
$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch {}
if ($null -eq $app) { $app = New-Object -ComObject Outlook.Application }
$mail = $app.CreateItem(0)
if ($payload.to) { $mail.To = [string]$payload.to }
if ($payload.cc) { $mail.CC = [string]$payload.cc }
if ($payload.bcc) { $mail.BCC = [string]$payload.bcc }
$mail.Subject = [string]$payload.subject
$mail.Importance = [int]$payload.importance
if ($payload.requestReadReceipt) { $mail.ReadReceiptRequested = $true }
$mail.BodyFormat = 2
$mail.HTMLBody = [string]$payload.htmlBody
$cidProp = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
foreach ($ia in $payload.inlineAtts) {
  if (Test-Path -LiteralPath $ia.filePath) {
    $att = $mail.Attachments.Add([string]$ia.filePath, 1, 0)
    try { $att.PropertyAccessor.SetProperty($cidProp, [string]$ia.cid) } catch {}
  }
}
foreach ($fp in $payload.fileAtts) {
  if (Test-Path -LiteralPath $fp) {
    $mail.Attachments.Add([string]$fp) | Out-Null
  }
}
if ($payload.action -eq 'display') {
  $mail.Display()
} else {
  $mail.Send()
}
Write-Output 'COM_OK'
`.trim();

      await new Promise<void>((resolve, reject) => {
        execFile(
          'powershell.exe',
          ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', psScript],
          { encoding: 'utf8', timeout: 25_000, windowsHide: true },
          (err, stdout, stderr) => {
            if (err) {
              reject(new Error(stderr?.trim() || err.message));
              return;
            }
            if (!String(stdout || '').includes('COM_OK')) {
              reject(new Error('Outlook COM 작업 응답을 확인하지 못했습니다.'));
              return;
            }
            resolve();
          }
        );
      });

      if (action === 'send' && localDb.isReady()) {
        try {
          const now = new Date();
          const sentId = `outlook-com-sent-${now.getTime()}`;
          const toArr = toStr.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
          const ccArr = ccStr.split(/[;,]/).map((s) => s.trim()).filter(Boolean);
          localDb.upsertEmailsBatch([
            {
              id: sentId,
              storeName: 'PC Outlook COM',
              folder: 'sent',
              folderName: '보낸 편지함',
              subject: subjectStr,
              fromName: '나 (PC Outlook)',
              fromEmail: 'me@outlook.local',
              to: toArr,
              cc: ccArr.length > 0 ? ccArr : undefined,
              date: now.toISOString().slice(0, 10),
              timestamp: now.toISOString(),
              timeLabel: now.toTimeString().slice(0, 5),
              snippet: String(body.body || '').slice(0, 120),
              body: String(body.body || ''),
              importance: body.importance === 'high' ? 'high' : 'normal',
              hasAttachment: fileAtts.length > 0,
              isRead: true,
              source: 'outlook-live',
            },
          ]);
          localDb.setEmailRichCache(sentId, sanitizeEmailHtml(rawHtml), []);
        } catch {
          // Best-effort local SQLite sent persistence
        }
      }

      res.json({ ok: true, action, attachmentsCount: fileAtts.length, inlineImagesCount: inlineAtts.length });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Outlook COM 메일 처리 실패';
      res.status(400).json({ error: msg });
    } finally {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch {
        // ignore temp cleanup failure
      }
    }
  });

  app.get('/api/pop3/session', (req, res) => {
    const token = extractPop3Bearer(req);
    const byToken = token ? pop3Sessions.get(token) : null;
    const session = byToken || (isLoopbackRequest(req) ? pop3Sessions.getLatest() : null);
    if (!session) {
      res.json({ connected: false });
      return;
    }
    res.json({
      connected: true,
      sessionToken: session.token,
      account: toPublicPop3AccountInfo(session),
    });
  });

  app.post('/api/pop3/disconnect', (req, res) => {
    const token = extractPop3Bearer(req);
    pop3Sessions.delete(token);
    res.json({ ok: true });
  });

  // --- Local Windows PC Outlook Profile / PST / Address Book Inspector & API Health Check ---
  app.get('/api/outlook/pc-inspector.ps1', (_req, res) => {
    const script = buildWindowsOutlookInspectorScript();
    res.setHeader('Content-Type', 'text/plain; charset=utf-8');
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="bmail-outlook-pc-inspector.ps1"'
    );
    res.send(`\uFEFF${script}`);
  });

  app.post('/api/outlook/pc-profile/import', (req, res) => {
    try {
      const profile = parseWindowsOutlookProfileJson(req.body);
      res.json({ profile });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'PC Outlook 프로필 JSON 파싱 실패';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/config/test-connections', (req, res) => {
    const cfg = readAppConfig();
    const verifiedEmp = resolveVerifiedEmployeeId(req);
    const runtimeEnv = emvVault.resolveRuntimeEnv(verifiedEmp);
    const providers = getProviderStatuses(cfg, runtimeEnv);
    const hybrid = getHybridCreditStatus(new Date(), { ...process.env, ...runtimeEnv });
    const msConfigured = Boolean(
      (runtimeEnv.MICROSOFT_CLIENT_ID || process.env.MICROSOFT_CLIENT_ID)?.trim() &&
      (runtimeEnv.MICROSOFT_CLIENT_SECRET || process.env.MICROSOFT_CLIENT_SECRET)?.trim()
    );
    res.json({
      checkedAt: new Date().toISOString(),
      verifiedEmployeeId: verifiedEmp,
      connections: [
        {
          id: 'gemini',
          label: 'Gemini 4 Hydrogen (1순위 실시간 AI)',
          configured: hybrid.realtimeGeminiCredit.configured,
          scope: verifiedEmp ? 'per-user' : 'org-wide',
          detail: hybrid.realtimeGeminiCredit.configured
            ? `활성 · 모델 ${hybrid.realtimeGeminiCredit.model} (${verifiedEmp ? `사원 ${verifiedEmp} 기준 적용` : '전사 공통 EMV/.env 적용'})`
            : '미설정 · EMV 관리 콘솔 또는 서버 .env에 GEMINI_API_KEY 등록 필요',
        },
        {
          id: 'agent-builder',
          label: 'Vertex AI Agent Builder (과거 메일 RAG)',
          configured: hybrid.agentBuilderRagCredit.configured,
          scope: 'org-wide',
          detail: hybrid.agentBuilderRagCredit.configured
            ? `${hybrid.agentBuilderRagCredit.mode} 활성 · 만료 가드 ${hybrid.agentBuilderRagCredit.expiryGuard.expiryDate}`
            : `미연결 · Gemini 1순위 폴백 동작 중 (만료 가드 ${hybrid.agentBuilderRagCredit.expiryGuard.expiryDate})`,
        },
        ...providers
          .filter((p) => p.id !== 'builtin')
          .map((p) => ({
            id: p.id,
            label: `${p.name} (2순위 폴백)`,
            configured: p.configured,
            scope: verifiedEmp ? 'per-user' : 'org-wide',
            detail: p.configured ? `활성 · 모델 ${p.model}` : `미설정 (${p.envKey})`,
          })),
        {
          id: 'microsoft',
          label: 'Microsoft 365 Outlook OAuth',
          configured: msConfigured,
          scope: 'org-wide',
          detail: msConfigured
            ? 'Azure Entra ID OAuth 준비됨 (직원 개별 로그인 가능)'
            : '미설정 (대신 메일플러그·회사 POP3 자동 연결 또는 PST 불러오기 사용 가능)',
        },
        {
          id: 'pop3-auto',
          label: '회사 메일 POP3 · SMTP 자동 탐지 엔진',
          configured: true,
          scope: 'per-user',
          detail: '항시 활성 · 직원별 이메일+비밀번호만으로 메일플러그/하이웍스/자사 도메인 자동 연결',
        },
      ],
    });
  });

  // --- Encrypted EMV (Environment Variable) Admin Vault Routes ---
  app.use('/api/emv', (req, res, next) => {
    // Strictly restrict CORS to same-origin on admin secret management endpoints
    const origin = req.headers.origin;
    const host = req.headers['x-forwarded-host']?.toString().split(',')[0] || req.headers.host || '';
    if (origin && host && origin.includes(host)) {
      res.setHeader('Access-Control-Allow-Origin', origin);
    } else {
      res.removeHeader('Access-Control-Allow-Origin');
    }
    res.setHeader('Cache-Control', 'no-store');
    next();
  });

  const requireEmvAdmin: express.RequestHandler = (req, res, next) => {
    const token = extractBearerToken(req);
    const session = emvVault.verifyAdminSession(token);
    if (!session) {
      res.status(401).json({ error: '종합 관리자 인증이 필요합니다. 먼저 관리자 로그인을 진행해 주세요.' });
      return;
    }
    next();
  };

  app.get('/api/emv/status', (req, res) => {
    const token = extractBearerToken(req);
    const auth = emvVault.getAuthStatus(token);
    const verifiedEmployeeId = resolveVerifiedEmployeeId(req);
    const publicData = auth.authenticated
      ? emvVault.listPublicEntries()
      : { globalEntries: [], employees: [] };

    res.json({
      auth,
      catalog: EMV_KEY_CATALOG,
      verifiedEmployeeId,
      globalEntries: publicData.globalEntries,
      employees: publicData.employees,
      currentResolution: emvVault.getResolutionPreview(verifiedEmployeeId),
    });
  });

  app.post('/api/emv/admin/claim', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as { password?: string };
      const session = emvVault.claimInitialAdminPassword(String(body.password || ''));
      const publicData = emvVault.listPublicEntries();
      res.json({
        token: session.token,
        expiresAt: session.expiresAt,
        auth: emvVault.getAuthStatus(session.token),
        globalEntries: publicData.globalEntries,
        employees: publicData.employees,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '종합 관리자 초기 비밀번호 등록에 실패했습니다.';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/emv/admin/login', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as { password?: string };
      const session = emvVault.authenticateAdmin(String(body.password || ''));
      const publicData = emvVault.listPublicEntries();
      res.json({
        token: session.token,
        expiresAt: session.expiresAt,
        auth: emvVault.getAuthStatus(session.token),
        globalEntries: publicData.globalEntries,
        employees: publicData.employees,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '종합 관리자 인증에 실패했습니다.';
      res.status(401).json({ error: msg });
    }
  });

  app.post('/api/emv/admin/logout', (req, res) => {
    const token = extractBearerToken(req);
    emvVault.revokeAdminSession(token);
    res.json({ ok: true });
  });

  app.put('/api/emv/entries', requireEmvAdmin, (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        scope?: EmvScope;
        employeeId?: string;
        key?: string;
        value?: string;
      };
      const session = emvVault.verifyAdminSession(extractBearerToken(req));
      const entry = emvVault.setEntry({
        scope: body.scope === 'employee' ? 'employee' : 'global',
        employeeId: body.employeeId,
        key: String(body.key || ''),
        value: String(body.value || ''),
        updatedBy: session?.label || '종합 관리자',
      });
      const publicData = emvVault.listPublicEntries();
      res.json({
        entry,
        globalEntries: publicData.globalEntries,
        employees: publicData.employees,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'EMV 키 저장에 실패했습니다.';
      res.status(400).json({ error: msg });
    }
  });

  app.delete('/api/emv/entries', requireEmvAdmin, (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        scope?: EmvScope;
        employeeId?: string;
        key?: string;
      };
      const deleted = emvVault.deleteEntry({
        scope: body.scope === 'employee' ? 'employee' : 'global',
        employeeId: body.employeeId,
        key: String(body.key || ''),
      });
      const publicData = emvVault.listPublicEntries();
      res.json({
        deleted,
        globalEntries: publicData.globalEntries,
        employees: publicData.employees,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'EMV 키 삭제에 실패했습니다.';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/emv/resolve-preview', requireEmvAdmin, (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as { employeeId?: string };
    res.json(emvVault.getResolutionPreview(body.employeeId));
  });

  // --- Compact Admin LLM Usage Telemetry Routes (Zero Secret Exposure) ---
  app.get('/api/admin/llm-usage', (req, res) => {
    const rawLimit = Number(req.query.limit);
    const limit = Number.isFinite(rawLimit) && rawLimit > 0 ? Math.min(50, rawLimit) : 12;
    const summary = getLlmUsageSummary(limit);
    const groqStatus = getGroqKeyPoolStatus();
    const claudeConfigured = Boolean((process.env.ANTHROPIC_API_KEY || '').trim());
    const geminiConfigured = Boolean((process.env.GEMINI_API_KEY || '').trim());
    res.json({
      ...summary,
      routingPolicy: {
        groqSlotsConfigured: groqStatus.configuredCount,
        groqSlotsHealthy: groqStatus.healthyCount,
        groqParallelMax: Math.max(1, groqStatus.configuredCount),
        claudeConfigured,
        geminiConfigured,
        defaultFastProvider:
          groqStatus.configuredCount > 0
            ? `Groq ${groqStatus.configuredCount}-키 병렬 (openai/gpt-oss-120b)`
            : claudeConfigured
              ? 'Claude Haiku 5.5 (Groq 미설정 폴백)'
              : '로컬 결정론적 엔진',
        advancedAnalysisProvider: claudeConfigured
          ? 'Claude 5.5 (Sonnet/Opus 자동 티어)'
          : groqStatus.configuredCount > 0
            ? 'Groq 5-키 폴백 (Claude 키 미설정)'
            : '로컬 결정론적 엔진',
      },
    });
  });

  app.post('/api/admin/llm-usage/reset', (_req, res) => {
    resetLlmUsageLedger();
    res.json({ ok: true, summary: getLlmUsageSummary(12) });
  });

  // --- Offline Local SQLite Daily Compressed Backup Snapshot Routes ---
  app.get('/api/admin/backups', (_req, res) => {
    res.json(localBackupScheduler.getStatus());
  });

  app.post('/api/admin/backups/run', (_req, res) => {
    const snapshot = localBackupScheduler.runSnapshotNow('manual');
    const status = localBackupScheduler.getStatus();
    if (!snapshot) {
      res.status(400).json({
        ok: false,
        error: status.lastError || 'SQLite 백업 스냅샷을 생성하지 못했습니다.',
        status,
      });
      return;
    }
    res.json({ ok: true, snapshot, status });
  });

  // --- Local Windows Outlook SQLite + FTS5 Store, Direct PST Attachment & Migration Routes ---
  let didAutoDiscoverBootstrapSettings = false;
  app.get('/api/local/bootstrap', (req, res) => {
    const rawLimit = Number(req.query.limit);
    const limit = Number.isFinite(rawLimit) && rawLimit > 0 ? Math.min(1000, rawLimit) : 300;
    let payload = localDb.getBootstrapPayload(limit);
    if (
      payload.available &&
      process.platform === 'win32' &&
      !didAutoDiscoverBootstrapSettings &&
      !payload.signatures.some((s) => s.htmlContent && s.htmlContent.includes('data:image/'))
    ) {
      didAutoDiscoverBootstrapSettings = true;
      try {
        discoverAndApplyAllOutlookSettings(localDb);
        payload = localDb.getBootstrapPayload(limit);
      } catch {
        // Keep bootstrap resilient if Registry/Signatures unavailable
      }
    } else {
      didAutoDiscoverBootstrapSettings = true;
    }
    res.json(payload);
  });

  app.get('/api/local/emails', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    const q = typeof req.query.q === 'string' ? req.query.q : undefined;
    const folder = typeof req.query.folder === 'string' ? req.query.folder : undefined;
    const store = typeof req.query.store === 'string' ? req.query.store : undefined;
    const includeDeleted = req.query.includeDeleted === 'true' || req.query.includeDeleted === '1';
    const deletedOnly = req.query.deletedOnly === 'true' || req.query.deletedOnly === '1';
    const limit = Number(req.query.limit) || 200;
    const offset = Number(req.query.offset) || 0;
    res.json(localDb.queryEmails({ q, folder, store, includeDeleted, deletedOnly, limit, offset }));
  });

  app.get('/api/local/emails/:id', (req, res) => {
    const email = localDb.getEmailById(req.params.id);
    if (!email) {
      res.status(404).json({ error: '요청한 메일을 찾을 수 없습니다.' });
      return;
    }
    backgroundMonitor.markViewed(req.params.id);
    res.json({ email });
  });

  app.patch('/api/local/emails/:id', (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      isRead?: boolean;
      isFlagged?: boolean;
      isDeleted?: boolean;
    };
    const updated = localDb.patchEmailState(req.params.id, {
      isRead: typeof body.isRead === 'boolean' ? body.isRead : undefined,
      isFlagged: typeof body.isFlagged === 'boolean' ? body.isFlagged : undefined,
      isDeleted: typeof body.isDeleted === 'boolean' ? body.isDeleted : undefined,
    });
    if (!updated) {
      res.status(404).json({ error: '수정할 메일을 찾을 수 없거나 변경 필드가 없습니다.' });
      return;
    }
    if (body.isRead === true) {
      backgroundMonitor.markViewed(req.params.id);
    }
    res.json({ ok: true });
  });

  const inferAccurateMimeType = (filename: string, rawMime?: string): string => {
    const cleanMime = (rawMime || '').trim().toLowerCase();
    if (cleanMime && cleanMime !== 'application/octet-stream' && cleanMime !== 'binary/octet-stream') {
      return cleanMime;
    }
    const ext = path.extname(filename || '').toLowerCase();
    switch (ext) {
      case '.png':
        return 'image/png';
      case '.jpg':
      case '.jpeg':
        return 'image/jpeg';
      case '.gif':
        return 'image/gif';
      case '.webp':
        return 'image/webp';
      case '.svg':
        return 'image/svg+xml';
      case '.bmp':
        return 'image/bmp';
      case '.pdf':
        return 'application/pdf';
      case '.xlsx':
        return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      case '.xls':
        return 'application/vnd.ms-excel';
      case '.csv':
        return 'text/csv; charset=utf-8';
      case '.docx':
        return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
      case '.doc':
        return 'application/msword';
      case '.txt':
      case '.md':
        return 'text/plain; charset=utf-8';
      case '.html':
      case '.htm':
        return 'text/html; charset=utf-8';
      default:
        return 'application/octet-stream';
    }
  };

  app.get('/api/local/pst/attachment', (req, res) => {
    const emailId = typeof req.query.emailId === 'string' ? req.query.emailId.trim() : '';
    const idx = Number(req.query.idx);
    const isInline = req.query.inline === '1' || req.query.inline === 'true';
    if (!emailId || !Number.isInteger(idx) || idx < 0) {
      res.status(400).json({ error: '유효한 emailId와 첨부파일 인덱스(idx)가 필요합니다.' });
      return;
    }
    const email = localDb.getEmailById(emailId);
    if (!email || !email.pstPath || email.descriptorNodeId === null) {
      res.status(404).json({ error: '원본 PST 메일 위치 정보를 찾을 수 없습니다.' });
      return;
    }
    if (!localDb.isAllowedPstPath(email.pstPath)) {
      res.status(403).json({ error: '등록되지 않은 PST 경로 접근이 차단되었습니다.' });
      return;
    }
    try {
      const att = readPstAttachmentBytes(email.pstPath, email.descriptorNodeId, idx);
      const mimeType = inferAccurateMimeType(att.filename, att.mimeType);
      res.setHeader('Content-Type', mimeType);
      res.setHeader(
        'Content-Disposition',
        `${isInline ? 'inline' : 'attachment'}; filename*=UTF-8''${encodeURIComponent(att.filename)}`
      );
      res.setHeader('Cache-Control', isInline ? 'private, max-age=600' : 'no-store');
      res.send(att.buffer);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'PST 첨부파일 추출 실패';
      res.status(400).json({ error: msg });
    }
  });

  app.get('/api/local/emails/:id/render', (req, res) => {
    const emailId = String(req.params.id || '').trim();
    if (!emailId) {
      res.status(400).json({ error: '유효한 메일 ID가 필요합니다.' });
      return;
    }
    if (!localDb.isReady()) {
      res.status(503).json({ error: '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    const cached = localDb.getEmailRichCache(emailId);
    backgroundMonitor.markViewed(emailId);
    if (cached && cached.bodyHtml) {
      res.json({
        id: emailId,
        bodyHtml: cached.bodyHtml,
        attachments: cached.attachments,
        fromCache: true,
      });
      return;
    }
    const email = localDb.getEmailById(emailId);
    if (!email) {
      res.status(404).json({ error: '해당 메일을 로컬 DB에서 찾을 수 없습니다.' });
      return;
    }
    if (
      email.pstPath &&
      email.descriptorNodeId !== null &&
      localDb.isAllowedPstPath(email.pstPath)
    ) {
      try {
        const rich = readPstMessageRichView(email.pstPath, email.descriptorNodeId, emailId);
        localDb.setEmailRichCache(emailId, rich.bodyHtml, rich.attachments);
        res.json({
          id: emailId,
          bodyHtml: rich.bodyHtml,
          attachments: rich.attachments,
          fromCache: false,
        });
        return;
      } catch {
        // Fall through to safe plain text HTML conversion below
      }
    }
    const safePlain = String(email.body || '(본문 없음)')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    const fallbackHtml = sanitizeEmailHtml(
      `<div style="white-space:pre-wrap;font-family:'Malgun Gothic','Segoe UI',sans-serif;line-height:1.65;color:#242424;">${safePlain}</div>`
    );
    res.json({
      id: emailId,
      bodyHtml: fallbackHtml,
      attachments: email.attachments || [],
      fromCache: false,
    });
  });

  app.post('/api/local/pst/attachment/parse', async (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      emailId?: string;
      idx?: number;
    };
    const emailId = typeof body.emailId === 'string' ? body.emailId.trim() : '';
    const idx = Number(body.idx);
    if (!emailId || !Number.isInteger(idx) || idx < 0) {
      res.status(400).json({ error: '유효한 emailId와 첨부파일 인덱스(idx)가 필요합니다.' });
      return;
    }
    const email = localDb.getEmailById(emailId);
    if (!email || !email.pstPath || email.descriptorNodeId === null) {
      res.status(404).json({ error: '원본 PST 메일 위치 정보를 찾을 수 없습니다.' });
      return;
    }
    if (!localDb.isAllowedPstPath(email.pstPath)) {
      res.status(403).json({ error: '등록되지 않은 PST 경로 접근이 차단되었습니다.' });
      return;
    }

    let tmpFile = '';
    try {
      const cfg = readAppConfig();
      const att = readPstAttachmentBytes(email.pstPath, email.descriptorNodeId, idx);
      const ext = path.extname(att.filename) || '.bin';
      tmpFile = path.join(os.tmpdir(), `bmail-pst-att-${crypto.randomUUID()}${ext}`);
      fs.writeFileSync(tmpFile, att.buffer);
      const parsed = await parseUploadedDocument(
        {
          path: tmpFile,
          originalname: att.filename,
          mimetype: att.mimeType,
          size: att.buffer.length,
        },
        cfg.externalEngines
      );
      res.json({ document: parsed });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'PST 첨부파일 로컬 파싱 실패';
      res.status(400).json({ error: msg });
    } finally {
      if (tmpFile) {
        try {
          fs.unlinkSync(tmpFile);
        } catch {
          // ignore cleanup failure
        }
      }
    }
  });

  app.post('/api/local/files/parse-paths', async (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      paths?: unknown;
    };
    const rawPaths = Array.isArray(body.paths)
      ? body.paths.filter((p): p is string => typeof p === 'string' && p.trim().length > 0).slice(0, 15)
      : [];
    if (rawPaths.length === 0) {
      res.status(400).json({ error: '파싱할 로컬 파일 경로(paths)가 필요합니다.' });
      return;
    }

    const cfg = readAppConfig();
    const MAX_LOCAL_FILE_BYTES = 50 * 1024 * 1024;
    const documents: ParsedDocumentSummary[] = [];
    const errors: Array<{ filename: string; error: string }> = [];

    for (const rawPath of rawPaths) {
      const resolved = path.resolve(rawPath.trim());
      const basename = path.basename(resolved) || resolved;
      try {
        if (!fs.existsSync(resolved)) {
          throw new Error('파일을 찾을 수 없습니다.');
        }
        const st = fs.statSync(resolved);
        if (!st.isFile()) {
          throw new Error('일반 파일만 직접 파싱할 수 있습니다.');
        }
        if (st.size > MAX_LOCAL_FILE_BYTES) {
          throw new Error('단일 파일 최대 크기(50MB)를 초과했습니다.');
        }
        const parsed = await parseUploadedDocument(
          {
            path: resolved,
            originalname: basename,
            mimetype: 'application/octet-stream',
            size: st.size,
          },
          cfg.externalEngines
        );
        documents.push(parsed);
      } catch (err) {
        errors.push({
          filename: basename,
          error: err instanceof Error ? err.message : '로컬 파일 파싱 실패',
        });
      }
    }

    if (documents.length === 0 && errors.length > 0) {
      res.status(400).json({
        error: errors.map((e) => `${e.filename}: ${e.error}`).join(' / '),
        errors,
      });
      return;
    }

    res.json({ documents, errors });
  });

  app.get('/api/local/contacts', (_req, res) => {
    res.json({ contacts: localDb.getAllContacts() });
  });

  app.get('/api/local/state/:key', (req, res) => {
    try {
      const key = String(req.params.key || '').trim();
      if (!localDb.isAllowedStateKey(key)) {
        res.status(400).json({ error: `허용되지 않은 상태 키입니다: ${key}` });
        return;
      }
      res.json({ key, value: localDb.getUserState(key) });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '로컬 상태 조회 실패';
      res.status(400).json({ error: msg });
    }
  });

  app.put('/api/local/state', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        key?: string;
        value?: unknown;
      };
      if (!body.key) {
        res.status(400).json({ error: '저장할 상태 key가 필요합니다.' });
        return;
      }
      localDb.setUserState(String(body.key), body.value);
      res.json({ ok: true });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '로컬 상태 저장 실패';
      res.status(400).json({ error: msg });
    }
  });

  let migrationChild: ChildProcess | null = null;
  const stopMigrationWorker = () => {
    if (!migrationChild) return;
    try {
      migrationChild.kill();
    } catch {
      // ignore
    }
    migrationChild = null;
  };
  process.once('exit', stopMigrationWorker);
  process.once('SIGTERM', () => {
    stopMigrationWorker();
    process.exit(0);
  });

  const startBackgroundMigration = (extraPstPaths?: string[]) => {
    if (migrationChild || migrationProgress.status === 'running') return;

    migrationProgress = {
      status: 'running',
      startedAt: new Date().toISOString(),
      completedAt: null,
      storesProcessed: 0,
      totalEmailsIndexed: 0,
      totalContactsIndexed: 0,
      totalFoldersIndexed: 0,
      totalSignaturesIndexed: 0,
      discoveredProjectsCount: 0,
      error: null,
    };

    const compiledWorker = path.resolve(APP_ROOT, 'dist-server', 'pst-worker.mjs');
    const sourceWorker = path.resolve(APP_ROOT, 'src', 'services', 'pstMigrationWorker.ts');
    const useCompiled = fs.existsSync(compiledWorker);
    const workerArgs = useCompiled
      ? [compiledWorker]
      : [path.resolve(APP_ROOT, 'node_modules', 'tsx', 'dist', 'cli.mjs'), sourceWorker];

    const child = spawn(process.execPath, workerArgs, {
      cwd: APP_ROOT,
      env: {
        ...process.env,
        BMAIL_DATA_DIR: DATA_DIR,
        BMAIL_EXTRA_PST_PATHS_JSON:
          extraPstPaths && extraPstPaths.length > 0 ? JSON.stringify(extraPstPaths) : '',
      },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    });
    migrationChild = child;

    let lineBuffer = '';
    child.stdout?.setEncoding('utf8');
    child.stdout?.on('data', (chunk: string) => {
      lineBuffer += chunk;
      const lines = lineBuffer.split(/\r?\n/);
      lineBuffer = lines.pop() ?? '';
      for (const line of lines) {
        if (!line.startsWith('__BMAIL_MIGRATION_PROGRESS__:')) continue;
        try {
          const parsed = JSON.parse(
            line.slice('__BMAIL_MIGRATION_PROGRESS__:'.length)
          ) as MigrationProgressSummary;
          if (parsed && typeof parsed.status === 'string') {
            migrationProgress = parsed;
          }
        } catch {
          // ignore malformed progress JSON line
        }
      }
    });

    let workerStderr = '';
    child.stderr?.setEncoding('utf8');
    child.stderr?.on('data', (chunk: string) => {
      if (workerStderr.length < 4096) workerStderr += chunk;
    });

    child.once('error', (err) => {
      migrationChild = null;
      migrationProgress = {
        ...migrationProgress,
        status: 'failed',
        completedAt: new Date().toISOString(),
        error: err.message,
      };
    });

    child.once('close', (code) => {
      migrationChild = null;
      if (migrationProgress.status === 'running') {
        migrationProgress = {
          ...migrationProgress,
          status: code === 0 ? 'completed' : 'failed',
          completedAt: new Date().toISOString(),
          error:
            code === 0
              ? null
              : workerStderr.trim().slice(0, 300) ||
                `PST 워커가 종료 코드 ${code ?? 'unknown'}로 중단되었습니다.`,
        };
      }
    });
  };

  app.post('/api/local/outlook/migrate', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 SQLite DB를 초기화할 수 없습니다.' });
      return;
    }
    if (migrationProgress.status === 'running') {
      res.status(202).json({ started: false, alreadyRunning: true, progress: migrationProgress });
      return;
    }

    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      folderPath?: string;
    };
    let extraPstPaths: string[] | undefined;
    if (typeof body.folderPath === 'string' && body.folderPath.trim()) {
      const targetDir = path.resolve(body.folderPath.trim());
      if (!fs.existsSync(targetDir) || !fs.statSync(targetDir).isDirectory()) {
        res.status(400).json({ error: `지정한 로컬 폴더를 찾을 수 없습니다: ${targetDir}` });
        return;
      }
      const entries = fs.readdirSync(targetDir, { withFileTypes: true });
      extraPstPaths = entries
        .filter(
          (e) =>
            e.isFile() &&
            !e.name.startsWith('~') &&
            /\.(pst|ost)$/i.test(e.name)
        )
        .map((e) => path.join(targetDir, e.name));
      if (extraPstPaths.length === 0) {
        res.status(400).json({ error: `지정한 폴더에서 .pst 또는 .ost 파일을 찾지 못했습니다: ${targetDir}` });
        return;
      }
    }

    startBackgroundMigration(extraPstPaths);
    res.status(202).json({
      started: true,
      scannedFiles: extraPstPaths ? extraPstPaths.length : 0,
      progress: migrationProgress,
    });
  });

  app.post('/api/local/outlook/recompute-projects', (_req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    const projects = recomputeProjectsFromLocalDb(localDb);
    res.json({ ok: true, projectsCount: projects.length });
  });

  app.get('/api/local/outlook/migrate/status', (_req, res) => {
    const trackedPaths = localDb.isReady() ? localDb.getState<string[]>('pstPaths') || [] : [];
    const savedStates = localDb.isReady() ? localDb.getState<LocalPstFileState[]>('pst_file_states') || [] : [];
    const currentStates = trackedPaths.length > 0 ? inspectPstFileStates(trackedPaths) : [];
    const savedMap = new Map(savedStates.map((s) => [s.path.toLowerCase(), s]));
    const hasNewerPstChanges =
      currentStates.length > 0 &&
      (savedStates.length === 0 ||
        currentStates.some((c) => {
          const prev = savedMap.get(c.path.toLowerCase());
          return !prev || prev.size !== c.size || Math.abs(prev.mtimeMs - c.mtimeMs) > 2000;
        }));
    res.json({ progress: migrationProgress, hasNewerPstChanges });
  });

  app.post('/api/local/outlook/sync-live', (_req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const summary = syncIncrementalLocalOutlook(localDb);
      res.json(summary);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '실시간 Outlook 동기화 실패';
      res.status(500).json({ error: msg });
    }
  });

  const handleDiscoverOutlookSettings = (_req: express.Request, res: express.Response) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const report = discoverAndApplyAllOutlookSettings(localDb);
      res.json(report);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Outlook 전체 설정 탐색 실패';
      res.status(500).json({ error: msg });
    }
  };
  app.get('/api/local/outlook/discover-settings', handleDiscoverOutlookSettings);
  app.post('/api/local/outlook/discover-settings', handleDiscoverOutlookSettings);

  let incrementalSyncInFlight: Promise<Record<string, unknown>> | null = null;
  app.post('/api/local/outlook/sync-incremental', async (_req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      if (!incrementalSyncInFlight) {
        incrementalSyncInFlight = (async () => {
          const comRes = await pollLiveOutlookInboxAsync(localDb, {
            maxItems: 25,
            timeoutMs: 12_000,
          });
          const syncedAt = new Date().toISOString();
          const totalEmails = localDb.queryEmails({ limit: 1 }).total;
          if (comRes.newEmails.length > 0) {
            void backgroundMonitor.ingestNewEmails(comRes.newEmails);
          }
          return {
            ok: true,
            outlookRunning: comRes.outlookRunning,
            addedPstEmails: 0,
            addedComEmails: comRes.newEmails.length,
            skippedUnchangedPsts: 0,
            syncedFolders: 0,
            favoritesCount: 0,
            categoriesCount: 0,
            rulesCount: 0,
            totalEmails,
            newEmails: comRes.newEmails,
            syncedAt,
          };
        })().finally(() => {
          incrementalSyncInFlight = null;
        });
      }
      const summary = await incrementalSyncInFlight;
      res.json(summary);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '증분 Outlook 동기화 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.get('/api/local/dashboard/executive', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const rawPeriod = typeof req.query.periodType === 'string' ? req.query.periodType : 'all';
      const periodType: ExecutivePeriodType =
        rawPeriod === 'weekly' || rawPeriod === 'monthly' || rawPeriod === 'custom'
          ? rawPeriod
          : 'all';
      const dateFrom = typeof req.query.dateFrom === 'string' ? req.query.dateFrom.trim() : undefined;
      const dateTo = typeof req.query.dateTo === 'string' ? req.query.dateTo.trim() : undefined;
      const projectCode =
        typeof req.query.projectCode === 'string' ? req.query.projectCode.trim() : undefined;
      const topic = typeof req.query.topic === 'string' ? req.query.topic.trim() : undefined;
      const rawDir = typeof req.query.direction === 'string' ? req.query.direction : 'all';
      const direction: 'all' | 'received' | 'sent' =
        rawDir === 'received' || rawDir === 'sent' ? rawDir : 'all';

      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;

      const report = buildExecutiveDashboardReport(
        localDb,
        { periodType, dateFrom, dateTo, projectCode, topic, direction },
        projects,
        overrides
      );
      res.json(report);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '대시보드 분석 리포트 생성 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.get('/api/local/unread-briefing', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const rawLimit = Number(req.query.limit || 250);
      const limit = Number.isFinite(rawLimit) ? Math.max(20, Math.min(500, rawLimit)) : 250;
      const unread = localDb.queryUnreadInboxEmails(limit);
      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      res.json(buildUnreadPriorityBriefing(unread, projects, overrides));
    } catch (err) {
      const msg = err instanceof Error ? err.message : '미열람 메일 브리핑 생성 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/local/dashboard/executive/ai-narrative', async (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        periodType?: ExecutivePeriodType;
        dateFrom?: string;
        dateTo?: string;
        projectCode?: string;
        topic?: string;
        direction?: 'all' | 'received' | 'sent';
      };
      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      const report = buildExecutiveDashboardReport(localDb, body, projects, overrides);
      const narrative = await generateAiExecutiveNarrative(report, process.env.ANTHROPIC_API_KEY);
      res.json({ narrativeReport: narrative });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'AI 분석글 작성 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/local/dashboard/executive/save', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        title?: string;
        periodType?: ExecutivePeriodType;
        dateFrom?: string;
        dateTo?: string;
        projectCode?: string;
        topic?: string;
        direction?: 'all' | 'received' | 'sent';
      };
      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      const report = buildExecutiveDashboardReport(localDb, body, projects, overrides);
      const savedReports = saveExecutiveReportSnapshot(localDb, report, body.title);
      res.json({ ok: true, savedReports });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '보고서 스냅샷 저장 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/local/mail/assistant', async (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      email?: Partial<MailItem>;
      mode?: 'analyze' | 'reply' | 'chat' | 'translate' | 'actions';
      prompt?: string;
      targetLang?: string;
      tone?: string;
      history?: Array<{ role: string; content: string }>;
      highPerformance?: boolean;
    };
    const mail = body.email || {};
    const mode = body.mode || 'analyze';
    const subject = String(mail.subject || '(제목 없음)').trim();
    const sender = String(mail.fromName || mail.fromEmail || '발신자').trim();
    const textBody = String(mail.body || mail.snippet || '').trim();
    const drawings = extractShipbuildingDrawingRefs(`${subject}\n${textBody}`).map(
      (d) => `${d.raw}${d.revision ? ` (${d.revision})` : ''}`
    );
    const deadlines = Array.from(
      new Set([
        ...(textBody.match(/\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b/g) || []),
        ...(textBody.match(/\d{1,2}\s*월\s*\d{1,2}\s*일/g) || []),
      ])
    ).slice(0, 4);
    const isUrgent =
      mail.importance === 'high' || /긴급|urgent|asap|미결|금일|즉시/i.test(`${subject} ${textBody}`);

    const buildOfflineAssistantResult = () => {
      const summaryLines = [
        `[${isUrgent ? '긴급·우선 검토' : '일반 업무'}] ${sender} 측에서 "${subject}" 관련 메일을 송부했습니다.`,
        drawings.length > 0
          ? `식별된 도면·문서 코드: ${drawings.join(', ')}`
          : '식별된 도면 번호 없음 (일반 사양/업무 협의)',
        deadlines.length > 0
          ? `언급된 주요 일정·기한: ${deadlines.join(', ')}`
          : '명시적 마감일 미검출 (본문 확인 권장)',
      ];
      const actionChecklist = [
        '첨부 문서(POS / 도면 / 단가표) 개정 내역 및 사양 일치 여부 확인',
        deadlines.length > 0
          ? `기한(${deadlines[0]}) 내 검토 코멘트 및 승인 여부 회신`
          : '담당 부서 검토 후 회신 일정 안내',
      ];
      const replyDrafts = [
        {
          id: 'formal',
          toneLabel: '정중·공식 회신',
          subject: subject.startsWith('RE:') ? subject : `RE: ${subject}`,
          body: `안녕하세요, ${sender}님.\n\n보내주신 "${subject}" 관련 메일 및 첨부 자료 잘 받았습니다.\n요청하신 사항을 내부 검토한 후 조속히 확정 회신드리겠습니다.\n\n감사합니다.`,
        },
        {
          id: 'technical',
          toneLabel: '기술·도면 검토 회신',
          subject: subject.startsWith('RE:') ? subject : `RE: ${subject}`,
          body: `안녕하세요, ${sender}님.\n\n송부해 주신 "${subject}"${drawings.length > 0 ? ` (${drawings.join(', ')})` : ''} 자료 접수하여 사양 및 개정 사항을 확인 중입니다.\n검토 완료 즉시 코멘트 반영하여 회신드리겠습니다.\n\n감사합니다.`,
        },
        {
          id: 'concise',
          toneLabel: '간결·일정 조율 회신',
          subject: subject.startsWith('RE:') ? subject : `RE: ${subject}`,
          body: `${sender}님, 안녕하세요.\n\n"${subject}" 건 확인했습니다. ${body.prompt ? `${body.prompt}\n` : '내부 확인 후 금주 내로 회신드리겠습니다.'}\n\n감사합니다.`,
        },
      ];
      const targetLang = body.targetLang || 'ko';
      const translatedText =
        targetLang === 'en'
          ? `[English Summary & Translation]\nSubject: ${subject}\nFrom: ${sender}\n\n${textBody.slice(0, 1200) || '(No body)'}`
          : `[한국어 요약·번역 해설]\n제목: ${subject}\n발신: ${sender}\n\n${summaryLines.join('\n')}\n\n--- 원문 발췌 ---\n${textBody.slice(0, 1200)}`;
      const chatAnswer = body.prompt
        ? `"${subject}" 메일 기준으로 질문("${body.prompt}")을 검토했습니다.\n- 발신자: ${sender}\n- 핵심 요지: ${summaryLines[0]}\n-${drawings.length > 0 ? ` 관련 도면: ${drawings.join(', ')}` : ' 본문 첫 단락: ' + textBody.slice(0, 180)}`
        : summaryLines.join('\n');

      return {
        ok: true,
        mode,
        model: 'deterministic-local',
        modelLabel: '로컬 결정론적 엔진 (무오류 오프라인 보장)',
        tier: 'low',
        usedFallback: true,
        urgency: isUrgent ? 'high' : 'normal',
        drawings,
        deadlines,
        summaryLines,
        actionChecklist,
        replyDrafts,
        translatedText,
        chatAnswer,
      };
    };

    const fallbackPayload = buildOfflineAssistantResult();
    const claudeApiKey = (process.env.ANTHROPIC_API_KEY || '').trim();
    const groqConfigured = getGroqKeyPoolStatus().configuredCount > 0;
    if (!groqConfigured && !claudeApiKey) {
      res.json(fallbackPayload);
      return;
    }

    try {
      const routingInput = {
        emailCount: 1,
        promptChars: subject.length + textBody.length + (body.prompt?.length || 0),
        customFocus: body.prompt || mode,
      };
      const highPerformance = shouldEscalateToClaude({
        highPerformance: Boolean(body.highPerformance),
        routing: routingInput,
        promptChars: routingInput.promptChars,
      });

      const systemPrompt =
        '당신은 조선해양·설계·영업 Outlook AI 어시스턴트입니다. 반드시 순수 JSON 객체만 출력하세요. 키 규격: {"summaryLines":["..."],"actionChecklist":["..."],"replyDrafts":[{"id":"formal","toneLabel":"정중·공식 회신","subject":"...","body":"..."},{"id":"technical","toneLabel":"기술·도면 검토 회신","subject":"...","body":"..."},{"id":"concise","toneLabel":"간결·일정 조율 회신","subject":"...","body":"..."}],"translatedText":"...","chatAnswer":"..."}';
      const userPrompt = JSON.stringify({
        mode,
        targetLang: body.targetLang || 'ko',
        tone: body.tone || 'formal',
        userQuestion: body.prompt || '',
        history: (body.history || []).slice(-6),
        email: {
          subject,
          sender,
          date: mail.date || '',
          drawings,
          deadlines,
          body: textBody.slice(0, 4500),
        },
      });

      let rawResponseText = '';
      let usedModelId = '';
      let usedModelLabel = '';
      let usedTier: 'low' | 'medium' | 'high' = 'medium';

      // Default: Groq 5-key streaming pool first unless highPerformance requires Claude
      if (!highPerformance && groqConfigured) {
        try {
          const groqRes = await invokeGroqStreamingCompletion({
            systemPrompt,
            userPrompt: userPrompt.slice(0, 6000),
            jsonMode: true,
            maxCompletionTokens: 2048,
            feature: 'mail-assistant',
          });
          rawResponseText = groqRes.text;
          usedModelId = groqRes.model;
          usedModelLabel = `${groqRes.model} (5-키 순환 기본)`;
          usedTier = 'medium';
        } catch {
          // Fail over to Claude below
        }
      }

      if (!rawResponseText && claudeApiKey) {
        const claudeRes = await invokeClaudeMessagesApi({
          apiKey: claudeApiKey,
          routing: routingInput,
          systemPrompt,
          userPrompt,
          maxTokens: 2048,
        });
        rawResponseText = claudeRes.text;
        usedModelId = claudeRes.model;
        usedModelLabel = `Claude · ${claudeRes.model} (${claudeRes.decision.tierLabel})`;
        usedTier = claudeRes.decision.tier;
      }

      if (!rawResponseText) {
        res.json(fallbackPayload);
        return;
      }

      const jsonMatch = rawResponseText.match(/\{[\s\S]*\}/);
      const parsed = jsonMatch ? (JSON.parse(jsonMatch[0]) as Record<string, unknown>) : {};

      res.json({
        ...fallbackPayload,
        model: usedModelId,
        modelLabel: usedModelLabel,
        tier: usedTier,
        usedFallback: false,
        summaryLines:
          Array.isArray(parsed.summaryLines) && parsed.summaryLines.length > 0
            ? parsed.summaryLines.map(String)
            : fallbackPayload.summaryLines,
        actionChecklist:
          Array.isArray(parsed.actionChecklist) && parsed.actionChecklist.length > 0
            ? parsed.actionChecklist.map(String)
            : fallbackPayload.actionChecklist,
        replyDrafts:
          Array.isArray(parsed.replyDrafts) && parsed.replyDrafts.length > 0
            ? parsed.replyDrafts
            : fallbackPayload.replyDrafts,
        translatedText:
          typeof parsed.translatedText === 'string' && parsed.translatedText.trim()
            ? parsed.translatedText.trim()
            : fallbackPayload.translatedText,
        chatAnswer:
          typeof parsed.chatAnswer === 'string' && parsed.chatAnswer.trim()
            ? parsed.chatAnswer.trim()
            : fallbackPayload.chatAnswer,
      });
    } catch {
      // Never return 5xx or trigger Failed to fetch: gracefully serve deterministic result
      res.json(fallbackPayload);
    }
  });

  app.post('/api/local/mail/quick-summary', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        emails?: MailItem[];
        emailIds?: string[];
        enableAi?: boolean;
      };
      let targets: MailItem[] = Array.isArray(body.emails)
        ? body.emails.filter((m): m is MailItem => Boolean(m && m.id && m.subject))
        : [];
      if (targets.length === 0 && Array.isArray(body.emailIds) && localDb.isReady()) {
        targets = body.emailIds
          .map((id) => localDb.getEmailById(String(id || '')))
          .filter((m): m is NonNullable<typeof m> => Boolean(m));
      }
      const summaries = await summarizeNewMailsForNotification(targets, {
        maxItems: 5,
        enableAi: body.enableAi !== false,
      });
      res.json({ ok: true, summaries });
    } catch {
      res.json({ ok: true, summaries: [] });
    }
  });

  // --- Backend Proactive Mail Monitor, Pre-Organized Dashboard Digest & SGW Mobile Alert Endpoints ---
  app.get('/api/local/monitor/status', (_req, res) => {
    res.json(backgroundMonitor.getStatus());
  });

  app.put('/api/local/monitor/settings', (req, res) => {
    try {
      const patch = req.body && typeof req.body === 'object' ? req.body : {};
      const next = backgroundMonitor.updateSettings(patch);
      res.json({ ok: true, settings: next, status: backgroundMonitor.getStatus() });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '모니터 설정 저장 실패';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/local/monitor/ack', (req, res) => {
    const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
      emailId?: string;
    };
    if (body.emailId) {
      backgroundMonitor.markViewed(String(body.emailId));
    }
    res.json({ ok: true });
  });

  app.post('/api/local/monitor/poll-now', async (_req, res) => {
    try {
      const result = await backgroundMonitor.runBackgroundPollCycle();
      res.json({ ok: true, ...result, status: backgroundMonitor.getStatus() });
    } catch (err) {
      const msg = err instanceof Error ? err.message : '백그라운드 수신 점검 실패';
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/local/monitor/test-sgw-alert', async (_req, res) => {
    try {
      const result = await backgroundMonitor.sendTestSgwAlert();
      res.json(result);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'SGW 모바일 테스트 알림 실패';
      res.status(400).json({ ok: false, reason: msg });
    }
  });

  app.post('/api/local/monitor/consume-tray-alerts', (_req, res) => {
    res.json({ alerts: backgroundMonitor.consumeUnnotifiedTrayAlerts() });
  });

  app.post('/api/local/attachments/export-plan', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const options = (req.body && typeof req.body === 'object'
        ? req.body
        : {}) as AttachmentExportFilterOptions;
      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      const plan = buildAttachmentExportPlan(localDb, options, projects, overrides);
      res.json(plan);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '첨부파일 저장 계획 생성 실패';
      res.status(400).json({ error: msg });
    }
  });

  app.post('/api/local/attachments/export-run', (req, res) => {
    if (!localDb.isReady()) {
      res.status(503).json({ error: localDb.getInitError() || '로컬 메일 DB가 준비되지 않았습니다.' });
      return;
    }
    try {
      const options = (req.body && typeof req.body === 'object'
        ? req.body
        : {}) as AttachmentExportFilterOptions;
      const projects = localDb.getState<ProjectDefinition[]>('projects') || [];
      const overrides =
        localDb.getState<ProjectMailOverrides>('projectMailOverrides') ||
        EMPTY_PROJECT_MAIL_OVERRIDES;
      const result = executeAttachmentExportPlan(localDb, options, projects, overrides);
      res.json(result);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '첨부파일 일괄 저장 및 리포트 생성 실패';
      res.status(400).json({ error: msg });
    }
  });

  // --- SGW Live Chat Bridge (https://sgw.seastar.work) & Desktop Update Check ---
  const sendSgwChatErr = (res: express.Response, err: unknown) => {
    if (err instanceof SgwBridgeError) {
      res.status(err.status).json({ error: err.message });
      return;
    }
    res.status(500).json({ error: err instanceof Error ? err.message : 'SGW 채팅 연동 오류가 발생했습니다.' });
  };

  app.get('/api/sgw-chat/status', (_req, res) => {
    res.json(sgwChatBridge.getStatus());
  });

  app.post('/api/sgw-chat/session-connect', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        sgwBaseUrl?: string;
        token?: string;
        cookieHeader?: string;
        user?: Record<string, unknown>;
      };
      const status = await sgwChatBridge.connectWithSession({
        sgwBaseUrl: typeof body.sgwBaseUrl === 'string' ? body.sgwBaseUrl : undefined,
        token: typeof body.token === 'string' ? body.token : undefined,
        cookieHeader: typeof body.cookieHeader === 'string' ? body.cookieHeader : undefined,
      });
      res.json(status);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/login', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        email?: string;
        password?: string;
        sgwBaseUrl?: string;
      };
      const status = await sgwChatBridge.login({
        email: String(body.email || ''),
        password: String(body.password || ''),
        sgwBaseUrl: typeof body.sgwBaseUrl === 'string' ? body.sgwBaseUrl : undefined,
      });
      res.json(status);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/logout', (_req, res) => {
    sgwChatBridge.clearSession();
    res.json({ ok: true });
  });

  app.get('/api/sgw-chat/employees', async (req, res) => {
    try {
      const employees = await sgwChatBridge.listEmployees(req.query.refresh === '1');
      res.json({ employees });
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.get('/api/sgw-chat/rooms', async (_req, res) => {
    try {
      res.json(await sgwChatBridge.listRooms());
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/rooms', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        participants?: string[];
        roomId?: string;
        lastMessage?: string;
      };
      const result = await sgwChatBridge.createOrOpenRoom({
        participants: Array.isArray(body.participants) ? body.participants : [],
        roomId: typeof body.roomId === 'string' ? body.roomId : undefined,
        lastMessage: typeof body.lastMessage === 'string' ? body.lastMessage : undefined,
      });
      res.json(result);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.get('/api/sgw-chat/rooms/:roomId/messages', async (req, res) => {
    try {
      const msgId = typeof req.query.msgId === 'string' ? req.query.msgId : undefined;
      res.json(await sgwChatBridge.getRoomMessages(req.params.roomId, msgId));
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/rooms/:roomId/messages', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        content?: string;
        fileName?: string;
        fileData?: string;
        replyToId?: string;
        mentions?: string[];
      };
      res.json(
        await sgwChatBridge.sendRoomMessage(req.params.roomId, {
          content: String(body.content || ''),
          fileName: body.fileName,
          fileData: body.fileData,
          replyToId: body.replyToId,
          mentions: body.mentions,
        })
      );
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/rooms/:roomId/read', async (req, res) => {
    try {
      const lastId = String((req.body as { lastReadMessageId?: string })?.lastReadMessageId || '');
      res.json(await sgwChatBridge.markRoomRead(req.params.roomId, lastId));
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/rooms/:roomId/react', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        id?: string;
        emoji?: string;
        on?: boolean;
      };
      res.json(
        await sgwChatBridge.toggleRoomReaction(
          req.params.roomId,
          String(body.id || ''),
          String(body.emoji || ''),
          body.on !== false
        )
      );
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.get('/api/sgw-chat/dm', async (req, res) => {
    try {
      const since = req.query.since ? Number(req.query.since) : undefined;
      const id = typeof req.query.id === 'string' ? req.query.id : undefined;
      res.json(await sgwChatBridge.listDirectMessages({ since, id }));
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/dm', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        receiverId?: string;
        receiverName?: string;
        content?: string;
        fileName?: string;
        fileData?: string;
        replyToId?: string;
      };
      res.json(
        await sgwChatBridge.sendDirectMessage({
          receiverId: String(body.receiverId || ''),
          receiverName: body.receiverName,
          content: String(body.content || ''),
          fileName: body.fileName,
          fileData: body.fileData,
          replyToId: body.replyToId,
        })
      );
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/dm/read', async (req, res) => {
    try {
      const otherId = String(
        (req.body as { otherId?: string })?.otherId || req.query.otherId || ''
      );
      res.json(await sgwChatBridge.markDmRead(otherId));
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-chat/share-mail', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        emailId?: string;
        mail?: Partial<MailItem> & { id?: string; subject?: string };
        targetType?: 'room' | 'dm';
        roomId?: string;
        receiverId?: string;
        receiverName?: string;
        note?: string;
        attachmentIdx?: number;
      };
      const emailId = typeof body.emailId === 'string' ? body.emailId.trim() : '';
      const storedMail = emailId && localDb.isReady() ? localDb.getEmailById(emailId) : null;
      const targetMail = storedMail || (body.mail?.id ? (body.mail as Partial<MailItem> & { id: string; subject: string }) : null);
      if (!targetMail) {
        res.status(404).json({ error: '공유할 메일 정보를 찾을 수 없습니다.' });
        return;
      }

      const content = buildMailShareText(targetMail, { note: body.note, includeDeepLink: true });
      let fileName: string | undefined;
      let fileData: string | undefined;

      if (
        body.attachmentIdx !== undefined &&
        body.attachmentIdx !== null &&
        Number.isInteger(Number(body.attachmentIdx)) &&
        Number(body.attachmentIdx) >= 0
      ) {
        if (storedMail?.pstPath && storedMail.descriptorNodeId !== null) {
          if (!localDb.isAllowedPstPath(storedMail.pstPath)) {
            res.status(403).json({ error: '등록되지 않은 PST 경로 접근이 차단되었습니다.' });
            return;
          }
          const extracted = readPstAttachmentBytes(
            storedMail.pstPath,
            storedMail.descriptorNodeId,
            Number(body.attachmentIdx)
          );
          if (extracted.buffer.length > SGW_MAX_RAW_ATTACHMENT_BYTES) {
            res.status(413).json({ error: 'SGW 채팅 직접 첨부는 5MB 이하 파일만 지원됩니다.' });
            return;
          }
          fileName = extracted.filename;
          fileData = extracted.buffer.toString('base64');
        }
      }

      if (body.targetType === 'dm') {
        const result = await sgwChatBridge.sendDirectMessage({
          receiverId: String(body.receiverId || ''),
          receiverName: body.receiverName,
          content,
          fileName,
          fileData,
        });
        res.json(result);
        return;
      }

      const result = await sgwChatBridge.sendRoomMessage(String(body.roomId || ''), {
        content,
        fileName,
        fileData,
      });
      res.json(result);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.get('/api/desktop/update-check', async (_req, res) => {
    let currentVersion = '1.1.0';
    try {
      const pkgRaw = fs.readFileSync(path.resolve(APP_ROOT, 'package.json'), 'utf8');
      const parsed = JSON.parse(pkgRaw) as { version?: string };
      if (parsed.version) currentVersion = parsed.version;
    } catch {
      // ignore
    }

    const manifestUrl =
      process.env.BMAIL_UPDATE_URL?.trim() || 'https://mail.seastar.work/releases/latest.json';
    const parseParts = (v: string) =>
      String(v || '')
        .trim()
        .replace(/^v/i, '')
        .split('-')[0]
        .split('.')
        .map((n) => Number.parseInt(n, 10) || 0);
    const isNewer = (latest: string, current: string) => {
      const a = parseParts(latest);
      const b = parseParts(current);
      for (let i = 0; i < 3; i += 1) {
        if ((a[i] || 0) !== (b[i] || 0)) return (a[i] || 0) > (b[i] || 0);
      }
      return false;
    };

    try {
      const upstream = await fetch(manifestUrl, {
        headers: { Accept: 'application/json' },
        signal: AbortSignal.timeout(6000),
      });
      if (!upstream.ok) {
        res.json({
          currentVersion,
          latestVersion: currentVersion,
          updateAvailable: false,
          error: `업데이트 서버 응답 HTTP ${upstream.status}`,
        });
        return;
      }
      const manifest = (await upstream.json()) as {
        version?: string;
        downloadUrl?: string;
        latestDownloadUrl?: string;
        releaseNotes?: string;
        sizeBytes?: number;
        publishedAt?: string;
      };
      const latestVersion = String(manifest.version || currentVersion).trim();
      const downloadUrl = String(manifest.latestDownloadUrl || manifest.downloadUrl || '').trim();
      res.json({
        currentVersion,
        latestVersion,
        updateAvailable: isNewer(latestVersion, currentVersion),
        releaseNotes: manifest.releaseNotes || '',
        downloadUrl,
        sizeBytes: Number(manifest.sizeBytes || 0),
        publishedAt: manifest.publishedAt || '',
      });
    } catch (err) {
      res.json({
        currentVersion,
        latestVersion: currentVersion,
        updateAvailable: false,
        error: err instanceof Error ? err.message : '업데이트 확인에 실패했습니다.',
      });
    }
  });

  // --- SGW Live Mail Ledger + Weekly Sync & SEDMS (sedms.app) Live Bridge Routes ---
  app.get('/api/sgw-bridge/ledger', async (req, res) => {
    try {
      const data = await sgwChatBridge.getLedgerDashboard({
        project: typeof req.query.project === 'string' ? req.query.project : undefined,
        direction: typeof req.query.direction === 'string' ? req.query.direction : undefined,
        q: typeof req.query.q === 'string' ? req.query.q : undefined,
      });
      res.json(data);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-bridge/ledger/sync', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        emailIds?: string[];
        limit?: number;
        emails?: MailItem[];
      };
      let targetEmails: MailItem[] = [];
      if (Array.isArray(body.emails) && body.emails.length > 0) {
        targetEmails = body.emails.slice(0, 100);
      } else if (localDb.isReady()) {
        if (Array.isArray(body.emailIds) && body.emailIds.length > 0) {
          targetEmails = body.emailIds
            .slice(0, 100)
            .map((id) => localDb.getEmailById(String(id)))
            .filter((m): m is NonNullable<typeof m> => Boolean(m));
        } else {
          const lim = Math.min(Math.max(Number(body.limit || 30), 1), 100);
          targetEmails = localDb.queryEmails({ limit: lim }).items;
        }
      }
      const result = await sgwChatBridge.syncEmailsToLedger(targetEmails);
      res.json(result);
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  app.post('/api/sgw-bridge/ledger/push-weekly', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        logDate?: string;
      };
      res.json(await sgwChatBridge.pushDailyLogToWeekly(body.logDate));
    } catch (err) {
      sendSgwChatErr(res, err);
    }
  });

  const sendSedmsBridgeErr = (res: express.Response, err: unknown) => {
    if (err instanceof SedmsBridgeError) {
      res.status(err.status).json({ error: err.message });
      return;
    }
    res.status(500).json({ error: err instanceof Error ? err.message : 'SEDMS 연동 오류가 발생했습니다.' });
  };

  app.get('/api/sedms-live/status', async (_req, res) => {
    const status = sedmsLiveBridge.getStatus();
    const health = await sedmsLiveBridge.checkHealth();
    res.json({ ...status, health });
  });

  app.put('/api/sedms-live/config', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        sedmsBaseUrl?: string;
        bearerToken?: string;
        tenantId?: string;
      };
      const status = sedmsLiveBridge.saveConfig({
        sedmsBaseUrl: typeof body.sedmsBaseUrl === 'string' ? body.sedmsBaseUrl : undefined,
        bearerToken: typeof body.bearerToken === 'string' ? body.bearerToken : undefined,
        tenantId: typeof body.tenantId === 'string' ? body.tenantId : undefined,
      });
      res.json(status);
    } catch (err) {
      sendSedmsBridgeErr(res, err);
    }
  });

  app.post('/api/sedms-live/logout', (_req, res) => {
    res.json(sedmsLiveBridge.clearToken());
  });

  app.post('/api/sedms-live/scan-local-mails', (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        limit?: number;
        tenantId?: string;
        emails?: MailItem[];
      };
      let mails: MailItem[] = [];
      if (Array.isArray(body.emails) && body.emails.length > 0) {
        mails = body.emails.slice(0, 500);
      } else if (localDb.isReady()) {
        const lim = Math.min(Math.max(Number(body.limit || 150), 1), 500);
        mails = localDb.queryEmails({ limit: lim }).items;
      }
      const report = sedmsLiveBridge.scanLocalEmailsForSedms(mails, body.tenantId);
      res.json(report);
    } catch (err) {
      sendSedmsBridgeErr(res, err);
    }
  });

  app.post('/api/sedms-live/contacts/preview', async (_req, res) => {
    try {
      const contacts = localDb.isReady() ? localDb.getAllContacts() : [];
      res.json(await sedmsLiveBridge.previewContactsImport(contacts));
    } catch (err) {
      sendSedmsBridgeErr(res, err);
    }
  });

  app.post('/api/sedms-live/contacts/commit', async (req, res) => {
    try {
      const body = (req.body && typeof req.body === 'object' ? req.body : {}) as {
        confirmed?: boolean;
      };
      if (body.confirmed !== true) {
        res.status(400).json({ error: 'SEDMS 수신처 일괄 등록에는 명시적 확인(confirmed: true)이 필요합니다.' });
        return;
      }
      const contacts = localDb.isReady() ? localDb.getAllContacts() : [];
      res.json(await sedmsLiveBridge.commitContactsImport(contacts));
    } catch (err) {
      sendSedmsBridgeErr(res, err);
    }
  });

  // --- Dev-Only Source Archive Routes (Disabled in Production; Never Expose HANDOFF) ---
  app.get('/api/export/source-archive.zip', (_req, res) => {
    if (process.env.BMAIL_DEV_EXPORT !== 'true') {
      res.status(404).json({ error: 'Not found' });
      return;
    }
    try {
      const archive = buildSafeSourceZipArchive(APP_ROOT, 'bmail-outlook-source');
      const stamp = new Date().toISOString().slice(0, 10);
      res.setHeader('Content-Type', 'application/zip');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="bmail-outlook-full-source-${stamp}.zip"`
      );
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('X-Archive-File-Count', String(archive.fileCount));
      res.send(archive.buffer);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'ZIP 소스 아카이브 생성에 실패했습니다.';
      res.status(500).json({ error: msg });
    }
  });

  app.get('/api/export/source-archive', (_req, res) => {
    if (process.env.BMAIL_DEV_EXPORT !== 'true') {
      res.status(404).json({ error: 'Not found' });
      return;
    }
    try {
      const archive = buildSafeSourceArchive(APP_ROOT, 'bmail-outlook-source');
      const stamp = new Date().toISOString().slice(0, 10);
      res.setHeader('Content-Type', 'application/gzip');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="bmail-outlook-full-source-${stamp}.tar.gz"`
      );
      res.setHeader('Cache-Control', 'no-store');
      res.setHeader('X-Archive-File-Count', String(archive.fileCount));
      res.send(archive.buffer);
    } catch (err) {
      const msg = err instanceof Error ? err.message : '소스 아카이브 생성에 실패했습니다.';
      res.status(500).json({ error: msg });
    }
  });

  app.get('/api/export/setup-manual', (_req, res) => {
    res.redirect(302, '/?manual=1');
  });

  // --- Bmail SGW Groupware Gateway, SEDMS Mail Hub, Mini-Assist & AI Inbox Routes ---
  app.use(createSgwSedmsRouter(APP_ROOT));

  // Ensure unmatched API paths return clean 404 JSON instead of SPA index.html
  app.all('/api/*', (_req, res) => {
    res.status(404).json({ error: '요청한 API 엔드포인트를 찾을 수 없습니다.' });
  });

  // --- Frontend SPA Serving (Vite in dev, dist/ in prod) ---
  if (isProduction) {
    const distPath = path.resolve(APP_ROOT, 'dist');
    app.use(express.static(distPath));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  } else {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);

    app.get(['/taskpane', '/sgw/mail', '/sgw/mail/*'], async (req, res, next) => {
      try {
        const htmlPath = path.resolve(APP_ROOT, 'index.html');
        const template = fs.readFileSync(htmlPath, 'utf-8');
        const html = await vite.transformIndexHtml(req.originalUrl, template);
        res.status(200).set({ 'Content-Type': 'text/html' }).end(html);
      } catch (e) {
        next(e);
      }
    });
  }

  app.listen(PORT, BIND_HOST, () => {
    console.log(`Bmail (Outlook AI · Voice · SGW · SEDMS Workspace) running on http://${BIND_HOST}:${PORT}`);
    if (process.env.BMAIL_DISABLE_AUTO_ONBOARD !== '1') {
      backgroundMonitor.start();
    }
    if (
      process.platform === 'win32' &&
      localDb.isReady() &&
      process.env.BMAIL_DISABLE_AUTO_ONBOARD !== '1'
    ) {
      setImmediate(() => {
        try {
          const inspected = syncOutlookSettingsOnly(localDb);
          const bootCheck = localDb.getBootstrapPayload(1);
          if (
            bootCheck.totalEmails === 0 &&
            inspected.pstPaths.length > 0 &&
            migrationProgress.status === 'idle'
          ) {
            startBackgroundMigration();
          }
        } catch (err) {
          console.error('[Bmail Auto-Onboard] Outlook settings detection skipped:', err);
        }
      });
    }
  });
}

startServer();
