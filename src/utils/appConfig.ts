import type { AiProviderId, AppConfig, BlogTone, DateMode, EnvSecretStatus, MailScope } from '../types/mail';

export const DEFAULT_APP_CONFIG: AppConfig = {
  plugin: {
    displayName: 'MailLog Studio AI',
    ribbonGroupLabel: 'MailLog AI',
    ribbonButtonLabel: '메일 분석·일지',
    supportsPinning: true,
    supportsSharedFolders: true,
    trustedCatalogPath: '\\\\fileserver\\outlook-addins\\maillog',
  },
  ai: {
    defaultProvider: 'groq',
    fallbackToBuiltIn: true,
    maxEmailsPerRun: 40,
    models: {
      groq: 'openai/gpt-oss-120b',
      claude: 'claude-haiku-5-5',
      openai: 'gpt-4o-mini',
      builtin: 'gemini-4-hydrogen',
    },
  },
  agent: {
    enabled: true,
    autoRunOnImport: false,
    autoSaveArchive: true,
    autoMaskSensitiveInfo: true,
    includeActionItems: true,
    defaultScope: 'all',
    defaultDateMode: 'single',
    defaultBlogTone: 'insight',
    customSystemDirective: '',
  },
  pst: {
    maxExtractedMessages: 500,
    maxUploadSizeMb: 150,
    autoDeleteTempFiles: true,
  },
  externalEngines: {
    pdfEngineDir: '',
    pdfEngineCommand: '',
    excelEngineDir: '',
    excelEngineCommand: '',
    fallbackToBuiltIn: true,
  },
};

const VALID_PROVIDERS: AiProviderId[] = ['groq', 'claude', 'openai', 'builtin'];
const VALID_SCOPES: MailScope[] = ['all', 'received', 'sent', 'selected'];
const VALID_DATE_MODES: DateMode[] = ['single', 'range'];
const VALID_TONES: BlogTone[] = ['insight', 'tech-dev', 'essay', 'newsletter'];

const SECRET_KEY_PATTERN = /(api[_-]?key|secret|password|passwd|token|credential|private[_-]?key)/i;

/**
 * Recursively strips any key that looks like a secret/password/API key so config.json
 * strictly stays non-secret while credentials live exclusively in .env.
 */
export function stripSecretKeysFromObject(input: unknown): unknown {
  if (Array.isArray(input)) {
    return input.map((item) => stripSecretKeysFromObject(item));
  }
  if (input && typeof input === 'object') {
    const out: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(input as Record<string, unknown>)) {
      if (SECRET_KEY_PATTERN.test(key)) {
        continue;
      }
      out[key] = stripSecretKeysFromObject(value);
    }
    return out;
  }
  return input;
}

function asString(value: unknown, fallback: string, maxLen = 240): string {
  if (typeof value !== 'string') return fallback;
  const trimmed = value.trim();
  return trimmed ? trimmed.slice(0, maxLen) : fallback;
}

function asOptionalString(value: unknown, fallback: string, maxLen = 300): string {
  if (typeof value !== 'string') return fallback;
  return value.trim().slice(0, maxLen);
}

function asBool(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback;
}

function asInt(value: unknown, fallback: number, min: number, max: number): number {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(min, Math.min(max, Math.round(n)));
}

/**
 * Validates and merges a raw JSON object into a strict, safe AppConfig.
 */
export function sanitizeAppConfig(raw: unknown): AppConfig {
  const cleaned = stripSecretKeysFromObject(raw) as Record<string, unknown> | null;
  const src = cleaned && typeof cleaned === 'object' ? cleaned : {};

  const pluginRaw = (src.plugin && typeof src.plugin === 'object' ? src.plugin : {}) as Record<string, unknown>;
  const aiRaw = (src.ai && typeof src.ai === 'object' ? src.ai : {}) as Record<string, unknown>;
  const modelsRaw = (aiRaw.models && typeof aiRaw.models === 'object' ? aiRaw.models : {}) as Record<string, unknown>;
  const agentRaw = (src.agent && typeof src.agent === 'object' ? src.agent : {}) as Record<string, unknown>;
  const pstRaw = (src.pst && typeof src.pst === 'object' ? src.pst : {}) as Record<string, unknown>;
  const extRaw = (src.externalEngines && typeof src.externalEngines === 'object'
    ? src.externalEngines
    : {}) as Record<string, unknown>;

  const defaultProvider = VALID_PROVIDERS.includes(aiRaw.defaultProvider as AiProviderId)
    ? (aiRaw.defaultProvider as AiProviderId)
    : DEFAULT_APP_CONFIG.ai.defaultProvider;

  const defaultScope = VALID_SCOPES.includes(agentRaw.defaultScope as MailScope)
    ? (agentRaw.defaultScope as MailScope)
    : DEFAULT_APP_CONFIG.agent.defaultScope;

  const defaultDateMode = VALID_DATE_MODES.includes(agentRaw.defaultDateMode as DateMode)
    ? (agentRaw.defaultDateMode as DateMode)
    : DEFAULT_APP_CONFIG.agent.defaultDateMode;

  const defaultBlogTone = VALID_TONES.includes(agentRaw.defaultBlogTone as BlogTone)
    ? (agentRaw.defaultBlogTone as BlogTone)
    : DEFAULT_APP_CONFIG.agent.defaultBlogTone;

  return {
    plugin: {
      displayName: asString(pluginRaw.displayName, DEFAULT_APP_CONFIG.plugin.displayName, 64),
      ribbonGroupLabel: asString(pluginRaw.ribbonGroupLabel, DEFAULT_APP_CONFIG.plugin.ribbonGroupLabel, 32),
      ribbonButtonLabel: asString(pluginRaw.ribbonButtonLabel, DEFAULT_APP_CONFIG.plugin.ribbonButtonLabel, 32),
      supportsPinning: asBool(pluginRaw.supportsPinning, DEFAULT_APP_CONFIG.plugin.supportsPinning),
      supportsSharedFolders: asBool(
        pluginRaw.supportsSharedFolders,
        DEFAULT_APP_CONFIG.plugin.supportsSharedFolders
      ),
      trustedCatalogPath: asString(
        pluginRaw.trustedCatalogPath,
        DEFAULT_APP_CONFIG.plugin.trustedCatalogPath,
        240
      ),
    },
    ai: {
      defaultProvider,
      fallbackToBuiltIn: asBool(aiRaw.fallbackToBuiltIn, DEFAULT_APP_CONFIG.ai.fallbackToBuiltIn),
      maxEmailsPerRun: asInt(aiRaw.maxEmailsPerRun, DEFAULT_APP_CONFIG.ai.maxEmailsPerRun, 1, 100),
      models: {
        groq: asString(modelsRaw.groq, DEFAULT_APP_CONFIG.ai.models.groq, 80),
        claude: asString(modelsRaw.claude, DEFAULT_APP_CONFIG.ai.models.claude, 80),
        openai: asString(modelsRaw.openai, DEFAULT_APP_CONFIG.ai.models.openai, 80),
        builtin: asString(modelsRaw.builtin, DEFAULT_APP_CONFIG.ai.models.builtin, 80),
      },
    },
    agent: {
      enabled: asBool(agentRaw.enabled, DEFAULT_APP_CONFIG.agent.enabled),
      autoRunOnImport: asBool(agentRaw.autoRunOnImport, DEFAULT_APP_CONFIG.agent.autoRunOnImport),
      autoSaveArchive: asBool(agentRaw.autoSaveArchive, DEFAULT_APP_CONFIG.agent.autoSaveArchive),
      autoMaskSensitiveInfo: asBool(
        agentRaw.autoMaskSensitiveInfo,
        DEFAULT_APP_CONFIG.agent.autoMaskSensitiveInfo
      ),
      includeActionItems: asBool(agentRaw.includeActionItems, DEFAULT_APP_CONFIG.agent.includeActionItems),
      defaultScope,
      defaultDateMode,
      defaultBlogTone,
      customSystemDirective:
        typeof agentRaw.customSystemDirective === 'string'
          ? agentRaw.customSystemDirective.trim().slice(0, 500)
          : DEFAULT_APP_CONFIG.agent.customSystemDirective,
    },
    pst: {
      maxExtractedMessages: asInt(
        pstRaw.maxExtractedMessages,
        DEFAULT_APP_CONFIG.pst.maxExtractedMessages,
        20,
        2000
      ),
      maxUploadSizeMb: asInt(pstRaw.maxUploadSizeMb, DEFAULT_APP_CONFIG.pst.maxUploadSizeMb, 10, 500),
      autoDeleteTempFiles: asBool(pstRaw.autoDeleteTempFiles, DEFAULT_APP_CONFIG.pst.autoDeleteTempFiles),
    },
    externalEngines: {
      pdfEngineDir: asOptionalString(extRaw.pdfEngineDir, DEFAULT_APP_CONFIG.externalEngines.pdfEngineDir, 360),
      pdfEngineCommand: asOptionalString(
        extRaw.pdfEngineCommand,
        DEFAULT_APP_CONFIG.externalEngines.pdfEngineCommand,
        360
      ),
      excelEngineDir: asOptionalString(
        extRaw.excelEngineDir,
        DEFAULT_APP_CONFIG.externalEngines.excelEngineDir,
        360
      ),
      excelEngineCommand: asOptionalString(
        extRaw.excelEngineCommand,
        DEFAULT_APP_CONFIG.externalEngines.excelEngineCommand,
        360
      ),
      fallbackToBuiltIn: asBool(
        extRaw.fallbackToBuiltIn,
        DEFAULT_APP_CONFIG.externalEngines.fallbackToBuiltIn
      ),
    },
  };
}

/**
 * Inspects environment variables and returns ONLY boolean readiness status (never raw values).
 */
export function buildEnvSecretStatuses(env: Record<string, string | undefined>): EnvSecretStatus[] {
  const isSet = (key: string) => {
    const val = env[key]?.trim();
    return Boolean(val && val !== `MY_${key}`);
  };

  return [
    {
      key: 'GEMINI_API_KEY',
      label: 'Gemini API 키 (1순위 실시간 요약·일지·음성)',
      category: 'ai',
      configured: isSet('GEMINI_API_KEY'),
      description: '13.43만 원 크레딧 우선 소진 · gemini-4-hydrogen 실시간 업무일지/요약 및 음성 전사',
    },
    {
      key: 'AGENT_BUILDER_WORKER_URL',
      label: 'Agent Builder Cloudflare Worker 프록시 URL',
      category: 'ai',
      configured: isSet('AGENT_BUILDER_WORKER_URL'),
      description: '143만 원 크레딧 우선 소진 · 과거 누적 메일 RAG 검색/맥락 요약용 Worker 엔드포인트',
    },
    {
      key: 'GCP_PROJECT_ID',
      label: 'GCP 프로젝트 ID (Discovery Engine / GCS)',
      category: 'ai',
      configured: isSet('GCP_PROJECT_ID'),
      description: 'Generative Language API 및 Discovery Engine API가 활성화된 GCP 프로젝트',
    },
    {
      key: 'AGENT_BUILDER_DATA_STORE_ID',
      label: 'Vertex AI Agent Builder 데이터 스토어 ID',
      category: 'ai',
      configured: isSet('AGENT_BUILDER_DATA_STORE_ID'),
      description: '과거 누적 메일 아카이브 버킷이 연결된 Discovery Engine 데이터 스토어',
    },
    {
      key: 'GCS_MAIL_ARCHIVE_BUCKET',
      label: 'Cloud Storage 메일 아카이브 버킷명',
      category: 'ai',
      configured: isSet('GCS_MAIL_ARCHIVE_BUCKET'),
      description: '누적 메일 JSONL(.jsonl) 업로드를 위한 Google Cloud Storage 버킷',
    },
    {
      key: 'GROQ_API_KEY',
      label: 'Groq 무료 API 키 (LLM + Whisper STT)',
      category: 'ai',
      configured: isSet('GROQ_API_KEY'),
      description: '무료 openai/gpt-oss-120b (8K TPM Map-Reduce) 및 whisper-large-v3-turbo 음성 전사용',
    },
    {
      key: 'ANTHROPIC_API_KEY',
      label: 'Claude (Anthropic) API 키',
      category: 'ai',
      configured: isSet('ANTHROPIC_API_KEY'),
      description: '정밀 한국어 업무 일지·회의록·블로그 작성용',
    },
    {
      key: 'OPENAI_API_KEY',
      label: 'ChatGPT (OpenAI) API 키 (LLM + Whisper)',
      category: 'ai',
      configured: isSet('OPENAI_API_KEY'),
      description: '긴 스레드·문서 분석 및 whisper-1 음성 전사용',
    },
    {
      key: 'MICROSOFT_CLIENT_ID',
      label: 'Microsoft Entra ID Client ID',
      category: 'microsoft',
      configured: isSet('MICROSOFT_CLIENT_ID'),
      description: 'Microsoft 365 Outlook 실시간 메일함 OAuth / MSAL NAA 앱 ID',
    },
    {
      key: 'MICROSOFT_CLIENT_SECRET',
      label: 'Microsoft Entra ID Client Secret',
      category: 'microsoft',
      configured: isSet('MICROSOFT_CLIENT_SECRET'),
      description: 'Microsoft Graph 토큰 교환용 서버 비밀 키',
    },
    {
      key: 'EXTERNAL_PDF_ENGINE_DIR',
      label: '외부 PDF 분석 프로젝트 폴더 (환경변수 오버라이드)',
      category: 'engine',
      configured: isSet('EXTERNAL_PDF_ENGINE_DIR'),
      description: '이미 보유한 외부 PDF 분석 프로젝트 폴더 경로 (미설정 시 config.json → 내장 pdf-parse 사용)',
    },
    {
      key: 'EXTERNAL_EXCEL_ENGINE_DIR',
      label: '외부 Excel 분석 프로젝트 폴더 (환경변수 오버라이드)',
      category: 'engine',
      configured: isSet('EXTERNAL_EXCEL_ENGINE_DIR'),
      description: '이미 보유한 외부 Excel 분석 프로젝트 폴더 경로 (미설정 시 config.json → 내장 exceljs 사용)',
    },
    {
      key: 'CONFIG_ADMIN_PASSWORD',
      label: '사내 config.json 변경 보호 비밀번호 (선택)',
      category: 'security',
      configured: isSet('CONFIG_ADMIN_PASSWORD'),
      description: '설정 시 config.json 저장 요청에 관리자 비밀번호 검증 적용',
    },
  ];
}

/**
 * Generates a clean .env template containing ONLY secret/password placeholders (never real values).
 */
export function buildEnvTemplateText(appUrl: string): string {
  const cleanUrl = appUrl.trim() || 'https://your-outlook-addin-host.example.com';
  return [
    '# ==============================================================================',
    '# MailLog Studio for Outlook — 비밀번호 / API 키 / 로컬 에이전트 전용 .env 템플릿',
    '# ==============================================================================',
    '# 일반 앱·플러그인·AI 에이전트·외부 폴더 경로는 config.json에서 분리 관리합니다.',
    '# 이 파일을 .env로 저장하고 로컬 에이전트가 필요한 키와 경로를 입력하세요 (Git 커밋 금지).',
    '',
    `APP_URL="${cleanUrl}"`,
    '',
    '# [AI 엔진 및 하이브리드 GCP 크레딧 우선 소진 설정]',
    '# 1순위 실시간 요약·업무일지·블로그 (13.43만 원 Gemini API 크레딧)',
    'GEMINI_API_KEY=""',
    'GEMINI_MODEL="gemini-4-hydrogen"',
    '# 1순위 과거 누적 메일 RAG 검색 (143만 원 Agent Builder 체험판 크레딧 · 2027-02-02 자동 차단)',
    'AGENT_BUILDER_WORKER_URL=""',
    'AGENT_BUILDER_WORKER_TOKEN=""',
    'GCP_PROJECT_ID=""',
    'AGENT_BUILDER_LOCATION="global"',
    'AGENT_BUILDER_DATA_STORE_ID=""',
    'AGENT_BUILDER_SERVING_CONFIG="default_search"',
    'AGENT_BUILDER_ACCESS_TOKEN=""',
    'GCS_MAIL_ARCHIVE_BUCKET=""',
    'AGENT_BUILDER_CREDIT_EXPIRY_DATE="2027-02-02"',
    'AGENT_BUILDER_HARD_STOP_AFTER_EXPIRY="true"',
    '# 2순위 폴백 AI 프로바이더',
    'GROQ_API_KEY=""',
    'ANTHROPIC_API_KEY=""',
    'OPENAI_API_KEY=""',
    '',
    '# [Microsoft 365 Outlook 실시간 연동 (Azure Entra ID)]',
    `# Azure Redirect URI: ${cleanUrl}/auth/microsoft/callback`,
    `# MSAL NAA (Outlook Desktop/OWA SSO): brk-multihub://${cleanUrl.replace(/^https?:\/\//, '')}`,
    'MICROSOFT_CLIENT_ID=""',
    'MICROSOFT_CLIENT_SECRET=""',
    'MICROSOFT_TENANT_ID="common"',
    '',
    '# [선택: 이미 보유한 외부 PDF / Excel 분석 프로젝트 폴더 직접 연동]',
    '# 폴더 경로만 지정하면 내장 파서 대신 해당 프로젝트의 PDF/Excel 분석 엔진을 우선 호출합니다.',
    'EXTERNAL_PDF_ENGINE_DIR=""',
    'EXTERNAL_PDF_ENGINE_CMD=""',
    'EXTERNAL_EXCEL_ENGINE_DIR=""',
    'EXTERNAL_EXCEL_ENGINE_CMD=""',
    '',
    '# [선택: 사내 배포 시 config.json 수정 보호 비밀번호]',
    'CONFIG_ADMIN_PASSWORD=""',
    '',
  ].join('\n');
}
