export type EmailFolder = 'inbox' | 'sent';

export type AttachmentKind = 'pdf' | 'excel' | 'word' | 'audio' | 'text' | 'other';

export interface MailAttachmentMeta {
  id: string;
  filename: string;
  mimeType: string;
  sizeBytes?: number;
  kind: AttachmentKind;
  extractedSummary?: string;
  isInline?: boolean;
  contentId?: string;
}

export interface MailItem {
  id: string;
  folder: EmailFolder;
  subject: string;
  fromName: string;
  fromEmail: string;
  to: string[];
  cc?: string[];
  /** Local date string in YYYY-MM-DD */
  date: string;
  /** ISO 8601 timestamp */
  timestamp: string;
  /** Short display time like "09:42" */
  timeLabel: string;
  snippet: string;
  body: string;
  bodyHtml?: string;
  hasAttachment?: boolean;
  attachments?: MailAttachmentMeta[];
  importance?: 'normal' | 'high';
  category?: string;
  folderName?: string;
  storeName?: string | null;
  conversationKey?: string;
  isRead: boolean;
  isFlagged?: boolean;
  source: 'outlook-live' | 'graph-live' | 'pst-import';
}

export type DateMode = 'single' | 'range';

export type MailScope = 'all' | 'received' | 'sent' | 'selected';

export type AiProviderId = 'groq' | 'claude' | 'openai' | 'builtin';

export type BlogTone = 'insight' | 'tech-dev' | 'essay' | 'newsletter';

export interface ExternalEnginesConfig {
  pdfEngineDir: string;
  pdfEngineCommand: string;
  excelEngineDir: string;
  excelEngineCommand: string;
  fallbackToBuiltIn: boolean;
}

export interface AppConfig {
  plugin: {
    displayName: string;
    ribbonGroupLabel: string;
    ribbonButtonLabel: string;
    supportsPinning: boolean;
    supportsSharedFolders: boolean;
    trustedCatalogPath: string;
  };
  ai: {
    defaultProvider: AiProviderId;
    fallbackToBuiltIn: boolean;
    maxEmailsPerRun: number;
    models: {
      groq: string;
      claude: string;
      openai: string;
      builtin: string;
    };
  };
  agent: {
    enabled: boolean;
    autoRunOnImport: boolean;
    autoSaveArchive: boolean;
    autoMaskSensitiveInfo: boolean;
    includeActionItems: boolean;
    defaultScope: MailScope;
    defaultDateMode: DateMode;
    defaultBlogTone: BlogTone;
    customSystemDirective: string;
  };
  pst: {
    maxExtractedMessages: number;
    maxUploadSizeMb: number;
    autoDeleteTempFiles: boolean;
  };
  externalEngines: ExternalEnginesConfig;
}

export interface ExternalEngineProbeResult {
  kind: 'pdf' | 'excel';
  configuredDir: string;
  exists: boolean;
  detectedRuntime: 'node' | 'python' | 'cli' | 'none';
  entryFile: string | null;
  message: string;
}

export interface ParsedDocumentSummary {
  id: string;
  filename: string;
  kind: AttachmentKind;
  mimeType: string;
  sizeBytes: number;
  engineUsed: string;
  pageCount?: number;
  sheetNames?: string[];
  rowCount?: number;
  wordCount?: number;
  numericHighlights?: string[];
  extractedText: string;
  warnings?: string[];
}

export interface DocumentAnalysisResult {
  id: string;
  createdAt: string;
  providerUsed: AiProviderId;
  providerModel: string;
  files: Array<{
    filename: string;
    kind: AttachmentKind;
    engineUsed: string;
    statsLabel: string;
  }>;
  linkedEmailCount: number;
  overview: string;
  keyFindings: Array<{
    sourceFile: string;
    category: string;
    summary: string;
    metrics: string[];
  }>;
  risksAndGaps: string[];
  actionItems: Array<{
    task: string;
    owner: string;
    priority: '높음' | '보통' | '참고';
    basis: string;
  }>;
  emailCrossInsights: string[];
  markdownReport: string;
}

export interface MeetingActionItem {
  owner: string;
  task: string;
  deadline: string;
  priority: '높음' | '보통' | '참고';
}

export interface MeetingMinutesResult {
  id: string;
  createdAt: string;
  providerUsed: AiProviderId;
  providerModel: string;
  transcriptionEngine?: string;
  title: string;
  meetingDate: string;
  participants: string[];
  executiveSummary: string;
  agendaTopics: Array<{
    topic: string;
    discussion: string;
    conclusion: string;
  }>;
  decisions: string[];
  actionItems: MeetingActionItem[];
  risksAndFollowups: string[];
  linkedEmailInsights: string[];
  sourceFiles: Array<{
    filename: string;
    kind: AttachmentKind;
    engineUsed: string;
  }>;
  transcript: string;
  markdownMinutes: string;
}

export interface EnvSecretStatus {
  key: string;
  label: string;
  category: 'ai' | 'microsoft' | 'security' | 'engine';
  configured: boolean;
  description: string;
}

export interface ProviderStatus {
  id: AiProviderId;
  name: string;
  vendor: string;
  model: string;
  summary: string;
  tierLabel: string;
  configured: boolean;
  envKey: string;
}

export interface AnalysisOptions {
  dateMode: DateMode;
  singleDate: string;
  startDate: string;
  endDate: string;
  scope: MailScope;
  selectedIds: string[];
  provider: AiProviderId;
  blogTone: BlogTone;
  maskSensitiveInfo: boolean;
  includeActionItems: boolean;
  customFocus: string;
}

export interface AnalyzedWorkItem {
  title: string;
  status: '완료' | '진행 중' | '후속 조치 필요';
  detail: string;
  relatedEmails: string[];
}

export interface IncomingMailSummary {
  sender: string;
  subject: string;
  intent: string;
  priority: '높음' | '보통' | '참고';
  actionNeeded: string;
}

export interface WorkLogDocument {
  title: string;
  periodLabel: string;
  executiveSummary: string;
  completedTasks: string[];
  inProgressTasks: string[];
  keyCommunications: string[];
  nextStepPlan: string[];
  issuesAndRisks: string[];
}

export interface BlogPostDocument {
  title: string;
  subtitle: string;
  readingTimeMinutes: number;
  tags: string[];
  intro: string;
  sections: {
    heading: string;
    body: string;
  }[];
  conclusion: string;
}

export interface AgentExecutionStep {
  id: string;
  title: string;
  status: 'pending' | 'running' | 'completed' | 'skipped' | 'error';
  detail: string;
  timestamp?: string;
}

export interface ProjectWorkSummary {
  projectCode: string;
  projectName: string;
  status: '완료' | '진행 중' | '후속 조치 필요';
  myActions: string[];
  keyDecisions: string[];
  attachmentsMentioned: string[];
  nextAction: string;
}

export interface InheritedOutlookProfile {
  storeName: string;
  ownerEmail: string;
  ownerName: string;
  folderStats: Array<{ name: string; count: number; isSystem: boolean }>;
  discoveredProjects: Array<{
    code: string;
    name: string;
    keywords: string[];
    color: string;
    isFavorite: boolean;
    source: 'folder' | 'subject-code';
    mailCount: number;
  }>;
  dateRange: { earliest: string; latest: string };
  receivedCount: number;
  sentCount: number;
  attachmentMailCount: number;
  calendarHintCount: number;
}

export interface AnalysisResult {
  id: string;
  createdAt: string;
  providerUsed: AiProviderId;
  providerModel: string;
  periodLabel: string;
  scopeLabel: string;
  analyzedCount: number;
  receivedCount: number;
  sentCount: number;
  overview: string;
  workPerformed: AnalyzedWorkItem[];
  projectSummaries?: ProjectWorkSummary[];
  incomingHighlights: IncomingMailSummary[];
  workLog: WorkLogDocument;
  blogPost: BlogPostDocument;
  agentSteps?: AgentExecutionStep[];
  redactedCount?: number;
  usageMetadata?: {
    promptTokenCount?: number;
    candidatesTokenCount?: number;
    totalTokenCount?: number;
    thoughtsTokenCount?: number;
  };
}

export interface SavedArchiveEntry {
  id: string;
  savedAt: string;
  result: AnalysisResult;
}

export interface OutlookHostInfo {
  isInsideOutlook: boolean;
  hostName: string;
  userEmail?: string;
  currentItemSubject?: string;
  supportsPinning?: boolean;
  pinningSupported?: boolean;
  isPinned?: boolean;
}
