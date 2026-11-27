import React, { lazy, Suspense, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import {
  ArrowUpCircle,
  BarChart3,
  BookOpen,
  Briefcase,
  Building2,
  Calendar,
  Check,
  CheckSquare,
  Database,
  Download,
  FileSearch,
  FolderOpen,
  Loader2,
  Mail,
  MessageSquare,
  Mic,
  Plug,
  Puzzle,
  RefreshCw,
  Reply,
  Search,
  Settings,
  Shield,
  ShieldCheck,
  Smartphone,
  Sparkles,
  StickyNote,
  Users,
  X,
} from 'lucide-react';
import { GlobalSearchFilterDropdown } from './components/GlobalSearchFilterDropdown';
import type { LocalFolderRecord } from './services/localMailDb';
import type { OutlookFavoriteLink } from './services/localOutlookMigrator';
import type { UnreadPriorityBriefing } from './services/unreadBriefingEngine';
import type {
  ComposeDraftState,
  OutlookContactEntry,
  OutlookSignatureEntry,
} from './components/OutlookComposeWindow';
import { OutlookTaskpane } from './components/OutlookTaskpane';
import { OutlookWorkbench, type OutlookDraftItem } from './components/OutlookWorkbench';
import type { Pop3ConnectedAccount } from './components/Pop3QuickConnectModal';

const AiInboxIntelligenceView = lazy(() =>
  import('./components/AiInboxIntelligenceView').then((m) => ({ default: m.AiInboxIntelligenceView }))
);
const AnalysisResultPanel = lazy(() =>
  import('./components/AnalysisResultPanel').then((m) => ({ default: m.AnalysisResultPanel }))
);
const ArchiveView = lazy(() =>
  import('./components/ArchiveView').then((m) => ({ default: m.ArchiveView }))
);
const ConfigEnvManager = lazy(() =>
  import('./components/ConfigEnvManager').then((m) => ({ default: m.ConfigEnvManager }))
);
const DashboardView = lazy(() =>
  import('./components/DashboardView').then((m) => ({ default: m.DashboardView }))
);
const IntelligenceStudio = lazy(() =>
  import('./components/IntelligenceStudio').then((m) => ({ default: m.IntelligenceStudio }))
);
const InteractiveManualModal = lazy(() =>
  import('./components/InteractiveManualModal').then((m) => ({ default: m.InteractiveManualModal }))
);
const MobileMailShell = lazy(() =>
  import('./components/MobileMailShell').then((m) => ({ default: m.MobileMailShell }))
);
const OutlookCalendarView = lazy(() =>
  import('./components/OutlookCalendarView').then((m) => ({ default: m.OutlookCalendarView }))
);
const OutlookPeopleView = lazy(() =>
  import('./components/OutlookPeopleView').then((m) => ({ default: m.OutlookPeopleView }))
);
const OutlookSettingsModal = lazy(() =>
  import('./components/OutlookSettingsModal').then((m) => ({ default: m.OutlookSettingsModal }))
);
const OutlookComposeWindow = lazy(() =>
  import('./components/OutlookComposeWindow').then((m) => ({ default: m.OutlookComposeWindow }))
);
const OutlookCompanionPanel = lazy(() =>
  import('./components/OutlookCompanionPanel').then((m) => ({ default: m.OutlookCompanionPanel }))
);
const OutlookTasksView = lazy(() =>
  import('./components/OutlookTasksView').then((m) => ({ default: m.OutlookTasksView }))
);
const PluginInstallCenter = lazy(() =>
  import('./components/PluginInstallCenter').then((m) => ({ default: m.PluginInstallCenter }))
);
const Pop3QuickConnectModal = lazy(() =>
  import('./components/Pop3QuickConnectModal').then((m) => ({ default: m.Pop3QuickConnectModal }))
);
const ProjectWorkspaceView = lazy(() =>
  import('./components/ProjectWorkspaceView').then((m) => ({ default: m.ProjectWorkspaceView }))
);
const PstImportModal = lazy(() =>
  import('./components/PstImportModal').then((m) => ({ default: m.PstImportModal }))
);
const SgwChatPanel = lazy(() =>
  import('./components/SgwChatPanel').then((m) => ({ default: m.SgwChatPanel }))
);
const SgwSedmsHubView = lazy(() =>
  import('./components/SgwSedmsHubView').then((m) => ({ default: m.SgwSedmsHubView }))
);
const StickyNotesWidget = lazy(() =>
  import('./components/StickyNotesWidget').then((m) => ({ default: m.StickyNotesWidget }))
);
import type { OutlookPcProfileSnapshot } from './services/outlookLocalBridge';
import {
  applyTaskpaneAlwaysOpenBehavior,
  getSelectedOutlookRestIds,
  initializeOfficeBridge,
  isTaskpaneRoute,
  loadTaskpanePinnedPreference,
  readCurrentOutlookItem,
  saveTaskpanePinnedPreference,
} from './services/officeBridge';
import type {
  AgentExecutionStep,
  AiProviderId,
  AnalysisOptions,
  AnalysisResult,
  AppConfig,
  EmailFolder,
  EnvSecretStatus,
  InheritedOutlookProfile,
  MailItem,
  MeetingMinutesResult,
  OutlookHostInfo,
  ParsedDocumentSummary,
  ProviderStatus,
  SavedArchiveEntry,
} from './types/mail';
import { DEFAULT_APP_CONFIG } from './utils/appConfig';
import {
  type CalendarScheduleEvent,
  extractDeterministicSchedules,
  type ProjectDefinition,
} from './utils/deterministicAnalyzer';
import {
  DEFAULT_GLOBAL_SEARCH_FILTERS,
  type GlobalSearchFilters,
} from './utils/mailFilter';
import {
  ACCENT_THEME_COLORS,
  DEFAULT_FAVORITE_FOLDER_IDS,
  EMPTY_PROJECT_MAIL_OVERRIDES,
  assignEmailToProject,
  createProjectVaultFile,
  removeEmailFromProject,
  removeProjectSubfolderSafely,
  sanitizeOutlookPreferences,
  toggleFavoriteFolder,
  type OutlookPreferences,
  type OutlookTaskItem,
  type ProjectMailOverrides,
  type ProjectTreeSubfolder,
  type ProjectVaultFile,
} from './utils/outlookWorkspaceModel';
import {
  buildHeaderAccountSyncView,
  type HeaderSyncSource,
} from './utils/headerSyncState';
import { mergeInheritedProjects } from './utils/pstInheritance';

type AppPage =
  | 'workbench'
  | 'companion'
  | 'calendar'
  | 'people'
  | 'tasks'
  | 'ai-inbox'
  | 'sgw-chat'
  | 'sgw-sedms'
  | 'dashboard'
  | 'projects'
  | 'studio'
  | 'install'
  | 'config'
  | 'archive';

function isSgwEmbeddedRoute(): boolean {
  if (typeof window === 'undefined') return false;
  const params = new URLSearchParams(window.location.search);
  return (
    params.get('embed') === 'sgw' ||
    window.location.pathname === '/sgw/mail' ||
    window.location.pathname.startsWith('/sgw/mail/')
  );
}

interface ServerConfigResponse {
  appUrl: string;
  manifestXmlUrl: string;
  manifestJsonUrl: string;
  taskpaneUrl: string;
  oauthCallbackUrl: string;
  config: AppConfig;
  envSecrets: EnvSecretStatus[];
  configPasswordRequired: boolean;
  providers: ProviderStatus[];
  microsoftAuth: {
    configured: boolean;
    connected: boolean;
    userEmail: string | null;
    displayName: string | null;
  };
}

const ARCHIVE_STORAGE_KEY = 'maillog_outlook_archive_v1';
const DRAFTS_STORAGE_KEY = 'maillog_outlook_drafts_v1';
const PROJECTS_STORAGE_KEY = 'maillog_outlook_projects_v1';
const CALENDAR_STORAGE_KEY = 'maillog_outlook_calendar_v1';
const MINUTES_STORAGE_KEY = 'maillog_outlook_minutes_v1';
const OUTLOOK_PREFS_STORAGE_KEY = 'bmail_outlook_prefs_v1';
const FAVORITE_FOLDERS_STORAGE_KEY = 'bmail_outlook_favorites_v1';
const PINNED_EMAILS_STORAGE_KEY = 'bmail_outlook_pinned_v1';
const PROJECT_SUBFOLDERS_STORAGE_KEY = 'bmail_project_subfolders_v1';
const PROJECT_VAULT_FILES_STORAGE_KEY = 'bmail_project_vault_files_v1';
const PROJECT_MAIL_OVERRIDES_STORAGE_KEY = 'bmail_project_mail_overrides_v1';
const OUTLOOK_TASKS_STORAGE_KEY = 'bmail_outlook_tasks_v1';
const MAIL_CATEGORIES_MAP_STORAGE_KEY = 'bmail_mail_categories_map_v1';
const MAIL_SNOOZE_MAP_STORAGE_KEY = 'bmail_mail_snooze_map_v1';

const FALLBACK_PROVIDERS: ProviderStatus[] = [
  {
    id: 'groq',
    name: 'Groq (무료 API)',
    vendor: 'Groq Cloud',
    model: 'openai/gpt-oss-120b',
    summary: '무료 고속 추론 티어 + Map-Reduce 자동 청킹(8K TPM HTTP 413 방지)',
    tierLabel: '무료 API 키 지원',
    configured: false,
    envKey: 'GROQ_API_KEY',
  },
  {
    id: 'claude',
    name: 'Claude',
    vendor: 'Anthropic',
    model: 'claude-sonnet-4-20250514',
    summary: '자연스러운 한국어 블로그 문체와 정밀 보고서 생성',
    tierLabel: 'Anthropic API',
    configured: false,
    envKey: 'ANTHROPIC_API_KEY',
  },
  {
    id: 'openai',
    name: 'ChatGPT',
    vendor: 'OpenAI',
    model: 'gpt-4o-mini',
    summary: '긴 스레드 요약 및 액션 아이템 추출',
    tierLabel: 'OpenAI API',
    configured: false,
    envKey: 'OPENAI_API_KEY',
  },
  {
    id: 'builtin',
    name: '기본 내장 엔진',
    vendor: 'Built-in AI',
    model: 'gemini-4-hydrogen',
    summary: '별도 외부 키 등록 없이 즉시 실행 가능한 기본 분석 엔진',
    tierLabel: '즉시 실행 가능',
    configured: true,
    envKey: 'GEMINI_API_KEY',
  },
];

function loadJsonStorage<T>(key: string): T[] {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? (parsed as T[]) : [];
  } catch {
    return [];
  }
}

function loadObjectStorage<T>(key: string, fallback: T): T {
  try {
    const raw = window.localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw) as T;
  } catch {
    return fallback;
  }
}

function todayIsoDate(): string {
  return new Date().toISOString().slice(0, 10);
}

interface MailtoDraftPayload {
  to: string;
  cc: string;
  subject: string;
  body: string;
}

function parseMailtoUri(rawUri: string): MailtoDraftPayload | null {
  const trimmed = rawUri.trim();
  if (!/^mailto:/i.test(trimmed)) return null;
  const withoutScheme = trimmed.slice(7);
  const qIdx = withoutScheme.indexOf('?');
  const rawTo = qIdx === -1 ? withoutScheme : withoutScheme.slice(0, qIdx);
  const query = qIdx === -1 ? '' : withoutScheme.slice(qIdx + 1);
  const params = new URLSearchParams(query);
  const decodeSafe = (s: string) => {
    try {
      return decodeURIComponent(s);
    } catch {
      return s;
    }
  };
  return {
    to: decodeSafe(rawTo || params.get('to') || '').trim(),
    cc: (params.get('cc') || '').trim(),
    subject: (params.get('subject') || '').trim(),
    body: (params.get('body') || '').trim(),
  };
}

interface LocalBootstrapResponse {
  available: boolean;
  migratedAt: string | null;
  totalEmails: number;
  totalContacts: number;
  accounts: Array<{
    email: string;
    displayName: string;
    providerPreset: string;
    protocol: 'pop3' | 'imap';
    pop3Host: string;
    pop3Port: number;
    smtpHost: string;
    smtpPort: number;
  }>;
  signatures: Array<{
    name: string;
    htmlContent: string;
    textContent: string;
    isDefaultNew: boolean;
    isDefaultReply: boolean;
  }>;
  activeSignatureName: string | null;
  folders: LocalFolderRecord[];
  pstPaths: string[];
  contacts: OutlookPcProfileSnapshot['contacts'];
  projects: ProjectDefinition[];
  emails: MailItem[];
  kvState?: Record<string, unknown>;
}

function isManualPopupRoute(): boolean {
  if (typeof window === 'undefined') return false;
  return new URLSearchParams(window.location.search).get('manual') === '1';
}

function isComposePopupRoute(): boolean {
  if (typeof window === 'undefined') return false;
  return new URLSearchParams(window.location.search).get('compose') === '1';
}

class LazyViewErrorBoundary extends React.Component<
  { viewKey: string; onRetry: () => void; children: React.ReactNode },
  { errorMessage: string | null }
> {
  constructor(props: { viewKey: string; onRetry: () => void; children: React.ReactNode }) {
    super(props);
    this.state = { errorMessage: null };
  }

  static getDerivedStateFromError(error: Error): { errorMessage: string } {
    return { errorMessage: error?.message || '화면을 불러오는 중 오류가 발생했습니다.' };
  }

  componentDidUpdate(prevProps: { viewKey: string }) {
    if (prevProps.viewKey !== this.props.viewKey && this.state.errorMessage) {
      this.setState({ errorMessage: null });
    }
  }

  render() {
    if (this.state.errorMessage) {
      return (
        <div className="flex-1 flex flex-col items-center justify-center gap-3 p-6 bg-[#FAF9F8] text-center">
          <div className="text-sm font-semibold text-[#242424]">화면 로딩 중 일시적 오류가 발생했습니다.</div>
          <div className="text-xs text-[#616161] max-w-md">{this.state.errorMessage}</div>
          <button
            type="button"
            onClick={() => {
              this.setState({ errorMessage: null });
              this.props.onRetry();
            }}
            className="px-4 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
          >
            다시 불러오기
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function App() {
  const standaloneManualPopup = isManualPopupRoute();
  const standaloneComposePopup = isComposePopupRoute();
  const standaloneTaskpane = isTaskpaneRoute();
  const sgwEmbedded = isSgwEmbeddedRoute();
  const [page, setPage] = useState<AppPage>('companion');
  const [requestedActiveEmailId, setRequestedActiveEmailId] = useState<string | null>(null);
  const [requestedToastReplyEmailId, setRequestedToastReplyEmailId] = useState<string | null>(null);
  const [isStickyNotesOpen, setIsStickyNotesOpen] = useState(false);
  const [isManualModalOpen, setIsManualModalOpen] = useState(false);
  const [localFolders, setLocalFolders] = useState<LocalFolderRecord[]>([]);
  const [outlookFavorites, setOutlookFavorites] = useState<OutlookFavoriteLink[]>([]);
  const [isSyncingLiveOutlook, setIsSyncingLiveOutlook] = useState(false);
  const hasAutoIncrementalSyncedRef = useRef(false);
  const hasAutoSyncedPop3Ref = useRef(false);
  const seenEmailIdsRef = useRef<Set<string>>(new Set());
  const [arrivalToasts, setArrivalToasts] = useState<
    Array<{
      emailId: string;
      sender: string;
      subject: string;
      summary: string;
      urgency: 'high' | 'normal';
      summarizedBy?: string;
    }>
  >([]);

  // Strictly ZERO simulation data: starts empty until real PST/EML/Outlook/Graph emails are loaded
  const [emails, setEmails] = useState<MailItem[]>([]);
  const [deletedIds, setDeletedIds] = useState<string[]>([]);
  const [drafts, setDrafts] = useState<OutlookDraftItem[]>(() =>
    loadJsonStorage<OutlookDraftItem>(DRAFTS_STORAGE_KEY)
  );

  const [appConfig, setAppConfig] = useState<AppConfig>(DEFAULT_APP_CONFIG);
  const [envSecrets, setEnvSecrets] = useState<EnvSecretStatus[]>([]);
  const [configPasswordRequired, setConfigPasswordRequired] = useState(false);
  const [providers, setProviders] = useState<ProviderStatus[]>(FALLBACK_PROVIDERS);
  const [appUrl, setAppUrl] = useState<string>(
    typeof window !== 'undefined' ? window.location.origin : 'https://localhost:3000'
  );
  const [oauthCallbackUrl, setOauthCallbackUrl] = useState<string>(
    typeof window !== 'undefined' ? `${window.location.origin}/auth/microsoft/callback` : ''
  );
  const [microsoftAuth, setMicrosoftAuth] = useState({
    configured: false,
    connected: false,
    userEmail: null as string | null,
    displayName: null as string | null,
  });
  const [hostInfo, setHostInfo] = useState<OutlookHostInfo>({
    isInsideOutlook: false,
    hostName: standaloneTaskpane ? 'Outlook Taskpane' : 'Outlook Workbench',
  });

  const today = todayIsoDate();
  const [options, setOptions] = useState<AnalysisOptions>({
    dateMode: DEFAULT_APP_CONFIG.agent.defaultDateMode,
    singleDate: today,
    startDate: today,
    endDate: today,
    scope: DEFAULT_APP_CONFIG.agent.defaultScope,
    selectedIds: [],
    provider: 'builtin',
    blogTone: DEFAULT_APP_CONFIG.agent.defaultBlogTone,
    maskSensitiveInfo: DEFAULT_APP_CONFIG.agent.autoMaskSensitiveInfo,
    includeActionItems: DEFAULT_APP_CONFIG.agent.includeActionItems,
    customFocus: '',
  });

  const [analysisResult, setAnalysisResult] = useState<AnalysisResult | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [agentSteps, setAgentSteps] = useState<AgentExecutionStep[]>([]);
  const [analysisError, setAnalysisError] = useState<string | null>(null);
  const [archive, setArchive] = useState<SavedArchiveEntry[]>(() =>
    loadJsonStorage<SavedArchiveEntry>(ARCHIVE_STORAGE_KEY)
  );
  const [isSyncingGraph, setIsSyncingGraph] = useState(false);
  const [graphNotice, setGraphNotice] = useState<string | null>(null);
  const [taskpaneSubView, setTaskpaneSubView] = useState<'options' | 'result'>('options');
  const [isPstModalOpen, setIsPstModalOpen] = useState(false);
  const [isPop3ModalOpen, setIsPop3ModalOpen] = useState(false);
  const [pop3SessionToken, setPop3SessionToken] = useState<string | null>(null);
  const [pop3Account, setPop3Account] = useState<Pop3ConnectedAccount | null>(null);
  const [isSyncingPop3, setIsSyncingPop3] = useState(false);
  const [lastSyncedAt, setLastSyncedAt] = useState<string | null>(null);
  const [lastSyncSource, setLastSyncSource] = useState<HeaderSyncSource>(null);
  const [outlookPcProfile, setOutlookPcProfile] = useState<OutlookPcProfileSnapshot | null>(null);
  const [uiMode, setUiMode] = useState<'auto' | 'desktop' | 'mobile'>('auto');
  const [isNarrowViewport, setIsNarrowViewport] = useState<boolean>(() =>
    typeof window !== 'undefined' ? window.innerWidth < 768 : false
  );
  const [connectedStoreName, setConnectedStoreName] = useState<string | null>(null);
  const [pstFolders, setPstFolders] = useState<string[]>([]);
  const [globalSearchQuery, setGlobalSearchQuery] = useState('');
  const [globalSearchFilters, setGlobalSearchFilters] = useState<GlobalSearchFilters>(
    DEFAULT_GLOBAL_SEARCH_FILTERS
  );
  const [isAppLauncherOpen, setIsAppLauncherOpen] = useState(false);
  const [isOutlookSettingsOpen, setIsOutlookSettingsOpen] = useState(false);
  const [outlookPrefs, setOutlookPrefs] = useState<OutlookPreferences>(() =>
    sanitizeOutlookPreferences(loadObjectStorage<unknown>(OUTLOOK_PREFS_STORAGE_KEY, null))
  );
  const [outlookTasks, setOutlookTasks] = useState<OutlookTaskItem[]>(() =>
    loadJsonStorage<OutlookTaskItem>(OUTLOOK_TASKS_STORAGE_KEY)
  );
  const [mailCategoriesMap, setMailCategoriesMap] = useState<Record<string, string[]>>(() =>
    loadObjectStorage<Record<string, string[]>>(MAIL_CATEGORIES_MAP_STORAGE_KEY, {})
  );
  const [mailSnoozeMap, setMailSnoozeMap] = useState<Record<string, string>>(() =>
    loadObjectStorage<Record<string, string>>(MAIL_SNOOZE_MAP_STORAGE_KEY, {})
  );
  const [favoriteFolderIds, setFavoriteFolderIds] = useState<string[]>(() => {
    const stored = loadJsonStorage<string>(FAVORITE_FOLDERS_STORAGE_KEY);
    return stored.length > 0 ? stored : DEFAULT_FAVORITE_FOLDER_IDS;
  });
  const [pinnedEmailIds, setPinnedEmailIds] = useState<string[]>(() =>
    loadJsonStorage<string>(PINNED_EMAILS_STORAGE_KEY)
  );
  const [projectSubfolders, setProjectSubfolders] = useState<ProjectTreeSubfolder[]>(() =>
    loadJsonStorage<ProjectTreeSubfolder>(PROJECT_SUBFOLDERS_STORAGE_KEY)
  );
  const [projectVaultFiles, setProjectVaultFiles] = useState<ProjectVaultFile[]>(() =>
    loadJsonStorage<ProjectVaultFile>(PROJECT_VAULT_FILES_STORAGE_KEY)
  );
  const [projectMailOverrides, setProjectMailOverrides] = useState<ProjectMailOverrides>(() =>
    loadObjectStorage<ProjectMailOverrides>(
      PROJECT_MAIL_OVERRIDES_STORAGE_KEY,
      EMPTY_PROJECT_MAIL_OVERRIDES
    )
  );
  const [isTaskpanePinned, setIsTaskpanePinned] = useState<boolean>(() =>
    loadTaskpanePinnedPreference(DEFAULT_APP_CONFIG.plugin.supportsPinning)
  );
  // Keep right AI Taskpane collapsed by default on full desktop so Outlook's 3-pane workspace has full authentic width
  const [isTaskpaneOpen, setIsTaskpaneOpen] = useState<boolean>(() =>
    standaloneTaskpane ? loadTaskpanePinnedPreference(DEFAULT_APP_CONFIG.plugin.supportsPinning) : false
  );

  // Full Outlook Suite Shared State: Projects, Calendar Schedules, Parsed Documents, Meeting Minutes
  const [projects, setProjects] = useState<ProjectDefinition[]>(() =>
    loadJsonStorage<ProjectDefinition>(PROJECTS_STORAGE_KEY)
  );
  const [manualEvents, setManualEvents] = useState<CalendarScheduleEvent[]>(() =>
    loadJsonStorage<CalendarScheduleEvent>(CALENDAR_STORAGE_KEY)
  );
  const [completedEventIds, setCompletedEventIds] = useState<string[]>([]);
  const [parsedDocuments, setParsedDocuments] = useState<ParsedDocumentSummary[]>([]);
  const [meetingMinutesHistory, setMeetingMinutesHistory] = useState<MeetingMinutesResult[]>(() =>
    loadJsonStorage<MeetingMinutesResult>(MINUTES_STORAGE_KEY)
  );
  const [studioPreset, setStudioPreset] = useState<{
    tab: 'documents' | 'meeting' | 'search';
    title?: string;
    date?: string;
  }>({ tab: 'documents' });
  const [localDbAvailable, setLocalDbAvailable] = useState(false);
  const [localTotalEmails, setLocalTotalEmails] = useState(0);
  const [localDefaultAccount, setLocalDefaultAccount] = useState<
    LocalBootstrapResponse['accounts'][number] | null
  >(null);
  const [localSignatures, setLocalSignatures] = useState<LocalBootstrapResponse['signatures']>([]);
  const [unreadBriefing, setUnreadBriefing] = useState<UnreadPriorityBriefing | null>(null);
  const [isDiscoveringOutlookSettings, setIsDiscoveringOutlookSettings] = useState(false);
  const [requestedMailto, setRequestedMailto] = useState<MailtoDraftPayload | null>(null);
  const [activeWorkbenchFolder, setActiveWorkbenchFolder] = useState<string>('inbox');
  const [localMigrationProgress, setLocalMigrationProgress] = useState<{
    status: string;
    totalEmailsIndexed: number;
    totalFoldersIndexed: number;
    totalContactsIndexed: number;
  } | null>(null);
  const [pstAttachmentNotice, setPstAttachmentNotice] = useState<string | null>(null);
  const [isSgwChatDrawerOpen, setIsSgwChatDrawerOpen] = useState(false);
  const [sgwUnreadCount, setSgwUnreadCount] = useState(0);
  const [pendingShareMailToChat, setPendingShareMailToChat] = useState<MailItem | null>(null);
  const [sgwChatInitialPeer, setSgwChatInitialPeer] = useState<{
    email: string;
    name?: string;
  } | null>(null);
  const [desktopUpdateInfo, setDesktopUpdateInfo] = useState<{
    currentVersion: string;
    latestVersion: string;
    updateAvailable: boolean;
    releaseNotes?: string;
    downloadUrl?: string;
  } | null>(null);
  const [desktopUpdateProgress, setDesktopUpdateProgress] = useState<{
    phase: string;
    percent: number;
    message?: string;
  } | null>(null);
  const prevSgwUnreadRef = useRef(0);

  const allCalendarEvents = useMemo(() => {
    const activeMails = emails.filter((m) => !deletedIds.includes(m.id));
    const autoEvents = extractDeterministicSchedules(activeMails, projects, meetingMinutesHistory);
    const merged = [...manualEvents, ...autoEvents].map((ev) => ({
      ...ev,
      completed: ev.completed || completedEventIds.includes(ev.id),
    }));
    return merged.sort((a, b) => `${a.date} ${a.timeLabel}`.localeCompare(`${b.date} ${b.timeLabel}`));
  }, [emails, deletedIds, projects, meetingMinutesHistory, manualEvents, completedEventIds]);

  const fetchServerConfig = useCallback(async () => {
    try {
      const res = await fetch('/api/config');
      if (!res.ok) return;
      const data = (await res.json()) as ServerConfigResponse;
      setAppUrl(data.appUrl);
      setOauthCallbackUrl(data.oauthCallbackUrl);
      setAppConfig(data.config);
      setEnvSecrets(data.envSecrets);
      setConfigPasswordRequired(Boolean(data.configPasswordRequired));
      setProviders(data.providers);
      setMicrosoftAuth(data.microsoftAuth);

      const preferredFromCfg = data.config.ai.defaultProvider;
      const isPreferredReady = data.providers.some(
        (p) => p.id === preferredFromCfg && p.configured
      );
      const fallbackOrder: AiProviderId[] = ['groq', 'claude', 'openai', 'builtin'];
      const chosenProvider = isPreferredReady
        ? preferredFromCfg
        : fallbackOrder.find((id) => data.providers.some((p) => p.id === id && p.configured)) ||
          'builtin';

      setOptions((prev) => ({
        ...prev,
        provider: chosenProvider,
        scope: data.config.agent.defaultScope,
        dateMode: data.config.agent.defaultDateMode,
        blogTone: data.config.agent.defaultBlogTone,
        maskSensitiveInfo: data.config.agent.autoMaskSensitiveInfo,
        includeActionItems: data.config.agent.includeActionItems,
      }));
    } catch {
      // keep defaults if momentarily unreachable
    }
  }, []);

  useEffect(() => {
    void fetchServerConfig();
  }, [fetchServerConfig]);

  // Auto-restore encrypted persistent POP3/IMAP login session across app restarts
  useEffect(() => {
    if (standaloneTaskpane || sgwEmbedded) return;
    let cancelled = false;
    fetch('/api/pop3/session')
      .then((r) => (r.ok ? r.json() : null))
      .then(
        (
          data: {
            connected?: boolean;
            sessionToken?: string;
            account?: Pop3ConnectedAccount;
          } | null
        ) => {
          if (cancelled || !data?.connected || !data.sessionToken || !data.account) return;
          setPop3SessionToken(data.sessionToken);
          setPop3Account(data.account);
          setConnectedStoreName(`${data.account.providerName} (${data.account.email})`);
          setHostInfo((prev) => ({ ...prev, userEmail: data.account!.email }));
        }
      )
      .catch(() => {
        // Ignore offline/cloud-only mode
      });
    return () => {
      cancelled = true;
    };
  }, [standaloneTaskpane, sgwEmbedded]);

  const notifyArrivedEmails = useCallback(
    async (candidates: MailItem[]) => {
      if (!Array.isArray(candidates) || candidates.length === 0) return;
      const unseenInbox = candidates
        .filter((m) => m && m.id && m.folder === 'inbox' && !seenEmailIdsRef.current.has(m.id))
        .sort((a, b) =>
          String(b.timestamp || b.date || '').localeCompare(String(a.timestamp || a.date || ''))
        );
      for (const m of candidates) {
        if (m?.id) seenEmailIdsRef.current.add(m.id);
      }
      if (unseenInbox.length === 0 || !outlookPrefs.general.desktopNotifications) return;

      const targets = unseenInbox.slice(0, 3);
      let summaries: Array<{
        emailId: string;
        sender: string;
        subject: string;
        summary: string;
        urgency: 'high' | 'normal';
        summarizedBy?: string;
      }> = targets.map((m) => ({
        emailId: m.id,
        sender: m.fromName || m.fromEmail || '발신자',
        subject: m.subject || '(제목 없음)',
        summary: (m.snippet || m.body || m.subject || '새 메일이 접수되었습니다.').slice(0, 120),
        urgency: m.importance === 'high' ? 'high' : 'normal',
        summarizedBy: 'deterministic',
      }));

      try {
        const res = await fetch('/api/local/mail/quick-summary', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ emails: targets, enableAi: true }),
        });
        if (res.ok) {
          const data = (await res.json()) as {
            summaries?: typeof summaries;
          };
          if (Array.isArray(data.summaries) && data.summaries.length > 0) {
            summaries = data.summaries;
          }
        }
      } catch {
        // Keep fast deterministic summary above
      }

      setArrivalToasts((prev) => {
        const existingIds = new Set(summaries.map((s) => s.emailId));
        return [...summaries, ...prev.filter((p) => !existingIds.has(p.emailId))].slice(0, 3);
      });
      const dismissIds = new Set(summaries.map((s) => s.emailId));
      window.setTimeout(() => {
        setArrivalToasts((prev) => prev.filter((item) => !dismissIds.has(item.emailId)));
      }, 12_000);

      const top = summaries[0];
      const desktopBridge = (
        window as unknown as {
          bmailDesktop?: {
            showMailToast?: (payload: {
              emailId: string;
              sender: string;
              subject: string;
              snippet: string;
              summary?: string;
              urgency?: 'high' | 'normal';
            }) => Promise<unknown>;
          };
        }
      ).bmailDesktop;
      void desktopBridge?.showMailToast?.({
        emailId: top.emailId,
        sender: top.sender,
        subject: top.subject,
        snippet: top.summary,
        summary: top.summary,
        urgency: top.urgency,
      });
    },
    [outlookPrefs.general.desktopNotifications]
  );

  // Zero-Click Local SQLite Outlook Archive Hydration (real migrated emails/accounts/contacts/signatures)
  const hydrateLocalBootstrap = useCallback(async () => {
    if (standaloneTaskpane || sgwEmbedded) return;
    try {
      const res = await fetch('/api/local/bootstrap?limit=100');
      if (!res.ok) return;
      const data = (await res.json()) as LocalBootstrapResponse;
      if (!data || !data.available) return;

      setLocalDbAvailable(true);
      setLocalTotalEmails(Number(data.totalEmails || 0));

      if (Array.isArray(data.emails) && data.emails.length > 0) {
        for (const m of data.emails) {
          if (m?.id) seenEmailIdsRef.current.add(m.id);
        }
        setEmails((prev) => {
          if (prev.length === 0) return data.emails;
          const seen = new Set(prev.map((m) => m.id));
          return [...prev, ...data.emails.filter((m) => !seen.has(m.id))];
        });
      }

      if (Array.isArray(data.folders) && data.folders.length > 0) {
        setLocalFolders(data.folders);
        setPstFolders(Array.from(new Set(data.folders.map((f) => f.folderPath))));
      }

      if (data.kvState && Array.isArray(data.kvState.outlook_favorites)) {
        setOutlookFavorites(data.kvState.outlook_favorites as OutlookFavoriteLink[]);
      }

      if (data.kvState && Array.isArray(data.kvState.outlook_tasks)) {
        setOutlookTasks(data.kvState.outlook_tasks as OutlookTaskItem[]);
      }
      if (data.kvState && data.kvState.mail_categories_map && typeof data.kvState.mail_categories_map === 'object') {
        setMailCategoriesMap(data.kvState.mail_categories_map as Record<string, string[]>);
      }
      if (data.kvState && data.kvState.mail_snooze_map && typeof data.kvState.mail_snooze_map === 'object') {
        setMailSnoozeMap(data.kvState.mail_snooze_map as Record<string, string>);
      }

      let hasLocalPrefs = false;
      try {
        hasLocalPrefs = Boolean(window.localStorage.getItem(OUTLOOK_PREFS_STORAGE_KEY));
      } catch {
        hasLocalPrefs = false;
      }
      if (data.kvState?.outlookPrefs && typeof data.kvState.outlookPrefs === 'object') {
        setOutlookPrefs(sanitizeOutlookPreferences(data.kvState.outlookPrefs));
      } else if (
        !hasLocalPrefs &&
        data.kvState?.outlook_window_settings &&
        typeof data.kvState.outlook_window_settings === 'object'
      ) {
        const ws = data.kvState.outlook_window_settings as {
          enableSingleLineRibbon?: boolean;
          minimalHeaderOn?: boolean;
          calWorkStart?: string;
          calWorkEnd?: string;
        };
        setOutlookPrefs((prev) =>
          sanitizeOutlookPreferences({
            ...prev,
            mail: {
              ...prev.mail,
              ...(typeof ws.enableSingleLineRibbon === 'boolean'
                ? { singleLineRibbon: ws.enableSingleLineRibbon }
                : {}),
              ...(typeof ws.minimalHeaderOn === 'boolean'
                ? { minimalHeader: ws.minimalHeaderOn }
                : {}),
            },
            calendar: {
              ...prev.calendar,
              ...(ws.calWorkStart ? { workStart: ws.calWorkStart } : {}),
              ...(ws.calWorkEnd ? { workEnd: ws.calWorkEnd } : {}),
            },
          })
        );
      }

      const primaryAcc = Array.isArray(data.accounts) ? data.accounts[0] : null;
      if (primaryAcc) {
        setLocalDefaultAccount(primaryAcc);
        setConnectedStoreName(`${primaryAcc.displayName || 'Outlook'} · ${primaryAcc.email}`);
        setHostInfo((prev) => ({ ...prev, userEmail: primaryAcc.email }));
      }

      if (
        (Array.isArray(data.accounts) && data.accounts.length > 0) ||
        (Array.isArray(data.contacts) && data.contacts.length > 0)
      ) {
        setOutlookPcProfile({
          capturedAt: data.migratedAt || new Date().toISOString(),
          outlookVersion: '16.0',
          accounts: (data.accounts || []).map((a) => ({
            displayName: a.displayName,
            smtpAddress: a.email,
            accountType: a.protocol === 'imap' ? 'IMAP' : 'POP3',
            pop3Host: a.pop3Host,
            pop3Port: a.pop3Port,
            smtpHost: a.smtpHost,
            smtpPort: a.smtpPort,
          })),
          pstPaths: data.pstPaths || [],
          contacts: data.contacts || [],
          folders: (data.folders || []).map((f) => f.folderPath),
          signatureNames: (data.signatures || []).map((s) => s.name),
        });
        if (Array.isArray(data.contacts) && data.contacts.length > 0) {
          try {
            window.localStorage.setItem(
              'bmail_local_contacts_cache',
              JSON.stringify(data.contacts.slice(0, 500))
            );
          } catch {
            // ignore storage quota
          }
        }
      }

      if (Array.isArray(data.projects) && data.projects.length > 0) {
        setProjects((prev) => {
          const seenCodes = new Set(prev.map((p) => p.code.trim().toLowerCase()));
          const additions = data.projects.filter(
            (p) => p.code && !seenCodes.has(p.code.trim().toLowerCase())
          );
          if (additions.length === 0) return prev;
          const merged = [...additions, ...prev];
          try {
            window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(merged));
          } catch {
            // ignore
          }
          return merged;
        });
      }

      if (Array.isArray(data.signatures) && data.signatures.length > 0) {
        setLocalSignatures(data.signatures);
        try {
          window.localStorage.setItem('bmail_local_signatures_cache', JSON.stringify(data.signatures));
        } catch {
          // ignore
        }
        const activeSig =
          data.signatures.find((s) => s.isDefaultNew) || data.signatures[0];
        if (activeSig && (activeSig.htmlContent.trim() || activeSig.textContent.trim())) {
          setOutlookPrefs((prev) => {
            const hasExistingHtml = Boolean(prev.mail.signatureHtml?.trim());
            const hasNewHtml = Boolean(activeSig.htmlContent?.trim());
            if (prev.mail.signature.trim() && (!hasNewHtml || hasExistingHtml)) {
              return prev;
            }
            const next = sanitizeOutlookPreferences({
              ...prev,
              mail: {
                ...prev.mail,
                signature: prev.mail.signature.trim() || activeSig.textContent.trim(),
                signatureHtml: activeSig.htmlContent || prev.mail.signatureHtml || '',
                signatureName: activeSig.name || prev.mail.signatureName || '',
                autoSignatureNew: true,
                autoSignatureReply: true,
              },
            });
            try {
              window.localStorage.setItem(OUTLOOK_PREFS_STORAGE_KEY, JSON.stringify(next));
            } catch {
              // ignore
            }
            return next;
          });
        }
      }

      if (data.totalEmails > 0) {
        setGraphNotice(
          `로컬 Outlook 완전 이관 완료 · 메일 ${data.totalEmails.toLocaleString()}건 · 폴더 ${(data.folders || []).length}개 · 연락처 ${(data.totalContacts || 0).toLocaleString()}명 · 서명(${data.activeSignatureName || '기본'})`
        );
      }
    } catch {
      // Offline or edge deployment without local DB
    }
  }, [standaloneTaskpane, sgwEmbedded]);

  useEffect(() => {
    void hydrateLocalBootstrap();
  }, [hydrateLocalBootstrap]);

  useEffect(() => {
    if (!localDbAvailable || standaloneTaskpane || sgwEmbedded) return;
    let cancelled = false;
    fetch('/api/local/unread-briefing?limit=250')
      .then((r) => (r.ok ? r.json() : null))
      .then((data: UnreadPriorityBriefing | null) => {
        if (!cancelled && data && typeof data.totalUnread === 'number') {
          setUnreadBriefing(data);
        }
      })
      .catch(() => {});
    return () => {
      cancelled = true;
    };
  }, [localDbAvailable, standaloneTaskpane, sgwEmbedded]);

  const handleDiscoverOutlookSettings = useCallback(async (): Promise<string> => {
    setIsDiscoveringOutlookSettings(true);
    try {
      const res = await fetch('/api/local/outlook/discover-settings', { method: 'POST' });
      const data = (await res.json()) as {
        ok?: boolean;
        error?: string;
        signaturesCount?: number;
        inlinedSignatureImages?: number;
        defaultNewSignature?: string | null;
        accountsCount?: number;
        favoritesCount?: number;
        signatures?: LocalBootstrapResponse['signatures'];
        outlookFavorites?: OutlookFavoriteLink[];
        windowSettings?: {
          enableSingleLineRibbon?: boolean;
          minimalHeaderOn?: boolean;
          calWorkStart?: string;
          calWorkEnd?: string;
          readingPanePosition?: 'right' | 'bottom' | 'hidden';
          defaultImportance?: 'low' | 'normal' | 'high';
          requestReadReceipt?: boolean;
        };
      };
      if (!res.ok) {
        const errMsg = data.error || 'Outlook 현재 설정 탐색에 실패했습니다.';
        setGraphNotice(errMsg);
        return errMsg;
      }

      if (Array.isArray(data.signatures) && data.signatures.length > 0) {
        setLocalSignatures(data.signatures);
        try {
          window.localStorage.setItem('bmail_local_signatures_cache', JSON.stringify(data.signatures));
        } catch {
          // ignore
        }
      }
      if (Array.isArray(data.outlookFavorites) && data.outlookFavorites.length > 0) {
        setOutlookFavorites(data.outlookFavorites);
      }

      const activeSig =
        data.signatures?.find((s) => s.isDefaultNew) || data.signatures?.[0] || null;
      const ws = data.windowSettings || {};
      setOutlookPrefs((prev) => {
        const next = sanitizeOutlookPreferences({
          ...prev,
          mail: {
            ...prev.mail,
            ...(activeSig
              ? {
                  signature: activeSig.textContent.trim() || prev.mail.signature,
                  signatureHtml: activeSig.htmlContent || prev.mail.signatureHtml,
                  signatureName: activeSig.name,
                  autoSignatureNew: true,
                  autoSignatureReply: true,
                }
              : {}),
            ...(typeof ws.enableSingleLineRibbon === 'boolean'
              ? { singleLineRibbon: ws.enableSingleLineRibbon }
              : {}),
            ...(typeof ws.minimalHeaderOn === 'boolean'
              ? { minimalHeader: ws.minimalHeaderOn }
              : {}),
            ...(ws.readingPanePosition ? { readingPanePosition: ws.readingPanePosition } : {}),
            ...(ws.defaultImportance ? { defaultImportance: ws.defaultImportance } : {}),
            ...(typeof ws.requestReadReceipt === 'boolean'
              ? { requestReadReceipt: ws.requestReadReceipt }
              : {}),
          },
          calendar: {
            ...prev.calendar,
            ...(ws.calWorkStart ? { workStart: ws.calWorkStart } : {}),
            ...(ws.calWorkEnd ? { workEnd: ws.calWorkEnd } : {}),
          },
        });
        try {
          window.localStorage.setItem(OUTLOOK_PREFS_STORAGE_KEY, JSON.stringify(next));
        } catch {
          // ignore
        }
        void fetch('/api/local/state/outlookPrefs', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ value: next }),
        }).catch(() => {});
        return next;
      });

      const msg = `PC Outlook 설정 적용 완료 · 서명 ${data.signaturesCount ?? 0}개(명함 이미지 ${data.inlinedSignatureImages ?? 0}개) · 계정 ${data.accountsCount ?? 0}개 · 즐겨찾기 ${data.favoritesCount ?? 0}개`;
      setGraphNotice(msg);
      return msg;
    } catch {
      const msg = 'Outlook 현재 설정 탐색 중 오류가 발생했습니다.';
      setGraphNotice(msg);
      return msg;
    } finally {
      setIsDiscoveringOutlookSettings(false);
    }
  }, []);

  const handleSyncIncrementalOutlook = useCallback(
    async (silent = false, opts?: { sendAndReceive?: boolean }) => {
      if (standaloneTaskpane || sgwEmbedded) return;
      setIsSyncingLiveOutlook(true);
      try {
        const res = await fetch('/api/local/outlook/sync-incremental', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sendAndReceive: Boolean(opts?.sendAndReceive) }),
        });
        if (res.ok) {
          const summary = (await res.json()) as {
            outlookRunning?: boolean;
            sendReceiveTriggered?: boolean;
            addedPstEmails?: number;
            addedComEmails?: number;
            skippedUnchangedPsts?: number;
            favoritesCount?: number;
            totalEmails?: number;
            outlookFavorites?: OutlookFavoriteLink[];
            newEmails?: MailItem[];
          };
          if (Array.isArray(summary.outlookFavorites) && summary.outlookFavorites.length > 0) {
            setOutlookFavorites(summary.outlookFavorites);
          }
          const added = Number(summary.addedPstEmails || 0) + Number(summary.addedComEmails || 0);
          if (added > 0 || !silent) {
            await hydrateLocalBootstrap();
          }
          if (Array.isArray(summary.newEmails) && summary.newEmails.length > 0) {
            setEmails((prev) => {
              const existing = new Set(prev.map((m) => m.id));
              const additions = summary.newEmails!.filter((m) => !existing.has(m.id));
              return additions.length > 0 ? [...additions, ...prev] : prev;
            });
            void notifyArrivedEmails(summary.newEmails);
          }
          if (!silent || added > 0) {
            const syncTitle = summary.sendReceiveTriggered
              ? 'Outlook 보내기/받기 · 증분 동기화 완료'
              : 'Outlook 증분 동기화 완료';
            setGraphNotice(
              `${syncTitle} · 신규 메일 +${added}건 반영 (기존 PST ${
                summary.skippedUnchangedPsts || 0
              }개 유지 · 즐겨찾기 ${summary.favoritesCount || 0}개 · 총 ${(
                summary.totalEmails || localTotalEmails
              ).toLocaleString()}건)`
            );
          }
        }
      } catch {
        // Ignore silent background sync error
      } finally {
        setIsSyncingLiveOutlook(false);
      }
    },
    [standaloneTaskpane, sgwEmbedded, hydrateLocalBootstrap, localTotalEmails, notifyArrivedEmails]
  );

  // Immediate startup Send/Receive + incremental Outlook sync so new emails arrive right on launch
  useEffect(() => {
    if (!localDbAvailable || hasAutoIncrementalSyncedRef.current || standaloneTaskpane || sgwEmbedded) {
      return;
    }
    hasAutoIncrementalSyncedRef.current = true;
    const timer = window.setTimeout(() => {
      void handleSyncIncrementalOutlook(true, { sendAndReceive: true });
    }, 300);
    return () => window.clearTimeout(timer);
  }, [localDbAvailable, standaloneTaskpane, sgwEmbedded, handleSyncIncrementalOutlook]);

  useEffect(() => {
    if (!localDbAvailable || standaloneTaskpane || sgwEmbedded) return;
    let lastSyncMs = Date.now();
    const intervalId = window.setInterval(() => {
      if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return;
      lastSyncMs = Date.now();
      void handleSyncIncrementalOutlook(true);
    }, 180_000);
    const onVisibilityChange = () => {
      if (document.visibilityState === 'visible' && Date.now() - lastSyncMs >= 120_000) {
        lastSyncMs = Date.now();
        void handleSyncIncrementalOutlook(true);
      }
    };
    document.addEventListener('visibilitychange', onVisibilityChange);
    return () => {
      window.clearInterval(intervalId);
      document.removeEventListener('visibilitychange', onVisibilityChange);
    };
  }, [localDbAvailable, standaloneTaskpane, sgwEmbedded, handleSyncIncrementalOutlook]);

  useEffect(() => {
    if (standaloneTaskpane || sgwEmbedded) return;
    let cancelled = false;
    let timerId: number | null = null;

    const pollStatus = async () => {
      try {
        const res = await fetch('/api/local/outlook/migrate/status');
        if (!res.ok || cancelled) return;
        const data = (await res.json()) as {
          progress?: {
            status?: string;
            totalEmailsIndexed?: number;
            totalFoldersIndexed?: number;
            totalContactsIndexed?: number;
          };
        };
        const prog = data.progress;
        if (!prog || cancelled) return;
        const nextState = {
          status: prog.status || 'idle',
          totalEmailsIndexed: Number(prog.totalEmailsIndexed || 0),
          totalFoldersIndexed: Number(prog.totalFoldersIndexed || 0),
          totalContactsIndexed: Number(prog.totalContactsIndexed || 0),
        };
        setLocalMigrationProgress(nextState);
        if (nextState.status === 'running') {
          void hydrateLocalBootstrap();
          timerId = window.setTimeout(() => void pollStatus(), 4000);
        } else if (nextState.status === 'completed') {
          void hydrateLocalBootstrap();
        }
      } catch {
        // ignore transient startup poll errors
      }
    };

    void pollStatus();
    return () => {
      cancelled = true;
      if (timerId !== null) window.clearTimeout(timerId);
    };
  }, [hydrateLocalBootstrap, standaloneTaskpane, sgwEmbedded]);

  const handleAnalyzePstAttachment = useCallback(async (emailId: string, idx: number) => {
    setPstAttachmentNotice(null);
    try {
      const res = await fetch('/api/local/pst/attachment/parse', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emailId, idx }),
      });
      const data = (await res.json()) as {
        document?: ParsedDocumentSummary;
        error?: string;
      };
      if (!res.ok || !data.document) {
        setPstAttachmentNotice(data.error || 'PST 첨부파일 추출에 실패했습니다.');
        return;
      }
      const doc = data.document;
      setParsedDocuments((prev) => [doc, ...prev.filter((d) => d.id !== doc.id)]);
      setStudioPreset({ tab: 'documents' });
      setPage('studio');
    } catch (err) {
      setPstAttachmentNotice(
        err instanceof Error ? err.message : 'PST 첨부파일 추출 중 오류가 발생했습니다.'
      );
    }
  }, []);

  const handleOpenMailById = useCallback(
    async (emailId: string) => {
      const cleanId = emailId.trim();
      if (!cleanId) return;
      if (!emails.some((m) => m.id === cleanId)) {
        try {
          const res = await fetch(`/api/local/emails/${encodeURIComponent(cleanId)}`);
          if (res.ok) {
            const data = (await res.json()) as { email?: MailItem };
            if (data.email) {
              setEmails((prev) =>
                prev.some((m) => m.id === data.email!.id) ? prev : [data.email!, ...prev]
              );
            }
          }
        } catch {
          // ignore if absent locally
        }
      }
      setPage('workbench');
      setRequestedActiveEmailId(cleanId);
    },
    [emails]
  );

  const handleSgwUnreadChange = useCallback((nextCount: number) => {
    setSgwUnreadCount(nextCount);
    if (nextCount > prevSgwUnreadRef.current && typeof window !== 'undefined') {
      const notifier = (
        window as unknown as {
          bmailDesktop?: {
            notify?: (payload: { title: string; body: string; route?: string }) => void;
          };
        }
      ).bmailDesktop?.notify;
      notifier?.({
        title: 'SGW 사내 채팅 새 메시지',
        body: `읽지 않은 사내 채팅·쪽지 ${nextCount}건이 있습니다.`,
        route: 'sgw-chat',
      });
    }
    prevSgwUnreadRef.current = nextCount;
  }, []);

  // Poll SGW unread count quietly when Chat Drawer / Full Chat is closed
  useEffect(() => {
    if (standaloneTaskpane || sgwEmbedded || isSgwChatDrawerOpen || page === 'sgw-chat') return;
    let cancelled = false;
    const pollUnread = async () => {
      try {
        const stRes = await fetch('/api/sgw-chat/status');
        if (!stRes.ok || cancelled) return;
        const st = (await stRes.json()) as { connected?: boolean; user?: { sub: string } };
        if (!st.connected || cancelled) return;
        const [rRes, dRes] = await Promise.all([
          fetch('/api/sgw-chat/rooms'),
          fetch('/api/sgw-chat/dm'),
        ]);
        if (cancelled) return;
        let total = 0;
        if (rRes.ok) {
          const rData = (await rRes.json()) as { unreadTotal?: number };
          total += Number(rData.unreadTotal || 0);
        }
        if (dRes.ok && st.user?.sub) {
          const dData = (await dRes.json()) as {
            messages?: Array<{ receiverId: string; isRead: boolean; isDeleted?: boolean }>;
          };
          total += (dData.messages || []).filter(
            (m) => m.receiverId === st.user!.sub && !m.isRead && !m.isDeleted
          ).length;
        }
        handleSgwUnreadChange(total);
      } catch {
        // ignore offline background poll
      }
    };
    void pollUnread();
    const id = window.setInterval(() => {
      if (document.visibilityState === 'visible') void pollUnread();
    }, 20000);
    return () => {
      cancelled = true;
      window.clearInterval(id);
    };
  }, [standaloneTaskpane, sgwEmbedded, isSgwChatDrawerOpen, page, handleSgwUnreadChange]);

  // Handle ?mailto=... query param and Electron desktop IPC events (mailto, bmail:// deep links, auto-updates)
  useEffect(() => {
    if (typeof window === 'undefined') return;
    const params = new URLSearchParams(window.location.search);
    const mailtoParam = params.get('mailto');
    if (mailtoParam) {
      const parsed = parseMailtoUri(mailtoParam);
      if (parsed) {
        setRequestedMailto(parsed);
        setPage('workbench');
      }
    }
    const bridge = (
      window as unknown as {
        bmailDesktop?: {
          onMailto?: (cb: (uri: string) => void) => () => void;
          onDeepLink?: (cb: (uri: string) => void) => () => void;
          onOpenRoute?: (cb: (route: string) => void) => () => void;
          onUpdateAvailable?: (
            cb: (info: {
              currentVersion: string;
              latestVersion: string;
              updateAvailable: boolean;
              releaseNotes?: string;
              downloadUrl?: string;
            }) => void
          ) => () => void;
          onUpdateProgress?: (
            cb: (prog: { phase: string; percent: number; message?: string }) => void
          ) => () => void;
          onToastAction?: (
            cb: (payload: { action: 'markRead' | 'reply' | 'open'; emailId: string }) => void
          ) => () => void;
        };
      }
    ).bmailDesktop;
    if (!bridge) return undefined;

    const cleanups: Array<() => void> = [];
    if (bridge.onMailto) {
      cleanups.push(
        bridge.onMailto((uri) => {
          const parsed = parseMailtoUri(uri);
          if (parsed) {
            setRequestedMailto(parsed);
            setPage('workbench');
          }
        })
      );
    }
    if (bridge.onDeepLink) {
      cleanups.push(
        bridge.onDeepLink((uri) => {
          try {
            const u = new URL(uri);
            if (u.protocol === 'bmail:' || u.protocol === 'sgw-outlook:') {
              const target = (u.hostname || u.pathname.replace(/^\/+/, '')).toLowerCase();
              if (target === 'open') {
                const emailId = u.searchParams.get('id');
                if (emailId) void handleOpenMailById(emailId);
                else setPage('workbench');
              } else if (target === 'chat') {
                setIsSgwChatDrawerOpen(true);
              }
            }
          } catch {
            // ignore invalid deep-link URI
          }
        })
      );
    }
    if (bridge.onOpenRoute) {
      cleanups.push(
        bridge.onOpenRoute((route) => {
          if (route === 'sgw-chat') {
            setPage('sgw-chat');
          }
        })
      );
    }
    if (bridge.onUpdateAvailable) {
      cleanups.push(
        bridge.onUpdateAvailable((info) => {
          if (info?.updateAvailable) setDesktopUpdateInfo(info);
        })
      );
    }
    if (bridge.onUpdateProgress) {
      cleanups.push(
        bridge.onUpdateProgress((prog) => {
          setDesktopUpdateProgress(prog);
        })
      );
    }
    if (bridge.onToastAction) {
      cleanups.push(
        bridge.onToastAction((payload) => {
          if (!payload || !payload.emailId) return;
          if (payload.action === 'markRead') {
            setEmails((prev) =>
              prev.map((m) => (m.id === payload.emailId ? { ...m, isRead: true } : m))
            );
            fetch(`/api/local/emails/${encodeURIComponent(payload.emailId)}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ isRead: true }),
            }).catch(() => {});
          } else if (payload.action === 'reply') {
            setPage('workbench');
            setRequestedToastReplyEmailId(payload.emailId);
          } else if (payload.action === 'open') {
            void handleOpenMailById(payload.emailId);
          }
        })
      );
    }
    return () => {
      for (const fn of cleanups) fn();
    };
  }, [handleOpenMailById]);

  // Query local SQLite FTS5 store on search or folder switch when archive exceeds initial memory slice
  useEffect(() => {
    if (!localDbAvailable || localTotalEmails === 0) return;
    const trimmedQ = globalSearchQuery.trim();
    const isDirectFolder =
      activeWorkbenchFolder !== 'all' &&
      activeWorkbenchFolder !== 'drafts' &&
      activeWorkbenchFolder !== 'flagged' &&
      !activeWorkbenchFolder.startsWith('project:');
    if (!trimmedQ && !isDirectFolder) return;

    const delay = trimmedQ ? 250 : 0;
    const timer = window.setTimeout(() => {
      const qs = new URLSearchParams();
      if (trimmedQ) qs.set('q', trimmedQ);
      if (activeWorkbenchFolder === 'trash') {
        qs.set('deletedOnly', 'true');
      } else if (isDirectFolder) {
        qs.set('folder', activeWorkbenchFolder);
      }
      qs.set('limit', '300');

      void fetch(`/api/local/emails?${qs.toString()}`)
        .then((res) => (res.ok ? (res.json() as Promise<{ items?: MailItem[] }>) : null))
        .then((data) => {
          const items = data?.items;
          if (!Array.isArray(items) || items.length === 0) return;
          setEmails((prev) => {
            const byId = new Map(prev.map((m) => [m.id, m]));
            for (const item of items) {
              if (!byId.has(item.id)) {
                byId.set(item.id, item);
              }
            }
            return Array.from(byId.values()).sort((a, b) =>
              b.timestamp.localeCompare(a.timestamp)
            );
          });
          if (activeWorkbenchFolder === 'trash') {
            setDeletedIds((prev) =>
              Array.from(new Set([...prev, ...items.map((m) => m.id)]))
            );
          }
        })
        .catch(() => {
          // ignore transient local query failure
        });
    }, delay);

    return () => window.clearTimeout(timer);
  }, [localDbAvailable, localTotalEmails, globalSearchQuery, activeWorkbenchFolder]);

  useEffect(() => {
    if (typeof window === 'undefined' || !window.matchMedia) return;
    const mq = window.matchMedia('(max-width: 767px)');
    const syncViewport = () => setIsNarrowViewport(mq.matches);
    syncViewport();
    mq.addEventListener('change', syncViewport);
    return () => mq.removeEventListener('change', syncViewport);
  }, []);

  // Keep Taskpane pinned state synced with manifest SupportsPinning & persisted as user default
  useEffect(() => {
    const supportsPinning = appConfig.plugin.supportsPinning;
    const effectivePinned = supportsPinning ? isTaskpanePinned : false;

    void saveTaskpanePinnedPreference(effectivePinned);
    void applyTaskpaneAlwaysOpenBehavior(effectivePinned);

    if (standaloneTaskpane && effectivePinned) {
      setIsTaskpaneOpen(true);
    }

    setHostInfo((prev) => ({
      ...prev,
      supportsPinning,
      isPinned: effectivePinned,
    }));
  }, [standaloneTaskpane, appConfig.plugin.supportsPinning, isTaskpanePinned]);

  // Initialize Microsoft Office.js bridge with SupportsPinning always-open persistence & event cleanup
  useEffect(() => {
    if (!standaloneTaskpane) return;
    let active = true;
    let disposeBridge: (() => void) | null = null;
    const supportsPinning = appConfig.plugin.supportsPinning;

    void initializeOfficeBridge({
      supportsPinning,
      defaultPinned: isTaskpanePinned,
      onItemChanged: (liveMail) => {
        if (!active) return;
        setEmails((prev) => {
          const exists = prev.some((m) => m.id === liveMail.id);
          return exists
            ? prev.map((m) => (m.id === liveMail.id ? liveMail : m))
            : [liveMail, ...prev];
        });
        setOptions((prev) => ({
          ...prev,
          singleDate: liveMail.date,
          selectedIds: prev.selectedIds.includes(liveMail.id)
            ? prev.selectedIds
            : [liveMail.id, ...prev.selectedIds],
        }));
        setHostInfo((prev) => ({
          ...prev,
          currentItemSubject: liveMail.subject,
        }));
      },
      onSelectedItemsChanged: (restIds) => {
        if (!active || restIds.length === 0) return;
        setOptions((prev) => ({
          ...prev,
          scope: 'selected',
        }));
      },
    }).then(({ hostInfo: nextInfo, cleanup }) => {
      if (!active) {
        cleanup();
        return;
      }
      disposeBridge = cleanup;
      setHostInfo(nextInfo);
      if (typeof nextInfo.isPinned === 'boolean') {
        setIsTaskpanePinned(nextInfo.isPinned);
      }
    });

    return () => {
      active = false;
      disposeBridge?.();
    };
  }, [standaloneTaskpane, appConfig.plugin.supportsPinning]);

  // Listen for Microsoft OAuth popup callback
  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS' && event.data?.provider === 'microsoft') {
        void fetchServerConfig().then(() => {
          setGraphNotice('Microsoft 365 Outlook 계정이 연결되었습니다. 상단 동기화 버튼으로 실제 메일을 불러오세요.');
        });
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, [fetchServerConfig]);

  const handleChangeOptions = useCallback((patch: Partial<AnalysisOptions>) => {
    setAnalysisError(null);
    setOptions((prev) => ({ ...prev, ...patch }));
  }, []);

  const handleToggleSelectEmail = useCallback((id: string) => {
    setAnalysisError(null);
    setOptions((prev) => {
      const exists = prev.selectedIds.includes(id);
      return {
        ...prev,
        selectedIds: exists
          ? prev.selectedIds.filter((item) => item !== id)
          : [...prev.selectedIds, id],
      };
    });
  }, []);

  const handleToggleReadEmail = useCallback(
    (id: string) => {
      setEmails((prev) =>
        prev.map((m) => {
          if (m.id !== id) return m;
          const nextRead = !m.isRead;
          if (localDbAvailable) {
            void fetch(`/api/local/emails/${encodeURIComponent(id)}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ isRead: nextRead }),
            }).catch(() => {});
          }
          return { ...m, isRead: nextRead };
        })
      );
    },
    [localDbAvailable]
  );

  const handleToggleFlagEmail = useCallback(
    (id: string) => {
      setEmails((prev) =>
        prev.map((m) => {
          if (m.id !== id) return m;
          const nextFlag = !m.isFlagged;
          if (localDbAvailable) {
            void fetch(`/api/local/emails/${encodeURIComponent(id)}`, {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ isFlagged: nextFlag }),
            }).catch(() => {});
          }
          return { ...m, isFlagged: nextFlag };
        })
      );
    },
    [localDbAvailable]
  );

  const handleDeleteEmail = useCallback(
    (id: string) => {
      setDeletedIds((prev) => (prev.includes(id) ? prev : [...prev, id]));
      setOptions((prev) => ({
        ...prev,
        selectedIds: prev.selectedIds.filter((s) => s !== id),
      }));
      if (localDbAvailable) {
        void fetch(`/api/local/emails/${encodeURIComponent(id)}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ isDeleted: true }),
        }).catch(() => {});
      }
    },
    [localDbAvailable]
  );

  const handleRestoreEmail = useCallback(
    (id: string) => {
      setDeletedIds((prev) => prev.filter((item) => item !== id));
      if (localDbAvailable) {
        void fetch(`/api/local/emails/${encodeURIComponent(id)}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ isDeleted: false }),
        }).catch(() => {});
      }
    },
    [localDbAvailable]
  );

  const handleEmptyTrash = useCallback(() => {
    const trashedSet = new Set(deletedIds);
    setEmails((prev) => prev.filter((m) => !trashedSet.has(m.id)));
    setDeletedIds([]);
  }, [deletedIds]);

  const handleMoveEmailToFolder = useCallback((id: string, folder: EmailFolder) => {
    setDeletedIds((prev) => prev.filter((item) => item !== id));
    setEmails((prev) => prev.map((m) => (m.id === id ? { ...m, folder } : m)));
  }, []);

  const persistKvState = useCallback((key: string, value: unknown) => {
    if (!localDbAvailable) return;
    void fetch('/api/local/state', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key, value }),
    }).catch(() => {});
  }, [localDbAvailable]);

  const handleChangeOutlookPrefs = useCallback((next: OutlookPreferences) => {
    const clean = sanitizeOutlookPreferences(next);
    setOutlookPrefs(clean);
    try {
      window.localStorage.setItem(OUTLOOK_PREFS_STORAGE_KEY, JSON.stringify(clean));
    } catch {
      // ignore storage quota
    }
    persistKvState('outlookPrefs', clean);
  }, [persistKvState]);

  const handleChangeOutlookTasks = useCallback((next: OutlookTaskItem[]) => {
    setOutlookTasks(next);
    try {
      window.localStorage.setItem(OUTLOOK_TASKS_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore
    }
    persistKvState('outlook_tasks', next);
  }, [persistKvState]);

  const handleChangeMailCategoriesMap = useCallback((next: Record<string, string[]>) => {
    setMailCategoriesMap(next);
    try {
      window.localStorage.setItem(MAIL_CATEGORIES_MAP_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore
    }
    persistKvState('mail_categories_map', next);
  }, [persistKvState]);

  const handleChangeMailSnoozeMap = useCallback((next: Record<string, string>) => {
    setMailSnoozeMap(next);
    try {
      window.localStorage.setItem(MAIL_SNOOZE_MAP_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore
    }
    persistKvState('mail_snooze_map', next);
  }, [persistKvState]);

  const handleArchiveEmail = useCallback((id: string) => {
    setDeletedIds((prev) => prev.filter((item) => item !== id));
    setEmails((prev) =>
      prev.map((m) => (m.id === id ? { ...m, isRead: true, folderName: '보관함' } : m))
    );
    if (localDbAvailable) {
      void fetch(`/api/local/emails/${encodeURIComponent(id)}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ isRead: true }),
      }).catch(() => {});
    }
  }, [localDbAvailable]);

  const handleToggleFavoriteFolder = useCallback((folderId: string) => {
    setFavoriteFolderIds((prev) => {
      const next = toggleFavoriteFolder(prev, folderId);
      try {
        window.localStorage.setItem(FAVORITE_FOLDERS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleTogglePinEmail = useCallback((emailId: string) => {
    setPinnedEmailIds((prev) => {
      const next = prev.includes(emailId)
        ? prev.filter((id) => id !== emailId)
        : [emailId, ...prev];
      try {
        window.localStorage.setItem(PINNED_EMAILS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleAddProjectSubfolder = useCallback(
    (projectCode: string, name: string, parentId: string | null) => {
      const entry: ProjectTreeSubfolder = {
        id: `psf-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        projectCode,
        parentId,
        name: name.trim(),
        createdAt: new Date().toISOString(),
      };
      setProjectSubfolders((prev) => {
        const next = [...prev, entry];
        try {
          window.localStorage.setItem(PROJECT_SUBFOLDERS_STORAGE_KEY, JSON.stringify(next));
        } catch {
          // ignore
        }
        return next;
      });
    },
    []
  );

  const handleDeleteProjectSubfolder = useCallback((subfolderId: string) => {
    setProjectSubfolders((prevFolders) => {
      let updatedFolders = prevFolders;
      setProjectVaultFiles((prevFiles) => {
        const cleaned = removeProjectSubfolderSafely(prevFolders, prevFiles, subfolderId);
        updatedFolders = cleaned.subfolders;
        try {
          window.localStorage.setItem(PROJECT_SUBFOLDERS_STORAGE_KEY, JSON.stringify(cleaned.subfolders));
          window.localStorage.setItem(PROJECT_VAULT_FILES_STORAGE_KEY, JSON.stringify(cleaned.files));
        } catch {
          // ignore
        }
        return cleaned.files;
      });
      return updatedFolders;
    });
  }, []);

  const handleAddProjectVaultFiles = useCallback(
    (projectCode: string, subfolderId: string | null, fileList: File[]) => {
      const created = fileList.map((f) =>
        createProjectVaultFile({
          projectCode,
          subfolderId,
          filename: f.name,
          sizeBytes: f.size,
          mimeType: f.type || 'application/octet-stream',
          autoIndexDrawing: outlookPrefs.sedms.autoIndexProjectFiles,
        })
      );
      setProjectVaultFiles((prev) => {
        const next = [...created, ...prev];
        try {
          window.localStorage.setItem(PROJECT_VAULT_FILES_STORAGE_KEY, JSON.stringify(next));
        } catch {
          // ignore
        }
        return next;
      });
    },
    [outlookPrefs.sedms.autoIndexProjectFiles]
  );

  const handleRemoveProjectVaultFile = useCallback((fileId: string) => {
    setProjectVaultFiles((prev) => {
      const next = prev.filter((f) => f.id !== fileId);
      try {
        window.localStorage.setItem(PROJECT_VAULT_FILES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleAssignEmailToProject = useCallback((projectCode: string, emailId: string) => {
    setProjectMailOverrides((prev) => {
      const next = assignEmailToProject(prev, projectCode, emailId);
      try {
        window.localStorage.setItem(PROJECT_MAIL_OVERRIDES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleRemoveEmailFromProject = useCallback((projectCode: string, emailId: string) => {
    setProjectMailOverrides((prev) => {
      const next = removeEmailFromProject(prev, projectCode, emailId);
      try {
        window.localStorage.setItem(PROJECT_MAIL_OVERRIDES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleAddProject = useCallback((proj: Omit<ProjectDefinition, 'id' | 'createdAt'>) => {
    const entry: ProjectDefinition = {
      ...proj,
      id: `prj-${Date.now()}`,
      createdAt: new Date().toISOString(),
    };
    setProjects((prev) => {
      const next = [entry, ...prev];
      try {
        window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleToggleFavoriteProject = useCallback((id: string) => {
    setProjects((prev) => {
      const next = prev.map((p) => (p.id === id ? { ...p, isFavorite: !p.isFavorite } : p));
      try {
        window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleUpdateProject = useCallback((id: string, patch: Partial<ProjectDefinition>) => {
    setProjects((prev) => {
      const next = prev.map((p) => (p.id === id ? { ...p, ...patch } : p));
      try {
        window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleImportProjectRules = useCallback((imported: ProjectDefinition[]) => {
    setProjects((prev) => {
      const byId = new Map(prev.map((p) => [p.id, p]));
      const byCode = new Map(prev.map((p) => [p.code.trim().toLowerCase(), p]));
      const merged = [...prev];
      for (const item of imported) {
        const existing =
          byId.get(item.id) || byCode.get(item.code.trim().toLowerCase());
        if (existing) {
          const idx = merged.findIndex((p) => p.id === existing.id);
          if (idx >= 0) {
            merged[idx] = { ...existing, ...item, id: existing.id };
          }
        } else {
          merged.unshift({
            ...item,
            id: item.id || `prj-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
            createdAt: item.createdAt || new Date().toISOString(),
          });
        }
      }
      try {
        window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(merged));
      } catch {
        // ignore
      }
      return merged;
    });
  }, []);

  const handleDeleteProject = useCallback((id: string) => {
    setProjects((prev) => {
      const next = prev.filter((p) => p.id !== id);
      try {
        window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleAddManualEvent = useCallback((ev: Omit<CalendarScheduleEvent, 'id'>) => {
    const item: CalendarScheduleEvent = {
      ...ev,
      id: `cal-${Date.now()}`,
      category: ev.category || 'manual',
    };
    setManualEvents((prev) => {
      const next = [item, ...prev];
      try {
        window.localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleToggleCompleteEvent = useCallback((id: string) => {
    setCompletedEventIds((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  }, []);

  const handleDeleteManualEvent = useCallback((id: string) => {
    setManualEvents((prev) => {
      const next = prev.filter((e) => e.id !== id);
      try {
        window.localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleImportIcsEvents = useCallback((incoming: Array<Omit<CalendarScheduleEvent, 'id'>>) => {
    let added = 0;
    setManualEvents((prev) => {
      const existingKeys = new Set(prev.map((e) => `${e.date}::${e.title.trim()}::${e.timeLabel || ''}`));
      const toAdd: CalendarScheduleEvent[] = [];
      incoming.forEach((ev, idx) => {
        const key = `${ev.date}::${ev.title.trim()}::${ev.timeLabel || ''}`;
        if (!existingKeys.has(key)) {
          existingKeys.add(key);
          toAdd.push({
            ...ev,
            id: `ics-${Date.now()}-${idx}`,
          });
        }
      });
      added = toAdd.length;
      if (added === 0) return prev;
      const next = [...toAdd, ...prev];
      try {
        window.localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
    return added;
  }, []);

  const handleImportDroppedEmails = useCallback((incoming: MailItem[]) => {
    if (!Array.isArray(incoming) || incoming.length === 0) return;
    setEmails((prev) => {
      const seen = new Set(prev.map((m) => m.id));
      const fresh = incoming.filter((m) => !seen.has(m.id));
      return fresh.length > 0 ? [...fresh, ...prev] : prev;
    });
    setGraphNotice(`메일 파일(.msg/.eml) ${incoming.length}건을 즉시 파싱해 워크벤치에 추가했습니다.`);
  }, []);

  const handleAddMeetingMinutes = useCallback((result: MeetingMinutesResult) => {
    setMeetingMinutesHistory((prev) => {
      const next = [result, ...prev].slice(0, 30);
      try {
        window.localStorage.setItem(MINUTES_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleSyncMinutesToCalendar = useCallback((minutes: MeetingMinutesResult) => {
    const newItems: CalendarScheduleEvent[] = minutes.actionItems.map((it, idx) => {
      const iso = it.deadline.match(/(20\d{2})[-./](\d{1,2})[-./](\d{1,2})/);
      const dt = iso
        ? `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`
        : minutes.meetingDate;
      return {
        id: `mtg-sync-${minutes.id}-${idx}`,
        title: `[회의 실행] ${it.task}`,
        date: dt,
        timeLabel: '18:00',
        category: 'action-item',
        notes: `담당: ${it.owner} · 회의록: ${minutes.title}`,
      };
    });
    if (newItems.length === 0) return;
    setManualEvents((prev) => {
      const existingIds = new Set(prev.map((e) => e.id));
      const merged = [...newItems.filter((n) => !existingIds.has(n.id)), ...prev];
      try {
        window.localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(merged));
      } catch {
        // ignore
      }
      return merged;
    });
    setGraphNotice(`회의록 실행 과제 ${newItems.length}건을 Outlook 일정 달력에 자동 반영했습니다.`);
  }, []);

  const handleExportMbox = useCallback(async (subset: MailItem[], filename: string) => {
    try {
      const res = await fetch('/api/export/mbox', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ emails: subset, filename }),
      });
      if (!res.ok) return;
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch {
      // ignore
    }
  }, []);

  const handleApplyDeterministicWorkLog = useCallback((result: AnalysisResult) => {
    setAnalysisResult(result);
    setArchive((prev) => {
      const next: SavedArchiveEntry[] = [
        { id: result.id, savedAt: result.createdAt, result },
        ...prev,
      ].slice(0, 40);
      try {
        window.localStorage.setItem(ARCHIVE_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
    setGraphNotice('LLM 호출 없이 자동 코드 엔진으로 업무 일지를 생성해 기록함에 저장했습니다.');
    setPage('archive');
  }, []);

  const handleSaveDraft = useCallback(
    (draft: Omit<OutlookDraftItem, 'id' | 'updatedAt'>) => {
      const item: OutlookDraftItem = {
        ...draft,
        id: `draft-${Date.now()}`,
        updatedAt: new Date().toISOString(),
      };
      setDrafts((prev) => {
        const next = [item, ...prev].slice(0, 50);
        try {
          window.localStorage.setItem(DRAFTS_STORAGE_KEY, JSON.stringify(next));
        } catch {
          // ignore
        }
        return next;
      });
    },
    []
  );

  const handleDeleteDraft = useCallback((id: string) => {
    setDrafts((prev) => {
      const next = prev.filter((d) => d.id !== id);
      try {
        window.localStorage.setItem(DRAFTS_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handleCaptureCurrentOutlookItem = useCallback(async () => {
    // Support Mailbox 1.15 Multi-Select batch when connected to Graph, or current active item
    const selectedRestIds = await getSelectedOutlookRestIds();
    if (selectedRestIds.length > 1 && microsoftAuth.connected) {
      try {
        const res = await fetch('/api/outlook/batch-messages', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ restIds: selectedRestIds }),
        });
        const data = (await res.json()) as { emails?: MailItem[] };
        if (res.ok && data.emails && data.emails.length > 0) {
          setEmails((prev) => {
            const existingIds = new Set(prev.map((m) => m.id));
            return [...data.emails!.filter((m) => !existingIds.has(m.id)), ...prev];
          });
          setOptions((prev) => ({
            ...prev,
            scope: 'selected',
            selectedIds: Array.from(new Set([...data.emails!.map((m) => m.id), ...prev.selectedIds])),
          }));
          return;
        }
      } catch {
        // Fall through to single active item capture
      }
    }

    const current = await readCurrentOutlookItem();
    if (!current) return;
    setEmails((prev) => {
      const exists = prev.some((m) => m.id === current.id);
      return exists ? prev : [current, ...prev];
    });
    setOptions((prev) => ({
      ...prev,
      scope: 'selected',
      selectedIds: prev.selectedIds.includes(current.id)
        ? prev.selectedIds
        : [current.id, ...prev.selectedIds],
    }));
  }, [microsoftAuth.connected]);

  /**
   * Executes the Autonomous Multi-Step AI Agent over Server-Sent Events (/api/agent/run).
   */
  const executeAutonomousAgent = useCallback(
    async (
      targetEmails: MailItem[],
      targetOptions: AnalysisOptions,
      overrideProvider?: AiProviderId
    ) => {
      const activeList = targetEmails.filter((m) => !deletedIds.includes(m.id));
      if (activeList.length === 0) {
        setAnalysisError('연동된 실제 메일이 없습니다. 먼저 .pst/.eml 파일을 열거나 Microsoft 365 계정을 연결해 주세요.');
        return;
      }

      setIsAnalyzing(true);
      setAnalysisError(null);
      setAgentSteps([]);
      const effectiveOptions: AnalysisOptions = {
        ...targetOptions,
        provider: overrideProvider ?? targetOptions.provider,
      };

      try {
        const agentHeaders: Record<string, string> = { 'Content-Type': 'application/json' };
        if (pop3SessionToken) {
          agentHeaders.Authorization = `Bearer ${pop3SessionToken}`;
        }
        const res = await fetch('/api/agent/run', {
          method: 'POST',
          headers: agentHeaders,
          body: JSON.stringify({
            emails: activeList,
            options: effectiveOptions,
            agentGoal: effectiveOptions.customFocus,
            autoPickLatestDateIfEmpty: true,
          }),
        });

        if (!res.ok || !res.body) {
          throw new Error('AI 에이전트 스트림 연결에 실패했습니다.');
        }

        const reader = res.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let finalResult: AnalysisResult | null = null;
        let streamError: string | null = null;

        while (true) {
          const { value, done } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });

          const blocks = buffer.split('\n\n');
          buffer = blocks.pop() ?? '';

          for (const block of blocks) {
            const lines = block.split('\n');
            let eventType = 'message';
            let dataStr = '';
            for (const line of lines) {
              if (line.startsWith('event:')) eventType = line.slice(6).trim();
              if (line.startsWith('data:')) dataStr += line.slice(5).trim();
            }
            if (!dataStr) continue;

            try {
              const parsed = JSON.parse(dataStr);
              if (eventType === 'steps' && Array.isArray(parsed.steps)) {
                setAgentSteps(parsed.steps as AgentExecutionStep[]);
              } else if (eventType === 'result') {
                finalResult = parsed as AnalysisResult;
              } else if (eventType === 'error') {
                streamError = String(parsed.error || 'AI 에이전트 실행 오류');
              }
            } catch {
              // ignore malformed SSE chunk
            }
          }
        }

        if (streamError) {
          throw new Error(streamError);
        }
        if (!finalResult) {
          throw new Error('AI 에이전트가 결과를 반환하지 못했습니다.');
        }

        setAnalysisResult(finalResult);
        if (finalResult.agentSteps) {
          setAgentSteps(finalResult.agentSteps);
        }
        if (standaloneTaskpane) {
          setTaskpaneSubView('result');
        }

        if (appConfig.agent.autoSaveArchive) {
          const entry: SavedArchiveEntry = {
            id: finalResult.id,
            savedAt: finalResult.createdAt,
            result: finalResult,
          };
          setArchive((prev) => {
            const next = [entry, ...prev].slice(0, 40);
            try {
              window.localStorage.setItem(ARCHIVE_STORAGE_KEY, JSON.stringify(next));
            } catch {
              // ignore quota error
            }
            return next;
          });
        }
      } catch (err) {
        setAnalysisError(
          err instanceof Error ? err.message : 'AI 에이전트 실행 중 네트워크 오류가 발생했습니다.'
        );
      } finally {
        setIsAnalyzing(false);
      }
    },
    [deletedIds, standaloneTaskpane, appConfig.agent.autoSaveArchive, pop3SessionToken]
  );

  const handleRunAnalysis = useCallback(
    (overrideProvider?: AiProviderId) => {
      void executeAutonomousAgent(emails, options, overrideProvider);
    },
    [emails, options, executeAutonomousAgent]
  );

  const handleSaveAppConfig = useCallback(
    async (nextConfig: AppConfig, adminPassword?: string): Promise<{ ok: boolean; error?: string }> => {
      try {
        const headers: Record<string, string> = { 'Content-Type': 'application/json' };
        if (adminPassword) {
          headers['X-Config-Password'] = adminPassword;
        }
        const res = await fetch('/api/config', {
          method: 'PUT',
          headers,
          body: JSON.stringify(nextConfig),
        });
        const data = (await res.json()) as {
          config?: AppConfig;
          providers?: ProviderStatus[];
          error?: string;
        };
        if (!res.ok || !data.config) {
          return { ok: false, error: data.error || 'config.json 저장 실패' };
        }
        setAppConfig(data.config);
        if (data.providers) setProviders(data.providers);
        setOptions((prev) => ({
          ...prev,
          scope: data.config!.agent.defaultScope,
          dateMode: data.config!.agent.defaultDateMode,
          blogTone: data.config!.agent.defaultBlogTone,
          maskSensitiveInfo: data.config!.agent.autoMaskSensitiveInfo,
          includeActionItems: data.config!.agent.includeActionItems,
        }));
        return { ok: true };
      } catch {
        return { ok: false, error: '서버와 통신하지 못했습니다.' };
      }
    },
    []
  );

  const handleConnectMicrosoft = useCallback(async () => {
    try {
      const res = await fetch('/api/auth/microsoft/url');
      const data = (await res.json()) as { url?: string; error?: string };
      if (!res.ok || !data.url) {
        setGraphNotice(data.error || '.env(Secrets)의 Microsoft OAuth 설정을 확인해 주세요.');
        return;
      }
      const popup = window.open(data.url, 'ms_outlook_oauth', 'width=560,height=680');
      if (!popup) {
        setGraphNotice('팝업이 차단되었습니다. 브라우저에서 팝업을 허용한 뒤 다시 시도해 주세요.');
      }
    } catch {
      setGraphNotice('Microsoft 로그인 URL을 생성하지 못했습니다.');
    }
  }, []);

  const handleSyncMicrosoftEmails = useCallback(
    async (silent = false) => {
      setIsSyncingGraph(true);
      if (!silent) setGraphNotice(null);
      try {
        const startDate = options.dateMode === 'single' ? options.singleDate : options.startDate;
        const endDate = options.dateMode === 'single' ? options.singleDate : options.endDate;
        const res = await fetch('/api/outlook/messages', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ startDate, endDate, folders: ['inbox', 'sent'] }),
        });
        const data = (await res.json()) as { emails?: MailItem[]; error?: string };
        if (!res.ok || !data.emails) {
          throw new Error(data.error || '실제 Outlook 메일 조회에 실패했습니다.');
        }
        if (data.emails.length > 0) {
          if (silent) {
            void notifyArrivedEmails(data.emails);
          } else {
            for (const m of data.emails) {
              if (m?.id) seenEmailIdsRef.current.add(m.id);
            }
          }
          setEmails((prev) => {
            if (prev.length === 0) return data.emails!;
            const existingIds = new Set(prev.map((m) => m.id));
            const fresh = data.emails!.filter((m) => !existingIds.has(m.id));
            return fresh.length > 0 ? [...fresh, ...prev] : prev;
          });
          setDeletedIds([]);
          if (!silent) {
            setGraphNotice(`Microsoft 365 Outlook에서 ${data.emails.length}건의 실제 메일을 동기화했습니다.`);
          }
        } else if (!silent) {
          setGraphNotice('선택한 기간에 해당하는 실제 Outlook 메일이 없습니다.');
        }
        setLastSyncedAt(new Date().toISOString());
        setLastSyncSource('microsoft365');
      } catch (err) {
        if (!silent) {
          setGraphNotice(err instanceof Error ? err.message : '동기화 오류가 발생했습니다.');
        }
      } finally {
        setIsSyncingGraph(false);
      }
    },
    [options, notifyArrivedEmails]
  );

  const handleDisconnectMicrosoft = useCallback(async () => {
    await fetch('/api/auth/microsoft/logout', { method: 'POST' });
    setMicrosoftAuth((prev) => ({ ...prev, connected: false, userEmail: null, displayName: null }));
    setEmails([]);
    setDeletedIds([]);
    setGraphNotice('Microsoft 계정 연결을 해제했습니다.');
  }, []);

  const handlePop3Connected = useCallback(
    (payload: {
      sessionToken: string;
      account: Pop3ConnectedAccount;
      emails: MailItem[];
      totalMailboxCount: number;
    }) => {
      setPop3SessionToken(payload.sessionToken);
      setPop3Account(payload.account);
      setConnectedStoreName(`${payload.account.providerName} (${payload.account.email})`);
      setHostInfo((prev) => ({ ...prev, userEmail: payload.account.email }));
      if (payload.emails.length > 0) {
        for (const m of payload.emails) {
          if (m?.id) seenEmailIdsRef.current.add(m.id);
        }
        setEmails((prev) => {
          const existingIds = new Set(prev.map((m) => m.id));
          const fresh = payload.emails.filter((m) => !existingIds.has(m.id));
          return [...fresh, ...prev];
        });
      }
      setLastSyncedAt(new Date().toISOString());
      setLastSyncSource('pop3');
      setGraphNotice(
        `${payload.account.email} (${(payload.account.protocol || 'pop3').toUpperCase()} · ${
          payload.account.protocol === 'imap'
            ? `${payload.account.imapHost || payload.account.pop3Host}:${payload.account.imapPort || 993}`
            : `${payload.account.pop3Host}:${payload.account.pop3Port}`
        }) 연결 완료 · 최신 메일 ${payload.emails.length}건 불러옴 (전체 ${payload.totalMailboxCount}건)`
      );
    },
    []
  );

  const handleSyncPop3 = useCallback(
    async (silent = false) => {
      if (!pop3SessionToken) {
        if (!silent) setIsPop3ModalOpen(true);
        return;
      }
      setIsSyncingPop3(true);
      try {
        const res = await fetch('/api/pop3/sync', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${pop3SessionToken}`,
          },
          body: JSON.stringify({ maxMessages: 30 }),
        });
        const data = (await res.json()) as {
          emails?: MailItem[];
          totalMailboxCount?: number;
          error?: string;
        };
        if (!res.ok) {
          if (res.status === 401) {
            setPop3SessionToken(null);
            setPop3Account(null);
          }
          if (!silent) setGraphNotice(data.error || 'POP3/IMAP 새로고침에 실패했습니다.');
          return;
        }
        if (data.emails && data.emails.length > 0) {
          void notifyArrivedEmails(data.emails);
          setEmails((prev) => {
            const existingIds = new Set(prev.map((m) => m.id));
            const fresh = data.emails!.filter((m) => !existingIds.has(m.id));
            return fresh.length > 0 ? [...fresh, ...prev] : prev;
          });
        }
        setLastSyncedAt(new Date().toISOString());
        setLastSyncSource('pop3');
        if (!silent) {
          setGraphNotice(`메일함을 새로고침했습니다 (${data.emails?.length || 0}건 확인).`);
        }
      } catch {
        if (!silent) setGraphNotice('메일 서버와 통신하지 못했습니다.');
      } finally {
        setIsSyncingPop3(false);
      }
    },
    [pop3SessionToken, notifyArrivedEmails]
  );

  // Automatic initial & 60-second background mail arrival polling for always-logged-in POP3/IMAP & M365 sessions
  useEffect(() => {
    if (!pop3SessionToken || standaloneTaskpane || sgwEmbedded) return;
    if (!hasAutoSyncedPop3Ref.current) {
      hasAutoSyncedPop3Ref.current = true;
      window.setTimeout(() => {
        void handleSyncPop3(true);
      }, 2000);
    }
    const id = window.setInterval(() => {
      if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return;
      void handleSyncPop3(true);
    }, 60_000);
    const onVis = () => {
      if (document.visibilityState === 'visible') {
        void handleSyncPop3(true);
      }
    };
    document.addEventListener('visibilitychange', onVis);
    return () => {
      window.clearInterval(id);
      document.removeEventListener('visibilitychange', onVis);
    };
  }, [pop3SessionToken, standaloneTaskpane, sgwEmbedded, handleSyncPop3]);

  useEffect(() => {
    if (!microsoftAuth.connected || standaloneTaskpane || sgwEmbedded) return;
    const id = window.setInterval(() => {
      if (typeof document !== 'undefined' && document.visibilityState === 'hidden') return;
      void handleSyncMicrosoftEmails(true);
    }, 60_000);
    return () => window.clearInterval(id);
  }, [microsoftAuth.connected, standaloneTaskpane, sgwEmbedded, handleSyncMicrosoftEmails]);

  const handleDisconnectPop3 = useCallback(async () => {
    if (pop3SessionToken) {
      await fetch('/api/pop3/disconnect', {
        method: 'POST',
        headers: { Authorization: `Bearer ${pop3SessionToken}` },
      }).catch(() => {});
    }
    setPop3SessionToken(null);
    setPop3Account(null);
    setGraphNotice('POP3 메일 계정 연결을 해제하고 서버 세션에서 비밀번호를 삭제했습니다.');
  }, [pop3SessionToken]);

  const handleSendViaPop3Smtp = useCallback(
    async (payload: {
      to: string;
      cc: string;
      bcc?: string;
      subject: string;
      body: string;
      htmlBody?: string;
      importance?: 'low' | 'normal' | 'high';
      requestReadReceipt?: boolean;
      attachments?: Array<{ filename: string; contentType?: string; contentBase64: string }>;
    }): Promise<{ ok: boolean; message: string }> => {
      if (!pop3SessionToken || !pop3Account) {
        return { ok: false, message: '연결된 POP3/SMTP 계정이 없습니다.' };
      }
      try {
        const res = await fetch('/api/pop3/send', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${pop3SessionToken}`,
          },
          body: JSON.stringify(payload),
        });
        const data = (await res.json()) as { ok?: boolean; acceptedTo?: string[]; error?: string };
        if (!res.ok || !data.ok) {
          return { ok: false, message: data.error || 'SMTP 발송에 실패했습니다.' };
        }
        const now = new Date();
        const sentItem: MailItem = {
          id: `smtp-sent-${Date.now()}`,
          folder: 'sent',
          subject: payload.subject || '(제목 없음)',
          fromName: pop3Account.email,
          fromEmail: pop3Account.email,
          to: data.acceptedTo || [payload.to],
          date: now.toISOString().slice(0, 10),
          timestamp: now.toISOString(),
          timeLabel: now.toTimeString().slice(0, 5),
          snippet: payload.body.slice(0, 120),
          body: payload.body,
          isRead: true,
          source: 'outlook-live',
        };
        setEmails((prev) => [sentItem, ...prev]);
        return { ok: true, message: `${(data.acceptedTo || [payload.to]).join(', ')}로 메일을 발송했습니다.` };
      } catch {
        return { ok: false, message: 'SMTP 발송 중 네트워크 오류가 발생했습니다.' };
      }
    },
    [pop3SessionToken, pop3Account]
  );

  const handleDeleteArchive = useCallback((id: string) => {
    setArchive((prev) => {
      const next = prev.filter((item) => item.id !== id);
      try {
        window.localStorage.setItem(ARCHIVE_STORAGE_KEY, JSON.stringify(next));
      } catch {
        // ignore
      }
      return next;
    });
  }, []);

  const handlePstImportComplete = useCallback(
    (
      payload: {
        storeName: string;
        fileType: 'pst' | 'eml' | 'mbox' | 'msg';
        folders: string[];
        emails: MailItem[];
        skippedCount: number;
        inheritedProfile?: InheritedOutlookProfile;
      },
      mode: 'replace' | 'merge',
      autoAnalyze = false
    ) => {
      setConnectedStoreName(payload.storeName);
      setPstFolders((prev) =>
        mode === 'replace'
          ? payload.folders
          : Array.from(new Set([...payload.folders, ...prev]))
      );
      const nextEmails = mode === 'replace' ? payload.emails : [...payload.emails, ...emails];
      setEmails(nextEmails);
      if (mode === 'replace') setDeletedIds([]);

      // Inherit Outlook account email & auto-discovered projects directly from PST structure
      const profile = payload.inheritedProfile;
      let addedProjectCount = 0;
      if (profile) {
        if (profile.ownerEmail) {
          setHostInfo((prev) => ({
            ...prev,
            userEmail: profile.ownerEmail,
          }));
        }
        if (profile.discoveredProjects.length > 0) {
          setProjects((prev) => {
            const { merged, addedCount } = mergeInheritedProjects(prev, profile.discoveredProjects);
            addedProjectCount = addedCount;
            try {
              window.localStorage.setItem(PROJECTS_STORAGE_KEY, JSON.stringify(merged));
            } catch (err) {
              const reason = err instanceof Error ? err.message : String(err);
              console.warn(`[App] Failed to persist inherited projects: ${reason}`);
            }
            return merged;
          });
        }
      }

      const newestDate = profile?.dateRange.latest || payload.emails[0]?.date;
      const earliestDate = profile?.dateRange.earliest || payload.emails[payload.emails.length - 1]?.date || newestDate;
      const hasMultiDayRange = Boolean(earliestDate && newestDate && earliestDate !== newestDate);
      const nextOptions: AnalysisOptions = newestDate
        ? {
            ...options,
            dateMode: hasMultiDayRange ? 'range' : 'single',
            singleDate: newestDate,
            startDate: earliestDate || newestDate,
            endDate: newestDate,
            scope: 'all',
            selectedIds: payload.emails.slice(0, 3).map((m) => m.id),
          }
        : options;
      setOptions(nextOptions);

      const summaryParts = [
        `${payload.storeName} (${payload.fileType.toUpperCase()}) Outlook 설정 인수 완료`,
        `메일 ${payload.emails.length}건`,
        `폴더 ${payload.folders.length}개`,
      ];
      if (profile?.ownerEmail) summaryParts.push(`계정 ${profile.ownerEmail}`);
      if (profile && profile.discoveredProjects.length > 0) {
        summaryParts.push(`프로젝트 ${addedProjectCount || profile.discoveredProjects.length}개 자동 등록`);
      }
      if (profile && profile.calendarHintCount > 0) {
        summaryParts.push(`일정 ${profile.calendarHintCount}건 반영`);
      }
      setGraphNotice(summaryParts.join(' · '));
      setPage('workbench');

      if ((autoAnalyze || appConfig.agent.autoRunOnImport) && nextEmails.length > 0) {
        setIsPstModalOpen(false);
        setIsTaskpaneOpen(true);
        const aiOptions: AnalysisOptions = {
          ...nextOptions,
          customFocus:
            nextOptions.customFocus ||
            'PST에서 인수한 주고받은 모든 메일에서 내가 직접 한 일과 프로젝트별 진행·결정·첨부·후속 조치를 빠짐없이 정리해 업무 일지와 블로그 초안을 작성하세요.',
        };
        setOptions(aiOptions);
        void executeAutonomousAgent(nextEmails, aiOptions);
      }
    },
    [emails, options, appConfig.agent.autoRunOnImport, executeAutonomousAgent]
  );

  const handleToggleTaskpanePin = useCallback(() => {
    if (!appConfig.plugin.supportsPinning) return;
    setIsTaskpanePinned((prev) => {
      const next = !prev;
      if (next) setIsTaskpaneOpen(true);
      return next;
    });
  }, [appConfig.plugin.supportsPinning]);

  const handleRibbonIconClick = useCallback(() => {
    setIsTaskpaneOpen((prev) => {
      if (!prev) {
        // Opening from Outlook Ribbon icon restores always-open pinned default if SupportsPinning is active
        if (appConfig.plugin.supportsPinning) {
          setIsTaskpanePinned(true);
          void applyTaskpaneAlwaysOpenBehavior(true);
        }
        return true;
      }
      // If currently pinned, keep it open (matching Outlook SupportsPinning behavior); otherwise toggle closed
      if (appConfig.plugin.supportsPinning && isTaskpanePinned) {
        return true;
      }
      return false;
    });
  }, [appConfig.plugin.supportsPinning, isTaskpanePinned]);

  const handleCloseTaskpane = useCallback(() => {
    setIsTaskpaneOpen(false);
    setIsTaskpanePinned(false);
  }, []);

  // --- Embedded SGW Groupware Main-Menu Mode (?embed=sgw or /sgw/mail) ---
  if (sgwEmbedded) {
    return (
      <div className="h-screen overflow-hidden bg-[#F5F5F5] flex flex-col">
        <SgwSedmsHubView embeddedMode initialTab="sgw-mail" />
      </div>
    );
  }

  const accent = ACCENT_THEME_COLORS[outlookPrefs.general.accentTheme] ?? ACCENT_THEME_COLORS.blue;

  // --- Focused Standalone Outlook Plugin Taskpane (/taskpane) ---
  if (standaloneTaskpane) {
    return (
      <div className="h-screen overflow-hidden bg-[#F5F5F5] text-[#242424] flex flex-col">
        <div className="shrink-0 flex items-center justify-center gap-2 p-2 bg-[#FAF9F8] border-b border-[#EDEBE9]">
          <div className="inline-flex items-center p-0.5 bg-[#F0F0F0] rounded-[6px] gap-0.5">
            <button
              type="button"
              onClick={() => setTaskpaneSubView('options')}
              className={`px-3 py-1 text-xs rounded-[4px] whitespace-nowrap transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                taskpaneSubView === 'options'
                  ? 'bg-white text-[#0F6CBD] font-semibold shadow-[0_1px_2px_rgba(0,0,0,0.08)]'
                  : 'text-[#424242] hover:text-[#242424]'
              }`}
            >
              AI 에이전트 옵션
            </button>
            <button
              type="button"
              onClick={() => setTaskpaneSubView('result')}
              className={`px-3 py-1 text-xs rounded-[4px] whitespace-nowrap transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                taskpaneSubView === 'result'
                  ? 'bg-white text-[#0F6CBD] font-semibold shadow-[0_1px_2px_rgba(0,0,0,0.08)]'
                  : 'text-[#424242] hover:text-[#242424]'
              }`}
            >
              생성 결과 {analysisResult ? '(완료)' : ''}
            </button>
          </div>
        </div>
        {taskpaneSubView === 'options' ? (
          <div className="flex-1 min-h-0 flex">
            <OutlookTaskpane
              standalone
              emails={emails}
              options={options}
              providers={providers}
              isAnalyzing={isAnalyzing}
              agentSteps={agentSteps}
              error={analysisError}
              hostInfo={hostInfo}
              isPinned={isTaskpanePinned}
              supportsPinning={appConfig.plugin.supportsPinning}
              accentColor={accent.primary}
              onTogglePin={handleToggleTaskpanePin}
              onChangeOptions={handleChangeOptions}
              onRunAnalysis={handleRunAnalysis}
              onCaptureCurrentOutlookItem={handleCaptureCurrentOutlookItem}
            />
          </div>
        ) : (
          <div className="flex-1 min-h-0 overflow-y-auto max-w-[440px] w-full mx-auto bg-white border-x border-[#E0E0E0]">
            <Suspense fallback={<div className="p-4 text-xs text-[#616161]">불러오는 중...</div>}>
              <AnalysisResultPanel
                result={analysisResult}
                isAnalyzing={isAnalyzing}
                agentSteps={agentSteps}
                emails={emails}
                projects={projects}
              />
            </Suspense>
          </div>
        )}
      </div>
    );
  }

  const userInitials = microsoftAuth.connected
    ? (microsoftAuth.displayName || microsoftAuth.userEmail || 'BM')
        .slice(0, 2)
        .toUpperCase()
    : 'BM';

  const studioMenuGroups: Array<{
    groupId: string;
    items: Array<{
      id: string;
      label: string;
      icon: React.ReactNode;
      active: boolean;
      badge?: number;
      onClick: () => void;
    }>;
  }> = [
    {
      groupId: 'mail-ai',
      items: [
        {
          id: 'companion',
          label: 'Outlook 동반 검증',
          icon: <ShieldCheck className="w-4 h-4" />,
          active: page === 'companion',
          onClick: () => setPage('companion'),
        },
        {
          id: 'workbench',
          label: '메일함',
          icon: <Mail className="w-4 h-4" />,
          active: page === 'workbench',
          badge: emails.length - deletedIds.length > 0 ? emails.length - deletedIds.length : undefined,
          onClick: () => setPage('workbench'),
        },
      ],
    },
    {
      groupId: 'connected-projects',
      items: [
        {
          id: 'voice',
          label: '음성 · 회의록 (Voice)',
          icon: <Mic className="w-4 h-4" />,
          active: page === 'studio' && studioPreset.tab === 'meeting',
          badge: meetingMinutesHistory.length > 0 ? meetingMinutesHistory.length : undefined,
          onClick: () => {
            setStudioPreset({ tab: 'meeting' });
            setPage('studio');
          },
        },
        {
          id: 'documents',
          label: '문서 · 검색',
          icon: <FileSearch className="w-4 h-4" />,
          active: page === 'studio' && studioPreset.tab !== 'meeting',
          badge: parsedDocuments.length > 0 ? parsedDocuments.length : undefined,
          onClick: () => {
            setStudioPreset({ tab: 'documents' });
            setPage('studio');
          },
        },
        {
          id: 'sgw-sedms',
          label: 'SGW · SEDMS',
          icon: <Building2 className="w-4 h-4" />,
          active: page === 'sgw-sedms',
          onClick: () => setPage('sgw-sedms'),
        },
        {
          id: 'projects',
          label: '프로젝트 · 파일정리',
          icon: <Briefcase className="w-4 h-4" />,
          active: page === 'projects',
          badge: projects.length > 0 ? projects.length : undefined,
          onClick: () => setPage('projects'),
        },
      ],
    },
    {
      groupId: 'schedule-analytics',
      items: [
        {
          id: 'dashboard',
          label: '대시보드 · 일간브리핑',
          icon: <BarChart3 className="w-4 h-4" />,
          active: page === 'dashboard',
          onClick: () => setPage('dashboard'),
        },
      ],
    },
  ];

  const railItems: Array<{
    id: string;
    label: string;
    icon: React.ReactNode;
    active: boolean;
    badge?: number;
    onClick: () => void;
  }> = [
    {
      id: 'companion',
      label: 'Outlook 동반 검증',
      icon: (
        <ShieldCheck
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'companion' ? 2 : 1.65}
        />
      ),
      active: page === 'companion',
      onClick: () => setPage('companion'),
    },
    {
      id: 'workbench',
      label: '메일함 (Outlook)',
      icon: (
        <Mail
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'workbench' ? 2 : 1.65}
        />
      ),
      active: page === 'workbench',
      badge: emails.length - deletedIds.length > 0 ? emails.length - deletedIds.length : undefined,
      onClick: () => setPage('workbench'),
    },
    {
      id: 'dashboard',
      label: '대시보드 · 일간브리핑',
      icon: (
        <BarChart3
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'dashboard' ? 2 : 1.65}
        />
      ),
      active: page === 'dashboard',
      onClick: () => setPage('dashboard'),
    },
    {
      id: 'projects',
      label: '프로젝트 · 파일정리',
      icon: (
        <Briefcase
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'projects' ? 2 : 1.65}
        />
      ),
      active: page === 'projects',
      badge: projects.length > 0 ? projects.length : undefined,
      onClick: () => setPage('projects'),
    },
    {
      id: 'people',
      label: '사람 · 주소록',
      icon: (
        <Users
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'people' ? 2 : 1.65}
        />
      ),
      active: page === 'people',
      badge:
        outlookPcProfile && outlookPcProfile.contacts.length > 0
          ? outlookPcProfile.contacts.length
          : undefined,
      onClick: () => setPage('people'),
    },
    {
      id: 'voice',
      label: '음성 · 회의록 (Voice)',
      icon: (
        <Mic
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'studio' && studioPreset.tab === 'meeting' ? 2 : 1.65}
        />
      ),
      active: page === 'studio' && studioPreset.tab === 'meeting',
      onClick: () => {
        setStudioPreset({ tab: 'meeting' });
        setPage('studio');
      },
    },
    {
      id: 'documents',
      label: '문서 · 검색 스튜디오',
      icon: (
        <FileSearch
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'studio' && studioPreset.tab !== 'meeting' ? 2 : 1.65}
        />
      ),
      active: page === 'studio' && studioPreset.tab !== 'meeting',
      onClick: () => {
        setStudioPreset({ tab: 'documents' });
        setPage('studio');
      },
    },
    {
      id: 'sgw-chat',
      label: 'SGW 실시간 채팅',
      icon: (
        <MessageSquare
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'sgw-chat' ? 2 : 1.65}
        />
      ),
      badge: sgwUnreadCount > 0 ? sgwUnreadCount : undefined,
      active: page === 'sgw-chat',
      onClick: () => setPage('sgw-chat'),
    },
    {
      id: 'sgw-sedms',
      label: 'SGW 그룹웨어 · SEDMS',
      icon: (
        <Building2
          className="w-[18px] h-[18px]"
          strokeWidth={page === 'sgw-sedms' ? 2 : 1.65}
        />
      ),
      active: page === 'sgw-sedms',
      onClick: () => setPage('sgw-sedms'),
    },
  ];

  const activeEmployeeEmail =
    pop3Account?.email ||
    microsoftAuth.userEmail ||
    hostInfo.userEmail ||
    outlookPcProfile?.accounts[0]?.smtpAddress ||
    null;

  const headerSyncView = useMemo(
    () =>
      buildHeaderAccountSyncView({
        microsoftConnected: microsoftAuth.connected,
        microsoftEmail: microsoftAuth.userEmail,
        pop3Connected: Boolean(pop3Account),
        pop3Email: pop3Account?.email || null,
        pop3Protocol: pop3Account?.protocol,
        pop3ProviderName: pop3Account?.providerName || null,
        isSyncingGraph,
        isSyncingPop3,
        lastSyncedAt,
        lastSyncSource,
      }),
    [
      microsoftAuth.connected,
      microsoftAuth.userEmail,
      pop3Account,
      isSyncingGraph,
      isSyncingPop3,
      lastSyncedAt,
      lastSyncSource,
    ]
  );

  if (standaloneManualPopup) {
    return (
      <Suspense fallback={null}>
        <InteractiveManualModal
          isOpen
          standalonePopup
          onClose={() => {
            if (typeof window !== 'undefined') window.close();
          }}
        />
      </Suspense>
    );
  }

  if (standaloneComposePopup) {
    const params = typeof window !== 'undefined' ? new URLSearchParams(window.location.search) : null;
    const draftKey = params?.get('draftKey') || '';
    let initialComposeDraft: ComposeDraftState = {
      kind: 'new',
      title: '새 메일',
      to: '',
      cc: '',
      subject: '',
    };
    let cachedSigs: OutlookSignatureEntry[] = localSignatures;
    let cachedContacts: OutlookContactEntry[] = outlookPcProfile?.contacts || [];
    if (typeof window !== 'undefined') {
      try {
        if (draftKey) {
          const rawDraft = window.localStorage.getItem(draftKey);
          if (rawDraft) {
            initialComposeDraft = JSON.parse(rawDraft) as ComposeDraftState;
          }
        }
        if (cachedSigs.length === 0) {
          const rawSigs = window.localStorage.getItem('bmail_local_signatures_cache');
          if (rawSigs) cachedSigs = JSON.parse(rawSigs) as OutlookSignatureEntry[];
        }
        if (cachedContacts.length === 0) {
          const rawContacts = window.localStorage.getItem('bmail_local_contacts_cache');
          if (rawContacts) cachedContacts = JSON.parse(rawContacts) as OutlookContactEntry[];
        }
      } catch {
        // ignore storage parse fallback
      }
    }
    return (
      <Suspense fallback={<div className="w-screen h-screen flex items-center justify-center text-xs text-[#616161]">메세지 작성 창 준비 중...</div>}>
        <OutlookComposeWindow
          mode="standalone"
          draft={initialComposeDraft}
          signatures={cachedSigs}
          contacts={cachedContacts}
          outlookPrefs={outlookPrefs}
          pop3ConnectedEmail={pop3Account?.email || null}
          onSendSmtp={handleSendViaPop3Smtp}
          onSaveDraft={handleSaveDraft}
          onClose={() => {
            if (typeof window !== 'undefined') window.close();
          }}
        />
      </Suspense>
    );
  }

  const handleHeaderQuickSync = useCallback(() => {
    if (isSyncingGraph || isSyncingPop3) return;
    if (microsoftAuth.connected) {
      void handleSyncMicrosoftEmails();
    }
    if (pop3SessionToken) {
      void handleSyncPop3();
    }
  }, [
    isSyncingGraph,
    isSyncingPop3,
    microsoftAuth.connected,
    pop3SessionToken,
    handleSyncMicrosoftEmails,
    handleSyncPop3,
  ]);

  const isMobileModeActive =
    uiMode === 'mobile' || (uiMode === 'auto' && isNarrowViewport);

  // --- Dedicated Clean Mobile Mode (Icons + Mail Read/Write + POP3 Connect ONLY) ---
  if (isMobileModeActive) {
    return (
      <Suspense fallback={<div className="h-screen flex items-center justify-center text-xs text-[#616161]">불러오는 중...</div>}>
        <MobileMailShell
          emails={emails.filter((m) => !deletedIds.includes(m.id))}
          drafts={drafts}
          pop3Account={pop3Account}
          isSyncingPop3={isSyncingPop3}
          microsoftConnected={microsoftAuth.connected}
          microsoftEmail={microsoftAuth.userEmail}
          isSyncingGraph={isSyncingGraph}
          lastSyncedAt={lastSyncedAt}
          notice={graphNotice}
          onDismissNotice={() => setGraphNotice(null)}
          onOpenPop3Modal={() => setIsPop3ModalOpen(true)}
          onOpenPstModal={() => setIsPstModalOpen(true)}
          onSyncPop3={() => void handleSyncPop3()}
          onSwitchToDesktop={() => setUiMode('desktop')}
          onToggleRead={handleToggleReadEmail}
          onToggleFlag={handleToggleFlagEmail}
          onDeleteEmail={handleDeleteEmail}
          onSaveDraft={handleSaveDraft}
          onDeleteDraft={handleDeleteDraft}
          onSendViaSmtp={handleSendViaPop3Smtp}
        />
        <Pop3QuickConnectModal
          isOpen={isPop3ModalOpen}
          connectedAccount={pop3Account}
          onClose={() => setIsPop3ModalOpen(false)}
          onConnected={handlePop3Connected}
          onDisconnect={() => void handleDisconnectPop3()}
          onOpenPstFallback={() => setIsPstModalOpen(true)}
        />
        <PstImportModal
          isOpen={isPstModalOpen}
          onClose={() => setIsPstModalOpen(false)}
          onImportComplete={handlePstImportComplete}
          localDbAvailable={localDbAvailable}
          localTotalEmails={localTotalEmails}
          localPstPaths={outlookPcProfile?.pstPaths || []}
          onLocalSyncComplete={() => void hydrateLocalBootstrap()}
        />
      </Suspense>
    );
  }

  // --- Full Desktop 100% Authentic Microsoft New Outlook (Fluent 2) 4-Pane Shell ---
  return (
    <div className="h-screen flex flex-col bg-[#F5F5F5] text-[#242424] overflow-hidden">
      {/* Authentic Microsoft New Outlook Single Brand Suite Bar & Global Search (48px) */}
      <header
        style={{ backgroundColor: accent.primary }}
        className="h-12 flex items-center justify-between gap-3 px-1.5 text-white shrink-0 relative z-30 transition-colors duration-150 select-none"
      >
        {/* Left: Authentic Microsoft 365 9-dot Waffle Launcher + Outlook Brand Lockup */}
        <div className="flex items-center gap-0.5 shrink-0">
          <div className="relative">
            <button
              type="button"
              onClick={() => setIsAppLauncherOpen((v) => !v)}
              aria-label="Bmail 앱 시작 관리자"
              aria-expanded={isAppLauncherOpen}
              title="앱 시작 관리자"
              className={`w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
                isAppLauncherOpen
                  ? 'bg-white/25'
                  : 'hover:bg-white/15 active:bg-white/25'
              }`}
            >
              <svg
                viewBox="0 0 16 16"
                fill="currentColor"
                aria-hidden="true"
                className="w-[16px] h-[16px] text-white"
              >
                <circle cx="2.5" cy="2.5" r="1.35" />
                <circle cx="8" cy="2.5" r="1.35" />
                <circle cx="13.5" cy="2.5" r="1.35" />
                <circle cx="2.5" cy="8" r="1.35" />
                <circle cx="8" cy="8" r="1.35" />
                <circle cx="13.5" cy="8" r="1.35" />
                <circle cx="2.5" cy="13.5" r="1.35" />
                <circle cx="8" cy="13.5" r="1.35" />
                <circle cx="13.5" cy="13.5" r="1.35" />
              </svg>
            </button>

            {isAppLauncherOpen && (
              <div
                role="menu"
                aria-label="Bmail 통합 앱 메뉴"
                className="absolute left-0 top-[calc(100%+6px)] w-64 bg-white text-[#242424] rounded-[8px] border border-[#E0E0E0] shadow-[0_8px_16px_rgba(0,0,0,0.14),0_0_2px_rgba(0,0,0,0.12)] p-1.5 space-y-0.5 z-50"
              >
                <div className="px-2.5 py-1.5 text-[11px] font-semibold text-[#616161]">
                  Bmail 통합 워크스페이스
                </div>
                {studioMenuGroups.flatMap((g) => g.items).map((item) => (
                  <button
                    key={item.id}
                    type="button"
                    role="menuitem"
                    onClick={() => {
                      item.onClick();
                      setIsAppLauncherOpen(false);
                    }}
                    className={`w-full flex items-center gap-2.5 px-2.5 py-2 text-xs rounded-[4px] text-left transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                      item.active
                        ? 'bg-[#EBF3FC] text-[#0F6CBD] font-semibold'
                        : 'text-[#242424] hover:bg-black/[0.04] active:bg-black/[0.07]'
                    }`}
                  >
                    <span className="text-[#0F6CBD] shrink-0">{item.icon}</span>
                    <span className="truncate">{item.label}</span>
                  </button>
                ))}
                <div className="my-1 border-t border-[#EDEBE9]" />
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setIsOutlookSettingsOpen(true);
                    setIsAppLauncherOpen(false);
                  }}
                  className="w-full flex items-center gap-2.5 px-2.5 py-2 text-xs font-medium text-[#242424] rounded-[4px] hover:bg-black/[0.04] active:bg-black/[0.07] text-left transition-colors duration-150"
                >
                  <Settings className="w-4 h-4 text-[#0F6CBD] shrink-0" strokeWidth={1.75} />
                  Outlook 개인 설정 (테마 · 규칙 · 서명)
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setPage('config');
                    setIsAppLauncherOpen(false);
                  }}
                  className="w-full flex items-center gap-2.5 px-2.5 py-2 text-xs font-medium text-[#242424] rounded-[4px] hover:bg-black/[0.04] active:bg-black/[0.07] text-left transition-colors duration-150"
                >
                  <Shield className="w-4 h-4 text-[#0F6CBD] shrink-0" strokeWidth={1.75} />
                  EMV 볼트 · API 키 · 서버 설정
                </button>
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setPage('install');
                    setIsAppLauncherOpen(false);
                  }}
                  className="w-full flex items-center gap-2.5 px-2.5 py-2 text-xs font-medium text-[var(--fluent-fg-1)] rounded-[4px] hover:bg-black/[0.04] active:bg-black/[0.07] text-left transition-colors duration-150"
                >
                  <Puzzle className="w-4 h-4 text-[var(--fluent-brand)] shrink-0" strokeWidth={1.75} />
                  Outlook 플러그인 · Windows 설치
                </button>
                <div className="my-1 border-t border-[var(--fluent-stroke-subtle)]" />
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => {
                    setIsManualModalOpen(true);
                    setIsAppLauncherOpen(false);
                  }}
                  className="w-full flex items-center gap-2.5 px-2.5 py-2 text-xs font-semibold text-[var(--fluent-brand)] rounded-[4px] hover:bg-[var(--fluent-brand-tint)] text-left transition-colors duration-150"
                >
                  <BookOpen className="w-4 h-4 shrink-0" strokeWidth={1.75} />
                  인터랙티브 HTML 매뉴얼 (모션그래픽 가이드)
                </button>
                <a
                  href="https://bmail.seastar.work"
                  target="_blank"
                  rel="noreferrer"
                  role="menuitem"
                  onClick={() => setIsAppLauncherOpen(false)}
                  className="w-full flex items-center gap-2.5 px-2.5 py-2 text-xs font-medium text-[var(--fluent-fg-1)] rounded-[4px] hover:bg-black/[0.04] text-left transition-colors duration-150"
                >
                  <Download className="w-4 h-4 text-[var(--fluent-brand)] shrink-0" strokeWidth={1.75} />
                  Bmail 웹 설치 · 20페이지 소개 (bmail.seastar.work)
                </a>
              </div>
            )}
          </div>

          <a
            href="/"
            onClick={(e) => {
              e.preventDefault();
              setPage('workbench');
              setIsAppLauncherOpen(false);
            }}
            className="h-9 px-2 inline-flex items-center gap-2 rounded-[4px] hover:bg-white/12 active:bg-white/20 transition-colors duration-150 whitespace-nowrap shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80"
          >
            <img
              src="/favicon.svg"
              alt="Bmail 앱 아이콘"
              className="w-6 h-6 rounded-[6px] shadow-sm shrink-0"
            />
            <span className="text-[16px] font-semibold tracking-[-0.015em] text-white leading-none">
              Bmail
            </span>
            <span className="hidden sm:inline text-[11px] font-normal text-white/75 leading-none pt-0.5">
              Outlook
            </span>
          </a>
        </div>

        {/* Center: Unified New Outlook Fluent 2 Search Bar with Docked Filter Trigger */}
        <div className="flex-1 max-w-[580px] mx-2">
          <div className="w-full relative flex items-center bg-white/[0.94] hover:bg-white focus-within:bg-white focus-within:shadow-[0_2px_8px_rgba(0,0,0,0.16)] focus-within:ring-1 focus-within:ring-white rounded-[4px] h-8 pl-2.5 pr-1 transition-all duration-150">
            <Search
              className="w-4 h-4 text-[#0F6CBD] shrink-0"
              strokeWidth={1.75}
              aria-hidden="true"
            />
            <label htmlFor="outlook-global-search" className="sr-only">
              Bmail 전체 검색
            </label>
            <input
              id="outlook-global-search"
              type="search"
              value={globalSearchQuery}
              onChange={(e) => {
                setGlobalSearchQuery(e.target.value);
                if (page !== 'workbench') setPage('workbench');
              }}
              placeholder="검색 (제목, 보낸 사람, 도면·프로젝트, 본문)"
              className="w-full bg-transparent px-2 text-[13px] text-[#242424] placeholder:text-[#616161] focus:outline-none"
            />
            {globalSearchQuery && (
              <button
                type="button"
                onClick={() => setGlobalSearchQuery('')}
                aria-label="검색어 지우기"
                className="w-6 h-6 flex items-center justify-center rounded-[4px] text-[#616161] hover:bg-black/[0.06] hover:text-[#242424] transition-colors duration-150 shrink-0"
              >
                <X className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
            )}
            <span className="h-4 w-px bg-[#E0E0E0] mx-1 shrink-0" aria-hidden="true" />
            <GlobalSearchFilterDropdown
              integrated
              filters={globalSearchFilters}
              onChange={(next) => {
                setGlobalSearchFilters(next);
                if (page !== 'workbench') setPage('workbench');
              }}
              onApply={() => {
                if (page !== 'workbench') setPage('workbench');
              }}
              referenceDate={today}
            />
          </div>
        </div>

        {/* Right: Send/Receive + Companion Mode + Interactive Motion Manual + Executive Dashboard + Uniform 36px Fluent 2 Suite Utility Buttons */}
        <div className="flex items-center gap-1 shrink-0 pr-0.5">
          <button
            type="button"
            data-testid="btn-send-receive-refresh"
            disabled={isSyncingLiveOutlook}
            onClick={() => void handleSyncIncrementalOutlook(false, { sendAndReceive: true })}
            title="클래식 Outlook 보내기/받기를 실행하고 새로 도착한 메일을 즉시 가져옵니다"
            aria-label="Outlook 보내기/받기 새로고침"
            className="inline-flex items-center gap-1.5 h-8 px-2.5 rounded-[4px] border bg-white/15 hover:bg-white/25 active:bg-white/30 border-white/30 text-white text-xs font-semibold transition-colors duration-150 whitespace-nowrap disabled:opacity-65 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80"
          >
            <RefreshCw
              className={`w-3.5 h-3.5 shrink-0 ${isSyncingLiveOutlook ? 'animate-spin' : ''}`}
              strokeWidth={2}
            />
            <span>{isSyncingLiveOutlook ? '새 메일 확인 중...' : '보내기/받기'}</span>
          </button>

          <button
            type="button"
            data-testid="btn-companion-mode"
            onClick={() => setPage('companion')}
            title="오리지널 클래식 Outlook 옆에 붙어서 실시간 메일 분석 · 발송 전 검증 · 초안 열기를 수행합니다"
            aria-label="Outlook 동반 검증 모드 열기"
            className={`inline-flex items-center gap-1.5 h-8 px-2.5 rounded-[4px] border text-white text-xs font-semibold transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              page === 'companion'
                ? 'bg-emerald-500/80 border-white/60 shadow-sm'
                : 'bg-white/12 hover:bg-white/22 border-white/25'
            }`}
          >
            <ShieldCheck className="w-3.5 h-3.5 shrink-0" strokeWidth={1.75} />
            <span>Outlook 동반 검증</span>
          </button>

          <button
            type="button"
            onClick={() => setPage('dashboard')}
            title="주간·월간 종합 분석글 · 내게 온 미회신건 · 프로젝트/주제별 통합 대시보드"
            aria-label="종합 분석 대시보드 열기"
            className={`hidden md:inline-flex items-center gap-1.5 h-8 px-2.5 rounded-[4px] border text-white text-xs font-semibold transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              page === 'dashboard'
                ? 'bg-white/25 border-white/40'
                : 'bg-white/12 hover:bg-white/22 border-white/20'
            }`}
          >
            <BarChart3 className="w-3.5 h-3.5 shrink-0" strokeWidth={1.75} />
            <span>분석 대시보드</span>
          </button>

          <button
            type="button"
            onClick={() => setIsManualModalOpen(true)}
            title="실제 화면 1:1 인터랙티브 HTML 매뉴얼 · 모션그래픽 동영상 가이드 (v1.3.1)"
            aria-label="인터랙티브 매뉴얼 팝업 열기"
            className="inline-flex items-center gap-1.5 h-8 px-2.5 rounded-[4px] bg-white/15 hover:bg-white/25 active:bg-white/30 border border-white/25 text-white text-xs font-semibold transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80"
          >
            <BookOpen className="w-3.5 h-3.5 shrink-0" strokeWidth={1.75} />
            <span>매뉴얼</span>
          </button>

          <button
            type="button"
            onClick={() => setIsPop3ModalOpen(true)}
            title={
              isSyncingPop3
                ? `${pop3Account?.email || '회사 메일'} 백그라운드 동기화 중...`
                : pop3Account
                  ? `${(pop3Account.protocol || 'pop3').toUpperCase()} 연결됨 (${pop3Account.email})`
                  : '메일플러그 · 회사 POP3/IMAP 자동 연결'
            }
            aria-label={isSyncingPop3 ? '회사 메일 동기화 중' : '회사 POP3 메일 연결'}
            aria-busy={isSyncingPop3}
            className={`relative w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              pop3Account || isSyncingPop3
                ? 'bg-white/22'
                : 'hover:bg-white/15 active:bg-white/25'
            }`}
          >
            {isSyncingPop3 ? (
              <Loader2 className="w-[17px] h-[17px] text-white animate-spin" strokeWidth={2} />
            ) : (
              <Plug className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
            )}
            {pop3Account && (
              <span
                aria-hidden="true"
                className="absolute top-1.5 right-1.5 flex h-2 w-2 items-center justify-center"
              >
                {isSyncingPop3 && (
                  <span className="absolute inline-flex h-full w-full rounded-full bg-emerald-300 opacity-80 animate-ping" />
                )}
                <span className="relative inline-flex rounded-full h-1.5 w-1.5 bg-emerald-300" />
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={() => setUiMode('mobile')}
            title="모바일 모드 (아이콘 · 메일 읽기/쓰기 기본 기능)"
            aria-label="모바일 모드로 전환"
            className="w-9 h-9 flex items-center justify-center rounded-[4px] hover:bg-white/15 active:bg-white/25 transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80"
          >
            <Smartphone className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
          </button>

          <button
            type="button"
            onClick={() => setIsSgwChatDrawerOpen((v) => !v)}
            title="SGW 실시간 채팅 서랍 열기/닫기"
            aria-label="SGW 실시간 채팅"
            className={`relative w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              isSgwChatDrawerOpen || page === 'sgw-chat'
                ? 'bg-white/22'
                : 'hover:bg-white/15 active:bg-white/25'
            }`}
          >
            <MessageSquare className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
            {sgwUnreadCount > 0 && (
              <span className="absolute top-1 right-1 min-w-[14px] h-[14px] px-1 rounded-full bg-rose-500 text-white text-[9px] font-bold flex items-center justify-center leading-none">
                {sgwUnreadCount > 99 ? '99+' : sgwUnreadCount}
              </span>
            )}
          </button>

          <button
            type="button"
            onClick={() => setIsStickyNotesOpen((v) => !v)}
            title="스티커 메모 (Sticky Notes 플로팅 위젯)"
            aria-label="스티커 메모 열기/닫기"
            className={`w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              isStickyNotesOpen ? 'bg-white/22' : 'hover:bg-white/15 active:bg-white/25'
            }`}
          >
            <StickyNote className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
          </button>

          <button
            type="button"
            onClick={() => setPage('config')}
            title="EMV 볼트 · API 키 · 서버 설정 (관리자/사원)"
            aria-label="EMV 및 서버 설정"
            className={`w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              page === 'config' ? 'bg-white/22' : 'hover:bg-white/15 active:bg-white/25'
            }`}
          >
            <Shield className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
          </button>

          <button
            type="button"
            onClick={() => setIsOutlookSettingsOpen(true)}
            title="Outlook 설정 (레이아웃 · 테마 · 규칙 · 서명 · 자동 회신 · SEDMS)"
            aria-label="Outlook 설정 열기"
            className={`w-9 h-9 flex items-center justify-center rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white/80 ${
              isOutlookSettingsOpen ? 'bg-white/22' : 'hover:bg-white/15 active:bg-white/25'
            }`}
          >
            <Settings className="w-[17px] h-[17px] text-white" strokeWidth={1.75} />
          </button>

          {/* Real-time Connected Account & Background Sync Heartbeat Indicator */}
          <div
            className="ml-1.5 flex items-center gap-1 shrink-0"
            data-testid="header-account-sync-indicator"
            data-sync-mode={headerSyncView.mode}
            aria-live="polite"
          >
            {headerSyncView.isConnected && (
              <button
                type="button"
                onClick={handleHeaderQuickSync}
                disabled={headerSyncView.isSyncing}
                title={
                  headerSyncView.isSyncing
                    ? `${headerSyncView.activeSourcesLabel} 동기화 진행 중...`
                    : `${headerSyncView.activeSourcesLabel} 즉시 새로고침`
                }
                aria-label={
                  headerSyncView.isSyncing
                    ? `${headerSyncView.activeSourcesLabel} 동기화 진행 중`
                    : '연결된 계정 즉시 동기화'
                }
                className="w-7 h-7 flex items-center justify-center rounded-full bg-white/12 hover:bg-white/22 disabled:opacity-80 text-white transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white"
              >
                <RefreshCw
                  className={`w-3.5 h-3.5 ${headerSyncView.isSyncing ? 'animate-spin' : ''}`}
                  strokeWidth={2}
                />
              </button>
            )}

            <button
              type="button"
              onClick={
                microsoftAuth.configured && !microsoftAuth.connected
                  ? handleConnectMicrosoft
                  : () => setPage('config')
              }
              title={headerSyncView.tooltip}
              aria-label={headerSyncView.ariaLabel}
              aria-busy={headerSyncView.isSyncing}
              className={`relative flex items-center justify-center shrink-0 transition-all duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-white ${
                headerSyncView.mode === 'syncing'
                  ? 'h-8 pl-2 pr-1 gap-1.5 rounded-full bg-white/20 ring-1 ring-white/75 text-white'
                  : headerSyncView.mode === 'connected'
                    ? 'h-8 pl-2 pr-1 gap-1.5 rounded-full bg-black/22 hover:bg-black/32 ring-1 ring-emerald-300/70 text-white'
                    : 'w-7 h-7 rounded-full ring-1 ring-white/55 hover:ring-white bg-black/20 hover:bg-black/30 active:bg-black/40 text-[11px] font-semibold tracking-tight'
              }`}
            >
              {headerSyncView.mode === 'syncing' && (
                <>
                  <Loader2
                    aria-hidden="true"
                    className="w-3.5 h-3.5 text-sky-100 animate-spin shrink-0"
                    strokeWidth={2.25}
                  />
                  <span className="hidden sm:inline text-[11px] font-semibold tracking-tight whitespace-nowrap">
                    {headerSyncView.compactStatusLabel}
                  </span>
                  <span className="relative w-6 h-6 rounded-full bg-black/25 text-[10px] font-semibold flex items-center justify-center shrink-0">
                    {activeEmployeeEmail ? activeEmployeeEmail.slice(0, 2).toUpperCase() : userInitials}
                    <span
                      aria-hidden="true"
                      className="absolute -bottom-0.5 -right-0.5 flex h-2.5 w-2.5 items-center justify-center"
                    >
                      <span className="absolute inline-flex h-full w-full rounded-full bg-sky-200 opacity-85 animate-ping" />
                      <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-300 ring-1 ring-[#0F6CBD]" />
                    </span>
                  </span>
                </>
              )}

              {headerSyncView.mode === 'connected' && (
                <>
                  <span
                    aria-hidden="true"
                    className="relative flex h-2 w-2 items-center justify-center shrink-0"
                  >
                    <span className="absolute inline-flex h-full w-full rounded-full bg-emerald-300 opacity-55 animate-ping" />
                    <span className="relative inline-flex rounded-full h-2 w-2 bg-emerald-300" />
                  </span>
                  <span className="hidden xl:inline text-[11px] font-medium text-white/95 whitespace-nowrap">
                    {headerSyncView.compactStatusLabel}
                  </span>
                  <span className="w-6 h-6 rounded-full bg-black/25 text-[10px] font-semibold flex items-center justify-center shrink-0">
                    {activeEmployeeEmail ? activeEmployeeEmail.slice(0, 2).toUpperCase() : userInitials}
                  </span>
                </>
              )}

              {headerSyncView.mode === 'disconnected' && (
                <span>
                  {activeEmployeeEmail ? activeEmployeeEmail.slice(0, 2).toUpperCase() : userInitials}
                </span>
              )}
            </button>
          </div>
        </div>
      </header>

      {/* Lower Body: Authentic New Outlook Fluent 2 Left App Rail (48px) + Active Workspace */}
      <div className="flex-1 flex min-h-0 overflow-hidden">
        <nav
          aria-label="Outlook 모듈 전환 레일"
          className="w-12 shrink-0 bg-[#F0F0F0] border-r border-[#E0E0E0] flex flex-col items-center justify-between py-1.5 px-1 select-none"
        >
          <div className="flex flex-col items-center gap-0.5 w-full overflow-y-auto">
            {railItems.map((item) => (
              <button
                key={item.id}
                type="button"
                onClick={item.onClick}
                title={item.label}
                aria-label={item.label}
                aria-current={item.active ? 'page' : undefined}
                style={item.active ? { color: accent.primary } : undefined}
                className={`relative w-10 h-10 rounded-[4px] flex items-center justify-center transition-all duration-150 shrink-0 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/70 ${
                  item.active
                    ? 'bg-white/95 shadow-[0_1px_2px_rgba(0,0,0,0.06)]'
                    : 'text-[#424242] hover:bg-black/[0.05] hover:text-[#242424] active:bg-black/[0.08]'
                }`}
              >
                {item.active && (
                  <span
                    aria-hidden="true"
                    style={{ backgroundColor: accent.primary }}
                    className="absolute left-0 top-1/2 -translate-y-1/2 w-[3px] h-[18px] rounded-full"
                  />
                )}
                {item.icon}
                {item.badge !== undefined && outlookPrefs.general.showUnreadBadge && (
                  <span
                    style={{ backgroundColor: accent.primary }}
                    className="absolute top-1 right-1 min-w-[15px] h-[15px] px-1 rounded-full text-[9px] font-mono tabular-nums font-semibold text-white ring-1 ring-[#F0F0F0] flex items-center justify-center leading-none"
                  >
                    {item.badge > 99 ? '99+' : item.badge}
                  </span>
                )}
              </button>
            ))}
          </div>

          <div className="flex flex-col items-center gap-0.5 w-full pt-1.5 shrink-0">
            <div className="w-6 h-px bg-[#D1D1D1] mb-1" aria-hidden="true" />
            <button
              type="button"
              onClick={() => setIsPstModalOpen(true)}
              title="원본 .pst / .eml 파일 불러오기"
              aria-label="원본 PST 파일 불러오기"
              className="w-10 h-9 rounded-[4px] flex items-center justify-center text-[#424242] hover:bg-black/[0.05] hover:text-[#0F6CBD] active:bg-black/[0.08] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/70"
            >
              <FolderOpen className="w-[17px] h-[17px]" strokeWidth={1.65} />
            </button>
          </div>
        </nav>

        <main className="flex-1 flex flex-col min-w-0 min-h-0 overflow-hidden">
          {desktopUpdateInfo?.updateAvailable && (
            <div className="px-4 py-1.5 bg-[#0F6CBD] text-white text-xs flex flex-wrap items-center justify-between gap-2 shrink-0">
              <div className="flex items-center gap-2 min-w-0">
                <ArrowUpCircle className="w-4 h-4 shrink-0" />
                <span className="font-semibold truncate">
                  Bmail v{desktopUpdateInfo.latestVersion} 업데이트 가능 (현재 v
                  {desktopUpdateInfo.currentVersion})
                </span>
                {desktopUpdateProgress?.phase === 'downloading' && (
                  <span className="font-mono text-[11px] text-white/90">
                    다운로드 {desktopUpdateProgress.percent}%
                  </span>
                )}
                {desktopUpdateProgress?.phase === 'installing' && (
                  <span className="font-mono text-[11px] text-white/90">재설치 시작 중...</span>
                )}
                {desktopUpdateProgress?.phase === 'error' && desktopUpdateProgress.message && (
                  <span className="text-[11px] text-amber-200">{desktopUpdateProgress.message}</span>
                )}
              </div>
              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  disabled={
                    desktopUpdateProgress?.phase === 'downloading' ||
                    desktopUpdateProgress?.phase === 'installing'
                  }
                  onClick={() => {
                    const installer = (
                      window as unknown as {
                        bmailDesktop?: { installUpdate?: () => Promise<unknown> };
                      }
                    ).bmailDesktop?.installUpdate;
                    if (installer) void installer();
                    else if (desktopUpdateInfo.downloadUrl) {
                      window.open(desktopUpdateInfo.downloadUrl, '_blank', 'noopener,noreferrer');
                    }
                  }}
                  className="px-2.5 py-0.5 text-[11px] font-semibold bg-white text-[#0F6CBD] hover:bg-[#EBF3FC] disabled:opacity-60 rounded"
                >
                  지금 1클릭 업데이트
                </button>
                <button
                  type="button"
                  onClick={() => setDesktopUpdateInfo(null)}
                  className="text-[11px] text-white/80 hover:text-white underline"
                >
                  나중에
                </button>
              </div>
            </div>
          )}
          {localMigrationProgress?.status === 'running' && (
            <div className="px-4 py-1.5 bg-[#EBF3FC] border-b border-[#CFE4FA] text-xs text-[#0F6CBD] flex items-center justify-between shrink-0">
              <span>
                Outlook PST 자동 색인 진행 중: 메일{' '}
                {localMigrationProgress.totalEmailsIndexed.toLocaleString()}건 · 폴더{' '}
                {localMigrationProgress.totalFoldersIndexed}개 · 연락처{' '}
                {localMigrationProgress.totalContactsIndexed}명
              </span>
              <span className="text-[11px] font-mono">백그라운드 동기화</span>
            </div>
          )}
          {pstAttachmentNotice && (
            <div className="px-4 py-1.5 bg-amber-50 border-b border-amber-200 text-xs text-amber-800 flex items-center justify-between shrink-0">
              <span>{pstAttachmentNotice}</span>
              <button
                type="button"
                onClick={() => setPstAttachmentNotice(null)}
                className="text-[11px] underline hover:no-underline"
              >
                닫기
              </button>
            </div>
          )}
          {page === 'workbench' && (
            <OutlookWorkbench
              emails={emails}
              deletedIds={deletedIds}
              drafts={drafts}
              options={options}
              providers={providers}
              analysisResult={analysisResult}
              isAnalyzing={isAnalyzing}
              agentSteps={agentSteps}
              analysisError={
                analysisError && /failed to fetch/i.test(analysisError)
                  ? null
                  : analysisError
              }
              hostInfo={hostInfo}
              microsoftAuth={microsoftAuth}
              isSyncingGraph={isSyncingGraph}
              graphNotice={
                graphNotice && /failed to fetch/i.test(graphNotice)
                  ? null
                  : graphNotice
              }
              connectedStoreName={connectedStoreName}
              pstFolders={pstFolders}
              searchQuery={globalSearchQuery}
              onSearchQueryChange={setGlobalSearchQuery}
              searchFilters={globalSearchFilters}
              onSearchFiltersChange={setGlobalSearchFilters}
              onChangeOptions={handleChangeOptions}
              onToggleSelectEmail={handleToggleSelectEmail}
              onToggleReadEmail={handleToggleReadEmail}
              onToggleFlagEmail={handleToggleFlagEmail}
              onDeleteEmail={handleDeleteEmail}
              onRestoreEmail={handleRestoreEmail}
              onEmptyTrash={handleEmptyTrash}
              onSaveDraft={handleSaveDraft}
              onDeleteDraft={handleDeleteDraft}
              onRunAnalysis={handleRunAnalysis}
              onConnectMicrosoft={handleConnectMicrosoft}
              onSyncMicrosoftEmails={handleSyncMicrosoftEmails}
              onDisconnectMicrosoft={handleDisconnectMicrosoft}
              onOpenPstModal={() => setIsPstModalOpen(true)}
              pop3ConnectedEmail={pop3Account?.email || null}
              onOpenPop3Modal={() => setIsPop3ModalOpen(true)}
              onOpenInstallGuide={() => setPage('install')}
              onOpenConfigPage={() => setPage('config')}
              onOpenArchivePage={() => setPage('archive')}
              onOpenStudioPage={() => {
                setStudioPreset({ tab: 'documents' });
                setPage('studio');
              }}
              onOpenVoicePage={() => {
                setStudioPreset({ tab: 'meeting' });
                setPage('studio');
              }}
              onOpenAiBriefPage={() => setPage('ai-inbox')}
              onOpenSgwSedmsPage={() => setPage('sgw-sedms')}
              onOpenProjectsPage={() => setPage('projects')}
              onMoveEmailToFolder={handleMoveEmailToFolder}
              projects={projects}
              requestedActiveEmailId={requestedActiveEmailId}
              onClearRequestedActiveEmailId={() => setRequestedActiveEmailId(null)}
              isTaskpaneOpen={isTaskpaneOpen}
              isTaskpanePinned={isTaskpanePinned}
              supportsPinning={appConfig.plugin.supportsPinning}
              onRibbonIconClick={handleRibbonIconClick}
              onToggleTaskpanePin={handleToggleTaskpanePin}
              onCloseTaskpane={handleCloseTaskpane}
              outlookPrefs={outlookPrefs}
              onChangeOutlookPrefs={handleChangeOutlookPrefs}
              onOpenOutlookSettings={() => setIsOutlookSettingsOpen(true)}
              pinnedEmailIds={pinnedEmailIds}
              onTogglePinEmail={handleTogglePinEmail}
              favoriteFolderIds={favoriteFolderIds}
              onToggleFavoriteFolder={handleToggleFavoriteFolder}
              projectSubfolders={projectSubfolders}
              projectVaultFiles={projectVaultFiles}
              projectMailOverrides={projectMailOverrides}
              onAddProjectSubfolder={handleAddProjectSubfolder}
              onDeleteProjectSubfolder={handleDeleteProjectSubfolder}
              onAddProjectVaultFiles={handleAddProjectVaultFiles}
              onRemoveProjectVaultFile={handleRemoveProjectVaultFile}
              onAssignEmailToProject={handleAssignEmailToProject}
              onRemoveEmailFromProject={handleRemoveEmailFromProject}
              onSendViaSmtp={handleSendViaPop3Smtp}
              localDbAvailable={localDbAvailable}
              localTotalEmails={localTotalEmails}
              onActiveFolderChange={setActiveWorkbenchFolder}
              requestedMailto={requestedMailto}
              onAnalyzePstAttachment={(emailId, idx) => void handleAnalyzePstAttachment(emailId, idx)}
              onShareMailToSgwChat={(mail) => {
                setPendingShareMailToChat(mail);
                setIsSgwChatDrawerOpen(true);
              }}
              onClearRequestedMailto={() => setRequestedMailto(null)}
              localFolders={localFolders}
              outlookFavorites={outlookFavorites}
              isSyncingLiveOutlook={isSyncingLiveOutlook}
              onSyncLiveOutlook={() =>
                void handleSyncIncrementalOutlook(false, { sendAndReceive: true })
              }
              onOpenManual={() => setIsManualModalOpen(true)}
              onOpenExecutiveDashboard={() => setPage('dashboard')}
              onAddCalendarEvent={handleAddManualEvent}
              onImportDroppedEmails={handleImportDroppedEmails}
              requestedToastReplyEmailId={requestedToastReplyEmailId}
              onClearRequestedToastReplyEmailId={() => setRequestedToastReplyEmailId(null)}
              mailCategoriesMap={mailCategoriesMap}
              onChangeMailCategoriesMap={handleChangeMailCategoriesMap}
              mailSnoozeMap={mailSnoozeMap}
              onChangeMailSnoozeMap={handleChangeMailSnoozeMap}
              onArchiveEmail={handleArchiveEmail}
              unreadBriefing={unreadBriefing}
              onDiscoverOutlookSettings={() => void handleDiscoverOutlookSettings()}
              isDiscoveringOutlookSettings={isDiscoveringOutlookSettings}
              localSignatures={localSignatures}
              localContacts={outlookPcProfile?.contacts || []}
            />
          )}

          {page !== 'workbench' && (
            <LazyViewErrorBoundary viewKey={page} onRetry={() => setPage((prev) => prev)}>
              <Suspense fallback={<div className="flex-1 flex items-center justify-center text-xs text-[#616161]">화면을 불러오는 중...</div>}>
              {page === 'ai-inbox' && (
                <AiInboxIntelligenceView
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  projects={projects}
                  calendarEvents={allCalendarEvents}
                  onOpenPstModal={() => setIsPstModalOpen(true)}
                  onSaveDraft={(draft) => {
                    handleSaveDraft(draft);
                    setGraphNotice('임시 보관함에 AI 회신·작성 초안을 저장했습니다.');
                  }}
                  onAddCalendarEvent={(ev) => {
                    handleAddManualEvent(ev);
                    setGraphNotice(`"${ev.title}" 일정을 Outlook 달력에 추가했습니다.`);
                  }}
                  onDeleteEmail={handleDeleteEmail}
                  onOpenMailInWorkbench={(emailId) => {
                    void handleOpenMailById(emailId);
                  }}
                  onOpenExecutiveDashboard={() => setPage('dashboard')}
                />
              )}

              {page === 'sgw-chat' && (
                <SgwChatPanel
                  mode="full"
                  defaultEmail={
                    localDefaultAccount?.email ||
                    microsoftAuth.userEmail ||
                    hostInfo.userEmail ||
                    pop3Account?.email ||
                    ''
                  }
                  pendingShareMail={pendingShareMailToChat}
                  onClearPendingShare={() => setPendingShareMailToChat(null)}
                  initialPeerEmail={sgwChatInitialPeer}
                  onClearInitialPeer={() => setSgwChatInitialPeer(null)}
                  onComposeFromChat={(draft) => {
                    setRequestedMailto({ to: draft.to, cc: '', subject: draft.subject, body: draft.body });
                    setPage('workbench');
                  }}
                  onOpenMailById={(id) => void handleOpenMailById(id)}
                  onUnreadCountChange={handleSgwUnreadChange}
                />
              )}

              {page === 'sgw-sedms' && (
                <SgwSedmsHubView
                  initialTab="sgw-mail"
                  onOpenSgwChat={() => setIsSgwChatDrawerOpen(true)}
                  localTotalEmails={localTotalEmails}
                />
              )}

              {page === 'companion' && (
                <OutlookCompanionPanel
                  fallbackEmail={emails.find((m) => !deletedIds.includes(m.id)) ?? null}
                  onSwitchToWorkbench={() => setPage('workbench')}
                  onSwitchToDashboard={() => setPage('dashboard')}
                  onSendReceiveRefresh={() =>
                    void handleSyncIncrementalOutlook(false, { sendAndReceive: true })
                  }
                  onOpenAiSettings={() => setPage('config')}
                  isSyncingLiveOutlook={isSyncingLiveOutlook}
                />
              )}

              {page === 'dashboard' && (
                <DashboardView
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  documents={parsedDocuments}
                  projects={projects}
                  analysisResult={analysisResult}
                  archivedResults={archive.map((entry) => entry.result)}
                  isAnalyzing={isAnalyzing}
                  onRunAiAnalysis={() => {
                    setPage('workbench');
                    void handleRunAnalysis();
                  }}
                  onOpenPstModal={() => setIsPstModalOpen(true)}
                  onOpenStudioDocuments={() => {
                    setStudioPreset({ tab: 'documents' });
                    setPage('studio');
                  }}
                  onOpenProjects={() => setPage('projects')}
                  onOpenMailInWorkbench={(emailId) => {
                    void handleOpenMailById(emailId);
                  }}
                  onSyncIncrementalOutlook={() =>
                    handleSyncIncrementalOutlook(false, { sendAndReceive: true })
                  }
                  onSaveDraftToOutlook={(subject, body) => {
                    handleSaveDraft({ to: '', cc: '', subject, body, kind: 'worklog' });
                    setGraphNotice('Outlook 임시 보관함에 일간 업무 브리핑 초안을 저장했습니다.');
                    setPage('workbench');
                  }}
                />
              )}

              {page === 'calendar' && (
                <OutlookCalendarView
                  events={allCalendarEvents}
                  projects={projects}
                  onAddManualEvent={handleAddManualEvent}
                  onImportIcsEvents={handleImportIcsEvents}
                  onToggleCompleteEvent={handleToggleCompleteEvent}
                  onDeleteManualEvent={handleDeleteManualEvent}
                  onOpenMeetingAnalyzer={(title, date) => {
                    setStudioPreset({ tab: 'meeting', title, date });
                    setPage('studio');
                  }}
                />
              )}

              {page === 'people' && (
                <OutlookPeopleView
                  initialContacts={outlookPcProfile?.contacts || []}
                  onComposeTo={(email, name) => {
                    setRequestedMailto({
                      to: name ? `${name} <${email}>` : email,
                      cc: '',
                      subject: '',
                      body: '',
                    });
                    setPage('workbench');
                  }}
                  onFilterBySender={(email) => {
                    setGlobalSearchQuery(email);
                    setPage('workbench');
                  }}
                />
              )}

              {page === 'tasks' && (
                <OutlookTasksView
                  tasks={outlookTasks}
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  onChangeTasks={handleChangeOutlookTasks}
                  onOpenMail={(m) => {
                    void handleOpenMailById(m.id);
                  }}
                  onToggleMailFlag={handleToggleFlagEmail}
                />
              )}

              {page === 'projects' && (
                <ProjectWorkspaceView
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  documents={parsedDocuments}
                  projects={projects}
                  ownerAccountEmail={microsoftAuth.userEmail || hostInfo.userEmail || undefined}
                  onAddProject={handleAddProject}
                  onUpdateProject={handleUpdateProject}
                  onImportProjectRules={handleImportProjectRules}
                  onToggleFavoriteProject={handleToggleFavoriteProject}
                  onDeleteProject={handleDeleteProject}
                  onApplyDeterministicWorkLog={handleApplyDeterministicWorkLog}
                  onExportMbox={handleExportMbox}
                  onSaveDraftToOutlook={(subject, body) => {
                    handleSaveDraft({ to: '', cc: '', subject, body, kind: 'worklog' });
                    setGraphNotice('프로젝트 파일 정리 보고서 초안을 Outlook 임시 보관함에 저장했습니다.');
                    setPage('workbench');
                  }}
                />
              )}

              {page === 'studio' && (
                <IntelligenceStudio
                  key={`${studioPreset.tab}-${studioPreset.title || ''}-${studioPreset.date || ''}`}
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  providers={providers}
                  selectedProvider={options.provider}
                  maskSensitiveInfo={options.maskSensitiveInfo}
                  projects={projects}
                  parsedDocuments={parsedDocuments}
                  onAddParsedDocuments={(docs) => setParsedDocuments((prev) => [...docs, ...prev])}
                  onRemoveParsedDocument={(id) =>
                    setParsedDocuments((prev) => prev.filter((d) => d.id !== id))
                  }
                  meetingMinutesHistory={meetingMinutesHistory}
                  onAddMeetingMinutes={handleAddMeetingMinutes}
                  onSyncMinutesToCalendar={handleSyncMinutesToCalendar}
                  onCreateOutlookDraft={(subject, body) => {
                    handleSaveDraft({ to: '', cc: '', subject, body, kind: 'worklog' });
                    setGraphNotice('Outlook 임시 보관함에 보고서 초안을 저장했습니다.');
                    setPage('workbench');
                  }}
                  onOpenConfigExternalEngines={() => setPage('config')}
                  initialTab={studioPreset.tab}
                  prefillMeetingTitle={studioPreset.title}
                  prefillMeetingDate={studioPreset.date}
                />
              )}

              {page === 'install' && (
                <PluginInstallCenter
                  defaultAppUrl={appUrl}
                  oauthCallbackUrl={oauthCallbackUrl}
                  providers={providers}
                  microsoftConfigured={microsoftAuth.configured}
                  onOpenPstModal={() => setIsPstModalOpen(true)}
                />
              )}

              {page === 'config' && (
                <ConfigEnvManager
                  config={appConfig}
                  envSecrets={envSecrets}
                  oauthCallbackUrl={oauthCallbackUrl}
                  configPasswordRequired={configPasswordRequired}
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  activeEmployeeEmail={activeEmployeeEmail}
                  pop3SessionToken={pop3SessionToken}
                  outlookPcProfile={outlookPcProfile}
                  onImportPcProfile={(profile) => {
                    setOutlookPcProfile(profile);
                    if (profile.accounts[0]?.smtpAddress) {
                      setHostInfo((prev) => ({ ...prev, userEmail: profile.accounts[0].smtpAddress }));
                    }
                    setGraphNotice(
                      `PC Outlook 프로필을 반영했습니다 (계정 ${profile.accounts.length}개 · 주소록 ${profile.contacts.length}명).`
                    );
                  }}
                  onOpenPop3Modal={() => setIsPop3ModalOpen(true)}
                  onVaultUpdated={() => {
                    const headers: Record<string, string> = {};
                    if (pop3SessionToken) headers.Authorization = `Bearer ${pop3SessionToken}`;
                    void fetch('/api/config', { headers })
                      .then((r) => (r.ok ? r.json() : null))
                      .then((data) => {
                        if (!data) return;
                        if (Array.isArray(data.envSecrets)) setEnvSecrets(data.envSecrets);
                        if (Array.isArray(data.providers)) setProviders(data.providers);
                      })
                      .catch(() => {});
                  }}
                  onSaveConfig={handleSaveAppConfig}
                />
              )}

              {page === 'archive' && (
                <ArchiveView
                  entries={archive}
                  onDeleteEntry={handleDeleteArchive}
                  onGoToWorkbench={() => setPage('workbench')}
                />
              )}
              </Suspense>
            </LazyViewErrorBoundary>
          )}
        </main>
        {isSgwChatDrawerOpen && page !== 'sgw-chat' && (
          <Suspense fallback={null}>
            <SgwChatPanel
              mode="drawer"
              defaultEmail={
                localDefaultAccount?.email ||
                microsoftAuth.userEmail ||
                hostInfo.userEmail ||
                pop3Account?.email ||
                ''
              }
              pendingShareMail={pendingShareMailToChat}
              onClearPendingShare={() => setPendingShareMailToChat(null)}
              initialPeerEmail={sgwChatInitialPeer}
              onClearInitialPeer={() => setSgwChatInitialPeer(null)}
              onComposeFromChat={(draft) => {
                setRequestedMailto({ to: draft.to, cc: '', subject: draft.subject, body: draft.body });
                setPage('workbench');
              }}
              onOpenMailById={(id) => void handleOpenMailById(id)}
              onUnreadCountChange={handleSgwUnreadChange}
              onClose={() => setIsSgwChatDrawerOpen(false)}
            />
          </Suspense>
        )}
      </div>

      <Suspense fallback={null}>
        <OutlookSettingsModal
          isOpen={isOutlookSettingsOpen}
          preferences={outlookPrefs}
          projects={projects}
          onChangePreferences={handleChangeOutlookPrefs}
          onClose={() => setIsOutlookSettingsOpen(false)}
          onOpenServerConfig={() => setPage('config')}
          onOpenSgwSedms={() => setPage('sgw-sedms')}
          onDiscoverOutlookSettings={handleDiscoverOutlookSettings}
          isDiscoveringOutlookSettings={isDiscoveringOutlookSettings}
        />

        <Pop3QuickConnectModal
          isOpen={isPop3ModalOpen}
          connectedAccount={pop3Account}
          initialAccount={localDefaultAccount}
          onClose={() => setIsPop3ModalOpen(false)}
          onConnected={handlePop3Connected}
          onDisconnect={() => void handleDisconnectPop3()}
          onOpenPstFallback={() => setIsPstModalOpen(true)}
        />

        <PstImportModal
          isOpen={isPstModalOpen}
          onClose={() => setIsPstModalOpen(false)}
          onImportComplete={handlePstImportComplete}
          localDbAvailable={localDbAvailable}
          localTotalEmails={localTotalEmails}
          localPstPaths={outlookPcProfile?.pstPaths || []}
          onLocalSyncComplete={() => void hydrateLocalBootstrap()}
        />

        <InteractiveManualModal
          isOpen={isManualModalOpen}
          onClose={() => setIsManualModalOpen(false)}
        />

        <StickyNotesWidget
          open={isStickyNotesOpen}
          onClose={() => setIsStickyNotesOpen(false)}
          activeEmail={
            requestedActiveEmailId
              ? emails.find((m) => m.id === requestedActiveEmailId) || emails[0] || null
              : emails[0] || null
          }
          onOpenEmailById={(id) => void handleOpenMailById(id)}
        />
      </Suspense>

      {/* Bottom-right In-App New Mail Arrival & Concise Summary Toasts */}
      {arrivalToasts.length > 0 && (
        <div
          aria-live="polite"
          className="fixed bottom-4 right-4 z-50 w-[356px] max-w-[calc(100vw-2rem)] space-y-2"
        >
          {arrivalToasts.map((toast) => (
            <div
              key={toast.emailId}
              className={`bg-white rounded-[8px] border shadow-[0_10px_28px_rgba(0,0,0,0.16)] p-3 space-y-2 ${
                toast.urgency === 'high' ? 'border-[#D13438]' : 'border-[#0F6CBD]'
              }`}
            >
              <div className="flex items-center justify-between gap-2">
                <div className="flex items-center gap-1.5 min-w-0">
                  <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
                  <span className="text-[11px] font-bold text-[#0F6CBD] truncate">
                    새 메일 도착 · 간단 요약
                  </span>
                  {toast.urgency === 'high' && (
                    <span className="px-1.5 py-0.5 rounded bg-[#D13438] text-white text-[10px] font-bold shrink-0">
                      긴급
                    </span>
                  )}
                </div>
                <button
                  type="button"
                  aria-label="새 메일 알림 닫기"
                  onClick={() =>
                    setArrivalToasts((prev) => prev.filter((item) => item.emailId !== toast.emailId))
                  }
                  className="p-0.5 text-[#616161] hover:text-[#242424] rounded"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
              <div className="min-w-0">
                <div className="text-xs font-bold text-[#242424] truncate">{toast.sender}</div>
                <div className="text-[11px] font-semibold text-[#0F6CBD] truncate">{toast.subject}</div>
              </div>
              <div
                className={`text-[11px] text-[#323130] bg-[#FAF9F8] px-2.5 py-1.5 rounded border-l-2 leading-snug line-clamp-2 ${
                  toast.urgency === 'high' ? 'border-[#D13438]' : 'border-[#0F6CBD]'
                }`}
              >
                {toast.summary}
              </div>
              <div className="flex items-center gap-1.5 pt-0.5">
                <button
                  type="button"
                  onClick={() => {
                    setArrivalToasts((prev) => prev.filter((i) => i.emailId !== toast.emailId));
                    handleToggleReadEmail(toast.emailId);
                  }}
                  className="flex-1 inline-flex items-center justify-center gap-1 px-2 py-1 text-[11px] font-medium bg-[#FAF9F8] hover:bg-[#F0F0F0] border border-[#D1D1D1] rounded-[4px]"
                >
                  <Check className="w-3 h-3 text-[#107C10]" />
                  즉시 읽음
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setArrivalToasts((prev) => prev.filter((i) => i.emailId !== toast.emailId));
                    setPage('workbench');
                    setRequestedToastReplyEmailId(toast.emailId);
                  }}
                  className="flex-1 inline-flex items-center justify-center gap-1 px-2 py-1 text-[11px] font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[4px]"
                >
                  <Reply className="w-3 h-3" />
                  즉시 회신
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setArrivalToasts((prev) => prev.filter((i) => i.emailId !== toast.emailId));
                    void handleOpenMailById(toast.emailId);
                  }}
                  className="px-2.5 py-1 text-[11px] font-medium bg-white hover:bg-[#F5F5F5] border border-[#D1D1D1] rounded-[4px]"
                >
                  열기
                </button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
