import React, { lazy, Suspense, useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  Archive,
  ArrowDownUp,
  Bell,
  BookOpen,
  Briefcase,
  Building2,
  CheckSquare,
  ChevronDown,
  ChevronRight,
  Clock,
  CloudDownload,
  Database,
  Download,
  FileEdit,
  FileSpreadsheet,
  Flag,
  Folder,
  Forward,
  Inbox,
  Languages,
  LayoutList,
  LogOut,
  Mail,
  MailOpen,
  Menu,
  MessageSquare,
  Mic,
  PanelLeft,
  PanelRight,
  Paperclip,
  PenSquare,
  Pin,
  Plug,
  Puzzle,
  RefreshCw,
  Reply,
  ReplyAll,
  RotateCcw,
  Search,
  Send,
  Settings,
  ShieldCheck,
  Sliders,
  Sparkles,
  Square,
  Star,
  Tag,
  Trash2,
  Umbrella,
  Undo2,
  Upload,
  X,
  Zap,
} from 'lucide-react';
import type {
  AgentExecutionStep,
  AiProviderId,
  AnalysisOptions,
  AnalysisResult,
  EmailFolder,
  MailAttachmentMeta,
  MailItem,
  OutlookHostInfo,
  ProviderStatus,
} from '../types/mail';
import type { LocalFolderRecord } from '../services/localMailDb';
import type { OutlookFavoriteLink } from '../services/localOutlookMigrator';
import { exportDraftToOutlook } from '../services/officeBridge';
import {
  type CalendarScheduleEvent,
  normalizeConversationSubject,
  type ProjectDefinition,
} from '../utils/deterministicAnalyzer';
import {
  classifyEmailIntent,
  generateInstantReplyOptions,
} from '../utils/aiMailIntelligence';
import {
  countActiveGlobalSearchFilters,
  DEFAULT_GLOBAL_SEARCH_FILTERS,
  isPromotionalEmail,
  matchesGlobalSearchFilters,
  type GlobalSearchFilters,
} from '../utils/mailFilter';
import {
  ACCENT_THEME_COLORS,
  applyQuickStepActions,
  buildSweepRuleForSender,
  DEFAULT_FAVORITE_FOLDER_IDS,
  DEFAULT_OUTLOOK_PREFERENCES,
  EMPTY_PROJECT_MAIL_OVERRIDES,
  evaluateOutlookMailRules,
  extractShipbuildingDrawingRefs,
  filterEmailsBySearchFolder,
  isBlockedSender,
  isEmailInProject,
  isInlineBusinessCardImage,
  type OutlookDensity,
  type OutlookPreferences,
  type OutlookPreviewLines,
  type OutlookQuickStep,
  type OutlookReadingPanePosition,
  type OutlookSearchFolderId,
  type ProjectMailOverrides,
  type ProjectTreeSubfolder,
  type ProjectVaultFile,
} from '../utils/outlookWorkspaceModel';
import {
  cancelDelayedSendItem,
  createDelayedSendItem,
  formatCountdownLabel,
  getRemainingSeconds,
  type DelayedSendItem,
} from '../utils/delayedSendQueue';
import {
  classifyMailPriority,
  compareMailsByPriority,
  type MailPriorityAssessment,
  type MailPriorityLevel,
} from '../utils/mailPriority';
import type { UnreadPriorityBriefing } from '../services/unreadBriefingEngine';
import type {
  ComposeDraftState,
  OutlookContactEntry,
  OutlookSignatureEntry,
} from './OutlookComposeWindow';
import { OutlookStoreFolderTree } from './OutlookStoreFolderTree';
import { OutlookTaskpane } from './OutlookTaskpane';
import { ProjectVaultBar } from './ProjectVaultBar';
import { UnreadBriefingBanner } from './UnreadBriefingBanner';

const OutlookComposeWindow = lazy(() =>
  import('./OutlookComposeWindow').then((m) => ({ default: m.OutlookComposeWindow }))
);

const AnalysisResultPanel = lazy(() =>
  import('./AnalysisResultPanel').then((m) => ({ default: m.AnalysisResultPanel }))
);
const AttachmentPreviewDock = lazy(() =>
  import('./AttachmentPreviewDock').then((m) => ({ default: m.AttachmentPreviewDock }))
);
const MailAiDockPanel = lazy(() =>
  import('./MailAiDockPanel').then((m) => ({ default: m.MailAiDockPanel }))
);
const MailMiniAssistWindow = lazy(() =>
  import('./MailMiniAssistWindow').then((m) => ({ default: m.MailMiniAssistWindow }))
);

type ResizablePaneKey = 'folder' | 'list' | 'aiDock' | 'attachDock';

const PANE_WIDTHS_STORAGE_KEY = 'bmail_pane_widths_v1';

const DEFAULT_PANE_WIDTHS: Record<ResizablePaneKey, number> = {
  folder: 224,
  list: 360,
  aiDock: 300,
  attachDock: 316,
};

const PANE_WIDTH_LIMITS: Record<ResizablePaneKey, readonly [number, number]> = {
  folder: [150, 460],
  list: [220, 760],
  aiDock: [240, 640],
  attachDock: [240, 640],
};

function clampPaneWidth(pane: ResizablePaneKey, raw: unknown): number {
  const [minW, maxW] = PANE_WIDTH_LIMITS[pane];
  const n = typeof raw === 'number' ? raw : Number(raw);
  if (!Number.isFinite(n)) return DEFAULT_PANE_WIDTHS[pane];
  return Math.max(minW, Math.min(maxW, Math.round(n)));
}

function loadSavedPaneWidths(): Record<ResizablePaneKey, number> {
  try {
    const raw = localStorage.getItem(PANE_WIDTHS_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_PANE_WIDTHS };
    const parsed = JSON.parse(raw) as Partial<Record<ResizablePaneKey, unknown>>;
    return {
      folder: clampPaneWidth('folder', parsed?.folder),
      list: clampPaneWidth('list', parsed?.list),
      aiDock: clampPaneWidth('aiDock', parsed?.aiDock),
      attachDock: clampPaneWidth('attachDock', parsed?.attachDock),
    };
  } catch {
    return { ...DEFAULT_PANE_WIDTHS };
  }
}

interface MicrosoftAuthState {
  configured: boolean;
  connected: boolean;
  userEmail: string | null;
  displayName: string | null;
}

export interface OutlookDraftItem {
  id: string;
  to: string;
  cc: string;
  subject: string;
  body: string;
  updatedAt: string;
  kind: 'new' | 'reply' | 'replyAll' | 'forward' | 'worklog';
}

interface OutlookWorkbenchProps {
  emails: MailItem[];
  deletedIds: string[];
  drafts: OutlookDraftItem[];
  options: AnalysisOptions;
  providers: ProviderStatus[];
  analysisResult: AnalysisResult | null;
  isAnalyzing: boolean;
  agentSteps: AgentExecutionStep[];
  analysisError: string | null;
  hostInfo: OutlookHostInfo;
  microsoftAuth: MicrosoftAuthState;
  isSyncingGraph: boolean;
  graphNotice: string | null;
  connectedStoreName: string | null;
  pstFolders: string[];
  searchQuery: string;
  onSearchQueryChange: (q: string) => void;
  searchFilters?: GlobalSearchFilters;
  onSearchFiltersChange?: (next: GlobalSearchFilters) => void;
  onChangeOptions: (patch: Partial<AnalysisOptions>) => void;
  onToggleSelectEmail: (id: string) => void;
  onToggleReadEmail: (id: string) => void;
  onToggleFlagEmail: (id: string) => void;
  onDeleteEmail: (id: string) => void;
  onRestoreEmail: (id: string) => void;
  onEmptyTrash: () => void;
  onSaveDraft: (draft: Omit<OutlookDraftItem, 'id' | 'updatedAt'>) => void;
  onDeleteDraft: (id: string) => void;
  onRunAnalysis: (overrideProvider?: AiProviderId) => void;
  onConnectMicrosoft: () => void;
  onSyncMicrosoftEmails: () => void;
  onDisconnectMicrosoft: () => void;
  onOpenPstModal: () => void;
  pop3ConnectedEmail?: string | null;
  onOpenPop3Modal?: () => void;
  onOpenInstallGuide: () => void;
  onOpenConfigPage: () => void;
  onOpenArchivePage: () => void;
  onOpenStudioPage: () => void;
  onOpenVoicePage?: () => void;
  onOpenAiBriefPage?: () => void;
  onOpenSgwSedmsPage?: () => void;
  onOpenProjectsPage: () => void;
  onMoveEmailToFolder: (id: string, folder: EmailFolder) => void;
  projects: ProjectDefinition[];
  requestedActiveEmailId?: string | null;
  onClearRequestedActiveEmailId?: () => void;
  isTaskpaneOpen: boolean;
  isTaskpanePinned: boolean;
  supportsPinning: boolean;
  onRibbonIconClick: () => void;
  onToggleTaskpanePin: () => void;
  onCloseTaskpane: () => void;
  outlookPrefs?: OutlookPreferences;
  onChangeOutlookPrefs?: (next: OutlookPreferences) => void;
  onOpenOutlookSettings?: () => void;
  pinnedEmailIds?: string[];
  onTogglePinEmail?: (id: string) => void;
  favoriteFolderIds?: string[];
  onToggleFavoriteFolder?: (folderId: string) => void;
  projectSubfolders?: ProjectTreeSubfolder[];
  projectVaultFiles?: ProjectVaultFile[];
  projectMailOverrides?: ProjectMailOverrides;
  onAddProjectSubfolder?: (projectCode: string, name: string, parentId: string | null) => void;
  onDeleteProjectSubfolder?: (subfolderId: string) => void;
  onAddProjectVaultFiles?: (projectCode: string, subfolderId: string | null, fileList: File[]) => void;
  onRemoveProjectVaultFile?: (fileId: string) => void;
  onAssignEmailToProject?: (projectCode: string, emailId: string) => void;
  onRemoveEmailFromProject?: (projectCode: string, emailId: string) => void;
  onSendViaSmtp?: (payload: {
    to: string;
    cc: string;
    bcc?: string;
    subject: string;
    body: string;
    htmlBody?: string;
    importance?: 'low' | 'normal' | 'high';
    requestReadReceipt?: boolean;
    attachments?: Array<{ filename: string; contentType?: string; contentBase64: string }>;
  }) => Promise<{ ok: boolean; message: string }>;
  localDbAvailable?: boolean;
  localTotalEmails?: number;
  onActiveFolderChange?: (folder: string) => void;
  requestedMailto?: { to: string; cc: string; subject: string; body: string } | null;
  onClearRequestedMailto?: () => void;
  onAnalyzePstAttachment?: (emailId: string, idx: number) => void;
  onShareMailToSgwChat?: (mail: MailItem) => void;
  localFolders?: LocalFolderRecord[];
  outlookFavorites?: OutlookFavoriteLink[];
  isSyncingLiveOutlook?: boolean;
  onSyncLiveOutlook?: () => void;
  onOpenManual?: () => void;
  onOpenExecutiveDashboard?: () => void;
  onAddCalendarEvent?: (ev: Omit<CalendarScheduleEvent, 'id'>) => void;
  onImportDroppedEmails?: (items: MailItem[]) => void;
  requestedToastReplyEmailId?: string | null;
  onClearRequestedToastReplyEmailId?: () => void;
  mailCategoriesMap?: Record<string, string[]>;
  onChangeMailCategoriesMap?: (next: Record<string, string[]>) => void;
  mailSnoozeMap?: Record<string, string>;
  onChangeMailSnoozeMap?: (next: Record<string, string>) => void;
  onArchiveEmail?: (id: string) => void;
  unreadBriefing?: UnreadPriorityBriefing | null;
  onDiscoverOutlookSettings?: () => Promise<void> | void;
  isDiscoveringOutlookSettings?: boolean;
  localSignatures?: OutlookSignatureEntry[];
  localContacts?: OutlookContactEntry[];
}

type BuiltInFolder =
  | 'all'
  | EmailFolder
  | 'flagged'
  | 'archive'
  | 'drafts'
  | 'trash'
  | `search:${OutlookSearchFolderId}`;
type FocusedTab = 'focused' | 'other';
type ListFilterMode = 'all' | 'unread' | 'attachments' | 'urgent' | 'important' | 'normal';
type SortMode = 'priority' | 'newest' | 'oldest' | 'sender';
type RibbonTab = 'home' | 'view' | 'help';

const AVATAR_TONES = [
  'bg-[#EBF3FC] text-[#0F6CBD]',
  'bg-[#E8F5E9] text-[#1B5E20]',
  'bg-[#FFF3E0] text-[#B45309]',
  'bg-[#F3E8FF] text-[#6B21A8]',
  'bg-[#FCE7F3] text-[#9D174D]',
  'bg-[#E0F2FE] text-[#0369A1]',
];

function getAvatarTone(seed: string): string {
  let hash = 0;
  for (let i = 0; i < seed.length; i += 1) {
    hash = (hash * 31 + seed.charCodeAt(i)) >>> 0;
  }
  return AVATAR_TONES[hash % AVATAR_TONES.length];
}

function getInitials(name: string): string {
  const clean = name.trim();
  if (!clean) return '?';
  // Korean names: take up to 2 chars; Latin names: initials of words
  if (/^[가-힣]/.test(clean)) {
    return clean.slice(0, 2);
  }
  const parts = clean.split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    return `${parts[0][0]}${parts[1][0]}`.toUpperCase();
  }
  return clean.slice(0, 2).toUpperCase();
}

export const OutlookWorkbench: React.FC<OutlookWorkbenchProps> = ({
  emails,
  deletedIds,
  drafts,
  options,
  providers,
  analysisResult,
  isAnalyzing,
  agentSteps,
  analysisError,
  hostInfo,
  microsoftAuth,
  isSyncingGraph,
  graphNotice,
  connectedStoreName,
  pstFolders,
  searchQuery,
  onSearchQueryChange,
  searchFilters = DEFAULT_GLOBAL_SEARCH_FILTERS,
  onSearchFiltersChange,
  onChangeOptions,
  onToggleSelectEmail,
  onToggleReadEmail,
  onToggleFlagEmail,
  onDeleteEmail,
  onRestoreEmail,
  onEmptyTrash,
  onSaveDraft,
  onDeleteDraft,
  onRunAnalysis,
  onConnectMicrosoft,
  onSyncMicrosoftEmails,
  onDisconnectMicrosoft,
  onOpenPstModal,
  pop3ConnectedEmail,
  onOpenPop3Modal,
  onOpenInstallGuide,
  onOpenConfigPage,
  onOpenArchivePage,
  onOpenStudioPage,
  onOpenVoicePage,
  onOpenAiBriefPage,
  onOpenSgwSedmsPage,
  onOpenProjectsPage,
  onMoveEmailToFolder,
  projects,
  requestedActiveEmailId,
  onClearRequestedActiveEmailId,
  isTaskpaneOpen,
  isTaskpanePinned,
  supportsPinning,
  onRibbonIconClick,
  onToggleTaskpanePin,
  onCloseTaskpane,
  outlookPrefs = DEFAULT_OUTLOOK_PREFERENCES,
  onChangeOutlookPrefs,
  onOpenOutlookSettings,
  pinnedEmailIds = [],
  onTogglePinEmail,
  favoriteFolderIds = DEFAULT_FAVORITE_FOLDER_IDS,
  onToggleFavoriteFolder,
  projectSubfolders = [],
  projectVaultFiles = [],
  projectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES,
  onAddProjectSubfolder,
  onDeleteProjectSubfolder,
  onAddProjectVaultFiles,
  onRemoveProjectVaultFile,
  onAssignEmailToProject,
  onRemoveEmailFromProject,
  onSendViaSmtp,
  localDbAvailable = false,
  localTotalEmails = 0,
  onActiveFolderChange,
  requestedMailto,
  onClearRequestedMailto,
  onAnalyzePstAttachment,
  onShareMailToSgwChat,
  localFolders = [],
  outlookFavorites = [],
  mailCategoriesMap = {},
  onChangeMailCategoriesMap,
  mailSnoozeMap = {},
  onChangeMailSnoozeMap,
  onArchiveEmail,
  unreadBriefing = null,
  onDiscoverOutlookSettings,
  isDiscoveringOutlookSettings = false,
  localSignatures = [],
  localContacts = [],
  isSyncingLiveOutlook = false,
  onSyncLiveOutlook,
  onOpenManual,
  onOpenExecutiveDashboard,
  onAddCalendarEvent,
  onImportDroppedEmails,
  requestedToastReplyEmailId,
  onClearRequestedToastReplyEmailId,
}) => {
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home');
  const [showFolderPane, setShowFolderPane] = useState(true);
  const [showAiSplitDock, setShowAiSplitDock] = useState<boolean>(() => {
    try {
      return localStorage.getItem('bmail_pin_ai_split_dock') !== 'false';
    } catch {
      return true;
    }
  });
  const togglePinAiSplitDock = (nextValue?: boolean) => {
    setShowAiSplitDock((prev) => {
      const next = typeof nextValue === 'boolean' ? nextValue : !prev;
      try {
        localStorage.setItem('bmail_pin_ai_split_dock', String(next));
      } catch {
        // ignore storage errors
      }
      return next;
    });
  };
  const [showUnreadBriefing, setShowUnreadBriefing] = useState(false);
  const [paneWidths, setPaneWidths] = useState<Record<ResizablePaneKey, number>>(() =>
    loadSavedPaneWidths()
  );
  const paneWidthsRef = useRef(paneWidths);
  paneWidthsRef.current = paneWidths;
  const [activeResizePane, setActiveResizePane] = useState<ResizablePaneKey | null>(null);
  const activeResizeCleanupRef = useRef<(() => void) | null>(null);

  useEffect(() => {
    return () => {
      activeResizeCleanupRef.current?.();
      activeResizeCleanupRef.current = null;
    };
  }, []);

  const persistPaneWidths = (next: Record<ResizablePaneKey, number>) => {
    try {
      localStorage.setItem(PANE_WIDTHS_STORAGE_KEY, JSON.stringify(next));
    } catch {
      // ignore storage errors
    }
  };

  const resetPaneWidth = (pane: ResizablePaneKey) => {
    setPaneWidths((prev) => {
      const next = { ...prev, [pane]: DEFAULT_PANE_WIDTHS[pane] };
      persistPaneWidths(next);
      return next;
    });
  };

  const startPaneResize = (pane: ResizablePaneKey, e: React.MouseEvent<HTMLDivElement>) => {
    if (e.button !== 0) return;
    e.preventDefault();
    activeResizeCleanupRef.current?.();
    const startX = e.clientX;
    const startWidth = paneWidthsRef.current[pane];
    const dir = pane === 'folder' || pane === 'list' ? 1 : -1;
    setActiveResizePane(pane);
    let latestWidths = paneWidthsRef.current;

    const handleMouseMove = (ev: MouseEvent) => {
      const dx = ev.clientX - startX;
      const clamped = clampPaneWidth(pane, startWidth + dir * dx);
      setPaneWidths((prev) => {
        if (prev[pane] === clamped) return prev;
        const next = { ...prev, [pane]: clamped };
        latestWidths = next;
        return next;
      });
    };

    const cleanup = () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('blur', handleMouseUp);
      activeResizeCleanupRef.current = null;
    };

    const handleMouseUp = () => {
      cleanup();
      setActiveResizePane(null);
      persistPaneWidths(latestWidths);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('blur', handleMouseUp);
    activeResizeCleanupRef.current = cleanup;
  };

  const renderVerticalSplitter = (pane: ResizablePaneKey, testId: string, label: string) => {
    const isActive = activeResizePane === pane;
    return (
      <div
        role="separator"
        aria-orientation="vertical"
        aria-label={label}
        aria-valuenow={paneWidths[pane]}
        aria-valuemin={PANE_WIDTH_LIMITS[pane][0]}
        aria-valuemax={PANE_WIDTH_LIMITS[pane][1]}
        data-testid={testId}
        title={`${label} — 좌우 드래그로 가로 폭 조절 (더블클릭 기본 너비 복원)`}
        onMouseDown={(e) => startPaneResize(pane, e)}
        onDoubleClick={() => resetPaneWidth(pane)}
        className="hidden lg:flex w-[6px] -mx-[3px] z-20 cursor-col-resize shrink-0 items-center justify-center group select-none"
      >
        <div
          className={`h-full transition-all ${
            isActive
              ? 'w-[2px] bg-[#0F6CBD]'
              : 'w-[1px] bg-transparent group-hover:w-[2px] group-hover:bg-[#0F6CBD]'
          }`}
        />
      </div>
    );
  };
  const [briefingFetchedEmail, setBriefingFetchedEmail] = useState<MailItem | null>(null);
  const [showAttachSplitDock, setShowAttachSplitDock] = useState(false);
  const [selectedPreviewAttachmentIdx, setSelectedPreviewAttachmentIdx] = useState<number | null>(null);
  const [activeStoreFilter, setActiveStoreFilter] = useState<string>('');
  const [richEmailRender, setRichEmailRender] = useState<{
    id: string;
    bodyHtml: string;
    attachments: MailAttachmentMeta[];
  } | null>(null);
  const [favoritesOpen, setFavoritesOpen] = useState(true);
  const [foldersOpen, setFoldersOpen] = useState(true);
  const [searchFoldersOpen, setSearchFoldersOpen] = useState(true);
  const [projectsTreeOpen, setProjectsTreeOpen] = useState(true);
  const [pstTreeOpen, setPstTreeOpen] = useState(true);
  const [openToolbarMenu, setOpenToolbarMenu] = useState<
    'quicksteps' | 'category' | 'snooze' | 'sweep' | null
  >(null);
  const readDelayTimerRef = useRef<number | null>(null);
  const [pstFolderSearch, setPstFolderSearch] = useState('');
  const [showAllPstFolders, setShowAllPstFolders] = useState(false);

  const [activeFolder, setActiveFolder] = useState<BuiltInFolder | string>('inbox');
  const [activeProjectSubfolderId, setActiveProjectSubfolderId] = useState<string | null>(null);
  const [focusedTab, setFocusedTab] = useState<FocusedTab>('focused');
  const [listFilter, setListFilter] = useState<ListFilterMode>('all');
  const [excludeAds, setExcludeAds] = useState(false);
  const [sortMode, setSortMode] = useState<SortMode>(() =>
    outlookPrefs.mail.newestOnTop ? 'newest' : 'oldest'
  );
  const [activeEmailId, setActiveEmailId] = useState<string>('');
  const [activeDraftId, setActiveDraftId] = useState<string>('');
  const [centerMode, setCenterMode] = useState<'mail' | 'result'>('mail');
  const [isMiniAssistOpen, setIsMiniAssistOpen] = useState(false);
  const [inlineAiExpanded, setInlineAiExpanded] = useState(false);
  const [draggedEmailId, setDraggedEmailId] = useState<string | null>(null);
  const [dropTargetFolder, setDropTargetFolder] = useState<string | null>(null);
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    mail: MailItem;
  } | null>(null);

  const accent = ACCENT_THEME_COLORS[outlookPrefs.general.accentTheme] ?? ACCENT_THEME_COLORS.blue;
  const density: OutlookDensity = outlookPrefs.general.density;
  const readingPanePosition: OutlookReadingPanePosition = outlookPrefs.mail.readingPanePosition;
  const showReadingPane = readingPanePosition !== 'hidden';
  const wideReadingPane = outlookPrefs.mail.wideReadingPane;
  const conversationView = outlookPrefs.mail.conversationView;
  const focusedInboxEnabled = outlookPrefs.mail.focusedInbox;
  const previewLines: OutlookPreviewLines = outlookPrefs.mail.previewLines;
  const showSenderAvatars = outlookPrefs.mail.showSenderAvatars;
  const showUnreadBadge = outlookPrefs.general.showUnreadBadge;
  const showDrawingBadgeInList = outlookPrefs.sedms.showDrawingBadgeInList;
  const pinnedSet = useMemo(() => new Set(pinnedEmailIds), [pinnedEmailIds]);

  const patchMailPref = (patch: Partial<OutlookPreferences['mail']>) => {
    onChangeOutlookPrefs?.({
      ...outlookPrefs,
      mail: { ...outlookPrefs.mail, ...patch },
    });
  };
  const patchGeneralPref = (patch: Partial<OutlookPreferences['general']>) => {
    onChangeOutlookPrefs?.({
      ...outlookPrefs,
      general: { ...outlookPrefs.general, ...patch },
    });
  };

  useEffect(() => {
    setActiveProjectSubfolderId(null);
    onActiveFolderChange?.(activeFolder);
  }, [activeFolder, onActiveFolderChange]);

  // Authentic Outlook Pop-out HTML Compose / Reply / Forward Dialog
  const [composePopup, setComposePopup] = useState<ComposeDraftState & { open: boolean }>({
    open: false,
    title: '새 메일 작성',
    kind: 'new',
    to: '',
    cc: '',
    subject: '',
    introText: '',
  });
  const [delayedQueue, setDelayedQueue] = useState<DelayedSendItem[]>([]);
  const [queueNowMs, setQueueNowMs] = useState<number>(() => Date.now());
  const [isMsgDragActive, setIsMsgDragActive] = useState(false);
  const [inAppMailToast, setInAppMailToast] = useState<MailItem | null>(null);

  const deletedSet = useMemo(() => new Set(deletedIds), [deletedIds]);

  const activeEmails = useMemo(
    () =>
      emails.filter(
        (e) =>
          !deletedSet.has(e.id) &&
          !isBlockedSender(e, outlookPrefs.mail.blockedSenders, outlookPrefs.mail.safeSenders)
      ),
    [emails, deletedSet, outlookPrefs.mail.blockedSenders, outlookPrefs.mail.safeSenders]
  );
  const trashedEmails = useMemo(
    () =>
      emails.filter(
        (e) =>
          deletedSet.has(e.id) ||
          isBlockedSender(e, outlookPrefs.mail.blockedSenders, outlookPrefs.mail.safeSenders)
      ),
    [emails, deletedSet, outlookPrefs.mail.blockedSenders, outlookPrefs.mail.safeSenders]
  );

  // Live evaluation of user-configured Outlook Mail Rules & SEDMS drawing references
  const ruleEvaluationById = useMemo(() => {
    const map = new Map<string, ReturnType<typeof evaluateOutlookMailRules>>();
    if (outlookPrefs.mail.rules.length === 0) return map;
    for (const mail of activeEmails) {
      const evaluated = evaluateOutlookMailRules(mail, outlookPrefs.mail.rules);
      if (evaluated.matchedRuleIds.length > 0) {
        map.set(mail.id, evaluated);
      }
    }
    return map;
  }, [activeEmails, outlookPrefs.mail.rules]);

  const drawingRefsById = useMemo(() => {
    const map = new Map<string, ReturnType<typeof extractShipbuildingDrawingRefs>>();
    if (!showDrawingBadgeInList && !outlookPrefs.sedms.autoExtractDrawings) return map;
    for (const mail of activeEmails) {
      const refs = extractShipbuildingDrawingRefs(`${mail.subject} ${mail.snippet}`);
      if (refs.length > 0) {
        map.set(mail.id, refs);
      }
    }
    return map;
  }, [activeEmails, showDrawingBadgeInList, outlookPrefs.sedms.autoExtractDrawings]);

  // Automatic mail priority tagging ('긴급' | '중요' | '일반') based on content urgency & project relevance
  const priorityById = useMemo(() => {
    const map = new Map<string, MailPriorityAssessment>();
    const allForPriority = activeFolder === 'trash' ? trashedEmails : activeEmails;
    for (const mail of allForPriority) {
      const ruleEval = ruleEvaluationById.get(mail.id);
      const drawingCount = drawingRefsById.get(mail.id)?.length ?? 0;
      map.set(
        mail.id,
        classifyMailPriority(mail, {
          projects,
          projectMailOverrides,
          ruleAssignedProjectCode: ruleEval?.assignedProjectCode ?? null,
          ruleShouldFlag: ruleEval?.shouldFlag ?? false,
          drawingRefCount: drawingCount,
        })
      );
    }
    return map;
  }, [
    activeFolder,
    activeEmails,
    trashedEmails,
    projects,
    projectMailOverrides,
    ruleEvaluationById,
    drawingRefsById,
  ]);

  const priorityCounts = useMemo(() => {
    const counts: Record<MailPriorityLevel, number> = { urgent: 0, important: 0, normal: 0 };
    for (const assessment of priorityById.values()) {
      counts[assessment.level] += 1;
    }
    return counts;
  }, [priorityById]);

  const inboxCount = useMemo(
    () => activeEmails.filter((e) => e.folder === 'inbox').length,
    [activeEmails]
  );
  const unreadInboxCount = useMemo(
    () => activeEmails.filter((e) => e.folder === 'inbox' && !e.isRead).length,
    [activeEmails]
  );
  const sentCount = useMemo(
    () => activeEmails.filter((e) => e.folder === 'sent').length,
    [activeEmails]
  );
  const flaggedCount = useMemo(
    () => activeEmails.filter((e) => e.isFlagged || ruleEvaluationById.get(e.id)?.shouldFlag).length,
    [activeEmails, ruleEvaluationById]
  );

  const activeGlobalFilterCount = useMemo(
    () => countActiveGlobalSearchFilters(searchFilters),
    [searchFilters]
  );

  const searchFolderEmailIds = useMemo(() => {
    const nowIso = new Date().toISOString();
    const ids: OutlookSearchFolderId[] = [
      'unread',
      'flagged',
      'has-attachments',
      'large-mail',
      'snoozed',
    ];
    const map = new Map<OutlookSearchFolderId, Set<string>>();
    for (const id of ids) {
      map.set(
        id,
        new Set(filterEmailsBySearchFolder(activeEmails, id, mailSnoozeMap, nowIso).map((m) => m.id))
      );
    }
    return map;
  }, [activeEmails, mailSnoozeMap]);

  const activeProject = useMemo(() => {
    if (!activeFolder.startsWith('project:')) return null;
    const code = activeFolder.slice(8);
    return projects.find((p) => p.code === code) ?? null;
  }, [activeFolder, projects]);

  const projectMailCounts = useMemo(() => {
    const counts = new Map<string, number>();
    for (const p of projects) {
      let c = 0;
      for (const mail of activeEmails) {
        const ruleProj = ruleEvaluationById.get(mail.id)?.assignedProjectCode;
        if (
          ruleProj === p.code ||
          isEmailInProject(mail, p.code, projects, projectMailOverrides)
        ) {
          c += 1;
        }
      }
      counts.set(p.code, c);
    }
    return counts;
  }, [activeEmails, projects, projectMailOverrides, ruleEvaluationById]);

  const visibleEmails = useMemo(() => {
    const source = activeFolder === 'trash' ? trashedEmails : activeEmails;
    const nowIso = new Date().toISOString();

    const filtered = source.filter((mail) => {
      const ruleEval = ruleEvaluationById.get(mail.id);
      if (activeFolder !== 'trash' && ruleEval?.shouldMoveTrash) return false;

      if (activeFolder === 'inbox') {
        if (mail.folder !== 'inbox') return false;
        if (mail.folderName === '보관함' || ruleEval?.shouldArchive) return false;
        const snoozeUntil = mailSnoozeMap[mail.id];
        if (snoozeUntil && snoozeUntil > nowIso) return false;
      } else if (activeFolder === 'sent') {
        if (mail.folder !== 'sent') return false;
      } else if (activeFolder === 'archive') {
        if (mail.folderName !== '보관함' && !ruleEval?.shouldArchive) return false;
      } else if (activeFolder === 'flagged') {
        if (!mail.isFlagged && !ruleEval?.shouldFlag) return false;
      } else if (activeFolder.startsWith('search:')) {
        const sfId = activeFolder.slice(7) as OutlookSearchFolderId;
        if (!searchFolderEmailIds.get(sfId)?.has(mail.id)) return false;
      } else if (activeFolder.startsWith('project:')) {
        const targetCode = activeFolder.slice(8);
        const matchesProj =
          ruleEval?.assignedProjectCode === targetCode ||
          isEmailInProject(mail, targetCode, projects, projectMailOverrides);
        if (!matchesProj) return false;
      } else if (activeFolder !== 'all' && activeFolder !== 'trash' && activeFolder !== 'drafts') {
        if (mail.folderName !== activeFolder && mail.category !== activeFolder) return false;
      }

      if (listFilter === 'unread' && mail.isRead) return false;
      if (listFilter === 'attachments' && !mail.hasAttachment) return false;
      if (
        (listFilter === 'urgent' || listFilter === 'important' || listFilter === 'normal') &&
        priorityById.get(mail.id)?.level !== listFilter
      ) {
        return false;
      }
      if (excludeAds && isPromotionalEmail(mail)) return false;

      // Authentic New Outlook Focused / Other (중요 | 기타) split for Inbox when enabled in settings
      if (activeFolder === 'inbox' && focusedInboxEnabled) {
        const isOther = isPromotionalEmail(mail);
        if (focusedTab === 'focused' && isOther) return false;
        if (focusedTab === 'other' && !isOther) return false;
      }

      return matchesGlobalSearchFilters(mail, searchQuery, searchFilters);
    });

    return [...filtered].sort((a, b) => {
      if (sortMode === 'priority') return compareMailsByPriority(a, b, priorityById);
      if (sortMode === 'oldest') return a.timestamp.localeCompare(b.timestamp);
      if (sortMode === 'sender') return a.fromName.localeCompare(b.fromName, 'ko');
      return b.timestamp.localeCompare(a.timestamp);
    });
  }, [
    activeEmails,
    trashedEmails,
    activeFolder,
    focusedTab,
    focusedInboxEnabled,
    listFilter,
    excludeAds,
    searchQuery,
    searchFilters,
    sortMode,
    projects,
    projectMailOverrides,
    ruleEvaluationById,
    priorityById,
    mailSnoozeMap,
    searchFolderEmailIds,
  ]);

  // Conversation thread counts by normalized subject
  const threadCountByKey = useMemo(() => {
    const counts = new Map<string, number>();
    for (const mail of visibleEmails) {
      const key = mail.conversationKey || normalizeConversationSubject(mail.subject);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return counts;
  }, [visibleEmails]);

  const displayedEmails = useMemo(() => {
    if (!conversationView) return visibleEmails;
    const seen = new Set<string>();
    const collapsed: MailItem[] = [];
    for (const mail of visibleEmails) {
      const key = mail.conversationKey || normalizeConversationSubject(mail.subject);
      if (seen.has(key)) continue;
      seen.add(key);
      collapsed.push(mail);
    }
    return collapsed;
  }, [visibleEmails, conversationView]);

  // Group displayed emails into Pinned section + chronological date headers (Authentic New Outlook)
  const groupedEmails = useMemo(() => {
    const pinnedItems = displayedEmails.filter((m) => pinnedSet.has(m.id));
    const regularItems = displayedEmails.filter((m) => !pinnedSet.has(m.id));

    const result: Array<{ label: string; isPinnedGroup?: boolean; items: MailItem[] }> = [];
    if (pinnedItems.length > 0) {
      result.push({ label: '고정됨', isPinnedGroup: true, items: pinnedItems });
    }

    if (sortMode === 'priority') {
      const buckets: Record<MailPriorityLevel, MailItem[]> = {
        urgent: [],
        important: [],
        normal: [],
      };
      for (const mail of regularItems) {
        const lvl = priorityById.get(mail.id)?.level ?? 'normal';
        buckets[lvl].push(mail);
      }
      if (buckets.urgent.length > 0) {
        result.push({ label: `긴급 (${buckets.urgent.length})`, items: buckets.urgent });
      }
      if (buckets.important.length > 0) {
        result.push({ label: `중요 (${buckets.important.length})`, items: buckets.important });
      }
      if (buckets.normal.length > 0) {
        result.push({ label: `일반 (${buckets.normal.length})`, items: buckets.normal });
      }
      return result;
    }

    if (sortMode === 'sender') {
      if (regularItems.length > 0) {
        result.push({ label: '발신자순 정렬', items: regularItems });
      }
      return result;
    }

    const groups = new Map<string, MailItem[]>();
    for (const mail of regularItems) {
      const key = mail.date || '날짜 미상';
      const arr = groups.get(key) ?? [];
      arr.push(mail);
      groups.set(key, arr);
    }
    for (const [label, items] of groups.entries()) {
      result.push({ label, items });
    }
    return result;
  }, [displayedEmails, pinnedSet, sortMode, priorityById]);

  const activeEmail =
    (activeFolder === 'trash' ? trashedEmails : activeEmails).find((m) => m.id === activeEmailId) ??
    (briefingFetchedEmail && briefingFetchedEmail.id === activeEmailId ? briefingFetchedEmail : null) ??
    visibleEmails[0] ??
    null;

  const activeDraft = drafts.find((d) => d.id === activeDraftId) ?? drafts[0] ?? null;

  useEffect(() => {
    return () => {
      if (readDelayTimerRef.current !== null) {
        window.clearTimeout(readDelayTimerRef.current);
      }
    };
  }, []);

  const handleSelectMailRow = (mail: MailItem) => {
    if (readDelayTimerRef.current !== null) {
      window.clearTimeout(readDelayTimerRef.current);
      readDelayTimerRef.current = null;
    }
    setActiveEmailId(mail.id);
    if (readingPanePosition === 'hidden') {
      patchMailPref({ readingPanePosition: 'right' });
    }
    setCenterMode('mail');
    if (!mail.isRead) {
      const mode = outlookPrefs.mail.markAsReadMode;
      if (mode === 'immediate') {
        onToggleReadEmail(mail.id);
      } else if (mode === 'delay') {
        const waitMs = Math.max(1, outlookPrefs.mail.markAsReadDelaySec || 3) * 1000;
        const targetId = mail.id;
        readDelayTimerRef.current = window.setTimeout(() => {
          onToggleReadEmail(targetId);
          readDelayTimerRef.current = null;
        }, waitMs);
      }
    }
  };

  const handleBriefingOpenOrReply = async (emailId: string, mode: 'open' | 'reply') => {
    let target = emails.find((m) => m.id === emailId) || null;
    if (!target) {
      try {
        const res = await fetch(`/api/local/emails/${encodeURIComponent(emailId)}`);
        if (res.ok) {
          const data = (await res.json()) as { email?: MailItem };
          if (data.email) {
            target = data.email;
            setBriefingFetchedEmail(data.email);
          }
        }
      } catch {
        // ignore fetch failure
      }
    }
    if (!target) return;
    if (mode === 'open') {
      handleSelectMailRow(target);
      return;
    }
    handleReplyAction(target, 'reply');
  };

  const handleDeleteOrArchiveWithAdvance = (id: string, action: 'delete' | 'archive') => {
    const idx = displayedEmails.findIndex((m) => m.id === id);
    let nextSelected = '';
    if (outlookPrefs.mail.afterDeleteAction === 'next') {
      nextSelected = displayedEmails[idx + 1]?.id ?? displayedEmails[idx - 1]?.id ?? '';
    } else if (outlookPrefs.mail.afterDeleteAction === 'prev') {
      nextSelected = displayedEmails[idx - 1]?.id ?? displayedEmails[idx + 1]?.id ?? '';
    }
    if (outlookPrefs.mail.afterDeleteAction === 'list' || nextSelected !== undefined) {
      setActiveEmailId(nextSelected);
    }
    if (action === 'delete') {
      onDeleteEmail(id);
    } else if (onArchiveEmail) {
      onArchiveEmail(id);
    } else {
      onMoveEmailToFolder(id, 'sent');
    }
  };

  const handleToggleMailCategory = (mailId: string, catId: string) => {
    if (!onChangeMailCategoriesMap) return;
    const prev = mailCategoriesMap[mailId] || [];
    const nextList = prev.includes(catId) ? prev.filter((c) => c !== catId) : [...prev, catId];
    onChangeMailCategoriesMap({ ...mailCategoriesMap, [mailId]: nextList });
  };

  const handleSnoozeMail = (
    mailId: string,
    preset: '1h' | 'tomorrow9' | 'nextweek' | 'clear'
  ) => {
    if (!onChangeMailSnoozeMap) return;
    if (preset === 'clear') {
      const copy = { ...mailSnoozeMap };
      delete copy[mailId];
      onChangeMailSnoozeMap(copy);
      return;
    }
    const now = new Date();
    if (preset === '1h') {
      now.setHours(now.getHours() + 1);
    } else if (preset === 'tomorrow9') {
      now.setDate(now.getDate() + 1);
      now.setHours(9, 0, 0, 0);
    } else {
      const day = now.getDay();
      const daysUntilMon = day === 0 ? 1 : 8 - day;
      now.setDate(now.getDate() + daysUntilMon);
      now.setHours(9, 0, 0, 0);
    }
    onChangeMailSnoozeMap({ ...mailSnoozeMap, [mailId]: now.toISOString() });
    setOpenToolbarMenu(null);
  };

  const handleSweepSender = (
    senderEmail: string,
    mode: 'archive' | 'trash',
    createFutureRule: boolean
  ) => {
    const clean = senderEmail.trim().toLowerCase();
    if (!clean) return;
    const targets = activeEmails.filter((m) => m.fromEmail.trim().toLowerCase() === clean);
    for (const m of targets) {
      if (mode === 'trash') {
        onDeleteEmail(m.id);
      } else if (onArchiveEmail) {
        onArchiveEmail(m.id);
      }
    }
    if (createFutureRule) {
      const newRule = buildSweepRuleForSender(clean, mode);
      const exists = outlookPrefs.mail.rules.some((r) => r.id === newRule.id);
      if (!exists) {
        patchMailPref({ rules: [newRule, ...outlookPrefs.mail.rules] });
      }
    }
    setOpenToolbarMenu(null);
  };

  const handleRunQuickStep = (step: OutlookQuickStep, mail: MailItem) => {
    const patch = applyQuickStepActions(mail, step);
    if (patch.markRead && !mail.isRead) onToggleReadEmail(mail.id);
    if (patch.setFlagged && !mail.isFlagged) onToggleFlagEmail(mail.id);
    if (patch.setPinned && !pinnedSet.has(mail.id)) onTogglePinEmail?.(mail.id);
    if (patch.assignCategoryId && onChangeMailCategoriesMap) {
      const prev = mailCategoriesMap[mail.id] || [];
      if (!prev.includes(patch.assignCategoryId)) {
        onChangeMailCategoriesMap({
          ...mailCategoriesMap,
          [mail.id]: [...prev, patch.assignCategoryId],
        });
      }
    }
    if (patch.assignProjectCode && onAssignEmailToProject) {
      onAssignEmailToProject(patch.assignProjectCode, mail.id);
    }
    if (patch.moveFolder === 'archive') {
      handleDeleteOrArchiveWithAdvance(mail.id, 'archive');
    } else if (patch.moveFolder === 'trash') {
      handleDeleteOrArchiveWithAdvance(mail.id, 'delete');
    }
    setOpenToolbarMenu(null);
  };

  // Jump directly to a cited mail when clicked from Ask AI or Morning Brief
  useEffect(() => {
    if (!requestedActiveEmailId) return;
    const target = activeEmails.find((m) => m.id === requestedActiveEmailId);
    if (target) {
      setActiveFolder('all');
      setActiveEmailId(target.id);
      if (readingPanePosition === 'hidden') {
        patchMailPref({ readingPanePosition: 'right' });
      }
      setCenterMode('mail');
    }
    onClearRequestedActiveEmailId?.();
  }, [requestedActiveEmailId, activeEmails, onClearRequestedActiveEmailId, readingPanePosition]);

  useEffect(() => {
    if (!requestedMailto) return;
    openComposePopup('new', requestedMailto);
    onClearRequestedMailto?.();
  }, [requestedMailto, onClearRequestedMailto]);

  useEffect(() => {
    if (!requestedToastReplyEmailId) return;
    const target = activeEmails.find((m) => m.id === requestedToastReplyEmailId);
    if (target) {
      setActiveFolder('all');
      setActiveEmailId(target.id);
      setCenterMode('mail');
      handleReplyAction(target, 'reply');
    }
    onClearRequestedToastReplyEmailId?.();
  }, [requestedToastReplyEmailId, activeEmails, onClearRequestedToastReplyEmailId]);

  const dispatchMailPayloadNow = async (payload: {
    to: string;
    cc: string;
    subject: string;
    body: string;
    kind: OutlookDraftItem['kind'];
  }): Promise<string> => {
    if (pop3ConnectedEmail && onSendViaSmtp) {
      const smtpRes = await onSendViaSmtp({
        to: payload.to,
        cc: payload.cc,
        subject: payload.subject,
        body: payload.body,
      });
      return smtpRes.message;
    }
    const res = exportDraftToOutlook(payload.subject, payload.body);
    onSaveDraft({
      to: payload.to,
      cc: payload.cc,
      subject: payload.subject,
      body: payload.body,
      kind: payload.kind,
    });
    return res.deliveredVia === 'outlook-compose'
      ? 'Outlook 메일 작성 창으로 전달했으며 임시 보관함에도 사본을 저장했습니다.'
      : '지연 발송 큐에서 임시 보관함 사본으로 저장했습니다.';
  };

  useEffect(() => {
    const hasQueued = delayedQueue.some((q) => q.status === 'queued');
    if (!hasQueued) return;
    const timer = window.setInterval(() => {
      const now = Date.now();
      setQueueNowMs(now);
      const ready = delayedQueue.filter((q) => q.status === 'queued' && now >= q.sendAtMs);
      if (ready.length > 0) {
        setDelayedQueue((prev) =>
          prev.map((item) =>
            item.status === 'queued' && now >= item.sendAtMs ? { ...item, status: 'sending' } : item
          )
        );
        ready.forEach((item) => {
          void dispatchMailPayloadNow(item)
            .then(() => {
              setDelayedQueue((prev) => prev.filter((q) => q.id !== item.id));
            })
            .catch((err) => {
              setDelayedQueue((prev) =>
                prev.map((q) =>
                  q.id === item.id
                    ? {
                        ...q,
                        status: 'failed',
                        errorMessage: err instanceof Error ? err.message : '발송 실패',
                      }
                    : q
                )
              );
            });
        });
      }
    }, 500);
    return () => window.clearInterval(timer);
  }, [delayedQueue]);

  const handleDropMailFiles = async (fileList: FileList | null) => {
    if (!fileList || fileList.length === 0) return;
    const mailFiles = Array.from(fileList).filter((f) =>
      /\.(msg|eml|mbox|pst|ost)$/i.test(f.name)
    );
    if (mailFiles.length === 0) return;
    try {
      const formData = new FormData();
      mailFiles.forEach((f) => formData.append('files', f));
      const res = await fetch('/api/pst/import', { method: 'POST', body: formData });
      if (!res.ok) return;
      const data = (await res.json()) as { emails?: MailItem[] };
      if (Array.isArray(data.emails) && data.emails.length > 0) {
        onImportDroppedEmails?.(data.emails);
        setActiveFolder('all');
        setActiveEmailId(data.emails[0].id);
        setCenterMode('mail');
      }
    } catch {
      // ignore transient drop error
    }
  };

  const handleTriggerMailToast = (mail: MailItem) => {
    const showToastIpc = (
      window as unknown as {
        bmailDesktop?: {
          showMailToast?: (payload: {
            emailId: string;
            sender: string;
            subject: string;
            snippet: string;
          }) => Promise<boolean>;
        };
      }
    ).bmailDesktop?.showMailToast;
    if (showToastIpc) {
      void showToastIpc({
        emailId: mail.id,
        sender: `${mail.fromName} <${mail.fromEmail}>`,
        subject: mail.subject,
        snippet: mail.snippet || mail.body.slice(0, 90),
      });
    } else {
      setInAppMailToast(mail);
    }
  };

  const inlineAiAnalysis = useMemo(
    () => (activeEmail ? classifyEmailIntent(activeEmail) : null),
    [activeEmail]
  );

  // Fetch sanitized rich HTML body + inline cid: image mapping for active email
  useEffect(() => {
    setSelectedPreviewAttachmentIdx(null);
    if (!activeEmail || !localDbAvailable) {
      setRichEmailRender(null);
      return;
    }
    const controller = new AbortController();
    fetch(`/api/local/emails/${encodeURIComponent(activeEmail.id)}/render`, {
      signal: controller.signal,
    })
      .then(async (res) => {
        if (!res.ok) return null;
        return (await res.json()) as {
          id: string;
          bodyHtml: string;
          attachments: MailAttachmentMeta[];
        };
      })
      .then((data) => {
        if (data && data.id === activeEmail.id) {
          setRichEmailRender(data);
        }
      })
      .catch(() => {
        // Ignore abort or transient offline error; plain body remains displayed
      });
    return () => {
      controller.abort();
    };
  }, [activeEmail?.id, localDbAvailable]);

  const inlineReplyOptions = useMemo(
    () => (activeEmail ? generateInstantReplyOptions(activeEmail) : []),
    [activeEmail]
  );

  const currentFolderTitle = useMemo(() => {
    switch (activeFolder) {
      case 'all':
        return '전체 메일함';
      case 'inbox':
        return '받은 편지함';
      case 'sent':
        return '보낸 편지함';
      case 'drafts':
        return '임시 보관함';
      case 'flagged':
        return '플래그 지정됨';
      case 'archive':
        return '보관함';
      case 'trash':
        return '지운 편지함';
      case 'search:unread':
        return '검색 폴더 · 읽지 않은 메일';
      case 'search:flagged':
        return '검색 폴더 · 플래그된 메일';
      case 'search:has-attachments':
        return '검색 폴더 · 첨부파일 있음';
      case 'search:large-mail':
        return '검색 폴더 · 대용량 메일 (1MB+)';
      case 'search:snoozed':
        return '검색 폴더 · 다시 알림 예약됨';
      default:
        return activeFolder.startsWith('project:') ? `프로젝트 ${activeFolder.slice(8)}` : activeFolder;
    }
  }, [activeFolder]);

  // Dismiss context menu on global click / Escape & support Outlook keyboard shortcuts
  useEffect(() => {
    const closeMenu = () => {
      setContextMenu(null);
      setOpenToolbarMenu(null);
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        setContextMenu(null);
        setOpenToolbarMenu(null);
        setComposePopup((prev) => (prev.open ? { ...prev, open: false } : prev));
        return;
      }
      const tag = (e.target as HTMLElement | null)?.tagName;
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return;

      if (e.key === 'Delete' && activeEmail && activeFolder !== 'trash' && activeFolder !== 'drafts') {
        handleDeleteOrArchiveWithAdvance(activeEmail.id, 'delete');
      } else if ((e.key === 'ArrowDown' || e.key === 'ArrowUp') && displayedEmails.length > 0) {
        e.preventDefault();
        const idx = displayedEmails.findIndex((m) => m.id === activeEmail?.id);
        const nextIdx =
          e.key === 'ArrowDown'
            ? Math.min(displayedEmails.length - 1, idx + 1)
            : Math.max(0, idx - 1);
        const target = displayedEmails[nextIdx];
        if (target) {
          handleSelectMailRow(target);
        }
      }
    };
    window.addEventListener('click', closeMenu);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('click', closeMenu);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [activeEmail, activeFolder, displayedEmails, onDeleteEmail]);

  const handleDropOnFolder = (targetFolder: BuiltInFolder, e: React.DragEvent) => {
    e.preventDefault();
    const id = e.dataTransfer.getData('text/plain') || draggedEmailId;
    setDropTargetFolder(null);
    setDraggedEmailId(null);
    if (!id) return;
    if (targetFolder === 'trash') {
      onDeleteEmail(id);
    } else if (targetFolder === 'archive') {
      handleDeleteOrArchiveWithAdvance(id, 'archive');
    } else if (targetFolder === 'flagged') {
      onToggleFlagEmail(id);
    } else if (targetFolder === 'inbox' || targetFolder === 'sent') {
      onMoveEmailToFolder(id, targetFolder);
    }
  };

  const handleRunAndShowResult = (overrideProvider?: AiProviderId) => {
    if (readingPanePosition === 'hidden') {
      patchMailPref({ readingPanePosition: 'right' });
    }
    setCenterMode('result');
    onRunAnalysis(overrideProvider);
  };

  const openPopOutComposeWindow = (snapshot: ComposeDraftState) => {
    const key = `bmail_compose_draft_${Date.now()}`;
    try {
      window.localStorage.setItem(key, JSON.stringify(snapshot));
    } catch {
      // ignore storage quota errors
    }
    const desktopBridge = (
      window as unknown as {
        bmailDesktop?: {
          openComposeWindow?: (draftKey: string) => Promise<{ ok?: boolean }>;
        };
      }
    ).bmailDesktop;
    if (desktopBridge?.openComposeWindow) {
      void desktopBridge.openComposeWindow(key).then((res) => {
        if (!res?.ok) {
          setComposePopup({ ...snapshot, open: true });
        }
      });
      return true;
    }
    const win = window.open(
      `/?compose=1&draftKey=${encodeURIComponent(key)}`,
      '_blank',
      'width=1080,height=780,noopener'
    );
    if (win) return true;
    return false;
  };

  const openComposePopup = (
    kind: OutlookDraftItem['kind'],
    preset?: Partial<ComposeDraftState> & { body?: string; preferInline?: boolean }
  ) => {
    const titleMap: Record<OutlookDraftItem['kind'], string> = {
      new: '새 메일',
      reply: '회신',
      replyAll: '전체 회신',
      forward: '전달',
      worklog: '업무 일지 메일 작성',
    };
    const nextDraft: ComposeDraftState = {
      kind,
      title: preset?.title || titleMap[kind],
      to: preset?.to ?? '',
      cc: preset?.cc ?? '',
      subject: preset?.subject ?? '',
      introText: preset?.introText ?? (preset?.sourceEmailId ? '' : (preset?.body ?? '')),
      sourceEmailId: preset?.sourceEmailId,
      quoteFromLabel: preset?.quoteFromLabel,
      quoteSentLabel: preset?.quoteSentLabel,
      quoteToLabel: preset?.quoteToLabel,
      quoteCcLabel: preset?.quoteCcLabel,
      quoteSubjectLabel: preset?.quoteSubjectLabel,
      quotePlainBody: preset?.quotePlainBody,
      quoteHtmlBody: preset?.quoteHtmlBody,
      attachments: preset?.attachments,
      importance: preset?.importance || outlookPrefs.mail.defaultImportance,
    };

    const desktopBridge = (
      window as unknown as {
        bmailDesktop?: { openComposeWindow?: (draftKey: string) => Promise<{ ok?: boolean }> };
      }
    ).bmailDesktop;
    if (!preset?.preferInline && desktopBridge?.openComposeWindow) {
      openPopOutComposeWindow(nextDraft);
      return;
    }
    setComposePopup({ ...nextDraft, open: true });
  };

  const handleReplyAction = (mail: MailItem, mode: 'reply' | 'replyAll' | 'forward') => {
    const quoteFromLabel = `${mail.fromName} <${mail.fromEmail}>`;
    const quoteSentLabel = `${mail.date} ${mail.timeLabel}`;
    const quoteToLabel = mail.to.join('; ');
    const quoteCcLabel = mail.cc?.join('; ') || '';
    if (mode === 'reply') {
      openComposePopup('reply', {
        to: mail.fromEmail,
        subject: mail.subject.startsWith('RE:') ? mail.subject : `RE: ${mail.subject}`,
        sourceEmailId: mail.id,
        quoteFromLabel,
        quoteSentLabel,
        quoteToLabel,
        quoteCcLabel,
        quoteSubjectLabel: mail.subject,
        quotePlainBody: mail.body,
      });
    } else if (mode === 'replyAll') {
      const others = mail.to.filter((addr) => addr !== mail.fromEmail).join(', ');
      openComposePopup('replyAll', {
        to: mail.fromEmail,
        cc: others,
        subject: mail.subject.startsWith('RE:') ? mail.subject : `RE: ${mail.subject}`,
        sourceEmailId: mail.id,
        quoteFromLabel,
        quoteSentLabel,
        quoteToLabel,
        quoteCcLabel,
        quoteSubjectLabel: mail.subject,
        quotePlainBody: mail.body,
      });
    } else {
      openComposePopup('forward', {
        to: '',
        subject: mail.subject.startsWith('FW:') ? mail.subject : `FW: ${mail.subject}`,
        sourceEmailId: mail.id,
        quoteFromLabel,
        quoteSentLabel,
        quoteToLabel,
        quoteCcLabel,
        quoteSubjectLabel: mail.subject,
        quotePlainBody: mail.body,
        attachments: mail.attachments,
      });
    }
  };

  const handleSaveWorklogDraftAndPopup = (subject: string, body: string) => {
    onSaveDraft({
      to: '',
      cc: '',
      subject,
      body,
      kind: 'worklog',
    });
    openComposePopup('worklog', { subject, body });
  };

  const renderFolderButton = (
    id: BuiltInFolder,
    label: string,
    icon: React.ReactNode,
    count: number,
    unreadHighlight?: number,
    suppressActiveHighlight = false,
    allowFavoriteStar = false
  ) => {
    const isActive = activeFolder === id && !suppressActiveHighlight;
    const isDropOver = dropTargetFolder === id;
    const canDrop = id === 'inbox' || id === 'sent' || id === 'flagged' || id === 'trash';
    const isFavorited = favoriteFolderIds.includes(id);
    const displayUnread = showUnreadBadge && unreadHighlight && unreadHighlight > 0;
    return (
      <div
        key={`${id}-${allowFavoriteStar ? 'tree' : 'fav'}`}
        onDragOver={(e) => {
          if (!canDrop) return;
          e.preventDefault();
          setDropTargetFolder(id);
        }}
        onDragLeave={() => {
          if (dropTargetFolder === id) setDropTargetFolder(null);
        }}
        onDrop={(e) => canDrop && handleDropOnFolder(id, e)}
        className={`group/folder flex items-center rounded-[4px] transition-colors duration-150 ${
          isDropOver
            ? 'bg-[#CFE4FA] ring-1 ring-[#0F6CBD] text-[#0F6CBD] font-semibold'
            : isActive
              ? 'bg-[#CFE4FA] text-[#0F6CBD] font-semibold'
              : 'text-[#242424] hover:bg-black/[0.04] active:bg-black/[0.07]'
        }`}
      >
        <button
          type="button"
          onClick={() => setActiveFolder(id)}
          className="flex-1 min-w-0 flex items-center justify-between gap-2 px-2.5 py-1.5 text-xs whitespace-nowrap text-left rounded-[4px] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
        >
          <span className="inline-flex items-center gap-2.5 min-w-0">
            <span className={isActive ? 'text-[#0F6CBD]' : 'text-[#616161]'}>{icon}</span>
            <span className="truncate">{label}</span>
          </span>
          {count > 0 && (
            <span
              className={`font-mono tabular-nums text-[11px] shrink-0 ${
                displayUnread
                  ? 'text-[#0F6CBD] font-semibold'
                  : isActive
                    ? 'text-[#0F6CBD]'
                    : 'text-[#616161]'
              }`}
            >
              {displayUnread ? unreadHighlight : count}
            </span>
          )}
        </button>
        {allowFavoriteStar && onToggleFavoriteFolder && (
          <button
            type="button"
            onClick={(e) => {
              e.stopPropagation();
              onToggleFavoriteFolder(id);
            }}
            aria-label={isFavorited ? `${label} 즐겨찾기 해제` : `${label} 즐겨찾기 추가`}
            title={isFavorited ? '즐겨찾기에서 제거' : '즐겨찾기에 추가'}
            className={`mr-1 p-1 rounded transition-opacity shrink-0 ${
              isFavorited
                ? 'text-amber-500 opacity-100'
                : 'text-[#9E9E9E] opacity-0 group-hover/folder:opacity-100 hover:text-amber-500'
            }`}
          >
            <Star className="w-3 h-3" fill={isFavorited ? 'currentColor' : 'none'} />
          </button>
        )}
      </div>
    );
  };

  return (
    <div
      className="relative flex flex-col flex-1 min-h-0 bg-[#F5F5F5]"
      onDragOver={(e) => {
        if (e.dataTransfer?.types?.includes('Files')) {
          e.preventDefault();
          setIsMsgDragActive(true);
        }
      }}
      onDragLeave={(e) => {
        if (!e.currentTarget.contains(e.relatedTarget as Node | null)) {
          setIsMsgDragActive(false);
        }
      }}
      onDrop={(e) => {
        if (e.dataTransfer?.files && e.dataTransfer.files.length > 0) {
          e.preventDefault();
          setIsMsgDragActive(false);
          void handleDropMailFiles(e.dataTransfer.files);
        } else {
          setIsMsgDragActive(false);
        }
      }}
    >
      {isMsgDragActive && (
        <div className="pointer-events-none absolute inset-0 z-50 bg-[#0F6CBD]/15 backdrop-blur-[1px] border-2 border-dashed border-[#0F6CBD] flex items-center justify-center">
          <div className="bg-white px-6 py-4 rounded-xl shadow-xl border border-[#B4D6FA] flex items-center gap-3 text-sm font-bold text-[#0F6CBD]">
            <Upload className="w-5 h-5" />
            Outlook .msg / .eml / .pst 파일을 놓으면 즉시 파싱하여 열어봅니다
          </div>
        </div>
      )}
      {/* Row 1: Authentic New Outlook Ribbon Tab Strip (홈 / 보기 / 도움말) */}
      <div className="h-9 px-2.5 flex items-center justify-between gap-4 shrink-0 select-none">
        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => setShowFolderPane((v) => !v)}
            aria-label="폴더 창 표시 전환"
            aria-pressed={showFolderPane}
            title="탐색 창 숨기기/표시"
            className="w-8 h-7 flex items-center justify-center rounded-[4px] hover:bg-black/[0.05] active:bg-black/[0.08] text-[#424242] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
          >
            <Menu className="w-4 h-4" strokeWidth={1.75} />
          </button>

          <div role="tablist" aria-label="Outlook 리본 탭" className="flex items-center gap-1 ml-1">
            {(
              [
                { id: 'home', label: '홈' },
                { id: 'view', label: '보기' },
                { id: 'help', label: '도움말 · 설치' },
              ] as const
            ).map((tab) => (
              <button
                key={tab.id}
                type="button"
                role="tab"
                aria-selected={ribbonTab === tab.id}
                onClick={() => setRibbonTab(tab.id)}
                className={`relative px-3 py-1.5 text-xs transition-colors duration-150 whitespace-nowrap rounded-[4px] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60 ${
                  ribbonTab === tab.id
                    ? 'font-semibold text-[#242424]'
                    : 'text-[#616161] hover:text-[#242424] hover:bg-black/[0.04]'
                }`}
              >
                {tab.label}
                {ribbonTab === tab.id && (
                  <span
                    style={{ backgroundColor: accent.primary }}
                    className="absolute left-2.5 right-2.5 bottom-0 h-[2px] rounded-full"
                  />
                )}
              </button>
            ))}
          </div>
        </div>

        {/* Quiet live account / connected PST store status on right of tab row */}
        <div className="flex items-center gap-2 text-xs text-[#616161] min-w-0">
          {localDbAvailable && localTotalEmails > 0 && (
            <span className="hidden lg:inline-flex items-center gap-1 px-2 py-0.5 rounded-[4px] bg-[#EBF3FC] text-[#0F6CBD] font-medium whitespace-nowrap">
              로컬 보관 {localTotalEmails.toLocaleString()}건
            </span>
          )}
          {connectedStoreName && (
            <span className="hidden md:inline truncate max-w-[240px]">
              연결된 데이터 파일: <strong className="text-[#242424] font-medium">{connectedStoreName}</strong>
            </span>
          )}
          {microsoftAuth.connected ? (
            <div className="flex items-center gap-1.5">
              <span className="truncate max-w-[180px] text-[#242424] font-medium">
                {microsoftAuth.userEmail}
              </span>
              <button
                type="button"
                disabled={isSyncingGraph}
                onClick={onSyncMicrosoftEmails}
                className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] active:bg-[#EBEBEB] rounded-[4px] transition-colors duration-150 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
              >
                <RefreshCw className={`w-3 h-3 ${isSyncingGraph ? 'animate-spin' : ''}`} strokeWidth={1.75} />
                동기화
              </button>
              <button
                type="button"
                onClick={onDisconnectMicrosoft}
                aria-label="Microsoft 계정 연결 해제"
                className="p-1 text-[#616161] hover:text-[#242424] rounded-[4px] hover:bg-black/[0.05] active:bg-black/[0.08] transition-colors duration-150"
              >
                <LogOut className="w-3.5 h-3.5" strokeWidth={1.75} />
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={microsoftAuth.configured ? onConnectMicrosoft : onOpenInstallGuide}
              className="inline-flex items-center gap-1.5 px-2.5 py-1 text-[11px] font-medium text-[#0F6CBD] hover:bg-[#EBF3FC] active:bg-[#DCEBFA] rounded-[4px] transition-colors duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]/60"
            >
              <CloudDownload className="w-3.5 h-3.5" strokeWidth={1.75} />
              {microsoftAuth.configured ? 'Microsoft 365 연결' : '실제 Outlook 플러그인 설치'}
            </button>
          )}
        </div>
      </div>

      {/* Row 2: Authentic New Outlook Floating White Rounded Command Ribbon */}
      <div className="relative mx-2 mb-2 shrink-0">
      <div
        role="toolbar"
        aria-label="Outlook 명령 리본"
        className={`${
          outlookPrefs.mail.singleLineRibbon ? 'px-1.5 py-0.5 gap-1' : 'px-2 py-1.5 gap-2'
        } bg-white border border-[#E0E0E0] rounded-[8px] shadow-[0_1px_2px_rgba(0,0,0,0.06),0_0_2px_rgba(0,0,0,0.04)] flex items-center justify-between overflow-x-auto`}
      >
        {ribbonTab === 'home' && (
          <div className="flex items-center gap-0.5 shrink-0">
            {/* Signature Outlook Blue New Mail Button */}
            <button
              type="button"
              onClick={() => openComposePopup('new')}
              style={{ backgroundColor: accent.primary }}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white hover:brightness-95 active:brightness-90 rounded-[4px] transition-all duration-150 whitespace-nowrap focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-offset-1 focus-visible:ring-[#0F6CBD]"
            >
              <PenSquare className="w-3.5 h-3.5" strokeWidth={1.75} />
              새 메일
            </button>

            <div className="h-5 w-px bg-[#E0E0E0] mx-1.5" aria-hidden="true" />

            {/* Authentic Outlook Message Actions */}
            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() =>
                activeEmail &&
                (activeFolder === 'trash'
                  ? onRestoreEmail(activeEmail.id)
                  : handleDeleteOrArchiveWithAdvance(activeEmail.id, 'delete'))
              }
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              {activeFolder === 'trash' ? (
                <>
                  <RotateCcw className="w-3.5 h-3.5 text-[#0F6CBD]" />
                  복원
                </>
              ) : (
                <>
                  <Trash2 className="w-3.5 h-3.5 text-[#616161]" />
                  삭제
                </>
              )}
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts' || activeFolder === 'trash'}
              onClick={() => activeEmail && handleDeleteOrArchiveWithAdvance(activeEmail.id, 'archive')}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <Archive className="w-3.5 h-3.5 text-[#0F6CBD]" />
              보관
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() => activeEmail && handleReplyAction(activeEmail, 'reply')}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <Reply className="w-3.5 h-3.5 text-[#0F6CBD]" />
              회신
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() => activeEmail && handleReplyAction(activeEmail, 'replyAll')}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <ReplyAll className="w-3.5 h-3.5 text-[#0F6CBD]" />
              전체 회신
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() => activeEmail && handleReplyAction(activeEmail, 'forward')}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <Forward className="w-3.5 h-3.5 text-[#0F6CBD]" />
              전달
            </button>

            <div className="h-5 w-px bg-[#E0E0E0] mx-1.5" aria-hidden="true" />

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() => activeEmail && onToggleReadEmail(activeEmail.id)}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <MailOpen className="w-3.5 h-3.5 text-[#616161]" />
              읽음/안읽음
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={() => activeEmail && onToggleFlagEmail(activeEmail.id)}
              className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
            >
              <Flag
                className={`w-3.5 h-3.5 ${
                  activeEmail?.isFlagged ? 'text-red-600' : 'text-[#616161]'
                }`}
              />
              플래그
            </button>

            {onTogglePinEmail && (
              <button
                type="button"
                disabled={!activeEmail || activeFolder === 'drafts'}
                onClick={() => activeEmail && onTogglePinEmail(activeEmail.id)}
                className="inline-flex items-center gap-1.5 px-2 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] disabled:text-[#BDBDBD] disabled:hover:bg-transparent rounded-[5px] whitespace-nowrap"
              >
                <Pin
                  className={`w-3.5 h-3.5 ${
                    activeEmail && pinnedSet.has(activeEmail.id)
                      ? 'text-[#0F6CBD]'
                      : 'text-[#616161]'
                  }`}
                />
                고정
              </button>
            )}

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={(e) => {
                e.stopPropagation();
                setOpenToolbarMenu((v) => (v === 'quicksteps' ? null : 'quicksteps'));
              }}
              className={`inline-flex items-center gap-1 px-2 py-1.5 text-xs rounded-[5px] whitespace-nowrap disabled:text-[#BDBDBD] ${
                openToolbarMenu === 'quicksteps' ? 'bg-[#EBF3FC] text-[#0F6CBD]' : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <Zap className="w-3.5 h-3.5 text-[#0F6CBD]" />
              빠른 단계
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={(e) => {
                e.stopPropagation();
                setOpenToolbarMenu((v) => (v === 'category' ? null : 'category'));
              }}
              className={`inline-flex items-center gap-1 px-2 py-1.5 text-xs rounded-[5px] whitespace-nowrap disabled:text-[#BDBDBD] ${
                openToolbarMenu === 'category' ? 'bg-[#EBF3FC] text-[#0F6CBD]' : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <Tag className="w-3.5 h-3.5 text-[#616161]" />
              범주
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={(e) => {
                e.stopPropagation();
                setOpenToolbarMenu((v) => (v === 'snooze' ? null : 'snooze'));
              }}
              className={`inline-flex items-center gap-1 px-2 py-1.5 text-xs rounded-[5px] whitespace-nowrap disabled:text-[#BDBDBD] ${
                openToolbarMenu === 'snooze' ? 'bg-[#EBF3FC] text-[#0F6CBD]' : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <Clock className="w-3.5 h-3.5 text-[#616161]" />
              다시 알림
            </button>

            <button
              type="button"
              disabled={!activeEmail || activeFolder === 'drafts'}
              onClick={(e) => {
                e.stopPropagation();
                setOpenToolbarMenu((v) => (v === 'sweep' ? null : 'sweep'));
              }}
              className={`inline-flex items-center gap-1 px-2 py-1.5 text-xs rounded-[5px] whitespace-nowrap disabled:text-[#BDBDBD] ${
                openToolbarMenu === 'sweep' ? 'bg-[#EBF3FC] text-[#0F6CBD]' : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD]" />
              일괄 정리
            </button>

            <div className="h-5 w-px bg-[#E0E0E0] mx-1.5" aria-hidden="true" />

            <button
              type="button"
              onClick={onOpenPstModal}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-[#242424] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
            >
              <Database className="w-3.5 h-3.5 text-[#0F6CBD]" />
              PST · EML 열기
            </button>

            {onOpenPop3Modal && (
              <button
                type="button"
                onClick={onOpenPop3Modal}
                className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-[4px] border transition-colors whitespace-nowrap ${
                  pop3ConnectedEmail
                    ? 'bg-[#EDEBE9] text-[#242424] border-[#D1D1D1]'
                    : 'bg-white border-transparent text-[#242424] hover:bg-[#F3F2F1]'
                }`}
              >
                <Plug className={`w-3.5 h-3.5 ${pop3ConnectedEmail ? 'text-emerald-600' : 'text-[#0F6CBD]'}`} />
                {pop3ConnectedEmail ? `POP3 연결됨` : 'POP3 · 메일플러그'}
              </button>
            )}

            <div className="h-5 w-px bg-[#E1DFDD] mx-1.5" aria-hidden="true" />

            {/* Classic Monochrome Ribbon Add-in, Split Docks, Live Sync & Briefing Toggle */}
            {onDiscoverOutlookSettings && (
              <button
                type="button"
                disabled={isDiscoveringOutlookSettings}
                onClick={() => void onDiscoverOutlookSettings()}
                title="현재 PC Outlook의 서명(명함)·계정·즐겨찾기·읽기창·근무시간 등 모든 설정을 탐색해 즉시 반영합니다"
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] bg-white hover:bg-[#F3F2F1] border border-transparent text-[#242424] text-xs font-medium transition-colors whitespace-nowrap disabled:opacity-60"
              >
                <Sliders className={`w-3.5 h-3.5 text-[#424242] ${isDiscoveringOutlookSettings ? 'animate-spin' : ''}`} />
                <span>{isDiscoveringOutlookSettings ? '설정 탐색 중...' : '아웃룩 설정 탐색·적용'}</span>
              </button>
            )}

            {onSyncLiveOutlook && (
              <button
                type="button"
                disabled={isSyncingLiveOutlook}
                onClick={onSyncLiveOutlook}
                title="1회 불러온 데이터를 영구 저장하고 변경된 PST 및 실행 중인 Outlook(COM) 신규 메일만 초고속 증분 동기화합니다"
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] bg-white hover:bg-[#F3F2F1] border border-transparent text-[#242424] text-xs font-medium transition-colors whitespace-nowrap disabled:opacity-60"
              >
                <RefreshCw className={`w-3.5 h-3.5 text-[#0F6CBD] ${isSyncingLiveOutlook ? 'animate-spin' : ''}`} />
                <span>{isSyncingLiveOutlook ? '증분 동기화 중...' : '신규 메일 증분 동기화'}</span>
              </button>
            )}

            <button
              type="button"
              aria-pressed={showAiSplitDock}
              onClick={() => togglePinAiSplitDock()}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] border text-xs font-medium transition-colors whitespace-nowrap ${
                showAiSplitDock
                  ? 'bg-[#EDEBE9] border-[#D1D1D1] text-[#0F6CBD] font-semibold'
                  : 'bg-white border-transparent text-[#242424] hover:bg-[#F3F2F1]'
              }`}
            >
              <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD]" />
              <span>AI 분석 도크 (고정)</span>
            </button>

            {centerMode === 'mail' && (
              <button
                type="button"
                aria-pressed={showUnreadBriefing}
                onClick={() => setShowUnreadBriefing((v) => !v)}
                title="최근 14일 미확인 메일 브리핑 팝업 열기/닫기"
                className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] border text-xs font-medium transition-colors whitespace-nowrap ${
                  showUnreadBriefing
                    ? 'bg-[#EDEBE9] border-[#D1D1D1] text-[#0F6CBD] font-semibold'
                    : 'bg-white border-transparent text-[#242424] hover:bg-[#F3F2F1]'
                }`}
              >
                <Sparkles className="w-3.5 h-3.5 text-[#424242]" />
                <span>미확인 브리핑</span>
                {unreadBriefing && unreadBriefing.totalUnread > 0 && (
                  <span className="px-1.5 py-0 rounded-full bg-[#F3F2F1] text-[#424242] text-[10px] font-semibold tabular-nums">
                    {unreadBriefing.totalUnread}
                  </span>
                )}
              </button>
            )}

            <button
              type="button"
              aria-pressed={showAttachSplitDock}
              onClick={() => setShowAttachSplitDock((v) => !v)}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] border text-xs font-medium transition-colors whitespace-nowrap ${
                showAttachSplitDock
                  ? 'bg-[#EDEBE9] border-[#D1D1D1] text-[#0F6CBD] font-semibold'
                  : 'bg-white border-transparent text-[#242424] hover:bg-[#F3F2F1]'
              }`}
            >
              <Paperclip className="w-3.5 h-3.5 text-[#424242]" />
              <span>첨부·폴더저장 도크</span>
            </button>

            {onOpenExecutiveDashboard && (
              <button
                type="button"
                onClick={onOpenExecutiveDashboard}
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[4px] bg-white hover:bg-[#F3F2F1] border border-transparent text-[#242424] text-xs font-medium transition-colors whitespace-nowrap"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-[#424242]" />
                <span>주간·월간 대시보드</span>
              </button>
            )}

            {onOpenManual && (
              <button
                type="button"
                onClick={onOpenManual}
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[5px] bg-white hover:bg-[#F5F5F5] text-[#242424] text-xs font-semibold transition-colors whitespace-nowrap"
              >
                <BookOpen className="w-3.5 h-3.5 text-[#0F6CBD]" />
                <span>모션 매뉴얼</span>
              </button>
            )}

            <button
              type="button"
              aria-pressed={isMiniAssistOpen}
              onClick={() => setIsMiniAssistOpen((v) => !v)}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[5px] border text-xs font-semibold transition-colors whitespace-nowrap ${
                isMiniAssistOpen
                  ? 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]'
                  : 'bg-white border-transparent text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <Languages className="w-3.5 h-3.5 text-[#0F6CBD]" />
              번역·회신 미니창
            </button>

            <button
              type="button"
              aria-pressed={isTaskpaneOpen}
              onClick={onRibbonIconClick}
              title={
                isTaskpaneOpen && isTaskpanePinned
                  ? 'Bmail AI 작업 창이 고정되어 열려 있습니다'
                  : 'Bmail AI 작업 창 열기 (SupportsPinning)'
              }
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[5px] border text-xs font-semibold transition-colors whitespace-nowrap ${
                isTaskpaneOpen
                  ? 'bg-[#EBF3FC] border-[#B4D6FA] text-[#0F6CBD]'
                  : 'bg-white border-transparent text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <img
                src="/assets/icon-32.png"
                alt="Bmail AI 작업 창"
                width={16}
                height={16}
                className="w-4 h-4 rounded"
              />
              <span>AI 작업창</span>
            </button>

            <button
              type="button"
              onClick={() => {
                if (readingPanePosition === 'hidden') {
                  patchMailPref({ readingPanePosition: 'right' });
                }
                setCenterMode(centerMode === 'mail' ? 'result' : 'mail');
              }}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-[5px] text-xs font-medium transition-colors whitespace-nowrap ${
                centerMode === 'result'
                  ? 'bg-[#242424] text-white'
                  : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              {centerMode === 'result' ? '메일 읽기' : '일지·블로그'}
            </button>
          </div>
        )}

        {ribbonTab === 'view' && (
          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              aria-pressed={showFolderPane}
              onClick={() => setShowFolderPane((v) => !v)}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs rounded-[5px] whitespace-nowrap ${
                showFolderPane
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] font-medium'
                  : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <PanelLeft className="w-3.5 h-3.5" />
              폴더 창
            </button>

            <div className="inline-flex items-center gap-1 px-2 py-1 text-xs text-[#616161]">
              <PanelRight className="w-3.5 h-3.5 text-[#0F6CBD]" />
              <select
                aria-label="읽기 창 위치"
                value={readingPanePosition}
                onChange={(e) =>
                  patchMailPref({
                    readingPanePosition: e.target.value as OutlookReadingPanePosition,
                  })
                }
                className="text-xs text-[#242424] bg-transparent border border-[#D1D1D1] rounded px-1.5 py-0.5"
              >
                <option value="right">읽기 창: 오른쪽</option>
                <option value="bottom">읽기 창: 아래쪽</option>
                <option value="hidden">읽기 창: 숨김</option>
              </select>
            </div>

            {readingPanePosition === 'right' && (
              <button
                type="button"
                aria-pressed={wideReadingPane}
                onClick={() => patchMailPref({ wideReadingPane: !wideReadingPane })}
                className={`inline-flex items-center gap-1 px-2.5 py-1.5 text-xs rounded-[5px] whitespace-nowrap ${
                  wideReadingPane
                    ? 'bg-[#EBF3FC] text-[#0F6CBD] font-medium'
                    : 'text-[#242424] hover:bg-[#F5F5F5]'
                }`}
              >
                넓은 읽기 창
              </button>
            )}

            <div className="inline-flex items-center gap-1 px-1.5 py-1 text-xs text-[#616161]">
              <LayoutList className="w-3.5 h-3.5 text-[#0F6CBD]" />
              <select
                aria-label="목록 밀도"
                value={density}
                onChange={(e) => patchGeneralPref({ density: e.target.value as OutlookDensity })}
                className="text-xs text-[#242424] bg-transparent border border-[#D1D1D1] rounded px-1.5 py-0.5"
              >
                <option value="roomy">밀도: 넓게</option>
                <option value="comfortable">밀도: 표준</option>
                <option value="compact">밀도: 간략히</option>
              </select>
            </div>

            <select
              aria-label="미리보기 줄 수"
              value={previewLines}
              onChange={(e) =>
                patchMailPref({ previewLines: Number(e.target.value) as OutlookPreviewLines })
              }
              className="text-xs text-[#242424] bg-transparent border border-[#D1D1D1] rounded px-1.5 py-1"
            >
              <option value={0}>미리보기 없음</option>
              <option value={1}>미리보기 1줄</option>
              <option value={2}>미리보기 2줄</option>
            </select>

            <div className="h-5 w-px bg-[#E0E0E0] mx-0.5" aria-hidden="true" />

            <button
              type="button"
              aria-pressed={focusedInboxEnabled}
              onClick={() => patchMailPref({ focusedInbox: !focusedInboxEnabled })}
              className={`inline-flex items-center gap-1 px-2.5 py-1.5 text-xs rounded-[5px] whitespace-nowrap ${
                focusedInboxEnabled
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] font-medium'
                  : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              중요 메일함
            </button>

            <button
              type="button"
              aria-pressed={conversationView}
              onClick={() => patchMailPref({ conversationView: !conversationView })}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs rounded-[5px] whitespace-nowrap ${
                conversationView
                  ? 'bg-[#EBF3FC] text-[#0F6CBD] font-medium'
                  : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              <MessageSquare className="w-3.5 h-3.5" />
              대화 스레드
            </button>

            <button
              type="button"
              aria-pressed={excludeAds}
              onClick={() => setExcludeAds((v) => !v)}
              className={`inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs rounded-[5px] whitespace-nowrap ${
                excludeAds
                  ? 'bg-emerald-50 text-emerald-700 font-medium'
                  : 'text-[#242424] hover:bg-[#F5F5F5]'
              }`}
            >
              광고 제외
            </button>

            <div className="inline-flex items-center gap-1 px-1.5 py-1 text-xs text-[#616161]">
              <ArrowDownUp className="w-3.5 h-3.5" />
              <select
                aria-label="리본 메일 정렬"
                value={sortMode}
                onChange={(e) => setSortMode(e.target.value as SortMode)}
                className="text-xs text-[#242424] bg-transparent border border-[#D1D1D1] rounded px-1.5 py-0.5"
              >
                <option value="newest">최신 날짜순</option>
                <option value="oldest">오래된 날짜순</option>
                <option value="sender">보낸 사람순</option>
              </select>
            </div>
          </div>
        )}

        {ribbonTab === 'help' && (
          <div className="flex items-center gap-1.5 shrink-0">
            <a
              href="/manifest.xml?download=1"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[5px] whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5" />
              manifest.xml 다운로드
            </a>
            <button
              type="button"
              onClick={onOpenInstallGuide}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
            >
              <Puzzle className="w-3.5 h-3.5 text-[#0F6CBD]" />
              Outlook 플러그인 설치 센터
            </button>
            <button
              type="button"
              onClick={onOpenConfigPage}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
            >
              <Settings className="w-3.5 h-3.5 text-[#0F6CBD]" />
              EMV 볼트 · 서버 설정
            </button>
            <button
              type="button"
              onClick={onOpenArchivePage}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-[#242424] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
            >
              <BookOpen className="w-3.5 h-3.5 text-[#0F6CBD]" />
              생성된 일지·블로그 기록
            </button>
          </div>
        )}

        {/* Right side of Command Ribbon: Quick Outlook Settings trigger */}
        {onOpenOutlookSettings && (
          <button
            type="button"
            onClick={onOpenOutlookSettings}
            title="Bmail Outlook 상세 설정 (레이아웃 · 서명 · 규칙 · 자동 회신 · SEDMS)"
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium text-[#424242] hover:bg-[#F5F5F5] hover:text-[#0F6CBD] rounded-[5px] whitespace-nowrap shrink-0 border-l border-[#EDEBE9] pl-3"
          >
            <Sliders className="w-3.5 h-3.5 text-[#0F6CBD]" />
            보기 설정
          </button>
        )}
      </div>

      {/* Floating Command Ribbon Popovers: Quick Steps / Categories / Snooze / Sweep */}
      {openToolbarMenu && activeEmail && (
        <div
          onClick={(e) => e.stopPropagation()}
          className="mt-1.5 p-2.5 bg-white border border-[#D1D1D1] rounded-[8px] shadow-[0_8px_20px_rgba(0,0,0,0.12)] flex flex-wrap items-center gap-2 z-30"
        >
          {openToolbarMenu === 'quicksteps' && (
            <>
              <span className="text-[11px] font-bold text-[#0F6CBD] px-1">빠른 단계 실행:</span>
              {outlookPrefs.mail.quickSteps.map((step) => (
                <button
                  key={step.id}
                  type="button"
                  onClick={() => handleRunQuickStep(step, activeEmail)}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1 text-xs font-medium bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] hover:border-[#0F6CBD] rounded-[4px]"
                >
                  <Zap className="w-3 h-3 text-[#0F6CBD]" />
                  <span>{step.name}</span>
                  {step.shortcut && (
                    <span className="text-[10px] font-mono text-[#616161]">({step.shortcut})</span>
                  )}
                </button>
              ))}
            </>
          )}

          {openToolbarMenu === 'category' && (
            <>
              <span className="text-[11px] font-bold text-[#0F6CBD] px-1">색 범주 태그:</span>
              {outlookPrefs.mail.categories.map((cat) => {
                const selected = (mailCategoriesMap[activeEmail.id] || []).includes(cat.id);
                return (
                  <button
                    key={cat.id}
                    type="button"
                    onClick={() => handleToggleMailCategory(activeEmail.id, cat.id)}
                    className={`inline-flex items-center gap-1.5 px-2.5 py-1 text-xs rounded-[4px] border ${
                      selected
                        ? 'bg-[#EBF3FC] border-[#0F6CBD] font-semibold'
                        : 'bg-white border-[#D1D1D1] hover:bg-[#F5F5F5]'
                    }`}
                  >
                    <span
                      className="w-2.5 h-2.5 rounded-full"
                      style={{ backgroundColor: cat.color }}
                    />
                    <span>{cat.name}</span>
                  </button>
                );
              })}
            </>
          )}

          {openToolbarMenu === 'snooze' && (
            <>
              <span className="text-[11px] font-bold text-[#0F6CBD] px-1">다시 알림 예약:</span>
              <button
                type="button"
                onClick={() => handleSnoozeMail(activeEmail.id, '1h')}
                className="px-2.5 py-1 text-xs bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] rounded-[4px]"
              >
                1시간 후
              </button>
              <button
                type="button"
                onClick={() => handleSnoozeMail(activeEmail.id, 'tomorrow9')}
                className="px-2.5 py-1 text-xs bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] rounded-[4px]"
              >
                내일 오전 9:00
              </button>
              <button
                type="button"
                onClick={() => handleSnoozeMail(activeEmail.id, 'nextweek')}
                className="px-2.5 py-1 text-xs bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] rounded-[4px]"
              >
                다음 주 월요일
              </button>
              {mailSnoozeMap[activeEmail.id] && (
                <button
                  type="button"
                  onClick={() => handleSnoozeMail(activeEmail.id, 'clear')}
                  className="px-2.5 py-1 text-xs bg-red-50 text-red-700 border border-red-200 rounded-[4px]"
                >
                  보류 해제
                </button>
              )}
            </>
          )}

          {openToolbarMenu === 'sweep' && (
            <>
              <span className="text-[11px] font-bold text-[#0F6CBD] px-1">
                발신자 일괄 정리 ({activeEmail.fromEmail}):
              </span>
              <button
                type="button"
                onClick={() => handleSweepSender(activeEmail.fromEmail, 'archive', false)}
                className="px-2.5 py-1 text-xs bg-[#FAF9F8] hover:bg-[#EBF3FC] border border-[#D1D1D1] rounded-[4px]"
              >
                받은 모든 메일 보관함 이동
              </button>
              <button
                type="button"
                onClick={() => handleSweepSender(activeEmail.fromEmail, 'archive', true)}
                className="px-2.5 py-1 text-xs bg-[#EBF3FC] text-[#0F6CBD] font-semibold border border-[#B4D6FA] rounded-[4px]"
              >
                모든 메일 + 앞으로 자동 보관 규칙 추가
              </button>
              <button
                type="button"
                onClick={() => handleSweepSender(activeEmail.fromEmail, 'trash', true)}
                className="px-2.5 py-1 text-xs bg-red-50 text-red-700 border border-red-200 rounded-[4px]"
              >
                모든 메일 삭제 + 자동 차단 규칙
              </button>
            </>
          )}

          <button
            type="button"
            onClick={() => setOpenToolbarMenu(null)}
            aria-label="메뉴 닫기"
            className="ml-auto p-1 text-[#616161] hover:text-[#242424] rounded"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}
      </div>

      {/* Authentic Outlook Automatic Replies (Out of Office) Status Bar */}
      {outlookPrefs.mail.outOfOfficeEnabled && (
        <div className="mx-2 mb-2 px-3.5 py-1.5 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-900 flex items-center justify-between gap-2 shrink-0">
          <span className="inline-flex items-center gap-2 truncate">
            <Umbrella className="w-3.5 h-3.5 text-amber-700 shrink-0" />
            <span className="font-semibold">자동 회신(부재중) 켜짐</span>
            {(outlookPrefs.mail.outOfOfficeStart || outlookPrefs.mail.outOfOfficeEnd) && (
              <span className="font-mono text-[11px] text-amber-800">
                ({outlookPrefs.mail.outOfOfficeStart || '지금'} ~ {outlookPrefs.mail.outOfOfficeEnd || '해제 시'})
              </span>
            )}
          </span>
          <button
            type="button"
            onClick={() => patchMailPref({ outOfOfficeEnabled: false })}
            className="text-[11px] font-semibold text-amber-900 hover:underline whitespace-nowrap shrink-0"
          >
            자동 회신 끄기
          </button>
        </div>
      )}

      {graphNotice && (
        <div className="px-3 py-1 bg-[#FAF9F8] border-b border-[#E1DFDD] text-[11px] text-[#424242] flex items-center justify-between shrink-0">
          <span className="truncate">{graphNotice}</span>
        </div>
      )}

      {/* Row 3: Authentic Classic Outlook Flat 1px Divider Multi-Pane Workspace */}
      <div
        style={
          {
            '--pane-folder-w': `${paneWidths.folder}px`,
            '--pane-list-w': `${paneWidths.list}px`,
            '--pane-aidock-w': `${paneWidths.aiDock}px`,
            '--pane-attachdock-w': `${paneWidths.attachDock}px`,
          } as React.CSSProperties
        }
        className="flex-1 flex flex-col lg:flex-row gap-0 px-0 pb-0 bg-white border-t border-[#E1DFDD] min-h-0 overflow-hidden"
      >
        {activeResizePane && (
          <div className="fixed inset-0 z-[999] cursor-col-resize select-none" aria-hidden="true" />
        )}
        {/* Pane 1: Collapsible Outlook Folder Tree (#FAF9F8 Classic Nav Rail) + Bottom Mini-Window Dock */}
        {showFolderPane && (
          <nav
            aria-label="Outlook 폴더"
            className="w-full lg:w-[var(--pane-folder-w)] shrink-0 bg-[#FAF9F8] border-r border-[#E1DFDD] rounded-none flex lg:flex-col justify-between gap-2 overflow-x-auto lg:overflow-hidden select-none"
          >
            <div className="flex lg:flex-col gap-3 lg:overflow-y-auto lg:flex-1 lg:min-h-0 p-1">
            {localFolders.length > 0 && (
              <div className="shrink-0 border-b border-[#E0E0E0] pb-2">
                <OutlookStoreFolderTree
                  folders={localFolders}
                  outlookFavorites={outlookFavorites}
                  favoriteFolderIds={favoriteFolderIds}
                  activeFolderFilter={String(activeFolder)}
                  activeStoreFilter={activeStoreFilter}
                  onSelectFolder={(folderPath, storeName) => {
                    setActiveStoreFilter(storeName || '');
                    setActiveFolder(folderPath);
                  }}
                  onToggleFavoriteFolder={onToggleFavoriteFolder}
                  isSyncingLiveOutlook={isSyncingLiveOutlook}
                  onSyncLiveOutlook={onSyncLiveOutlook}
                  totalEmailsCount={localTotalEmails || activeEmails.length}
                  totalUnreadCount={unreadInboxCount}
                />
              </div>
            )}
            {/* 즐겨찾기 Section */}
            <div className="flex lg:flex-col gap-0.5 shrink-0">
              <button
                type="button"
                onClick={() => setFavoritesOpen((v) => !v)}
                className="hidden lg:flex items-center gap-1.5 px-2 py-1 text-xs font-semibold text-[#424242] hover:bg-black/[0.04] rounded"
              >
                {favoritesOpen ? (
                  <ChevronDown className="w-3.5 h-3.5 text-[#616161]" />
                ) : (
                  <ChevronRight className="w-3.5 h-3.5 text-[#616161]" />
                )}
                즐겨찾기
              </button>
              {favoritesOpen && (
                <div className="flex lg:flex-col gap-0.5 lg:pl-1">
                  {favoriteFolderIds.includes('inbox') &&
                    renderFolderButton(
                      'inbox',
                      '받은 편지함',
                      <Inbox className="w-4 h-4" />,
                      inboxCount,
                      unreadInboxCount
                    )}
                  {favoriteFolderIds.includes('sent') &&
                    renderFolderButton(
                      'sent',
                      '보낸 편지함',
                      <Send className="w-4 h-4" />,
                      sentCount
                    )}
                  {favoriteFolderIds.includes('drafts') &&
                    renderFolderButton(
                      'drafts',
                      '임시 보관함',
                      <FileEdit className="w-4 h-4" />,
                      drafts.length
                    )}
                  {favoriteFolderIds.includes('flagged') &&
                    renderFolderButton(
                      'flagged',
                      '플래그 지정됨',
                      <Flag className="w-4 h-4" />,
                      flaggedCount
                    )}
                  {favoriteFolderIds.includes('all') &&
                    renderFolderButton(
                      'all',
                      '전체 메일함',
                      <Mail className="w-4 h-4" />,
                      activeEmails.length
                    )}
                  {favoriteFolderIds.includes('trash') &&
                    renderFolderButton(
                      'trash',
                      '지운 편지함',
                      <Trash2 className="w-4 h-4" />,
                      trashedEmails.length
                    )}
                  {projects
                    .filter((p) => p.isFavorite)
                    .map((p) => (
                      <button
                        key={p.id}
                        type="button"
                        onClick={() => setActiveFolder(`project:${p.code}`)}
                        onDragOver={(e) => {
                          e.preventDefault();
                          setDropTargetFolder(`project:${p.code}`);
                        }}
                        onDragLeave={() => {
                          if (dropTargetFolder === `project:${p.code}`) setDropTargetFolder(null);
                        }}
                        onDrop={(e) => {
                          e.preventDefault();
                          const id = e.dataTransfer.getData('text/plain') || draggedEmailId;
                          setDropTargetFolder(null);
                          setDraggedEmailId(null);
                          if (id) onAssignEmailToProject?.(p.code, id);
                        }}
                        className={`w-full flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-[5px] text-xs transition-colors whitespace-nowrap ${
                          dropTargetFolder === `project:${p.code}` || activeFolder === `project:${p.code}`
                            ? 'bg-[#CFE4FA] text-[#0F6CBD] font-semibold'
                            : 'text-[#242424] hover:bg-black/[0.04]'
                        }`}
                      >
                        <span className="inline-flex items-center gap-2 min-w-0">
                          <Star className="w-3.5 h-3.5 text-amber-500 shrink-0" fill="currentColor" />
                          <span className="truncate font-mono">{p.code}</span>
                        </span>
                        <span className="text-[10px] font-mono tabular-nums text-[#616161]">
                          {projectMailCounts.get(p.code) ?? 0}
                        </span>
                      </button>
                    ))}
                </div>
              )}
            </div>

            {/* 폴더 Section */}
            <div className="hidden lg:flex flex-col gap-0.5 shrink-0">
              <button
                type="button"
                onClick={() => setFoldersOpen((v) => !v)}
                className="flex items-center gap-1.5 px-2 py-1 text-xs font-semibold text-[#424242] hover:bg-black/[0.04] rounded"
              >
                {foldersOpen ? (
                  <ChevronDown className="w-3.5 h-3.5 text-[#616161]" />
                ) : (
                  <ChevronRight className="w-3.5 h-3.5 text-[#616161]" />
                )}
                <span className="truncate">{pop3ConnectedEmail || microsoftAuth.userEmail || '사서함 폴더'}</span>
              </button>
              {foldersOpen && (
                <div className="flex flex-col gap-0.5 pl-1">
                  {renderFolderButton(
                    'all',
                    '전체 메일함',
                    <Mail className="w-4 h-4" />,
                    activeEmails.length,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('all'),
                    true
                  )}
                  {renderFolderButton(
                    'inbox',
                    '받은 편지함',
                    <Inbox className="w-4 h-4" />,
                    inboxCount,
                    unreadInboxCount,
                    favoritesOpen && favoriteFolderIds.includes('inbox'),
                    true
                  )}
                  {renderFolderButton(
                    'sent',
                    '보낸 편지함',
                    <Send className="w-4 h-4" />,
                    sentCount,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('sent'),
                    true
                  )}
                  {renderFolderButton(
                    'drafts',
                    '임시 보관함',
                    <FileEdit className="w-4 h-4" />,
                    drafts.length,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('drafts'),
                    true
                  )}
                  {renderFolderButton(
                    'flagged',
                    '플래그 지정됨',
                    <Flag className="w-4 h-4" />,
                    flaggedCount,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('flagged'),
                    true
                  )}
                  {renderFolderButton(
                    'archive',
                    '보관함',
                    <Archive className="w-4 h-4" />,
                    activeEmails.filter((m) => m.folderName === '보관함').length,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('archive'),
                    true
                  )}
                  {renderFolderButton(
                    'trash',
                    '지운 편지함',
                    <Trash2 className="w-4 h-4" />,
                    trashedEmails.length,
                    undefined,
                    favoritesOpen && favoriteFolderIds.includes('trash'),
                    true
                  )}
                </div>
              )}
            </div>

            {/* 검색 폴더 (Search Folders) Section */}
            <div className="hidden lg:flex flex-col gap-0.5 pt-2 border-t border-[#E0E0E0] shrink-0">
              <button
                type="button"
                onClick={() => setSearchFoldersOpen((v) => !v)}
                className="flex items-center gap-1.5 px-2 py-1 text-xs font-semibold text-[#424242] hover:bg-black/[0.04] rounded"
              >
                {searchFoldersOpen ? (
                  <ChevronDown className="w-3.5 h-3.5 text-[#616161]" />
                ) : (
                  <ChevronRight className="w-3.5 h-3.5 text-[#616161]" />
                )}
                검색 폴더
              </button>
              {searchFoldersOpen && (
                <div className="flex flex-col gap-0.5 pl-1">
                  {renderFolderButton(
                    'search:unread',
                    '읽지 않은 메일',
                    <Mail className="w-3.5 h-3.5 text-[#0F6CBD]" />,
                    searchFolderEmailIds.get('unread')?.size ?? 0
                  )}
                  {renderFolderButton(
                    'search:flagged',
                    '후속 작업 (플래그)',
                    <Flag className="w-3.5 h-3.5 text-red-600" />,
                    searchFolderEmailIds.get('flagged')?.size ?? 0
                  )}
                  {renderFolderButton(
                    'search:has-attachments',
                    '첨부파일 포함',
                    <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD]" />,
                    searchFolderEmailIds.get('has-attachments')?.size ?? 0
                  )}
                  {renderFolderButton(
                    'search:large-mail',
                    '대용량 메일 (1MB+)',
                    <Search className="w-3.5 h-3.5 text-[#616161]" />,
                    searchFolderEmailIds.get('large-mail')?.size ?? 0
                  )}
                  {renderFolderButton(
                    'search:snoozed',
                    '다시 알림 예약됨',
                    <Clock className="w-3.5 h-3.5 text-amber-600" />,
                    searchFolderEmailIds.get('snoozed')?.size ?? 0
                  )}
                </div>
              )}
            </div>

            {/* 프로젝트 · 파일함 트리 Section */}
            <div className="hidden lg:flex flex-col gap-0.5 pt-2 border-t border-[#E0E0E0] shrink-0">
              <div className="flex items-center justify-between pr-1">
                <button
                  type="button"
                  onClick={() => setProjectsTreeOpen((v) => !v)}
                  className="flex items-center gap-1.5 px-2 py-1 text-xs font-semibold text-[#424242] hover:bg-black/[0.04] rounded"
                >
                  {projectsTreeOpen ? (
                    <ChevronDown className="w-3.5 h-3.5 text-[#616161]" />
                  ) : (
                    <ChevronRight className="w-3.5 h-3.5 text-[#616161]" />
                  )}
                  프로젝트 트리
                </button>
                <button
                  type="button"
                  onClick={onOpenProjectsPage}
                  title="프로젝트 관리 · 도메인 규칙 추가"
                  className="px-1.5 py-0.5 text-[10px] font-medium text-[#0F6CBD] hover:bg-[#EBF3FC] rounded"
                >
                  관리
                </button>
              </div>
              {projectsTreeOpen && (
                <div className="flex flex-col gap-0.5 pl-1">
                  {projects.length === 0 ? (
                    <button
                      type="button"
                      onClick={onOpenProjectsPage}
                      className="w-full flex items-center gap-2 px-2.5 py-1.5 text-[11px] text-[#616161] hover:text-[#0F6CBD] hover:bg-black/[0.04] rounded-[5px]"
                    >
                      <Briefcase className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      프로젝트 추가하기
                    </button>
                  ) : (
                    projects.map((p) => {
                      const folderKey = `project:${p.code}`;
                      const isProjActive = activeFolder === folderKey;
                      const mailCount = projectMailCounts.get(p.code) ?? 0;
                      const fileCount = projectVaultFiles.filter((f) => f.projectCode === p.code).length;
                      return (
                        <button
                          key={p.id}
                          type="button"
                          onClick={() => setActiveFolder(folderKey)}
                          onDragOver={(e) => {
                            e.preventDefault();
                            setDropTargetFolder(folderKey);
                          }}
                          onDragLeave={() => {
                            if (dropTargetFolder === folderKey) setDropTargetFolder(null);
                          }}
                          onDrop={(e) => {
                            e.preventDefault();
                            const id = e.dataTransfer.getData('text/plain') || draggedEmailId;
                            setDropTargetFolder(null);
                            setDraggedEmailId(null);
                            if (id) onAssignEmailToProject?.(p.code, id);
                          }}
                          className={`w-full flex items-center justify-between gap-1.5 px-2.5 py-1.5 rounded-[5px] text-xs transition-colors whitespace-nowrap ${
                            dropTargetFolder === folderKey || isProjActive
                              ? 'bg-[#CFE4FA] text-[#0F6CBD] font-semibold'
                              : 'text-[#242424] hover:bg-black/[0.04]'
                          }`}
                        >
                          <span className="inline-flex items-center gap-2 min-w-0">
                            <span
                              className="w-2 h-2 rounded-full shrink-0"
                              style={{ backgroundColor: p.color || '#0F6CBD' }}
                            />
                            <span className="truncate font-mono">{p.code}</span>
                          </span>
                          <span className="text-[10px] font-mono tabular-nums text-[#616161] shrink-0">
                            {mailCount}
                            {fileCount > 0 ? `·F${fileCount}` : ''}
                          </span>
                        </button>
                      );
                    })
                  )}
                </div>
              )}
            </div>

            {/* Imported PST Data File Folders */}
            {pstFolders.length > 0 && (
              <div className="hidden lg:flex flex-col gap-0.5 pt-2 border-t border-[#E0E0E0] shrink-0">
                <button
                  type="button"
                  onClick={() => setPstTreeOpen((v) => !v)}
                  className="flex items-center gap-1.5 px-2 py-1 text-xs font-semibold text-[#424242] hover:bg-black/[0.04] rounded"
                >
                  {pstTreeOpen ? (
                    <ChevronDown className="w-3.5 h-3.5 text-[#616161]" />
                  ) : (
                    <ChevronRight className="w-3.5 h-3.5 text-[#616161]" />
                  )}
                  <span className="truncate">{connectedStoreName || 'PST 데이터 파일'}</span>
                </button>
                {pstTreeOpen && (
                  <div className="flex flex-col gap-1 pl-1">
                    {pstFolders.length > 10 && (
                      <div className="px-1 pt-0.5 pb-1">
                        <input
                          type="search"
                          value={pstFolderSearch}
                          onChange={(e) => setPstFolderSearch(e.target.value)}
                          placeholder={`PST 폴더 검색 (${pstFolders.length}개)...`}
                          aria-label="PST 폴더 검색"
                          className="w-full px-2 py-1 text-[11px] text-[#242424] bg-white border border-[#D1D1D1] rounded-[4px] focus:outline-none focus:border-[#0F6CBD]"
                        />
                      </div>
                    )}
                    {(() => {
                      const q = pstFolderSearch.trim().toLowerCase();
                      const filtered = q
                        ? pstFolders.filter((f) => f.toLowerCase().includes(q))
                        : pstFolders;
                      const visible = showAllPstFolders || q ? filtered : filtered.slice(0, 15);
                      return (
                        <>
                          <div
                            className={`flex flex-col gap-0.5 ${
                              showAllPstFolders || q ? 'max-h-64 overflow-y-auto pr-0.5' : ''
                            }`}
                          >
                            {visible.map((folderName) => (
                              <button
                                key={folderName}
                                type="button"
                                onClick={() => setActiveFolder(folderName)}
                                title={folderName}
                                className={`w-full flex items-center gap-2 px-2.5 py-1.5 rounded-[5px] text-xs text-left truncate transition-colors ${
                                  activeFolder === folderName
                                    ? 'bg-[#CFE4FA] text-[#0F6CBD] font-semibold'
                                    : 'text-[#424242] hover:bg-black/[0.04]'
                                }`}
                              >
                                <Folder className="w-3.5 h-3.5 shrink-0 text-[#616161]" />
                                <span className="truncate">{folderName}</span>
                              </button>
                            ))}
                            {visible.length === 0 && (
                              <div className="px-2.5 py-1 text-[11px] text-[#707070]">
                                일치하는 PST 폴더가 없습니다.
                              </div>
                            )}
                          </div>
                          {!q && filtered.length > 15 && (
                            <button
                              type="button"
                              onClick={() => setShowAllPstFolders((v) => !v)}
                              className="w-full text-left px-2.5 py-1 text-[11px] font-medium text-[#0F6CBD] hover:bg-[#EBF3FC] rounded-[4px] transition-colors"
                            >
                              {showAllPstFolders
                                ? '상위 15개만 보기'
                                : `전체 ${filtered.length}개 폴더 보기`}
                            </button>
                          )}
                        </>
                      );
                    })()}
                  </div>
                )}
              </div>
            )}
            </div>
          </nav>
        )}
        {showFolderPane &&
          renderVerticalSplitter('folder', 'splitter-folder', '폴더 창 가로 크기 조절')}

        {/* Right Split Container: Optional Project Vault Bar + Message List + Reading Pane */}
        <div className="flex-1 flex flex-col gap-0 min-w-0 min-h-0 overflow-hidden bg-white">
          {activeProject && onAddProjectSubfolder && onDeleteProjectSubfolder && onAddProjectVaultFiles && onRemoveProjectVaultFile && (
            <div className="bg-white border-b border-[#E1DFDD] rounded-none overflow-hidden shrink-0">
              <ProjectVaultBar
                project={activeProject}
                activeSubfolderId={activeProjectSubfolderId}
                subfolders={projectSubfolders}
                files={projectVaultFiles}
                linkedEmailCount={projectMailCounts.get(activeProject.code) ?? 0}
                onSelectSubfolder={setActiveProjectSubfolderId}
                onAddSubfolder={onAddProjectSubfolder}
                onDeleteSubfolder={onDeleteProjectSubfolder}
                onAddFiles={onAddProjectVaultFiles}
                onRemoveFile={onRemoveProjectVaultFile}
                onOpenSedmsHub={onOpenSgwSedmsPage}
              />
            </div>
          )}

          <div
            className={`flex-1 flex gap-0 min-w-0 min-h-0 overflow-hidden ${
              readingPanePosition === 'bottom' ? 'flex-col' : 'flex-col lg:flex-row'
            }`}
          >
        {/* Pane 2: Classic Flat White Outlook Message List */}
        <section
          aria-label="메일 목록"
          className={`bg-white border-r border-[#E1DFDD] rounded-none flex flex-col min-h-[220px] lg:min-h-0 overflow-hidden ${
            !showReadingPane
              ? 'flex-1 min-w-0'
              : readingPanePosition === 'bottom'
                ? 'h-[42%] shrink-0 w-full'
                : 'w-full lg:w-[var(--pane-list-w)] shrink-0'
          }`}
        >
          {/* Outlook List Header: Underlined Filter Tabs + Select / Sort */}
          <div className="px-3.5 pt-2.5 pb-2 border-b border-[#EDEBE9] space-y-2 shrink-0">
            <div className="flex items-center justify-between gap-2">
              <div className="flex items-center gap-2 min-w-0">
                {activeFolder !== 'drafts' && visibleEmails.length > 0 && (
                  <button
                    type="button"
                    onClick={() => {
                      if (options.selectedIds.length === visibleEmails.length) {
                        onChangeOptions({ selectedIds: [] });
                      } else {
                        onChangeOptions({ selectedIds: visibleEmails.map((m) => m.id) });
                      }
                    }}
                    aria-label={
                      options.selectedIds.length === visibleEmails.length
                        ? '전체 선택 해제'
                        : '표시된 메일 전체 선택'
                    }
                    className="text-[#616161] hover:text-[#0F6CBD] shrink-0"
                  >
                    {options.selectedIds.length === visibleEmails.length ? (
                      <CheckSquare className="w-4 h-4 text-[#0F6CBD]" />
                    ) : (
                      <Square className="w-4 h-4" />
                    )}
                  </button>
                )}
                <h2 className="text-sm font-semibold text-[#242424] truncate">
                  {currentFolderTitle}
                </h2>
              </div>

              {activeFolder === 'trash' && trashedEmails.length > 0 ? (
                <button
                  type="button"
                  onClick={onEmptyTrash}
                  className="text-[11px] font-medium text-red-700 hover:underline whitespace-nowrap"
                >
                  폴더 비우기
                </button>
              ) : (
                <span className="text-[11px] font-mono tabular-nums text-[#616161] whitespace-nowrap">
                  {activeFolder === 'drafts' ? `${drafts.length}건` : `${visibleEmails.length}건`}
                </span>
              )}
            </div>

            {activeFolder !== 'drafts' && (
              <div className="flex items-center justify-between gap-2">
                {activeFolder === 'inbox' && focusedInboxEnabled ? (
                  <div className="flex items-center gap-3.5 text-xs" role="tablist" aria-label="중요 및 기타 메일함">
                    {(
                      [
                        { id: 'focused', label: '중요' },
                        { id: 'other', label: '기타' },
                      ] as const
                    ).map((tab) => (
                      <button
                        key={tab.id}
                        type="button"
                        role="tab"
                        aria-selected={focusedTab === tab.id}
                        onClick={() => setFocusedTab(tab.id)}
                        className={`relative pb-1 transition-colors whitespace-nowrap ${
                          focusedTab === tab.id
                            ? 'font-bold text-[#242424]'
                            : 'text-[#616161] hover:text-[#242424]'
                        }`}
                      >
                        {tab.label}
                        {focusedTab === tab.id && (
                          <span
                            style={{ backgroundColor: accent.primary }}
                            className="absolute left-0 right-0 bottom-0 h-[2px] rounded-full"
                          />
                        )}
                      </button>
                    ))}
                  </div>
                ) : (
                  <div className="flex items-center gap-3 text-xs">
                    {(
                      [
                        { id: 'all', label: '전체' },
                        { id: 'unread', label: '읽지 않음' },
                        { id: 'attachments', label: '첨부' },
                      ] as const
                    ).map((tab) => (
                      <button
                        key={tab.id}
                        type="button"
                        onClick={() => setListFilter(tab.id)}
                        className={`relative pb-1 transition-colors whitespace-nowrap ${
                          listFilter === tab.id
                            ? 'font-semibold text-[#242424]'
                            : 'text-[#616161] hover:text-[#242424]'
                        }`}
                      >
                        {tab.label}
                        {listFilter === tab.id && (
                          <span
                            style={{ backgroundColor: accent.primary }}
                            className="absolute left-0 right-0 bottom-0 h-[2px] rounded-full"
                          />
                        )}
                      </button>
                    ))}
                  </div>
                )}

                <div className="flex items-center gap-1">
                  <select
                    aria-label="메일 필터"
                    value={listFilter}
                    onChange={(e) => setListFilter(e.target.value as ListFilterMode)}
                    className="text-[11px] text-[var(--fluent-fg-2)] bg-transparent hover:bg-[var(--fluent-bg-canvas)] rounded-[4px] px-1.5 py-0.5 focus:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                  >
                    <option value="all">필터: 전체</option>
                    <option value="urgent">긴급 ({priorityCounts.urgent})</option>
                    <option value="important">중요 ({priorityCounts.important})</option>
                    <option value="normal">일반 ({priorityCounts.normal})</option>
                    <option value="unread">읽지 않음</option>
                    <option value="attachments">첨부 있음</option>
                  </select>
                  <select
                    aria-label="메일 정렬"
                    value={sortMode}
                    onChange={(e) => setSortMode(e.target.value as SortMode)}
                    className="text-[11px] text-[var(--fluent-fg-2)] bg-transparent hover:bg-[var(--fluent-bg-canvas)] rounded-[4px] px-1.5 py-0.5 focus:outline-none focus-visible:ring-2 focus-visible:ring-[#0F6CBD]"
                  >
                    <option value="newest">최신순</option>
                    <option value="priority">중요도순 (긴급→중요→일반)</option>
                    <option value="oldest">오래된순</option>
                    <option value="sender">발신자순</option>
                  </select>
                </div>
              </div>
            )}

            {(searchQuery.trim() || activeGlobalFilterCount > 0) && (
              <div
                aria-label="적용된 검색 필터"
                className="space-y-1.5 bg-[#EBF3FC] border border-[#CFE4FA] px-2.5 py-2 rounded-[5px]"
              >
                <div className="flex items-center justify-between gap-2 text-[11px] text-[#0F6CBD]">
                  <span className="font-semibold">
                    검색 조건 적용 중
                    {activeGlobalFilterCount > 0 ? ` (${activeGlobalFilterCount}개 필터)` : ''}
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      onSearchQueryChange('');
                      onSearchFiltersChange?.({ ...DEFAULT_GLOBAL_SEARCH_FILTERS });
                    }}
                    className="font-semibold hover:underline shrink-0"
                  >
                    전체 초기화
                  </button>
                </div>
                <div className="flex flex-wrap items-center gap-1">
                  {searchQuery.trim() && (
                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-[#B4D6FA] rounded text-[11px] text-[#242424]">
                      <span>검색: {searchQuery.trim()}</span>
                      <button
                        type="button"
                        onClick={() => onSearchQueryChange('')}
                        aria-label="검색어 조건 제거"
                        className="text-[#616161] hover:text-[#242424]"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  )}
                  {searchFilters.from.trim() && (
                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-[#B4D6FA] rounded text-[11px] text-[#242424]">
                      <span>보낸 사람: {searchFilters.from.trim()}</span>
                      <button
                        type="button"
                        onClick={() =>
                          onSearchFiltersChange?.({ ...searchFilters, from: '' })
                        }
                        aria-label="보낸 사람 조건 제거"
                        className="text-[#616161] hover:text-[#242424]"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  )}
                  {searchFilters.to.trim() && (
                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-[#B4D6FA] rounded text-[11px] text-[#242424]">
                      <span>받는 사람: {searchFilters.to.trim()}</span>
                      <button
                        type="button"
                        onClick={() =>
                          onSearchFiltersChange?.({ ...searchFilters, to: '' })
                        }
                        aria-label="받는 사람 조건 제거"
                        className="text-[#616161] hover:text-[#242424]"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  )}
                  {searchFilters.hasAttachments && (
                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-[#B4D6FA] rounded text-[11px] text-[#242424]">
                      <Paperclip className="w-3 h-3 text-[#0F6CBD]" />
                      <span>첨부 파일 있음</span>
                      <button
                        type="button"
                        onClick={() =>
                          onSearchFiltersChange?.({ ...searchFilters, hasAttachments: false })
                        }
                        aria-label="첨부 파일 조건 제거"
                        className="text-[#616161] hover:text-[#242424]"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  )}
                  {(searchFilters.dateFrom.trim() || searchFilters.dateTo.trim()) && (
                    <span className="inline-flex items-center gap-1 px-1.5 py-0.5 bg-white border border-[#B4D6FA] rounded text-[11px] text-[#242424] font-mono">
                      <span>
                        기간: {searchFilters.dateFrom || '처음'} ~ {searchFilters.dateTo || '현재'}
                      </span>
                      <button
                        type="button"
                        onClick={() =>
                          onSearchFiltersChange?.({
                            ...searchFilters,
                            datePreset: 'any',
                            dateFrom: '',
                            dateTo: '',
                          })
                        }
                        aria-label="기간 조건 제거"
                        className="text-[#616161] hover:text-[#242424]"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </span>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Scrollable Chronological Message List */}
          <div className="flex-1 overflow-y-auto">
            {activeFolder === 'drafts' ? (
              drafts.length === 0 ? (
                <div className="p-8 text-center text-xs text-[#616161]">
                  임시 보관함에 저장된 메일이 없습니다.
                </div>
              ) : (
                <div className="divide-y divide-[#F0F0F0]">
                  {drafts.map((draft) => (
                    <button
                      key={draft.id}
                      type="button"
                      onClick={() => {
                        setActiveDraftId(draft.id);
                        if (readingPanePosition === 'hidden') {
                          patchMailPref({ readingPanePosition: 'right' });
                        }
                        setCenterMode('mail');
                      }}
                      className={`w-full text-left px-3.5 py-3 space-y-1 transition-colors ${
                        activeDraft?.id === draft.id ? 'bg-[#EBF3FC]' : 'hover:bg-[#F5F5F5]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-xs font-semibold text-red-700">[임시 보관]</span>
                        <span className="text-[11px] font-mono tabular-nums text-[#616161]">
                          {draft.updatedAt.slice(11, 16)}
                        </span>
                      </div>
                      <div className="text-xs font-medium text-[#242424] truncate">{draft.subject}</div>
                      <div className="text-[11px] text-[#616161] truncate">
                        받는 사람: {draft.to || '(미지정)'}
                      </div>
                    </button>
                  ))}
                </div>
              )
            ) : visibleEmails.length === 0 ? (
              <div className="p-7 text-center space-y-3">
                <p className="text-xs text-[#616161] leading-relaxed break-keep">
                  {emails.length === 0
                    ? '연동된 실제 메일이 없습니다. 시뮬레이션 데이터는 비워져 있습니다.'
                    : '현재 폴더 및 검색·필터 조건에 맞는 메일이 없습니다.'}
                </p>
                {emails.length === 0 ? (
                  <div className="flex flex-col items-center gap-2">
                    {onOpenPop3Modal && (
                      <button
                        type="button"
                        onClick={onOpenPop3Modal}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[5px]"
                      >
                        <Plug className="w-3.5 h-3.5" />
                        메일플러그 · 회사 POP3 연결
                      </button>
                    )}
                    <button
                      type="button"
                      onClick={onOpenPstModal}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[5px]"
                    >
                      <Database className="w-3.5 h-3.5" />
                      원본 .pst / .eml 불러오기
                    </button>
                  </div>
                ) : (searchQuery.trim() || activeGlobalFilterCount > 0) ? (
                  <button
                    type="button"
                    onClick={() => {
                      onSearchQueryChange('');
                      onSearchFiltersChange?.({ ...DEFAULT_GLOBAL_SEARCH_FILTERS });
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[5px]"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    검색·필터 초기화
                  </button>
                ) : null}
              </div>
            ) : (
              groupedEmails.map((group) => (
                <div key={group.label}>
                  {/* Outlook Date or Pinned Group Header */}
                  <div className="sticky top-0 z-10 px-3.5 py-1.5 bg-[#FAF9F8] border-y border-[#EDEBE9] text-[11px] font-semibold text-[#616161] font-mono tabular-nums flex items-center gap-1.5">
                    {group.isPinnedGroup && <Pin className="w-3 h-3 text-[#0F6CBD]" />}
                    <span>{group.label}</span>
                  </div>
                  <div className="divide-y divide-[#F0F0F0]">
                    {group.items.map((mail) => {
                      const isSelected = options.selectedIds.includes(mail.id);
                      const isActive = activeEmail?.id === mail.id && centerMode === 'mail';
                      const isPinned = pinnedSet.has(mail.id);
                      const isFlagged = mail.isFlagged || Boolean(ruleEvaluationById.get(mail.id)?.shouldFlag);
                      const avatarTone = getAvatarTone(mail.fromEmail || mail.fromName);
                      const initials = getInitials(mail.fromName);
                      const threadKey = mail.conversationKey || normalizeConversationSubject(mail.subject);
                      const threadCount = conversationView ? threadCountByKey.get(threadKey) ?? 1 : 1;
                      const drawingRefs = showDrawingBadgeInList ? drawingRefsById.get(mail.id) : undefined;
                      const priority = priorityById.get(mail.id);
                      const priorityTagClass =
                        priority?.level === 'urgent'
                          ? 'fluent-tag-urgent'
                          : priority?.level === 'important'
                            ? 'fluent-tag-important'
                            : 'fluent-tag-normal';
                      const rowPadClass =
                        density === 'compact' ? 'py-1.5' : density === 'roomy' ? 'py-3.5' : 'py-2.5';

                      return (
                        <div
                          key={mail.id}
                          draggable
                          onDragStart={(e) => {
                            setDraggedEmailId(mail.id);
                            e.dataTransfer.setData('text/plain', mail.id);
                            e.dataTransfer.effectAllowed = 'move';
                          }}
                          onDragEnd={() => {
                            setDraggedEmailId(null);
                            setDropTargetFolder(null);
                          }}
                          onContextMenu={(e) => {
                            e.preventDefault();
                            setActiveEmailId(mail.id);
                            setContextMenu({
                              x: Math.min(e.clientX, window.innerWidth - 220),
                              y: Math.min(e.clientY, window.innerHeight - 320),
                              mail,
                            });
                          }}
                          style={!mail.isRead ? { borderLeftColor: accent.primary } : undefined}
                          className={`group flex items-start gap-2.5 pl-2.5 pr-2.5 border-l-[3px] transition-colors ${
                            !mail.isRead ? '' : 'border-l-transparent'
                          } ${rowPadClass} ${
                            isActive
                              ? 'bg-[#EBF3FC]'
                              : isPinned
                                ? 'bg-[#FFFDF5] hover:bg-[#FEF8E4]'
                                : 'hover:bg-[#F5F5F5]'
                          }`}
                        >
                          {/* Authentic New Outlook Avatar-to-Checkbox Control */}
                          {showSenderAvatars ? (
                            <button
                              type="button"
                              onClick={() => onToggleSelectEmail(mail.id)}
                              aria-label={`'${mail.subject}' 분석 대상 선택`}
                              aria-pressed={isSelected}
                              className="relative mt-0.5 w-8 h-8 shrink-0 flex items-center justify-center rounded-full"
                            >
                              {isSelected ? (
                                <span
                                  style={{ backgroundColor: accent.primary }}
                                  className="w-8 h-8 rounded-full text-white flex items-center justify-center"
                                >
                                  <CheckSquare className="w-4 h-4" />
                                </span>
                              ) : (
                                <>
                                  <span
                                    className={`w-8 h-8 rounded-full text-[11px] font-semibold flex items-center justify-center group-hover:hidden ${avatarTone}`}
                                  >
                                    {initials}
                                  </span>
                                  <span className="hidden group-hover:flex w-8 h-8 rounded-full border border-[#B4D6FA] bg-white text-[#0F6CBD] items-center justify-center">
                                    <Square className="w-4 h-4" />
                                  </span>
                                </>
                              )}
                            </button>
                          ) : (
                            <button
                              type="button"
                              onClick={() => onToggleSelectEmail(mail.id)}
                              aria-label={`'${mail.subject}' 분석 대상 선택`}
                              aria-pressed={isSelected}
                              className={`mt-1 shrink-0 ${
                                isSelected
                                  ? 'text-[#0F6CBD]'
                                  : 'text-[#BDBDBD] opacity-0 group-hover:opacity-100 hover:text-[#0F6CBD]'
                              }`}
                            >
                              {isSelected ? (
                                <CheckSquare className="w-4 h-4" />
                              ) : (
                                <Square className="w-4 h-4" />
                              )}
                            </button>
                          )}

                          <button
                            type="button"
                            onClick={() => handleSelectMailRow(mail)}
                            className="flex-1 min-w-0 text-left space-y-0.5"
                          >
                            <div className="flex items-center justify-between gap-2">
                              <span className="inline-flex items-center gap-1.5 min-w-0">
                                {!mail.isRead && (
                                  <span
                                    aria-label="읽지 않음"
                                    style={{ backgroundColor: accent.primary }}
                                    className="w-2 h-2 rounded-full shrink-0"
                                  />
                                )}
                                <span
                                  className={`text-xs truncate ${
                                    !mail.isRead
                                      ? 'font-bold text-[var(--fluent-fg-1)]'
                                      : 'font-medium text-[var(--fluent-fg-1)]'
                                  }`}
                                >
                                  {mail.fromName}
                                </span>
                                {priority && (
                                  <span
                                    title={`자동 중요도 분류 (${priority.score}점): ${priority.reasons.join(' · ')}`}
                                    className={`inline-flex items-center px-1.5 py-[1px] rounded-[4px] text-[10px] font-semibold leading-tight shrink-0 ${priorityTagClass}`}
                                  >
                                    {priority.label}
                                  </span>
                                )}
                              </span>
                              <span className="text-[11px] font-mono tabular-nums text-[var(--fluent-fg-3)] shrink-0 group-hover:hidden">
                                {mail.timeLabel}
                              </span>
                            </div>

                            <p
                              className={`text-xs truncate flex items-center gap-1.5 ${
                                !mail.isRead
                                  ? 'font-semibold text-[var(--fluent-brand)]'
                                  : 'font-normal text-[var(--fluent-fg-1)]'
                              }`}
                            >
                              <span className="truncate">{mail.subject}</span>
                              {priority?.matchedProjectCode && (
                                <span className="font-mono text-[10px] font-medium text-[var(--fluent-fg-3)] shrink-0">
                                  [{priority.matchedProjectCode}]
                                </span>
                              )}
                              {threadCount > 1 && (
                                <span className="font-mono text-[10px] font-bold text-[var(--fluent-brand)] shrink-0">
                                  ({threadCount})
                                </span>
                              )}
                            </p>

                            {previewLines > 0 && density !== 'compact' && (
                              <div className="flex items-start gap-1.5 text-[11px] text-[#616161] min-w-0">
                                <span
                                  className={
                                    previewLines === 2 ? 'line-clamp-2 leading-snug' : 'truncate'
                                  }
                                >
                                  {mail.snippet}
                                </span>
                                {mail.hasAttachment && (
                                  <Paperclip className="w-3 h-3 text-[#616161] shrink-0 mt-0.5" />
                                )}
                              </div>
                            )}

                            {drawingRefs && drawingRefs.length > 0 && (
                              <div className="text-[10px] font-mono tabular-nums text-[#0F6CBD] truncate">
                                도면 {drawingRefs.slice(0, 2).map((d) => d.raw).join(' · ')}
                              </div>
                            )}

                            {((mailCategoriesMap[mail.id]?.length || 0) > 0 || mailSnoozeMap[mail.id]) && (
                              <div className="flex flex-wrap items-center gap-1 pt-0.5">
                                {(mailCategoriesMap[mail.id] || []).map((catId) => {
                                  const cat = outlookPrefs.mail.categories.find((c) => c.id === catId);
                                  if (!cat) return null;
                                  return (
                                    <span
                                      key={cat.id}
                                      className="inline-flex items-center gap-1 px-1.5 py-[1px] rounded-[3px] text-[10px] font-semibold border"
                                      style={{
                                        borderColor: `${cat.color}44`,
                                        backgroundColor: `${cat.color}14`,
                                        color: cat.color,
                                      }}
                                    >
                                      <span
                                        className="w-1.5 h-1.5 rounded-full"
                                        style={{ backgroundColor: cat.color }}
                                      />
                                      {cat.name}
                                    </span>
                                  );
                                })}
                                {mailSnoozeMap[mail.id] && (
                                  <span className="inline-flex items-center gap-0.5 px-1.5 py-[1px] rounded-[3px] bg-amber-50 border border-amber-200 text-[10px] text-amber-800">
                                    <Clock className="w-2.5 h-2.5" />
                                    보류됨
                                  </span>
                                )}
                              </div>
                            )}
                          </button>

                          {/* Authentic Microsoft New Outlook 4-Action Row Hover Bar (Delete / Read / Flag / Pin) */}
                          <div className="flex items-center gap-0.5 shrink-0 pt-0.5">
                            <button
                              type="button"
                              onClick={() =>
                                activeFolder === 'trash'
                                  ? onRestoreEmail(mail.id)
                                  : handleDeleteOrArchiveWithAdvance(mail.id, 'delete')
                              }
                              title={activeFolder === 'trash' ? '복원' : '삭제'}
                              aria-label={activeFolder === 'trash' ? `'${mail.subject}' 복원` : `'${mail.subject}' 삭제`}
                              className="p-1 rounded text-[#757575] opacity-0 group-hover:opacity-100 hover:bg-black/[0.06] hover:text-red-600 transition-opacity"
                            >
                              {activeFolder === 'trash' ? (
                                <RotateCcw className="w-3.5 h-3.5" />
                              ) : (
                                <Trash2 className="w-3.5 h-3.5" />
                              )}
                            </button>
                            <button
                              type="button"
                              onClick={() => onToggleReadEmail(mail.id)}
                              title={mail.isRead ? '읽지 않음으로 표시' : '읽음으로 표시'}
                              aria-label={`'${mail.subject}' 읽음 상태 전환`}
                              className="p-1 rounded text-[#757575] opacity-0 group-hover:opacity-100 hover:bg-black/[0.06] hover:text-[#0F6CBD] transition-opacity"
                            >
                              <MailOpen className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              onClick={() => onToggleFlagEmail(mail.id)}
                              title={isFlagged ? '플래그 해제' : '플래그 지정'}
                              aria-label={`'${mail.subject}' 플래그 전환`}
                              className={`p-1 rounded hover:bg-black/[0.06] transition-opacity ${
                                isFlagged
                                  ? 'text-red-600 opacity-100'
                                  : 'text-[#757575] opacity-0 group-hover:opacity-100 hover:text-red-600'
                              }`}
                            >
                              <Flag className="w-3.5 h-3.5" />
                            </button>
                            {onTogglePinEmail && (
                              <button
                                type="button"
                                onClick={() => onTogglePinEmail(mail.id)}
                                title={isPinned ? '고정 해제' : '상단 고정'}
                                aria-label={`'${mail.subject}' 고정 전환`}
                                className={`p-1 rounded hover:bg-black/[0.06] transition-opacity ${
                                  isPinned
                                    ? 'text-[#0F6CBD] opacity-100'
                                    : 'text-[#757575] opacity-0 group-hover:opacity-100 hover:text-[#0F6CBD]'
                                }`}
                              >
                                <Pin className="w-3.5 h-3.5" />
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))
            )}
          </div>
        </section>
        {showReadingPane &&
          readingPanePosition !== 'bottom' &&
          renderVerticalSplitter('list', 'splitter-list', '메일 목록 · 읽기 창 가로 크기 조절')}

        {/* Pane 3: Classic Flat White Outlook Reading Pane OR AI Result Viewer */}
        {showReadingPane && (
          <section
            aria-label={centerMode === 'result' ? 'AI 분석 및 생성 문서' : '메일 읽기 창'}
            className="flex-1 min-w-0 bg-white rounded-none flex flex-col overflow-hidden"
          >
            {centerMode === 'result' ? (
              <Suspense fallback={<div className="p-6 text-xs text-[#616161]">AI 분석 결과 불러오는 중...</div>}>
                <AnalysisResultPanel
                  result={analysisResult}
                  isAnalyzing={isAnalyzing}
                  agentSteps={agentSteps}
                  emails={emails.filter((m) => !deletedIds.includes(m.id))}
                  projects={projects}
                  onSaveDraftToOutlook={handleSaveWorklogDraftAndPopup}
                  onOpenPstModal={onOpenPstModal}
                />
              </Suspense>
            ) : activeFolder === 'drafts' && activeDraft ? (
              <div className="flex-1 overflow-y-auto p-6 space-y-5">
                <div className="flex items-center justify-between pb-4 border-b border-[#EDEBE9]">
                  <div>
                    <span className="text-xs font-semibold text-red-700">임시 보관함 초안</span>
                    <h2 className="text-lg font-bold text-[#242424] mt-0.5">{activeDraft.subject}</h2>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() =>
                        openComposePopup(activeDraft.kind, {
                          to: activeDraft.to,
                          cc: activeDraft.cc,
                          subject: activeDraft.subject,
                          body: activeDraft.body,
                        })
                      }
                      className="px-3 py-1.5 text-xs font-semibold text-white bg-[#0F6CBD] hover:bg-[#115EA3] rounded-[5px]"
                    >
                      계속 작성 / 보내기
                    </button>
                    <button
                      type="button"
                      onClick={() => onDeleteDraft(activeDraft.id)}
                      className="px-3 py-1.5 text-xs font-medium text-red-700 bg-white border border-[#D1D1D1] hover:bg-red-50 rounded-[5px]"
                    >
                      초안 삭제
                    </button>
                  </div>
                </div>
                <div className="text-xs text-[#616161] space-y-1">
                  <div>받는 사람: {activeDraft.to || '(미지정)'}</div>
                  {activeDraft.cc && <div>참조: {activeDraft.cc}</div>}
                </div>
                <div className="text-sm text-[#242424] leading-[1.75] whitespace-pre-line max-w-2xl">
                  {activeDraft.body}
                </div>
              </div>
            ) : activeEmail ? (
              <div className="flex-1 flex flex-col min-h-0 overflow-hidden">
                {/* Outlook Subject Header Strip */}
                <div
                  className={`${
                    outlookPrefs.mail.minimalHeader ? 'px-5 py-2.5' : 'px-6 py-4'
                  } border-b border-[#EDEBE9] shrink-0`}
                >
                  {!outlookPrefs.mail.minimalHeader && (
                    <div className="text-[11px] text-[#616161] mb-1">
                      {activeEmail.folder === 'inbox' ? '받은 편지함' : '보낸 편지함'}
                      {activeEmail.folderName ? ` · ${activeEmail.folderName}` : ''}
                    </div>
                  )}
                  <h2
                    className={`${
                      outlookPrefs.mail.minimalHeader ? 'text-[15px]' : 'text-[17px]'
                    } font-bold text-[#242424] leading-snug`}
                  >
                    {activeEmail.subject}
                  </h2>
                </div>

                {/* Scrollable Message Thread Body */}
                <div className="flex-1 overflow-y-auto px-6 py-5 space-y-5">
                  <div className="flex items-start justify-between gap-4">
                    <div className="flex items-start gap-3 min-w-0">
                      <div
                        className={`w-10 h-10 rounded-full text-xs font-semibold flex items-center justify-center shrink-0 ${getAvatarTone(
                          activeEmail.fromEmail || activeEmail.fromName
                        )}`}
                      >
                        {getInitials(activeEmail.fromName)}
                      </div>
                      <div className="min-w-0 space-y-0.5 text-xs">
                        <div className="flex flex-wrap items-baseline gap-1.5">
                          <span className="font-semibold text-[#242424] text-sm">
                            {activeEmail.fromName}
                          </span>
                          <span className="font-mono text-[#616161] text-[11px]">
                            &lt;{activeEmail.fromEmail}&gt;
                          </span>
                        </div>
                        <div className="text-[#616161] truncate">
                          받는 사람: <span className="text-[#242424]">{activeEmail.to.join(', ')}</span>
                        </div>
                        <div className="text-[11px] font-mono tabular-nums text-[#616161]">
                          {activeEmail.date} {activeEmail.timeLabel}
                        </div>
                      </div>
                    </div>

                    {/* Outlook Compact Message Action Buttons */}
                    <div className="flex items-center gap-1 shrink-0">
                      <button
                        type="button"
                        onClick={() => setIsMiniAssistOpen(true)}
                        title="번역 · 요약 · 회신 미니창 열기"
                        aria-label="번역 · 요약 · 회신 미니창 열기"
                        className="inline-flex items-center gap-1 px-2 py-1.5 text-[11px] font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[5px] whitespace-nowrap"
                      >
                        <Languages className="w-3.5 h-3.5" />
                        번역·요약
                      </button>
                      <button
                        type="button"
                        onClick={() => handleReplyAction(activeEmail, 'reply')}
                        title="회신"
                        aria-label="회신"
                        className="p-2 text-[#424242] hover:bg-[#F5F5F5] rounded-[5px]"
                      >
                        <Reply className="w-4 h-4 text-[#0F6CBD]" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleReplyAction(activeEmail, 'replyAll')}
                        title="전체 회신"
                        aria-label="전체 회신"
                        className="p-2 text-[#424242] hover:bg-[#F5F5F5] rounded-[5px]"
                      >
                        <ReplyAll className="w-4 h-4 text-[#0F6CBD]" />
                      </button>
                      <button
                        type="button"
                        onClick={() => handleReplyAction(activeEmail, 'forward')}
                        title="전달"
                        aria-label="전달"
                        className="p-2 text-[#424242] hover:bg-[#F5F5F5] rounded-[5px]"
                      >
                        <Forward className="w-4 h-4 text-[#0F6CBD]" />
                      </button>
                      {inlineAiAnalysis && (
                        <button
                          type="button"
                          onClick={() => setInlineAiExpanded((v) => !v)}
                          title="이 메일의 즉시 요약 및 원클릭 회신 초안 펼치기/접기"
                          aria-pressed={inlineAiExpanded}
                          className={`inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium rounded-[4px] border transition-colors ${
                            inlineAiExpanded
                              ? 'bg-[#EDEBE9] border-[#D1D1D1] text-[#0F6CBD]'
                              : 'bg-white border-transparent text-[#616161] hover:text-[#242424] hover:bg-[#F3F2F1]'
                          }`}
                        >
                          <Sparkles className="w-3 h-3 text-[#0F6CBD]" />
                          <span>{inlineAiExpanded ? 'AI 요약 숨기기' : 'AI 요약 보기'}</span>
                        </button>
                      )}
                      {onShareMailToSgwChat && (
                        <button
                          type="button"
                          onClick={() => onShareMailToSgwChat(activeEmail)}
                          title="이 메일을 SGW 사내 채팅으로 공유"
                          aria-label="SGW 채팅 공유"
                          className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-medium text-[#424242] hover:bg-[#F3F2F1] rounded-[4px]"
                        >
                          <MessageSquare className="w-3.5 h-3.5 text-[#0F6CBD]" />
                          <span>채팅 공유</span>
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Collapsible Inline Bmail AI Copilot Summary & 3-Option Instant Reply Strip (Hidden by default for clean Classic Outlook reading flow) */}
                  {inlineAiAnalysis && inlineAiExpanded && (
                    <div className="pl-0 sm:pl-[52px]">
                      <div className="rounded-lg border border-[#D8E6F5] bg-[#F8FBFF] p-3.5 space-y-2.5">
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2 min-w-0 text-xs">
                            <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD] shrink-0" />
                            <span className="font-semibold text-[#242424] whitespace-nowrap">
                              Bmail AI 즉시 요약
                            </span>
                            <span className="text-[#616161] truncate">
                              {inlineAiAnalysis.intentLabel} · 긴급도{' '}
                              {inlineAiAnalysis.urgency === 'urgent'
                                ? '높음'
                                : inlineAiAnalysis.urgency === 'high'
                                  ? '우선'
                                  : inlineAiAnalysis.urgency === 'low'
                                    ? '낮음'
                                    : '보통'}
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => setInlineAiExpanded(false)}
                            className="text-[11px] font-medium text-[#0F6CBD] hover:underline whitespace-nowrap shrink-0"
                          >
                            접기
                          </button>
                        </div>

                        <div className="space-y-2.5 text-xs">
                            <p className="text-[#242424] leading-relaxed">
                              {inlineAiAnalysis.oneLineSummary}
                            </p>

                            {(inlineAiAnalysis.extractedDeadlines.length > 0 ||
                              inlineAiAnalysis.extractedAmounts.length > 0 ||
                              inlineAiAnalysis.extractedDrawingNos.length > 0) && (
                              <div className="text-[11px] text-[#424242] font-mono tabular-nums space-x-3">
                                {inlineAiAnalysis.extractedDeadlines.length > 0 && (
                                  <span>기한: {inlineAiAnalysis.extractedDeadlines.join(', ')}</span>
                                )}
                                {inlineAiAnalysis.extractedAmounts.length > 0 && (
                                  <span>금액: {inlineAiAnalysis.extractedAmounts.join(', ')}</span>
                                )}
                                {inlineAiAnalysis.extractedDrawingNos.length > 0 && (
                                  <span>도면: {inlineAiAnalysis.extractedDrawingNos.join(', ')}</span>
                                )}
                              </div>
                            )}

                            {inlineAiAnalysis.riskFlags.length > 0 && (
                              <div className="flex items-center gap-1.5 text-[11px] text-amber-800">
                                <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                                <span>{inlineAiAnalysis.riskFlags.join(' · ')}</span>
                              </div>
                            )}

                            {inlineReplyOptions.length > 0 && (
                              <div className="pt-2 border-t border-[#E1ECF8] flex flex-wrap items-center gap-1.5">
                                <span className="text-[11px] font-medium text-[#616161] mr-1 whitespace-nowrap">
                                  원클릭 회신 초안:
                                </span>
                                {inlineReplyOptions.map((opt) => (
                                  <button
                                    key={opt.id}
                                    type="button"
                                    onClick={() =>
                                      openComposePopup('reply', {
                                        to: activeEmail.fromEmail,
                                        subject: opt.subject,
                                        body: opt.body,
                                      })
                                    }
                                    className="px-2.5 py-1 text-[11px] font-medium text-[#0F6CBD] bg-white border border-[#C7E0F4] hover:bg-[#EBF3FC] rounded-[5px] transition-colors whitespace-nowrap"
                                  >
                                    {opt.label}
                                  </button>
                                ))}
                              </div>
                            )}
                          </div>
                      </div>
                    </div>
                  )}

                  {(() => {
                    const effectiveAtts =
                      richEmailRender && richEmailRender.id === activeEmail.id && richEmailRender.attachments.length > 0
                        ? richEmailRender.attachments
                        : activeEmail.attachments || [];
                    const bodyHtmlHint =
                      richEmailRender && richEmailRender.id === activeEmail.id
                        ? richEmailRender.bodyHtml
                        : undefined;
                    const realAtts = effectiveAtts
                      .map((att, idx) => {
                        const m = String(att.id || '').match(/-att-(\d+)$/);
                        const realIdx = m ? Number(m[1]) : idx;
                        const isInlineSig = isInlineBusinessCardImage(att, bodyHtmlHint);
                        return { att, realIdx, isInlineSig };
                      })
                      .filter((x) => !x.isInlineSig);

                    if (realAtts.length === 0) return null;
                    const isLocalPst =
                      localDbAvailable &&
                      activeEmail.source === 'pst-import' &&
                      activeEmail.id.startsWith('pst-');

                    return (
                      <div className="pl-0 sm:pl-[52px] flex flex-wrap items-center gap-2">
                        {realAtts.map(({ att, realIdx }) => (
                          <div
                            key={att.id || realIdx}
                            className="inline-flex items-center rounded-[5px] border border-[#D1D1D1] bg-[#FAF9F8] overflow-hidden text-xs text-[#242424]"
                          >
                            <button
                              type="button"
                              onClick={() => {
                                setShowAttachSplitDock(true);
                                setSelectedPreviewAttachmentIdx(realIdx);
                              }}
                              title="우측 첨부 미리보기 · 분석 도크에서 바로 보기"
                              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 hover:bg-[#EBF3FC] transition-colors"
                            >
                              <Paperclip className="w-3.5 h-3.5 text-[#0F6CBD]" />
                              <span className="font-medium">{att.filename}</span>
                              <span className="text-[10px] font-mono text-[#616161] uppercase">
                                {att.kind}
                              </span>
                            </button>
                            {isLocalPst && (
                              <a
                                href={`/api/local/pst/attachment?emailId=${encodeURIComponent(activeEmail.id)}&idx=${realIdx}`}
                                download={att.filename}
                                title={`${att.filename} 원본 다운로드`}
                                className="px-2 py-1.5 border-l border-[#D1D1D1] text-[#0F6CBD] hover:bg-[#EBF3FC] inline-flex items-center"
                              >
                                <Download className="w-3.5 h-3.5" />
                              </a>
                            )}
                          </div>
                        ))}
                      </div>
                    );
                  })()}

                  {richEmailRender &&
                  richEmailRender.id === activeEmail.id &&
                  richEmailRender.bodyHtml ? (
                    <div
                      className={`pl-0 sm:pl-[52px] text-[#242424] leading-[1.7] max-w-4xl overflow-x-auto [&_img]:max-w-full [&_img]:h-auto [&_img]:rounded [&_img]:my-2 [&_table]:border-collapse ${
                        outlookPrefs.mail.readingFontSize === 'sm'
                          ? 'text-xs'
                          : outlookPrefs.mail.readingFontSize === 'lg'
                            ? 'text-[15px]'
                            : 'text-sm'
                      }`}
                      dangerouslySetInnerHTML={{ __html: richEmailRender.bodyHtml }}
                    />
                  ) : (
                    <div
                      className={`pl-0 sm:pl-[52px] text-[#242424] leading-[1.75] whitespace-pre-line max-w-3xl ${
                        outlookPrefs.mail.readingFontSize === 'sm'
                          ? 'text-xs'
                          : outlookPrefs.mail.readingFontSize === 'lg'
                            ? 'text-[15px]'
                            : 'text-sm'
                      }`}
                    >
                      {activeEmail.body}
                    </div>
                  )}

                  {/* Authentic Outlook Bottom Quick-Reply Action Strip */}
                  <div className="pl-0 sm:pl-[52px] pt-4 border-t border-[#EDEBE9] flex flex-wrap items-center gap-2">
                    <button
                      type="button"
                      onClick={() => handleReplyAction(activeEmail, 'reply')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
                    >
                      <Reply className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      회신
                    </button>
                    <button
                      type="button"
                      onClick={() => handleReplyAction(activeEmail, 'replyAll')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
                    >
                      <ReplyAll className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      전체 회신
                    </button>
                    <button
                      type="button"
                      onClick={() => handleReplyAction(activeEmail, 'forward')}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px]"
                    >
                      <Forward className="w-3.5 h-3.5 text-[#0F6CBD]" />
                      전달
                    </button>
                    <button
                      type="button"
                      onClick={() => setIsMiniAssistOpen(true)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] border border-[#C7E0F4] hover:bg-[#DCEBFA] rounded-[5px]"
                    >
                      <Languages className="w-3.5 h-3.5" />
                      번역 · 회신 미니창
                    </button>

                    {analysisResult && (
                      <button
                        type="button"
                        onClick={() => setCenterMode('result')}
                        className="ml-auto inline-flex items-center gap-1.5 text-xs font-semibold text-[#0F6CBD] hover:underline"
                      >
                        <FileSpreadsheet className="w-3.5 h-3.5" />
                        생성된 AI 일지·블로그 보기
                      </button>
                    )}
                  </div>
                </div>
              </div>
            ) : (
              /* Zero-Simulation Authentic Outlook Empty Reading Pane */
              <div className="flex-1 flex flex-col items-center justify-center p-8 text-center">
                <div className="max-w-xl w-full space-y-6">
                  <div className="space-y-2">
                    <img
                      src="/assets/icon-64.png"
                      alt="Bmail"
                      width={44}
                      height={44}
                      className="w-11 h-11 rounded-xl mx-auto"
                    />
                    <h2 className="text-base font-bold text-[#242424] break-keep">
                      Bmail — 실제 메일함을 연결해 일일 브리핑과 번역·회신을 시작하세요
                    </h2>
                    <p className="text-xs text-[#616161] leading-relaxed break-keep">
                      샘플·더미 데이터는 0건입니다. 로컬 Outlook{' '}
                      <code className="font-mono">.pst/.ost/.eml/.mbox</code> 파일을 열거나 Microsoft 365 · SGW 사내 메일을 연동하면 오른쪽 읽기 창에 매일 메일 요약, 3옵션 즉시 회신, 번역 미니창, 음성·회의록 및 SEDMS 도면 연결이 즉시 반영됩니다.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5 text-left items-stretch">
                    <button
                      type="button"
                      onClick={onOpenPstModal}
                      className="p-3.5 rounded-lg border border-[#E0E0E0] bg-[#FAF9F8] hover:border-[#0F6CBD] hover:bg-[#EBF3FC]/50 transition-colors space-y-1.5 flex flex-col justify-start min-w-0"
                    >
                      <Database className="w-5 h-5 text-[#0F6CBD] shrink-0" />
                      <div className="text-xs font-bold text-[#242424] break-keep">
                        1. 원본 PST · EML 열기
                      </div>
                      <p className="text-[11px] text-[#616161] leading-relaxed break-keep">
                        Outlook 폴더·프로젝트·일정 자동 인수
                      </p>
                    </button>

                    <button
                      type="button"
                      onClick={microsoftAuth.configured ? onConnectMicrosoft : onOpenInstallGuide}
                      className="p-3.5 rounded-lg border border-[#E0E0E0] bg-[#FAF9F8] hover:border-[#0F6CBD] hover:bg-[#EBF3FC]/50 transition-colors space-y-1.5 flex flex-col justify-start min-w-0"
                    >
                      <CloudDownload className="w-5 h-5 text-[#0F6CBD] shrink-0" />
                      <div className="text-xs font-bold text-[#242424] break-keep">
                        2. Microsoft 365 연동
                      </div>
                      <p className="text-[11px] text-[#616161] leading-relaxed break-keep">
                        Graph OAuth 실시간 메일 동기화
                      </p>
                    </button>

                    <button
                      type="button"
                      onClick={onOpenSgwSedmsPage || onOpenInstallGuide}
                      className="p-3.5 rounded-lg border border-[#E0E0E0] bg-[#FAF9F8] hover:border-[#0F6CBD] hover:bg-[#EBF3FC]/50 transition-colors space-y-1.5 flex flex-col justify-start min-w-0"
                    >
                      <Building2 className="w-5 h-5 text-[#0F6CBD] shrink-0" />
                      <div className="text-xs font-bold text-[#242424] break-keep">
                        3. SGW · SEDMS 연동
                      </div>
                      <p className="text-[11px] text-[#616161] leading-relaxed break-keep">
                        사내 그룹웨어 SSO · 도면 자동 연결
                      </p>
                    </button>
                  </div>

                  <div className="flex flex-wrap items-center justify-center gap-2 pt-1">
                    <button
                      type="button"
                      onClick={() => setIsMiniAssistOpen(true)}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#0F6CBD] bg-[#EBF3FC] hover:bg-[#DCEBFA] rounded-[5px] whitespace-nowrap"
                    >
                      <Languages className="w-3.5 h-3.5" />
                      영업·외국어 번역 미니창 열기
                    </button>
                    {onOpenAiBriefPage && (
                      <button
                        type="button"
                        onClick={onOpenAiBriefPage}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
                      >
                        <Sparkles className="w-3.5 h-3.5 text-[#0F6CBD]" />
                        AI 모닝 브리핑 · Ask AI
                      </button>
                    )}
                    {onOpenVoicePage && (
                      <button
                        type="button"
                        onClick={onOpenVoicePage}
                        className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-[#242424] bg-white border border-[#D1D1D1] hover:bg-[#F5F5F5] rounded-[5px] whitespace-nowrap"
                      >
                        <Mic className="w-3.5 h-3.5 text-[#0F6CBD]" />
                        음성 · 회의록 분석 (Voice)
                      </button>
                    )}
                  </div>
                </div>
              </div>
            )}
          </section>
        )}

        {/* Right Split 1: Vertical AI Analysis & Reply Dock (Horizontally Resizable) */}
        {showReadingPane && centerMode === 'mail' && showAiSplitDock && (
          <>
            {renderVerticalSplitter('aiDock', 'splitter-ai-dock', '읽기 창 · AI 도크 가로 크기 조절')}
            <Suspense fallback={null}>
              <div className="w-full lg:w-[var(--pane-aidock-w)] border-l border-[#E1DFDD] overflow-hidden flex min-h-0 shrink-0">
                <MailAiDockPanel
                  email={activeEmail}
                  onInsertReplyDraft={(draft) =>
                    openComposePopup('reply', {
                      to: draft.to,
                      subject: draft.subject,
                      body: draft.body,
                    })
                  }
                  onAddCalendarEvent={onAddCalendarEvent}
                  onClose={() => togglePinAiSplitDock(false)}
                />
              </div>
            </Suspense>
          </>
        )}

        {/* Right Split 2: Vertical Attachment Preview, AI Doc Analysis & Folder Batch Exporter Dock (Horizontally Resizable) */}
        {showReadingPane && centerMode === 'mail' && showAttachSplitDock && (
          <>
            {renderVerticalSplitter('attachDock', 'splitter-attach-dock', '첨부·폴더저장 도크 가로 크기 조절')}
            <Suspense fallback={null}>
              <div className="w-full lg:w-[var(--pane-attachdock-w)] border-l border-[#E1DFDD] overflow-hidden flex min-h-0 shrink-0">
                <AttachmentPreviewDock
                  email={activeEmail}
                  selectedAttachmentIdx={selectedPreviewAttachmentIdx}
                  onSelectAttachmentIdx={setSelectedPreviewAttachmentIdx}
                  activeFolderFilter={String(activeFolder)}
                  activeStoreFilter={activeStoreFilter}
                  projects={projects}
                  onClose={() => setShowAttachSplitDock(false)}
                />
              </div>
            </Suspense>
          </>
        )}
          </div>
        </div>

        {/* Pane 4: Floating White Rounded Docked Outlook Add-in Taskpane */}
        {isTaskpaneOpen && (
          <OutlookTaskpane
            emails={activeEmails}
            options={options}
            providers={providers}
            isAnalyzing={isAnalyzing}
            agentSteps={agentSteps}
            error={analysisError}
            hostInfo={hostInfo}
            isPinned={isTaskpanePinned}
            supportsPinning={supportsPinning}
            accentColor={accent.primary}
            onTogglePin={onToggleTaskpanePin}
            onChangeOptions={onChangeOptions}
            onRunAnalysis={handleRunAndShowResult}
            onClose={onCloseTaskpane}
          />
        )}
      </div>

      {/* Authentic Outlook Pop-out Rich HTML Compose / Reply / Forward Window */}
      {composePopup.open && (
        <Suspense fallback={null}>
          <OutlookComposeWindow
            mode="modal"
            draft={composePopup}
            signatures={localSignatures}
            contacts={localContacts}
            outlookPrefs={outlookPrefs}
            pop3ConnectedEmail={pop3ConnectedEmail}
            onSendSmtp={onSendViaSmtp}
            onSaveDraft={(draft) => {
              onSaveDraft(draft);
              setComposePopup((prev) => ({ ...prev, open: false }));
              setActiveFolder('drafts');
            }}
            onPopOutNewWindow={(snapshot) => {
              if (openPopOutComposeWindow(snapshot)) {
                setComposePopup((prev) => ({ ...prev, open: false }));
              }
            }}
            onClose={() => setComposePopup((prev) => ({ ...prev, open: false }))}
          />
        </Suspense>
      )}

      {/* Authentic Outlook Right-Click Context Menu */}
      {contextMenu && (
        <div
          role="menu"
          aria-label="메일 우클릭 메뉴"
          style={{ left: contextMenu.x, top: contextMenu.y }}
          onClick={(e) => e.stopPropagation()}
          className="fixed z-50 w-52 bg-white border border-[#E0E0E0] rounded-lg shadow-xl py-1.5 text-xs text-[#242424]"
        >
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              handleReplyAction(contextMenu.mail, 'reply');
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <Reply className="w-3.5 h-3.5" />
            회신
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              handleReplyAction(contextMenu.mail, 'replyAll');
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <ReplyAll className="w-3.5 h-3.5" />
            전체 회신
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              handleReplyAction(contextMenu.mail, 'forward');
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <Forward className="w-3.5 h-3.5" />
            전달
          </button>
          <div className="my-1 border-t border-[#EDEBE9]" />
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onToggleReadEmail(contextMenu.mail.id);
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <MailOpen className="w-3.5 h-3.5" />
            {contextMenu.mail.isRead ? '읽지 않음으로 표시' : '읽음으로 표시'}
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onToggleFlagEmail(contextMenu.mail.id);
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <Flag className="w-3.5 h-3.5" />
            {contextMenu.mail.isFlagged ? '플래그 해제' : '플래그 지정'}
          </button>
          {onTogglePinEmail && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                onTogglePinEmail(contextMenu.mail.id);
                setContextMenu(null);
              }}
              className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
            >
              <Pin className="w-3.5 h-3.5" />
              {pinnedSet.has(contextMenu.mail.id) ? '고정 해제' : '상단 고정'}
            </button>
          )}
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onMoveEmailToFolder(
                contextMenu.mail.id,
                contextMenu.mail.folder === 'inbox' ? 'sent' : 'inbox'
              );
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
          >
            <Folder className="w-3.5 h-3.5" />
            {contextMenu.mail.folder === 'inbox' ? '보낸 편지함으로 이동' : '받은 편지함으로 이동'}
          </button>
          {activeProject && onRemoveEmailFromProject && (
            <button
              type="button"
              role="menuitem"
              onClick={() => {
                onRemoveEmailFromProject(activeProject.code, contextMenu.mail.id);
                setContextMenu(null);
              }}
              className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] hover:text-[#0F6CBD] text-left"
            >
              <Briefcase className="w-3.5 h-3.5" />
              프로젝트({activeProject.code})에서 제외
            </button>
          )}
          <div className="my-1 border-t border-[#EDEBE9]" />
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onChangeOptions({ scope: 'selected', selectedIds: [contextMenu.mail.id] });
              handleRunAndShowResult();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] text-[#0F6CBD] font-semibold text-left"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            AI로 이 메일 즉시 분석
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              onOpenStudioPage();
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] text-[#0F6CBD] text-left"
          >
            <Mic className="w-3.5 h-3.5" />
            문서·음성 AI 스튜디오 열기
          </button>
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              handleTriggerMailToast(contextMenu.mail);
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-[#EBF3FC] text-[#0F6CBD] text-left"
          >
            <Bell className="w-3.5 h-3.5" />
            트레이 1-클릭 알림 띄우기
          </button>
          <div className="my-1 border-t border-[#EDEBE9]" />
          <button
            type="button"
            role="menuitem"
            onClick={() => {
              if (activeFolder === 'trash') {
                onRestoreEmail(contextMenu.mail.id);
              } else {
                onDeleteEmail(contextMenu.mail.id);
              }
              setContextMenu(null);
            }}
            className="w-full flex items-center gap-2 px-3 py-1.5 hover:bg-red-50 text-red-700 text-left"
          >
            <Trash2 className="w-3.5 h-3.5" />
            {activeFolder === 'trash' ? '메일 복원' : '삭제 (지운 편지함)'}
          </button>
        </div>
      )}
      {/* Bottom-Left Floating Cancellable Delayed Send Queue Bar */}
      {delayedQueue.length > 0 && (
        <div
          role="region"
          aria-label="지연 발송 취소 대기열"
          className="fixed bottom-3 left-4 z-40 w-[380px] max-w-[calc(100vw-2rem)] space-y-2"
        >
          {delayedQueue.map((item) => {
            const remSec = getRemainingSeconds(item, queueNowMs);
            return (
              <div
                key={item.id}
                className="bg-[#242424] text-white border border-white/15 rounded-lg px-3.5 py-2.5 shadow-2xl flex items-center justify-between gap-3 text-xs"
              >
                <div className="min-w-0 flex-1 space-y-0.5">
                  <div className="flex items-center gap-1.5 font-semibold">
                    <Clock className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                    <span className="truncate">{item.subject}</span>
                  </div>
                  <div className="text-[11px] text-white/75 truncate">
                    {item.status === 'sending'
                      ? 'SMTP 발송 진행 중...'
                      : item.status === 'failed'
                        ? `오류: ${item.errorMessage || '실패'}`
                        : `수신: ${item.to || '(임시저장)'} · ${formatCountdownLabel(remSec)} 후 자동 발송`}
                  </div>
                </div>
                {item.status === 'queued' && (
                  <div className="flex items-center gap-1.5 shrink-0">
                    <button
                      type="button"
                      onClick={() => {
                        const res = cancelDelayedSendItem(delayedQueue, item.id);
                        setDelayedQueue(res.remainingQueue);
                        if (res.cancelled) {
                          openComposePopup(res.cancelled.kind, {
                            to: res.cancelled.to,
                            cc: res.cancelled.cc,
                            subject: res.cancelled.subject,
                            body: res.cancelled.body,
                          });
                        }
                      }}
                      className="inline-flex items-center gap-1 px-2 py-1 text-[11px] font-bold bg-amber-400 hover:bg-amber-300 text-slate-900 rounded"
                    >
                      <Undo2 className="w-3 h-3" />
                      발송 취소
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setDelayedQueue((prev) =>
                          prev.map((q) => (q.id === item.id ? { ...q, status: 'sending' } : q))
                        );
                        void dispatchMailPayloadNow(item).finally(() => {
                          setDelayedQueue((prev) => prev.filter((q) => q.id !== item.id));
                        });
                      }}
                      className="px-2 py-1 text-[11px] font-semibold bg-white/15 hover:bg-white/25 rounded"
                    >
                      즉시
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* Web-Mode Compact 1-Click Mail Notification Toast */}
      {inAppMailToast && (
        <div
          role="status"
          aria-label="빠른 메일 1-클릭 알림"
          className="fixed bottom-14 right-3 z-40 w-[340px] max-w-[calc(100vw-2rem)] bg-white border border-[#0F6CBD] rounded-lg shadow-2xl p-3 space-y-2 text-xs"
        >
          <div className="flex items-center justify-between gap-2">
            <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-[#EBF3FC] text-[#0F6CBD]">
              Bmail 1-클릭 알림
            </span>
            <button
              type="button"
              onClick={() => setInAppMailToast(null)}
              aria-label="알림 닫기"
              className="text-[#707070] hover:text-[#242424]"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
          <div className="space-y-0.5">
            <div className="font-bold text-[#242424] truncate">{inAppMailToast.fromName}</div>
            <div className="font-semibold text-[#0F6CBD] truncate">{inAppMailToast.subject}</div>
          </div>
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={() => {
                if (!inAppMailToast.isRead) onToggleReadEmail(inAppMailToast.id);
                setInAppMailToast(null);
              }}
              className="flex-1 py-1.5 px-2 text-[11px] font-semibold bg-[#FAF9F8] hover:bg-[#EDEBE9] border border-[#D1D1D1] rounded"
            >
              ✓ 즉시 읽음
            </button>
            <button
              type="button"
              onClick={() => {
                const target = inAppMailToast;
                setInAppMailToast(null);
                setActiveEmailId(target.id);
                handleReplyAction(target, 'reply');
              }}
              className="flex-1 py-1.5 px-2 text-[11px] font-semibold bg-[#0F6CBD] hover:bg-[#115EA3] text-white rounded"
            >
              ↩ 즉시 회신
            </button>
          </div>
        </div>
      )}

      {/* Bottom-Right Compact Floating Sales / Foreign-Language Translation & Reply Mini-Window */}
      {centerMode === 'mail' && (
        <UnreadBriefingBanner
          open={showUnreadBriefing}
          onOpenChange={setShowUnreadBriefing}
          emails={emails}
          projects={projects}
          projectMailOverrides={projectMailOverrides}
          serverBriefing={unreadBriefing}
          onOpenMail={(emailId) => void handleBriefingOpenOrReply(emailId, 'open')}
          onQuickReplyMail={(emailId) => void handleBriefingOpenOrReply(emailId, 'reply')}
        />
      )}
      {!isMiniAssistOpen && (
        <button
          type="button"
          onClick={() => setIsMiniAssistOpen(true)}
          title="영업·외국어 번역 및 회신 초안 미니창 열기 (우측 하단 컴팩트 고정)"
          className="fixed bottom-3 right-3 z-30 px-3 py-2 rounded-full bg-[#0F6CBD] hover:bg-[#115EA3] text-white text-xs font-bold shadow-lg border border-white/25 flex items-center gap-1.5 transition-all"
        >
          <Languages className="w-3.5 h-3.5 shrink-0" />
          <span>번역·요약 미니창</span>
        </button>
      )}
      <Suspense fallback={null}>
        <MailMiniAssistWindow
          isOpen={isMiniAssistOpen}
          onClose={() => setIsMiniAssistOpen(false)}
          activeEmail={activeEmail}
          onSaveReplyDraft={(draft) => {
            onSaveDraft({
              to: draft.to,
              cc: '',
              subject: draft.subject,
              body: draft.body,
              kind: 'reply',
            });
            setActiveFolder('drafts');
          }}
          onOpenComposeWithDraft={(draft) => {
            setIsMiniAssistOpen(false);
            openComposePopup('reply', {
              to: draft.to,
              subject: draft.subject,
              body: draft.body,
            });
          }}
        />
      </Suspense>
    </div>
  );
};
