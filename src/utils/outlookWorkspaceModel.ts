import type { AttachmentKind, MailAttachmentMeta, MailItem } from '../types/mail';
import { detectProjectCode, type ProjectDefinition } from './deterministicAnalyzer';

// ============================================================================
// 1. Authentic Microsoft New Outlook User Preferences & Rules Model
// ============================================================================

export type OutlookDensity = 'roomy' | 'comfortable' | 'compact';
export type OutlookReadingPanePosition = 'right' | 'bottom' | 'hidden';
export type OutlookAccentTheme = 'blue' | 'navy' | 'teal' | 'slate';
export type OutlookPreviewLines = 0 | 1 | 2;
export type OutlookFontSize = 'sm' | 'md' | 'lg';
export type OutlookUndoSendSeconds = 0 | 5 | 10;

export type MailRuleConditionField = 'from' | 'subject' | 'drawing' | 'hasAttachment';
export type MailRuleActionType = 'assign-project' | 'flag' | 'mark-read' | 'archive' | 'move-trash';

export interface OutlookCategoryTag {
  id: string;
  name: string;
  color: string;
}

export type OutlookQuickStepAction =
  | 'mark-read'
  | 'flag'
  | 'pin'
  | 'archive'
  | 'move-trash'
  | 'assign-category'
  | 'assign-project';

export interface OutlookQuickStep {
  id: string;
  name: string;
  shortcut?: string;
  actions: OutlookQuickStepAction[];
  categoryId?: string;
  projectCode?: string;
}

export type OutlookSearchFolderId =
  | 'unread'
  | 'flagged'
  | 'has-attachments'
  | 'large-mail'
  | 'snoozed';

export interface OutlookTaskItem {
  id: string;
  title: string;
  dueDate?: string;
  priority: 'low' | 'normal' | 'high';
  completed: boolean;
  linkedMailId?: string;
  createdAt: string;
}

export const DEFAULT_OUTLOOK_CATEGORIES: OutlookCategoryTag[] = [
  { id: 'cat-urgent', name: '긴급', color: '#D13438' },
  { id: 'cat-drawing', name: '도면검토', color: '#0F6CBD' },
  { id: 'cat-quote', name: '견적·발주', color: '#107C10' },
  { id: 'cat-pending', name: '승인대기', color: '#CA5010' },
  { id: 'cat-meeting', name: '일정·회의', color: '#8764B8' },
  { id: 'cat-ref', name: '참고자료', color: '#8A8886' },
];

export const DEFAULT_OUTLOOK_QUICK_STEPS: OutlookQuickStep[] = [
  {
    id: 'qs-done',
    name: '완료 처리',
    shortcut: 'Ctrl+Shift+1',
    actions: ['mark-read', 'archive'],
  },
  {
    id: 'qs-priority',
    name: '중요 플래그 + 고정',
    shortcut: 'Ctrl+Shift+2',
    actions: ['flag', 'pin'],
  },
  {
    id: 'qs-drawing',
    name: '도면 분류 + 읽음',
    shortcut: 'Ctrl+Shift+3',
    actions: ['mark-read', 'assign-category'],
    categoryId: 'cat-drawing',
  },
  {
    id: 'qs-archive',
    name: '즉시 보관함 이동',
    shortcut: 'Ctrl+Shift+4',
    actions: ['archive'],
  },
];

export interface OutlookMailRule {
  id: string;
  name: string;
  enabled: boolean;
  conditionField: MailRuleConditionField;
  conditionValue: string;
  actionType: MailRuleActionType;
  actionTarget?: string;
}

export interface OutlookPreferences {
  general: {
    language: 'ko-KR' | 'en-US';
    timezone: string;
    density: OutlookDensity;
    accentTheme: OutlookAccentTheme;
    desktopNotifications: boolean;
    notificationSound: boolean;
    showUnreadBadge: boolean;
  };
  mail: {
    focusedInbox: boolean;
    readingPanePosition: OutlookReadingPanePosition;
    wideReadingPane: boolean;
    readingFontSize: OutlookFontSize;
    previewLines: OutlookPreviewLines;
    showSenderAvatars: boolean;
    conversationView: boolean;
    newestOnTop: boolean;
    defaultReplyAction: 'reply' | 'replyAll';
    undoSendSeconds: OutlookUndoSendSeconds;
    composeFontFamily: 'sans' | 'serif' | 'mono';
    composeFontSizePt: 10 | 11 | 12;
    signature: string;
    signatureHtml: string;
    signatureName: string;
    autoSignatureNew: boolean;
    autoSignatureReply: boolean;
    outOfOfficeEnabled: boolean;
    outOfOfficeStart: string;
    outOfOfficeEnd: string;
    outOfOfficeMessage: string;
    blockedSenders: string[];
    safeSenders: string[];
    rules: OutlookMailRule[];
    singleLineRibbon: boolean;
    minimalHeader: boolean;
    markAsReadMode: 'immediate' | 'delay' | 'manual';
    markAsReadDelaySec: number;
    afterDeleteAction: 'next' | 'prev' | 'list';
    requestReadReceipt: boolean;
    defaultImportance: 'low' | 'normal' | 'high';
    categories: OutlookCategoryTag[];
    quickSteps: OutlookQuickStep[];
  };
  calendar: {
    weekStartsOn: 'sunday' | 'monday';
    defaultMeetingMinutes: 30 | 60;
    workStart: string;
    workEnd: string;
  };
  sedms: {
    autoExtractDrawings: boolean;
    autoIndexProjectFiles: boolean;
    showDrawingBadgeInList: boolean;
    defaultTenantId: string;
  };
}

export const DEFAULT_OUTLOOK_PREFERENCES: OutlookPreferences = {
  general: {
    language: 'ko-KR',
    timezone: 'Asia/Seoul (UTC+09:00)',
    density: 'comfortable',
    accentTheme: 'blue',
    desktopNotifications: true,
    notificationSound: false,
    showUnreadBadge: true,
  },
  mail: {
    focusedInbox: true,
    readingPanePosition: 'right',
    wideReadingPane: true,
    readingFontSize: 'md',
    previewLines: 1,
    showSenderAvatars: false,
    conversationView: false,
    newestOnTop: true,
    defaultReplyAction: 'reply',
    undoSendSeconds: 5,
    composeFontFamily: 'sans',
    composeFontSizePt: 11,
    signature: '',
    signatureHtml: '',
    signatureName: '',
    autoSignatureNew: true,
    autoSignatureReply: true,
    outOfOfficeEnabled: false,
    outOfOfficeStart: '',
    outOfOfficeEnd: '',
    outOfOfficeMessage: '현재 부재중으로 메일 확인이 지연될 수 있습니다. 긴급 건은 사내 그룹웨어로 연락 부탁드립니다.',
    blockedSenders: [],
    safeSenders: [],
    rules: [],
    singleLineRibbon: false,
    minimalHeader: false,
    markAsReadMode: 'immediate',
    markAsReadDelaySec: 3,
    afterDeleteAction: 'next',
    requestReadReceipt: false,
    defaultImportance: 'normal',
    categories: structuredClone(DEFAULT_OUTLOOK_CATEGORIES),
    quickSteps: structuredClone(DEFAULT_OUTLOOK_QUICK_STEPS),
  },
  calendar: {
    weekStartsOn: 'monday',
    defaultMeetingMinutes: 30,
    workStart: '09:00',
    workEnd: '18:00',
  },
  sedms: {
    autoExtractDrawings: true,
    autoIndexProjectFiles: true,
    showDrawingBadgeInList: true,
    defaultTenantId: 'shipyard-main',
  },
};

export const ACCENT_THEME_COLORS: Record<
  OutlookAccentTheme,
  { primary: string; hover: string; soft: string; border: string; label: string }
> = {
  blue: {
    primary: '#0F6CBD',
    hover: '#115EA3',
    soft: '#EBF3FC',
    border: '#B4D6FA',
    label: 'Outlook 블루',
  },
  navy: {
    primary: '#1E3A5F',
    hover: '#162C49',
    soft: '#EDF2F8',
    border: '#B8CCE0',
    label: '마린 네이비',
  },
  teal: {
    primary: '#036973',
    hover: '#02555D',
    soft: '#E6F5F6',
    border: '#A7DADF',
    label: '조선 틸',
  },
  slate: {
    primary: '#334155',
    hover: '#1E293B',
    soft: '#F1F5F9',
    border: '#CBD5E1',
    label: '슬레이트 그레이',
  },
};

export function sanitizeOutlookPreferences(raw: unknown): OutlookPreferences {
  if (!raw || typeof raw !== 'object') {
    return structuredClone(DEFAULT_OUTLOOK_PREFERENCES);
  }
  const obj = raw as Record<string, unknown>;
  const general = (obj.general && typeof obj.general === 'object' ? obj.general : {}) as Record<string, unknown>;
  const mail = (obj.mail && typeof obj.mail === 'object' ? obj.mail : {}) as Record<string, unknown>;
  const calendar = (obj.calendar && typeof obj.calendar === 'object' ? obj.calendar : {}) as Record<string, unknown>;
  const sedms = (obj.sedms && typeof obj.sedms === 'object' ? obj.sedms : {}) as Record<string, unknown>;

  const validDensities: OutlookDensity[] = ['roomy', 'comfortable', 'compact'];
  const validPositions: OutlookReadingPanePosition[] = ['right', 'bottom', 'hidden'];
  const validAccents: OutlookAccentTheme[] = ['blue', 'navy', 'teal', 'slate'];
  const validPreview: OutlookPreviewLines[] = [0, 1, 2];
  const validFontSizes: OutlookFontSize[] = ['sm', 'md', 'lg'];
  const validUndo: OutlookUndoSendSeconds[] = [0, 5, 10];

  const sanitizeStringList = (val: unknown): string[] =>
    Array.isArray(val)
      ? Array.from(new Set(val.map((v) => String(v || '').trim().toLowerCase()).filter(Boolean)))
      : [];

  const sanitizeRules = (val: unknown): OutlookMailRule[] => {
    if (!Array.isArray(val)) return [];
    return val
      .filter((r): r is Record<string, unknown> => Boolean(r && typeof r === 'object'))
      .map((r, idx) => ({
        id: String(r.id || `rule-${idx}`),
        name: String(r.name || '새 규칙').slice(0, 80),
        enabled: Boolean(r.enabled),
        conditionField: (['from', 'subject', 'drawing', 'hasAttachment'].includes(String(r.conditionField))
          ? r.conditionField
          : 'subject') as MailRuleConditionField,
        conditionValue: String(r.conditionValue || '').trim(),
        actionType: (['assign-project', 'flag', 'mark-read', 'archive', 'move-trash'].includes(String(r.actionType))
          ? r.actionType
          : 'flag') as MailRuleActionType,
        actionTarget: r.actionTarget ? String(r.actionTarget).trim() : undefined,
      }));
  };

  const sanitizeCategories = (val: unknown): OutlookCategoryTag[] => {
    if (!Array.isArray(val)) return structuredClone(DEFAULT_OUTLOOK_CATEGORIES);
    return val
      .filter((c): c is Record<string, unknown> => Boolean(c && typeof c === 'object'))
      .map((c, idx) => ({
        id: String(c.id || `cat-${idx}`),
        name: String(c.name || '범주').trim().slice(0, 40) || '범주',
        color:
          typeof c.color === 'string' && /^#[0-9a-fA-F]{6}$/.test(c.color.trim())
            ? c.color.trim()
            : '#0F6CBD',
      }));
  };

  const validQuickStepActions = new Set<string>([
    'mark-read',
    'flag',
    'pin',
    'archive',
    'move-trash',
    'assign-category',
    'assign-project',
  ]);

  const sanitizeQuickSteps = (val: unknown): OutlookQuickStep[] => {
    if (!Array.isArray(val)) return structuredClone(DEFAULT_OUTLOOK_QUICK_STEPS);
    return val
      .filter((q): q is Record<string, unknown> => Boolean(q && typeof q === 'object'))
      .map((q, idx) => ({
        id: String(q.id || `qs-${idx}`),
        name: String(q.name || '빠른 단계').trim().slice(0, 60) || '빠른 단계',
        shortcut: typeof q.shortcut === 'string' ? q.shortcut.trim().slice(0, 30) : undefined,
        actions: Array.isArray(q.actions)
          ? (q.actions
              .map(String)
              .filter((a) => validQuickStepActions.has(a)) as OutlookQuickStepAction[])
          : ['mark-read'],
        categoryId: typeof q.categoryId === 'string' ? q.categoryId.trim() : undefined,
        projectCode: typeof q.projectCode === 'string' ? q.projectCode.trim() : undefined,
      }));
  };

  return {
    general: {
      language: general.language === 'en-US' ? 'en-US' : 'ko-KR',
      timezone: typeof general.timezone === 'string' && general.timezone.trim()
        ? general.timezone.trim()
        : DEFAULT_OUTLOOK_PREFERENCES.general.timezone,
      density: validDensities.includes(general.density as OutlookDensity)
        ? (general.density as OutlookDensity)
        : DEFAULT_OUTLOOK_PREFERENCES.general.density,
      accentTheme: validAccents.includes(general.accentTheme as OutlookAccentTheme)
        ? (general.accentTheme as OutlookAccentTheme)
        : DEFAULT_OUTLOOK_PREFERENCES.general.accentTheme,
      desktopNotifications:
        typeof general.desktopNotifications === 'boolean'
          ? general.desktopNotifications
          : DEFAULT_OUTLOOK_PREFERENCES.general.desktopNotifications,
      notificationSound:
        typeof general.notificationSound === 'boolean'
          ? general.notificationSound
          : DEFAULT_OUTLOOK_PREFERENCES.general.notificationSound,
      showUnreadBadge:
        typeof general.showUnreadBadge === 'boolean'
          ? general.showUnreadBadge
          : DEFAULT_OUTLOOK_PREFERENCES.general.showUnreadBadge,
    },
    mail: {
      focusedInbox:
        typeof mail.focusedInbox === 'boolean'
          ? mail.focusedInbox
          : DEFAULT_OUTLOOK_PREFERENCES.mail.focusedInbox,
      readingPanePosition: validPositions.includes(mail.readingPanePosition as OutlookReadingPanePosition)
        ? (mail.readingPanePosition as OutlookReadingPanePosition)
        : DEFAULT_OUTLOOK_PREFERENCES.mail.readingPanePosition,
      wideReadingPane:
        typeof mail.wideReadingPane === 'boolean'
          ? mail.wideReadingPane
          : DEFAULT_OUTLOOK_PREFERENCES.mail.wideReadingPane,
      readingFontSize: validFontSizes.includes(mail.readingFontSize as OutlookFontSize)
        ? (mail.readingFontSize as OutlookFontSize)
        : DEFAULT_OUTLOOK_PREFERENCES.mail.readingFontSize,
      previewLines: validPreview.includes(mail.previewLines as OutlookPreviewLines)
        ? (mail.previewLines as OutlookPreviewLines)
        : DEFAULT_OUTLOOK_PREFERENCES.mail.previewLines,
      showSenderAvatars:
        typeof mail.showSenderAvatars === 'boolean'
          ? mail.showSenderAvatars
          : DEFAULT_OUTLOOK_PREFERENCES.mail.showSenderAvatars,
      conversationView:
        typeof mail.conversationView === 'boolean'
          ? mail.conversationView
          : DEFAULT_OUTLOOK_PREFERENCES.mail.conversationView,
      newestOnTop:
        typeof mail.newestOnTop === 'boolean'
          ? mail.newestOnTop
          : DEFAULT_OUTLOOK_PREFERENCES.mail.newestOnTop,
      defaultReplyAction: mail.defaultReplyAction === 'replyAll' ? 'replyAll' : 'reply',
      undoSendSeconds: validUndo.includes(mail.undoSendSeconds as OutlookUndoSendSeconds)
        ? (mail.undoSendSeconds as OutlookUndoSendSeconds)
        : DEFAULT_OUTLOOK_PREFERENCES.mail.undoSendSeconds,
      composeFontFamily:
        mail.composeFontFamily === 'serif' || mail.composeFontFamily === 'mono'
          ? mail.composeFontFamily
          : 'sans',
      composeFontSizePt:
        mail.composeFontSizePt === 10 || mail.composeFontSizePt === 12 ? mail.composeFontSizePt : 11,
      signature: typeof mail.signature === 'string' ? mail.signature.slice(0, 1500) : '',
      signatureHtml:
        typeof mail.signatureHtml === 'string' ? mail.signatureHtml.slice(0, 150_000) : '',
      signatureName: typeof mail.signatureName === 'string' ? mail.signatureName.slice(0, 120) : '',
      autoSignatureNew:
        typeof mail.autoSignatureNew === 'boolean'
          ? mail.autoSignatureNew
          : DEFAULT_OUTLOOK_PREFERENCES.mail.autoSignatureNew,
      autoSignatureReply:
        typeof mail.autoSignatureReply === 'boolean'
          ? mail.autoSignatureReply
          : DEFAULT_OUTLOOK_PREFERENCES.mail.autoSignatureReply,
      outOfOfficeEnabled: Boolean(mail.outOfOfficeEnabled),
      outOfOfficeStart: typeof mail.outOfOfficeStart === 'string' ? mail.outOfOfficeStart : '',
      outOfOfficeEnd: typeof mail.outOfOfficeEnd === 'string' ? mail.outOfOfficeEnd : '',
      outOfOfficeMessage:
        typeof mail.outOfOfficeMessage === 'string'
          ? mail.outOfOfficeMessage.slice(0, 1500)
          : DEFAULT_OUTLOOK_PREFERENCES.mail.outOfOfficeMessage,
      blockedSenders: sanitizeStringList(mail.blockedSenders),
      safeSenders: sanitizeStringList(mail.safeSenders),
      rules: sanitizeRules(mail.rules),
      singleLineRibbon:
        typeof mail.singleLineRibbon === 'boolean'
          ? mail.singleLineRibbon
          : DEFAULT_OUTLOOK_PREFERENCES.mail.singleLineRibbon,
      minimalHeader:
        typeof mail.minimalHeader === 'boolean'
          ? mail.minimalHeader
          : DEFAULT_OUTLOOK_PREFERENCES.mail.minimalHeader,
      markAsReadMode:
        mail.markAsReadMode === 'delay' || mail.markAsReadMode === 'manual'
          ? mail.markAsReadMode
          : 'immediate',
      markAsReadDelaySec:
        typeof mail.markAsReadDelaySec === 'number' && Number.isFinite(mail.markAsReadDelaySec)
          ? Math.max(0, Math.min(30, Math.round(mail.markAsReadDelaySec)))
          : DEFAULT_OUTLOOK_PREFERENCES.mail.markAsReadDelaySec,
      afterDeleteAction:
        mail.afterDeleteAction === 'prev' || mail.afterDeleteAction === 'list'
          ? mail.afterDeleteAction
          : 'next',
      requestReadReceipt: Boolean(mail.requestReadReceipt),
      defaultImportance:
        mail.defaultImportance === 'low' || mail.defaultImportance === 'high'
          ? mail.defaultImportance
          : 'normal',
      categories: sanitizeCategories(mail.categories),
      quickSteps: sanitizeQuickSteps(mail.quickSteps),
    },
    calendar: {
      weekStartsOn: calendar.weekStartsOn === 'sunday' ? 'sunday' : 'monday',
      defaultMeetingMinutes: calendar.defaultMeetingMinutes === 60 ? 60 : 30,
      workStart: typeof calendar.workStart === 'string' && /^\d{2}:\d{2}$/.test(calendar.workStart)
        ? calendar.workStart
        : '09:00',
      workEnd: typeof calendar.workEnd === 'string' && /^\d{2}:\d{2}$/.test(calendar.workEnd)
        ? calendar.workEnd
        : '18:00',
    },
    sedms: {
      autoExtractDrawings:
        typeof sedms.autoExtractDrawings === 'boolean' ? sedms.autoExtractDrawings : true,
      autoIndexProjectFiles:
        typeof sedms.autoIndexProjectFiles === 'boolean' ? sedms.autoIndexProjectFiles : true,
      showDrawingBadgeInList:
        typeof sedms.showDrawingBadgeInList === 'boolean' ? sedms.showDrawingBadgeInList : true,
      defaultTenantId:
        typeof sedms.defaultTenantId === 'string' && sedms.defaultTenantId.trim()
          ? sedms.defaultTenantId.trim()
          : 'shipyard-main',
    },
  };
}

// ============================================================================
// 2. Shipbuilding / SEDMS Drawing Reference Extractor (Browser-Safe)
// ============================================================================

export interface ShipbuildingDrawingRef {
  raw: string;
  canonKey: string;
  discipline: '선체/구조' | '배관' | '전장' | '기계/의장' | '도면/일반';
  revision?: string;
}

const DRAWING_PATTERNS = [
  /\b(?:DWG|DRW|BLK|HULL|STR|PIP|ELC|MCH|OUT)[-_ ]?[A-Z0-9]{2,6}(?:[-_](?!REV\.?[A-Z0-9]+\b|R\d+\b)[A-Z0-9]{1,6}){0,3}(?:[-_ ]?(?:REV\.?[A-Z0-9]+|R\d+))?\b/gi,
  /\b[HSEPMO][-_]?\d{3,5}[-_][A-Z0-9]{1,5}(?:[-_](?:R\d+|REV\.?[A-Z0-9]+))?\b/gi,
];

export function normalizeDrawingKeyBrowser(input: string): string {
  return String(input || '')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function classifyDiscipline(raw: string): ShipbuildingDrawingRef['discipline'] {
  const upper = raw.toUpperCase();
  if (/^(HULL|BLK|STR|H[-_]?|S[-_]?)/.test(upper)) return '선체/구조';
  if (/^(PIP|P[-_]?)/.test(upper)) return '배관';
  if (/^(ELC|E[-_]?)/.test(upper)) return '전장';
  if (/^(MCH|OUT|M[-_]?|O[-_]?)/.test(upper)) return '기계/의장';
  return '도면/일반';
}

export function extractShipbuildingDrawingRefs(text: string): ShipbuildingDrawingRef[] {
  const source = String(text || '')
    // Strip email addresses and URLs so usernames like `elcon0955@naver.com` are never mistaken for ELC drawing codes
    .replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g, ' ')
    .replace(/https?:\/\/\S+/gi, ' ');
  if (!source.trim()) return [];
  const seen = new Set<string>();
  const refs: ShipbuildingDrawingRef[] = [];

  for (const regex of DRAWING_PATTERNS) {
    regex.lastIndex = 0;
    let match: RegExpExecArray | null = regex.exec(source);
    while (match) {
      const raw = match[0].trim();
      const revMatch = raw.match(/(?:REV\.?|R)([A-Z0-9]+)$/i);
      const baseRaw = revMatch ? raw.slice(0, revMatch.index).replace(/[-_\s.]+$/, '') : raw;
      const canonKey = normalizeDrawingKeyBrowser(baseRaw || raw);
      if (canonKey.length >= 4 && /\d/.test(canonKey) && !seen.has(canonKey)) {
        seen.add(canonKey);
        refs.push({
          raw,
          canonKey,
          discipline: classifyDiscipline(raw),
          revision: revMatch ? revMatch[0].toUpperCase() : undefined,
        });
      }
      match = regex.exec(source);
    }
  }
  return refs;
}

// ============================================================================
// 3. Favorites, Hierarchical Project Tree & Project File Vault (Add / Remove)
// ============================================================================

export const DEFAULT_FAVORITE_FOLDER_IDS = ['inbox', 'sent', 'drafts'];

export interface ProjectTreeSubfolder {
  id: string;
  projectCode: string;
  parentId: string | null;
  name: string;
  createdAt: string;
}

export interface ProjectVaultFile {
  id: string;
  projectCode: string;
  subfolderId: string | null;
  filename: string;
  sizeBytes: number;
  mimeType: string;
  kind: AttachmentKind;
  addedAt: string;
  drawingNoKey?: string;
  discipline?: ShipbuildingDrawingRef['discipline'];
  sedmsStatus: 'linked' | 'pending' | 'unmatched';
  sourceEmailId?: string;
  notes?: string;
}

export interface ProjectMailOverrides {
  /** Explicitly added email IDs per projectCode */
  added: Record<string, string[]>;
  /** Explicitly removed (excluded) email IDs per projectCode */
  removed: Record<string, string[]>;
}

export const EMPTY_PROJECT_MAIL_OVERRIDES: ProjectMailOverrides = {
  added: {},
  removed: {},
};

export function inferAttachmentKindFromFilename(filename: string): AttachmentKind {
  const lower = filename.toLowerCase();
  if (lower.endsWith('.pdf')) return 'pdf';
  if (/\.(xlsx|xls|csv)$/.test(lower)) return 'excel';
  if (/\.(docx|doc|hwp)$/.test(lower)) return 'word';
  if (/\.(mp3|wav|m4a|ogg|webm)$/.test(lower)) return 'audio';
  if (/\.(txt|md|json|xml|eml)$/.test(lower)) return 'text';
  return 'other';
}

const IMAGE_ATTACHMENT_EXT_RE = /\.(png|jpe?g|gif|bmp|webp|svg|ico)$/i;
const DOC_PROTECTED_ATTACHMENT_EXT_RE = /\.(pdf|docx?|xlsx?|csv|pptx?|dwg|dxf|zip|7z|rar|txt|md|hwp)$/i;
const INLINE_SIGNATURE_FILENAME_RE =
  /^(image\d*|img_?\d*|outlook[-_ ]?logo|sig[-_ ]?\d*|signature[-_ a-z0-9]*|name[-_ ]?card[-_ a-z0-9]*|card[-_ ]?\d*|logo[-_ a-z0-9]*|banner[-_ a-z0-9]*|clip_image\d*|oledata\d*|header[-_ ]?\d*|footer[-_ ]?\d*|spacer\d*|icon[-_ ]?\d*|명함[-_ 0-9가-힣]*|서명[-_ 0-9가-힣]*|로고[-_ 0-9가-힣]*)\.(png|jpe?g|gif|bmp|webp|svg)$/i;

/**
 * Shared classifier identifying Outlook inline signature / business-card / name-tag images
 * so they never clutter real business document attachment lists or batch exports.
 */
export function isInlineBusinessCardImage(
  att: Pick<MailAttachmentMeta, 'filename' | 'mimeType' | 'sizeBytes' | 'contentId' | 'isInline'>,
  bodyHtml?: string
): boolean {
  if (att.isInline === true) return true;
  const filename = String(att.filename || '').trim();
  if (!filename || DOC_PROTECTED_ATTACHMENT_EXT_RE.test(filename)) return false;

  const mime = String(att.mimeType || '').toLowerCase();
  const isImage = IMAGE_ATTACHMENT_EXT_RE.test(filename) || mime.startsWith('image/');
  if (!isImage) return false;

  if (att.contentId) {
    if (!bodyHtml) return true;
    const cleanCid = att.contentId.replace(/^<|>$/g, '');
    if (cleanCid && bodyHtml.includes(`cid:${cleanCid}`)) {
      return true;
    }
  }

  const size = Number(att.sizeBytes || 0);
  if (INLINE_SIGNATURE_FILENAME_RE.test(filename) && (size === 0 || size < 250_000)) {
    return true;
  }
  if (size > 0 && size < 12_000 && isImage) {
    return true;
  }
  return false;
}

export function toggleFavoriteFolder(favorites: string[], folderId: string): string[] {
  const clean = folderId.trim();
  if (!clean) return favorites;
  return favorites.includes(clean)
    ? favorites.filter((id) => id !== clean)
    : [...favorites, clean];
}

export function assignEmailToProject(
  overrides: ProjectMailOverrides,
  projectCode: string,
  emailId: string
): ProjectMailOverrides {
  const code = projectCode.trim();
  if (!code || !emailId) return overrides;
  const currentAdded = overrides.added[code] ?? [];
  const currentRemoved = overrides.removed[code] ?? [];
  return {
    added: {
      ...overrides.added,
      [code]: currentAdded.includes(emailId) ? currentAdded : [emailId, ...currentAdded],
    },
    removed: {
      ...overrides.removed,
      [code]: currentRemoved.filter((id) => id !== emailId),
    },
  };
}

export function removeEmailFromProject(
  overrides: ProjectMailOverrides,
  projectCode: string,
  emailId: string
): ProjectMailOverrides {
  const code = projectCode.trim();
  if (!code || !emailId) return overrides;
  const currentAdded = overrides.added[code] ?? [];
  const currentRemoved = overrides.removed[code] ?? [];
  return {
    added: {
      ...overrides.added,
      [code]: currentAdded.filter((id) => id !== emailId),
    },
    removed: {
      ...overrides.removed,
      [code]: currentRemoved.includes(emailId) ? currentRemoved : [...currentRemoved, emailId],
    },
  };
}

export function isEmailInProject(
  mail: Pick<MailItem, 'id' | 'subject' | 'body' | 'category'>,
  projectCode: string,
  projects: ProjectDefinition[],
  overrides: ProjectMailOverrides
): boolean {
  const removedList = overrides.removed[projectCode] ?? [];
  if (removedList.includes(mail.id)) return false;

  const addedList = overrides.added[projectCode] ?? [];
  if (addedList.includes(mail.id)) return true;

  const detected = detectProjectCode(
    `${mail.subject} ${mail.body} ${mail.category || ''}`,
    projects
  );
  return detected === projectCode;
}

export function createProjectVaultFile(params: {
  projectCode: string;
  subfolderId?: string | null;
  filename: string;
  sizeBytes?: number;
  mimeType?: string;
  sourceEmailId?: string;
  notes?: string;
  autoIndexDrawing?: boolean;
}): ProjectVaultFile {
  const cleanName = params.filename.trim() || 'untitled.bin';
  const refs = params.autoIndexDrawing === false ? [] : extractShipbuildingDrawingRefs(cleanName);
  const firstRef = refs[0];
  return {
    id: `pfile-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    projectCode: params.projectCode.trim(),
    subfolderId: params.subfolderId ?? null,
    filename: cleanName,
    sizeBytes: Math.max(0, Number(params.sizeBytes ?? 0)),
    mimeType: params.mimeType || 'application/octet-stream',
    kind: inferAttachmentKindFromFilename(cleanName),
    addedAt: new Date().toISOString(),
    drawingNoKey: firstRef?.canonKey,
    discipline: firstRef?.discipline,
    sedmsStatus: firstRef ? 'linked' : 'pending',
    sourceEmailId: params.sourceEmailId,
    notes: params.notes,
  };
}

export function removeProjectSubfolderSafely(
  subfolders: ProjectTreeSubfolder[],
  files: ProjectVaultFile[],
  subfolderId: string
): {
  subfolders: ProjectTreeSubfolder[];
  files: ProjectVaultFile[];
} {
  // Collect target subfolder and any nested children
  const toDelete = new Set<string>([subfolderId]);
  let expanded = true;
  while (expanded) {
    expanded = false;
    for (const sf of subfolders) {
      if (sf.parentId && toDelete.has(sf.parentId) && !toDelete.has(sf.id)) {
        toDelete.add(sf.id);
        expanded = true;
      }
    }
  }
  const target = subfolders.find((s) => s.id === subfolderId);
  const fallbackParentId = target?.parentId ?? null;

  return {
    subfolders: subfolders.filter((s) => !toDelete.has(s.id)),
    // Move orphaned files safely up to parent (or project root) rather than losing them
    files: files.map((f) =>
      f.subfolderId && toDelete.has(f.subfolderId) ? { ...f, subfolderId: fallbackParentId } : f
    ),
  };
}

// ============================================================================
// 4. Outlook Rules & Blocked Sender Evaluation
// ============================================================================

export function isBlockedSender(
  mail: Pick<MailItem, 'fromEmail'>,
  blockedSenders: string[],
  safeSenders: string[] = []
): boolean {
  const email = (mail.fromEmail || '').trim().toLowerCase();
  if (!email) return false;
  const domain = email.includes('@') ? email.split('@')[1] : '';

  const isSafe = safeSenders.some((entry) => {
    const clean = entry.trim().toLowerCase();
    return Boolean(clean) && (email === clean || (domain && domain === clean.replace(/^@/, '')));
  });
  if (isSafe) return false;

  return blockedSenders.some((entry) => {
    const clean = entry.trim().toLowerCase();
    if (!clean) return false;
    if (clean.startsWith('@')) {
      return domain === clean.slice(1);
    }
    return email === clean || (domain && domain === clean);
  });
}

export interface EvaluatedMailRuleResult {
  matchedRuleIds: string[];
  assignedProjectCode?: string;
  shouldFlag: boolean;
  shouldMarkRead: boolean;
  shouldArchive?: boolean;
  shouldMoveTrash: boolean;
}

export function evaluateOutlookMailRules(
  mail: Pick<MailItem, 'subject' | 'fromName' | 'fromEmail' | 'body' | 'hasAttachment' | 'attachments'>,
  rules: OutlookMailRule[]
): EvaluatedMailRuleResult {
  const result: EvaluatedMailRuleResult = {
    matchedRuleIds: [],
    shouldFlag: false,
    shouldMarkRead: false,
    shouldArchive: false,
    shouldMoveTrash: false,
  };

  for (const rule of rules) {
    if (!rule.enabled) continue;
    const val = rule.conditionValue.trim().toLowerCase();
    let matched = false;

    if (rule.conditionField === 'from') {
      matched = Boolean(val) && `${mail.fromName} ${mail.fromEmail}`.toLowerCase().includes(val);
    } else if (rule.conditionField === 'subject') {
      matched = Boolean(val) && (mail.subject || '').toLowerCase().includes(val);
    } else if (rule.conditionField === 'drawing') {
      const refs = extractShipbuildingDrawingRefs(`${mail.subject} ${mail.body}`);
      matched = val
        ? refs.some((r) => r.canonKey.toLowerCase().includes(val.replace(/[^a-z0-9]/g, '')))
        : refs.length > 0;
    } else if (rule.conditionField === 'hasAttachment') {
      matched = Boolean(mail.hasAttachment) || Boolean(mail.attachments && mail.attachments.length > 0);
    }

    if (!matched) continue;
    result.matchedRuleIds.push(rule.id);

    if (rule.actionType === 'assign-project' && rule.actionTarget && !result.assignedProjectCode) {
      result.assignedProjectCode = rule.actionTarget;
    } else if (rule.actionType === 'flag') {
      result.shouldFlag = true;
    } else if (rule.actionType === 'mark-read') {
      result.shouldMarkRead = true;
    } else if (rule.actionType === 'archive') {
      result.shouldArchive = true;
    } else if (rule.actionType === 'move-trash') {
      result.shouldMoveTrash = true;
    }
  }

  return result;
}

// ============================================================================
// 5. Search Folders, Quick Steps & Sweep Helpers (Outlook Parity)
// ============================================================================

export function filterEmailsBySearchFolder(
  emails: MailItem[],
  folderId: OutlookSearchFolderId,
  snoozeMap?: Record<string, string>,
  nowIso?: string
): MailItem[] {
  return emails.filter((mail) => {
    if (folderId === 'unread') {
      return !mail.isRead;
    }
    if (folderId === 'flagged') {
      return Boolean(mail.isFlagged);
    }
    if (folderId === 'has-attachments') {
      return Boolean(mail.hasAttachment || (mail.attachments && mail.attachments.length > 0));
    }
    if (folderId === 'large-mail') {
      const totalBytes = (mail.attachments || []).reduce((acc, att) => acc + (Number(att.sizeBytes) || 0), 0);
      return totalBytes >= 1024 * 1024 || (mail.body || '').length >= 20000;
    }
    if (folderId === 'snoozed') {
      const until = snoozeMap?.[mail.id];
      if (!until) return false;
      return !nowIso || until > nowIso;
    }
    return false;
  });
}

export interface AppliedQuickStepPatch {
  markRead?: boolean;
  setFlagged?: boolean;
  setPinned?: boolean;
  moveFolder?: 'archive' | 'trash';
  assignCategoryId?: string;
  assignProjectCode?: string;
}

export function applyQuickStepActions(
  _mail: Pick<MailItem, 'id' | 'isRead' | 'isFlagged'>,
  step: OutlookQuickStep
): AppliedQuickStepPatch {
  const patch: AppliedQuickStepPatch = {};
  for (const action of step.actions) {
    if (action === 'mark-read') patch.markRead = true;
    else if (action === 'flag') patch.setFlagged = true;
    else if (action === 'pin') patch.setPinned = true;
    else if (action === 'archive') patch.moveFolder = 'archive';
    else if (action === 'move-trash') patch.moveFolder = 'trash';
    else if (action === 'assign-category' && step.categoryId) {
      patch.assignCategoryId = step.categoryId;
    } else if (action === 'assign-project' && step.projectCode) {
      patch.assignProjectCode = step.projectCode;
    }
  }
  return patch;
}

export function buildSweepRuleForSender(
  senderEmail: string,
  mode: 'archive' | 'trash'
): OutlookMailRule {
  const cleanEmail = senderEmail.trim().toLowerCase();
  const slug = cleanEmail.replace(/[^a-z0-9]/g, '-').slice(0, 32) || 'sender';
  return {
    id: `sweep-${slug}-${mode}`,
    name: `일괄 정리 · ${cleanEmail || '발신자'} → ${mode === 'trash' ? '휴지통' : '보관함'}`,
    enabled: true,
    conditionField: 'from',
    conditionValue: cleanEmail,
    actionType: mode === 'trash' ? 'move-trash' : 'archive',
  };
}
