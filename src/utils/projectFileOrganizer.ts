import type { AttachmentKind, MailItem } from '../types/mail';
import {
  detectProjectCode,
  type ProjectDefinition,
  type ProjectFolderPreset,
} from './deterministicAnalyzer';
import {
  extractShipbuildingDrawingRefs,
  inferAttachmentKindFromFilename,
} from './outlookWorkspaceModel';

export type ProjectMatchSource =
  | 'manual'
  | 'sender'
  | 'domain'
  | 'account'
  | 'code-or-keyword'
  | 'unassigned';

export interface ProjectClassificationResult {
  projectCode: string; // 'UNASSIGNED' when unmatched
  projectName: string;
  matchSource: ProjectMatchSource;
  matchEvidence: string;
}

export interface ManualOrganizedFile {
  id: string;
  filename: string;
  sizeBytes: number;
  kind?: AttachmentKind;
  date: string; // YYYY-MM-DD
  direction: 'received' | 'sent';
  senderName: string;
  senderEmail: string;
  recipientEmails: string[];
  accountEmail?: string;
  subject?: string;
  manualProjectCode?: string;
  addedAt: string;
}

export interface OrganizedFileEntry {
  id: string;
  sourceType: 'mail-attachment' | 'manual-upload';
  sourceEmailId?: string;
  emailSubject: string;
  filename: string;
  kind: AttachmentKind;
  sizeBytes: number;
  date: string; // YYYY-MM-DD
  yearMonth: string; // YYYY-MM
  direction: 'received' | 'sent';
  senderName: string;
  senderEmail: string;
  senderDomain: string;
  recipientEmails: string[];
  accountEmail: string;
  projectCode: string;
  projectName: string;
  matchSource: ProjectMatchSource;
  matchEvidence: string;
  drawingKeys: string[];
  folderPath: string;
}

export interface FolderTreeNode {
  id: string;
  name: string;
  path: string;
  depth: number;
  fileCount: number;
  totalBytes: number;
  drawingCount: number;
  files: OrganizedFileEntry[];
  children: FolderTreeNode[];
}

export interface ProjectFileReportRow {
  projectCode: string;
  projectName: string;
  fileCount: number;
  totalBytes: number;
  receivedCount: number;
  sentCount: number;
  uniqueSenders: string[];
  domains: string[];
  accounts: string[];
  drawingKeys: string[];
  earliestDate: string;
  latestDate: string;
  matchBreakdown: Record<ProjectMatchSource, number>;
  files: OrganizedFileEntry[];
}

export interface DateFileReportRow {
  date: string;
  fileCount: number;
  totalBytes: number;
  receivedCount: number;
  sentCount: number;
  projectCodes: string[];
  senders: string[];
  files: OrganizedFileEntry[];
}

export interface ParticipantFileReportRow {
  email: string;
  displayName: string;
  domain: string;
  role: 'sender' | 'recipient' | 'both';
  fileCount: number;
  totalBytes: number;
  receivedCount: number;
  sentCount: number;
  projectCodes: string[];
  latestDate: string;
  files: OrganizedFileEntry[];
}

export interface FileOrganizationReport {
  generatedAt: string;
  preset: ProjectFolderPreset;
  summary: {
    totalFiles: number;
    totalBytes: number;
    receivedFileCount: number;
    sentFileCount: number;
    projectCount: number;
    unassignedFileCount: number;
    uniqueSendersCount: number;
    uniqueDomainsCount: number;
    uniqueDatesCount: number;
    totalDrawingsDetected: number;
  };
  byProject: ProjectFileReportRow[];
  byDate: DateFileReportRow[];
  byParticipant: ParticipantFileReportRow[];
  byFileKind: Record<AttachmentKind, { count: number; totalBytes: number }>;
}

export interface FileOrganizerFilter {
  projectCode?: string; // 'ALL' | 'UNASSIGNED' | specific code
  dateFrom?: string;
  dateTo?: string;
  direction?: 'all' | 'received' | 'sent';
  senderOrDomain?: string;
  account?: string;
  fileKind?: 'all' | AttachmentKind;
  query?: string;
}

export const FOLDER_PRESET_META: Record<
  ProjectFolderPreset,
  { label: string; patternLabel: string; description: string }
> = {
  'project-month-sender': {
    label: '프로젝트 / 연월 / 송신인',
    patternLabel: '/{프로젝트}/{YYYY-MM}/{도메인_송신인}/{파일명}',
    description: '프로젝트별로 월간·거래처(도메인/송신인) 폴더를 자동 생성합니다.',
  },
  'project-date-kind': {
    label: '프로젝트 / 일자 / 파일유형',
    patternLabel: '/{프로젝트}/{YYYY-MM-DD}/{문서유형}/{파일명}',
    description: '프로젝트 하위 일자별로 PDF·Excel·Word·도면 파일을 정리합니다.',
  },
  'domain-project-date': {
    label: '도메인 / 프로젝트 / 일자',
    patternLabel: '/{송신인도메인}/{프로젝트}/{YYYY-MM-DD}/{파일명}',
    description: '발주처·협력사 도메인 기준으로 프로젝트 수신 파일을 묶어 보관합니다.',
  },
  'date-project-sender': {
    label: '일자 / 프로젝트 / 송신인',
    patternLabel: '/{YYYY-MM-DD}/{프로젝트}/{송신인}/{파일명}',
    description: '일일 보고서·마감 검수용으로 날짜별 프로젝트 수신 파일을 정리합니다.',
  },
};

export const MATCH_SOURCE_LABELS: Record<ProjectMatchSource, string> = {
  manual: '수동 지정',
  sender: '송신인 규칙',
  domain: '도메인 규칙',
  account: '계정 규칙',
  'code-or-keyword': '코드·키워드',
  unassigned: '미분류',
};

const FILE_KIND_FOLDER_LABELS: Record<AttachmentKind, string> = {
  pdf: 'PDF_문서',
  excel: 'Excel_집계표',
  word: 'Word_보고서',
  audio: '음성_회의록',
  text: '텍스트_로그',
  other: '기타_첨부',
};

/**
 * Extracts a lowercase email address from raw headers like `"Kim" <Kim@Vendor.CO.KR>`.
 */
export function extractNormalizedEmail(raw?: string): string {
  const src = String(raw || '').trim();
  if (!src) return '';
  const match = src.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (match) return match[0].toLowerCase();
  return src.includes('@') ? src.toLowerCase() : '';
}

/**
 * Extracts a normalized domain without leading `@`.
 */
export function extractEmailDomain(raw?: string): string {
  const src = String(raw || '').trim().toLowerCase();
  if (!src) return '';
  const email = extractNormalizedEmail(src);
  if (email && email.includes('@')) {
    return email.split('@')[1].trim();
  }
  const bare = src.replace(/^@+/, '').replace(/^https?:\/\//, '').split('/')[0].trim();
  return /^[a-z0-9.-]+\.[a-z]{2,}$/i.test(bare) ? bare : '';
}

/**
 * Normalizes a list of emails, domains, or account identifiers.
 */
export function normalizeRuleList(
  rawList: string[] | undefined,
  kind: 'email' | 'domain' | 'account'
): string[] {
  if (!Array.isArray(rawList)) return [];
  const out: string[] = [];
  const seen = new Set<string>();
  for (const item of rawList) {
    const trimmed = String(item || '').trim().toLowerCase();
    if (!trimmed) continue;
    let normalized = trimmed;
    if (kind === 'email') {
      normalized = extractNormalizedEmail(trimmed) || trimmed;
    } else if (kind === 'domain') {
      normalized = extractEmailDomain(trimmed) || trimmed.replace(/^@+/, '');
    } else if (kind === 'account') {
      normalized = extractNormalizedEmail(trimmed) || trimmed;
    }
    if (normalized && !seen.has(normalized)) {
      seen.add(normalized);
      out.push(normalized);
    }
  }
  return out;
}

/**
 * Sanitizes a single folder segment to prevent path traversal or shell-unsafe characters.
 */
export function sanitizeFolderSegment(segment: string, fallback = '미지정'): string {
  const clean = String(segment || '')
    .replace(/\.\.+/g, '_')
    .replace(/[\\/:*?"<>|'`$;!&()]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 64);
  return clean || fallback;
}

/**
 * Formats byte counts cleanly for report tables.
 */
export function formatFileSize(bytes: number): string {
  const safe = Math.max(0, Number.isFinite(bytes) ? bytes : 0);
  if (safe < 1024) return `${safe} B`;
  if (safe < 1024 * 1024) return `${(safe / 1024).toFixed(1)} KB`;
  if (safe < 1024 * 1024 * 1024) return `${(safe / (1024 * 1024)).toFixed(2)} MB`;
  return `${(safe / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

/**
 * Deterministic multi-tier classifier using per-project Sender, Domain, Account, and Keyword/Code rules.
 * Priority: Manual Override > Sender Email > Domain > Account > Code/Keyword > Unassigned.
 */
export function classifyByProjectRoutingRules(params: {
  senderEmail?: string;
  recipientEmails?: string[];
  accountEmail?: string;
  subject?: string;
  body?: string;
  category?: string;
  filenames?: string[];
  projects: ProjectDefinition[];
  manualProjectCode?: string;
}): ProjectClassificationResult {
  const { projects } = params;
  const projectByCode = new Map(projects.map((p) => [p.code, p]));

  // 1. Explicit manual assignment
  const manualCode = (params.manualProjectCode || '').trim();
  if (manualCode && manualCode !== 'UNASSIGNED') {
    const found = projectByCode.get(manualCode);
    return {
      projectCode: manualCode,
      projectName: found?.name || manualCode,
      matchSource: 'manual',
      matchEvidence: `사용자 수동 지정 (${manualCode})`,
    };
  }

  const senderEmail = extractNormalizedEmail(params.senderEmail);
  const recipientEmails = (params.recipientEmails ?? [])
    .map((r) => extractNormalizedEmail(r))
    .filter(Boolean);
  const participantEmails = Array.from(new Set([senderEmail, ...recipientEmails].filter(Boolean)));
  const senderDomain = extractEmailDomain(senderEmail);
  const participantDomains = Array.from(
    new Set(participantEmails.map((e) => extractEmailDomain(e)).filter(Boolean))
  );
  const accountCandidates = Array.from(
    new Set(
      [
        extractNormalizedEmail(params.accountEmail) || (params.accountEmail || '').trim().toLowerCase(),
        ...participantEmails,
      ].filter(Boolean)
    )
  );

  // 2. Exact sender / participant email rule match (sender checked first)
  for (const proj of projects) {
    const allowedSenders = normalizeRuleList(proj.senderEmails, 'email');
    if (allowedSenders.length === 0) continue;
    if (senderEmail && allowedSenders.includes(senderEmail)) {
      return {
        projectCode: proj.code,
        projectName: proj.name,
        matchSource: 'sender',
        matchEvidence: `송신인 일치 (${senderEmail})`,
      };
    }
    const matchedParticipant = recipientEmails.find((e) => allowedSenders.includes(e));
    if (matchedParticipant) {
      return {
        projectCode: proj.code,
        projectName: proj.name,
        matchSource: 'sender',
        matchEvidence: `수신인 일치 (${matchedParticipant})`,
      };
    }
  }

  // 3. Domain rule match (sender domain prioritized, then participant domains)
  for (const proj of projects) {
    const allowedDomains = normalizeRuleList(proj.domains, 'domain');
    if (allowedDomains.length === 0) continue;
    if (senderDomain && allowedDomains.includes(senderDomain)) {
      return {
        projectCode: proj.code,
        projectName: proj.name,
        matchSource: 'domain',
        matchEvidence: `도메인 일치 (@${senderDomain})`,
      };
    }
    const matchedDomain = participantDomains.find((d) => allowedDomains.includes(d));
    if (matchedDomain) {
      return {
        projectCode: proj.code,
        projectName: proj.name,
        matchSource: 'domain',
        matchEvidence: `참여자 도메인 일치 (@${matchedDomain})`,
      };
    }
  }

  // 4. Account rule match (bound Outlook/M365/mailplug account)
  for (const proj of projects) {
    const allowedAccounts = normalizeRuleList(proj.accounts, 'account');
    if (allowedAccounts.length === 0) continue;
    const matchedAccount = accountCandidates.find((acc) => allowedAccounts.includes(acc));
    if (matchedAccount) {
      return {
        projectCode: proj.code,
        projectName: proj.name,
        matchSource: 'account',
        matchEvidence: `연결 계정 일치 (${matchedAccount})`,
      };
    }
  }

  // 5. Fallback to project code / name / keywords in subject, body, category, or filenames
  const combinedText = [
    params.subject || '',
    params.body || '',
    params.category || '',
    ...(params.filenames ?? []),
  ].join(' ');
  const detectedCode = detectProjectCode(combinedText, projects);
  if (detectedCode) {
    const found = projectByCode.get(detectedCode);
    return {
      projectCode: detectedCode,
      projectName: found?.name || detectedCode,
      matchSource: 'code-or-keyword',
      matchEvidence: found ? `코드·키워드 일치 (${found.code})` : `제목/본문 코드 감지 (${detectedCode})`,
    };
  }

  return {
    projectCode: 'UNASSIGNED',
    projectName: '미분류 프로젝트',
    matchSource: 'unassigned',
    matchEvidence: '일치하는 송신인·도메인·계정·키워드 규칙 없음',
  };
}

/**
 * Computes the virtual/local target folder path for an organized file entry.
 */
export function buildTargetFolderPath(params: {
  preset: ProjectFolderPreset;
  projectCode: string;
  defaultSubfolder?: string;
  date: string;
  senderName: string;
  senderDomain: string;
  kind: AttachmentKind;
}): string {
  const safeProject = sanitizeFolderSegment(
    params.projectCode === 'UNASSIGNED' ? '99_미분류' : params.projectCode,
    '99_미분류'
  );
  const safeDate = /^\d{4}-\d{2}-\d{2}$/.test(params.date) ? params.date : '날짜미상';
  const safeMonth = /^\d{4}-\d{2}/.test(safeDate) ? safeDate.slice(0, 7) : '연월미상';
  const safeDomain = sanitizeFolderSegment(params.senderDomain || '내부_기타도메인', '내부_기타도메인');
  const safeSender = sanitizeFolderSegment(params.senderName || '발신자미상', '발신자미상');
  const safeKind = FILE_KIND_FOLDER_LABELS[params.kind] || '기타_첨부';
  const sub = params.defaultSubfolder?.trim() ? `/${sanitizeFolderSegment(params.defaultSubfolder)}` : '';

  switch (params.preset) {
    case 'project-date-kind':
      return `/${safeProject}${sub}/${safeDate}/${safeKind}`;
    case 'domain-project-date':
      return `/${safeDomain}/${safeProject}${sub}/${safeDate}`;
    case 'date-project-sender':
      return `/${safeDate}/${safeProject}${sub}/${safeSender}`;
    case 'project-month-sender':
    default:
      return `/${safeProject}${sub}/${safeMonth}/${safeDomain}_${safeSender}`;
  }
}

/**
 * Collects all attachments from real emails + manually added received files, applies routing rules,
 * computes target folder paths, and applies optional UI filters.
 */
export function collectOrganizedFiles(params: {
  emails: MailItem[];
  manualFiles?: ManualOrganizedFile[];
  projects: ProjectDefinition[];
  preset: ProjectFolderPreset;
  ownerAccountEmail?: string;
  fileProjectOverrides?: Record<string, string>;
  filter?: FileOrganizerFilter;
}): OrganizedFileEntry[] {
  const {
    emails,
    manualFiles = [],
    projects,
    preset,
    ownerAccountEmail = '',
    fileProjectOverrides = {},
    filter,
  } = params;

  const projectByCode = new Map(projects.map((p) => [p.code, p]));
  const entries: OrganizedFileEntry[] = [];

  for (const mail of emails) {
    const attachments = mail.attachments ?? [];
    if (attachments.length === 0) continue;

    const direction: 'received' | 'sent' = mail.folder === 'sent' ? 'sent' : 'received';
    const senderEmail = extractNormalizedEmail(mail.fromEmail) || mail.fromEmail || '';
    const senderDomain = extractEmailDomain(senderEmail);
    const recipients = [...(mail.to ?? []), ...(mail.cc ?? [])]
      .map((r) => extractNormalizedEmail(r) || r.trim())
      .filter(Boolean);
    const accountEmail =
      extractNormalizedEmail(ownerAccountEmail) ||
      (direction === 'sent' ? senderEmail : recipients[0] || '');
    const mailDate = /^\d{4}-\d{2}-\d{2}$/.test(mail.date) ? mail.date : '';
    const yearMonth = mailDate ? mailDate.slice(0, 7) : '연월미상';

    for (const att of attachments) {
      const fileId = `mailfile::${mail.id}::${att.id}`;
      const classification = classifyByProjectRoutingRules({
        senderEmail,
        recipientEmails: recipients,
        accountEmail,
        subject: mail.subject,
        body: mail.body,
        category: mail.folderName || mail.category,
        filenames: [att.filename],
        projects,
        manualProjectCode: fileProjectOverrides[fileId],
      });

      const matchedProject = projectByCode.get(classification.projectCode);
      const effectivePreset = matchedProject?.folderTemplate ?? preset;
      const drawingKeys = Array.from(
        new Set(
          extractShipbuildingDrawingRefs(`${mail.subject} ${att.filename}`).map((r) => r.canonKey)
        )
      );
      const kind = att.kind || inferAttachmentKindFromFilename(att.filename);

      const folderPath = buildTargetFolderPath({
        preset: effectivePreset,
        projectCode: classification.projectCode,
        defaultSubfolder: matchedProject?.defaultSubfolder,
        date: mailDate,
        senderName: mail.fromName || senderEmail,
        senderDomain,
        kind,
      });

      entries.push({
        id: fileId,
        sourceType: 'mail-attachment',
        sourceEmailId: mail.id,
        emailSubject: mail.subject || '(제목 없음)',
        filename: att.filename,
        kind,
        sizeBytes: Math.max(0, Number(att.sizeBytes) || 0),
        date: mailDate || '날짜미상',
        yearMonth,
        direction,
        senderName: mail.fromName || senderEmail || '발신자미상',
        senderEmail,
        senderDomain,
        recipientEmails: recipients,
        accountEmail,
        projectCode: classification.projectCode,
        projectName: classification.projectName,
        matchSource: classification.matchSource,
        matchEvidence: classification.matchEvidence,
        drawingKeys,
        folderPath,
      });
    }
  }

  for (const mf of manualFiles) {
    const senderEmail = extractNormalizedEmail(mf.senderEmail) || mf.senderEmail || '';
    const senderDomain = extractEmailDomain(senderEmail);
    const recipients = (mf.recipientEmails ?? [])
      .map((r) => extractNormalizedEmail(r) || r.trim())
      .filter(Boolean);
    const accountEmail =
      extractNormalizedEmail(mf.accountEmail || ownerAccountEmail) ||
      (mf.accountEmail || ownerAccountEmail || '').trim();
    const kind = mf.kind || inferAttachmentKindFromFilename(mf.filename);

    const classification = classifyByProjectRoutingRules({
      senderEmail,
      recipientEmails: recipients,
      accountEmail,
      subject: mf.subject || mf.filename,
      filenames: [mf.filename],
      projects,
      manualProjectCode: fileProjectOverrides[mf.id] ?? mf.manualProjectCode,
    });

    const matchedProject = projectByCode.get(classification.projectCode);
    const effectivePreset = matchedProject?.folderTemplate ?? preset;
    const drawingKeys = Array.from(
      new Set(extractShipbuildingDrawingRefs(`${mf.subject || ''} ${mf.filename}`).map((r) => r.canonKey))
    );
    const safeDate = /^\d{4}-\d{2}-\d{2}$/.test(mf.date) ? mf.date : '날짜미상';

    const folderPath = buildTargetFolderPath({
      preset: effectivePreset,
      projectCode: classification.projectCode,
      defaultSubfolder: matchedProject?.defaultSubfolder,
      date: safeDate,
      senderName: mf.senderName || senderEmail,
      senderDomain,
      kind,
    });

    entries.push({
      id: mf.id,
      sourceType: 'manual-upload',
      emailSubject: mf.subject || '(직접 등록 수신 파일)',
      filename: mf.filename,
      kind,
      sizeBytes: Math.max(0, Number(mf.sizeBytes) || 0),
      date: safeDate,
      yearMonth: /^\d{4}-\d{2}/.test(safeDate) ? safeDate.slice(0, 7) : '연월미상',
      direction: mf.direction,
      senderName: mf.senderName || senderEmail || '직접 등록',
      senderEmail,
      senderDomain,
      recipientEmails: recipients,
      accountEmail,
      projectCode: classification.projectCode,
      projectName: classification.projectName,
      matchSource: classification.matchSource,
      matchEvidence: classification.matchEvidence,
      drawingKeys,
      folderPath,
    });
  }

  // Apply optional filters
  const filtered = entries.filter((entry) => {
    if (!filter) return true;
    if (filter.projectCode && filter.projectCode !== 'ALL' && entry.projectCode !== filter.projectCode) {
      return false;
    }
    if (filter.dateFrom && entry.date !== '날짜미상' && entry.date < filter.dateFrom) return false;
    if (filter.dateTo && entry.date !== '날짜미상' && entry.date > filter.dateTo) return false;
    if (filter.direction && filter.direction !== 'all' && entry.direction !== filter.direction) {
      return false;
    }
    if (filter.fileKind && filter.fileKind !== 'all' && entry.kind !== filter.fileKind) {
      return false;
    }
    if (filter.senderOrDomain && filter.senderOrDomain.trim()) {
      const needle = filter.senderOrDomain.trim().toLowerCase().replace(/^@/, '');
      const hay = `${entry.senderName} ${entry.senderEmail} ${entry.senderDomain} ${entry.recipientEmails.join(' ')}`.toLowerCase();
      if (!hay.includes(needle)) return false;
    }
    if (filter.account && filter.account.trim()) {
      const needle = filter.account.trim().toLowerCase();
      if (!entry.accountEmail.toLowerCase().includes(needle)) return false;
    }
    if (filter.query && filter.query.trim()) {
      const q = filter.query.trim().toLowerCase();
      const hay = `${entry.filename} ${entry.emailSubject} ${entry.projectCode} ${entry.projectName} ${entry.senderName} ${entry.senderEmail} ${entry.drawingKeys.join(' ')}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  return filtered.sort((a, b) => {
    const dateCmp = b.date.localeCompare(a.date);
    if (dateCmp !== 0) return dateCmp;
    return a.filename.localeCompare(b.filename);
  });
}

/**
 * Builds a recursive virtual folder tree from organized file entries.
 */
export function buildFolderOrganizationTree(files: OrganizedFileEntry[]): FolderTreeNode[] {
  const rootMap = new Map<string, FolderTreeNode>();

  const ensureNode = (
    container: Map<string, FolderTreeNode>,
    name: string,
    fullPath: string,
    depth: number
  ): FolderTreeNode => {
    let existing = container.get(fullPath);
    if (!existing) {
      existing = {
        id: `folder::${fullPath}`,
        name,
        path: fullPath,
        depth,
        fileCount: 0,
        totalBytes: 0,
        drawingCount: 0,
        files: [],
        children: [],
      };
      container.set(fullPath, existing);
    }
    return existing;
  };

  const childMaps = new Map<string, Map<string, FolderTreeNode>>();

  for (const file of files) {
    const segments = file.folderPath.split('/').filter(Boolean);
    if (segments.length === 0) continue;

    let currentMap = rootMap;
    let accumPath = '';

    segments.forEach((seg, idx) => {
      accumPath += `/${seg}`;
      const node = ensureNode(currentMap, seg, accumPath, idx);
      node.fileCount += 1;
      node.totalBytes += file.sizeBytes;
      node.drawingCount += file.drawingKeys.length;
      if (idx === segments.length - 1) {
        node.files.push(file);
      }
      if (!childMaps.has(accumPath)) {
        childMaps.set(accumPath, new Map());
      }
      currentMap = childMaps.get(accumPath)!;
    });
  }

  const finalize = (map: Map<string, FolderTreeNode>): FolderTreeNode[] => {
    const list = Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
    for (const node of list) {
      const subMap = childMaps.get(node.path);
      if (subMap && subMap.size > 0) {
        node.children = finalize(subMap);
      }
    }
    return list;
  };

  return finalize(rootMap);
}

/**
 * Aggregates organized files into an executive report grouped by Project, Date, Participant, and File Kind.
 */
export function buildFileOrganizationReport(
  files: OrganizedFileEntry[],
  preset: ProjectFolderPreset
): FileOrganizationReport {
  const projectMap = new Map<string, ProjectFileReportRow>();
  const dateMap = new Map<string, DateFileReportRow>();
  const participantMap = new Map<string, ParticipantFileReportRow>();
  const byFileKind: Record<AttachmentKind, { count: number; totalBytes: number }> = {
    pdf: { count: 0, totalBytes: 0 },
    excel: { count: 0, totalBytes: 0 },
    word: { count: 0, totalBytes: 0 },
    audio: { count: 0, totalBytes: 0 },
    text: { count: 0, totalBytes: 0 },
    other: { count: 0, totalBytes: 0 },
  };

  const uniqueSenders = new Set<string>();
  const uniqueDomains = new Set<string>();
  const uniqueDates = new Set<string>();
  const uniqueDrawings = new Set<string>();
  let totalBytes = 0;
  let receivedFileCount = 0;
  let sentFileCount = 0;
  let unassignedFileCount = 0;

  for (const file of files) {
    totalBytes += file.sizeBytes;
    if (file.direction === 'received') receivedFileCount += 1;
    else sentFileCount += 1;
    if (file.projectCode === 'UNASSIGNED') unassignedFileCount += 1;
    if (file.senderEmail) uniqueSenders.add(file.senderEmail);
    if (file.senderDomain) uniqueDomains.add(file.senderDomain);
    if (file.date && file.date !== '날짜미상') uniqueDates.add(file.date);
    for (const dk of file.drawingKeys) uniqueDrawings.add(dk);

    byFileKind[file.kind].count += 1;
    byFileKind[file.kind].totalBytes += file.sizeBytes;

    // 1. Project aggregation
    let projRow = projectMap.get(file.projectCode);
    if (!projRow) {
      projRow = {
        projectCode: file.projectCode,
        projectName: file.projectName,
        fileCount: 0,
        totalBytes: 0,
        receivedCount: 0,
        sentCount: 0,
        uniqueSenders: [],
        domains: [],
        accounts: [],
        drawingKeys: [],
        earliestDate: file.date,
        latestDate: file.date,
        matchBreakdown: {
          manual: 0,
          sender: 0,
          domain: 0,
          account: 0,
          'code-or-keyword': 0,
          unassigned: 0,
        },
        files: [],
      };
      projectMap.set(file.projectCode, projRow);
    }
    projRow.fileCount += 1;
    projRow.totalBytes += file.sizeBytes;
    if (file.direction === 'received') projRow.receivedCount += 1;
    else projRow.sentCount += 1;
    projRow.matchBreakdown[file.matchSource] += 1;
    if (file.senderEmail && !projRow.uniqueSenders.includes(file.senderEmail)) {
      projRow.uniqueSenders.push(file.senderEmail);
    }
    if (file.senderDomain && !projRow.domains.includes(file.senderDomain)) {
      projRow.domains.push(file.senderDomain);
    }
    if (file.accountEmail && !projRow.accounts.includes(file.accountEmail)) {
      projRow.accounts.push(file.accountEmail);
    }
    for (const dk of file.drawingKeys) {
      if (!projRow.drawingKeys.includes(dk)) projRow.drawingKeys.push(dk);
    }
    if (file.date !== '날짜미상') {
      if (projRow.earliestDate === '날짜미상' || file.date < projRow.earliestDate) {
        projRow.earliestDate = file.date;
      }
      if (projRow.latestDate === '날짜미상' || file.date > projRow.latestDate) {
        projRow.latestDate = file.date;
      }
    }
    projRow.files.push(file);

    // 2. Date aggregation
    let dateRow = dateMap.get(file.date);
    if (!dateRow) {
      dateRow = {
        date: file.date,
        fileCount: 0,
        totalBytes: 0,
        receivedCount: 0,
        sentCount: 0,
        projectCodes: [],
        senders: [],
        files: [],
      };
      dateMap.set(file.date, dateRow);
    }
    dateRow.fileCount += 1;
    dateRow.totalBytes += file.sizeBytes;
    if (file.direction === 'received') dateRow.receivedCount += 1;
    else dateRow.sentCount += 1;
    if (!dateRow.projectCodes.includes(file.projectCode)) {
      dateRow.projectCodes.push(file.projectCode);
    }
    const senderLabel = file.senderEmail || file.senderName;
    if (senderLabel && !dateRow.senders.includes(senderLabel)) {
      dateRow.senders.push(senderLabel);
    }
    dateRow.files.push(file);

    // 3. Participant aggregation (Sender & Recipients)
    const upsertParticipant = (
      emailRaw: string,
      displayName: string,
      role: 'sender' | 'recipient'
    ) => {
      const key = (emailRaw || displayName || '미상').toLowerCase();
      let pRow = participantMap.get(key);
      if (!pRow) {
        pRow = {
          email: emailRaw || displayName || '미상',
          displayName: displayName || emailRaw || '미상',
          domain: extractEmailDomain(emailRaw),
          role,
          fileCount: 0,
          totalBytes: 0,
          receivedCount: 0,
          sentCount: 0,
          projectCodes: [],
          latestDate: file.date,
          files: [],
        };
        participantMap.set(key, pRow);
      } else if (pRow.role !== role) {
        pRow.role = 'both';
      }
      if (!pRow.files.some((f) => f.id === file.id)) {
        pRow.fileCount += 1;
        pRow.totalBytes += file.sizeBytes;
        if (file.direction === 'received') pRow.receivedCount += 1;
        else pRow.sentCount += 1;
        pRow.files.push(file);
      }
      if (!pRow.projectCodes.includes(file.projectCode)) {
        pRow.projectCodes.push(file.projectCode);
      }
      if (file.date !== '날짜미상' && (pRow.latestDate === '날짜미상' || file.date > pRow.latestDate)) {
        pRow.latestDate = file.date;
      }
    };

    upsertParticipant(file.senderEmail, file.senderName, 'sender');
    for (const rec of file.recipientEmails) {
      upsertParticipant(rec, rec, 'recipient');
    }
  }

  const byProject = Array.from(projectMap.values()).sort((a, b) => {
    if (a.projectCode === 'UNASSIGNED') return 1;
    if (b.projectCode === 'UNASSIGNED') return -1;
    return b.fileCount - a.fileCount || a.projectCode.localeCompare(b.projectCode);
  });

  const byDate = Array.from(dateMap.values()).sort((a, b) => b.date.localeCompare(a.date));
  const byParticipant = Array.from(participantMap.values()).sort(
    (a, b) => b.fileCount - a.fileCount || a.email.localeCompare(b.email)
  );

  const assignedProjectCount = byProject.filter((p) => p.projectCode !== 'UNASSIGNED').length;

  return {
    generatedAt: new Date().toISOString(),
    preset,
    summary: {
      totalFiles: files.length,
      totalBytes,
      receivedFileCount,
      sentFileCount,
      projectCount: assignedProjectCount,
      unassignedFileCount,
      uniqueSendersCount: uniqueSenders.size,
      uniqueDomainsCount: uniqueDomains.size,
      uniqueDatesCount: uniqueDates.size,
      totalDrawingsDetected: uniqueDrawings.size,
    },
    byProject,
    byDate,
    byParticipant,
    byFileKind,
  };
}

/**
 * Formats the aggregated report into a clean, structured Markdown executive document.
 */
export function formatOrganizationReportMarkdown(report: FileOrganizationReport): string {
  const lines: string[] = [
    '# 프로젝트·일자·송수신인별 파일 정리 분석 보고서',
    '',
    `- 생성 일시: ${report.generatedAt.slice(0, 19).replace('T', ' ')} UTC`,
    `- 폴더 정리 기준: ${FOLDER_PRESET_META[report.preset].label} (\`${FOLDER_PRESET_META[report.preset].patternLabel}\`)`,
    '',
    '## 1. 핵심 요약 (Executive Summary)',
    '',
    '| 항목 | 수치 |',
    '| --- | ---: |',
    `| 분류 완료 프로젝트 | ${report.summary.projectCount}개 |`,
    `| 총 정리 파일 수 | ${report.summary.totalFiles}건 (수신 ${report.summary.receivedFileCount} · 발신 ${report.summary.sentFileCount}) |`,
    `| 총 파일 용량 | ${formatFileSize(report.summary.totalBytes)} |`,
    `| 미분류 파일 | ${report.summary.unassignedFileCount}건 |`,
    `| 관련 송신인 / 도메인 | ${report.summary.uniqueSendersCount}명 / ${report.summary.uniqueDomainsCount}개 도메인 |`,
    `| 식별된 도면 번호 | ${report.summary.totalDrawingsDetected}건 |`,
    '',
    '## 2. 프로젝트별 파일 정리 현황',
    '',
  ];

  if (report.byProject.length === 0) {
    lines.push('_정리 대상 파일이 없습니다._', '');
  } else {
    lines.push(
      '| 프로젝트 | 파일 수 | 수신/발신 | 용량 | 주요 도메인 | 식별 도면 | 기간 |',
      '| --- | ---: | :---: | ---: | --- | --- | --- |'
    );
    for (const p of report.byProject) {
      const codeLabel = p.projectCode === 'UNASSIGNED' ? '미분류' : `${p.projectCode} (${p.projectName})`;
      lines.push(
        `| ${codeLabel} | ${p.fileCount} | ${p.receivedCount}/${p.sentCount} | ${formatFileSize(p.totalBytes)} | ${p.domains.slice(0, 3).join(', ') || '-'} | ${p.drawingKeys.slice(0, 3).join(', ') || '-'} | ${p.earliestDate} ~ ${p.latestDate} |`
      );
    }
    lines.push('');
  }

  lines.push('## 3. 일자별 파일 수발신 타임라인', '');
  if (report.byDate.length === 0) {
    lines.push('_일자별 파일 내역이 없습니다._', '');
  } else {
    lines.push(
      '| 일자 | 파일 수 | 수신/발신 | 용량 | 관련 프로젝트 | 주요 송신인 |',
      '| --- | ---: | :---: | ---: | --- | --- |'
    );
    for (const d of report.byDate) {
      lines.push(
        `| ${d.date} | ${d.fileCount} | ${d.receivedCount}/${d.sentCount} | ${formatFileSize(d.totalBytes)} | ${d.projectCodes.join(', ')} | ${d.senders.slice(0, 3).join(', ') || '-'} |`
      );
    }
    lines.push('');
  }

  lines.push('## 4. 송수신인 · 도메인별 파일 집계', '');
  if (report.byParticipant.length === 0) {
    lines.push('_송수신인별 파일 내역이 없습니다._', '');
  } else {
    lines.push(
      '| 송수신인 | 이메일 | 도메인 | 역할 | 파일 수 | 용량 | 관련 프로젝트 |',
      '| --- | --- | --- | :---: | ---: | ---: | --- |'
    );
    for (const u of report.byParticipant.slice(0, 40)) {
      const roleKr = u.role === 'both' ? '송·수신' : u.role === 'sender' ? '송신' : '수신';
      lines.push(
        `| ${u.displayName} | ${u.email} | ${u.domain || '-'} | ${roleKr} | ${u.fileCount} | ${formatFileSize(u.totalBytes)} | ${u.projectCodes.join(', ')} |`
      );
    }
    lines.push('');
  }

  return lines.join('\n');
}

function escapeCsvCell(value: string | number): string {
  const str = String(value ?? '');
  if (/[",\r\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/**
 * Generates an RFC-4180 CSV file table (Excel-ready with BOM).
 */
export function formatOrganizationReportCsv(files: OrganizedFileEntry[]): string {
  const header = [
    '프로젝트코드',
    '프로젝트명',
    '분류근거',
    '일자',
    '구분',
    '송신인',
    '송신인이메일',
    '도메인',
    '수신인',
    '연결계정',
    '파일명',
    '파일유형',
    '크기(Bytes)',
    '도면번호',
    '정리폴더경로',
  ];
  const rows = [header.map(escapeCsvCell).join(',')];
  for (const f of files) {
    rows.push(
      [
        f.projectCode === 'UNASSIGNED' ? '미분류' : f.projectCode,
        f.projectName,
        `${MATCH_SOURCE_LABELS[f.matchSource]} (${f.matchEvidence})`,
        f.date,
        f.direction === 'received' ? '수신' : '발신',
        f.senderName,
        f.senderEmail,
        f.senderDomain,
        f.recipientEmails.join('; '),
        f.accountEmail,
        f.filename,
        f.kind.toUpperCase(),
        f.sizeBytes,
        f.drawingKeys.join('; '),
        `${f.folderPath}/${sanitizeFolderSegment(f.filename, 'file')}`,
      ]
        .map(escapeCsvCell)
        .join(',')
    );
  }
  return `\uFEFF${rows.join('\n')}`;
}

/**
 * Generates a safe local folder-organization script (PowerShell .ps1 or POSIX .sh)
 * so users can mirror the exact Bmail received-file folder tree on Windows/macOS/NAS.
 */
export function generateFolderOrganizerScript(
  files: OrganizedFileEntry[],
  shell: 'ps1' | 'sh',
  rootDirName = 'Bmail_Organized_Files'
): string {
  const safeRoot = sanitizeFolderSegment(rootDirName, 'Bmail_Organized_Files');
  const uniqueFolders = Array.from(new Set(files.map((f) => f.folderPath))).sort();

  if (shell === 'ps1') {
    const lines = [
      '# Bmail 수신파일 자동 폴더 정리 스크립트 (Windows PowerShell)',
      `# 생성 파일 수: ${files.length}건 · 대상 폴더 수: ${uniqueFolders.length}개`,
      'param([string]$SourceDir = ".\\Downloads", [string]$TargetRoot = ".\\' + safeRoot + '")',
      '$ErrorActionPreference = "Continue"',
      '',
      '# 1. 프로젝트·일자·송신인 폴더 트리 생성',
    ];
    for (const folder of uniqueFolders) {
      const winRel = folder.split('/').filter(Boolean).map((s) => sanitizeFolderSegment(s)).join('\\');
      lines.push(`New-Item -ItemType Directory -Force -Path (Join-Path $TargetRoot '${winRel}') | Out-Null`);
    }
    lines.push('', '# 2. 다운로드 폴더의 일치 파일 자동 복사');
    for (const f of files) {
      const winRel = f.folderPath.split('/').filter(Boolean).map((s) => sanitizeFolderSegment(s)).join('\\');
      const safeFile = sanitizeFolderSegment(f.filename, 'file');
      lines.push(
        `if (Test-Path (Join-Path $SourceDir '${safeFile}')) { Copy-Item -Path (Join-Path $SourceDir '${safeFile}') -Destination (Join-Path $TargetRoot '${winRel}\\${safeFile}') -Force }`
      );
    }
    return lines.join('\n');
  }

  const lines = [
    '#!/usr/bin/env bash',
    '# Bmail 수신파일 자동 폴더 정리 스크립트 (POSIX Shell)',
    `# 생성 파일 수: ${files.length}건 · 대상 폴더 수: ${uniqueFolders.length}개`,
    'set -euo pipefail',
    'SOURCE_DIR="${1:-./Downloads}"',
    `TARGET_ROOT="\${2:-./${safeRoot}}"`,
    '',
    '# 1. 프로젝트·일자·송신인 폴더 트리 생성',
  ];
  for (const folder of uniqueFolders) {
    const posixRel = folder.split('/').filter(Boolean).map((s) => sanitizeFolderSegment(s)).join('/');
    lines.push(`mkdir -p "$TARGET_ROOT/${posixRel}"`);
  }
  lines.push('', '# 2. 다운로드 폴더의 일치 파일 자동 복사');
  for (const f of files) {
    const posixRel = f.folderPath.split('/').filter(Boolean).map((s) => sanitizeFolderSegment(s)).join('/');
    const safeFile = sanitizeFolderSegment(f.filename, 'file');
    lines.push(
      `[ -f "$SOURCE_DIR/${safeFile}" ] && cp -f "$SOURCE_DIR/${safeFile}" "$TARGET_ROOT/${posixRel}/${safeFile}" || true`
    );
  }
  return lines.join('\n');
}

/**
 * Extracts distinct sender emails, domains, and accounts discovered across loaded emails
 * to offer 1-click quick-add suggestions inside the Per-Project Routing Setup UI.
 */
export function discoverRoutingSuggestions(
  emails: MailItem[],
  ownerAccountEmail?: string
): {
  senders: string[];
  domains: string[];
  accounts: string[];
} {
  const senders = new Set<string>();
  const domains = new Set<string>();
  const accounts = new Set<string>();

  if (ownerAccountEmail && extractNormalizedEmail(ownerAccountEmail)) {
    accounts.add(extractNormalizedEmail(ownerAccountEmail));
  }

  for (const mail of emails) {
    const from = extractNormalizedEmail(mail.fromEmail);
    if (from) {
      senders.add(from);
      const dom = extractEmailDomain(from);
      if (dom) domains.add(dom);
      if (mail.folder === 'sent') accounts.add(from);
    }
    for (const rec of [...(mail.to ?? []), ...(mail.cc ?? [])]) {
      const addr = extractNormalizedEmail(rec);
      if (addr) {
        const dom = extractEmailDomain(addr);
        if (dom) domains.add(dom);
        if (mail.folder === 'inbox') accounts.add(addr);
      }
    }
  }

  return {
    senders: Array.from(senders).sort(),
    domains: Array.from(domains).sort(),
    accounts: Array.from(accounts).sort(),
  };
}
