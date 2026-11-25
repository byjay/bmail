import type {
  AnalysisResult,
  AttachmentKind,
  MailItem,
  ParsedDocumentSummary,
} from '../types/mail';
import { detectProjectCode, type ProjectDefinition } from './deterministicAnalyzer';

export interface DailyBriefingItem {
  id: string;
  text: string;
  source: 'ai' | 'sent-mail' | 'received-mail' | 'project-summary';
  projectCode?: string;
  priority?: '높음' | '보통' | '참고';
  relatedSubject?: string;
}

export interface DailyBriefingDay {
  date: string; // YYYY-MM-DD
  weekdayLabel: string;
  receivedCount: number;
  sentCount: number;
  attachmentCount: number;
  myWorkItems: DailyBriefingItem[];
  followUpItems: DailyBriefingItem[];
  projectsActive: string[];
}

export interface DomainExchangeStat {
  domain: string;
  received: number;
  sent: number;
  total: number;
  attachmentCount: number;
  lastDate: string;
  accounts: string[];
}

export interface AccountExchangeStat {
  email: string;
  displayName: string;
  domain: string;
  received: number;
  sent: number;
  total: number;
  lastDate: string;
}

export interface ProjectMetricStat {
  projectCode: string;
  projectName: string;
  color: string;
  received: number;
  sent: number;
  total: number;
  attachmentCount: number;
  sharePercent: number;
}

export interface AttachmentKindStat {
  kind: AttachmentKind;
  label: string;
  count: number;
  totalBytes: number;
  sharePercent: number;
  sampleFiles: string[];
}

export interface AttachmentInventoryRow {
  id: string;
  filename: string;
  kind: AttachmentKind;
  sizeBytes: number;
  date: string;
  direction: 'received' | 'sent' | 'uploaded';
  counterpart: string;
  projectCode?: string;
  sourceSubject: string;
  engineSummary?: string;
}

export interface DailyVolumePoint {
  date: string;
  shortLabel: string;
  received: number;
  sent: number;
  total: number;
}

export interface MailboxDashboardMetrics {
  totalEmails: number;
  receivedCount: number;
  sentCount: number;
  uniqueDomainsCount: number;
  uniqueAccountsCount: number;
  activeProjectsCount: number;
  totalAttachmentsCount: number;
  parsedDocumentsCount: number;
  totalFollowUpsCount: number;
  dateRangeLabel: string;
  dailyVolume: DailyVolumePoint[];
  projectStats: ProjectMetricStat[];
  domainStats: DomainExchangeStat[];
  accountStats: AccountExchangeStat[];
  attachmentKindStats: AttachmentKindStat[];
  attachmentInventory: AttachmentInventoryRow[];
  executiveSummaryLines: string[];
}

const WEEKDAYS_KO = ['일', '월', '화', '수', '목', '금', '토'] as const;

const ATTACHMENT_KIND_LABELS: Record<AttachmentKind, string> = {
  pdf: 'PDF 문서',
  excel: 'Excel · CSV',
  word: 'Word 문서',
  audio: '음성 · 녹취',
  text: '텍스트 · MD',
  other: '기타 파일',
};

const PROJECT_PALETTE = [
  '#0F6CBD',
  '#0B7A75',
  '#6B4C9A',
  '#B45309',
  '#BE123C',
  '#1D4ED8',
  '#047857',
];

const FOLLOW_UP_SIGNAL = /(요청|회신|검토|확인|승인|마감|제출|공유|부탁|피드백|urgent|asap|deadline|review)/i;

export function extractEmailDomain(rawAddress: string): string | null {
  if (!rawAddress || typeof rawAddress !== 'string') return null;
  const match = rawAddress.trim().toLowerCase().match(/[a-z0-9._%+-]+@([a-z0-9.-]+\.[a-z]{2,})/i);
  if (!match) return null;
  const domain = match[1];
  if (domain.endsWith('.local') || domain === 'unknown') return null;
  return domain;
}

function normalizeEmailAddress(rawAddress: string): string | null {
  if (!rawAddress || typeof rawAddress !== 'string') return null;
  const match = rawAddress.trim().toLowerCase().match(/([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})/i);
  if (!match) return null;
  const addr = match[1];
  if (addr.endsWith('.local')) return null;
  return addr;
}

function getKoreanWeekday(isoDate: string): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(isoDate)) return '';
  const dt = new Date(`${isoDate}T00:00:00Z`);
  if (Number.isNaN(dt.getTime())) return '';
  return WEEKDAYS_KO[dt.getUTCDay()] ?? '';
}

function resolveResultAnchorDate(result: AnalysisResult | null | undefined, fallbackDate: string): string {
  if (!result) return fallbackDate;
  const isoMatches = result.periodLabel.match(/\b(20\d{2}-\d{2}-\d{2})\b/g);
  if (isoMatches && isoMatches.length > 0) {
    return isoMatches[isoMatches.length - 1];
  }
  if (result.createdAt && /^\d{4}-\d{2}-\d{2}/.test(result.createdAt)) {
    return result.createdAt.slice(0, 10);
  }
  return fallbackDate;
}

function dedupeItems(items: DailyBriefingItem[]): DailyBriefingItem[] {
  const seen = new Set<string>();
  const out: DailyBriefingItem[] = [];
  for (const item of items) {
    const key = item.text.replace(/\s+/g, ' ').trim().toLowerCase();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    out.push(item);
  }
  return out;
}

/**
 * Extracts a date-grouped Daily Work Briefing timeline containing "What I Did" (내가 한 일)
 * and "Follow-up Actions" (후속 조치) from both mailbox records and AI AnalysisResult(s).
 */
export function buildDailyBriefingTimeline(params: {
  emails: MailItem[];
  analysisResult?: AnalysisResult | null;
  archivedResults?: AnalysisResult[];
  projects?: ProjectDefinition[];
}): DailyBriefingDay[] {
  const emails = Array.isArray(params.emails) ? params.emails : [];
  const projects = Array.isArray(params.projects) ? params.projects : [];
  const byDate = new Map<
    string,
    {
      receivedCount: number;
      sentCount: number;
      attachmentCount: number;
      myWorkItems: DailyBriefingItem[];
      followUpItems: DailyBriefingItem[];
      projectsActive: Set<string>;
    }
  >();

  const ensureDay = (date: string) => {
    const safeDate = /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : '날짜 미상';
    let entry = byDate.get(safeDate);
    if (!entry) {
      entry = {
        receivedCount: 0,
        sentCount: 0,
        attachmentCount: 0,
        myWorkItems: [],
        followUpItems: [],
        projectsActive: new Set<string>(),
      };
      byDate.set(safeDate, entry);
    }
    return entry;
  };

  // Build subject -> date index so AI work items can anchor to the exact email date
  const subjectToDate = new Map<string, string>();
  for (const mail of emails) {
    if (mail.subject && /^\d{4}-\d{2}-\d{2}$/.test(mail.date)) {
      subjectToDate.set(mail.subject.trim().toLowerCase(), mail.date);
    }
  }

  // 1. Deterministic extraction from every real email (Sent = 내가 한 일, Received with signal = 후속 조치)
  for (const mail of emails) {
    const day = ensureDay(mail.date);
    const fullText = `${mail.subject} ${mail.body} ${mail.folderName || mail.category || ''}`;
    const projectCode = detectProjectCode(fullText, projects) ?? undefined;
    if (projectCode) day.projectsActive.add(projectCode);

    const attCount = mail.attachments?.length ?? (mail.hasAttachment ? 1 : 0);
    day.attachmentCount += attCount;

    if (mail.folder === 'sent') {
      day.sentCount += 1;
      const attHint =
        mail.attachments && mail.attachments.length > 0
          ? ` (첨부: ${mail.attachments.map((a) => a.filename).slice(0, 2).join(', ')})`
          : '';
      day.myWorkItems.push({
        id: `sent-${mail.id}`,
        text: `${mail.subject}${attHint}`,
        source: 'sent-mail',
        projectCode,
        relatedSubject: mail.subject,
      });
    } else {
      day.receivedCount += 1;
      if (FOLLOW_UP_SIGNAL.test(`${mail.subject} ${mail.snippet} ${mail.body}`)) {
        const isHigh = /(긴급|마감|승인|urgent|asap|deadline)/i.test(`${mail.subject} ${mail.body}`);
        day.followUpItems.push({
          id: `recv-${mail.id}`,
          text: `[${mail.fromName}] ${mail.subject}`,
          source: 'received-mail',
          projectCode,
          priority: isHigh ? '높음' : '보통',
          relatedSubject: mail.subject,
        });
      }
    }
  }

  // 2. Merge structured AI AnalysisResult(s) ("내가 한 일", "프로젝트별 조치", "수신 대응", "익일 계획")
  const allResults: AnalysisResult[] = [];
  if (params.analysisResult) allResults.push(params.analysisResult);
  if (Array.isArray(params.archivedResults)) {
    for (const r of params.archivedResults) {
      if (!allResults.some((existing) => existing.id === r.id)) {
        allResults.push(r);
      }
    }
  }

  const latestEmailDate =
    emails.find((m) => /^\d{4}-\d{2}-\d{2}$/.test(m.date))?.date ||
    new Date().toISOString().slice(0, 10);

  for (const res of allResults) {
    const anchorDate = resolveResultAnchorDate(res, latestEmailDate);

    // AI workPerformed -> 내가 한 일
    for (let i = 0; i < (res.workPerformed ?? []).length; i += 1) {
      const wp = res.workPerformed[i];
      const matchedDate =
        (wp.relatedEmails ?? [])
          .map((subj) => subjectToDate.get(subj.trim().toLowerCase()))
          .find(Boolean) || anchorDate;
      const day = ensureDay(matchedDate);
      const prj = detectProjectCode(`${wp.title} ${wp.detail}`, projects) ?? undefined;
      if (prj) day.projectsActive.add(prj);
      day.myWorkItems.unshift({
        id: `ai-wp-${res.id}-${i}`,
        text: `${wp.title}${wp.detail ? ` — ${wp.detail}` : ''}`,
        source: 'ai',
        projectCode: prj,
      });
      if (wp.status === '후속 조치 필요') {
        day.followUpItems.unshift({
          id: `ai-wp-fu-${res.id}-${i}`,
          text: `[후속 필요] ${wp.title}`,
          source: 'ai',
          projectCode: prj,
          priority: '높음',
        });
      }
    }

    // AI projectSummaries -> 내가 한 일 + 프로젝트별 다음 조치
    for (let i = 0; i < (res.projectSummaries ?? []).length; i += 1) {
      const ps = res.projectSummaries![i];
      const day = ensureDay(anchorDate);
      if (ps.projectCode) day.projectsActive.add(ps.projectCode);
      for (let j = 0; j < ps.myActions.length; j += 1) {
        day.myWorkItems.unshift({
          id: `ai-ps-act-${res.id}-${i}-${j}`,
          text: `[${ps.projectCode}] ${ps.myActions[j]}`,
          source: 'project-summary',
          projectCode: ps.projectCode,
        });
      }
      if (ps.nextAction && ps.nextAction.trim()) {
        day.followUpItems.unshift({
          id: `ai-ps-next-${res.id}-${i}`,
          text: `[${ps.projectCode}] ${ps.nextAction}`,
          source: 'project-summary',
          projectCode: ps.projectCode,
          priority: ps.status === '후속 조치 필요' ? '높음' : '보통',
        });
      }
    }

    // AI incomingHighlights -> 후속 조치
    for (let i = 0; i < (res.incomingHighlights ?? []).length; i += 1) {
      const inc = res.incomingHighlights[i];
      const matchedDate = subjectToDate.get(inc.subject.trim().toLowerCase()) || anchorDate;
      const day = ensureDay(matchedDate);
      const prj = detectProjectCode(`${inc.subject} ${inc.intent}`, projects) ?? undefined;
      if (prj) day.projectsActive.add(prj);
      day.followUpItems.unshift({
        id: `ai-inc-${res.id}-${i}`,
        text: `[${inc.sender}] ${inc.actionNeeded} (${inc.subject})`,
        source: 'ai',
        projectCode: prj,
        priority: inc.priority,
        relatedSubject: inc.subject,
      });
    }

    // AI workLog.nextStepPlan -> 후속 조치
    for (let i = 0; i < (res.workLog?.nextStepPlan ?? []).length; i += 1) {
      const step = res.workLog.nextStepPlan[i];
      if (!step.trim()) continue;
      const day = ensureDay(anchorDate);
      day.followUpItems.push({
        id: `ai-next-${res.id}-${i}`,
        text: step,
        source: 'ai',
        priority: '보통',
      });
    }
  }

  return Array.from(byDate.entries())
    .map(([date, bucket]) => ({
      date,
      weekdayLabel: getKoreanWeekday(date),
      receivedCount: bucket.receivedCount,
      sentCount: bucket.sentCount,
      attachmentCount: bucket.attachmentCount,
      myWorkItems: dedupeItems(bucket.myWorkItems),
      followUpItems: dedupeItems(bucket.followUpItems),
      projectsActive: Array.from(bucket.projectsActive),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));
}

/**
 * Computes comprehensive dashboard analytics across projects, email domains/accounts,
 * attachments (mail + parsed studio docs), daily volume charts, and executive summary lines.
 */
export function buildMailboxDashboardMetrics(params: {
  emails: MailItem[];
  documents?: ParsedDocumentSummary[];
  projects?: ProjectDefinition[];
  analysisResult?: AnalysisResult | null;
}): MailboxDashboardMetrics {
  const emails = Array.isArray(params.emails) ? params.emails : [];
  const documents = Array.isArray(params.documents) ? params.documents : [];
  const projects = Array.isArray(params.projects) ? params.projects : [];

  const receivedCount = emails.filter((m) => m.folder === 'inbox').length;
  const sentCount = emails.filter((m) => m.folder === 'sent').length;

  // 1. Daily Volume
  const volumeMap = new Map<string, { received: number; sent: number }>();
  for (const mail of emails) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(mail.date)) continue;
    const prev = volumeMap.get(mail.date) ?? { received: 0, sent: 0 };
    if (mail.folder === 'sent') prev.sent += 1;
    else prev.received += 1;
    volumeMap.set(mail.date, prev);
  }
  const sortedDates = Array.from(volumeMap.keys()).sort();
  const recentDates = sortedDates.slice(-14);
  const dailyVolume: DailyVolumePoint[] = recentDates.map((date) => {
    const v = volumeMap.get(date)!;
    return {
      date,
      shortLabel: date.slice(5),
      received: v.received,
      sent: v.sent,
      total: v.received + v.sent,
    };
  });

  // 2. Project stats
  const projectBuckets = new Map<
    string,
    { name: string; color: string; received: number; sent: number; attachmentCount: number }
  >();
  for (const mail of emails) {
    const fullText = `${mail.subject} ${mail.body} ${mail.folderName || mail.category || ''}`;
    const detected = detectProjectCode(fullText, projects) || '미분류';
    const def = projects.find((p) => p.code.toLowerCase() === detected.toLowerCase());
    const prev = projectBuckets.get(detected) ?? {
      name: def?.name || (detected === '미분류' ? '미분류 일반' : detected),
      color: def?.color || PROJECT_PALETTE[projectBuckets.size % PROJECT_PALETTE.length],
      received: 0,
      sent: 0,
      attachmentCount: 0,
    };
    if (mail.folder === 'sent') prev.sent += 1;
    else prev.received += 1;
    prev.attachmentCount += mail.attachments?.length ?? (mail.hasAttachment ? 1 : 0);
    projectBuckets.set(detected, prev);
  }

  const projectStats: ProjectMetricStat[] = Array.from(projectBuckets.entries())
    .map(([code, b]) => {
      const total = b.received + b.sent;
      return {
        projectCode: code,
        projectName: b.name,
        color: b.color,
        received: b.received,
        sent: b.sent,
        total,
        attachmentCount: b.attachmentCount,
        sharePercent: emails.length > 0 ? Math.round((total / emails.length) * 100) : 0,
      };
    })
    .sort((a, b) => b.total - a.total);

  // 3. Domain & Account exchange counts (수신: fromEmail, 발신: each recipient in to[])
  const domainMap = new Map<
    string,
    { received: number; sent: number; attachmentCount: number; lastDate: string; accounts: Set<string> }
  >();
  const accountMap = new Map<
    string,
    { displayName: string; domain: string; received: number; sent: number; lastDate: string }
  >();

  const recordInteraction = (
    rawAddr: string,
    displayName: string,
    direction: 'received' | 'sent',
    date: string,
    attCount: number
  ) => {
    const email = normalizeEmailAddress(rawAddr);
    const domain = extractEmailDomain(rawAddr);
    if (!email || !domain) return;

    const dPrev = domainMap.get(domain) ?? {
      received: 0,
      sent: 0,
      attachmentCount: 0,
      lastDate: '',
      accounts: new Set<string>(),
    };
    if (direction === 'received') dPrev.received += 1;
    else dPrev.sent += 1;
    dPrev.attachmentCount += attCount;
    if (date > dPrev.lastDate) dPrev.lastDate = date;
    dPrev.accounts.add(email);
    domainMap.set(domain, dPrev);

    const aPrev = accountMap.get(email) ?? {
      displayName: displayName || email.split('@')[0],
      domain,
      received: 0,
      sent: 0,
      lastDate: '',
    };
    if (direction === 'received') aPrev.received += 1;
    else aPrev.sent += 1;
    if (date > aPrev.lastDate) aPrev.lastDate = date;
    if (displayName && aPrev.displayName === email.split('@')[0]) {
      aPrev.displayName = displayName;
    }
    accountMap.set(email, aPrev);
  };

  for (const mail of emails) {
    const attCount = mail.attachments?.length ?? (mail.hasAttachment ? 1 : 0);
    if (mail.folder === 'inbox') {
      recordInteraction(mail.fromEmail, mail.fromName, 'received', mail.date, attCount);
    } else {
      const recipients = Array.isArray(mail.to) ? mail.to : [];
      for (const recipient of recipients) {
        recordInteraction(recipient, '', 'sent', mail.date, attCount);
      }
    }
  }

  const domainStats: DomainExchangeStat[] = Array.from(domainMap.entries())
    .map(([domain, d]) => ({
      domain,
      received: d.received,
      sent: d.sent,
      total: d.received + d.sent,
      attachmentCount: d.attachmentCount,
      lastDate: d.lastDate,
      accounts: Array.from(d.accounts),
    }))
    .sort((a, b) => b.total - a.total || a.domain.localeCompare(b.domain));

  const accountStats: AccountExchangeStat[] = Array.from(accountMap.entries())
    .map(([email, a]) => ({
      email,
      displayName: a.displayName,
      domain: a.domain,
      received: a.received,
      sent: a.sent,
      total: a.received + a.sent,
      lastDate: a.lastDate,
    }))
    .sort((a, b) => b.total - a.total || a.email.localeCompare(b.email));

  // 4. Attachments & Parsed Documents inventory + kind stats
  const inventory: AttachmentInventoryRow[] = [];
  for (const mail of emails) {
    const fullText = `${mail.subject} ${mail.body} ${mail.folderName || mail.category || ''}`;
    const projectCode = detectProjectCode(fullText, projects) ?? undefined;
    if (mail.attachments && mail.attachments.length > 0) {
      for (const att of mail.attachments) {
        inventory.push({
          id: `${mail.id}-${att.id}`,
          filename: att.filename,
          kind: att.kind,
          sizeBytes: att.sizeBytes ?? 0,
          date: mail.date,
          direction: mail.folder === 'sent' ? 'sent' : 'received',
          counterpart: mail.folder === 'sent' ? mail.to[0] || '수신자' : mail.fromName,
          projectCode,
          sourceSubject: mail.subject,
        });
      }
    } else if (mail.hasAttachment) {
      inventory.push({
        id: `${mail.id}-att-generic`,
        filename: `${mail.subject.slice(0, 24)} 첨부파일`,
        kind: 'other',
        sizeBytes: 0,
        date: mail.date,
        direction: mail.folder === 'sent' ? 'sent' : 'received',
        counterpart: mail.folder === 'sent' ? mail.to[0] || '수신자' : mail.fromName,
        projectCode,
        sourceSubject: mail.subject,
      });
    }
  }

  for (const doc of documents) {
    const projectCode = detectProjectCode(`${doc.filename} ${doc.extractedText}`, projects) ?? undefined;
    const summaryBits: string[] = [doc.engineUsed];
    if (doc.pageCount) summaryBits.push(`${doc.pageCount}쪽`);
    if (doc.sheetNames && doc.sheetNames.length > 0) {
      summaryBits.push(`시트 ${doc.sheetNames.length}개 (${doc.rowCount ?? 0}행)`);
    }
    if (doc.numericHighlights && doc.numericHighlights.length > 0) {
      summaryBits.push(doc.numericHighlights[0]);
    }
    inventory.push({
      id: `doc-${doc.id}`,
      filename: doc.filename,
      kind: doc.kind,
      sizeBytes: doc.sizeBytes,
      date: new Date().toISOString().slice(0, 10),
      direction: 'uploaded',
      counterpart: '문서 AI 스튜디오',
      projectCode,
      sourceSubject: doc.filename,
      engineSummary: summaryBits.join(' · '),
    });
  }

  inventory.sort((a, b) => b.date.localeCompare(a.date));

  const kindMap = new Map<AttachmentKind, { count: number; totalBytes: number; samples: string[] }>();
  for (const row of inventory) {
    const prev = kindMap.get(row.kind) ?? { count: 0, totalBytes: 0, samples: [] };
    prev.count += 1;
    prev.totalBytes += Math.max(0, row.sizeBytes || 0);
    if (prev.samples.length < 3 && !prev.samples.includes(row.filename)) {
      prev.samples.push(row.filename);
    }
    kindMap.set(row.kind, prev);
  }

  const totalAttachmentsCount = inventory.length;
  const attachmentKindStats: AttachmentKindStat[] = (
    ['pdf', 'excel', 'word', 'audio', 'text', 'other'] as AttachmentKind[]
  )
    .map((kind) => {
      const bucket = kindMap.get(kind);
      const count = bucket?.count ?? 0;
      return {
        kind,
        label: ATTACHMENT_KIND_LABELS[kind],
        count,
        totalBytes: bucket?.totalBytes ?? 0,
        sharePercent: totalAttachmentsCount > 0 ? Math.round((count / totalAttachmentsCount) * 100) : 0,
        sampleFiles: bucket?.samples ?? [],
      };
    })
    .filter((s) => s.count > 0)
    .sort((a, b) => b.count - a.count);

  const briefing = buildDailyBriefingTimeline({
    emails,
    analysisResult: params.analysisResult,
    projects,
  });
  const totalFollowUpsCount = briefing.reduce((acc, d) => acc + d.followUpItems.length, 0);

  const dateRangeLabel =
    sortedDates.length >= 2
      ? `${sortedDates[0]} ~ ${sortedDates[sortedDates.length - 1]}`
      : sortedDates[0] || '연동 기간 없음';

  // 5. Auto executive summary lines
  const executiveSummaryLines: string[] = [];
  if (emails.length > 0) {
    executiveSummaryLines.push(
      `${dateRangeLabel} 기간 동안 총 ${emails.length}건(수신 ${receivedCount}건 · 발신 ${sentCount}건)의 메일을 ${domainStats.length}개 도메인(${accountStats.length}개 계정)과 교신했습니다.`
    );
  }
  if (projectStats.length > 0) {
    const topProject = projectStats[0];
    executiveSummaryLines.push(
      `가장 활발한 프로젝트는 [${topProject.projectCode}] ${topProject.projectName} (${topProject.total}건, 전체의 ${topProject.sharePercent}%)입니다.`
    );
  }
  if (domainStats.length > 0) {
    const topDomain = domainStats[0];
    executiveSummaryLines.push(
      `최다 교신 도메인은 ${topDomain.domain} (총 ${topDomain.total}회: 수신 ${topDomain.received} · 발신 ${topDomain.sent})입니다.`
    );
  }
  if (totalAttachmentsCount > 0) {
    const topKind = attachmentKindStats[0];
    executiveSummaryLines.push(
      `첨부파일·업로드 문서 총 ${totalAttachmentsCount}건 중 ${topKind.label} 비중이 ${topKind.sharePercent}%(${topKind.count}건)로 가장 높습니다.`
    );
  }
  if (params.analysisResult?.overview) {
    executiveSummaryLines.push(`AI 핵심 요약: ${params.analysisResult.overview}`);
  }

  return {
    totalEmails: emails.length,
    receivedCount,
    sentCount,
    uniqueDomainsCount: domainStats.length,
    uniqueAccountsCount: accountStats.length,
    activeProjectsCount: projectStats.filter((p) => p.projectCode !== '미분류').length,
    totalAttachmentsCount,
    parsedDocumentsCount: documents.length,
    totalFollowUpsCount,
    dateRangeLabel,
    dailyVolume,
    projectStats,
    domainStats,
    accountStats,
    attachmentKindStats,
    attachmentInventory: inventory,
    executiveSummaryLines,
  };
}

export function formatDailyBriefingMarkdown(days: DailyBriefingDay[]): string {
  if (days.length === 0) return '# 일간 업무 브리핑\n\n표시할 브리핑 항목이 없습니다.';
  const lines: string[] = ['# 일간 업무 브리핑 타임라인', ''];
  for (const day of days) {
    lines.push(
      `## ${day.date}${day.weekdayLabel ? ` (${day.weekdayLabel})` : ''} — 수신 ${day.receivedCount}건 · 발신 ${day.sentCount}건 · 첨부 ${day.attachmentCount}건`
    );
    lines.push('### 내가 한 일');
    if (day.myWorkItems.length === 0) {
      lines.push('- 기록된 발신/처리 내역 없음');
    } else {
      for (const item of day.myWorkItems) {
        lines.push(`- [x] ${item.text}`);
      }
    }
    lines.push('### 후속 조치');
    if (day.followUpItems.length === 0) {
      lines.push('- 대기 중인 후속 조치 없음');
    } else {
      for (const item of day.followUpItems) {
        lines.push(`- [ ] ${item.priority ? `[${item.priority}] ` : ''}${item.text}`);
      }
    }
    lines.push('');
  }
  return lines.join('\n');
}

export interface ProjectProgressPoint {
  projectCode: string;
  projectName: string;
  color: string;
  received: number;
  sent: number;
  total: number;
  readCount: number;
  pendingFollowUps: number;
  progressPercent: number;
  status: '완료' | '진행 중' | '후속 조치 필요';
}

export interface DailyEfficiencyPoint {
  date: string;
  shortLabel: string;
  completedWork: number;
  pendingFollowUps: number;
  totalHandled: number;
  efficiencyPercent: number;
}

export interface ProjectResponseComparisonPoint {
  projectCode: string;
  projectName: string;
  color: string;
  mailCount: number;
  received: number;
  sent: number;
  sharePercent: number;
  avgResponseHours: number | null; // null if no matched inbox->sent reply thread exists
  responseSpeedScore: number; // 0..100 (higher = faster response / higher responsiveness)
  pairedThreads: number;
}

export interface MonthlyMailTrendPoint {
  month: string; // YYYY-MM
  shortLabel: string; // YY.MM
  received: number;
  sent: number;
  total: number;
  readRatePercent: number;
  efficiencyPercent: number;
}

export interface MonthlyMailTrendSummary {
  currentMonthLabel: string; // YYYY-MM or '-'
  currentMonthTotal: number;
  currentMonthReceived: number;
  currentMonthSent: number;
  previousMonthLabel: string | null;
  previousMonthTotal: number;
  momDeltaPercent: number | null;
  avgResponseHours: number | null;
  fastestResponseHours: number | null;
  measuredReplyPairs: number;
  monthlyReadRatePercent: number;
  monthlyEfficiencyPercent: number;
  topProjectCode: string | null;
  topProjectSharePercent: number;
  monthlySeries: MonthlyMailTrendPoint[];
}

export interface WorkEfficiencyMetrics {
  overallEfficiencyPercent: number;
  readRatePercent: number;
  responseRatePercent: number;
  followUpClosurePercent: number;
  totalEmails: number;
  readInboxCount: number;
  totalInboxCount: number;
  sentCount: number;
  flaggedPendingCount: number;
  completedWorkCount: number;
  pendingFollowUpCount: number;
  projectProgress: ProjectProgressPoint[];
  dailyTrend: DailyEfficiencyPoint[];
  projectResponseComparison: ProjectResponseComparisonPoint[];
  monthlyTrend: MonthlyMailTrendSummary;
}

function clampPercent(value: number): number {
  if (!Number.isFinite(value)) return 0;
  return Math.max(0, Math.min(100, Math.round(value)));
}

function normalizeSubjectThreadKey(subject: string): string {
  return (subject || '')
    .replace(/^(?:(?:re|fw|fwd|답장|전달)\s*:\s*)+/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function parseMailTimeMs(mail: MailItem): number | null {
  if (mail.timestamp) {
    const t = Date.parse(mail.timestamp);
    if (Number.isFinite(t)) return t;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(mail.date)) {
    const t = Date.parse(`${mail.date}T09:00:00Z`);
    if (Number.isFinite(t)) return t;
  }
  return null;
}

function hoursToResponseSpeedScore(avgHours: number | null, fallbackSentRatio: number, fallbackReadRatio: number): number {
  if (avgHours !== null && Number.isFinite(avgHours)) {
    // <= 1h => 100, 4h => 88, 12h => 72, 24h => 55, 48h => 35, >= 96h => 15
    if (avgHours <= 1) return 100;
    const decayed = 100 * Math.exp(-avgHours / 40);
    return clampPercent(Math.max(12, Math.min(100, decayed)));
  }
  return clampPercent((fallbackSentRatio * 0.6 + fallbackReadRatio * 0.4) * 100);
}

/**
 * Pure deterministic calculator for D3 Project Progress, Multi-Series Response Speed vs Distribution,
 * Monthly Trend Summary KPI cards, and Mail Processing Efficiency charts.
 * Strictly zero-simulation: returns zeroed metrics and empty arrays when emails is empty.
 */
export function buildWorkEfficiencyMetrics(params: {
  emails: MailItem[];
  projects?: ProjectDefinition[];
  analysisResult?: AnalysisResult | null;
  archivedResults?: AnalysisResult[];
}): WorkEfficiencyMetrics {
  const emails = Array.isArray(params.emails) ? params.emails : [];
  const projects = Array.isArray(params.projects) ? params.projects : [];

  if (emails.length === 0) {
    return {
      overallEfficiencyPercent: 0,
      readRatePercent: 0,
      responseRatePercent: 0,
      followUpClosurePercent: 0,
      totalEmails: 0,
      readInboxCount: 0,
      totalInboxCount: 0,
      sentCount: 0,
      flaggedPendingCount: 0,
      completedWorkCount: 0,
      pendingFollowUpCount: 0,
      projectProgress: [],
      dailyTrend: [],
      projectResponseComparison: [],
      monthlyTrend: {
        currentMonthLabel: '-',
        currentMonthTotal: 0,
        currentMonthReceived: 0,
        currentMonthSent: 0,
        previousMonthLabel: null,
        previousMonthTotal: 0,
        momDeltaPercent: null,
        avgResponseHours: null,
        fastestResponseHours: null,
        measuredReplyPairs: 0,
        monthlyReadRatePercent: 0,
        monthlyEfficiencyPercent: 0,
        topProjectCode: null,
        topProjectSharePercent: 0,
        monthlySeries: [],
      },
    };
  }

  const inboxEmails = emails.filter((m) => m.folder === 'inbox');
  const sentEmails = emails.filter((m) => m.folder === 'sent');
  const readInboxCount = inboxEmails.filter((m) => m.isRead).length;
  const flaggedPendingCount = emails.filter((m) => m.isFlagged).length;

  const readRatePercent =
    inboxEmails.length > 0
      ? clampPercent((readInboxCount / inboxEmails.length) * 100)
      : 100;

  // Response balance: how actively sent replies match received volume (1:1 or higher = 100%)
  const responseRatePercent =
    inboxEmails.length > 0
      ? clampPercent(Math.min(1, sentEmails.length / inboxEmails.length) * 100)
      : sentEmails.length > 0
        ? 100
        : 0;

  const briefingDays = buildDailyBriefingTimeline({
    emails,
    analysisResult: params.analysisResult,
    archivedResults: params.archivedResults,
    projects,
  });

  const completedWorkCount = briefingDays.reduce((sum, d) => sum + d.myWorkItems.length, 0);
  const pendingFollowUpCount = briefingDays.reduce((sum, d) => sum + d.followUpItems.length, 0);
  const totalBriefingTasks = completedWorkCount + pendingFollowUpCount;
  const followUpClosurePercent =
    totalBriefingTasks > 0
      ? clampPercent((completedWorkCount / totalBriefingTasks) * 100)
      : readRatePercent;

  // Weighted composite efficiency score (Read 40% + Follow-up Closure 35% + Response Balance 25%)
  const overallEfficiencyPercent = clampPercent(
    readRatePercent * 0.4 + followUpClosurePercent * 0.35 + responseRatePercent * 0.25
  );

  // Build AI project summary status map if available
  const aiStatusByCode = new Map<string, '완료' | '진행 중' | '후속 조치 필요'>();
  for (const ps of params.analysisResult?.projectSummaries ?? []) {
    if (ps.projectCode) {
      aiStatusByCode.set(ps.projectCode.toLowerCase(), ps.status);
    }
  }

  const resolveProjectCode = (mail: MailItem): string => {
    const fullText = `${mail.subject} ${mail.body} ${mail.folderName || mail.category || ''}`;
    let detected = detectProjectCode(fullText, projects);
    if (!detected && aiStatusByCode.size > 0) {
      const lowerText = fullText.toLowerCase();
      for (const ps of params.analysisResult?.projectSummaries ?? []) {
        if (ps.projectCode && lowerText.includes(ps.projectCode.toLowerCase())) {
          detected = ps.projectCode;
          break;
        }
      }
    }
    return detected || '미분류';
  };

  // Compute deterministic response hours by pairing inbox messages with subsequent sent replies in the same thread
  const threadsByKey = new Map<string, Array<{ mail: MailItem; timeMs: number; projectCode: string }>>();
  for (const mail of emails) {
    const timeMs = parseMailTimeMs(mail);
    if (timeMs === null) continue;
    const threadKey = (mail.conversationKey || normalizeSubjectThreadKey(mail.subject)).trim();
    if (!threadKey) continue;
    const list = threadsByKey.get(threadKey) ?? [];
    list.push({ mail, timeMs, projectCode: resolveProjectCode(mail) });
    threadsByKey.set(threadKey, list);
  }

  const responseHoursByProject = new Map<string, number[]>();
  const allResponseHours: number[] = [];

  for (const entries of threadsByKey.values()) {
    entries.sort((a, b) => a.timeMs - b.timeMs);
    let lastInbox: { timeMs: number; projectCode: string } | null = null;
    for (const item of entries) {
      if (item.mail.folder === 'inbox') {
        lastInbox = { timeMs: item.timeMs, projectCode: item.projectCode };
      } else if (item.mail.folder === 'sent' && lastInbox) {
        const diffHours = (item.timeMs - lastInbox.timeMs) / (1000 * 60 * 60);
        // Accept realistic response windows (0.05h to 336h = 14 days)
        if (diffHours >= 0.05 && diffHours <= 336) {
          const rounded = Math.round(diffHours * 10) / 10;
          allResponseHours.push(rounded);
          const code = item.projectCode !== '미분류' ? item.projectCode : lastInbox.projectCode;
          const arr = responseHoursByProject.get(code) ?? [];
          arr.push(rounded);
          responseHoursByProject.set(code, arr);
        }
        lastInbox = null;
      }
    }
  }

  // Group project progress deterministically from emails
  const projectMap = new Map<
    string,
    {
      name: string;
      color: string;
      received: number;
      sent: number;
      readCount: number;
      pendingFollowUps: number;
    }
  >();

  for (const mail of emails) {
    const resolvedCode = resolveProjectCode(mail);
    const def = projects.find((p) => p.code.toLowerCase() === resolvedCode.toLowerCase());
    const aiSummary = params.analysisResult?.projectSummaries?.find(
      (ps) => ps.projectCode?.toLowerCase() === resolvedCode.toLowerCase()
    );
    const bucket = projectMap.get(resolvedCode) ?? {
      name:
        def?.name ||
        aiSummary?.projectName ||
        (resolvedCode === '미분류' ? '미분류 일반' : resolvedCode),
      color: def?.color || PROJECT_PALETTE[projectMap.size % PROJECT_PALETTE.length],
      received: 0,
      sent: 0,
      readCount: 0,
      pendingFollowUps: 0,
    };

    if (mail.folder === 'sent') {
      bucket.sent += 1;
      bucket.readCount += 1;
    } else {
      bucket.received += 1;
      if (mail.isRead) bucket.readCount += 1;
      if (mail.isFlagged || FOLLOW_UP_SIGNAL.test(`${mail.subject} ${mail.snippet}`)) {
        bucket.pendingFollowUps += 1;
      }
    }
    projectMap.set(resolvedCode, bucket);
  }

  const projectProgress: ProjectProgressPoint[] = Array.from(projectMap.entries())
    .map(([code, b]) => {
      const total = b.received + b.sent;
      const readRatio = total > 0 ? b.readCount / total : 0;
      const sentRatio = b.received > 0 ? Math.min(1, b.sent / b.received) : b.sent > 0 ? 1 : 0;
      const pendingPenalty = total > 0 ? Math.min(0.45, (b.pendingFollowUps / total) * 0.45) : 0;

      const aiStatus = aiStatusByCode.get(code.toLowerCase());
      let progressPercent: number;
      let status: '완료' | '진행 중' | '후속 조치 필요';

      if (aiStatus === '완료') {
        progressPercent = clampPercent(Math.max(88, (readRatio * 0.5 + sentRatio * 0.5) * 100));
        status = '완료';
      } else if (aiStatus === '후속 조치 필요') {
        progressPercent = clampPercent(Math.min(55, (readRatio * 0.5 + sentRatio * 0.3) * 100));
        status = '후속 조치 필요';
      } else {
        progressPercent = clampPercent(
          (readRatio * 0.55 + sentRatio * 0.45 - pendingPenalty) * 100
        );
        if (b.pendingFollowUps > 0 && progressPercent < 55) {
          status = '후속 조치 필요';
        } else if (progressPercent >= 85 && b.pendingFollowUps === 0) {
          status = '완료';
        } else {
          status = '진행 중';
        }
      }

      return {
        projectCode: code,
        projectName: b.name,
        color: b.color,
        received: b.received,
        sent: b.sent,
        total,
        readCount: b.readCount,
        pendingFollowUps: b.pendingFollowUps,
        progressPercent,
        status,
      };
    })
    .sort((a, b) => b.total - a.total || b.progressPercent - a.progressPercent)
    .slice(0, 7);

  // Multi-series dataset: Project mail distribution vs Mail response speed score
  const projectResponseComparison: ProjectResponseComparisonPoint[] = projectProgress.map((p) => {
    const hoursList = responseHoursByProject.get(p.projectCode) ?? [];
    const avgHours =
      hoursList.length > 0
        ? Math.round((hoursList.reduce((acc, h) => acc + h, 0) / hoursList.length) * 10) / 10
        : null;
    const sentRatio = p.received > 0 ? Math.min(1, p.sent / p.received) : p.sent > 0 ? 1 : 0;
    const readRatio = p.total > 0 ? p.readCount / p.total : 0;

    return {
      projectCode: p.projectCode,
      projectName: p.projectName,
      color: p.color,
      mailCount: p.total,
      received: p.received,
      sent: p.sent,
      sharePercent: emails.length > 0 ? clampPercent((p.total / emails.length) * 100) : 0,
      avgResponseHours: avgHours,
      responseSpeedScore: hoursToResponseSpeedScore(avgHours, sentRatio, readRatio),
      pairedThreads: hoursList.length,
    };
  });

  // Monthly mail processing trend summary (YYYY-MM aggregation)
  const monthBuckets = new Map<
    string,
    { received: number; sent: number; readInbox: number; totalInbox: number }
  >();

  for (const mail of emails) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(mail.date)) continue;
    const month = mail.date.slice(0, 7);
    const prev = monthBuckets.get(month) ?? {
      received: 0,
      sent: 0,
      readInbox: 0,
      totalInbox: 0,
    };
    if (mail.folder === 'sent') {
      prev.sent += 1;
    } else {
      prev.received += 1;
      prev.totalInbox += 1;
      if (mail.isRead) prev.readInbox += 1;
    }
    monthBuckets.set(month, prev);
  }

  const sortedMonths = Array.from(monthBuckets.keys()).sort();
  const monthlySeries: MonthlyMailTrendPoint[] = sortedMonths.slice(-6).map((month) => {
    const b = monthBuckets.get(month)!;
    const total = b.received + b.sent;
    const mReadRate = b.totalInbox > 0 ? clampPercent((b.readInbox / b.totalInbox) * 100) : 100;
    const mRespRate =
      b.received > 0 ? clampPercent(Math.min(1, b.sent / b.received) * 100) : b.sent > 0 ? 100 : 0;
    const mEff = clampPercent(mReadRate * 0.55 + mRespRate * 0.45);
    return {
      month,
      shortLabel: `${month.slice(2, 4)}.${month.slice(5, 7)}`,
      received: b.received,
      sent: b.sent,
      total,
      readRatePercent: mReadRate,
      efficiencyPercent: mEff,
    };
  });

  const latestMonth = monthlySeries[monthlySeries.length - 1] ?? null;
  const prevMonth = monthlySeries.length >= 2 ? monthlySeries[monthlySeries.length - 2] : null;
  const momDeltaPercent =
    latestMonth && prevMonth && prevMonth.total > 0
      ? Math.round(((latestMonth.total - prevMonth.total) / prevMonth.total) * 100)
      : null;

  const avgResponseHours =
    allResponseHours.length > 0
      ? Math.round((allResponseHours.reduce((acc, h) => acc + h, 0) / allResponseHours.length) * 10) /
        10
      : null;
  const fastestResponseHours =
    allResponseHours.length > 0 ? Math.min(...allResponseHours) : null;

  const topProject = projectResponseComparison[0] ?? null;

  const monthlyTrend: MonthlyMailTrendSummary = {
    currentMonthLabel: latestMonth?.month ?? '-',
    currentMonthTotal: latestMonth?.total ?? 0,
    currentMonthReceived: latestMonth?.received ?? 0,
    currentMonthSent: latestMonth?.sent ?? 0,
    previousMonthLabel: prevMonth?.month ?? null,
    previousMonthTotal: prevMonth?.total ?? 0,
    momDeltaPercent,
    avgResponseHours,
    fastestResponseHours,
    measuredReplyPairs: allResponseHours.length,
    monthlyReadRatePercent: latestMonth?.readRatePercent ?? readRatePercent,
    monthlyEfficiencyPercent: latestMonth?.efficiencyPercent ?? overallEfficiencyPercent,
    topProjectCode: topProject?.projectCode ?? null,
    topProjectSharePercent: topProject?.sharePercent ?? 0,
    monthlySeries,
  };

  // Chronological daily trend (up to latest 10 active days)
  const chronologicalDays = [...briefingDays]
    .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d.date))
    .sort((a, b) => a.date.localeCompare(b.date))
    .slice(-10);

  const dailyTrend: DailyEfficiencyPoint[] = chronologicalDays.map((day) => {
    const completedWork = day.myWorkItems.length;
    const pendingFollowUps = day.followUpItems.length;
    const totalHandled = day.receivedCount + day.sentCount;
    const denominator = completedWork + pendingFollowUps;
    const efficiencyPercent =
      denominator > 0
        ? clampPercent((completedWork / denominator) * 100)
        : day.sentCount > 0
          ? 100
          : 0;

    return {
      date: day.date,
      shortLabel: day.date.slice(5),
      completedWork,
      pendingFollowUps,
      totalHandled,
      efficiencyPercent,
    };
  });

  return {
    overallEfficiencyPercent,
    readRatePercent,
    responseRatePercent,
    followUpClosurePercent,
    totalEmails: emails.length,
    readInboxCount,
    totalInboxCount: inboxEmails.length,
    sentCount: sentEmails.length,
    flaggedPendingCount,
    completedWorkCount,
    pendingFollowUpCount,
    projectProgress,
    dailyTrend,
    projectResponseComparison,
    monthlyTrend,
  };
}


