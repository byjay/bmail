import crypto from 'node:crypto';
import type { MailItem } from '../types/mail';
import { normalizeConversationSubject, type ProjectDefinition } from '../utils/deterministicAnalyzer';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  isEmailInProject,
  type ProjectMailOverrides,
} from '../utils/outlookWorkspaceModel';
import { classifyAttachmentInline } from './attachmentBatchExporter';
import { invokeClaudeMessagesApi } from './claudeModelRouter';
import {
  getGroqKeyPoolStatus,
  invokeGroqStreamingCompletion,
} from './groqKeyRotator';
import type { LocalMailDb } from './localMailDb';

export type ExecutivePeriodType = 'weekly' | 'monthly' | 'custom' | 'all';

export interface ExecutivePeriodRow {
  periodKey: string;
  periodLabel: string;
  total: number;
  received: number;
  sent: number;
  unrepliedIncoming: number;
  attachmentCount: number;
  topProject: string;
  topTopic: string;
}

export interface ExecutiveUnrepliedItem {
  emailId: string;
  date: string;
  timeLabel: string;
  daysPending: number;
  fromName: string;
  fromEmail: string;
  projectCode: string;
  topic: string;
  subject: string;
  hasAttachment: boolean;
  urgency: 'high' | 'normal';
  snippet: string;
}

export interface ExecutiveProjectRow {
  projectCode: string;
  projectName: string;
  total: number;
  received: number;
  sent: number;
  unrepliedIncoming: number;
  attachmentCount: number;
  latestDate: string;
  latestSubject: string;
  topSenders: string[];
  topTopics: string[];
}

export interface ExecutiveTopicRow {
  topic: string;
  total: number;
  received: number;
  sent: number;
  unrepliedIncoming: number;
  attachmentCount: number;
  topProjects: string[];
  latestSubject: string;
}

export interface ExecutiveMailLedgerRow {
  emailId: string;
  direction: 'received' | 'sent';
  date: string;
  timeLabel: string;
  fromName: string;
  fromEmail: string;
  to: string[];
  projectCode: string;
  topic: string;
  subject: string;
  realAttachmentCount: number;
  replyStatus: 'replied' | 'unreplied-incoming' | 'sent' | 'fyi';
  isRead: boolean;
}

export interface ExecutiveAttachmentRow {
  emailId: string;
  attachmentIndex: number;
  filename: string;
  kind: string;
  sizeBytes: number;
  date: string;
  fromName: string;
  projectCode: string;
  topic: string;
  subject: string;
}

export interface ExecutiveNarrativeReport {
  title: string;
  executiveSummary: string;
  periodTrendNote: string;
  projectFocusNote: string;
  unrepliedRiskNote: string;
  keyHighlights: string[];
  recommendedActions: string[];
  markdownFull: string;
  generatedBy: string;
  generatedAt: string;
}

export interface SavedExecutiveReportSnapshot {
  id: string;
  title: string;
  savedAt: string;
  periodType: ExecutivePeriodType;
  dateRangeLabel: string;
  projectCode?: string;
  topic?: string;
  kpis: ExecutiveDashboardReport['kpis'];
  narrativeMarkdown: string;
}

export interface ExecutiveDashboardReport {
  filters: {
    periodType: ExecutivePeriodType;
    dateFrom?: string;
    dateTo?: string;
    projectCode?: string;
    topic?: string;
    direction: 'all' | 'received' | 'sent';
  };
  kpis: {
    totalEmails: number;
    receivedCount: number;
    sentCount: number;
    unrepliedIncomingCount: number;
    awaitingCounterpartyCount: number;
    realAttachmentCount: number;
    projectCount: number;
    topicCount: number;
    unreadCount: number;
    dateRangeLabel: string;
    persistedTotalInDb: number;
  };
  narrativeReport: ExecutiveNarrativeReport;
  weeklyTable: ExecutivePeriodRow[];
  monthlyTable: ExecutivePeriodRow[];
  unrepliedIncoming: ExecutiveUnrepliedItem[];
  projectRows: ExecutiveProjectRow[];
  topicRows: ExecutiveTopicRow[];
  mailRows: ExecutiveMailLedgerRow[];
  attachmentRows: ExecutiveAttachmentRow[];
  availableProjects: Array<{ code: string; name: string }>;
  availableTopics: string[];
  savedReports: SavedExecutiveReportSnapshot[];
  lastSyncedAt: string | null;
}

const NEWSLETTER_OR_NOREPLY_RE =
  /noreply|no-reply|newsletter|mailer-daemon|donotreply|수신거부|광고|웹메일 안내|뉴스레터|notification/i;

export function classifyEmailTopic(
  email: Pick<MailItem, 'subject' | 'body' | 'folderName'>
): string {
  const text = `${email.subject || ''} ${(email.body || '').slice(0, 600)} ${email.folderName || ''}`.toLowerCase();
  if (/(dwg|drawing|승인도|출도|도면|rev\.?\s?[a-z0-9]|revision)/.test(text)) {
    return '도면·승인도·REV';
  }
  if (/(pos|p\.o\.s|rfq|견적|발주|단가|계약|purchase|quotation|price)/.test(text)) {
    return '견적·발주·POS';
  }
  if (/(선급|kr|abs|dnv|lr|bv|nk|검사|미결|comment|class)/.test(text)) {
    return '선급·검사·미결';
  }
  if (/(3d|modeling|모델링|생산설계|tribon|aveva|cad|block|블록)/.test(text)) {
    return '3D 모델링·생산설계';
  }
  if (/(주간보고|월간|일정|납기|schedule|delivery|meeting|회의|화상회의)/.test(text)) {
    return '납기·주간보고·일정';
  }
  if (/(urgent|긴급|하자|클레임|defect|issue|지연|delay|asap)/.test(text)) {
    return '하자·긴급·이슈';
  }
  return '일반 업무·공유';
}

function getIsoWeekKey(dateStr: string): string {
  const m = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return 'UNDATED';
  const dt = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  const dayNum = dt.getUTCDay() || 7;
  dt.setUTCDate(dt.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(dt.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((dt.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return `${dt.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function formatWeekLabel(weekKey: string): string {
  const m = weekKey.match(/^(\d{4})-W(\d{2})$/);
  if (!m) return weekKey;
  const year = Number(m[1]);
  const week = Number(m[2]);
  const simple = new Date(Date.UTC(year, 0, 1 + (week - 1) * 7));
  const dow = simple.getUTCDay();
  const monday = new Date(simple);
  monday.setUTCDate(simple.getUTCDate() - (dow === 0 ? 6 : dow - 1));
  const sunday = new Date(monday);
  sunday.setUTCDate(monday.getUTCDate() + 6);
  const mmdd = (d: Date) =>
    `${String(d.getUTCMonth() + 1).padStart(2, '0')}/${String(d.getUTCDate()).padStart(2, '0')}`;
  return `${weekKey} (${mmdd(monday)}~${mmdd(sunday)})`;
}

function resolveProjectCodeAndName(
  email: MailItem,
  projects: ProjectDefinition[],
  overrides: ProjectMailOverrides
): { code: string; name: string } {
  for (const p of projects) {
    if (isEmailInProject(email, p.code, projects, overrides)) {
      return { code: p.code, name: p.name };
    }
  }
  const codeMatch = `${email.subject} ${email.folderName || ''}`.match(
    /\b([A-Z]{2,6}[-_]\d{2,6}|\d{2}K[-_\s]?FD|H\d{3,4})\b/i
  );
  if (codeMatch) {
    const code = codeMatch[1].toUpperCase().replace(/\s+/g, '');
    return { code, name: code };
  }
  const cleanFolder = (email.folderName || '').split('/').pop()?.trim();
  if (cleanFolder && !/받은 편지함|보낸 편지함|지운 편지함|임시 보관함/i.test(cleanFolder)) {
    return { code: cleanFolder, name: cleanFolder };
  }
  return { code: 'GENERAL', name: '일반·기타 업무' };
}

export function buildExecutiveDashboardReport(
  db: LocalMailDb,
  options?: {
    periodType?: ExecutivePeriodType;
    dateFrom?: string;
    dateTo?: string;
    projectCode?: string;
    topic?: string;
    direction?: 'all' | 'received' | 'sent';
  },
  projects?: ProjectDefinition[],
  overrides: ProjectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES
): ExecutiveDashboardReport {
  const resolvedProjects =
    projects ?? db.getState<ProjectDefinition[]>('projects') ?? [];
  const periodType: ExecutivePeriodType = options?.periodType || 'all';
  const direction: 'all' | 'received' | 'sent' = options?.direction || 'all';

  let effFrom = options?.dateFrom;
  let effTo = options?.dateTo;
  const allDbEmails = db.queryAllEmailsForAnalytics();
  const persistedTotalInDb = allDbEmails.length;

  const anchorDateStr = allDbEmails[0]?.date || new Date().toISOString().slice(0, 10);
  const anchorMs = Date.parse(`${anchorDateStr}T00:00:00Z`) || Date.now();
  if (!effFrom && !effTo) {
    if (periodType === 'weekly') {
      effFrom = new Date(anchorMs - 6 * 86400000).toISOString().slice(0, 10);
      effTo = anchorDateStr;
    } else if (periodType === 'monthly') {
      effFrom = new Date(anchorMs - 29 * 86400000).toISOString().slice(0, 10);
      effTo = anchorDateStr;
    }
  }

  // Track latest sent & received timestamps per normalized conversation key across FULL DB
  const latestSentByConv = new Map<string, string>();
  const latestRecvByConv = new Map<string, string>();
  for (const m of allDbEmails) {
    const conv = m.conversationKey || normalizeConversationSubject(m.subject);
    if (!conv) continue;
    const ts = m.timestamp || `${m.date}T00:00:00Z`;
    if (m.folder === 'sent') {
      const prev = latestSentByConv.get(conv);
      if (!prev || ts > prev) latestSentByConv.set(conv, ts);
    } else {
      const prev = latestRecvByConv.get(conv);
      if (!prev || ts > prev) latestRecvByConv.set(conv, ts);
    }
  }

  const availableTopicSet = new Set<string>();
  const availableProjMap = new Map<string, string>();

  interface Enriched {
    mail: MailItem;
    projCode: string;
    projName: string;
    topic: string;
    realAtts: Array<{ idx: number; filename: string; kind: string; sizeBytes: number }>;
    replyStatus: 'replied' | 'unreplied-incoming' | 'sent' | 'fyi';
  }

  const matched: Enriched[] = [];
  for (const m of allDbEmails) {
    if (effFrom && m.date < effFrom) continue;
    if (effTo && m.date > effTo) continue;
    if (direction === 'received' && m.folder !== 'inbox') continue;
    if (direction === 'sent' && m.folder !== 'sent') continue;

    const proj = resolveProjectCodeAndName(m, resolvedProjects, overrides);
    const topic = classifyEmailTopic(m);
    availableTopicSet.add(topic);
    availableProjMap.set(proj.code, proj.name);

    if (options?.projectCode && options.projectCode !== 'ALL' && proj.code !== options.projectCode) {
      continue;
    }
    if (options?.topic && options.topic !== 'ALL' && topic !== options.topic) {
      continue;
    }

    const conv = m.conversationKey || normalizeConversationSubject(m.subject);
    const ts = m.timestamp || `${m.date}T00:00:00Z`;
    const isNewsletter = NEWSLETTER_OR_NOREPLY_RE.test(`${m.fromEmail} ${m.subject}`);

    let replyStatus: Enriched['replyStatus'] = 'sent';
    if (m.folder !== 'sent') {
      if (isNewsletter) {
        replyStatus = 'fyi';
      } else {
        const sentAfter = conv ? latestSentByConv.get(conv) : undefined;
        replyStatus = sentAfter && sentAfter >= ts ? 'replied' : 'unreplied-incoming';
      }
    }

    const realAtts: Enriched['realAtts'] = [];
    if (Array.isArray(m.attachments)) {
      m.attachments.forEach((att, idx) => {
        if (!classifyAttachmentInline(att)) {
          realAtts.push({
            idx,
            filename: att.filename,
            kind: att.kind,
            sizeBytes: Number(att.sizeBytes || 0),
          });
        }
      });
    }

    matched.push({
      mail: m,
      projCode: proj.code,
      projName: proj.name,
      topic,
      realAtts,
      replyStatus,
    });
  }

  let receivedCount = 0;
  let sentCount = 0;
  let unreadCount = 0;
  let realAttachmentCount = 0;
  const unrepliedIncoming: ExecutiveUnrepliedItem[] = [];
  const awaitingConvSet = new Set<string>();

  type PeriodAcc = {
    total: number;
    received: number;
    sent: number;
    unrepliedIncoming: number;
    attachmentCount: number;
    projCounts: Map<string, number>;
    topicCounts: Map<string, number>;
  };
  const weeklyMap = new Map<string, PeriodAcc>();
  const monthlyMap = new Map<string, PeriodAcc>();
  const projectMap = new Map<
    string,
    {
      name: string;
      total: number;
      received: number;
      sent: number;
      unrepliedIncoming: number;
      attachmentCount: number;
      latestDate: string;
      latestSubject: string;
      senderCounts: Map<string, number>;
      topicCounts: Map<string, number>;
    }
  >();
  const topicMap = new Map<
    string,
    {
      total: number;
      received: number;
      sent: number;
      unrepliedIncoming: number;
      attachmentCount: number;
      projCounts: Map<string, number>;
      latestSubject: string;
    }
  >();

  const bumpPeriod = (map: Map<string, PeriodAcc>, key: string, e: Enriched) => {
    let acc = map.get(key);
    if (!acc) {
      acc = {
        total: 0,
        received: 0,
        sent: 0,
        unrepliedIncoming: 0,
        attachmentCount: 0,
        projCounts: new Map(),
        topicCounts: new Map(),
      };
      map.set(key, acc);
    }
    acc.total += 1;
    if (e.mail.folder === 'sent') acc.sent += 1;
    else acc.received += 1;
    if (e.replyStatus === 'unreplied-incoming') acc.unrepliedIncoming += 1;
    acc.attachmentCount += e.realAtts.length;
    acc.projCounts.set(e.projCode, (acc.projCounts.get(e.projCode) || 0) + 1);
    acc.topicCounts.set(e.topic, (acc.topicCounts.get(e.topic) || 0) + 1);
  };

  for (const e of matched) {
    const m = e.mail;
    if (m.folder === 'sent') {
      sentCount += 1;
      const conv = m.conversationKey || normalizeConversationSubject(m.subject);
      const ts = m.timestamp || `${m.date}T00:00:00Z`;
      const recvAfter = conv ? latestRecvByConv.get(conv) : undefined;
      if (conv && (!recvAfter || recvAfter < ts)) {
        awaitingConvSet.add(conv);
      }
    } else {
      receivedCount += 1;
    }
    if (!m.isRead && m.folder !== 'sent') unreadCount += 1;
    realAttachmentCount += e.realAtts.length;

    if (e.replyStatus === 'unreplied-incoming') {
      const mMs = Date.parse(`${m.date}T00:00:00Z`) || anchorMs;
      const daysPending = Math.max(0, Math.round((anchorMs - mMs) / 86400000));
      const isHigh =
        m.importance === 'high' || /긴급|urgent|asap|미결|승인|기한|금일/i.test(`${m.subject} ${m.snippet}`);
      unrepliedIncoming.push({
        emailId: m.id,
        date: m.date,
        timeLabel: m.timeLabel,
        daysPending,
        fromName: m.fromName,
        fromEmail: m.fromEmail,
        projectCode: e.projCode,
        topic: e.topic,
        subject: m.subject,
        hasAttachment: e.realAtts.length > 0,
        urgency: isHigh ? 'high' : 'normal',
        snippet: m.snippet,
      });
    }

    bumpPeriod(weeklyMap, getIsoWeekKey(m.date), e);
    bumpPeriod(monthlyMap, (m.date || '').slice(0, 7) || 'UNDATED', e);

    let pAcc = projectMap.get(e.projCode);
    if (!pAcc) {
      pAcc = {
        name: e.projName,
        total: 0,
        received: 0,
        sent: 0,
        unrepliedIncoming: 0,
        attachmentCount: 0,
        latestDate: m.date,
        latestSubject: m.subject,
        senderCounts: new Map(),
        topicCounts: new Map(),
      };
      projectMap.set(e.projCode, pAcc);
    }
    pAcc.total += 1;
    if (m.folder === 'sent') pAcc.sent += 1;
    else pAcc.received += 1;
    if (e.replyStatus === 'unreplied-incoming') pAcc.unrepliedIncoming += 1;
    pAcc.attachmentCount += e.realAtts.length;
    if (m.date >= pAcc.latestDate) {
      pAcc.latestDate = m.date;
      pAcc.latestSubject = m.subject;
    }
    pAcc.senderCounts.set(m.fromName || m.fromEmail, (pAcc.senderCounts.get(m.fromName || m.fromEmail) || 0) + 1);
    pAcc.topicCounts.set(e.topic, (pAcc.topicCounts.get(e.topic) || 0) + 1);

    let tAcc = topicMap.get(e.topic);
    if (!tAcc) {
      tAcc = {
        total: 0,
        received: 0,
        sent: 0,
        unrepliedIncoming: 0,
        attachmentCount: 0,
        projCounts: new Map(),
        latestSubject: m.subject,
      };
      topicMap.set(e.topic, tAcc);
    }
    tAcc.total += 1;
    if (m.folder === 'sent') tAcc.sent += 1;
    else tAcc.received += 1;
    if (e.replyStatus === 'unreplied-incoming') tAcc.unrepliedIncoming += 1;
    tAcc.attachmentCount += e.realAtts.length;
    tAcc.projCounts.set(e.projCode, (tAcc.projCounts.get(e.projCode) || 0) + 1);
  }

  const pickTopKey = (map: Map<string, number>): string => {
    let best = '-';
    let bestCount = -1;
    for (const [k, c] of map.entries()) {
      if (c > bestCount) {
        best = k;
        bestCount = c;
      }
    }
    return best;
  };

  const weeklyTable: ExecutivePeriodRow[] = Array.from(weeklyMap.entries())
    .sort((a, b) => b[0].localeCompare(a[0]))
    .slice(0, 26)
    .map(([key, acc]) => ({
      periodKey: key,
      periodLabel: formatWeekLabel(key),
      total: acc.total,
      received: acc.received,
      sent: acc.sent,
      unrepliedIncoming: acc.unrepliedIncoming,
      attachmentCount: acc.attachmentCount,
      topProject: pickTopKey(acc.projCounts),
      topTopic: pickTopKey(acc.topicCounts),
    }));

  const monthlyTable: ExecutivePeriodRow[] = Array.from(monthlyMap.entries())
    .sort((a, b) => b[0].localeCompare(a[0]))
    .slice(0, 24)
    .map(([key, acc]) => ({
      periodKey: key,
      periodLabel: key,
      total: acc.total,
      received: acc.received,
      sent: acc.sent,
      unrepliedIncoming: acc.unrepliedIncoming,
      attachmentCount: acc.attachmentCount,
      topProject: pickTopKey(acc.projCounts),
      topTopic: pickTopKey(acc.topicCounts),
    }));

  unrepliedIncoming.sort((a, b) => {
    if (a.urgency !== b.urgency) return a.urgency === 'high' ? -1 : 1;
    return b.date.localeCompare(a.date);
  });

  const projectRows: ExecutiveProjectRow[] = Array.from(projectMap.entries())
    .map(([code, acc]) => ({
      projectCode: code,
      projectName: acc.name,
      total: acc.total,
      received: acc.received,
      sent: acc.sent,
      unrepliedIncoming: acc.unrepliedIncoming,
      attachmentCount: acc.attachmentCount,
      latestDate: acc.latestDate,
      latestSubject: acc.latestSubject,
      topSenders: Array.from(acc.senderCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([n]) => n),
      topTopics: Array.from(acc.topicCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([t]) => t),
    }))
    .sort((a, b) => b.total - a.total);

  const topicRows: ExecutiveTopicRow[] = Array.from(topicMap.entries())
    .map(([topic, acc]) => ({
      topic,
      total: acc.total,
      received: acc.received,
      sent: acc.sent,
      unrepliedIncoming: acc.unrepliedIncoming,
      attachmentCount: acc.attachmentCount,
      topProjects: Array.from(acc.projCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 3)
        .map(([p]) => p),
      latestSubject: acc.latestSubject,
    }))
    .sort((a, b) => b.total - a.total);

  const mailRows: ExecutiveMailLedgerRow[] = matched.slice(0, 300).map((e) => ({
    emailId: e.mail.id,
    direction: e.mail.folder === 'sent' ? 'sent' : 'received',
    date: e.mail.date,
    timeLabel: e.mail.timeLabel,
    fromName: e.mail.fromName,
    fromEmail: e.mail.fromEmail,
    to: e.mail.to,
    projectCode: e.projCode,
    topic: e.topic,
    subject: e.mail.subject,
    realAttachmentCount: e.realAtts.length,
    replyStatus: e.replyStatus,
    isRead: e.mail.isRead,
  }));

  const attachmentRows: ExecutiveAttachmentRow[] = [];
  for (const e of matched) {
    if (attachmentRows.length >= 300) break;
    for (const att of e.realAtts) {
      if (attachmentRows.length >= 300) break;
      attachmentRows.push({
        emailId: e.mail.id,
        attachmentIndex: att.idx,
        filename: att.filename,
        kind: att.kind,
        sizeBytes: att.sizeBytes,
        date: e.mail.date,
        fromName: e.mail.fromName,
        projectCode: e.projCode,
        topic: e.topic,
        subject: e.mail.subject,
      });
    }
  }

  const minDate = matched[matched.length - 1]?.mail.date || '';
  const maxDate = matched[0]?.mail.date || '';
  const dateRangeLabel =
    effFrom || effTo ? `${effFrom || minDate || '~'} ~ ${effTo || maxDate || '~'}` : minDate && maxDate ? `${minDate} ~ ${maxDate}` : '전체 기간';

  const topProj = projectRows[0];
  const topTopic = topicRows[0];
  const latestWeek = weeklyTable[0];
  const urgentCount = unrepliedIncoming.filter((u) => u.urgency === 'high').length;

  const title = `Bmail 메일함 ${periodType === 'weekly' ? '주간' : periodType === 'monthly' ? '월간' : '종합'} 업무 분석 보고서 (${dateRangeLabel})`;
  const executiveSummary = `조회 기간(${dateRangeLabel}) 동안 총 ${matched.length.toLocaleString()}건의 메일이 집계되었습니다(수신 ${receivedCount.toLocaleString()}건 · 발신 ${sentCount.toLocaleString()}건 · 실첨부 파일 ${realAttachmentCount.toLocaleString()}개). 로컬 SQLite DB에는 총 ${persistedTotalInDb.toLocaleString()}건이 영구 보관되어 있어 재시작 없이 즉시 조회되며 신규 도착 메일만 증분 반영됩니다.`;
  const periodTrendNote = latestWeek
    ? `최근 주차(${latestWeek.periodLabel})에는 ${latestWeek.total.toLocaleString()}건(수신 ${latestWeek.received} · 발신 ${latestWeek.sent} · 첨부 ${latestWeek.attachmentCount}개)이 집중되었으며, 핵심 프로젝트는 ${latestWeek.topProject}, 최다 주제는 '${latestWeek.topTopic}'입니다.`
    : '선택된 기간 내 주차별 집계 데이터가 없습니다.';
  const projectFocusNote = topProj
    ? `전체 업무 비중 1위 프로젝트는 '${topProj.projectCode}'(${topProj.total}건, 미회신 ${topProj.unrepliedIncoming}건, 실첨부 ${topProj.attachmentCount}개)이며, 주요 소통 상대는 ${topProj.topSenders.join(', ') || '미상'}입니다.`
    : '프로젝트별 분류 데이터가 없습니다.';
  const unrepliedRiskNote =
    unrepliedIncoming.length > 0
      ? `내가 회신해야 할 미회신 수신 메일은 총 ${unrepliedIncoming.length}건(긴급·기한 언급 ${urgentCount}건)입니다. 최우선 건은 '${unrepliedIncoming[0].subject}' (${unrepliedIncoming[0].fromName}, ${unrepliedIncoming[0].date})입니다.`
      : '조회 범위 내 미회신 수신 메일이 없습니다.';

  const keyHighlights = [
    `수신 ${receivedCount.toLocaleString()}건 vs 발신 ${sentCount.toLocaleString()}건 (미읽음 ${unreadCount.toLocaleString()}건 · 상대 회신 대기 스레드 ${awaitingConvSet.size}건)`,
    topTopic
      ? `최다 발생 업무 주제: ${topTopic.topic} (${topTopic.total}건 · 첨부 ${topTopic.attachmentCount}개)`
      : '주제 분류: 일반 업무',
    topProj
      ? `최다 교신 프로젝트: ${topProj.projectCode} (${topProj.total}건 · 최근 안건 "${topProj.latestSubject.slice(0, 40)}")`
      : '프로젝트 미분류',
  ];

  const recommendedActions = unrepliedIncoming.slice(0, 5).map(
    (u, idx) =>
      `${idx + 1}. [${u.projectCode} · ${u.topic}] ${u.fromName} (${u.date}${u.daysPending > 0 ? `, D+${u.daysPending}` : ''}) — "${u.subject}" 회신 검토`
  );
  if (recommendedActions.length === 0) {
    recommendedActions.push('현재 긴급 미회신 수신 건이 없으며, 주차별 첨부 도면·POS 정산 내역을 주기적으로 아카이빙하세요.');
  }

  const markdownFull = [
    `# ${title}`,
    '',
    '## 1. 종합 브리핑 (분석글)',
    executiveSummary,
    '',
    periodTrendNote,
    '',
    projectFocusNote,
    '',
    unrepliedRiskNote,
    '',
    '## 2. 핵심 요약 지표',
    ...keyHighlights.map((h) => `- ${h}`),
    '',
    '## 3. 내게 온 미회신건 우선 조치 리스트',
    ...recommendedActions.map((a) => `- ${a}`),
  ].join('\n');

  const meta = db.getState<{ migratedAt?: string }>('migration_meta');
  const savedReports =
    db.getState<SavedExecutiveReportSnapshot[]>('saved_executive_reports') || [];

  return {
    filters: {
      periodType,
      dateFrom: effFrom,
      dateTo: effTo,
      projectCode: options?.projectCode || 'ALL',
      topic: options?.topic || 'ALL',
      direction,
    },
    kpis: {
      totalEmails: matched.length,
      receivedCount,
      sentCount,
      unrepliedIncomingCount: unrepliedIncoming.length,
      awaitingCounterpartyCount: awaitingConvSet.size,
      realAttachmentCount,
      projectCount: projectRows.length,
      topicCount: topicRows.length,
      unreadCount,
      dateRangeLabel,
      persistedTotalInDb,
    },
    narrativeReport: {
      title,
      executiveSummary,
      periodTrendNote,
      projectFocusNote,
      unrepliedRiskNote,
      keyHighlights,
      recommendedActions,
      markdownFull,
      generatedBy: '로컬 결정론적 분석 엔진',
      generatedAt: new Date().toISOString(),
    },
    weeklyTable: weeklyTable.slice(0, 16),
    monthlyTable: monthlyTable.slice(0, 12),
    unrepliedIncoming: unrepliedIncoming.slice(0, 150),
    projectRows,
    topicRows,
    mailRows,
    attachmentRows,
    availableProjects: Array.from(availableProjMap.entries())
      .map(([code, name]) => ({ code, name }))
      .sort((a, b) => a.code.localeCompare(b.code, 'ko')),
    availableTopics: Array.from(availableTopicSet).sort((a, b) => a.localeCompare(b, 'ko')),
    savedReports: savedReports.slice(0, 30),
    lastSyncedAt: meta?.migratedAt || null,
  };
}

export async function generateAiExecutiveNarrative(
  report: ExecutiveDashboardReport,
  apiKey?: string
): Promise<ExecutiveNarrativeReport> {
  const effectiveClaudeKey = (apiKey || process.env.ANTHROPIC_API_KEY || '').trim();
  const groqConfigured = getGroqKeyPoolStatus().configuredCount > 0;
  if (!groqConfigured && !effectiveClaudeKey) {
    return report.narrativeReport;
  }

  try {
    const promptStats = {
      dateRangeLabel: report.kpis.dateRangeLabel,
      totalEmails: report.kpis.totalEmails,
      receivedCount: report.kpis.receivedCount,
      sentCount: report.kpis.sentCount,
      unrepliedIncomingCount: report.kpis.unrepliedIncomingCount,
      realAttachmentCount: report.kpis.realAttachmentCount,
      topWeekly: report.weeklyTable.slice(0, 4),
      topProjects: report.projectRows.slice(0, 6).map((p) => ({
        projectCode: p.projectCode,
        total: p.total,
        received: p.received,
        sent: p.sent,
        unrepliedIncoming: p.unrepliedIncoming,
        attachmentCount: p.attachmentCount,
      })),
      topTopics: report.topicRows.slice(0, 6),
    };

    const systemPrompt =
      '당신은 조선해양·엔지니어링 메일 분석 총괄 어시스턴트입니다. 주어진 집계 통계(JSON)만을 바탕으로 과장 없이 정돈된 한국어 주간/월간 분석 보고서 본문을 JSON({"executiveSummary":"...","periodTrendNote":"...","projectFocusNote":"...","unrepliedRiskNote":"...","recommendedActions":["..."]})으로 작성하세요.';
    const userPrompt = JSON.stringify(promptStats);

    let rawText = '';
    let generatedBy = '';

    if (groqConfigured) {
      try {
        const groqRes = await invokeGroqStreamingCompletion({
          systemPrompt,
          userPrompt: userPrompt.slice(0, 6000),
          jsonMode: true,
          maxCompletionTokens: 1600,
          feature: 'executive-narrative',
        });
        rawText = groqRes.text;
        generatedBy = groqRes.model;
      } catch {
        // Fail over to Claude below
      }
    }

    if (!rawText && effectiveClaudeKey) {
      const res = await invokeClaudeMessagesApi({
        apiKey: effectiveClaudeKey,
        routing: {
          taskKind: 'mail-reduce',
          promptChars: userPrompt.length,
          customFocus: '주간 월간 보고서 분석',
        },
        systemPrompt,
        userPrompt,
        maxTokens: 1600,
      });
      rawText = res.text;
      generatedBy = `Claude · ${res.model}`;
    }

    const jsonMatch = rawText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return report.narrativeReport;
    const parsed = JSON.parse(jsonMatch[0]) as Partial<ExecutiveNarrativeReport>;

    const next: ExecutiveNarrativeReport = {
      ...report.narrativeReport,
      executiveSummary: parsed.executiveSummary || report.narrativeReport.executiveSummary,
      periodTrendNote: parsed.periodTrendNote || report.narrativeReport.periodTrendNote,
      projectFocusNote: parsed.projectFocusNote || report.narrativeReport.projectFocusNote,
      unrepliedRiskNote: parsed.unrepliedRiskNote || report.narrativeReport.unrepliedRiskNote,
      recommendedActions:
        Array.isArray(parsed.recommendedActions) && parsed.recommendedActions.length > 0
          ? parsed.recommendedActions.map(String)
          : report.narrativeReport.recommendedActions,
      generatedBy: generatedBy || 'Groq · openai/gpt-oss-120b',
      generatedAt: new Date().toISOString(),
    };

    next.markdownFull = [
      `# ${next.title}`,
      '',
      `> 생성 엔진: ${next.generatedBy}`,
      '',
      '## 1. 종합 브리핑 (분석글)',
      next.executiveSummary,
      '',
      next.periodTrendNote,
      '',
      next.projectFocusNote,
      '',
      next.unrepliedRiskNote,
      '',
      '## 2. 권장 실행 과제',
      ...next.recommendedActions.map((a, idx) => `${idx + 1}. ${a}`),
    ].join('\n');

    return next;
  } catch {
    return report.narrativeReport;
  }
}

export function saveExecutiveReportSnapshot(
  db: LocalMailDb,
  report: ExecutiveDashboardReport,
  customTitle?: string
): SavedExecutiveReportSnapshot[] {
  const existing =
    db.getState<SavedExecutiveReportSnapshot[]>('saved_executive_reports') || [];
  const snapshot: SavedExecutiveReportSnapshot = {
    id: `exec-${crypto.randomUUID().slice(0, 8)}`,
    title: customTitle?.trim() || report.narrativeReport.title,
    savedAt: new Date().toISOString(),
    periodType: report.filters.periodType,
    dateRangeLabel: report.kpis.dateRangeLabel,
    projectCode: report.filters.projectCode,
    topic: report.filters.topic,
    kpis: report.kpis,
    narrativeMarkdown: report.narrativeReport.markdownFull,
  };
  const next = [snapshot, ...existing].slice(0, 30);
  db.setState('saved_executive_reports', next);
  return next;
}
