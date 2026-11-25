import type { MailItem } from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import { cleanEmailBody } from '../utils/mailFilter';
import { classifyMailPriority } from '../utils/mailPriority';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  extractShipbuildingDrawingRefs,
  type ProjectMailOverrides,
} from '../utils/outlookWorkspaceModel';
import { classifyByProjectRoutingRules } from '../utils/projectFileOrganizer';

export type UnreadUrgencyTier = 'urgent-reply' | 'important' | 'normal';

export interface UnreadBriefingItem {
  emailId: string;
  subject: string;
  fromName: string;
  fromEmail: string;
  date: string;
  timeLabel: string;
  timestamp: string;
  projectCode: string;
  projectName: string;
  urgencyTier: UnreadUrgencyTier;
  urgencyLabel: '반드시 긴급회신' | '중요' | '일반';
  priorityScore: number;
  mustReply: boolean;
  reasons: string[];
  drawings: string[];
  deadlines: string[];
  summary: string;
}

export interface UnreadProjectGroup {
  projectCode: string;
  projectName: string;
  totalCount: number;
  urgentReplyCount: number;
  importantCount: number;
  normalCount: number;
  highestScore: number;
  items: UnreadBriefingItem[];
}

export const DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS = 14;
export const DEFAULT_URGENT_REPLY_MAX_AGE_DAYS = 14;

export interface UnreadPriorityBriefing {
  generatedAt: string;
  maxAgeDays: number;
  excludedOldUnreadCount: number;
  totalUnread: number;
  urgentReplyCount: number;
  importantCount: number;
  normalCount: number;
  projectCount: number;
  briefingSummary: string;
  urgentReplyAlerts: UnreadBriefingItem[];
  projectGroups: UnreadProjectGroup[];
}

const REPLY_ACTION_RE =
  /(회신\s*(요망|요청|바랍니다|부탁)|답변\s*(요망|요청|바랍니다|부탁)|검토\s*(후|요청|바랍니다)|승인\s*(요청|대기|미결)|미결\s*사항|조치\s*요망|확인\s*(후|요청|바랍니다)|urgent|asap|금일\s*중|금일\s*까지|오늘\s*까지|즉시|please\s+(?:advise|reply|confirm)|action\s+required)/i;

const DEADLINE_DATE_RE = /\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b/g;
const DEADLINE_KR_RE = /\b\d{1,2}\s*월\s*\d{1,2}\s*일\b/g;

function tierRank(tier: UnreadUrgencyTier): number {
  if (tier === 'urgent-reply') return 3;
  if (tier === 'important') return 2;
  return 1;
}

/**
 * Pure read-only triage of RECENT unread inbox emails (default within 14 days so archived old mails
 * from previous months/years are never mistakenly surfaced as urgent/important):
 * - Never mutates `isRead` or marks items as read
 * - Identifies mandatory urgent-reply items (strictly recent)
 * - Groups and sorts by Project and Importance/Urgency
 */
export function buildUnreadPriorityBriefing(
  emails: MailItem[],
  projects: ProjectDefinition[] = [],
  overrides: ProjectMailOverrides = EMPTY_PROJECT_MAIL_OVERRIDES,
  options?: { maxAgeDays?: number; nowMs?: number }
): UnreadPriorityBriefing {
  const maxAgeDays =
    typeof options?.maxAgeDays === 'number' && Number.isFinite(options.maxAgeDays)
      ? Math.max(0, options.maxAgeDays)
      : DEFAULT_UNREAD_BRIEFING_MAX_AGE_DAYS;
  const nowMs =
    typeof options?.nowMs === 'number' && Number.isFinite(options.nowMs)
      ? options.nowMs
      : Date.now();
  const maxAgeMs = maxAgeDays * 86_400_000;
  const urgentCutoffMs = DEFAULT_URGENT_REPLY_MAX_AGE_DAYS * 86_400_000;

  let excludedOldUnreadCount = 0;
  const unreadInbox = emails.filter((m) => {
    if (!m || m.folder !== 'inbox' || m.isRead) return false;
    if (maxAgeDays <= 0) return true;
    const rawTs = String(m.timestamp || m.date || '').trim();
    const parsedMs = rawTs ? Date.parse(rawTs) : Number.NaN;
    if (!Number.isFinite(parsedMs) || nowMs - parsedMs > maxAgeMs) {
      excludedOldUnreadCount += 1;
      return false;
    }
    return true;
  });

  const projectMap = new Map(projects.map((p) => [p.code, p]));
  const items: UnreadBriefingItem[] = [];

  for (const mail of unreadInbox) {
    const rawTs = String(mail.timestamp || mail.date || '').trim();
    const parsedMs = rawTs ? Date.parse(rawTs) : Number.NaN;
    const isWithinUrgentWindow = Number.isFinite(parsedMs) && nowMs - parsedMs <= urgentCutoffMs;
    const subject = String(mail.subject || '(제목 없음)').trim();
    const cleanedBody = cleanEmailBody(String(mail.body || mail.snippet || ''), 700)
      .replace(/\s+/g, ' ')
      .trim();
    const combined = `${subject}\n${cleanedBody}`;
    const drawings = extractShipbuildingDrawingRefs(combined)
      .slice(0, 4)
      .map((d) => `${d.raw}${d.revision ? ` (${d.revision})` : ''}`);
    const deadlines = Array.from(
      new Set([
        ...(combined.match(DEADLINE_DATE_RE) || []),
        ...(combined.match(DEADLINE_KR_RE) || []),
      ])
    ).slice(0, 2);

    const routing = classifyByProjectRoutingRules({
      senderEmail: mail.fromEmail,
      recipientEmails: mail.to,
      subject,
      body: cleanedBody,
      category: mail.category || mail.folderName,
      filenames: (mail.attachments || []).map((a) => a.filename),
      projects,
    });

    const projectCode =
      assessmentCodeFallback(mail, projects, overrides, routing.projectCode, drawings.length);
    void projectCode;
    const assessment = classifyMailPriority(mail, {
      projects,
      projectMailOverrides: overrides,
      ruleAssignedProjectCode: routing.projectCode !== 'UNASSIGNED' ? routing.projectCode : null,
      drawingRefCount: drawings.length,
    });

    const resolvedProjectCode =
      assessment.matchedProjectCode ||
      (routing.projectCode !== 'UNASSIGNED' ? routing.projectCode : 'GENERAL');
    const matchedProj = projectMap.get(resolvedProjectCode);
    const projectName =
      matchedProj?.name ||
      (resolvedProjectCode === 'GENERAL'
        ? '일반 · 기타 업무'
        : routing.projectCode === resolvedProjectCode
          ? routing.projectName
          : resolvedProjectCode);

    const needsReply = REPLY_ACTION_RE.test(combined);
    const isUrgentReply =
      isWithinUrgentWindow &&
      (assessment.level === 'urgent' || (assessment.score >= 52 && needsReply));
    const isImportant = !isUrgentReply && (assessment.level === 'important' || needsReply || mail.isFlagged);

    const urgencyTier: UnreadUrgencyTier = isUrgentReply
      ? 'urgent-reply'
      : isImportant
        ? 'important'
        : 'normal';
    const urgencyLabel: UnreadBriefingItem['urgencyLabel'] =
      urgencyTier === 'urgent-reply'
        ? '반드시 긴급회신'
        : urgencyTier === 'important'
          ? '중요'
          : '일반';

    const reasons = [...assessment.reasons];
    if (needsReply && !reasons.some((r) => r.includes('회신'))) {
      reasons.push('회신/조치 요구 표현 감지');
    }
    if (deadlines.length > 0 && !reasons.some((r) => r.includes('기한'))) {
      reasons.push(`기한 언급 (${deadlines[0]})`);
    }

    const tags: string[] = [];
    if (drawings.length > 0) tags.push(`도면 ${drawings.slice(0, 2).join(', ')}`);
    if (deadlines.length > 0) tags.push(`기한 ${deadlines[0]}`);
    const prefix = tags.length > 0 ? `[${tags.join(' · ')}] ` : '';
    const snippetText = (cleanedBody || subject).slice(0, 95);
    const summary = `${prefix}${snippetText}${cleanedBody.length > 95 ? '…' : ''}`;

    items.push({
      emailId: mail.id,
      subject,
      fromName: String(mail.fromName || mail.fromEmail || '발신자').trim(),
      fromEmail: String(mail.fromEmail || '').trim(),
      date: mail.date,
      timeLabel: mail.timeLabel,
      timestamp: mail.timestamp,
      projectCode: resolvedProjectCode,
      projectName,
      urgencyTier,
      urgencyLabel,
      priorityScore: assessment.score + (needsReply ? 12 : 0),
      mustReply: urgencyTier === 'urgent-reply',
      reasons,
      drawings,
      deadlines,
      summary,
    });
  }

  items.sort(
    (a, b) =>
      tierRank(b.urgencyTier) - tierRank(a.urgencyTier) ||
      b.priorityScore - a.priorityScore ||
      String(b.timestamp || '').localeCompare(String(a.timestamp || ''))
  );

  const urgentReplyAlerts = items.filter((item) => item.urgencyTier === 'urgent-reply');
  const importantCount = items.filter((item) => item.urgencyTier === 'important').length;
  const normalCount = items.filter((item) => item.urgencyTier === 'normal').length;

  const groupMap = new Map<string, UnreadProjectGroup>();
  for (const item of items) {
    let group = groupMap.get(item.projectCode);
    if (!group) {
      group = {
        projectCode: item.projectCode,
        projectName: item.projectName,
        totalCount: 0,
        urgentReplyCount: 0,
        importantCount: 0,
        normalCount: 0,
        highestScore: item.priorityScore,
        items: [],
      };
      groupMap.set(item.projectCode, group);
    }
    group.totalCount += 1;
    if (item.urgencyTier === 'urgent-reply') group.urgentReplyCount += 1;
    else if (item.urgencyTier === 'important') group.importantCount += 1;
    else group.normalCount += 1;
    group.highestScore = Math.max(group.highestScore, item.priorityScore);
    group.items.push(item);
  }

  const projectGroups = Array.from(groupMap.values()).sort(
    (a, b) =>
      b.urgentReplyCount - a.urgentReplyCount ||
      b.importantCount - a.importantCount ||
      b.highestScore - a.highestScore ||
      b.totalCount - a.totalCount
  );

  const briefingSummary =
    items.length === 0
      ? `최근 ${maxAgeDays}일 내 확인하지 않은 수신 메일이 없습니다.`
      : `최근 ${maxAgeDays}일 미확인(안 읽음 유지) ${items.length}건 중 반드시 긴급회신 ${urgentReplyAlerts.length}건 · 중요 ${importantCount}건 · ${projectGroups.length}개 프로젝트로 정리되었습니다.`;

  return {
    generatedAt: new Date().toISOString(),
    maxAgeDays,
    excludedOldUnreadCount,
    totalUnread: items.length,
    urgentReplyCount: urgentReplyAlerts.length,
    importantCount,
    normalCount,
    projectCount: projectGroups.length,
    briefingSummary,
    urgentReplyAlerts,
    projectGroups,
  };
}

function assessmentCodeFallback(
  _mail: MailItem,
  _projects: ProjectDefinition[],
  _overrides: ProjectMailOverrides,
  _ruleCode: string,
  _drawingsLen: number
): string {
  return '';
}
