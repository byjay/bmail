import type { MailItem } from '../types/mail';
import { detectProjectCode, type ProjectDefinition } from './deterministicAnalyzer';
import { isPromotionalEmail } from './mailFilter';
import {
  EMPTY_PROJECT_MAIL_OVERRIDES,
  type ProjectMailOverrides,
} from './outlookWorkspaceModel';

export type MailPriorityLevel = 'urgent' | 'important' | 'normal';
export type MailPriorityLabel = '긴급' | '중요' | '일반';

export interface MailPriorityAssessment {
  level: MailPriorityLevel;
  label: MailPriorityLabel;
  rank: number; // 3 = urgent, 2 = important, 1 = normal
  score: number; // 0..100 deterministic priority score
  matchedProjectCode: string | null;
  reasons: string[];
}

export interface MailPriorityContext {
  projects?: ProjectDefinition[];
  projectMailOverrides?: ProjectMailOverrides;
  ruleAssignedProjectCode?: string | null;
  ruleShouldFlag?: boolean;
  drawingRefCount?: number;
}

const URGENT_KEYWORDS =
  /(긴급|초긴급|즉시|오늘까지|금일\s*중|금일\s*마감|마감\s*임박|기한\s*초과|지연\s*발생|장애\s*발생|클레임|긴급\s*수정|긴급\s*요청|urgent|asap|critical|overdue|action\s*required|immediate|escalation)/i;

const DEADLINE_PRESSURE =
  /(마감|납기|오늘\s*오후|금일\s*내|D-[0-2]\b|시급|재요청|독촉|deadline|due\s*today|eod\b)/i;

const IMPORTANT_BUSINESS_KEYWORDS =
  /(결재|승인|도면|출도|rev\.?\s*[a-z0-9]|견적|발주|계약|검토\s*요청|회신\s*요청|보고서|회의록|사양서|선주|선급|class\s*comment|comment|specification|approval|purchase\s*order|rfq|quotation|contract)/i;

/**
 * Pure deterministic classifier that assigns '긴급' | '중요' | '일반' tags
 * based on content urgency/deadline signals and project relevance.
 */
export function classifyMailPriority(
  mail: MailItem,
  context: MailPriorityContext = {}
): MailPriorityAssessment {
  const projects = Array.isArray(context.projects) ? context.projects : [];
  const overrides: ProjectMailOverrides =
    context.projectMailOverrides ?? EMPTY_PROJECT_MAIL_OVERRIDES;
  const subject = mail.subject || '';
  const bodySnippet = `${mail.snippet || ''} ${(mail.body || '').slice(0, 600)}`;
  const fullText = `${subject} ${bodySnippet} ${mail.folderName || mail.category || ''}`;

  const promotional = isPromotionalEmail(mail);
  const reasons: string[] = [];
  let score = 10;

  // 1. Project relevance check (explicit override > rule assignment > deterministic project code detection)
  let matchedProjectCode: string | null = null;
  if (context.ruleAssignedProjectCode) {
    matchedProjectCode = context.ruleAssignedProjectCode;
  } else {
    for (const [code, ids] of Object.entries(overrides.added ?? {})) {
      if (Array.isArray(ids) && ids.includes(mail.id)) {
        matchedProjectCode = code;
        break;
      }
    }
  }

  if (!matchedProjectCode) {
    const detected = detectProjectCode(fullText, projects);
    if (detected) {
      const removedByOverride = overrides.removed?.[detected]?.includes(mail.id);
      if (!removedByOverride) {
        matchedProjectCode = detected;
      }
    }
  }

  if (matchedProjectCode) {
    score += 26;
    reasons.push(`프로젝트 [${matchedProjectCode}] 연관`);
  }

  // 2. SEDMS / Engineering drawing reference boost
  if ((context.drawingRefCount ?? 0) > 0) {
    score += 12;
    reasons.push(`도면 번호 ${context.drawingRefCount}건 감지`);
  }

  // 3. Explicit urgency & deadline signals
  const hasUrgentInSubject = URGENT_KEYWORDS.test(subject);
  const hasUrgentInBody = !hasUrgentInSubject && URGENT_KEYWORDS.test(bodySnippet);
  const hasDeadlinePressure = DEADLINE_PRESSURE.test(fullText);

  if (hasUrgentInSubject) {
    score += 56;
    reasons.push('제목 내 긴급 키워드');
  } else if (hasUrgentInBody) {
    score += 44;
    reasons.push('본문 내 긴급 요청');
  }

  if (hasDeadlinePressure) {
    score += 26;
    reasons.push('납기·마감 임박 신호');
  }

  // 4. Business / decision / approval importance keywords
  if (IMPORTANT_BUSINESS_KEYWORDS.test(fullText)) {
    score += 20;
    reasons.push('승인·검토·도면 핵심 업무');
  }

  // 5. User / rule flag state
  const isFlagged = Boolean(mail.isFlagged || context.ruleShouldFlag);
  if (isFlagged) {
    score += 24;
    reasons.push('플래그 지정됨');
  }

  // 6. Unread inbox + attachment slight contextual lift
  if (!promotional && mail.folder === 'inbox' && !mail.isRead) {
    score += 6;
  }
  if (!promotional && mail.hasAttachment) {
    score += 6;
    reasons.push('첨부 문서 포함');
  }

  // 7. Reserve '긴급' (>=70) strictly for explicit urgency, deadline pressure, or flagged escalations
  const hasEscalationSignal =
    hasUrgentInSubject || hasUrgentInBody || hasDeadlinePressure || isFlagged;
  if (!hasEscalationSignal) {
    score = Math.min(score, 68);
  }

  // 8. Promotional / automated newsletter damping unless explicitly flagged or urgent in subject
  if (promotional && !isFlagged && !hasUrgentInSubject) {
    score = Math.min(score, 20);
    reasons.length = 0;
    reasons.push('일반 안내·뉴스레터');
  }

  const clampedScore = Math.max(0, Math.min(100, Math.round(score)));

  let level: MailPriorityLevel;
  let label: MailPriorityLabel;
  let rank: number;

  if (clampedScore >= 70) {
    level = 'urgent';
    label = '긴급';
    rank = 3;
  } else if (clampedScore >= 36) {
    level = 'important';
    label = '중요';
    rank = 2;
  } else {
    level = 'normal';
    label = '일반';
    rank = 1;
  }

  if (reasons.length === 0) {
    reasons.push(level === 'normal' ? '일반 업무 메일' : '업무 중요도 기준 충족');
  }

  return {
    level,
    label,
    rank,
    score: clampedScore,
    matchedProjectCode,
    reasons,
  };
}

/**
 * Deterministic comparator for priority sorting:
 * Urgent (긴급) -> Important (중요) -> Normal (일반), then score desc, then timestamp desc.
 */
export function compareMailsByPriority(
  a: MailItem,
  b: MailItem,
  priorityById: Map<string, MailPriorityAssessment>
): number {
  const pa = priorityById.get(a.id);
  const pb = priorityById.get(b.id);
  const rankDiff = (pb?.rank ?? 1) - (pa?.rank ?? 1);
  if (rankDiff !== 0) return rankDiff;

  const scoreDiff = (pb?.score ?? 0) - (pa?.score ?? 0);
  if (scoreDiff !== 0) return scoreDiff;

  return b.timestamp.localeCompare(a.timestamp);
}
