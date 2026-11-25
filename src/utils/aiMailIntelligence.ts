import type { MailItem } from '../types/mail';
import {
  type CalendarScheduleEvent,
  detectProjectCode,
  normalizeConversationSubject,
  type ProjectDefinition,
} from './deterministicAnalyzer';
import { isPromotionalEmail, maskSensitiveData } from './mailFilter';

export type MailIntentCategory =
  | 'action_required'
  | 'decision_needed'
  | 'waiting_on'
  | 'sales_quote'
  | 'drawing_tech'
  | 'meeting_schedule'
  | 'fyi_newsletter';

export type MailUrgencyLevel = 'urgent' | 'high' | 'normal' | 'low';

export interface CustomAutoLabelRule {
  id: string;
  name: string;
  promptDescription: string;
  keywords: string[];
  enabled: boolean;
}

export interface MailIntelligenceProfile {
  emailId: string;
  intent: MailIntentCategory;
  intentLabel: string;
  urgency: MailUrgencyLevel;
  urgencyLabel: string;
  isForeignLanguage: boolean;
  detectedLanguage: 'ko' | 'en' | 'ja' | 'zh' | 'mixed';
  oneLineSummary: string;
  actionRequiredText: string | null;
  riskFlags: string[];
  extractedAmounts: string[];
  extractedDeadlines: string[];
  extractedDrawingNos: string[];
  matchedLabels: string[];
}

export interface WaitingOnItem {
  emailId: string;
  subject: string;
  recipientLabel: string;
  sentDate: string;
  daysWaiting: number;
  reminderSubject: string;
  reminderBody: string;
}

export interface ExtractedCommitment {
  id: string;
  emailId: string;
  owner: string;
  task: string;
  deadline: string;
  sourceSubject: string;
  sourceDate: string;
}

export interface MorningBriefReport {
  generatedAt: string;
  totalAnalyzed: number;
  headlineSummary: string;
  actionRequired: Array<{ mail: MailItem; profile: MailIntelligenceProfile }>;
  decisionNeeded: Array<{ mail: MailItem; profile: MailIntelligenceProfile }>;
  waitingOnOthers: WaitingOnItem[];
  fyiSweepable: Array<{ mail: MailItem; profile: MailIntelligenceProfile }>;
  commitments: ExtractedCommitment[];
  labelCounts: Array<{ label: string; count: number; emailIds: string[] }>;
  keyHighlights: string[];
}

export interface InstantReplyOption {
  id: 'confirm' | 'clarify' | 'schedule';
  label: string;
  intentHint: string;
  subject: string;
  body: string;
}

export interface AskInboxCitation {
  index: number;
  emailId: string;
  subject: string;
  fromName: string;
  date: string;
  snippet: string;
}

export interface AskInboxResponse {
  question: string;
  answer: string;
  bulletFindings: string[];
  citations: AskInboxCitation[];
  matchedCount: number;
}

export interface DraftCoachingWarning {
  severity: 'high' | 'medium' | 'info';
  code:
    | 'missing_attachment'
    | 'sensitive_data'
    | 'missing_subject'
    | 'missing_recipient'
    | 'harsh_tone'
    | 'missing_cta'
    | 'too_long';
  title: string;
  suggestion: string;
}

export interface DraftCoachingReport {
  readyToSend: boolean;
  warnings: DraftCoachingWarning[];
  wordCount: number;
  estimatedReadSec: number;
  toneSummary: string;
}

export const DEFAULT_AUTO_LABEL_RULES: CustomAutoLabelRule[] = [
  {
    id: 'lbl-urgent',
    name: '긴급 · 즉시 회신',
    promptDescription: '긴급 요청, 오늘 마감, 즉시 회신 또는 승인이 필요한 메일',
    keywords: ['긴급', 'urgent', 'asap', '금일', '회신 부탁', '승인 요청', '결재', 'action required'],
    enabled: true,
  },
  {
    id: 'lbl-sales',
    name: '영업 · 견적 · 발주',
    promptDescription: 'RFQ, 견적서, 단가, 발주(PO), 납기 및 해외 바이어 영업 메일',
    keywords: ['견적', '발주', '단가', '계약', 'quotation', 'rfq', 'po', 'price', 'delivery', 'invoice', 'quote'],
    enabled: true,
  },
  {
    id: 'lbl-drawing',
    name: '도면 · 출도 · 기술검토',
    promptDescription: '도면 번호, Rev 개정, 출도, 사양 검토, 기술 코멘트 메일',
    keywords: ['도면', '출도', '개정', 'rev', 'dwg', 'drawing', 'comment', 'spec', '시방', '전장'],
    enabled: true,
  },
  {
    id: 'lbl-meeting',
    name: '일정 · 회의 · 킥오프',
    promptDescription: '회의 소집, 미팅 일정, 화상 회의 링크, 방문 일정 조율 메일',
    keywords: ['회의', '미팅', '일정', 'meeting', 'schedule', 'kickoff', 'zoom', 'teams', '참석'],
    enabled: true,
  },
  {
    id: 'lbl-foreign',
    name: '외국어 · 해외 수신',
    promptDescription: '영문/일문/중문 등 번역 및 다국어 회신 검토가 필요한 메일',
    keywords: ['dear ', 'best regards', 'please find attached', 'thank you', 'kind regards', 'sincerely'],
    enabled: true,
  },
];

const INTENT_LABELS: Record<MailIntentCategory, string> = {
  action_required: '즉시 회신·조치',
  decision_needed: '의사결정·검토',
  waiting_on: '상대 회신 대기',
  sales_quote: '영업·견적·발주',
  drawing_tech: '도면·기술검토',
  meeting_schedule: '일정·회의',
  fyi_newsletter: '참고·자동정리',
};

const URGENCY_LABELS: Record<MailUrgencyLevel, string> = {
  urgent: '긴급',
  high: '높음',
  normal: '보통',
  low: '낮음',
};

function detectLanguage(text: string): {
  lang: MailIntelligenceProfile['detectedLanguage'];
  isForeign: boolean;
} {
  const sample = String(text || '').slice(0, 1500);
  const koCount = (sample.match(/[가-힣]/g) || []).length;
  const enCount = (sample.match(/[a-zA-Z]/g) || []).length;
  const jaCount = (sample.match(/[\u3040-\u30ff]/g) || []).length;
  const zhCount = (sample.match(/[\u4e00-\u9fff]/g) || []).length;

  if (jaCount >= 8 && jaCount > koCount) return { lang: 'ja', isForeign: true };
  if (zhCount >= 8 && koCount < 5) return { lang: 'zh', isForeign: true };
  if (enCount >= 40 && koCount < 10) return { lang: 'en', isForeign: true };
  if (enCount >= 40 && koCount >= 10) return { lang: 'mixed', isForeign: enCount > koCount * 3 };
  return { lang: 'ko', isForeign: false };
}

function extractAmounts(text: string): string[] {
  const matches = text.match(
    /(?:USD|EUR|JPY|KRW|\$|€|₩)\s?[0-9,]+(?:\.[0-9]+)?(?:\s?(?:K|M|억|만))?|[0-9,]+(?:\.[0-9]+)?\s?(?:원|만원|억원|달러|USD|EUR)/gi
  );
  if (!matches) return [];
  return Array.from(new Set(matches.map((m) => m.trim()))).slice(0, 5);
}

function extractDeadlines(text: string): string[] {
  const results: string[] = [];
  const patterns = [
    /(20\d{2}[-./]\d{1,2}[-./]\d{1,2})/g,
    /(\d{1,2}월\s*\d{1,2}일(?:\s*(?:까지|오전|오후|\d{1,2}시))?)/g,
    /\b(?:by|before|due|deadline)\s+([A-Z][a-z]{2,8}\s+\d{1,2}(?:,\s*20\d{2})?|\d{1,2}\/\d{1,2}|Friday|Monday|Tuesday|Wednesday|Thursday|EOD|COB)/gi,
    /(금일\s*중|오늘\s*(?:까지|오후\s*\d{1,2}시)|내일\s*(?:오전|오후|까지)|이번\s*주\s*(?:금요일|내))/g,
  ];
  for (const pat of patterns) {
    let m: RegExpExecArray | null;
    while ((m = pat.exec(text)) !== null) {
      const val = (m[1] || m[0]).trim();
      if (val && !results.includes(val)) results.push(val);
      if (results.length >= 4) break;
    }
    if (results.length >= 4) break;
  }
  return results;
}

function extractDrawingNos(text: string): string[] {
  const matches = text.match(/\b(?:DWG|DRW|EL|ME|PI|HU)[-_]?[A-Z0-9]{2,8}(?:[-_][A-Z0-9]{1,5})?\b/gi);
  if (!matches) return [];
  return Array.from(new Set(matches.map((m) => m.toUpperCase()))).slice(0, 5);
}

function buildOneLineSummary(mail: MailItem, intent: MailIntentCategory, deadlines: string[], amounts: string[]): string {
  const cleanBody = String(mail.body || '')
    .replace(/\r/g, '')
    .split('\n')
    .map((line) => line.trim())
    .filter(
      (line) =>
        line.length >= 8 &&
        !/^(안녕하세요|수고하십니다|감사합니다|dear\s+|hi\s+|hello\s+|best regards|sincerely|kind regards|thanks)/i.test(
          line
        ) &&
        !line.startsWith('>')
    );

  const coreSentence = (cleanBody[0] || mail.subject || '(본문 없음)').slice(0, 110);
  const extras: string[] = [];
  if (deadlines[0]) extras.push(`기한 ${deadlines[0]}`);
  if (amounts[0]) extras.push(`금액 ${amounts[0]}`);
  const suffix = extras.length > 0 ? ` (${extras.join(' · ')})` : '';
  return `[${INTENT_LABELS[intent]}] ${coreSentence}${suffix}`;
}

export function classifyEmailIntent(
  mail: MailItem,
  customRules: CustomAutoLabelRule[] = DEFAULT_AUTO_LABEL_RULES
): MailIntelligenceProfile {
  const subject = String(mail.subject || '');
  const body = String(mail.body || '');
  const attachmentsText = (mail.attachments || []).map((a) => a.filename).join(' ');
  const combined = `${subject}\n${body}\n${attachmentsText}`;
  const lower = combined.toLowerCase();

  const { lang, isForeign } = detectLanguage(combined);
  const extractedAmounts = extractAmounts(combined);
  const extractedDeadlines = extractDeadlines(combined);
  const extractedDrawingNos = extractDrawingNos(combined);

  const isPromo = isPromotionalEmail(mail);

  const hasUrgentCue = /(긴급|urgent|asap|즉시|장애|클레임|claim|penalty|지연\s*경고|금일\s*마감|eod|cob)/i.test(
    lower
  );
  const hasActionCue =
    /(회신\s*부탁|답변\s*부탁|확인\s*부탁|검토\s*후\s*회신|승인\s*요청|결재|please\s+(confirm|reply|advise|review|approve|send)|kindly\s+(confirm|provide|reply)|action\s+required|\?)/i.test(
      combined
    );
  const hasDecisionCue =
    /(견적|단가|계약|발주|승인|의사결정|비교표|quotation|rfq|proposal|approve|budget|purchase order|\bpo\b)/i.test(
      lower
    );
  const hasDrawingCue =
    extractedDrawingNos.length > 0 ||
    /(도면|출도|개정|rev\.?\s*\d|시방서|drawing|specification|transmittal)/i.test(lower);
  const hasMeetingCue = /(회의|미팅|일정\s*조율|킥오프|화상\s*회의|meeting|zoom|teams|calendar|invite)/i.test(
    lower
  );

  let intent: MailIntentCategory = 'action_required';
  if (isPromo) {
    intent = 'fyi_newsletter';
  } else if (mail.folder === 'sent') {
    intent = 'waiting_on';
  } else if (hasUrgentCue || (hasActionCue && !hasDecisionCue && !hasMeetingCue)) {
    intent = 'action_required';
  } else if (hasDecisionCue) {
    intent = extractedAmounts.length > 0 || /(견적|quotation|rfq|발주|po\b)/i.test(lower) ? 'sales_quote' : 'decision_needed';
  } else if (hasDrawingCue) {
    intent = 'drawing_tech';
  } else if (hasMeetingCue) {
    intent = 'meeting_schedule';
  } else if (hasActionCue) {
    intent = 'action_required';
  } else {
    intent = 'fyi_newsletter';
  }

  let urgency: MailUrgencyLevel = 'normal';
  if (hasUrgentCue) {
    urgency = 'urgent';
  } else if (
    intent === 'action_required' ||
    intent === 'decision_needed' ||
    intent === 'sales_quote' ||
    extractedDeadlines.length > 0
  ) {
    urgency = 'high';
  } else if (intent === 'fyi_newsletter') {
    urgency = 'low';
  }

  const riskFlags: string[] = [];
  if (/(지연|delay|overdue|납기\s*초과|연기)/i.test(lower)) riskFlags.push('납기·일정 지연 리스크');
  if (/(클레임|claim|penalty|배상|불량|defect|하자)/i.test(lower)) riskFlags.push('클레임·품질 이슈');
  if (/(충돌|conflict|수정\s*요청|재검토|rev\.?\s*변경|reject)/i.test(lower)) riskFlags.push('개정·재검토 필요');
  if (isForeign) riskFlags.push(`외국어(${lang.toUpperCase()}) 번역 권장`);
  if (extractedDeadlines.length > 0) riskFlags.push(`기한 감지(${extractedDeadlines[0]})`);

  // Evaluate natural-language Auto-Label rules
  const matchedLabels: string[] = [];
  for (const rule of customRules) {
    if (!rule.enabled) continue;
    const kws = rule.keywords.map((k) => k.trim().toLowerCase()).filter(Boolean);
    if (kws.length === 0) continue;
    if (kws.some((kw) => lower.includes(kw))) {
      matchedLabels.push(rule.name);
    }
  }
  if (isForeign && !matchedLabels.some((l) => l.includes('외국어'))) {
    matchedLabels.push('외국어 · 해외 수신');
  }

  let actionRequiredText: string | null = null;
  if (intent === 'action_required') {
    actionRequiredText = extractedDeadlines[0]
      ? `${extractedDeadlines[0]}까지 발신자(${mail.fromName}) 요청 사항 확인 및 회신`
      : `발신자(${mail.fromName}) 질의·요청 사항 확인 및 회신 필요`;
  } else if (intent === 'sales_quote' || intent === 'decision_needed') {
    actionRequiredText = extractedAmounts[0]
      ? `견적/조건(${extractedAmounts[0]}) 검토 후 승인·회신 결정`
      : '제안·견적 조건 검토 및 의사결정 필요';
  } else if (intent === 'drawing_tech') {
    actionRequiredText = extractedDrawingNos[0]
      ? `도면(${extractedDrawingNos[0]}) 기술 검토 및 코멘트 반영`
      : '도면·기술 문서 검토 및 피드백 회신';
  } else if (intent === 'meeting_schedule') {
    actionRequiredText = extractedDeadlines[0]
      ? `${extractedDeadlines[0]} 회의 참석 여부 확정 및 일정 등록`
      : '회의 참석 가능 시간 회신 및 일정 등록';
  }

  return {
    emailId: mail.id,
    intent,
    intentLabel: INTENT_LABELS[intent],
    urgency,
    urgencyLabel: URGENCY_LABELS[urgency],
    isForeignLanguage: isForeign,
    detectedLanguage: lang,
    oneLineSummary: buildOneLineSummary(mail, intent, extractedDeadlines, extractedAmounts),
    actionRequiredText,
    riskFlags,
    extractedAmounts,
    extractedDeadlines,
    extractedDrawingNos,
    matchedLabels: Array.from(new Set(matchedLabels)),
  };
}

function diffCalendarDays(earlierIso: string, laterIso: string): number {
  const a = new Date(earlierIso).getTime();
  const b = new Date(laterIso).getTime();
  if (Number.isNaN(a) || Number.isNaN(b)) return 0;
  return Math.max(0, Math.floor((b - a) / (24 * 60 * 60 * 1000)));
}

/**
 * Generates a Genspark-surpassing Morning Executive Brief + Waiting-On SLA Radar + Commitment Tracker.
 */
export function buildMorningBrief(
  emails: MailItem[],
  customRules: CustomAutoLabelRule[] = DEFAULT_AUTO_LABEL_RULES,
  referenceIsoDate?: string
): MorningBriefReport {
  const nowIso = referenceIsoDate || new Date().toISOString();
  const safeEmails = Array.isArray(emails) ? emails : [];

  const actionRequired: MorningBriefReport['actionRequired'] = [];
  const decisionNeeded: MorningBriefReport['decisionNeeded'] = [];
  const fyiSweepable: MorningBriefReport['fyiSweepable'] = [];
  const commitments: ExtractedCommitment[] = [];

  // Index inbox timestamps by normalized subject to detect unanswered sent mails (Waiting-On Radar)
  const latestInboxByThread = new Map<string, string>();
  for (const mail of safeEmails) {
    if (mail.folder === 'inbox') {
      const key = (mail.conversationKey || normalizeConversationSubject(mail.subject)).toLowerCase();
      const prev = latestInboxByThread.get(key);
      if (!prev || mail.timestamp > prev) {
        latestInboxByThread.set(key, mail.timestamp);
      }
    }
  }

  const waitingOnOthers: WaitingOnItem[] = [];
  const labelMap = new Map<string, string[]>();
  const allHighlights: string[] = [];

  for (const mail of safeEmails) {
    const profile = classifyEmailIntent(mail, customRules);

    for (const lbl of profile.matchedLabels) {
      const ids = labelMap.get(lbl) ?? [];
      ids.push(mail.id);
      labelMap.set(lbl, ids);
    }

    if (mail.folder === 'sent') {
      const key = (mail.conversationKey || normalizeConversationSubject(mail.subject)).toLowerCase();
      const latestReplyTs = latestInboxByThread.get(key);
      const hasNewerReply = Boolean(latestReplyTs && latestReplyTs > mail.timestamp);
      const asksForResponse = /(\?|부탁|회신|확인|검토|요청|please|confirm|reply|quote|review)/i.test(
        `${mail.subject} ${mail.body}`
      );
      if (!hasNewerReply && asksForResponse) {
        const daysWaiting = diffCalendarDays(mail.timestamp || mail.date, nowIso);
        const recipientLabel = mail.to?.[0] || '수신자';
        waitingOnOthers.push({
          emailId: mail.id,
          subject: mail.subject,
          recipientLabel,
          sentDate: mail.date,
          daysWaiting,
          reminderSubject: mail.subject.startsWith('Re:') ? mail.subject : `Re: ${mail.subject}`,
          reminderBody: [
            `안녕하세요, ${recipientLabel}님.`,
            '',
            `지난 ${mail.date} 발송드린 "${mail.subject}" 건의 진행 현황을 확인하고자 다시 연락드립니다.`,
            '검토 완료 일정 또는 회신 가능 시점을 공유해 주시면 감사하겠습니다.',
            '',
            '감사합니다.',
          ].join('\n'),
        });
      }
      continue;
    }

    if (profile.intent === 'fyi_newsletter') {
      fyiSweepable.push({ mail, profile });
    } else if (profile.intent === 'sales_quote' || profile.intent === 'decision_needed') {
      decisionNeeded.push({ mail, profile });
    } else {
      actionRequired.push({ mail, profile });
    }

    if (profile.actionRequiredText) {
      commitments.push({
        id: `cmt-${mail.id}`,
        emailId: mail.id,
        owner: '본인',
        task: profile.actionRequiredText,
        deadline: profile.extractedDeadlines[0] || mail.date,
        sourceSubject: mail.subject,
        sourceDate: mail.date,
      });
    }

    if (profile.extractedAmounts.length > 0) {
      allHighlights.push(`${mail.subject.slice(0, 32)} · 금액 ${profile.extractedAmounts.join(', ')}`);
    }
    if (profile.extractedDrawingNos.length > 0) {
      allHighlights.push(`${mail.subject.slice(0, 32)} · 도면 ${profile.extractedDrawingNos.join(', ')}`);
    }
  }

  // Sort action/decision by urgency descending then newest first
  const urgencyRank: Record<MailUrgencyLevel, number> = { urgent: 0, high: 1, normal: 2, low: 3 };
  const sortItems = (
    a: { mail: MailItem; profile: MailIntelligenceProfile },
    b: { mail: MailItem; profile: MailIntelligenceProfile }
  ) =>
    urgencyRank[a.profile.urgency] - urgencyRank[b.profile.urgency] ||
    b.mail.timestamp.localeCompare(a.mail.timestamp);

  actionRequired.sort(sortItems);
  decisionNeeded.sort(sortItems);
  waitingOnOthers.sort((a, b) => b.daysWaiting - a.daysWaiting);

  const labelCounts = Array.from(labelMap.entries())
    .map(([label, emailIds]) => ({ label, count: emailIds.length, emailIds }))
    .sort((a, b) => b.count - a.count);

  const headlineSummary =
    safeEmails.length === 0
      ? '연동된 실제 메일이 없습니다. PST/EML을 열거나 Microsoft 365를 연결하면 즉시 모닝 브리프가 생성됩니다.'
      : `총 ${safeEmails.length}건 분석 완료 — 즉시 회신·조치 ${actionRequired.length}건, 견적·의사결정 ${decisionNeeded.length}건, 상대 회신 대기 ${waitingOnOthers.length}건, 자동 정리 후보 ${fyiSweepable.length}건입니다.`;

  return {
    generatedAt: nowIso,
    totalAnalyzed: safeEmails.length,
    headlineSummary,
    actionRequired,
    decisionNeeded,
    waitingOnOthers,
    fyiSweepable,
    commitments: commitments.slice(0, 20),
    labelCounts,
    keyHighlights: Array.from(new Set(allHighlights)).slice(0, 8),
  };
}

/**
 * Superhuman-surpassing Instant Reply generator: produces 3 full-sentence contextual drafts
 * matching the email's language (Korean or English) & domain intent. Never auto-sends.
 */
export function generateInstantReplyOptions(
  mail: MailItem,
  userSignature = ''
): InstantReplyOption[] {
  const profile = classifyEmailIntent(mail);
  const replySubject = /^re\s*:/i.test(mail.subject) ? mail.subject : `Re: ${mail.subject}`;
  const sig = userSignature.trim() ? `\n\n${userSignature.trim()}` : '';
  const sender = mail.fromName || mail.fromEmail || '담당자';
  const refNote = profile.extractedDrawingNos[0]
    ? ` (${profile.extractedDrawingNos[0]})`
    : profile.extractedAmounts[0]
      ? ` (${profile.extractedAmounts[0]})`
      : '';
  const deadlineNote = profile.extractedDeadlines[0] ? ` (${profile.extractedDeadlines[0]})` : '';

  if (profile.isForeignLanguage) {
    return [
      {
        id: 'confirm',
        label: '수락 · 진행 확정 (EN)',
        intentHint: '영문 긍정 회신 및 진행 확정',
        subject: replySubject,
        body: [
          `Dear ${sender},`,
          '',
          `Thank you for your email regarding "${mail.subject}"${refNote}.`,
          `We have reviewed the details and confirm we can proceed as requested${deadlineNote}. We will share the final confirmation and updated documents shortly.`,
          '',
          'Best regards,',
        ].join('\n') + sig,
      },
      {
        id: 'clarify',
        label: '조건 확인 · 추가 질의 (EN)',
        intentHint: '영문 견적/도면/납기 세부 조건 확인',
        subject: replySubject,
        body: [
          `Dear ${sender},`,
          '',
          `Thank you for sharing the details on "${mail.subject}"${refNote}.`,
          'Before finalizing our confirmation, could you please clarify the following items?',
          '1. Latest delivery lead time and commercial validity',
          '2. Technical revision status and required deliverables',
          '',
          'Looking forward to your prompt reply.',
          '',
          'Best regards,',
        ].join('\n') + sig,
      },
      {
        id: 'schedule',
        label: '검토 후 회신 · 일정 제안 (EN)',
        intentHint: '내부 검토 시간 확보 및 회신 일정 안내',
        subject: replySubject,
        body: [
          `Dear ${sender},`,
          '',
          `Thank you for your message regarding "${mail.subject}".`,
          'We are currently reviewing the technical and commercial points internally and will revert with our detailed feedback by tomorrow afternoon.',
          '',
          'Best regards,',
        ].join('\n') + sig,
      },
    ];
  }

  return [
    {
      id: 'confirm',
      label: '수락 · 진행 확정',
      intentHint: '요청 사항 승인 및 진행 일정 확정',
      subject: replySubject,
      body: [
        `안녕하세요, ${sender}님.`,
        '',
        `공유해 주신 "${mail.subject}"${refNote} 건 잘 확인했습니다.`,
        `요청하신 내용대로 진행 확정하며${deadlineNote ? `, ${profile.extractedDeadlines[0]} 일정에 맞춰` : ''} 후속 자료를 정리해 회신드리겠습니다.`,
        '',
        '감사합니다.',
      ].join('\n') + sig,
    },
    {
      id: 'clarify',
      label: '세부 조건 · 추가 확인',
      intentHint: '견적/도면/일정 쟁점 확인 질의',
      subject: replySubject,
      body: [
        `안녕하세요, ${sender}님.`,
        '',
        `보내주신 "${mail.subject}"${refNote} 관련하여 최종 확정 전 아래 사항 확인 부탁드립니다.`,
        '1. 최종 적용 사양(또는 개정 Rev) 및 첨부 기준 일치 여부',
        '2. 납기/회신 마감 일정 및 비용 변동 여부',
        '',
        '확인 회신 주시는 대로 즉시 다음 단계 진행하겠습니다.',
        '',
        '감사합니다.',
      ].join('\n') + sig,
    },
    {
      id: 'schedule',
      label: '내부 검토 · 일정 조율',
      intentHint: '검토 시간 확보 또는 미팅 제안',
      subject: replySubject,
      body: [
        `안녕하세요, ${sender}님.`,
        '',
        `전달해 주신 "${mail.subject}" 건 접수했습니다.`,
        '현재 유관 담당자와 세부 사항을 검토 중이며, 검토를 마치는 대로 정리된 회신을 드리겠습니다. 필요 시 짧은 온라인 미팅으로 조율 가능합니다.',
        '',
        '감사합니다.',
      ].join('\n') + sig,
    },
  ];
}

/**
 * Superhuman Ask AI / Shortwave AI Search:
 * Grounded natural-language Q&A across ALL loaded emails, attachments, projects, and calendar events.
 */
export function answerInboxQuestion(
  question: string,
  emails: MailItem[],
  projects: ProjectDefinition[] = [],
  calendarEvents: CalendarScheduleEvent[] = []
): AskInboxResponse {
  const q = String(question || '').trim();
  const safeEmails = Array.isArray(emails) ? emails : [];

  if (!q) {
    return {
      question: '',
      answer: '질문을 입력해 주세요. 예: "회신해야 할 긴급 메일 알려줘", "견적·금액이 포함된 메일 정리해줘", "첨부 도면 목록 찾아줘".',
      bulletFindings: [],
      citations: [],
      matchedCount: 0,
    };
  }

  if (safeEmails.length === 0) {
    return {
      question: q,
      answer: '현재 로드된 실제 메일이 없습니다. 먼저 .pst/.eml 파일을 열거나 Microsoft 365 계정을 동기화해 주세요.',
      bulletFindings: [],
      citations: [],
      matchedCount: 0,
    };
  }

  const qLower = q.toLowerCase();
  const isUrgentQuery = /(긴급|회신|답장|해야\s*할|할\s*일|urgent|action|reply|todo)/i.test(qLower);
  const isMoneyQuery = /(견적|금액|가격|단가|비용|발주|quote|price|usd|krw|budget|cost|영업)/i.test(qLower);
  const isAttachmentQuery = /(첨부|파일|pdf|엑셀|xlsx|도면|dwg|attachment)/i.test(qLower);
  const isScheduleQuery = /(일정|회의|미팅|마감|기한|스케줄|schedule|meeting|deadline)/i.test(qLower);
  const isWaitingQuery = /(대기|미회신|답변\s*안|waiting|follow\s*up)/i.test(qLower);

  // Extract content search tokens (excluding common stop words)
  const stopWords = new Set([
    '알려줘', '정리해줘', '찾아줘', '보여줘', '메일', '이메일', '관련', '전체', '이번', '오늘', '어제',
    'what', 'show', 'find', 'list', 'email', 'emails', 'about', 'from', 'with', 'the', 'and', 'for',
  ]);
  const tokens = qLower
    .replace(/[^\w가-힣-]/g, ' ')
    .split(/\s+/)
    .map((t) => t.trim())
    .filter((t) => t.length >= 2 && !stopWords.has(t));

  const scored = safeEmails
    .map((mail) => {
      const profile = classifyEmailIntent(mail);
      const hay = `${mail.subject}\n${mail.fromName}\n${mail.fromEmail}\n${mail.body}\n${(
        mail.attachments || []
      )
        .map((a) => a.filename)
        .join(' ')}`.toLowerCase();
      const projCode = detectProjectCode(`${mail.subject} ${mail.body}`, projects);

      let score = 0;
      for (const tok of tokens) {
        if (mail.subject.toLowerCase().includes(tok)) score += 6;
        if (mail.fromName.toLowerCase().includes(tok) || mail.fromEmail.toLowerCase().includes(tok)) score += 5;
        if (projCode && projCode.toLowerCase().includes(tok)) score += 7;
        if (hay.includes(tok)) score += 3;
      }

      if (isUrgentQuery && (profile.urgency === 'urgent' || profile.intent === 'action_required')) score += 6;
      if (isMoneyQuery && (profile.extractedAmounts.length > 0 || profile.intent === 'sales_quote')) score += 7;
      if (isAttachmentQuery && (mail.hasAttachment || (mail.attachments?.length ?? 0) > 0 || profile.extractedDrawingNos.length > 0)) {
        score += 6;
      }
      if (isScheduleQuery && (profile.extractedDeadlines.length > 0 || profile.intent === 'meeting_schedule')) score += 6;
      if (isWaitingQuery && mail.folder === 'sent') score += 6;

      return { mail, profile, score, projCode };
    })
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score || b.mail.timestamp.localeCompare(a.mail.timestamp));

  const top = scored.slice(0, 6);
  const citations: AskInboxCitation[] = top.map((item, idx) => ({
    index: idx + 1,
    emailId: item.mail.id,
    subject: item.mail.subject,
    fromName: item.mail.fromName,
    date: item.mail.date,
    snippet: item.profile.oneLineSummary,
  }));

  const bulletFindings: string[] = top.map((item, idx) => {
    const extras: string[] = [];
    if (item.projCode) extras.push(`프로젝트 ${item.projCode}`);
    if (item.profile.extractedAmounts[0]) extras.push(`금액 ${item.profile.extractedAmounts[0]}`);
    if (item.profile.extractedDeadlines[0]) extras.push(`기한 ${item.profile.extractedDeadlines[0]}`);
    if (item.profile.extractedDrawingNos[0]) extras.push(`도면 ${item.profile.extractedDrawingNos[0]}`);
    const extraText = extras.length > 0 ? ` — ${extras.join(' · ')}` : '';
    return `[${idx + 1}] ${item.mail.date} · ${item.mail.fromName} · ${item.mail.subject}${extraText}`;
  });

  if (isScheduleQuery && calendarEvents.length > 0) {
    const matchedEvents = calendarEvents.slice(0, 3);
    for (const ev of matchedEvents) {
      bulletFindings.push(`[일정] ${ev.date} ${ev.timeLabel} · ${ev.title}`);
    }
  }

  const answer =
    top.length === 0
      ? `"${q}"와 직접 일치하는 메일을 찾지 못했습니다. 발신자 이름, 프로젝트 코드, 도면번호 또는 핵심 키워드로 다시 질문해 보세요.`
      : `"${q}" 기준으로 관련 메일 ${scored.length}건을 찾았습니다. 상위 ${top.length}건의 핵심 근거와 조치 사항을 아래에 인용([1]~[${top.length}])으로 정리했습니다.`;

  return {
    question: q,
    answer,
    bulletFindings,
    citations,
    matchedCount: scored.length,
  };
}

/**
 * Superhuman / Shortwave "Write with AI":
 * Expands rough bullet points into a complete, well-structured business email.
 */
export function expandBulletsToDraft(
  rawBullets: string,
  options: {
    tone: 'formal' | 'concise' | 'english';
    recipientName?: string;
    subjectHint?: string;
    signature?: string;
  }
): { subject: string; body: string } {
  const lines = String(rawBullets || '')
    .split(/\r?\n/)
    .map((l) => l.replace(/^[-*•\d.)\s]+/, '').trim())
    .filter(Boolean);

  const recipient = options.recipientName?.trim() || (options.tone === 'english' ? 'Team' : '담당자');
  const sig = options.signature?.trim() ? `\n\n${options.signature.trim()}` : '';
  const firstLine = lines[0] || '업무 협조 요청';
  const subject =
    options.subjectHint?.trim() ||
    (options.tone === 'english'
      ? `Follow-up: ${firstLine.slice(0, 60)}`
      : `[업무 공유] ${firstLine.slice(0, 50)}`);

  if (lines.length === 0) {
    return {
      subject,
      body:
        options.tone === 'english'
          ? `Dear ${recipient},\n\nPlease enter key bullet points to generate a draft.\n\nBest regards,${sig}`
          : `안녕하세요, ${recipient}님.\n\n전달할 핵심 항목을 입력하시면 완성된 메일 초안을 생성합니다.\n\n감사합니다.${sig}`,
    };
  }

  if (options.tone === 'english') {
    const bulletBlock = lines.map((l, idx) => `${idx + 1}. ${l}`).join('\n');
    return {
      subject,
      body:
        [
          `Dear ${recipient},`,
          '',
          'I hope this message finds you well. I am writing to share an update and confirm the following items:',
          '',
          bulletBlock,
          '',
          'Please let us know if you have any questions or require additional clarification.',
          '',
          'Best regards,',
        ].join('\n') + sig,
    };
  }

  if (options.tone === 'concise') {
    const bulletBlock = lines.map((l) => `- ${l}`).join('\n');
    return {
      subject,
      body:
        [
          `안녕하세요, ${recipient}님.`,
          '',
          '요청 및 공유 사항을 간략히 정리해 드립니다.',
          bulletBlock,
          '',
          '확인 후 회신 부탁드립니다. 감사합니다.',
        ].join('\n') + sig,
    };
  }

  const bulletBlock = lines.map((l, idx) => `${idx + 1}. ${l}`).join('\n');
  return {
    subject,
    body:
      [
        `안녕하세요, ${recipient}님.`,
        '평소 원활한 업무 협조에 깊이 감사드립니다.',
        '',
        '관련 건의 진행 사항 및 요청 내용을 아래와 같이 정리하여 송부드립니다.',
        '',
        bulletBlock,
        '',
        '내용 검토 후 확인 의견 또는 회신 가능 일정을 공유해 주시면 후속 업무를 신속히 진행하겠습니다.',
        '',
        '감사합니다.',
      ].join('\n') + sig,
  };
}

/**
 * Microsoft Copilot for Outlook "Coaching by Copilot":
 * Pre-send quality, missing attachment, sensitive-data leak, and tone risk check.
 */
export function coachEmailDraft(draft: {
  to: string;
  subject: string;
  body: string;
  hasAttachment?: boolean;
}): DraftCoachingReport {
  const warnings: DraftCoachingWarning[] = [];
  const to = String(draft.to || '').trim();
  const subject = String(draft.subject || '').trim();
  const body = String(draft.body || '').trim();

  if (!to) {
    warnings.push({
      severity: 'high',
      code: 'missing_recipient',
      title: '받는 사람(수신자)이 비어 있습니다',
      suggestion: '발송 전 수신자 이메일 주소 또는 사내 담당자를 지정하세요.',
    });
  }

  if (!subject) {
    warnings.push({
      severity: 'high',
      code: 'missing_subject',
      title: '메일 제목이 비어 있습니다',
      suggestion: '프로젝트 코드나 핵심 요청이 드러나는 한 줄 제목을 입력하세요.',
    });
  }

  const mentionsAttachment = /(첨부|붙임|동봉|attached|attachment|enclosed|please find)/i.test(body);
  if (mentionsAttachment && !draft.hasAttachment) {
    warnings.push({
      severity: 'high',
      code: 'missing_attachment',
      title: '본문에 첨부 언급이 있으나 첨부파일이 감지되지 않았습니다',
      suggestion: 'Outlook/Bmail 발송 전 실제 첨부파일(PDF/도면/견적서)이 첨부되었는지 확인하세요.',
    });
  }

  const masked = maskSensitiveData(body);
  if (masked !== body) {
    warnings.push({
      severity: 'medium',
      code: 'sensitive_data',
      title: '본문에 연락처·API 키·금액 등 민감 정보 패턴이 포함되어 있습니다',
      suggestion: '외부 수신자가 포함된 경우 불필요한 전화번호·내부 키·계정 노출 여부를 재확인하세요.',
    });
  }

  if (/(왜\s*안|당장|책임\s*지|말이\s*안\s*됨|최악|ridiculous|unacceptable|immediately\s+or\s+else)/i.test(body)) {
    warnings.push({
      severity: 'medium',
      code: 'harsh_tone',
      title: '공격적이거나 단정적인 어조가 감지되었습니다',
      suggestion: '사실 기반의 정중한 요청 표현("일정 확인 부탁드립니다", "원인 파악 후 공유 요청")으로 완화하는 것을 권장합니다.',
    });
  }

  if (body.length > 0 && !/(부탁|회신|확인|공유|감사|please|confirm|regards|thank)/i.test(body)) {
    warnings.push({
      severity: 'info',
      code: 'missing_cta',
      title: '명확한 회신 요청(Call-to-Action) 또는 마감 안내가 보이지 않습니다',
      suggestion: '상대방이 언제까지 무엇을 회신해야 하는지 마지막 문장에 명시하면 응답률이 높아집니다.',
    });
  }

  const wordCount = body ? body.split(/\s+/).filter(Boolean).length : 0;
  if (wordCount > 350) {
    warnings.push({
      severity: 'info',
      code: 'too_long',
      title: '본문이 다소 깁니다 (350단어 초과)',
      suggestion: '상단 2줄에 핵심 결론(TL;DR)을 배치하고 세부 항목을 번호 목록으로 분리해 보세요.',
    });
  }

  const hasHigh = warnings.some((w) => w.severity === 'high');
  const estimatedReadSec = Math.max(5, Math.round((wordCount / 3.2) * 1));
  const toneSummary = warnings.some((w) => w.code === 'harsh_tone')
    ? '다소 강경함 (완화 권장)'
    : /(감사|부탁드립니다|best regards|sincerely)/i.test(body)
      ? '정중·비즈니스 적합'
      : '중립·간결';

  return {
    readyToSend: !hasHigh && body.length > 0,
    warnings,
    wordCount,
    estimatedReadSec,
    toneSummary,
  };
}
