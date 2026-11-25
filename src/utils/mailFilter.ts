import type {
  AnalysisOptions,
  AnalysisResult,
  MailItem,
  MailScope,
} from '../types/mail';

export interface DateSummary {
  date: string;
  total: number;
  received: number;
  sent: number;
}

export interface FilterValidationResult {
  valid: boolean;
  error: string | null;
  matchedEmails: MailItem[];
  receivedCount: number;
  sentCount: number;
  periodLabel: string;
  scopeLabel: string;
}

export const SCOPE_LABELS: Record<MailScope, string> = {
  all: '받은 메일 + 보낸 메일',
  received: '받은 메일만',
  sent: '보낸 메일만',
  selected: '선택한 메일',
};

export function isValidIsoDate(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(value) && !Number.isNaN(Date.parse(`${value}T00:00:00Z`));
}

export type SearchDateRangePreset = 'any' | 'today' | '7d' | '30d' | 'custom';

export interface GlobalSearchFilters {
  from: string;
  to: string;
  hasAttachments: boolean;
  datePreset: SearchDateRangePreset;
  dateFrom: string;
  dateTo: string;
}

export const DEFAULT_GLOBAL_SEARCH_FILTERS: GlobalSearchFilters = {
  from: '',
  to: '',
  hasAttachments: false,
  datePreset: 'any',
  dateFrom: '',
  dateTo: '',
};

export interface SearchFilterValidation {
  valid: boolean;
  error: string | null;
}

function shiftIsoDate(isoDate: string, deltaDays: number): string {
  const [y, m, d] = isoDate.split('-').map(Number);
  const utc = Date.UTC(y, m - 1, d + deltaDays);
  return new Date(utc).toISOString().slice(0, 10);
}

/**
 * Resolves inclusive YYYY-MM-DD bounds for a date range preset relative to a reference ISO date.
 */
export function resolveSearchDateRange(
  preset: SearchDateRangePreset,
  referenceIsoDate: string
): { dateFrom: string; dateTo: string } {
  if (preset === 'any' || !isValidIsoDate(referenceIsoDate)) {
    return { dateFrom: '', dateTo: '' };
  }
  if (preset === 'today') {
    return { dateFrom: referenceIsoDate, dateTo: referenceIsoDate };
  }
  if (preset === '7d') {
    return { dateFrom: shiftIsoDate(referenceIsoDate, -6), dateTo: referenceIsoDate };
  }
  if (preset === '30d') {
    return { dateFrom: shiftIsoDate(referenceIsoDate, -29), dateTo: referenceIsoDate };
  }
  return { dateFrom: '', dateTo: '' };
}

export function validateGlobalSearchFilters(
  filters: GlobalSearchFilters
): SearchFilterValidation {
  const fromDate = filters.dateFrom.trim();
  const toDate = filters.dateTo.trim();

  if (fromDate && !isValidIsoDate(fromDate)) {
    return { valid: false, error: '시작 날짜 형식이 올바르지 않습니다 (YYYY-MM-DD).' };
  }
  if (toDate && !isValidIsoDate(toDate)) {
    return { valid: false, error: '종료 날짜 형식이 올바르지 않습니다 (YYYY-MM-DD).' };
  }
  if (fromDate && toDate && fromDate > toDate) {
    return { valid: false, error: '시작 날짜가 종료 날짜보다 늦을 수 없습니다.' };
  }
  return { valid: true, error: null };
}

export function countActiveGlobalSearchFilters(filters: GlobalSearchFilters): number {
  let count = 0;
  if (filters.from.trim()) count += 1;
  if (filters.to.trim()) count += 1;
  if (filters.hasAttachments) count += 1;
  if (filters.dateFrom.trim() || filters.dateTo.trim()) count += 1;
  return count;
}

/**
 * Pure predicate scoping a MailItem by global search text plus From, To, Has Attachments, and Date Range filters.
 */
export function matchesGlobalSearchFilters(
  mail: Pick<
    MailItem,
    'subject' | 'fromName' | 'fromEmail' | 'to' | 'cc' | 'body' | 'date' | 'hasAttachment' | 'attachments' | 'category'
  >,
  query: string,
  filters: GlobalSearchFilters
): boolean {
  const fromTerm = filters.from.trim().toLowerCase();
  if (fromTerm) {
    const senderHaystack = `${mail.fromName || ''} ${mail.fromEmail || ''}`.toLowerCase();
    if (!senderHaystack.includes(fromTerm)) return false;
  }

  const toTerm = filters.to.trim().toLowerCase();
  if (toTerm) {
    const recipients = [...(Array.isArray(mail.to) ? mail.to : []), ...(Array.isArray(mail.cc) ? mail.cc : [])];
    const matchedRecipient = recipients.some((addr) => addr.toLowerCase().includes(toTerm));
    if (!matchedRecipient) return false;
  }

  if (filters.hasAttachments) {
    const hasFiles = Boolean(mail.hasAttachment) || (Array.isArray(mail.attachments) && mail.attachments.length > 0);
    if (!hasFiles) return false;
  }

  const dateFrom = filters.dateFrom.trim();
  const dateTo = filters.dateTo.trim();
  if (dateFrom || dateTo) {
    const validation = validateGlobalSearchFilters(filters);
    if (!validation.valid) return false;
    if (!isValidIsoDate(mail.date)) return false;
    if (dateFrom && mail.date < dateFrom) return false;
    if (dateTo && mail.date > dateTo) return false;
  }

  const q = query.trim().toLowerCase();
  if (!q) return true;

  const recipientText = [...(Array.isArray(mail.to) ? mail.to : []), ...(Array.isArray(mail.cc) ? mail.cc : [])]
    .join(' ')
    .toLowerCase();

  return (
    (mail.subject || '').toLowerCase().includes(q) ||
    (mail.fromName || '').toLowerCase().includes(q) ||
    (mail.fromEmail || '').toLowerCase().includes(q) ||
    recipientText.includes(q) ||
    (mail.body || '').toLowerCase().includes(q) ||
    (mail.category || '').toLowerCase().includes(q)
  );
}

export function filterEmails(emails: MailItem[], options: Pick<AnalysisOptions, 'dateMode' | 'singleDate' | 'startDate' | 'endDate' | 'scope' | 'selectedIds'>): MailItem[] {
  if (options.scope === 'selected') {
    const selectedSet = new Set(options.selectedIds);
    return emails
      .filter((mail) => selectedSet.has(mail.id))
      .sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  }

  return emails
    .filter((mail) => {
      if (options.scope === 'received' && mail.folder !== 'inbox') {
        return false;
      }
      if (options.scope === 'sent' && mail.folder !== 'sent') {
        return false;
      }
      if (options.dateMode === 'single') {
        return mail.date === options.singleDate;
      }
      return mail.date >= options.startDate && mail.date <= options.endDate;
    })
    .sort((a, b) => b.timestamp.localeCompare(a.timestamp));
}

export function validateAnalysisFilter(
  emails: MailItem[],
  options: Pick<AnalysisOptions, 'dateMode' | 'singleDate' | 'startDate' | 'endDate' | 'scope' | 'selectedIds'>
): FilterValidationResult {
  const scopeLabel = SCOPE_LABELS[options.scope];

  if (options.scope !== 'selected') {
    if (options.dateMode === 'single') {
      if (!isValidIsoDate(options.singleDate)) {
        return {
          valid: false,
          error: '분석할 일자를 올바르게 선택해 주세요.',
          matchedEmails: [],
          receivedCount: 0,
          sentCount: 0,
          periodLabel: '일자 미지정',
          scopeLabel,
        };
      }
    } else {
      if (!isValidIsoDate(options.startDate) || !isValidIsoDate(options.endDate)) {
        return {
          valid: false,
          error: '시작 일자와 종료 일자를 모두 입력해 주세요.',
          matchedEmails: [],
          receivedCount: 0,
          sentCount: 0,
          periodLabel: '기간 미지정',
          scopeLabel,
        };
      }
      if (options.startDate > options.endDate) {
        return {
          valid: false,
          error: '시작 일자가 종료 일자보다 늦을 수 없습니다.',
          matchedEmails: [],
          receivedCount: 0,
          sentCount: 0,
          periodLabel: `${options.startDate} ~ ${options.endDate}`,
          scopeLabel,
        };
      }
    }
  }

  const matchedEmails = filterEmails(emails, options);
  const receivedCount = matchedEmails.filter((m) => m.folder === 'inbox').length;
  const sentCount = matchedEmails.filter((m) => m.folder === 'sent').length;

  let periodLabel: string;
  if (options.scope === 'selected') {
    if (matchedEmails.length === 0) {
      periodLabel = '선택된 메일 없음';
    } else {
      const dates = matchedEmails.map((m) => m.date).sort();
      const first = dates[0];
      const last = dates[dates.length - 1];
      periodLabel = first === last ? first : `${first} ~ ${last}`;
    }
  } else {
    periodLabel =
      options.dateMode === 'single'
        ? options.singleDate
        : `${options.startDate} ~ ${options.endDate}`;
  }

  if (matchedEmails.length === 0) {
    const emptyReason =
      options.scope === 'selected'
        ? '선택한 메일이 없습니다. 메일 목록에서 분석할 메일을 1개 이상 체크해 주세요.'
        : '선택한 일자 및 메일 조건에 해당하는 메일이 없습니다.';
    return {
      valid: false,
      error: emptyReason,
      matchedEmails: [],
      receivedCount: 0,
      sentCount: 0,
      periodLabel,
      scopeLabel,
    };
  }

  return {
    valid: true,
    error: null,
    matchedEmails,
    receivedCount,
    sentCount,
    periodLabel,
    scopeLabel,
  };
}

export function getAvailableDates(emails: MailItem[]): DateSummary[] {
  const map = new Map<string, DateSummary>();
  for (const mail of emails) {
    const existing = map.get(mail.date) ?? {
      date: mail.date,
      total: 0,
      received: 0,
      sent: 0,
    };
    existing.total += 1;
    if (mail.folder === 'inbox') {
      existing.received += 1;
    } else {
      existing.sent += 1;
    }
    map.set(mail.date, existing);
  }
  return Array.from(map.values()).sort((a, b) => b.date.localeCompare(a.date));
}

const PROMO_PATTERNS = [/수신\s*거부/i, /unsubscribe/i, /\(광고\)/, /\[광고\]/, /noreply@/i];

export function isPromotionalEmail(mail: Pick<MailItem, 'subject' | 'fromEmail' | 'body'>): boolean {
  const sample = `${mail.subject} ${mail.fromEmail} ${mail.body.slice(0, 400)}`;
  return PROMO_PATTERNS.some((re) => re.test(sample));
}

export function cleanEmailBody(raw: string, maxChars = 2400): string {
  const cleaned = raw
    .replace(/\r\n/g, '\n')
    .replace(/\n(--\s*|___+|감사합니다\.?\s*\n.*드림|CONFIDENTIALITY NOTICE)[\s\S]*$/i, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  return cleaned.length > maxChars ? `${cleaned.slice(0, maxChars)}\n...[이하 생략]` : cleaned;
}

export function maskSensitiveData(input: string): string {
  return input
    // Mask email addresses
    .replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, '[이메일 비공개]')
    // Mask Korean mobile / office phone numbers
    .replace(/\b0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4}\b/g, '[연락처 비공개]')
    // Mask IPv4 addresses
    .replace(/\b(?:\d{1,3}\.){3}\d{1,3}\b/g, '[IP 비공개]')
    // Mask common API keys / tokens
    .replace(/\b(sk-[a-zA-Z0-9_-]{16,}|gsk_[a-zA-Z0-9_-]{16,}|ghp_[a-zA-Z0-9]{20,})\b/g, '[API키 비공개]')
    // Mask monetary amounts in KRW (e.g., 4,800만원, 120,000원, 3억원)
    .replace(/\b\d{1,3}(?:,\d{3})+\s?(?:만원|억원|원)|\b\d+\s?(?:만원|억원)/g, '[금액 비공개]');
}

/**
 * Splits serialized <email_record> items into token-safe chunks so Groq 8K TPM
 * single-request HTTP 413 limits are avoided via Map-Reduce.
 */
export function chunkRecordsByBudget(records: string[], maxCharsPerChunk = 6000): string[] {
  const chunks: string[] = [];
  let current = '';
  for (const rec of records) {
    if (current.length + rec.length > maxCharsPerChunk && current.length > 0) {
      chunks.push(current);
      current = rec;
    } else {
      current += (current ? '\n\n' : '') + rec;
    }
  }
  if (current) chunks.push(current);
  return chunks;
}

export function formatWorkLogMarkdown(result: AnalysisResult): string {
  const { workLog } = result;
  const lines: string[] = [
    `# ${workLog.title}`,
    '',
    `- **분석 기간**: ${workLog.periodLabel}`,
    `- **분석 범위**: ${result.scopeLabel} (총 ${result.analyzedCount}건 · 수신 ${result.receivedCount} / 발신 ${result.sentCount})`,
    `- **분석 엔진**: ${result.providerModel}`,
    '',
    '## 핵심 요약',
    workLog.executiveSummary,
    '',
    '## 완료한 주요 업무',
    ...(workLog.completedTasks.length > 0
      ? workLog.completedTasks.map((item) => `- [x] ${item}`)
      : ['- 해당 사항 없음']),
    '',
    '## 진행 중인 업무',
    ...(workLog.inProgressTasks.length > 0
      ? workLog.inProgressTasks.map((item) => `- [ ] ${item}`)
      : ['- 해당 사항 없음']),
    '',
    '## 주요 수신·발신 커뮤니케이션',
    ...(workLog.keyCommunications.length > 0
      ? workLog.keyCommunications.map((item) => `- ${item}`)
      : ['- 해당 사항 없음']),
    '',
    '## 이슈 및 리스크',
    ...(workLog.issuesAndRisks.length > 0
      ? workLog.issuesAndRisks.map((item) => `- ${item}`)
      : ['- 특이 리스크 없음']),
    '',
    '## 익일 / 후속 계획',
    ...(workLog.nextStepPlan.length > 0
      ? workLog.nextStepPlan.map((item) => `- ${item}`)
      : ['- 후속 일정 정리 예정']),
    ...(result.projectSummaries && result.projectSummaries.length > 0
      ? [
          '',
          '## 프로젝트별 전체 정리 (내가 한 일 · 결정 · 첨부)',
          ...result.projectSummaries.flatMap((prj) => [
            `### [${prj.projectCode}] ${prj.projectName} (${prj.status})`,
            ...(prj.myActions.length > 0
              ? prj.myActions.map((a) => `- **내가 한 일**: ${a}`)
              : []),
            ...(prj.keyDecisions.length > 0
              ? prj.keyDecisions.map((d) => `- **주요 협의/결정**: ${d}`)
              : []),
            ...(prj.attachmentsMentioned.length > 0
              ? [`- **관련 첨부/문서**: ${prj.attachmentsMentioned.join(', ')}`]
              : []),
            `- **다음 조치**: ${prj.nextAction}`,
            '',
          ]),
        ]
      : []),
  ];
  return lines.join('\n');
}

export function formatBlogPostMarkdown(result: AnalysisResult): string {
  const { blogPost } = result;
  const lines: string[] = [
    `# ${blogPost.title}`,
    '',
    `> ${blogPost.subtitle}`,
    '',
    `예상 읽기 시간: ${blogPost.readingTimeMinutes}분 · 태그: ${blogPost.tags.map((t) => `#${t.replace(/^#/, '')}`).join(' ')}`,
    '',
    blogPost.intro,
    '',
    ...blogPost.sections.flatMap((section) => [
      `## ${section.heading}`,
      '',
      section.body,
      '',
    ]),
    '## 마치며',
    '',
    blogPost.conclusion,
  ];
  return lines.join('\n');
}
