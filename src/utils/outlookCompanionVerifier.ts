import { extractShipbuildingDrawingRefs } from './outlookWorkspaceModel';

export type OutlookCompanionMode = 'composing' | 'reading' | 'idle';
export type OutlookVerificationSeverity = 'critical' | 'warning' | 'info';

export interface OutlookVerificationFinding {
  code:
    | 'missing_attachment'
    | 'missing_subject'
    | 'missing_recipient'
    | 'external_sensitive_leak'
    | 'drawing_and_deadline_check';
  severity: OutlookVerificationSeverity;
  title: string;
  detail: string;
}

export interface OutlookVerificationReport {
  status: 'safe' | 'caution' | 'block';
  score: number;
  summary: string;
  findings: OutlookVerificationFinding[];
  drawings: string[];
  deadlines: string[];
}

export interface OutlookActiveItemSnapshot {
  entryId: string;
  sourceType: 'inspector' | 'explorer' | 'local-fallback';
  subject: string;
  fromName: string;
  fromEmail: string;
  to: string[];
  cc: string[];
  bcc: string[];
  body: string;
  attachments: string[];
  attachmentsCount: number;
  receivedAt: string;
  sent: boolean;
  importance: 'low' | 'normal' | 'high';
}

export interface OutlookCompanionPollResult {
  outlookRunning: boolean;
  mode: OutlookCompanionMode;
  item: OutlookActiveItemSnapshot | null;
  verification: OutlookVerificationReport | null;
  polledAt: string;
  error?: string;
}

const ATTACHMENT_MENTION_RE =
  /(첨부\s*(?:드리|해|하|합|바랍|참조|확인|파일|문서|자료|도면)?|별첨|붙임|동봉|attached|enclosed|see\s+attach|도면\s*(?:송부|전달|보내|발송)|견적서\s*(?:송부|전달|보내|발송)|파일\s*(?:송부|전달|보내))/i;

const QUOTE_BOUNDARY_RE =
  /\r?\n(?:-----Original Message-----|________________________________|From:\s|보낸\s*사람:\s|수신\s*일시:|---\s*원본\s*메시지\s*---)/i;

const FREE_PUBLIC_EMAIL_DOMAIN_RE =
  /^(?:gmail\.com|naver\.com|daum\.net|hanmail\.net|nate\.com|hotmail\.com|outlook\.com|yahoo\.com|kakao\.com)$/i;

const DEADLINE_RE = /(?:\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b|\b\d{1,2}월\s*\d{1,2}일)/g;

const SENSITIVE_CATEGORIES: Array<{ label: string; pattern: RegExp }> = [
  { label: 'API 키/토큰', pattern: /(gsk_[A-Za-z0-9]{10,}|sk-ant-[A-Za-z0-9\-_]{10,}|AIza[0-9A-Za-z\-_]{15,})/ },
  { label: '비밀번호/계정', pattern: /(?:비밀번호|패스워드|password|passcode)\s*[:=]\s*\S+/i },
  { label: '전화번호', pattern: /\b01[016789]-?\d{3,4}-?\d{4}\b/ },
  { label: '계좌번호', pattern: /(?:계좌번호|입금계좌|account\s*no)[^\n]{0,20}\d{4,}|\b\d{3,6}-\d{2,6}-\d{4,8}\b/i },
  {
    label: '견적·계약 금액',
    pattern: /(?:(?:\$|USD|EUR|KRW|₩)\s*\d[\d,]*(?:\.\d+)?|\b\d[\d,]+\s*(?:원|만원|억원))/i,
  },
];

function extractTopMessageSegment(body: string): string {
  const clean = String(body || '');
  const m = clean.match(QUOTE_BOUNDARY_RE);
  if (m && typeof m.index === 'number' && m.index > 20) {
    return clean.slice(0, m.index);
  }
  return clean;
}

function splitRecipientTokens(list: unknown): string[] {
  if (!Array.isArray(list)) return [];
  const out: string[] = [];
  for (const raw of list) {
    const s = String(raw || '').trim();
    if (s) out.push(s);
  }
  return out;
}

function extractDomain(address: string): string {
  const m = String(address || '').match(/@([A-Za-z0-9.-]+\.[A-Za-z]{2,})/);
  return m ? m[1].toLowerCase() : '';
}

/**
 * Pure 0ms pre-send & incoming safety verification for active Outlook drafts/mails.
 * Browser-safe (zero Node built-ins) and zero side effects.
 */
export function verifyOutlookItemSafety(
  item: Partial<OutlookActiveItemSnapshot> & { attachmentNames?: string[] },
  mode: OutlookCompanionMode = 'reading'
): OutlookVerificationReport {
  const subject = String(item.subject || '').trim();
  const fromEmail = String(item.fromEmail || '').trim();
  const to = splitRecipientTokens(item.to);
  const cc = splitRecipientTokens(item.cc);
  const bcc = splitRecipientTokens(item.bcc);
  const fullBody = String(item.body || '');
  const topBody = extractTopMessageSegment(fullBody);
  const attNames = Array.isArray(item.attachments)
    ? item.attachments
    : Array.isArray(item.attachmentNames)
      ? item.attachmentNames
      : [];
  const attachmentsCount =
    typeof item.attachmentsCount === 'number' && Number.isFinite(item.attachmentsCount)
      ? Math.max(0, item.attachmentsCount)
      : attNames.length;

  const findings: OutlookVerificationFinding[] = [];
  const textToCheck = `${subject}\n${topBody}`;

  if (ATTACHMENT_MENTION_RE.test(textToCheck) && attachmentsCount === 0) {
    findings.push({
      code: 'missing_attachment',
      severity: mode === 'composing' ? 'critical' : 'info',
      title: mode === 'composing' ? '첨부파일 누락 위험' : '본문 내 첨부 언급 (파일 0건)',
      detail:
        mode === 'composing'
          ? '본문/제목에 "첨부·별첨·도면 송부" 표현이 있으나 실제 첨부된 파일이 0건입니다. 발송 전 파일을 첨부했는지 확인하세요.'
          : '수신 메일에 첨부 관련 언급이 있으나 첨부파일이 감지되지 않았습니다 (대용량 링크 또는 누락 가능).',
    });
  }

  const allRecipients = [...to, ...cc, ...bcc];
  if (mode === 'composing' && allRecipients.length === 0) {
    findings.push({
      code: 'missing_recipient',
      severity: 'critical',
      title: '수신자 미지정',
      detail: '받는 사람(To)·참조(Cc)·숨은 참조(Bcc)가 모두 비어 있습니다.',
    });
  }

  if (mode === 'composing' && (!subject || subject === '(제목 없음)')) {
    findings.push({
      code: 'missing_subject',
      severity: 'warning',
      title: '메일 제목 누락',
      detail: '제목이 비어 있으면 스팸 분류 또는 수신자 혼동 우려가 있습니다.',
    });
  }

  const senderDomain = extractDomain(fromEmail);
  const externalDomains = new Set<string>();
  for (const addr of allRecipients) {
    const dom = extractDomain(addr);
    if (!dom) continue;
    if (FREE_PUBLIC_EMAIL_DOMAIN_RE.test(dom) || (senderDomain && dom !== senderDomain)) {
      externalDomains.add(dom);
    }
  }
  const matchedSensitive = SENSITIVE_CATEGORIES.filter((c) => c.pattern.test(textToCheck)).map(
    (c) => c.label
  );
  if (externalDomains.size > 0 && matchedSensitive.length > 0) {
    findings.push({
      code: 'external_sensitive_leak',
      severity: 'warning',
      title: '외부 수신자 민감 정보 포함 점검',
      detail: `외부 도메인(${Array.from(externalDomains).slice(0, 3).join(', ')}) 수신자와 함께 [${matchedSensitive.join(' · ')}] 패턴이 감지되었습니다. 발송 대상이 맞는지 확인하세요.`,
    });
  }

  const drawings = extractShipbuildingDrawingRefs(textToCheck)
    .map((ref) => ref.raw || ref.canonKey)
    .filter(Boolean)
    .slice(0, 12);
  const rawDeadlines = textToCheck.match(DEADLINE_RE) || [];
  const deadlines = Array.from(new Set(rawDeadlines.map((d) => d.trim()))).slice(0, 8);
  if (drawings.length > 0 || deadlines.length > 0) {
    const parts: string[] = [];
    if (drawings.length > 0) parts.push(`도면 ${drawings.length}건(${drawings.slice(0, 4).join(', ')})`);
    if (deadlines.length > 0) parts.push(`일정/기한(${deadlines.slice(0, 4).join(', ')})`);
    findings.push({
      code: 'drawing_and_deadline_check',
      severity: 'info',
      title: '도면번호 · 일정 자동 추출 검증',
      detail: `${parts.join(' / ')} — 리비전·호기·납기일이 맞는지 대조 확인하세요.`,
    });
  }

  let score = 100;
  let hasCritical = false;
  let hasWarning = false;
  for (const f of findings) {
    if (f.severity === 'critical') {
      score -= 45;
      hasCritical = true;
    } else if (f.severity === 'warning') {
      score -= 20;
      hasWarning = true;
    }
  }
  score = Math.max(0, Math.min(100, score));

  const status: OutlookVerificationReport['status'] = hasCritical
    ? 'block'
    : hasWarning
      ? 'caution'
      : 'safe';

  const criticalCount = findings.filter((f) => f.severity === 'critical').length;
  const warningCount = findings.filter((f) => f.severity === 'warning').length;
  const summary =
    status === 'block'
      ? `발송 전 필수 확인 ${criticalCount}건 (보완 필요)`
      : status === 'caution'
        ? `주의 검토 ${warningCount}건 (외부·제목·민감정보 확인 권장)`
        : '발송·수신 검증 통과 (위험 요인 없음)';

  return {
    status,
    score,
    summary,
    findings,
    drawings,
    deadlines,
  };
}
