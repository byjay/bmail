import type {
  AnalysisResult,
  AttachmentKind,
  MailItem,
  MeetingMinutesResult,
  ParsedDocumentSummary,
} from '../types/mail';

export type ProjectFolderPreset =
  | 'project-month-sender'
  | 'project-date-kind'
  | 'domain-project-date'
  | 'date-project-sender';

export interface ProjectDefinition {
  id: string;
  code: string;
  name: string;
  keywords: string[];
  color: string;
  isFavorite: boolean;
  createdAt: string;
  /** Optional per-project exact sender/participant email addresses */
  senderEmails?: string[];
  /** Optional per-project domain rules (normalized without leading @) */
  domains?: string[];
  /** Optional per-project bound mailbox/account addresses (for multi-account routing) */
  accounts?: string[];
  /** Optional preferred received-file folder layout template */
  folderTemplate?: ProjectFolderPreset;
  /** Optional default subfolder name inside the project root */
  defaultSubfolder?: string;
}

export interface CalendarScheduleEvent {
  id: string;
  title: string;
  date: string; // YYYY-MM-DD
  timeLabel: string;
  category: 'meeting' | 'deadline' | 'email-extracted' | 'manual' | 'action-item';
  projectCode?: string;
  sourceEmailId?: string;
  sourceEmailSubject?: string;
  notes?: string;
  completed?: boolean;
}

export interface ParsedNaturalQuery {
  rawQuery: string;
  startDate?: string;
  endDate?: string;
  projectCode?: string;
  attachmentKind?: AttachmentKind;
  folderHint?: 'inbox' | 'sent';
  keywords: string[];
}

export interface ConversationalSearchHit {
  type: 'email' | 'document';
  id: string;
  title: string;
  date: string;
  subtitle: string;
  matchedReasons: string[];
  snippet: string;
  projectCode?: string;
}

/**
 * Normalizes email subjects (RE:, FW:, Fwd:, [회신], [전달]) to group Outlook conversation threads.
 */
export function normalizeConversationSubject(subject: string): string {
  return (subject || '')
    .replace(/^((\s*(re|fw|fwd|답장|회신|전달)\s*:\s*)|(\s*\[(회신|전달|답장|외부)\]\s*))+/gi, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * Deterministically extracts a project code from email subject/body/category or matches user-defined project rules.
 * Works 100% offline with ZERO LLM calls.
 */
export function detectProjectCode(
  text: string,
  projects: ProjectDefinition[]
): string | null {
  const lower = text.toLowerCase();

  // 1. Check user-registered projects first (exact code or keywords)
  for (const project of projects) {
    if (project.code && lower.includes(project.code.toLowerCase())) {
      return project.code;
    }
    if (project.name && lower.includes(project.name.toLowerCase())) {
      return project.code;
    }
    for (const kw of project.keywords) {
      if (kw.trim() && lower.includes(kw.trim().toLowerCase())) {
        return project.code;
      }
    }
  }

  // 2. Fallback regex for bracketed project codes like [PRJ-2026], [PJT-01], [DEV-104], #PRJ-01
  const codeMatch = text.match(/(?:\[|#)([A-Z]{2,6}[-_]\d{2,6}(?:[-_][A-Z0-9]{1,6})?)\]?/i);
  if (codeMatch) {
    return codeMatch[1].toUpperCase();
  }

  return null;
}

/**
 * Deterministically parses Korean/ISO natural-language queries like:
 * "11월 1일부터 11월 19일까지 PRJ-01 PDF 계약서 찾아줘"
 * "2026-11-10 ~ 2026-11-19 보낸 엑셀 견적"
 */
export function parseNaturalSearchQuery(
  rawQuery: string,
  referenceYear = new Date().getFullYear()
): ParsedNaturalQuery {
  const q = rawQuery.trim();
  let startDate: string | undefined;
  let endDate: string | undefined;

  // Match ISO range: 2026-11-01 ~ 2026-11-19 or 2026-11-01부터 2026-11-19
  const isoDates = q.match(/\b(20\d{2})[-./](\d{1,2})[-./](\d{1,2})\b/g);
  if (isoDates && isoDates.length >= 2) {
    const norm = (s: string) => {
      const [y, m, d] = s.split(/[-./]/);
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    };
    startDate = norm(isoDates[0]);
    endDate = norm(isoDates[1]);
  } else if (isoDates && isoDates.length === 1) {
    const [y, m, d] = isoDates[0].split(/[-./]/);
    startDate = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    endDate = startDate;
  } else {
    // Match Korean dates: 11월 1일 ~ 11월 19일
    const korMatches = Array.from(q.matchAll(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/g));
    if (korMatches.length >= 2) {
      startDate = `${referenceYear}-${korMatches[0][1].padStart(2, '0')}-${korMatches[0][2].padStart(2, '0')}`;
      endDate = `${referenceYear}-${korMatches[1][1].padStart(2, '0')}-${korMatches[1][2].padStart(2, '0')}`;
    } else if (korMatches.length === 1) {
      startDate = `${referenceYear}-${korMatches[0][1].padStart(2, '0')}-${korMatches[0][2].padStart(2, '0')}`;
      endDate = startDate;
    }
  }

  if (startDate && endDate && startDate > endDate) {
    const tmp = startDate;
    startDate = endDate;
    endDate = tmp;
  }

  // Project code detection
  const prjMatch = q.match(/\b([A-Z]{2,6}[-_]\d{2,6}(?:[-_][A-Z0-9]{1,6})?)\b/i);
  const projectCode = prjMatch ? prjMatch[1].toUpperCase() : undefined;

  // Attachment kind detection
  let attachmentKind: AttachmentKind | undefined;
  if (/\bpdf\b|피디에프/i.test(q)) attachmentKind = 'pdf';
  else if (/\b(xlsx?|csv|excel)\b|엑셀|스프레드시트/i.test(q)) attachmentKind = 'excel';
  else if (/\b(docx?|word)\b|워드|문서/i.test(q)) attachmentKind = 'word';
  else if (/\b(mp3|wav|m4a|audio)\b|음성|녹음|녹취/i.test(q)) attachmentKind = 'audio';

  // Folder hint
  let folderHint: 'inbox' | 'sent' | undefined;
  if (/(?:보낸(?:\s*(?:메일|엑셀|파일|문서|pdf|견적))?|발신|내가\s*보낸)/i.test(q)) folderHint = 'sent';
  else if (/(?:받은(?:\s*(?:메일|엑셀|파일|문서|pdf|견적))?|수신|들어온)/i.test(q)) folderHint = 'inbox';

  // Clean remaining search keywords
  const stripped = q
    .replace(/\b20\d{2}[-./]\d{1,2}[-./]\d{1,2}\b/g, ' ')
    .replace(/\d{1,2}\s*월\s*\d{1,2}\s*일/g, ' ')
    .replace(/(부터|까지|에서|중에|찾아줘|찾아|검색해줘|검색|보여줘|검증해줘|검증|분석해줘|분석|관련|메일|파일|첨부|보낸|받은|수신|발신|프로젝트|pdf|excel|xlsx|word|docx|엑셀|워드|음성|녹음|~|-)/gi, ' ')
    .replace(projectCode ? new RegExp(projectCode, 'gi') : /$^/, ' ')
    .trim();

  const keywords = stripped
    .split(/\s+/)
    .map((w) => w.trim())
    .filter((w) => w.length >= 2);

  return {
    rawQuery: q,
    startDate,
    endDate,
    projectCode,
    attachmentKind,
    folderHint,
    keywords,
  };
}

/**
 * Instant deterministic conversational search & verification across emails and parsed documents.
 */
export function runDeterministicConversationalSearch(
  rawQuery: string,
  emails: MailItem[],
  documents: ParsedDocumentSummary[],
  projects: ProjectDefinition[]
): { parsed: ParsedNaturalQuery; hits: ConversationalSearchHit[] } {
  const parsed = parseNaturalSearchQuery(rawQuery);
  const hits: ConversationalSearchHit[] = [];

  for (const mail of emails) {
    const reasons: string[] = [];
    if (parsed.startDate && parsed.endDate) {
      if (mail.date < parsed.startDate || mail.date > parsed.endDate) continue;
      reasons.push(`기간 일치 (${mail.date})`);
    }
    if (parsed.folderHint && mail.folder !== parsed.folderHint) continue;
    if (parsed.folderHint) reasons.push(parsed.folderHint === 'sent' ? '보낸 편지함' : '받은 편지함');

    const fullText = `${mail.subject} ${mail.fromName} ${mail.fromEmail} ${mail.body} ${(
      mail.attachments ?? []
    )
      .map((a) => a.filename)
      .join(' ')}`;
    const detectedProject = detectProjectCode(fullText, projects);

    if (parsed.projectCode) {
      const matched =
        detectedProject?.toLowerCase() === parsed.projectCode.toLowerCase() ||
        fullText.toLowerCase().includes(parsed.projectCode.toLowerCase());
      if (!matched) continue;
      reasons.push(`프로젝트 ${parsed.projectCode}`);
    }

    if (parsed.attachmentKind) {
      const hasKind =
        (mail.attachments ?? []).some((a) => a.kind === parsed.attachmentKind) ||
        fullText.toLowerCase().includes(parsed.attachmentKind);
      if (!hasKind) continue;
      reasons.push(`${parsed.attachmentKind.toUpperCase()} 첨부/언급`);
    }

    if (parsed.keywords.length > 0) {
      const lowerText = fullText.toLowerCase();
      const matchedKws = parsed.keywords.filter((kw) => lowerText.includes(kw.toLowerCase()));
      if (matchedKws.length === 0) continue;
      reasons.push(`키워드: ${matchedKws.join(', ')}`);
    }

    hits.push({
      type: 'email',
      id: mail.id,
      title: mail.subject,
      date: mail.date,
      subtitle: `${mail.fromName} (${mail.folder === 'sent' ? '발신' : '수신'} · ${mail.timeLabel})`,
      matchedReasons: reasons.length > 0 ? reasons : ['전체 조건 일치'],
      snippet: mail.snippet || mail.body.slice(0, 140),
      projectCode: detectedProject ?? undefined,
    });
  }

  for (const doc of documents) {
    const reasons: string[] = [];
    if (parsed.attachmentKind && doc.kind !== parsed.attachmentKind) continue;
    if (parsed.attachmentKind) reasons.push(`${doc.kind.toUpperCase()} 문서`);

    const fullText = `${doc.filename} ${doc.extractedText}`;
    const detectedProject = detectProjectCode(fullText, projects);
    if (parsed.projectCode && !fullText.toLowerCase().includes(parsed.projectCode.toLowerCase())) {
      continue;
    }
    if (parsed.projectCode) reasons.push(`프로젝트 ${parsed.projectCode}`);

    if (parsed.keywords.length > 0) {
      const lowerText = fullText.toLowerCase();
      const matchedKws = parsed.keywords.filter((kw) => lowerText.includes(kw.toLowerCase()));
      if (matchedKws.length === 0) continue;
      reasons.push(`키워드: ${matchedKws.join(', ')}`);
    }

    hits.push({
      type: 'document',
      id: doc.id,
      title: doc.filename,
      date: new Date().toISOString().slice(0, 10),
      subtitle: `${doc.kind.toUpperCase()} · ${doc.engineUsed}`,
      matchedReasons: reasons.length > 0 ? reasons : ['업로드 문서 일치'],
      snippet: doc.extractedText.replace(/\s+/g, ' ').slice(0, 150),
      projectCode: detectedProject ?? undefined,
    });
  }

  hits.sort((a, b) => b.date.localeCompare(a.date));
  return { parsed, hits };
}

/**
 * Deterministically extracts schedule / deadline / meeting events from emails & meeting minutes
 * without requiring any LLM API key.
 */
export function extractDeterministicSchedules(
  emails: MailItem[],
  projects: ProjectDefinition[],
  meetingMinutes: MeetingMinutesResult[] = []
): CalendarScheduleEvent[] {
  const events: CalendarScheduleEvent[] = [];
  const seen = new Set<string>();

  const SCHEDULE_SIGNAL = /(회의|미팅|일정|마감|제출|검토|브리핑|보고|킥오프|점검|kickoff|deadline|meeting|review)/i;

  for (const mail of emails) {
    const combined = `${mail.subject}\n${mail.body}`;
    if (!SCHEDULE_SIGNAL.test(combined)) continue;

    const projectCode = detectProjectCode(combined, projects) ?? undefined;

    // Extract explicit ISO dates or Korean (YYYY년) M월 D일 dates inside body/subject
    let eventDate = mail.date;
    const isoMatch = combined.match(/\b(20\d{2})[-./](\d{1,2})[-./](\d{1,2})\b/);
    const korMatch = combined.match(/(?:(20\d{2})\s*년\s*)?(\d{1,2})\s*월\s*(\d{1,2})\s*일/);
    if (isoMatch) {
      eventDate = `${isoMatch[1]}-${isoMatch[2].padStart(2, '0')}-${isoMatch[3].padStart(2, '0')}`;
    } else if (korMatch) {
      const year = korMatch[1] || mail.date.slice(0, 4) || String(new Date().getFullYear());
      eventDate = `${year}-${korMatch[2].padStart(2, '0')}-${korMatch[3].padStart(2, '0')}`;
    }

    const timeMatch = combined.match(
      /(?:(오전|오후|AM|PM)\s*)?(\d{1,2})\s*(?:시\s*(?:(\d{1,2})\s*분)?|:\s*(\d{2}))(?:\s*(AM|PM))?/i
    );
    let timeLabel = mail.timeLabel || '09:00';
    if (timeMatch) {
      const meridiem = (timeMatch[1] || timeMatch[5] || '').toUpperCase();
      let hour = Number.parseInt(timeMatch[2], 10);
      const minute = Number.parseInt(timeMatch[3] || timeMatch[4] || '0', 10);
      if (Number.isFinite(hour) && hour >= 0 && hour <= 23 && Number.isFinite(minute) && minute >= 0 && minute <= 59) {
        if ((meridiem === '오후' || meridiem === 'PM') && hour >= 1 && hour <= 11) {
          hour += 12;
        } else if ((meridiem === '오전' || meridiem === 'AM') && hour === 12) {
          hour = 0;
        }
        timeLabel = `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
      }
    }

    const key = `mail-${mail.id}-${eventDate}`;
    if (seen.has(key)) continue;
    seen.add(key);

    const isDeadline = /(마감|제출|deadline|기한)/i.test(combined);
    events.push({
      id: key,
      title: mail.subject,
      date: eventDate,
      timeLabel,
      category: isDeadline ? 'deadline' : 'email-extracted',
      projectCode,
      sourceEmailId: mail.id,
      sourceEmailSubject: mail.subject,
      notes: `${mail.fromName}: ${mail.snippet}`,
    });
  }

  for (const minutes of meetingMinutes) {
    for (let idx = 0; idx < minutes.actionItems.length; idx += 1) {
      const item = minutes.actionItems[idx];
      const iso = item.deadline.match(/(20\d{2})[-./](\d{1,2})[-./](\d{1,2})/);
      const eventDate = iso
        ? `${iso[1]}-${iso[2].padStart(2, '0')}-${iso[3].padStart(2, '0')}`
        : minutes.meetingDate || new Date().toISOString().slice(0, 10);
      const key = `minutes-${minutes.id}-${idx}`;
      if (seen.has(key)) continue;
      seen.add(key);

      events.push({
        id: key,
        title: `[회의 액션] ${item.task}`,
        date: eventDate,
        timeLabel: '18:00',
        category: 'action-item',
        notes: `담당: ${item.owner} · 출처 회의록: ${minutes.title}`,
      });
    }
  }

  return events.sort((a, b) => `${a.date} ${a.timeLabel}`.localeCompare(`${b.date} ${b.timeLabel}`));
}

/**
 * Generates a complete structured Work Log & Case Ledger report purely with deterministic code
 * (Zero LLM required!). Can be used standalone offline or refined by AI when desired.
 */
export function buildDeterministicWorkLog(params: {
  emails: MailItem[];
  documents?: ParsedDocumentSummary[];
  projects?: ProjectDefinition[];
  periodLabel: string;
  templateName?: 'standard' | 'project-case' | 'executive';
}): AnalysisResult {
  const { emails, documents = [], projects = [], periodLabel, templateName = 'standard' } = params;
  const received = emails.filter((e) => e.folder === 'inbox');
  const sent = emails.filter((e) => e.folder === 'sent');

  // Group by project code
  const byProject = new Map<string, MailItem[]>();
  for (const mail of emails) {
    const code = detectProjectCode(`${mail.subject} ${mail.body} ${mail.category || ''}`, projects) || '미분류 일반';
    const arr = byProject.get(code) ?? [];
    arr.push(mail);
    byProject.set(code, arr);
  }

  const workPerformed = Array.from(byProject.entries()).map(([code, items]) => {
    const sentItems = items.filter((i) => i.folder === 'sent');
    const attNames = items
      .flatMap((i) => i.attachments ?? [])
      .map((a) => a.filename);
    const status: '완료' | '진행 중' | '후속 조치 필요' =
      sentItems.length > 0 ? '진행 중' : '후속 조치 필요';
    return {
      title: `[${code}] 관련 메일 ${items.length}건 (발신 ${sentItems.length} · 수신 ${items.length - sentItems.length})`,
      status,
      detail: [
        `대표 건: ${items[0]?.subject || '-'}`,
        attNames.length > 0 ? `첨부파일 ${attNames.length}건(${attNames.slice(0, 3).join(', ')})` : '',
      ]
        .filter(Boolean)
        .join(' · '),
      relatedEmails: items.slice(0, 5).map((i) => `${i.date} ${i.subject}`),
    };
  });

  const incomingHighlights = received.slice(0, 8).map((mail) => {
    const isUrgent = /(긴급|마감|요청|확인|회신|승인|urgent|asap)/i.test(`${mail.subject} ${mail.body}`);
    return {
      sender: mail.fromName,
      subject: mail.subject,
      intent: mail.snippet || mail.body.slice(0, 100),
      priority: (isUrgent ? '높음' : '보통') as '높음' | '보통' | '참고',
      actionNeeded: isUrgent ? '기한 내 검토 및 회신 필요' : '내용 확인 및 프로젝트 기록 반영',
    };
  });

  const docLines = documents.map(
    (d) => `[문서:${d.kind.toUpperCase()}] ${d.filename} (${d.engineUsed}${d.rowCount ? `, ${d.rowCount}행` : ''}${d.pageCount ? `, ${d.pageCount}쪽` : ''})`
  );

  const projectSummaries = Array.from(byProject.entries()).map(([code, items]) => {
    const matchedDef = projects.find((p) => p.code.toLowerCase() === code.toLowerCase());
    const sentItems = items.filter((i) => i.folder === 'sent');
    const receivedItems = items.filter((i) => i.folder === 'inbox');
    const attNames = Array.from(
      new Set(items.flatMap((i) => i.attachments ?? []).map((a) => a.filename))
    );
    const status: '완료' | '진행 중' | '후속 조치 필요' =
      sentItems.length > 0 ? '진행 중' : '후속 조치 필요';
    return {
      projectCode: code,
      projectName: matchedDef?.name || code,
      status,
      myActions:
        sentItems.length > 0
          ? sentItems.slice(0, 5).map((m) => `[${m.date}] 발신·회신: ${m.subject}`)
          : [`수신 메일 ${receivedItems.length}건 접수 및 검토 대기`],
      keyDecisions: items.slice(0, 3).map((m) => `${m.date} ${m.fromName}: ${m.subject}`),
      attachmentsMentioned: attNames.slice(0, 6),
      nextAction:
        receivedItems.length > 0
          ? `최근 수신 건(${receivedItems[0].subject}) 검토 및 회신`
          : '후속 일정 모니터링',
    };
  });

  return {
    id: `det-${Date.now()}`,
    createdAt: new Date().toISOString(),
    providerUsed: 'builtin',
    providerModel: `코드 기반 결정론적 분석 엔진 (${templateName}, LLM 미사용)`,
    periodLabel,
    scopeLabel: `수신 ${received.length}건 · 발신 ${sent.length}건 · 문서 ${documents.length}건`,
    analyzedCount: emails.length,
    receivedCount: received.length,
    sentCount: sent.length,
    overview: `${periodLabel} 기간 동안 총 ${emails.length}건의 실제 메일(수신 ${received.length}건, 발신 ${sent.length}건)과 ${documents.length}건의 첨부·문서를 LLM 호출 없이 규칙 기반으로 자동 분류·집계했습니다. 총 ${byProject.size}개 프로젝트 코드 그룹이 식별되었습니다.`,
    workPerformed,
    projectSummaries,
    incomingHighlights,
    workLog: {
      title: `${periodLabel} 자동 코드 기반 업무 일지 (${templateName})`,
      periodLabel,
      executiveSummary: `총 ${emails.length}건의 메일과 ${documents.length}건의 문서를 프로젝트 코드(${Array.from(byProject.keys()).join(', ')})별로 자동 집계한 1차 구조화 일지입니다.`,
      completedTasks: sent.slice(0, 8).map((m) => `[${m.date} ${m.timeLabel}] 발신: ${m.subject}`),
      inProgressTasks: received
        .slice(0, 8)
        .map((m) => `[${m.date} ${m.timeLabel}] 수신(${m.fromName}): ${m.subject}`),
      keyCommunications: [
        ...emails
          .filter((m) => m.hasAttachment || (m.attachments && m.attachments.length > 0))
          .slice(0, 6)
          .map((m) => `[첨부포함] ${m.date} ${m.subject}`),
        ...docLines,
      ],
      issuesAndRisks: incomingHighlights
        .filter((h) => h.priority === '높음')
        .map((h) => `[우선 확인] ${h.sender} - ${h.subject}`),
      nextStepPlan: [
        '고우선순위 수신 메일 회신 및 일정 캘린더 반영',
        '필요 시 상단 AI 분석 실행 버튼으로 정밀 요약·블로그 초안 자동 고도화',
      ],
    },
    blogPost: {
      title: `${periodLabel} 실무 기록 아카이브 요약`,
      subtitle: '코드 기반 자동 분류 결과 (AI 다듬기 전 초안)',
      readingTimeMinutes: 3,
      tags: Array.from(byProject.keys()).slice(0, 4),
      intro: `${periodLabel} 기간의 Outlook 메일 ${emails.length}건과 첨부 문서 ${documents.length}건을 프로젝트 코드별로 정리한 1차 기록입니다.`,
      sections: workPerformed.slice(0, 3).map((item) => ({
        heading: item.title,
        body: `${item.detail}\n관련 메일: ${item.relatedEmails.join(' / ')}`,
      })),
      conclusion: '이 초안은 LLM API 호출 없이 자동 코드로 생성되었으며, AI 엔진 실행 시 자연스러운 발행용 문체로 자동 확장됩니다.',
    },
  };
}

/**
 * Exports loaded MailItem records to a standard RFC 4155 .mbox archive (importable into Thunderbird / Outlook / MailCraft).
 */
export function exportEmailsToMbox(emails: MailItem[]): string {
  if (!Array.isArray(emails)) return '';
  return emails
    .filter((mail): mail is MailItem => Boolean(mail && typeof mail === 'object'))
    .map((mail) => {
      const dateSeed = mail.timestamp || (mail.date ? `${mail.date}T09:00:00Z` : '');
      const dt = dateSeed ? new Date(dateSeed) : new Date();
      const fromLineDate = Number.isNaN(dt.getTime()) ? new Date().toUTCString() : dt.toUTCString();
      const safeBody = (typeof mail.body === 'string' ? mail.body : '').replace(/^From /gm, '>From ');
      const fromEmail =
        (typeof mail.fromEmail === 'string' && mail.fromEmail.trim()) || 'unknown@local';
      const fromName =
        (typeof mail.fromName === 'string' && mail.fromName.trim()) || fromEmail;
      const toList = Array.isArray(mail.to)
        ? mail.to.filter((item): item is string => typeof item === 'string' && item.trim().length > 0)
        : [];
      const subject = typeof mail.subject === 'string' ? mail.subject : '(제목 없음)';
      const folderName =
        (typeof mail.folderName === 'string' && mail.folderName.trim()) ||
        (mail.folder === 'sent' ? 'Sent Items' : 'Inbox');

      return [
        `From ${fromEmail} ${fromLineDate}`,
        `Date: ${fromLineDate}`,
        `From: ${fromName} <${fromEmail}>`,
        `To: ${toList.join(', ') || 'me@workspace.kr'}`,
        `Subject: ${subject}`,
        `X-Outlook-Folder: ${folderName}`,
        `Content-Type: text/plain; charset=UTF-8`,
        '',
        safeBody,
        '',
      ].join('\r\n');
    })
    .join('\r\n');
}
