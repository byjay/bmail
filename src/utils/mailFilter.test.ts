import test from 'node:test';
import assert from 'node:assert/strict';
import {
  chunkRecordsByBudget,
  cleanEmailBody,
  countActiveGlobalSearchFilters,
  DEFAULT_GLOBAL_SEARCH_FILTERS,
  filterEmails,
  formatBlogPostMarkdown,
  formatWorkLogMarkdown,
  getAvailableDates,
  isPromotionalEmail,
  maskSensitiveData,
  matchesGlobalSearchFilters,
  resolveSearchDateRange,
  validateAnalysisFilter,
  validateGlobalSearchFilters,
} from './mailFilter';
import {
  buildCommandsHtml,
  buildOutlookManifestXml,
  buildUnifiedManifestJson,
  buildWindowsRegistryCatalog,
  normalizeBaseUrl,
} from './outlookManifest';
import { createIconPngBuffer } from './iconPng';
import type { AnalysisResult, MailItem } from '../types/mail';

const TEST_EMAILS: MailItem[] = [
  {
    id: 'msg-20261119-01',
    folder: 'inbox',
    subject: '결제 API 지연 장애 원인 분석 요청',
    fromName: '박지훈',
    fromEmail: 'jh.park@finflow.kr',
    to: ['minseo@finflow.kr'],
    date: '2026-11-19',
    timeLabel: '10:00',
    timestamp: '2026-11-19T09:15:00+09:00',
    snippet: '피크타임 p95 응답속도가 1.8초까지 상승했습니다.',
    body: '피크타임 p95 응답속도가 1.8초까지 상승했습니다. Read Replica 분리 검토 부탁드립니다.',
    isRead: false,
    hasAttachment: true,
    importance: 'high',
    source: 'pst-import',
  },
  {
    id: 'msg-20261119-02',
    folder: 'sent',
    subject: 'Re: 결제 API 지연 장애 원인 분석 요청',
    fromName: '김민서',
    fromEmail: 'minseo@finflow.kr',
    to: ['jh.park@finflow.kr'],
    date: '2026-11-19',
    timeLabel: '10:00',
    timestamp: '2026-11-19T11:40:00+09:00',
    snippet: 'PR #482로 Read Replica 라우팅 및 백오프 패치 적용 완료했습니다.',
    body: 'PR #482로 Read Replica 라우팅 및 백오프 패치 적용 완료했습니다. p95가 210ms로 안정화되었습니다.',
    isRead: true,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
  },
  {
    id: 'msg-20261119-03',
    folder: 'inbox',
    subject: '4분기 파트너십 정산 일정 공유',
    fromName: '이서연',
    fromEmail: 'sy.lee@partner.co.kr',
    to: ['minseo@finflow.kr'],
    date: '2026-11-19',
    timeLabel: '10:00',
    timestamp: '2026-11-19T14:20:00+09:00',
    snippet: '정산서 마감은 금주 금요일 오후 3시입니다.',
    body: '정산서 마감은 금주 금요일 오후 3시입니다.',
    isRead: true,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
  },
  {
    id: 'msg-20261119-04',
    folder: 'sent',
    subject: '주간 개발 스프린트 완료 보고',
    fromName: '김민서',
    fromEmail: 'minseo@finflow.kr',
    to: ['team@finflow.kr'],
    date: '2026-11-19',
    timeLabel: '10:00',
    timestamp: '2026-11-19T17:50:00+09:00',
    snippet: '금주 목표 작업 5건 중 4건 완료했습니다.',
    body: '금주 목표 작업 5건 중 4건 완료했습니다.',
    isRead: true,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
  },
  {
    id: 'msg-20261118-01',
    folder: 'inbox',
    subject: '디자인 QA 피드백 리스트',
    fromName: '최유진',
    fromEmail: 'yj.choi@finflow.kr',
    to: ['minseo@finflow.kr'],
    date: '2026-11-18',
    timeLabel: '16:10',
    timestamp: '2026-11-18T16:10:00+09:00',
    snippet: '사이드바 간격 조정 요청드립니다.',
    body: '사이드바 간격 조정 요청드립니다.',
    isRead: true,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
  },
  {
    id: 'msg-20261118-02',
    folder: 'sent',
    subject: '스테이징 배포 공지',
    fromName: '김민서',
    fromEmail: 'minseo@finflow.kr',
    to: ['team@finflow.kr'],
    date: '2026-11-18',
    timeLabel: '16:10',
    timestamp: '2026-11-18T18:00:00+09:00',
    snippet: '스테이징 환경 v2.4 반영 완료했습니다.',
    body: '스테이징 환경 v2.4 반영 완료했습니다.',
    isRead: true,
    hasAttachment: false,
    importance: 'normal',
    source: 'pst-import',
  },
  {
    id: 'msg-20261117-01',
    folder: 'inbox',
    subject: '신규 보안 감사 체크리스트 안내',
    fromName: '정보보안팀',
    fromEmail: 'sec@finflow.kr',
    to: ['minseo@finflow.kr'],
    date: '2026-11-17',
    timeLabel: '10:00',
    timestamp: '2026-11-17T10:00:00+09:00',
    snippet: 'API 키 및 개인정보 마스킹 정책 점검 요청드립니다.',
    body: 'API 키 및 개인정보 마스킹 정책 점검 요청드립니다.',
    isRead: true,
    hasAttachment: true,
    importance: 'high',
    source: 'pst-import',
  },
];

test('single-day received/sent/all filtering returns accurate email subsets', () => {
  const allOn19 = filterEmails(TEST_EMAILS, {
    dateMode: 'single',
    singleDate: '2026-11-19',
    startDate: '2026-11-17',
    endDate: '2026-11-19',
    scope: 'all',
    selectedIds: [],
  });
  assert.equal(allOn19.length, 4);

  const receivedOn19 = filterEmails(TEST_EMAILS, {
    dateMode: 'single',
    singleDate: '2026-11-19',
    startDate: '2026-11-17',
    endDate: '2026-11-19',
    scope: 'received',
    selectedIds: [],
  });
  assert.equal(receivedOn19.length, 2);
  assert.ok(receivedOn19.every((m) => m.folder === 'inbox'));

  const sentOn19 = filterEmails(TEST_EMAILS, {
    dateMode: 'single',
    singleDate: '2026-11-19',
    startDate: '2026-11-17',
    endDate: '2026-11-19',
    scope: 'sent',
    selectedIds: [],
  });
  assert.equal(sentOn19.length, 2);
  assert.ok(sentOn19.every((m) => m.folder === 'sent'));
});

test('custom date range and selected emails filtering work accurately', () => {
  const rangeRes = validateAnalysisFilter(TEST_EMAILS, {
    dateMode: 'range',
    singleDate: '2026-11-19',
    startDate: '2026-11-18',
    endDate: '2026-11-19',
    scope: 'all',
    selectedIds: [],
  });
  assert.equal(rangeRes.valid, true);
  assert.equal(rangeRes.matchedEmails.length, 6);
  assert.equal(rangeRes.periodLabel, '2026-11-18 ~ 2026-11-19');

  const selectedRes = validateAnalysisFilter(TEST_EMAILS, {
    dateMode: 'single',
    singleDate: '2026-11-19',
    startDate: '2026-11-18',
    endDate: '2026-11-19',
    scope: 'selected',
    selectedIds: ['msg-20261119-01', 'msg-20261117-01'],
  });
  assert.equal(selectedRes.valid, true);
  assert.equal(selectedRes.matchedEmails.length, 2);
  assert.equal(selectedRes.periodLabel, '2026-11-17 ~ 2026-11-19');
});

test('validation rejects inverted date ranges and empty selections', () => {
  const inverted = validateAnalysisFilter(TEST_EMAILS, {
    dateMode: 'range',
    singleDate: '2026-11-19',
    startDate: '2026-11-20',
    endDate: '2026-11-18',
    scope: 'all',
    selectedIds: [],
  });
  assert.equal(inverted.valid, false);
  assert.match(inverted.error ?? '', /시작 일자가 종료 일자보다 늦을 수 없습니다/);

  const emptySelected = validateAnalysisFilter(TEST_EMAILS, {
    dateMode: 'single',
    singleDate: '2026-11-19',
    startDate: '2026-11-18',
    endDate: '2026-11-19',
    scope: 'selected',
    selectedIds: [],
  });
  assert.equal(emptySelected.valid, false);
  assert.match(emptySelected.error ?? '', /선택한 메일이 없습니다/);
});

test('maskSensitiveData hides emails, Korean phone numbers, IPv4, API keys, and KRW amounts', () => {
  const input =
    '연락처: 010-4821-9034, 이메일: jh.park@finflow.kr, 서버 IP 10.24.18.92, 토큰 gsk_test1234567890abcdef, 계약 금액 4,800만원 기준';
  const masked = maskSensitiveData(input);
  assert.ok(!masked.includes('010-4821-9034'));
  assert.ok(!masked.includes('jh.park@finflow.kr'));
  assert.ok(!masked.includes('10.24.18.92'));
  assert.ok(!masked.includes('gsk_test1234567890abcdef'));
  assert.ok(!masked.includes('4,800만원'));
  assert.ok(masked.includes('[연락처 비공개]'));
  assert.ok(masked.includes('[이메일 비공개]'));
  assert.ok(masked.includes('[IP 비공개]'));
  assert.ok(masked.includes('[API키 비공개]'));
  assert.ok(masked.includes('[금액 비공개]'));
});

test('promotional filter, email body cleaner, and Groq budget chunker behave accurately', () => {
  assert.equal(
    isPromotionalEmail({
      subject: '(광고) 연말 특가 뉴스레터',
      fromEmail: 'marketing@shop.kr',
      body: '할인 쿠폰을 확인하세요.',
    }),
    true
  );
  assert.equal(
    isPromotionalEmail({
      subject: '결제 API 성능 점검 결과',
      fromEmail: 'dev@finflow.kr',
      body: 'Read Replica 적용 후 응답 지연이 개선되었습니다.',
    }),
    false
  );

  const rawBody = '핵심 보고 내용입니다.\r\n\r\n\r\n--\n김민서 드림\nCONFIDENTIALITY NOTICE: Do not distribute';
  const cleaned = cleanEmailBody(rawBody);
  assert.equal(cleaned, '핵심 보고 내용입니다.');

  const records = Array.from({ length: 6 }, (_, i) => `<email_record id="${i}">${'가'.repeat(1800)}</email_record>`);
  const chunks = chunkRecordsByBudget(records, 4000);
  assert.ok(chunks.length >= 3);
  assert.ok(chunks.every((c) => c.length <= 4000));
});

test('getAvailableDates groups emails by descending date', () => {
  const dates = getAvailableDates(TEST_EMAILS);
  assert.equal(dates.length, 3);
  assert.equal(dates[0].date, '2026-11-19');
  assert.equal(dates[0].received, 2);
  assert.equal(dates[0].sent, 2);
});

test('Outlook XML & Unified JSON manifests include 2026 Mailbox 1.15, No-Item Context, Multi-Select, and Pinning', () => {
  assert.equal(normalizeBaseUrl('example.com/'), 'https://example.com');

  const xml = buildOutlookManifestXml({
    baseUrl: 'https://maillog.example.com/',
    displayName: 'MailLog AI',
  });
  assert.ok(xml.includes('MinVersion="1.15"'));
  assert.ok(xml.includes('<SupportsPinning>true</SupportsPinning>'));
  assert.ok(xml.includes('<SupportsNoItemContext>true</SupportsNoItemContext>'));
  assert.ok(xml.includes('<SupportsMultiSelect>true</SupportsMultiSelect>'));
  assert.ok(xml.includes('<SupportsSharedFolders>true</SupportsSharedFolders>'));
  assert.ok(xml.includes('https://maillog.example.com/taskpane'));
  assert.ok(xml.includes('https://maillog.example.com/assets/icon-16.png'));
  assert.ok(xml.includes('https://maillog.example.com/assets/icon-32.png'));
  assert.ok(xml.includes('https://maillog.example.com/assets/icon-80.png'));
  assert.ok(xml.includes('MessageReadCommandSurface'));
  assert.ok(xml.includes('MessageComposeCommandSurface'));

  const unified = JSON.parse(buildUnifiedManifestJson({ baseUrl: 'https://maillog.example.com' }));
  assert.equal(unified.manifestVersion, '1.19');
  assert.equal(unified.extensions[0].runtimes[0].actions[0].pinnable, true);

  const reg = buildWindowsRegistryCatalog('\\\\corp\\addins\\maillog');
  assert.ok(reg.includes('TrustedCatalogs'));
  assert.ok(reg.includes('\\\\\\\\corp\\\\addins\\\\maillog'));

  const cmdHtml = buildCommandsHtml();
  assert.ok(cmdHtml.includes('appsforoffice.microsoft.com/lib/1/hosted/office.js'));
});

test('createIconPngBuffer outputs valid PNG binary with exact dimensions', () => {
  for (const size of [16, 32, 64, 80, 128]) {
    const buf = createIconPngBuffer(size);
    // PNG magic bytes
    assert.deepEqual(Array.from(buf.subarray(0, 8)), [137, 80, 78, 71, 13, 10, 26, 10]);
    // Width & Height in IHDR
    assert.equal(buf.readUInt32BE(16), size);
    assert.equal(buf.readUInt32BE(20), size);
  }
});

test('Markdown formatters output structured work log and blog post', () => {
  const sample: AnalysisResult = {
    id: 'test-1',
    createdAt: '2026-11-19T12:00:00Z',
    providerUsed: 'groq',
    providerModel: 'Groq · openai/gpt-oss-120b',
    periodLabel: '2026-11-19',
    scopeLabel: '받은 메일 + 보낸 메일',
    analyzedCount: 4,
    receivedCount: 2,
    sentCount: 2,
    overview: '결제 API 지연 패치 완료',
    workPerformed: [],
    incomingHighlights: [],
    workLog: {
      title: '2026-11-19 업무 일지',
      periodLabel: '2026-11-19',
      executiveSummary: '핵심 요약',
      completedTasks: ['PR #482 리뷰 완료'],
      inProgressTasks: ['디자인 QA 수정'],
      keyCommunications: ['박지훈 수석과 배포 일정 확정'],
      issuesAndRisks: ['피크타임 DB 부하 주의'],
      nextStepPlan: ['모니터링 확인'],
    },
    blogPost: {
      title: '결제 API p95 응답속도를 1.8초에서 210ms로 줄인 기록',
      subtitle: 'Read Replica 분리와 Exponential Backoff 실무 적용기',
      readingTimeMinutes: 5,
      tags: ['백엔드', '성능최적화'],
      intro: '도입부 내용',
      sections: [{ heading: '원인 파악', body: '본문 내용' }],
      conclusion: '결론 내용',
    },
  };

  const logMd = formatWorkLogMarkdown(sample);
  assert.ok(logMd.includes('# 2026-11-19 업무 일지'));
  assert.ok(logMd.includes('- [x] PR #482 리뷰 완료'));

  const blogMd = formatBlogPostMarkdown(sample);
  assert.ok(blogMd.includes('# 결제 API p95 응답속도를 1.8초에서 210ms로 줄인 기록'));
  assert.ok(blogMd.includes('#백엔드 #성능최적화'));
});

test('EML, MBOX parsers and Korean Outlook folder classifier extract emails accurately', async () => {
  const { classifyOutlookFolder, decodeMimeHeader, parseEmlMessage, parseMboxArchive } = await import(
    './mailFileParsers'
  );

  assert.equal(classifyOutlookFolder('받은 편지함'), 'inbox');
  assert.equal(classifyOutlookFolder('Inbox'), 'inbox');
  assert.equal(classifyOutlookFolder('보낸 편지함'), 'sent');
  assert.equal(classifyOutlookFolder('Sent Items'), 'sent');

  const encodedSubject = `=?UTF-8?B?${Buffer.from('아웃룩 PST 복구 테스트', 'utf8').toString('base64')}?=`;
  const encodedFromName = `=?UTF-8?B?${Buffer.from('김민서', 'utf8').toString('base64')}?=`;
  assert.equal(decodeMimeHeader(encodedSubject), '아웃룩 PST 복구 테스트');

  const rawEml = [
    `From: ${encodedFromName} <minseo@finflow.kr>`,
    'To: team@finflow.kr',
    `Subject: ${encodedSubject}`,
    'Date: Thu, 19 Nov 2026 10:30:00 +0900',
    'X-Outlook-Folder: 보낸 편지함',
    'Content-Type: text/plain; charset=UTF-8',
    '',
    'PST 아카이브에서 추출된 실제 업무 메일 본문입니다.',
  ].join('\r\n');

  const parsedEml = parseEmlMessage(rawEml);
  assert.equal(parsedEml.folder, 'sent');
  assert.equal(parsedEml.subject, '아웃룩 PST 복구 테스트');
  assert.equal(parsedEml.fromName, '김민서');
  assert.equal(parsedEml.fromEmail, 'minseo@finflow.kr');
  assert.equal(parsedEml.date, '2026-11-19');
  assert.ok(parsedEml.body.includes('PST 아카이브에서 추출된'));

  const mbox = `From sender@example.com Thu Nov 19 09:00:00 2026\n${rawEml}\n\nFrom other@example.com Wed Nov 18 11:00:00 2026\nFrom: Alice <alice@example.com>\nSubject: Second Mail\nDate: Wed, 18 Nov 2026 11:00:00 +0900\n\nSecond message body.`;
  const parsedMbox = parseMboxArchive(mbox, 'backup.mbox');
  assert.equal(parsedMbox.emails.length, 2);
  assert.equal(parsedMbox.emails[1].subject, 'Second Mail');
});

test('global search filters accurately scope by From, To, Has Attachments, and Date Range', () => {
  // 1. From filter (name or email, case-insensitive)
  const byFromName = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', { ...DEFAULT_GLOBAL_SEARCH_FILTERS, from: '박지훈' })
  );
  assert.equal(byFromName.length, 1);
  assert.equal(byFromName[0].id, 'msg-20261119-01');

  const byFromEmail = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', { ...DEFAULT_GLOBAL_SEARCH_FILTERS, from: 'MINSEO@FINFLOW.KR' })
  );
  assert.equal(byFromEmail.length, 3);

  // 2. To filter (matches `to` recipients and `cc` recipients)
  const byTo = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', { ...DEFAULT_GLOBAL_SEARCH_FILTERS, to: 'team@finflow.kr' })
  );
  assert.equal(byTo.length, 2);

  const mailWithCc: MailItem = {
    ...TEST_EMAILS[0],
    id: 'msg-cc-test',
    cc: ['audit@finflow.kr'],
  };
  assert.equal(
    matchesGlobalSearchFilters(mailWithCc, '', { ...DEFAULT_GLOBAL_SEARCH_FILTERS, to: 'audit@' }),
    true
  );

  // 3. Has Attachments filter (checks hasAttachment flag and attachments array)
  const withAttachments = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', { ...DEFAULT_GLOBAL_SEARCH_FILTERS, hasAttachments: true })
  );
  assert.equal(withAttachments.length, 2);
  assert.deepEqual(
    withAttachments.map((m) => m.id),
    ['msg-20261119-01', 'msg-20261117-01']
  );

  const mailWithAttachmentArrayOnly: MailItem = {
    ...TEST_EMAILS[2],
    hasAttachment: false,
    attachments: [{ id: 'a1', filename: 'spec.pdf', mimeType: 'application/pdf', kind: 'pdf' }],
  };
  assert.equal(
    matchesGlobalSearchFilters(mailWithAttachmentArrayOnly, '', {
      ...DEFAULT_GLOBAL_SEARCH_FILTERS,
      hasAttachments: true,
    }),
    true
  );

  // 4. Date Range filter (inclusive both bounds, one-sided bounds, and presets)
  const dateRangeMatched = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', {
      ...DEFAULT_GLOBAL_SEARCH_FILTERS,
      datePreset: 'custom',
      dateFrom: '2026-11-18',
      dateTo: '2026-11-18',
    })
  );
  assert.equal(dateRangeMatched.length, 2);

  const fromOnlyMatched = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', {
      ...DEFAULT_GLOBAL_SEARCH_FILTERS,
      datePreset: 'custom',
      dateFrom: '2026-11-19',
      dateTo: '',
    })
  );
  assert.equal(fromOnlyMatched.length, 4);

  const toOnlyMatched = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', {
      ...DEFAULT_GLOBAL_SEARCH_FILTERS,
      datePreset: 'custom',
      dateFrom: '',
      dateTo: '2026-11-17',
    })
  );
  assert.equal(toOnlyMatched.length, 1);
  assert.equal(toOnlyMatched[0].id, 'msg-20261117-01');

  // 5. Combined query + multiple scoped filters
  const combined = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, 'API', {
      from: 'finflow.kr',
      to: 'minseo@finflow.kr',
      hasAttachments: true,
      datePreset: 'custom',
      dateFrom: '2026-11-17',
      dateTo: '2026-11-19',
    })
  );
  assert.equal(combined.length, 2);
});

test('global search date range presets, active count, and validation handle edge cases', () => {
  assert.deepEqual(resolveSearchDateRange('any', '2026-11-20'), { dateFrom: '', dateTo: '' });
  assert.deepEqual(resolveSearchDateRange('today', '2026-11-20'), {
    dateFrom: '2026-11-20',
    dateTo: '2026-11-20',
  });
  assert.deepEqual(resolveSearchDateRange('7d', '2026-11-20'), {
    dateFrom: '2026-11-14',
    dateTo: '2026-11-20',
  });
  assert.deepEqual(resolveSearchDateRange('30d', '2026-11-20'), {
    dateFrom: '2026-10-22',
    dateTo: '2026-11-20',
  });

  assert.equal(countActiveGlobalSearchFilters(DEFAULT_GLOBAL_SEARCH_FILTERS), 0);
  assert.equal(
    countActiveGlobalSearchFilters({
      from: 'park',
      to: 'minseo',
      hasAttachments: true,
      datePreset: '7d',
      dateFrom: '2026-11-14',
      dateTo: '2026-11-20',
    }),
    4
  );

  const inverted = validateGlobalSearchFilters({
    ...DEFAULT_GLOBAL_SEARCH_FILTERS,
    datePreset: 'custom',
    dateFrom: '2026-11-21',
    dateTo: '2026-11-18',
  });
  assert.equal(inverted.valid, false);
  assert.match(inverted.error ?? '', /시작 날짜가 종료 날짜보다 늦을 수 없습니다/);

  // Inverted range defensively matches 0 emails
  const matchedOnInverted = TEST_EMAILS.filter((m) =>
    matchesGlobalSearchFilters(m, '', {
      ...DEFAULT_GLOBAL_SEARCH_FILTERS,
      datePreset: 'custom',
      dateFrom: '2026-11-21',
      dateTo: '2026-11-18',
    })
  );
  assert.equal(matchedOnInverted.length, 0);
});
