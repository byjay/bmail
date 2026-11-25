import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import ExcelJS from 'exceljs';
import type { MailItem, MeetingMinutesResult, ParsedDocumentSummary } from '../types/mail';
import { sanitizeAppConfig, buildEnvSecretStatuses } from './appConfig';
import {
  buildDeterministicWorkLog,
  detectProjectCode,
  exportEmailsToMbox,
  extractDeterministicSchedules,
  normalizeConversationSubject,
  parseNaturalSearchQuery,
  runDeterministicConversationalSearch,
  type ProjectDefinition,
} from './deterministicAnalyzer';
import { parseEmlMessage } from './mailFileParsers';
import { parseUploadedDocument } from '../services/documentParser';
import { probeExternalEngine } from '../services/externalEngineBridge';

const PROJECTS: ProjectDefinition[] = [
  {
    id: 'p1',
    code: 'PRJ-01',
    name: '차세대 ERP 구축',
    keywords: ['ERP', '견적서', '발주'],
    color: '#0F6CBD',
    isFavorite: true,
    createdAt: '2026-11-01T00:00:00Z',
  },
];

const EMAILS: MailItem[] = [
  {
    id: 'm1',
    folder: 'inbox',
    folderName: '받은 편지함',
    subject: 'RE: [PRJ-01] ERP 최종 견적 PDF 검토 요청',
    fromName: '김민수',
    fromEmail: 'minsu@partner.co.kr',
    to: ['me@company.kr'],
    date: '2026-11-18',
    timeLabel: '14:30',
    timestamp: '2026-11-18T14:30:00Z',
    snippet: '11월 20일 15:00 회의 전까지 PDF 계약서 확인 부탁드립니다.',
    body: '안녕하세요.\n11월 20일 15:00 회의 전까지 첨부된 PDF 계약서와 견적서 검토 및 회신 요청드립니다.\nFrom 영업팀 드림',
    hasAttachment: true,
    attachments: [
      {
        id: 'att-1',
        filename: 'PRJ-01_계약서.pdf',
        mimeType: 'application/pdf',
        kind: 'pdf',
        sizeBytes: 24500,
      },
    ],
    isRead: false,
    source: 'pst-import',
  },
  {
    id: 'm2',
    folder: 'sent',
    folderName: '보낸 편지함',
    subject: '[PRJ-01] 발주 단가 엑셀 회신',
    fromName: '나',
    fromEmail: 'me@company.kr',
    to: ['minsu@partner.co.kr'],
    date: '2026-11-19',
    timeLabel: '09:15',
    timestamp: '2026-11-19T09:15:00Z',
    snippet: '엑셀 단가표 검토 완료해 회신합니다.',
    body: '요청하신 ERP 단가 엑셀 파일 검토 완료해 보냅니다.',
    hasAttachment: true,
    attachments: [
      {
        id: 'att-2',
        filename: '단가표.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        kind: 'excel',
        sizeBytes: 12000,
      },
    ],
    isRead: true,
    source: 'pst-import',
  },
];

test('normalizeConversationSubject strips RE/FW and Korean reply prefixes', () => {
  assert.equal(
    normalizeConversationSubject('RE: [회신] FW: [PRJ-01] 최종 견적 검토'),
    '[prj-01] 최종 견적 검토'
  );
});

test('detectProjectCode matches registered projects and fallback bracket codes', () => {
  assert.equal(detectProjectCode('ERP 단가표 검토 부탁드립니다', PROJECTS), 'PRJ-01');
  assert.equal(detectProjectCode('[DEV-204] API 장애 대응 보고', PROJECTS), 'DEV-204');
  assert.equal(detectProjectCode('일반 인사 안내', PROJECTS), null);
});

test('parseNaturalSearchQuery extracts Korean date ranges, project code, attachment type, and folder hint', () => {
  const parsed = parseNaturalSearchQuery('11월 1일부터 11월 19일까지 PRJ-01 PDF 계약서 찾아줘', 2026);
  assert.equal(parsed.startDate, '2026-11-01');
  assert.equal(parsed.endDate, '2026-11-19');
  assert.equal(parsed.projectCode, 'PRJ-01');
  assert.equal(parsed.attachmentKind, 'pdf');
  assert.ok(parsed.keywords.includes('계약서'));

  const sentQuery = parseNaturalSearchQuery('2026-11-18 ~ 2026-11-19 보낸 엑셀 단가', 2026);
  assert.equal(sentQuery.folderHint, 'sent');
  assert.equal(sentQuery.attachmentKind, 'excel');
  assert.equal(sentQuery.startDate, '2026-11-18');
  assert.equal(sentQuery.endDate, '2026-11-19');
});

test('runDeterministicConversationalSearch accurately filters emails and documents with zero LLM', () => {
  const docs: ParsedDocumentSummary[] = [
    {
      id: 'd1',
      filename: 'PRJ-01_계약서.pdf',
      kind: 'pdf',
      mimeType: 'application/pdf',
      sizeBytes: 2048,
      engineUsed: '내장 PDF 엔진 (pdf-parse · MIT)',
      wordCount: 12,
      extractedText: 'PRJ-01 ERP 최종 계약서 조항 및 납품 일정',
      pageCount: 4,
      warnings: [],
    },
  ];

  const { hits } = runDeterministicConversationalSearch(
    '2026-11-15 ~ 2026-11-19 PRJ-01 PDF 계약서',
    EMAILS,
    docs,
    PROJECTS
  );
  assert.equal(hits.length, 2);
  assert.ok(hits.some((h) => h.type === 'email' && h.id === 'm1'));
  assert.ok(hits.some((h) => h.type === 'document' && h.id === 'd1'));
});

test('extractDeterministicSchedules extracts dates from emails and meeting action items', () => {
  const minutes: MeetingMinutesResult[] = [
    {
      id: 'min-1',
      createdAt: '2026-11-19T10:00:00Z',
      providerUsed: 'builtin',
      providerModel: 'gemini-4-hydrogen',
      transcriptionEngine: 'gemini-3.5-transcribe',
      title: 'ERP 주간 점검 회의',
      meetingDate: '2026-11-19',
      participants: ['김민수', '박지훈'],
      executiveSummary: '계약서 최종 승인',
      agendaTopics: [{ topic: '일정 확인', discussion: '납품 일정 점검', conclusion: '11월 22일 확정' }],
      decisions: ['11월 22일 배포 확정'],
      actionItems: [
        { owner: '김민수', task: '계약서 날인 송부', deadline: '2026-11-21', priority: '높음' },
      ],
      risksAndFollowups: [],
      linkedEmailInsights: [],
      sourceFiles: [],
      transcript: '...',
      markdownMinutes: '# 회의록',
    },
  ];

  const events = extractDeterministicSchedules(EMAILS, PROJECTS, minutes);
  assert.ok(events.some((e) => e.sourceEmailId === 'm1' && e.date === '2026-11-20' && e.timeLabel === '15:00'));
  assert.ok(events.some((e) => e.category === 'action-item' && e.date === '2026-11-21'));
});

test('buildDeterministicWorkLog produces structured work log without LLM', () => {
  const report = buildDeterministicWorkLog({
    emails: EMAILS,
    projects: PROJECTS,
    periodLabel: '2026-11-18 ~ 2026-11-19',
    templateName: 'project-case',
  });
  assert.equal(report.analyzedCount, 2);
  assert.equal(report.receivedCount, 1);
  assert.equal(report.sentCount, 1);
  assert.ok(report.workPerformed.some((w) => w.title.includes('[PRJ-01]')));
  assert.ok(report.workLog.completedTasks.length >= 1);
});

test('exportEmailsToMbox generates valid RFC-4155 mbox and escapes From lines', () => {
  const mbox = exportEmailsToMbox(EMAILS);
  assert.match(mbox, /^From minsu@partner\.co\.kr /);
  assert.ok(mbox.includes('>From 영업팀 드림'));
  assert.ok(mbox.includes('Subject: [PRJ-01] 발주 단가 엑셀 회신'));
});

test('extractDeterministicSchedules converts Korean 오후/오전 and explicit YYYY년 accurately', () => {
  const pmMail: MailItem = {
    ...EMAILS[0],
    id: 'm-pm',
    date: '2026-11-19',
    subject: '[PRJ-01] 킥오프 미팅 일정 안내',
    body: '2027년 1월 5일 오후 2시 30분 킥오프 미팅 진행합니다.',
  };
  const midnightMail: MailItem = {
    ...EMAILS[0],
    id: 'm-midnight',
    date: '2026-11-19',
    subject: '[PRJ-01] 자정 배포 점검 회의',
    body: '2026년 11월 26일 오전 12시 배포 점검 회의 예정입니다.',
  };
  const events = extractDeterministicSchedules([pmMail, midnightMail], PROJECTS);
  assert.ok(events.some((e) => e.sourceEmailId === 'm-pm' && e.date === '2027-01-05' && e.timeLabel === '14:30'));
  assert.ok(events.some((e) => e.sourceEmailId === 'm-midnight' && e.date === '2026-11-26' && e.timeLabel === '00:00'));
});

test('exportEmailsToMbox handles partial or malformed mail payloads defensively without throwing', () => {
  const partial = [{ id: 'p1', subject: '부분 데이터', body: 'Hello' }] as unknown as MailItem[];
  const mbox = exportEmailsToMbox(partial);
  assert.ok(mbox.includes('From unknown@local '));
  assert.ok(mbox.includes('Subject: 부분 데이터'));
  assert.ok(mbox.includes('To: me@workspace.kr'));
});

test('parseEmlMessage extracts multipart attachment metadata (PDF/Excel)', () => {
  const rawEml = [
    'From: MinSu <minsu@partner.co.kr>',
    'To: me@company.kr',
    'Subject: [PRJ-01] Attachment Test',
    'Date: Wed, 18 Nov 2026 10:00:00 +0900',
    'Content-Type: multipart/mixed; boundary="BND123"',
    '',
    '--BND123',
    'Content-Type: text/plain; charset="utf-8"',
    '',
    'Please check attached quote.',
    '--BND123',
    'Content-Type: application/pdf; name="quote_2026.pdf"',
    'Content-Disposition: attachment; filename="quote_2026.pdf"',
    'Content-Transfer-Encoding: base64',
    '',
    'JVBERi0xLjQK',
    '--BND123--',
  ].join('\r\n');

  const parsed = parseEmlMessage(rawEml, { idPrefix: 'test-eml' });
  assert.equal(parsed.hasAttachment, true);
  assert.equal(parsed.attachments?.length, 1);
  assert.equal(parsed.attachments?.[0].filename, 'quote_2026.pdf');
  assert.equal(parsed.attachments?.[0].kind, 'pdf');
});

test('sanitizeAppConfig preserves externalEngines folder paths while stripping secrets', () => {
  const sanitized = sanitizeAppConfig({
    externalEngines: {
      pdfEngineDir: '/opt/my-pdf-engine',
      excelEngineDir: '/opt/my-excel-engine',
      apiKey: 'SHOULD_BE_REMOVED',
      fallbackToBuiltIn: true,
    },
  });
  assert.equal(sanitized.externalEngines.pdfEngineDir, '/opt/my-pdf-engine');
  assert.equal(sanitized.externalEngines.excelEngineDir, '/opt/my-excel-engine');
  assert.equal('apiKey' in (sanitized.externalEngines as unknown as Record<string, unknown>), false);

  const statuses = buildEnvSecretStatuses({ EXTERNAL_PDF_ENGINE_DIR: '/opt/my-pdf-engine' });
  const pdfStatus = statuses.find((s) => s.key === 'EXTERNAL_PDF_ENGINE_DIR');
  assert.equal(pdfStatus?.configured, true);
});

test('parseUploadedDocument parses real .xlsx via exceljs with numeric summary highlights', async () => {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('매출집계');
  sheet.addRow(['항목', '금액']);
  sheet.addRow(['1차 납품', 1000000]);
  sheet.addRow(['2차 납품', 2500000]);

  const tmpPath = path.join(os.tmpdir(), `maillog-test-${Date.now()}.xlsx`);
  await wb.xlsx.writeFile(tmpPath);

  try {
    const parsed = await parseUploadedDocument(
      {
        path: tmpPath,
        originalname: 'sales_2026.xlsx',
        mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
      {
        pdfEngineDir: '',
        pdfEngineCommand: '',
        excelEngineDir: '',
        excelEngineCommand: '',
        fallbackToBuiltIn: true,
      }
    );

    assert.equal(parsed.kind, 'excel');
    assert.ok(parsed.engineUsed.includes('exceljs'));
    assert.deepEqual(parsed.sheetNames, ['매출집계']);
    assert.equal(parsed.rowCount, 3);
    assert.ok((parsed.numericHighlights ?? []).some((h) => h.includes('3,500,000')));
  } finally {
    fs.rmSync(tmpPath, { force: true });
  }
});

test('probeExternalEngine safely reports missing directory without throwing', async () => {
  const probe = await probeExternalEngine('pdf', '/nonexistent/maillog-pdf-engine-99999');
  assert.equal(probe.configuredDir, '/nonexistent/maillog-pdf-engine-99999');
  assert.equal(probe.exists, false);
});
