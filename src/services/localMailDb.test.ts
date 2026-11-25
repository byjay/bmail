import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { runAutonomousMailAgent } from './agentService';
import { LocalMailDb, type LocalEmailRecord } from './localMailDb';
import { inspectPstFileStates, recomputeProjectsFromLocalDb } from './localOutlookMigrator';
import { analyzeMeetingMinutesWithAi, buildDeterministicDocumentAnalysis } from './multimodalAiService';
import { DEFAULT_APP_CONFIG } from '../utils/appConfig';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-localdb-test-'));
}

test('LocalMailDb: upserts emails idempotently, searches Korean/English via FTS5, and patches flags', (t) => {
  const dir = makeTempDir();
  const db = new LocalMailDb(dir);

  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable in current Node runtime: ${db.getInitError() || 'unknown'}`);
      return;
    }

    const sampleRecords: LocalEmailRecord[] = [
      {
        id: 'pst-test-1',
        storeName: 'test-store',
        pstPath: path.join(dir, 'sample.pst'),
        descriptorNodeId: 1001,
        folder: 'inbox',
        folderName: '받은 편지함 / 22K FD',
        subject: '[22K FD] 선급 승인도면 제출 요청',
        fromName: '홍길동',
        fromEmail: 'sender01@example.test',
        to: ['user01@example.test'],
        date: '2026-11-20',
        timestamp: '2026-11-20T09:15:00.000Z',
        timeLabel: '09:15',
        snippet: '도면 번호 DWG-2201 확인 부탁드립니다.',
        body: '안녕하세요, 선급 승인도면 DWG-2201 개정본 검토 요청드립니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'pst-test-1-att-0',
            filename: 'DWG-2201.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 20480,
            kind: 'pdf',
          },
        ],
        importance: 'high',
        category: '받은 편지함 / 22K FD',
        isRead: false,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
      {
        id: 'pst-test-2',
        storeName: 'test-store',
        pstPath: path.join(dir, 'sample.pst'),
        descriptorNodeId: 1002,
        folder: 'sent',
        folderName: '보낸 편지함',
        subject: 'Loadout beam structural calculation reply',
        fromName: '김엔지니어',
        fromEmail: 'user01@example.test',
        to: ['partner02@example.test'],
        date: '2026-11-21',
        timestamp: '2026-11-21T03:00:00.000Z',
        timeLabel: '03:00',
        snippet: 'Attached updated REV-B report.',
        body: 'Please find attached the updated Loadout beam REV-B structural calculation.',
        hasAttachment: false,
        importance: 'normal',
        category: '보낸 편지함',
        isRead: true,
        isFlagged: true,
        isDeleted: false,
        source: 'pst-import',
      },
    ];

    const inserted = db.upsertEmailsBatch(sampleRecords);
    assert.equal(inserted, 2);

    // Re-upserting the same IDs must remain strictly idempotent (no duplicate rows)
    db.upsertEmailsBatch(sampleRecords);
    const allRes = db.queryEmails({ limit: 50 });
    assert.equal(allRes.total, 2);
    assert.equal(allRes.items[0].id, 'pst-test-2');

    // FTS5 Korean + English search
    const koRes = db.queryEmails({ q: '선급 승인도면' });
    assert.equal(koRes.total, 1);
    assert.equal(koRes.items[0].id, 'pst-test-1');

    const enRes = db.queryEmails({ q: 'Loadout REV-B' });
    assert.equal(enRes.total, 1);
    assert.equal(enRes.items[0].id, 'pst-test-2');

    // Patch state (read + flag + soft-delete)
    assert.equal(db.patchEmailState('pst-test-1', { isRead: true, isFlagged: true }), true);
    const detail = db.getEmailById('pst-test-1');
    assert.equal(detail?.isRead, true);
    assert.equal(detail?.isFlagged, true);
    assert.equal(detail?.descriptorNodeId, 1001);

    assert.equal(db.patchEmailState('pst-test-1', { isDeleted: true }), true);
    assert.equal(db.queryEmails({ limit: 50 }).total, 1);
    assert.equal(db.queryEmails({ deletedOnly: true, limit: 50 }).total, 1);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('LocalMailDb: persists accounts, signatures, contacts, user state allowlist, and PST path checks', (t) => {
  const dir = makeTempDir();
  const db = new LocalMailDb(dir);

  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable in current Node runtime: ${db.getInitError() || 'unknown'}`);
      return;
    }

    const allowedPst = path.join(dir, 'archive.pst');
    db.setState('pstPaths', [allowedPst]);
    assert.equal(db.isAllowedPstPath(allowedPst), true);
    assert.equal(db.isAllowedPstPath(path.join(dir, '..', 'outside.pst')), false);

    db.upsertAccounts([
      {
        email: 'user01@example.test',
        displayName: 'Test User',
        providerPreset: 'mailplug',
        protocol: 'pop3',
        pop3Host: 'pop3.mailplug.co.kr',
        pop3Port: 995,
        pop3Tls: true,
        imapHost: '',
        imapPort: 993,
        imapTls: true,
        smtpHost: 'smtp.mailplug.co.kr',
        smtpPort: 465,
        smtpTls: true,
        leaveOnServer: true,
        deliveryStorePath: allowedPst,
      },
    ]);

    db.upsertSignatures([
      {
        name: '유일',
        htmlContent: '<p>유일 서명</p>',
        textContent: '유일 서명',
        isDefaultNew: true,
        isDefaultReply: true,
      },
    ]);

    db.upsertContacts([
      {
        id: 'c-1',
        fullName: '담당자',
        email: 'partner01@example.test',
        company: '테스트조선',
        source: 'outlook-pc',
      },
    ]);

    db.setUserState('outlookPrefs', { general: { language: 'ko-KR' } });
    assert.throws(() => db.setUserState('arbitrarySecretKey', { bad: true }));

    const bootstrap = db.getBootstrapPayload(10);
    assert.equal(bootstrap.available, true);
    assert.equal(bootstrap.accounts.length, 1);
    assert.equal(bootstrap.accounts[0].providerPreset, 'mailplug');
    assert.equal(bootstrap.activeSignatureName, '유일');
    assert.equal(bootstrap.totalContacts, 1);
    assert.equal(bootstrap.pstPaths.length, 1);
    assert.deepEqual(bootstrap.kvState.outlookPrefs, { general: { language: 'ko-KR' } });
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('recomputeProjectsFromLocalDb: discovers multiple non-system folders from SQLite folders table', (t) => {
  const dir = makeTempDir();
  const db = new LocalMailDb(dir);

  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable in current Node runtime: ${db.getInitError() || 'unknown'}`);
      return;
    }

    db.upsertFolders([
      {
        storeName: 'ArchiveStore',
        folderPath: '받은 편지함',
        classifiedFolder: 'inbox',
        itemCount: 1500,
        isSystem: true,
      },
      {
        storeName: 'ArchiveStore',
        folderPath: '받은 편지함 / SN2201',
        classifiedFolder: 'inbox',
        itemCount: 420,
        isSystem: false,
      },
      {
        storeName: 'ArchiveStore',
        folderPath: '받은 편지함 / 22K FD',
        classifiedFolder: 'inbox',
        itemCount: 310,
        isSystem: false,
      },
      {
        storeName: 'ArchiveStore',
        folderPath: '받은 편지함 / 프로젝트 Alpha-09',
        classifiedFolder: 'inbox',
        itemCount: 185,
        isSystem: false,
      },
    ]);

    const projects = recomputeProjectsFromLocalDb(db);
    assert.ok(projects.length >= 3, `expected >= 3 projects, got ${projects.length}`);
    const codes = new Set(projects.map((p) => p.code));
    assert.ok(codes.has('SN2201'));
    assert.ok(codes.has('ALPHA-09'));

    const persisted = db.getState<typeof projects>('projects') || [];
    assert.equal(persisted.length, projects.length);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('buildDeterministicDocumentAnalysis and inspectPstFileStates work 100% offline with zero API keys', () => {
  const dir = makeTempDir();
  const sampleFile = path.join(dir, 'sample.pst');
  fs.writeFileSync(sampleFile, 'PST_HEADER_STUB', 'utf8');

  try {
    const states = inspectPstFileStates([sampleFile, path.join(dir, 'missing.pst')]);
    assert.equal(states.length, 1);
    assert.equal(states[0].path, path.resolve(sampleFile));
    assert.ok(states[0].size > 0);

    const result = buildDeterministicDocumentAnalysis({
      documents: [
        {
          id: 'doc-1',
          filename: 'DWG-2201-REV2-Schedule.pdf',
          kind: 'pdf',
          mimeType: 'application/pdf',
          sizeBytes: 40960,
          engineUsed: '내장 PDF 엔진 (pdf-parse · MIT)',
          pageCount: 4,
          wordCount: 320,
          extractedText: 'SN2201 Hull Structural Drawing DWG-2201 REV-2 urgent approval schedule.',
          warnings: [],
        },
        {
          id: 'doc-2',
          filename: 'BOM-Estimate.xlsx',
          kind: 'excel',
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          sizeBytes: 18432,
          engineUsed: '내장 Excel 엔진 (exceljs · MIT)',
          sheetNames: ['Summary', 'Items'],
          rowCount: 45,
          wordCount: 180,
          numericHighlights: ['합계 금액: 125000'],
          extractedText: 'Summary total items quantity cost breakdown.',
          warnings: [],
        },
      ],
      emails: [
        {
          id: 'mail-1',
          folder: 'inbox',
          subject: '[SN2201] DWG-2201 승인도면 검토 요청',
          fromName: '김선급',
          fromEmail: 'class@example.test',
          to: ['design@example.test'],
          date: '2026-11-20',
          timestamp: '2026-11-20T08:00:00.000Z',
          timeLabel: '08:00',
          snippet: 'DWG-2201 검토 부탁드립니다.',
          body: '첨부 DWG-2201 확인 바랍니다.',
          hasAttachment: true,
          isRead: false,
          isFlagged: false,
          source: 'pst-import',
        },
      ],
      customFocus: '도면 리비전 및 수치 비교',
      maskSensitiveInfo: true,
    });

    assert.equal(result.providerUsed, 'builtin');
    assert.equal(result.keyFindings.length, 2);
    assert.ok(result.overview.includes('100% 로컬'));
    assert.ok(result.emailCrossInsights.length >= 1);
    assert.ok(result.actionItems.length >= 1);
    assert.ok(result.markdownReport.includes('Outlook 메일 교차 분석'));
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('runAutonomousMailAgent falls back offline via buildDeterministicWorkLog when cloud keys are empty', async () => {
  const prevKeys = {
    GEMINI_API_KEY: process.env.GEMINI_API_KEY,
    GROQ_API_KEY: process.env.GROQ_API_KEY,
    ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY,
    OPENAI_API_KEY: process.env.OPENAI_API_KEY,
  };
  delete process.env.GEMINI_API_KEY;
  delete process.env.GROQ_API_KEY;
  delete process.env.ANTHROPIC_API_KEY;
  delete process.env.OPENAI_API_KEY;

  try {
    const result = await runAutonomousMailAgent({
      emails: [
        {
          id: 'offline-mail-1',
          folder: 'inbox',
          subject: '[SN2201] DWG-2201 선체 구조도 검토 요청',
          fromName: '김선급',
          fromEmail: 'class@example.test',
          to: ['design@example.test'],
          date: '2026-11-22',
          timestamp: '2026-11-22T09:00:00.000Z',
          timeLabel: '09:00',
          snippet: '긴급 검토 및 회신 요청드립니다.',
          body: '첨부 DWG-2201 검토 후 2026-11-25까지 회신 부탁드립니다. 연락처 010-1234-5678',
          hasAttachment: true,
          isRead: false,
          isFlagged: false,
          source: 'pst-import',
        },
        {
          id: 'offline-mail-2',
          folder: 'sent',
          subject: 'RE: [SN2201] DWG-2201 검토 결과 회신',
          fromName: '설계담당',
          fromEmail: 'design@example.test',
          to: ['class@example.test'],
          date: '2026-11-22',
          timestamp: '2026-11-22T10:30:00.000Z',
          timeLabel: '10:30',
          snippet: '1차 수정 반영 완료했습니다.',
          body: '수정 도면 확인 바랍니다.',
          hasAttachment: false,
          isRead: true,
          isFlagged: false,
          source: 'pst-import',
        },
      ],
      options: {
        dateMode: 'single',
        singleDate: '2026-11-22',
        startDate: '2026-11-22',
        endDate: '2026-11-22',
        scope: 'all',
        selectedIds: [],
        provider: 'builtin',
        blogTone: 'insight',
        maskSensitiveInfo: true,
        includeActionItems: true,
        customFocus: '',
      },
      config: {
        ...DEFAULT_APP_CONFIG,
        ai: { ...DEFAULT_APP_CONFIG.ai, fallbackToBuiltIn: true },
      },
      projects: [
        {
          id: 'proj-sn2201',
          code: 'SN2201',
          name: '2201호선 구조설계',
          keywords: ['SN2201', 'DWG-2201'],
          color: '#0F6CBD',
          isFavorite: true,
          createdAt: '2026-11-22T00:00:00.000Z',
        },
      ],
      envOverrides: {
        GEMINI_API_KEY: '',
        GROQ_API_KEY: '',
        ANTHROPIC_API_KEY: '',
        OPENAI_API_KEY: '',
      },
    });

    assert.equal(result.providerUsed, 'builtin');
    assert.ok(result.workPerformed.length > 0);
    assert.equal(result.agentSteps?.length, 4);
    assert.ok(result.agentSteps?.every((step) => step.status === 'completed'));
    assert.ok(!JSON.stringify(result).includes('010-1234-5678'));
  } finally {
    for (const [k, v] of Object.entries(prevKeys)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  }
});

test('analyzeMeetingMinutesWithAi falls back offline via buildDeterministicMeetingMinutes when cloud keys are empty', async () => {
  const prevKeys = {
    GEMINI_API_KEY: process.env.GEMINI_API_KEY,
    GROQ_API_KEY: process.env.GROQ_API_KEY,
    ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY,
    OPENAI_API_KEY: process.env.OPENAI_API_KEY,
  };
  delete process.env.GEMINI_API_KEY;
  delete process.env.GROQ_API_KEY;
  delete process.env.ANTHROPIC_API_KEY;
  delete process.env.OPENAI_API_KEY;

  try {
    const minutes = await analyzeMeetingMinutesWithAi({
      titleHint: 'SN2201 설계 공정 점검 회의',
      meetingDate: '2026-11-22',
      transcript: '',
      notesText: [
        '참석자: 김설계, 이품질, 박공정',
        '1. SN2201 DWG-2201 리비전 2 일정 점검',
        '김설계: 리비전 2 출도 일정을 2026-11-26까지 확정하기로 합의함.',
        '2. 자재 수급 리스크 점검',
        '이품질: 긴급 자재 스펙 확인 필요 (담당: 박공정, 기한: 2026-11-24)',
      ].join('\n'),
      documents: [],
      emails: [],
      provider: 'builtin',
      config: {
        ...DEFAULT_APP_CONFIG,
        ai: { ...DEFAULT_APP_CONFIG.ai, fallbackToBuiltIn: true },
      },
      maskSensitiveInfo: true,
    });

    assert.equal(minutes.providerUsed, 'builtin');
    assert.ok(minutes.participants.includes('김설계'));
    assert.ok(minutes.decisions.length >= 1);
    assert.ok(minutes.actionItems.length >= 1);
    assert.ok(minutes.markdownMinutes.includes('확정 의사결정'));
  } finally {
    for (const [k, v] of Object.entries(prevKeys)) {
      if (v === undefined) delete process.env[k];
      else process.env[k] = v;
    }
  }
});
