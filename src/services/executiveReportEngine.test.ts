import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import {
  buildExecutiveDashboardReport,
  classifyEmailTopic,
  saveExecutiveReportSnapshot,
} from './executiveReportEngine';
import { LocalMailDb, type LocalEmailRecord } from './localMailDb';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-exec-report-test-'));
}

const TEST_PROJECTS: ProjectDefinition[] = [
  {
    id: 'p-22k',
    code: '22K-FD',
    name: '22K Floating Dock',
    keywords: ['22K', 'KYTS-1037'],
    color: '#0F6CBD',
    isFavorite: true,
    createdAt: '2026-11-01T00:00:00Z',
  },
  {
    id: 'p-h042',
    code: 'H042',
    name: 'Exmar H042 Project',
    keywords: ['H042', 'Exmar'],
    color: '#0B7A75',
    isFavorite: true,
    createdAt: '2026-11-01T00:00:00Z',
  },
];

test('classifyEmailTopic categorizes shipbuilding engineering and sales topics deterministically', () => {
  assert.equal(
    classifyEmailTopic({ subject: '[H042] 승인도 DWG-101 Rev.A 제출', body: '', folderName: '' }),
    '도면·승인도·REV'
  );
  assert.equal(
    classifyEmailTopic({ subject: 'KYTS-1037 수정 POS 및 견적 요청건', body: '', folderName: '' }),
    '견적·발주·POS'
  );
  assert.equal(
    classifyEmailTopic({ subject: '한국선급 승인 미결 항목 알림', body: '', folderName: '' }),
    '선급·검사·미결'
  );
  assert.equal(
    classifyEmailTopic({ subject: '3D Modeling 생산설계 검토', body: '', folderName: '' }),
    '3D 모델링·생산설계'
  );
});

test('buildExecutiveDashboardReport builds narrative report, weekly/monthly tables, unreplied incoming, project/topic tables, and persists snapshots', (t) => {
  const dir = makeTempDir();
  const db = new LocalMailDb(dir);

  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable: ${db.getInitError() || 'unknown'}`);
      return;
    }

    const records: LocalEmailRecord[] = [
      // Thread 1: Received and NOT replied -> should appear in unrepliedIncoming
      {
        id: 'exec-m1',
        storeName: '김봉정',
        folder: 'inbox',
        folderName: '받은 편지함 / 22K FD',
        subject: '[KYTS-1037] 22K FD 수정 POS 긴급 발행 요청',
        fromName: '최길용',
        fromEmail: 'choi@partner.example',
        to: ['me@seastar.example'],
        date: '2026-11-20',
        timestamp: '2026-11-20T05:28:00Z',
        timeLabel: '14:28',
        snippet: '수정 POS 발행을 요청드리오니 검토 바랍니다.',
        body: '초기 POS 대비 최종 적용 사양이 변경되어 수정 POS 발행 요청드립니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'exec-m1-att-0',
            filename: 'image001.png',
            mimeType: 'image/png',
            sizeBytes: 4200,
            kind: 'other',
          },
          {
            id: 'exec-m1-att-1',
            filename: 'KYTS-1037_POS_Deviation.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 340000,
            kind: 'pdf',
          },
        ],
        conversationKey: 'kyts-1037 22k fd 수정 pos 긴급 발행 요청',
        importance: 'high',
        category: '받은 편지함 / 22K FD',
        isRead: false,
        isFlagged: true,
        isDeleted: false,
        source: 'pst-import',
      },
      // Thread 2: Received on 11-18 AND replied on 11-19 -> should NOT appear in unrepliedIncoming
      {
        id: 'exec-m2',
        storeName: '김봉정',
        folder: 'inbox',
        folderName: '받은 편지함',
        subject: '[H042] Exmar 3D Modeling 자료 요청',
        fromName: '서하원',
        fromEmail: 'seo@vendor.example',
        to: ['me@seastar.example'],
        date: '2026-11-18',
        timestamp: '2026-11-18T02:00:00Z',
        timeLabel: '11:00',
        snippet: 'Modeling 자료 요청드립니다.',
        body: 'Exmar 3D Modeling 파일 송부 부탁드립니다.',
        hasAttachment: false,
        conversationKey: 'h042 exmar 3d modeling 자료 요청',
        importance: 'normal',
        category: '받은 편지함',
        isRead: true,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
      {
        id: 'exec-m3',
        storeName: '김봉정',
        folder: 'sent',
        folderName: '보낸 편지함',
        subject: 'RE: [H042] Exmar 3D Modeling 자료 요청',
        fromName: '김봉정',
        fromEmail: 'me@seastar.example',
        to: ['seo@vendor.example'],
        date: '2026-11-19',
        timestamp: '2026-11-19T04:00:00Z',
        timeLabel: '13:00',
        snippet: '요청하신 Modeling 파일 송부합니다.',
        body: '첨부 확인 바랍니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'exec-m3-att-0',
            filename: 'H042_3D_Model_Rev1.xlsx',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            sizeBytes: 128000,
            kind: 'excel',
          },
        ],
        conversationKey: 'h042 exmar 3d modeling 자료 요청',
        importance: 'normal',
        category: '보낸 편지함',
        isRead: true,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
      // Thread 3: Newsletter / no-reply -> classified as fyi, NOT unrepliedIncoming
      {
        id: 'exec-m4',
        storeName: '김봉정',
        folder: 'inbox',
        folderName: '받은 편지함',
        subject: '[무료 세미나] 수신거부 포함 뉴스레터 안내',
        fromName: '세미나센터',
        fromEmail: 'noreply@seminar.example',
        to: ['me@seastar.example'],
        date: '2026-11-16',
        timestamp: '2026-11-16T01:00:00Z',
        timeLabel: '10:00',
        snippet: '수신거부 링크 안내',
        body: '뉴스레터 본문입니다.',
        hasAttachment: false,
        conversationKey: '무료 세미나 수신거부 포함 뉴스레터 안내',
        importance: 'normal',
        category: '받은 편지함',
        isRead: true,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
    ];

    db.upsertEmailsBatch(records);

    const report = buildExecutiveDashboardReport(db, { periodType: 'all' }, TEST_PROJECTS);

    assert.equal(report.kpis.totalEmails, 4);
    assert.equal(report.kpis.receivedCount, 3);
    assert.equal(report.kpis.sentCount, 1);
    // Only exec-m1 is unreplied incoming (exec-m2 was replied by exec-m3, exec-m4 is noreply/newsletter)
    assert.equal(report.kpis.unrepliedIncomingCount, 1);
    assert.equal(report.unrepliedIncoming[0].emailId, 'exec-m1');
    assert.equal(report.unrepliedIncoming[0].projectCode, '22K-FD');
    // Inline image001.png excluded, only 2 real attachments (PDF + XLSX)
    assert.equal(report.kpis.realAttachmentCount, 2);
    assert.equal(report.attachmentRows.length, 2);
    assert.ok(report.narrativeReport.executiveSummary.includes('4건'));
    assert.ok(report.narrativeReport.markdownFull.includes('내게 온 미회신건'));
    assert.ok(report.weeklyTable.length >= 1);
    assert.ok(report.monthlyTable.length >= 1);
    assert.ok(report.projectRows.some((p) => p.projectCode === '22K-FD' && p.unrepliedIncoming === 1));
    assert.ok(report.projectRows.some((p) => p.projectCode === 'H042' && p.total === 2 && p.unrepliedIncoming === 0));

    // Filter by project H042
    const h042Only = buildExecutiveDashboardReport(
      db,
      { periodType: 'all', projectCode: 'H042' },
      TEST_PROJECTS
    );
    assert.equal(h042Only.kpis.totalEmails, 2);
    assert.equal(h042Only.kpis.unrepliedIncomingCount, 0);

    // Save snapshot into SQLite kv_state
    const snapshots = saveExecutiveReportSnapshot(db, report, '11월 3주차 주간 보고서');
    assert.equal(snapshots.length, 1);
    assert.equal(snapshots[0].title, '11월 3주차 주간 보고서');
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
