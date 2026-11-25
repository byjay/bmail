import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { LocalMailDb, type LocalEmailRecord } from './localMailDb';
import {
  buildAttachmentExportPlan,
  classifyAttachmentInline,
  executeAttachmentExportPlan,
  validateExportDestDir,
} from './attachmentBatchExporter';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';

function makeTempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-att-export-test-'));
}

const TEST_PROJECTS: ProjectDefinition[] = [
  {
    id: 'p-22k',
    code: '22K-FD',
    name: '22K Floating Dock',
    keywords: ['22K', 'CAISSON', 'KYTS-1037'],
    color: '#0F6CBD',
    isFavorite: true,
    createdAt: '2026-11-20T00:00:00Z',
  },
];

test('classifyAttachmentInline excludes signature images cleanly while keeping real documents', () => {
  assert.equal(
    classifyAttachmentInline({
      id: 'a0',
      filename: 'image001.png',
      mimeType: 'image/png',
      sizeBytes: 12400,
      kind: 'other',
    }),
    true
  );
  assert.equal(
    classifyAttachmentInline({
      id: 'a0b',
      filename: 'image.png',
      mimeType: 'image/png',
      sizeBytes: 25600,
      kind: 'other',
    }),
    true
  );
  assert.equal(
    classifyAttachmentInline(
      {
        id: 'a1',
        filename: 'logo.jpg',
        mimeType: 'image/jpeg',
        sizeBytes: 85000,
        kind: 'other',
        contentId: '<logo@sig>',
      },
      '<p>Hi <img src="cid:logo@sig"></p>'
    ),
    true
  );
  assert.equal(
    classifyAttachmentInline({
      id: 'a2',
      filename: '22K_POS_REV_A.pdf',
      mimeType: 'application/pdf',
      sizeBytes: 512000,
      kind: 'pdf',
    }),
    false
  );
  assert.equal(
    classifyAttachmentInline({
      id: 'a3',
      filename: 'H042_DRAWING.dwg',
      mimeType: 'application/acad',
      sizeBytes: 240000,
      kind: 'other',
    }),
    false
  );
});

test('validateExportDestDir blocks unsafe or system paths and accepts valid export directories', () => {
  assert.throws(() => validateExportDestDir(''));
  assert.throws(() => validateExportDestDir('C:\\Windows\\System32'));
  assert.throws(() => validateExportDestDir('C:\\temp\\..\\escape'));
  assert.throws(() => validateExportDestDir('D:\\'));
  const tempTarget = path.join(os.tmpdir(), 'bmail-safe-export-check');
  assert.equal(validateExportDestDir(tempTarget), path.resolve(tempTarget));
});

test('buildAttachmentExportPlan & executeAttachmentExportPlan filter, deduplicate, save files, and write MD/CSV/HTML reports', (t) => {
  const dir = makeTempDir();
  const dbDir = path.join(dir, 'db');
  const exportDir = path.join(dir, 'exports');
  const db = new LocalMailDb(dbDir);

  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable: ${db.getInitError() || 'unknown'}`);
      return;
    }

    const fakePst = path.join(dir, 'sample.pst');
    fs.writeFileSync(fakePst, 'pst-placeholder');

    const records: LocalEmailRecord[] = [
      {
        id: 'exp-m1',
        storeName: 'StoreA',
        pstPath: fakePst,
        descriptorNodeId: 101,
        folder: 'inbox',
        folderName: '받은 편지함 / 22K FD',
        subject: '[KYTS-1037] 22K FD POS 최종 송부',
        fromName: '홍엔지니어',
        fromEmail: 'hong@partner.example',
        to: ['me@seastar.example'],
        date: '2026-11-18',
        timestamp: '2026-11-18T10:00:00Z',
        timeLabel: '10:00',
        snippet: 'POS 첨부',
        body: '22K Floating Dock POS 문서 송부합니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'exp-m1-att-0',
            filename: 'image001.png',
            mimeType: 'image/png',
            sizeBytes: 9400,
            kind: 'other',
          },
          {
            id: 'exp-m1-att-1',
            filename: 'KYTS-1037_POS.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 12,
            kind: 'pdf',
          },
        ],
        importance: 'normal',
        category: '받은 편지함 / 22K FD',
        isRead: true,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
      {
        id: 'exp-m2',
        storeName: 'StoreA',
        pstPath: fakePst,
        descriptorNodeId: 102,
        folder: 'inbox',
        folderName: '받은 편지함 / 일반',
        subject: '주간 단가표 업데이트',
        fromName: '박담당',
        fromEmail: 'park@vendor.example',
        to: ['me@seastar.example'],
        date: '2026-11-10',
        timestamp: '2026-11-10T09:00:00Z',
        timeLabel: '09:00',
        snippet: '엑셀 첨부',
        body: '단가표 송부합니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'exp-m2-att-0',
            filename: 'BOM_Price.xlsx',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            sizeBytes: 15,
            kind: 'excel',
          },
        ],
        importance: 'normal',
        category: '받은 편지함 / 일반',
        isRead: true,
        isFlagged: false,
        isDeleted: false,
        source: 'pst-import',
      },
    ];
    db.upsertEmailsBatch(records);

    const plan = buildAttachmentExportPlan(
      db,
      {
        destDir: exportDir,
        scope: 'date-range',
        dateFrom: '2026-11-15',
        dateTo: '2026-11-23',
        senderQuery: 'hong',
        excludeInlineImages: true,
        folderStructure: 'by-project',
        filenamePattern: 'date-sender-original',
      },
      TEST_PROJECTS
    );

    assert.equal(plan.matchedEmails, 1);
    assert.equal(plan.matchedFiles, 1);
    assert.equal(plan.skippedInlineCount, 1);
    assert.equal(plan.items[0].filename, 'KYTS-1037_POS.pdf');
    assert.equal(plan.items[0].projectCode, '22K-FD');
    assert.equal(plan.items[0].attachmentIndex, 1);
    assert.ok(plan.items[0].targetRelativePath.startsWith('22K-FD'));

    const mockReader = (_p: string, _node: number, _idx: number) => ({
      filename: 'KYTS-1037_POS.pdf',
      mimeType: 'application/pdf',
      buffer: Buffer.from('PDF-BYTE-OK!'),
    });

    const firstRun = executeAttachmentExportPlan(
      db,
      {
        destDir: exportDir,
        scope: 'date-range',
        dateFrom: '2026-11-15',
        dateTo: '2026-11-23',
        senderQuery: 'hong',
        excludeInlineImages: true,
        folderStructure: 'by-project',
        filenamePattern: 'date-sender-original',
        conflictPolicy: 'skip-identical',
        reportFormats: ['md', 'csv', 'html'],
      },
      TEST_PROJECTS,
      undefined,
      mockReader
    );

    assert.equal(firstRun.exportedCount, 1);
    assert.equal(firstRun.skippedIdenticalCount, 0);
    assert.equal(firstRun.failedCount, 0);
    assert.equal(firstRun.reportPaths.length, 3);
    assert.ok(fs.existsSync(firstRun.sampleExportedFiles[0]));
    for (const rp of firstRun.reportPaths) {
      assert.ok(fs.existsSync(rp.path));
    }
    const csvContent = fs.readFileSync(firstRun.reportPaths.find((r) => r.format === 'csv')!.path, 'utf8');
    assert.ok(csvContent.startsWith('\uFEFF'));

    // Re-running with skip-identical skips already-saved identical file
    const secondRun = executeAttachmentExportPlan(
      db,
      {
        destDir: exportDir,
        scope: 'date-range',
        dateFrom: '2026-11-15',
        dateTo: '2026-11-23',
        senderQuery: 'hong',
        excludeInlineImages: true,
        folderStructure: 'by-project',
        filenamePattern: 'date-sender-original',
        conflictPolicy: 'skip-identical',
      },
      TEST_PROJECTS,
      undefined,
      mockReader
    );
    assert.equal(secondRun.exportedCount, 0);
    assert.equal(secondRun.skippedIdenticalCount, 1);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
