import assert from 'node:assert/strict';
import test from 'node:test';
import type { MailItem } from '../types/mail';
import { buildDeterministicWorkLog, type ProjectDefinition } from './deterministicAnalyzer';
import { formatWorkLogMarkdown } from './mailFilter';
import {
  buildInheritedOutlookProfile,
  isSystemOutlookFolder,
  mergeInheritedProjects,
} from './pstInheritance';

const SAMPLE_PST_EMAILS: MailItem[] = [
  {
    id: 'm1',
    folder: 'sent',
    folderName: '보낸 편지함',
    subject: '[PRJ-2026] 차세대 ERP 1차 견적서 제출',
    fromName: '홍길동',
    fromEmail: 'gildong@company.kr',
    to: ['client@partner.co.kr'],
    date: '2026-11-18',
    timeLabel: '16:10',
    timestamp: '2026-11-18T16:10:00Z',
    snippet: '1차 견적서 PDF 첨부하여 회신드립니다.',
    body: '요청하신 ERP 1차 견적서 검토 완료하여 제출합니다. 11월 20일 14:00 회의에서 최종 확정 부탁드립니다.',
    hasAttachment: true,
    attachments: [
      {
        id: 'a1',
        filename: 'ERP_견적서_v1.pdf',
        mimeType: 'application/pdf',
        kind: 'pdf',
        sizeBytes: 34000,
      },
    ],
    isRead: true,
    source: 'pst-import',
  },
  {
    id: 'm2',
    folder: 'inbox',
    folderName: '받은 편지함 / 스마트팩토리',
    subject: 'RE: 스마트팩토리 센서 발주 단가 검토 요청',
    fromName: '김과장',
    fromEmail: 'kim@vendor.co.kr',
    to: ['gildong@company.kr'],
    date: '2026-11-19',
    timeLabel: '09:30',
    timestamp: '2026-11-19T09:30:00Z',
    snippet: '엑셀 단가표 확인 부탁드립니다.',
    body: '11월 22일 마감 전까지 센서 발주 단가표 확인 및 승인 부탁드립니다.',
    hasAttachment: true,
    attachments: [
      {
        id: 'a2',
        filename: '센서_단가표.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        kind: 'excel',
        sizeBytes: 18500,
      },
    ],
    isRead: false,
    source: 'pst-import',
  },
];

test('isSystemOutlookFolder distinguishes system folders from custom project folders', () => {
  assert.equal(isSystemOutlookFolder('받은 편지함'), true);
  assert.equal(isSystemOutlookFolder('Inbox'), true);
  assert.equal(isSystemOutlookFolder('보낸 편지함'), true);
  assert.equal(isSystemOutlookFolder('받은 편지함 / 스마트팩토리'), false);
  assert.equal(isSystemOutlookFolder('PRJ-2026 ERP'), false);
});

test('buildInheritedOutlookProfile infers owner account, custom folder projects, subject codes, and schedule hints', () => {
  const profile = buildInheritedOutlookProfile({
    storeName: 'company_archive.pst',
    folders: ['보낸 편지함', '받은 편지함 / 스마트팩토리'],
    emails: SAMPLE_PST_EMAILS,
  });

  assert.equal(profile.ownerEmail, 'gildong@company.kr');
  assert.equal(profile.ownerName, '홍길동');
  assert.equal(profile.dateRange.earliest, '2026-11-18');
  assert.equal(profile.dateRange.latest, '2026-11-19');
  assert.equal(profile.attachmentMailCount, 2);
  assert.ok(profile.calendarHintCount >= 2);
  assert.ok(profile.discoveredProjects.some((p) => p.code === 'PRJ-2026'));
  assert.ok(profile.discoveredProjects.some((p) => p.name === '스마트팩토리' && p.source === 'folder'));
});

test('buildInheritedOutlookProfile handles empty input defensively without throwing', () => {
  const empty = buildInheritedOutlookProfile({ storeName: '', folders: [], emails: [] });
  assert.equal(empty.storeName, 'Outlook PST');
  assert.equal(empty.ownerEmail, '');
  assert.equal(empty.discoveredProjects.length, 0);
  assert.equal(empty.calendarHintCount, 0);
});

test('mergeInheritedProjects prevents duplicate project codes and preserves existing manual entries', () => {
  const existing: ProjectDefinition[] = [
    {
      id: 'manual-1',
      code: 'PRJ-2026',
      name: '수동 등록 ERP',
      keywords: ['ERP'],
      color: '#0F6CBD',
      isFavorite: true,
      createdAt: '2026-11-01T00:00:00Z',
    },
  ];
  const profile = buildInheritedOutlookProfile({
    storeName: 'company_archive.pst',
    folders: ['보낸 편지함', '받은 편지함 / 스마트팩토리'],
    emails: SAMPLE_PST_EMAILS,
  });

  const { merged, addedCount } = mergeInheritedProjects(existing, profile.discoveredProjects);
  assert.equal(addedCount, 1);
  assert.equal(merged.length, 2);
  assert.ok(merged.some((p) => p.id === 'manual-1' && p.name === '수동 등록 ERP'));
});

test('buildDeterministicWorkLog and formatWorkLogMarkdown include projectSummaries section', () => {
  const report = buildDeterministicWorkLog({
    emails: SAMPLE_PST_EMAILS,
    periodLabel: '2026-11-18 ~ 2026-11-19',
  });
  assert.ok(Array.isArray(report.projectSummaries) && report.projectSummaries.length >= 1);
  const prj2026 = report.projectSummaries?.find((p) => p.projectCode === 'PRJ-2026');
  assert.ok(prj2026);
  assert.ok(prj2026.myActions.some((a) => a.includes('견적서 제출')));
  assert.ok(prj2026.attachmentsMentioned.includes('ERP_견적서_v1.pdf'));

  const md = formatWorkLogMarkdown(report);
  assert.ok(md.includes('## 프로젝트별 전체 정리'));
  assert.ok(md.includes('[PRJ-2026]'));
  assert.ok(md.includes('ERP_견적서_v1.pdf'));
});
