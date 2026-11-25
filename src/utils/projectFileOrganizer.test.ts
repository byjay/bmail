import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import type { ProjectDefinition } from './deterministicAnalyzer';
import {
  buildFileOrganizationReport,
  buildFolderOrganizationTree,
  classifyByProjectRoutingRules,
  collectOrganizedFiles,
  discoverRoutingSuggestions,
  extractEmailDomain,
  extractNormalizedEmail,
  formatOrganizationReportCsv,
  formatOrganizationReportMarkdown,
  generateFolderOrganizerScript,
  normalizeRuleList,
  sanitizeFolderSegment,
} from './projectFileOrganizer';

const PROJECTS: ProjectDefinition[] = [
  {
    id: 'p1',
    code: 'PRJ-H101',
    name: '선체 101호기 배관공사',
    keywords: ['선체', '배관'],
    color: '#0F6CBD',
    isFavorite: true,
    createdAt: '2026-11-01T00:00:00Z',
    senderEmails: ['vip.buyer@alpha-yard.co.kr'],
    domains: ['alpha-supplier.com'],
    accounts: ['sales-h101@mycompany.co.kr'],
    defaultSubfolder: '01_수신자료',
  },
  {
    id: 'p2',
    code: 'PRJ-E202',
    name: '전장 케이블 포설',
    keywords: ['전장', '케이블'],
    color: '#107C10',
    isFavorite: false,
    createdAt: '2026-11-02T00:00:00Z',
    senderEmails: ['eng@beta-electric.com'],
    domains: ['beta-electric.com'],
    accounts: ['elec-team@mycompany.co.kr'],
  },
];

const SAMPLE_EMAILS: MailItem[] = [
  {
    id: 'm1',
    folder: 'inbox',
    subject: '도면 검토 요청 (DWG-H101-S02-REV.A)',
    fromName: '김철수',
    fromEmail: 'vip.buyer@alpha-yard.co.kr',
    to: ['sales-h101@mycompany.co.kr'],
    date: '2026-11-18',
    timeLabel: '09:30',
    timestamp: '2026-11-18T09:30:00Z',
    snippet: '도면 첨부합니다.',
    body: '첨부 도면 확인 부탁드립니다.',
    isRead: true,
    hasAttachment: true,
    attachments: [
      { id: 'a1', filename: 'DWG-H101-S02-REV.A.pdf', kind: 'pdf', mimeType: 'application/pdf', sizeBytes: 512_000 },
      {
        id: 'a2',
        filename: 'BOM_Summary.xlsx',
        kind: 'excel',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        sizeBytes: 120_000,
      },
    ],
    source: 'pst-import',
  },
  {
    id: 'm2',
    folder: 'inbox',
    subject: '견적서 송부',
    fromName: 'Partner Sales',
    fromEmail: 'quote@alpha-supplier.com',
    to: ['general@mycompany.co.kr'],
    date: '2026-11-19',
    timeLabel: '14:00',
    timestamp: '2026-11-19T14:00:00Z',
    snippet: '견적서 첨부',
    body: '도메인 규칙으로 PRJ-H101에 매칭되어야 합니다.',
    isRead: false,
    hasAttachment: true,
    attachments: [
      { id: 'a3', filename: 'Quotation_2026.pdf', kind: 'pdf', mimeType: 'application/pdf', sizeBytes: 256_000 },
    ],
    source: 'pst-import',
  },
  {
    id: 'm3',
    folder: 'sent',
    subject: '계정 기준 자동 분류 테스트',
    fromName: '전장팀',
    fromEmail: 'elec-team@mycompany.co.kr',
    to: ['vendor@unknown-partner.org'],
    date: '2026-11-20',
    timeLabel: '16:15',
    timestamp: '2026-11-20T16:15:00Z',
    snippet: '출도 서류 발송',
    body: '연결 계정(elec-team@mycompany.co.kr)으로 PRJ-E202 매칭.',
    isRead: true,
    hasAttachment: true,
    attachments: [
      {
        id: 'a4',
        filename: 'E-310-02_Spec.docx',
        kind: 'word',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        sizeBytes: 98_000,
      },
    ],
    source: 'pst-import',
  },
  {
    id: 'm4',
    folder: 'inbox',
    subject: '일반 사내 공지',
    fromName: '총무팀',
    fromEmail: 'notice@other-corp.net',
    to: ['all@mycompany.co.kr'],
    date: '2026-11-20',
    timeLabel: '17:00',
    timestamp: '2026-11-20T17:00:00Z',
    snippet: '공지 첨부',
    body: '프로젝트 규칙 없음',
    isRead: true,
    hasAttachment: true,
    attachments: [
      { id: 'a5', filename: 'holiday_schedule.pdf', kind: 'pdf', mimeType: 'application/pdf', sizeBytes: 40_000 },
    ],
    source: 'pst-import',
  },
];

test('email/domain normalization and rule list deduplication handle messy inputs defensively', () => {
  assert.equal(extractNormalizedEmail('Kim <VIP.Buyer@Alpha-Yard.CO.KR>'), 'vip.buyer@alpha-yard.co.kr');
  assert.equal(extractEmailDomain('@Vendor-A.CO.KR'), 'vendor-a.co.kr');
  assert.equal(extractEmailDomain('https://sub.vendor.com/path'), 'sub.vendor.com');
  assert.deepEqual(
    normalizeRuleList([' @Alpha.com ', 'alpha.com', 'BETA.CO.KR', ''], 'domain'),
    ['alpha.com', 'beta.co.kr']
  );
  assert.equal(sanitizeFolderSegment('../../etc/passwd; rm -rf *'), 'etc_passwd_rm_-rf');
});

test('classifyByProjectRoutingRules follows strict priority: manual > sender > domain > account > keyword > unassigned', () => {
  // 1. Manual override beats everything
  const manual = classifyByProjectRoutingRules({
    senderEmail: 'vip.buyer@alpha-yard.co.kr',
    projects: PROJECTS,
    manualProjectCode: 'PRJ-E202',
  });
  assert.equal(manual.projectCode, 'PRJ-E202');
  assert.equal(manual.matchSource, 'manual');

  // 2. Exact sender match beats domain & account
  const sender = classifyByProjectRoutingRules({
    senderEmail: 'vip.buyer@alpha-yard.co.kr',
    accountEmail: 'elec-team@mycompany.co.kr',
    projects: PROJECTS,
  });
  assert.equal(sender.projectCode, 'PRJ-H101');
  assert.equal(sender.matchSource, 'sender');

  // 3. Domain match
  const domain = classifyByProjectRoutingRules({
    senderEmail: 'anyone@alpha-supplier.com',
    projects: PROJECTS,
  });
  assert.equal(domain.projectCode, 'PRJ-H101');
  assert.equal(domain.matchSource, 'domain');

  // 4. Account match
  const account = classifyByProjectRoutingRules({
    senderEmail: 'unknown@external.io',
    accountEmail: 'elec-team@mycompany.co.kr',
    projects: PROJECTS,
  });
  assert.equal(account.projectCode, 'PRJ-E202');
  assert.equal(account.matchSource, 'account');

  // 5. Keyword match
  const keyword = classifyByProjectRoutingRules({
    senderEmail: 'unknown@external.io',
    subject: '케이블 트레이 확인 요청',
    projects: PROJECTS,
  });
  assert.equal(keyword.projectCode, 'PRJ-E202');
  assert.equal(keyword.matchSource, 'code-or-keyword');

  // 6. Unassigned
  const unassigned = classifyByProjectRoutingRules({
    senderEmail: 'nobody@random.example',
    subject: '점심 메뉴 안내',
    projects: PROJECTS,
  });
  assert.equal(unassigned.projectCode, 'UNASSIGNED');
  assert.equal(unassigned.matchSource, 'unassigned');
});

test('collectOrganizedFiles, folder tree, report aggregation, Markdown/CSV/Script exports work end-to-end', () => {
  const files = collectOrganizedFiles({
    emails: SAMPLE_EMAILS,
    projects: PROJECTS,
    preset: 'project-month-sender',
    manualFiles: [
      {
        id: 'man-1',
        filename: 'DWG-E202-01_Layout.pdf',
        sizeBytes: 300_000,
        date: '2026-11-19',
        direction: 'received',
        senderName: 'Beta Eng',
        senderEmail: 'eng@beta-electric.com',
        recipientEmails: ['elec-team@mycompany.co.kr'],
        addedAt: '2026-11-19T10:00:00Z',
      },
    ],
  });

  assert.equal(files.length, 6);
  // m1 attachments -> PRJ-H101 via sender rule + defaultSubfolder 01_수신자료
  const m1Pdf = files.find((f) => f.filename === 'DWG-H101-S02-REV.A.pdf');
  assert.ok(m1Pdf);
  assert.equal(m1Pdf.projectCode, 'PRJ-H101');
  assert.equal(m1Pdf.matchSource, 'sender');
  assert.ok(m1Pdf.folderPath.startsWith('/PRJ-H101/01_수신자료/2026-11/'));
  assert.ok(m1Pdf.drawingKeys.includes('DWGH101S02'));

  // m2 attachment -> PRJ-H101 via domain rule
  const m2Pdf = files.find((f) => f.filename === 'Quotation_2026.pdf');
  assert.equal(m2Pdf?.projectCode, 'PRJ-H101');
  assert.equal(m2Pdf?.matchSource, 'domain');

  // m3 attachment -> PRJ-E202 via account rule
  const m3Doc = files.find((f) => f.filename === 'E-310-02_Spec.docx');
  assert.equal(m3Doc?.projectCode, 'PRJ-E202');
  assert.equal(m3Doc?.matchSource, 'account');

  // m4 attachment -> UNASSIGNED
  const m4Pdf = files.find((f) => f.filename === 'holiday_schedule.pdf');
  assert.equal(m4Pdf?.projectCode, 'UNASSIGNED');

  // Folder tree structure
  const tree = buildFolderOrganizationTree(files);
  assert.ok(tree.some((n) => n.name === 'PRJ-H101' && n.fileCount === 3));
  assert.ok(tree.some((n) => n.name === 'PRJ-E202' && n.fileCount === 2));
  assert.ok(tree.some((n) => n.name === '99_미분류' && n.fileCount === 1));

  // Executive report aggregation
  const report = buildFileOrganizationReport(files, 'project-month-sender');
  assert.equal(report.summary.totalFiles, 6);
  assert.equal(report.summary.projectCount, 2);
  assert.equal(report.summary.unassignedFileCount, 1);
  assert.equal(report.summary.receivedFileCount, 5);
  assert.equal(report.summary.sentFileCount, 1);
  assert.equal(report.byDate.length, 3);
  assert.ok(report.byParticipant.some((p) => p.email === 'vip.buyer@alpha-yard.co.kr' && p.fileCount === 2));

  // Markdown & CSV exports
  const md = formatOrganizationReportMarkdown(report);
  assert.ok(md.includes('# 프로젝트·일자·송수신인별 파일 정리 분석 보고서'));
  assert.ok(md.includes('PRJ-H101'));
  assert.ok(md.includes('2026-11-18'));

  const csv = formatOrganizationReportCsv(files);
  assert.ok(csv.startsWith('\uFEFF프로젝트코드,프로젝트명'));
  assert.ok(csv.includes('DWG-H101-S02-REV.A.pdf'));

  // PowerShell & Shell organizer scripts
  const ps1 = generateFolderOrganizerScript(files, 'ps1');
  assert.ok(ps1.includes('New-Item -ItemType Directory -Force'));
  assert.ok(ps1.includes('PRJ-H101\\01_수신자료'));

  const sh = generateFolderOrganizerScript(files, 'sh');
  assert.ok(sh.includes('#!/usr/bin/env bash'));
  assert.ok(sh.includes('mkdir -p "$TARGET_ROOT/PRJ-H101/01_수신자료'));

  // Routing suggestions discovery
  const suggestions = discoverRoutingSuggestions(SAMPLE_EMAILS, 'owner@mycompany.co.kr');
  assert.ok(suggestions.senders.includes('vip.buyer@alpha-yard.co.kr'));
  assert.ok(suggestions.domains.includes('alpha-supplier.com'));
  assert.ok(suggestions.accounts.includes('owner@mycompany.co.kr'));
});

test('empty emails and files produce valid zero-state report without throwing', () => {
  const files = collectOrganizedFiles({
    emails: [],
    projects: [],
    preset: 'date-project-sender',
  });
  assert.deepEqual(files, []);
  const tree = buildFolderOrganizationTree(files);
  assert.deepEqual(tree, []);
  const report = buildFileOrganizationReport(files, 'date-project-sender');
  assert.equal(report.summary.totalFiles, 0);
  assert.ok(formatOrganizationReportMarkdown(report).includes('정리 대상 파일이 없습니다.'));
});
