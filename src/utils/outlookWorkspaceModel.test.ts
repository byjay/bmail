import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyQuickStepActions,
  assignEmailToProject,
  buildSweepRuleForSender,
  createProjectVaultFile,
  DEFAULT_FAVORITE_FOLDER_IDS,
  DEFAULT_OUTLOOK_CATEGORIES,
  DEFAULT_OUTLOOK_PREFERENCES,
  DEFAULT_OUTLOOK_QUICK_STEPS,
  EMPTY_PROJECT_MAIL_OVERRIDES,
  evaluateOutlookMailRules,
  extractShipbuildingDrawingRefs,
  filterEmailsBySearchFolder,
  isBlockedSender,
  isEmailInProject,
  removeEmailFromProject,
  removeProjectSubfolderSafely,
  sanitizeOutlookPreferences,
  toggleFavoriteFolder,
  type OutlookMailRule,
  type ProjectTreeSubfolder,
  type ProjectVaultFile,
} from './outlookWorkspaceModel';
import type { ProjectDefinition } from './deterministicAnalyzer';

const SAMPLE_PROJECTS: ProjectDefinition[] = [
  {
    id: 'p1',
    code: 'PRJ-H101',
    name: '선체 블록 설계',
    keywords: ['블록', '선체'],
    color: '#0F6CBD',
    isFavorite: true,
    createdAt: '2026-11-20T00:00:00Z',
  },
];

test('extractShipbuildingDrawingRefs detects hull, piping, electrical, and revision drawing numbers', () => {
  const text =
    '첨부 도면 DWG-H101-S02-REV.A 및 BLK-204-A, E-310-02, P-405-R2 검토 요청드립니다. 중복 dwg_h101_s02 확인.';
  const refs = extractShipbuildingDrawingRefs(text);
  assert.ok(refs.length >= 4);
  assert.ok(refs.some((r) => r.canonKey === 'DWGH101S02' && r.revision === 'REV.A'));
  assert.ok(refs.some((r) => r.canonKey === 'BLK204A' && r.discipline === '선체/구조'));
  assert.ok(refs.some((r) => r.canonKey === 'E31002' && r.discipline === '전장'));
  assert.ok(refs.some((r) => r.canonKey === 'P405R2' || r.canonKey === 'P405'));

  // Pure words without digits like MESSAGE / SUBJECT must never match as drawing refs
  assert.deepEqual(extractShipbuildingDrawingRefs('MESSAGE 확인 요청 · DRAFT STATUS'), []);

  // Email usernames (e.g. elcon0955@naver.com) and URLs must never be mistaken for ELC drawing refs
  assert.deepEqual(
    extractShipbuildingDrawingRefs('발신: elcon0955@naver.com 참고 https://example.com/elc-999'),
    []
  );

  // Empty input safely returns empty list
  assert.deepEqual(extractShipbuildingDrawingRefs(''), []);
});

test('toggleFavoriteFolder adds and removes folders and project nodes cleanly', () => {
  const added = toggleFavoriteFolder(DEFAULT_FAVORITE_FOLDER_IDS, 'project:PRJ-H101');
  assert.ok(added.includes('project:PRJ-H101'));

  const removed = toggleFavoriteFolder(added, 'drafts');
  assert.equal(removed.includes('drafts'), false);
  assert.ok(removed.includes('inbox'));
});

test('project email assign/remove overrides auto-detection accurately', () => {
  const autoMail = {
    id: 'm1',
    subject: '[PRJ-H101] 블록 조립 도면 승인',
    body: '선체 블록 확인 부탁드립니다.',
  };
  const unrelatedMail = {
    id: 'm2',
    subject: '일반 회의 안내',
    body: '내일 오전 회의 진행합니다.',
  };

  assert.equal(isEmailInProject(autoMail, 'PRJ-H101', SAMPLE_PROJECTS, EMPTY_PROJECT_MAIL_OVERRIDES), true);
  assert.equal(isEmailInProject(unrelatedMail, 'PRJ-H101', SAMPLE_PROJECTS, EMPTY_PROJECT_MAIL_OVERRIDES), false);

  // Put unrelated mail into project
  const afterAdd = assignEmailToProject(EMPTY_PROJECT_MAIL_OVERRIDES, 'PRJ-H101', 'm2');
  assert.equal(isEmailInProject(unrelatedMail, 'PRJ-H101', SAMPLE_PROJECTS, afterAdd), true);

  // Remove auto-detected mail out of project
  const afterRemove = removeEmailFromProject(afterAdd, 'PRJ-H101', 'm1');
  assert.equal(isEmailInProject(autoMail, 'PRJ-H101', SAMPLE_PROJECTS, afterRemove), false);
});

test('project file vault indexes SEDMS drawing codes and safely preserves files on subfolder deletion', () => {
  const file = createProjectVaultFile({
    projectCode: 'PRJ-H101',
    subfolderId: 'sf-child',
    filename: 'DWG-H101-S02-REV.B_HullBlock.pdf',
    sizeBytes: 204800,
    mimeType: 'application/pdf',
  });
  assert.equal(file.kind, 'pdf');
  assert.equal(file.drawingNoKey, 'DWGH101S02');
  assert.equal(file.sedmsStatus, 'linked');

  const subfolders: ProjectTreeSubfolder[] = [
    { id: 'sf-root', projectCode: 'PRJ-H101', parentId: null, name: '설계 도면', createdAt: '2026-11-20' },
    { id: 'sf-child', projectCode: 'PRJ-H101', parentId: 'sf-root', name: 'Rev.B', createdAt: '2026-11-20' },
  ];
  const files: ProjectVaultFile[] = [file];

  const cleaned = removeProjectSubfolderSafely(subfolders, files, 'sf-root');
  assert.equal(cleaned.subfolders.length, 0);
  // Orphaned file is preserved and moved safely up to project root (null)
  assert.equal(cleaned.files.length, 1);
  assert.equal(cleaned.files[0].subfolderId, null);
});

test('Outlook rules and blocked/safe sender lists evaluate deterministically', () => {
  const rules: OutlookMailRule[] = [
    {
      id: 'r1',
      name: '도면 메일 프로젝트 배정',
      enabled: true,
      conditionField: 'drawing',
      conditionValue: 'BLK204',
      actionType: 'assign-project',
      actionTarget: 'PRJ-H101',
    },
    {
      id: 'r2',
      name: '첨부 메일 플래그',
      enabled: true,
      conditionField: 'hasAttachment',
      conditionValue: '',
      actionType: 'flag',
    },
    {
      id: 'r3',
      name: '비활성 규칙',
      enabled: false,
      conditionField: 'subject',
      conditionValue: '블록',
      actionType: 'move-trash',
    },
  ];

  const evalRes = evaluateOutlookMailRules(
    {
      subject: '블록 BLK-204-A 도면 송부',
      fromName: '설계팀',
      fromEmail: 'design@yard.example',
      body: '검토 바랍니다.',
      hasAttachment: true,
    },
    rules
  );
  assert.equal(evalRes.assignedProjectCode, 'PRJ-H101');
  assert.equal(evalRes.shouldFlag, true);
  assert.equal(evalRes.shouldMoveTrash, false);

  assert.equal(isBlockedSender({ fromEmail: 'spam@promo.example' }, ['@promo.example'], []), true);
  // Safe sender overrides blocked domain
  assert.equal(
    isBlockedSender({ fromEmail: 'vip@promo.example' }, ['@promo.example'], ['vip@promo.example']),
    false
  );
});

test('sanitizeOutlookPreferences gracefully recovers from malformed or partial storage data', () => {
  const fromNull = sanitizeOutlookPreferences(null);
  assert.deepEqual(fromNull, DEFAULT_OUTLOOK_PREFERENCES);

  const fromPartial = sanitizeOutlookPreferences({
    general: { density: 'compact', accentTheme: 'invalid-color' },
    mail: {
      readingPanePosition: 'bottom',
      undoSendSeconds: 99,
      focusedInbox: false,
      singleLineRibbon: true,
      markAsReadMode: 'delay',
      markAsReadDelaySec: 99,
      afterDeleteAction: 'prev',
      defaultImportance: 'high',
    },
  });
  assert.equal(fromPartial.general.density, 'compact');
  assert.equal(fromPartial.general.accentTheme, 'blue');
  assert.equal(fromPartial.mail.readingPanePosition, 'bottom');
  assert.equal(fromPartial.mail.undoSendSeconds, 5);
  assert.equal(fromPartial.mail.focusedInbox, false);
  assert.equal(fromPartial.mail.singleLineRibbon, true);
  assert.equal(fromPartial.mail.markAsReadMode, 'delay');
  assert.equal(fromPartial.mail.markAsReadDelaySec, 30);
  assert.equal(fromPartial.mail.afterDeleteAction, 'prev');
  assert.equal(fromPartial.mail.defaultImportance, 'high');
  assert.equal(fromPartial.mail.categories.length, DEFAULT_OUTLOOK_CATEGORIES.length);
  assert.equal(fromPartial.mail.quickSteps.length, DEFAULT_OUTLOOK_QUICK_STEPS.length);
});

test('filterEmailsBySearchFolder, applyQuickStepActions, and buildSweepRuleForSender handle Outlook parity actions', () => {
  const sampleMails = [
    {
      id: 'm-unread',
      folder: 'inbox' as const,
      subject: '미읽은 메일',
      fromName: 'A',
      fromEmail: 'a@yard.test',
      to: [] as string[],
      date: '2026-11-24',
      timestamp: '2026-11-24T09:00:00.000Z',
      timeLabel: '09:00',
      snippet: '',
      body: '짧은 본문',
      hasAttachment: false,
      importance: 'normal' as const,
      isRead: false,
      isFlagged: false,
      source: 'pst-import' as const,
    },
    {
      id: 'm-large-flag',
      folder: 'inbox' as const,
      subject: '대용량 도면 첨부',
      fromName: 'B',
      fromEmail: 'b@yard.test',
      to: [] as string[],
      date: '2026-11-24',
      timestamp: '2026-11-24T10:00:00.000Z',
      timeLabel: '10:00',
      snippet: '',
      body: '본문',
      hasAttachment: true,
      attachments: [
        {
          id: 'att-1',
          filename: 'GA.pdf',
          mimeType: 'application/pdf',
          sizeBytes: 2 * 1024 * 1024,
          kind: 'pdf' as const,
        },
      ],
      importance: 'high' as const,
      isRead: true,
      isFlagged: true,
      source: 'pst-import' as const,
    },
  ];

  const snoozeMap = { 'm-unread': '2026-11-25T09:00:00.000Z' };
  assert.equal(filterEmailsBySearchFolder(sampleMails, 'unread').map((m) => m.id).join(','), 'm-unread');
  assert.equal(filterEmailsBySearchFolder(sampleMails, 'flagged').map((m) => m.id).join(','), 'm-large-flag');
  assert.equal(filterEmailsBySearchFolder(sampleMails, 'has-attachments').map((m) => m.id).join(','), 'm-large-flag');
  assert.equal(filterEmailsBySearchFolder(sampleMails, 'large-mail').map((m) => m.id).join(','), 'm-large-flag');
  assert.equal(
    filterEmailsBySearchFolder(sampleMails, 'snoozed', snoozeMap, '2026-11-24T12:00:00.000Z').map((m) => m.id).join(','),
    'm-unread'
  );

  const patch = applyQuickStepActions(sampleMails[0], DEFAULT_OUTLOOK_QUICK_STEPS[2]);
  assert.equal(patch.markRead, true);
  assert.equal(patch.assignCategoryId, 'cat-drawing');

  const sweepRule = buildSweepRuleForSender('Promo@Example.COM', 'archive');
  assert.equal(sweepRule.actionType, 'archive');
  assert.equal(sweepRule.conditionValue, 'promo@example.com');
  const ruleEval = evaluateOutlookMailRules(
    {
      subject: 'News',
      fromName: 'Promo',
      fromEmail: 'promo@example.com',
      body: 'hi',
      hasAttachment: false,
    },
    [sweepRule]
  );
  assert.equal(ruleEval.shouldArchive, true);
});
