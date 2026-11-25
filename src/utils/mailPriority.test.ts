import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import { classifyMailPriority, compareMailsByPriority } from './mailPriority';
import type { ProjectDefinition } from './deterministicAnalyzer';

const SAMPLE_PROJECTS: ProjectDefinition[] = [
  {
    id: 'p1',
    code: 'H1001',
    name: 'LNG 선체 설계',
    keywords: ['LNG', '선체', 'H1001'],
    color: '#0F6CBD',
    createdAt: '2026-11-01T00:00:00Z',
    isFavorite: true,
  },
];

function makeMail(overrides: Partial<MailItem>): MailItem {
  return {
    id: overrides.id ?? 'm1',
    folder: overrides.folder ?? 'inbox',
    folderName: overrides.folderName ?? '받은 편지함',
    fromName: overrides.fromName ?? '김담당',
    fromEmail: overrides.fromEmail ?? 'kim@company.co.kr',
    to: overrides.to ?? ['me@company.co.kr'],
    subject: overrides.subject ?? '일반 안내',
    date: overrides.date ?? '2026-11-20',
    timeLabel: overrides.timeLabel ?? '09:30',
    timestamp: overrides.timestamp ?? '2026-11-20T09:30:00Z',
    snippet: overrides.snippet ?? '',
    body: overrides.body ?? '',
    isRead: overrides.isRead ?? true,
    isFlagged: overrides.isFlagged ?? false,
    source: overrides.source ?? 'pst-import',
    hasAttachment: overrides.hasAttachment ?? false,
    attachments: overrides.attachments,
    category: overrides.category,
  };
}

test('classifyMailPriority: marks urgent deadline + project mail as 긴급', () => {
  const mail = makeMail({
    id: 'urgent-1',
    subject: '[H1001] 긴급: 선체 도면 마감 확인 요청 (오늘까지)',
    body: '금일 마감까지 승인 회신 부탁드립니다.',
    isRead: false,
  });

  const result = classifyMailPriority(mail, { projects: SAMPLE_PROJECTS });
  assert.equal(result.level, 'urgent');
  assert.equal(result.label, '긴급');
  assert.equal(result.matchedProjectCode, 'H1001');
  assert.ok(result.score >= 70);
});

test('classifyMailPriority: marks project-relevant approval mail as 중요', () => {
  const mail = makeMail({
    id: 'important-1',
    subject: '[H1001] 도면 검토 의견 공유',
    body: '첨부 사양서 검토 부탁드립니다.',
    hasAttachment: true,
    isRead: true,
  });

  const result = classifyMailPriority(mail, { projects: SAMPLE_PROJECTS });
  assert.equal(result.level, 'important');
  assert.equal(result.label, '중요');
  assert.equal(result.matchedProjectCode, 'H1001');
  assert.ok(result.score >= 38 && result.score < 70);
});

test('classifyMailPriority: marks routine or promotional newsletter as 일반', () => {
  const promo = makeMail({
    id: 'normal-1',
    fromName: '뉴스레터',
    fromEmail: 'noreply@newsletter.example.com',
    subject: '주간 뉴스레터 안내 (수신거부 가능)',
    body: 'Unsubscribe link at bottom.',
  });

  const result = classifyMailPriority(promo, { projects: SAMPLE_PROJECTS });
  assert.equal(result.level, 'normal');
  assert.equal(result.label, '일반');
  assert.ok(result.score < 38);
});

test('compareMailsByPriority: orders 긴급 -> 중요 -> 일반, then newer timestamp', () => {
  const normalMail = makeMail({
    id: 'm-normal',
    subject: '주간 공지',
    timestamp: '2026-11-20T14:00:00Z',
  });
  const importantMail = makeMail({
    id: 'm-important',
    subject: '[H1001] 도면 검토',
    timestamp: '2026-11-20T10:00:00Z',
  });
  const urgentOlder = makeMail({
    id: 'm-urgent-old',
    subject: '긴급 장애 발생 즉시 확인',
    timestamp: '2026-11-20T08:00:00Z',
  });
  const urgentNewer = makeMail({
    id: 'm-urgent-new',
    subject: '[H1001] 긴급 마감 지연 대응',
    timestamp: '2026-11-20T11:00:00Z',
  });

  const list = [normalMail, importantMail, urgentOlder, urgentNewer];
  const map = new Map(
    list.map((m) => [m.id, classifyMailPriority(m, { projects: SAMPLE_PROJECTS })])
  );

  const sorted = [...list].sort((a, b) => compareMailsByPriority(a, b, map));
  assert.deepEqual(
    sorted.map((m) => m.id),
    ['m-urgent-new', 'm-urgent-old', 'm-important', 'm-normal']
  );
});
