import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import type { ProjectDefinition } from '../utils/deterministicAnalyzer';
import { buildUnreadPriorityBriefing } from './unreadBriefingEngine';

const SAMPLE_PROJECTS: ProjectDefinition[] = [
  {
    id: 'p1',
    code: 'H042',
    name: 'H042 Exmar LPG',
    keywords: ['H042', 'EXMAR', '생산설계'],
    color: '#0F6CBD',
    isFavorite: false,
    createdAt: '2026-01-01T00:00:00.000Z',
  },
  {
    id: 'p2',
    code: 'KYTS1037',
    name: 'KYTS1037 선급',
    keywords: ['KYTS1037', '한국선급'],
    color: '#D83B01',
    isFavorite: false,
    createdAt: '2026-01-01T00:00:00.000Z',
  },
];

function makeMail(partial: Partial<MailItem>): MailItem {
  return {
    id: partial.id || 'm-1',
    folder: partial.folder || 'inbox',
    folderName: '받은 편지함',
    subject: partial.subject || '테스트',
    fromName: partial.fromName || '테스터',
    fromEmail: partial.fromEmail || 'user@example.com',
    to: ['me@seastargo.com'],
    date: '2026-11-25',
    timestamp: partial.timestamp || '2026-11-25T09:00:00.000Z',
    timeLabel: '09:00',
    snippet: partial.snippet || partial.body || '',
    body: partial.body || '본문',
    hasAttachment: false,
    isRead: partial.isRead ?? false,
    isFlagged: partial.isFlagged ?? false,
    importance: partial.importance || 'normal',
    category: '',
    source: partial.source || 'outlook-live',
    ...partial,
  };
}

test('buildUnreadPriorityBriefing analyzes unread mails read-only, prioritizes urgent replies, and groups by project', () => {
  const mails: MailItem[] = [
    makeMail({
      id: 'read-ignore',
      subject: '[긴급] 이미 읽은 메일',
      isRead: true,
    }),
    makeMail({
      id: 'old-4yr-ignore',
      subject: '[긴급] 상상인 선박기계 - SCA 35K FD - 생산 설계 도면 출도 회신 요망',
      body: '2023-03-09 대용량 첨부파일 긴급 회신 요망',
      date: '2023-03-09',
      timestamp: '2023-03-09T09:00:00.000Z',
      importance: 'high',
      isRead: false,
    }),
    makeMail({
      id: 'normal-gen',
      subject: '사내 공지: 11월 식단표',
      body: '이번 달 식단 안내드립니다.',
      isRead: false,
    }),
    makeMail({
      id: 'urgent-ky',
      subject: '[긴급] KYTS1037 한국선급 승인 미결 도면 DWG-KY10-E01 금일 회신 요망',
      body: '2026-11-25 금일 중 반드시 승인 회신 부탁드립니다.',
      importance: 'high',
      isRead: false,
    }),
    makeMail({
      id: 'imp-h042',
      subject: 'H042 EXMAR 생산설계 도면 검토',
      body: 'DWG-H042-P102 도면 검토 부탁드립니다.',
      isRead: false,
    }),
  ];

  const briefing = buildUnreadPriorityBriefing(mails, SAMPLE_PROJECTS, undefined, {
    maxAgeDays: 14,
    nowMs: Date.parse('2026-11-25T15:00:00.000Z'),
  });

  // Never marks unread emails as read
  for (const orig of mails) {
    if (orig.id !== 'read-ignore') {
      assert.equal(orig.isRead, false);
    }
  }

  assert.equal(briefing.excludedOldUnreadCount, 1);
  assert.equal(briefing.totalUnread, 3);
  assert.equal(briefing.urgentReplyCount, 1);
  assert.equal(briefing.urgentReplyAlerts[0].emailId, 'urgent-ky');
  assert.equal(briefing.urgentReplyAlerts[0].urgencyLabel, '반드시 긴급회신');

  // Top project group is KYTS1037 (has urgent-reply item)
  assert.equal(briefing.projectGroups[0].projectCode, 'KYTS1037');
  assert.equal(briefing.projectGroups[0].urgentReplyCount, 1);
  assert.ok(briefing.projectGroups.some((g) => g.projectCode === 'H042'));
});
