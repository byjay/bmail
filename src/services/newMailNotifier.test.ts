import test from 'node:test';
import assert from 'node:assert/strict';
import type { MailItem } from '../types/mail';
import {
  buildDeterministicQuickSummary,
  detectNewIncomingEmails,
  summarizeNewMailsForNotification,
} from './newMailNotifier';

function makeMail(overrides: Partial<MailItem> = {}): MailItem {
  return {
    id: 'm-1',
    folder: 'inbox',
    subject: '일반 안내',
    fromName: '홍길동',
    fromEmail: 'hong@seastargo.com',
    to: ['me@seastargo.com'],
    date: '2026-11-24',
    timestamp: '2026-11-24T09:00:00.000Z',
    timeLabel: '09:00',
    snippet: '안내 말씀 드립니다.',
    body: '안녕하세요, 내일 오전 10시 회의 진행 예정입니다.',
    importance: 'normal',
    isRead: false,
    source: 'pop3-live' as MailItem['source'],
    ...overrides,
  };
}

test('detectNewIncomingEmails returns only unseen inbox items sorted newest first', () => {
  const seen = new Set<string>(['m-old']);
  const incoming: MailItem[] = [
    makeMail({ id: 'm-old', timestamp: '2026-11-24T08:00:00.000Z' }),
    makeMail({ id: 'm-sent', folder: 'sent', timestamp: '2026-11-24T12:00:00.000Z' }),
    makeMail({ id: 'm-new-1', timestamp: '2026-11-24T10:00:00.000Z' }),
    makeMail({ id: 'm-new-2', timestamp: '2026-11-24T11:30:00.000Z' }),
  ];

  const fresh = detectNewIncomingEmails(incoming, seen);
  assert.deepEqual(
    fresh.map((m) => m.id),
    ['m-new-2', 'm-new-1']
  );
});

test('buildDeterministicQuickSummary extracts urgency, drawing refs, deadlines, and attachment count', () => {
  const urgentMail = makeMail({
    id: 'm-urgent',
    subject: '[긴급] H101 블록 도면 DWG-H101-S02-REV.A 승인 요청',
    body: '2026-11-26까지 회신 요망합니다. 수정 도면 첨부합니다.',
    importance: 'high',
    hasAttachment: true,
    attachments: [
      {
        id: 'att-1',
        filename: 'DWG-H101-S02.pdf',
        mimeType: 'application/pdf',
        sizeBytes: 102400,
        kind: 'pdf',
      },
    ],
  });

  const summary = buildDeterministicQuickSummary(urgentMail);
  assert.equal(summary.urgency, 'high');
  assert.ok(summary.drawings.some((d) => d.includes('DWG-H101-S02')));
  assert.deepEqual(summary.deadlines, ['2026-11-26']);
  assert.ok(summary.summary.includes('긴급'));
  assert.ok(summary.summary.includes('도면'));
  assert.ok(summary.summary.includes('첨부 1건'));
  assert.equal(summary.summarizedBy, 'deterministic');
});

test('summarizeNewMailsForNotification falls back cleanly offline and respects maxItems', async () => {
  const items = [
    makeMail({ id: 'n1', subject: '첫 번째 새 메일' }),
    makeMail({ id: 'n2', subject: '두 번째 새 메일' }),
    makeMail({ id: 'n3', subject: '세 번째 새 메일' }),
    makeMail({ id: 'n4', subject: '네 번째 메일 (제외)' }),
  ];

  const summaries = await summarizeNewMailsForNotification(items, {
    maxItems: 3,
    enableAi: false,
  });

  assert.equal(summaries.length, 3);
  assert.deepEqual(
    summaries.map((s) => s.emailId),
    ['n1', 'n2', 'n3']
  );
  assert.equal(summaries[0].summarizedBy, 'deterministic');
});
