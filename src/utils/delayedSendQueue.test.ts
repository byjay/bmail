import assert from 'node:assert/strict';
import test from 'node:test';
import {
  cancelDelayedSendItem,
  createDelayedSendItem,
  formatCountdownLabel,
  getRemainingSeconds,
} from './delayedSendQueue';

test('delayedSendQueue creates queued items, tracks countdown accurately, and supports clean cancellation', () => {
  const t0 = 1_700_000_000_000;
  const item = createDelayedSendItem(
    {
      to: 'partner@example.com',
      cc: '',
      subject: '  H2710 수정 POS 송부  ',
      body: '첨부 확인 부탁드립니다.',
      kind: 'reply',
    },
    30,
    t0
  );

  assert.equal(item.subject, 'H2710 수정 POS 송부');
  assert.equal(item.status, 'queued');
  assert.equal(getRemainingSeconds(item, t0), 30);
  assert.equal(getRemainingSeconds(item, t0 + 12_400), 18);
  assert.equal(getRemainingSeconds(item, t0 + 31_000), 0);
  assert.equal(formatCountdownLabel(75), '1분 15초');
  assert.equal(formatCountdownLabel(9), '9초');

  const other = createDelayedSendItem(
    { to: 'a@example.com', cc: '', subject: 'Other', body: 'hi', kind: 'new' },
    10,
    t0
  );
  const { cancelled, remainingQueue } = cancelDelayedSendItem([item, other], item.id);
  assert.equal(cancelled?.status, 'cancelled');
  assert.equal(cancelled?.to, 'partner@example.com');
  assert.equal(remainingQueue.length, 1);
  assert.equal(remainingQueue[0].id, other.id);
});
