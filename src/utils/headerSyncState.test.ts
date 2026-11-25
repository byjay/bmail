import test from 'node:test';
import assert from 'node:assert/strict';
import { buildHeaderAccountSyncView, formatSyncClockTime } from './headerSyncState';

test('buildHeaderAccountSyncView returns syncing state when Microsoft 365 background sync is active', () => {
  const view = buildHeaderAccountSyncView({
    microsoftConnected: true,
    microsoftEmail: 'kim@company.co.kr',
    pop3Connected: false,
    pop3Email: null,
    isSyncingGraph: true,
    isSyncingPop3: false,
    lastSyncedAt: '2026-11-21T04:15:30.000Z',
    lastSyncSource: 'microsoft365',
  });

  assert.equal(view.mode, 'syncing');
  assert.equal(view.isSyncing, true);
  assert.equal(view.isConnected, true);
  assert.equal(view.compactStatusLabel, '동기화 중');
  assert.match(view.activeSourcesLabel, /Microsoft 365/);
  assert.match(view.ariaLabel, /백그라운드 동기화 중/);
  assert.match(view.tooltip, /kim@company\.co\.kr/);
});

test('buildHeaderAccountSyncView returns syncing state when POP3/IMAP sync is active (including dual sync)', () => {
  const pop3Only = buildHeaderAccountSyncView({
    microsoftConnected: false,
    microsoftEmail: null,
    pop3Connected: true,
    pop3Email: 'designsir@seastar.co.kr',
    pop3Protocol: 'pop3',
    pop3ProviderName: '메일플러그 (Mailplug)',
    isSyncingGraph: false,
    isSyncingPop3: true,
    lastSyncedAt: null,
    lastSyncSource: null,
  });

  assert.equal(pop3Only.mode, 'syncing');
  assert.equal(pop3Only.isSyncing, true);
  assert.match(pop3Only.activeSourcesLabel, /메일플러그.*POP3/);

  const both = buildHeaderAccountSyncView({
    microsoftConnected: true,
    microsoftEmail: 'kim@company.co.kr',
    pop3Connected: true,
    pop3Email: 'designsir@seastar.co.kr',
    pop3Protocol: 'imap',
    pop3ProviderName: '메일플러그',
    isSyncingGraph: true,
    isSyncingPop3: true,
    lastSyncedAt: null,
    lastSyncSource: null,
  });

  assert.equal(both.mode, 'syncing');
  assert.match(both.activeSourcesLabel, /Microsoft 365/);
  assert.match(both.activeSourcesLabel, /IMAP/);
});

test('buildHeaderAccountSyncView returns connected-current state with formatted last-synced clock time', () => {
  const sampleDate = new Date(2026, 10, 21, 9, 5, 8);
  const view = buildHeaderAccountSyncView({
    microsoftConnected: false,
    microsoftEmail: null,
    pop3Connected: true,
    pop3Email: 'designsir@seastar.co.kr',
    pop3Protocol: 'pop3',
    pop3ProviderName: '메일플러그',
    isSyncingGraph: false,
    isSyncingPop3: false,
    lastSyncedAt: sampleDate.toISOString(),
    lastSyncSource: 'pop3',
  });

  assert.equal(view.mode, 'connected');
  assert.equal(view.isSyncing, false);
  assert.equal(view.isConnected, true);
  assert.equal(view.lastSyncedTimeLabel, '09:05:08');
  assert.equal(view.compactStatusLabel, '최신 09:05');
  assert.match(view.tooltip, /최신 동기화 09:05:08/);
});

test('buildHeaderAccountSyncView handles disconnected state and invalid timestamps defensively', () => {
  assert.equal(formatSyncClockTime('not-a-date'), null);
  assert.equal(formatSyncClockTime(null), null);

  const view = buildHeaderAccountSyncView({
    microsoftConnected: false,
    microsoftEmail: null,
    pop3Connected: false,
    pop3Email: null,
    isSyncingGraph: false,
    isSyncingPop3: false,
    lastSyncedAt: 'invalid',
    lastSyncSource: null,
  });

  assert.equal(view.mode, 'disconnected');
  assert.equal(view.isSyncing, false);
  assert.equal(view.isConnected, false);
  assert.equal(view.ariaLabel, '계정 및 연결 상태');
});
