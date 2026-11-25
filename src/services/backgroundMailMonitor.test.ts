import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { MailItem } from '../types/mail';
import {
  BackgroundMailMonitor,
  type SgwBridgeLike,
} from './backgroundMailMonitor';
import { LocalMailDb } from './localMailDb';
import { Pop3SessionStore } from './pop3AutoConnect';

function makeMail(overrides: Partial<MailItem> = {}): MailItem {
  return {
    id: 'bg-1',
    folder: 'inbox',
    subject: '선체 블록 도면 검토 요청',
    fromName: '김설계',
    fromEmail: 'kim@seastargo.com',
    to: ['designsir@seastargo.com'],
    date: '2026-11-24',
    timestamp: '2026-11-24T09:00:00.000Z',
    timeLabel: '09:00',
    snippet: 'DWG-H101 검토 부탁드립니다.',
    body: '안녕하세요, DWG-H101 개정본 확인 요청드립니다.',
    importance: 'normal',
    isRead: false,
    source: 'pop3-live' as MailItem['source'],
    ...overrides,
  };
}

function makeMockSgw(connected = true): {
  bridge: SgwBridgeLike;
  dmSent: Array<{ receiverId: string; content: string }>;
  roomSent: Array<{ roomId: string; content: string }>;
} {
  const dmSent: Array<{ receiverId: string; content: string }> = [];
  const roomSent: Array<{ roomId: string; content: string }> = [];
  const bridge: SgwBridgeLike = {
    getStatus: () =>
      connected
        ? {
            connected: true,
            sgwBaseUrl: 'https://sgw.seastar.work',
            user: { sub: 'emp-001', name: '김봉정', email: 'designsir@seastargo.com' },
          }
        : {
            connected: false,
            sgwBaseUrl: 'https://sgw.seastar.work',
            user: null,
          },
    sendDirectMessage: async (payload) => {
      dmSent.push({ receiverId: payload.receiverId, content: payload.content });
      return { ok: true, id: `dm-${dmSent.length}` };
    },
    sendRoomMessage: async (roomId, payload) => {
      roomSent.push({ roomId, content: payload.content });
      return { ok: true, id: `rm-${roomSent.length}` };
    },
  };
  return { bridge, dmSent, roomSent };
}

test('BackgroundMailMonitor pre-organizes dashboard digest and dispatches SGW self-DM only if user has not viewed', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-bg-mon-'));
  const db = new LocalMailDb(dir);
  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable: ${db.getInitError() || 'unknown'}`);
      return;
    }

    let nowMs = 1_700_000_000_000;
    const pop3Sessions = new Pop3SessionStore();
    const { bridge, dmSent } = makeMockSgw(true);

    const monitor = new BackgroundMailMonitor({
      localDb: db,
      pop3Sessions,
      sgwChatBridge: bridge,
      enableAiSummary: false,
      nowFn: () => nowMs,
      pollOutlookFn: () => [],
    });
    monitor.updateSettings({
      enabled: true,
      unviewedGraceSec: 60,
      sgwPushEnabled: true,
      sgwUrgentOnly: false,
    });

    const mViewed = makeMail({ id: 'mail-viewed', subject: '먼저 본 메일' });
    const mUnviewed = makeMail({
      id: 'mail-unviewed',
      subject: '[긴급] DWG-2201 승인 요청',
      importance: 'high',
    });

    db.upsertEmailsBatch([
      { ...mViewed, storeName: 'test-store' },
      { ...mUnviewed, storeName: 'test-store' },
    ]);
    await monitor.ingestNewEmails([mViewed, mUnviewed]);

    // Dashboard digest proactively computed & saved to SQLite KV
    const stBefore = monitor.getStatus();
    assert.ok(stBefore.digest);
    assert.equal(stBefore.digest.totalEmails, 2);
    assert.equal(stBefore.pendingUnviewedCount, 2);

    // User views/reads first email within grace period -> cancels its SGW alert
    monitor.markViewed('mail-viewed');
    const sentEarly = await monitor.dispatchDueSgwAlerts();
    assert.equal(sentEarly, 0);

    // Advance past 60s unviewed grace period -> only mail-unviewed sends SGW self-DM once
    nowMs += 65_000;
    const sentDue = await monitor.dispatchDueSgwAlerts();
    assert.equal(sentDue, 1);
    assert.equal(dmSent.length, 1);
    assert.equal(dmSent[0].receiverId, 'emp-001');
    assert.ok(dmSent[0].content.includes('[긴급 메일]'));
    assert.ok(dmSent[0].content.includes('DWG-2201'));

    // Second dispatch pass never double-sends
    const sentAgain = await monitor.dispatchDueSgwAlerts();
    assert.equal(sentAgain, 0);
    assert.equal(dmSent.length, 1);

    // Tray alerts consume unviewed mail-unviewed once
    const trayAlerts = monitor.consumeUnnotifiedTrayAlerts();
    assert.deepEqual(
      trayAlerts.map((a) => a.emailId),
      ['mail-unviewed']
    );
    assert.equal(monitor.consumeUnnotifiedTrayAlerts().length, 0);
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('BackgroundMailMonitor runBackgroundPollCycle fetches persisted POP3 session and respects sgwUrgentOnly', async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-bg-poll-'));
  const db = new LocalMailDb(dir);
  try {
    if (!db.isReady()) {
      t.skip(`node:sqlite unavailable: ${db.getInitError() || 'unknown'}`);
      return;
    }

    let nowMs = 1_800_000_000_000;
    const pop3Sessions = new Pop3SessionStore({ baseDir: dir, secretKey: 'bg-test-key' });
    pop3Sessions.create({
      email: 'user@seastargo.com',
      password: 'app-password-1234',
      providerId: 'mailplug',
      providerName: 'Mailplug',
      protocol: 'pop3',
      usernameFormat: 'full-email',
      pop3Host: 'pop3.mailplug.co.kr',
      pop3Port: 995,
      pop3Tls: true,
      imapHost: 'imap.mailplug.co.kr',
      imapPort: 993,
      imapTls: true,
      smtpHost: 'smtp.mailplug.co.kr',
      smtpPort: 465,
      smtpTls: true,
    });

    const { bridge, roomSent } = makeMockSgw(true);
    const monitor = new BackgroundMailMonitor({
      localDb: db,
      pop3Sessions,
      sgwChatBridge: bridge,
      enableAiSummary: false,
      nowFn: () => nowMs,
      pollPop3SessionFn: async () => [
        makeMail({ id: 'pop-normal', subject: '일반 뉴스레터', importance: 'normal' }),
        makeMail({ id: 'pop-urgent', subject: '[긴급] 당일 납기 변경', importance: 'high' }),
      ],
      pollOutlookFn: () => [],
    });

    monitor.updateSettings({
      enabled: true,
      unviewedGraceSec: 0,
      sgwPushEnabled: true,
      sgwUrgentOnly: true,
      sgwTargetRoomId: 'room-mobile-alerts',
    });

    const res = await monitor.runBackgroundPollCycle();
    assert.equal(res.addedCount, 2);
    // Immediate send (graceSec=0) with urgentOnly=true -> only pop-urgent dispatched to room
    assert.equal(res.sgwSentCount, 1);
    assert.equal(roomSent.length, 1);
    assert.equal(roomSent[0].roomId, 'room-mobile-alerts');
    assert.ok(roomSent[0].content.includes('당일 납기 변경'));
  } finally {
    db.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});
