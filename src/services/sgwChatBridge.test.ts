import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  buildMailShareText,
  normalizeBase64Payload,
  SgwBridgeError,
  SgwChatBridge,
} from './sgwChatBridge';

test('SgwChatBridge authenticates, encrypts session on disk, syncs rooms/DMs, and handles 401 cleanup', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-sgw-chat-'));
  let rejectNextWith401 = false;
  const savedRoomMessages: Array<Record<string, unknown>> = [];
  const savedDms: Array<Record<string, unknown>> = [];
  let capturedFileData: string | undefined;
  let readMarkedRoomId = '';
  let readMarkedDmOtherId = '';
  const ledgerIngested: Array<Record<string, unknown>> = [];
  let weeklyPushedDate = '';

  const server = http.createServer(async (req, res) => {
    const reqUrl = new URL(req.url || '/', 'http://127.0.0.1');
    const authHeader = req.headers.authorization || '';
    const rawChunks: Buffer[] = [];
    for await (const chunk of req) {
      rawChunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    const bodyText = Buffer.concat(rawChunks).toString('utf8');
    const body = bodyText ? (JSON.parse(bodyText) as Record<string, unknown>) : {};

    res.setHeader('Content-Type', 'application/json');

    if (reqUrl.pathname === '/api/auth/login' && req.method === 'POST') {
      if (body.email === 'kim@seastargo.com' && body.password === 'valid-pass') {
        res.end(
          JSON.stringify({
            ok: true,
            token: 'jwt_test_secret_token_value',
            user: {
              id: 'u-10203',
              name: '김봉정',
              email: 'kim@seastargo.com',
              role: 'manager',
              department: '전장설계부',
              position: '부장',
            },
          })
        );
        return;
      }
      res.statusCode = 401;
      res.end(JSON.stringify({ ok: false, error: 'invalid_credentials' }));
      return;
    }

    if (reqUrl.pathname === '/api/auth/me' && req.method === 'GET') {
      const cookieHdr = String(req.headers.cookie || '');
      if (authHeader === 'Bearer jwt_test_secret_token_value' || cookieHdr.includes('sgw_sess=ok')) {
        res.end(
          JSON.stringify({
            sub: 'u-10203',
            name: '김봉정',
            email: 'kim@seastargo.com',
            role: 'manager',
            department: '전장설계부',
            position: '부장',
          })
        );
        return;
      }
      res.statusCode = 401;
      res.end(JSON.stringify({ ok: false, error: 'auth_required' }));
      return;
    }

    if (rejectNextWith401 || authHeader !== 'Bearer jwt_test_secret_token_value') {
      res.statusCode = 401;
      res.end(JSON.stringify({ ok: false, error: 'auth_required' }));
      return;
    }

    if (reqUrl.pathname === '/api/employees' && req.method === 'GET') {
      res.end(
        JSON.stringify({
          employees: [
            {
              id: 'u-10203',
              name: '김봉정',
              email: 'kim@seastargo.com',
              role: 'manager',
              department: '전장설계부',
              position: '부장',
              status: 'active',
            },
            {
              id: 'u-10204',
              name: '이선우',
              email: 'lee@seastargo.com',
              role: 'user',
              department: '배관설계부',
              position: '부장',
              status: 'active',
            },
          ],
        })
      );
      return;
    }

    if (reqUrl.pathname === '/api/chat' && req.method === 'GET') {
      const roomId = reqUrl.searchParams.get('roomId');
      if (roomId) {
        res.end(JSON.stringify({ messages: savedRoomMessages }));
        return;
      }
      res.end(
        JSON.stringify({
          rooms: [
            {
              id: 'room-1',
              participants: ['u-10203', 'u-10204'],
              lastMessage: '도면 확인 부탁',
              lastMessageTime: 1795308000000,
              unreadCount: 2,
            },
          ],
        })
      );
      return;
    }

    if (reqUrl.pathname === '/api/chat' && req.method === 'POST') {
      const action = reqUrl.searchParams.get('action');
      if (action === 'room') {
        res.end(JSON.stringify({ ok: true, id: String(body.id || 'room-new') }));
        return;
      }
      if (action === 'read') {
        readMarkedRoomId = String(body.roomId || '');
        res.end(JSON.stringify({ ok: true }));
        return;
      }
      if (typeof body.fileData === 'string') {
        capturedFileData = body.fileData;
      }
      savedRoomMessages.push({
        id: body.id,
        roomId: body.roomId,
        senderId: 'u-10203',
        content: body.content,
        timestamp: 1795308100000,
        readBy: ['u-10203'],
        hasFile: Boolean(body.fileData),
        fileName: body.fileName,
      });
      res.end(JSON.stringify({ ok: true, id: String(body.id) }));
      return;
    }

    if (reqUrl.pathname === '/api/messages' && req.method === 'GET') {
      res.end(JSON.stringify({ messages: savedDms, serverNow: 1795308200000 }));
      return;
    }

    if (reqUrl.pathname === '/api/messages' && req.method === 'POST') {
      const action = reqUrl.searchParams.get('action');
      if (action === 'read') {
        readMarkedDmOtherId = reqUrl.searchParams.get('otherId') || '';
        res.end(JSON.stringify({ ok: true }));
        return;
      }
      savedDms.push({
        id: body.id,
        senderId: 'u-10203',
        senderName: '김봉정',
        receiverId: body.receiverId,
        receiverName: body.receiverName || '',
        content: body.content,
        timestamp: 1795308300000,
        isRead: false,
        hasFile: Boolean(body.fileData),
        fileName: body.fileName,
      });
      res.end(JSON.stringify({ ok: true, id: String(body.id) }));
      return;
    }

    if (reqUrl.pathname === '/api/bmail-ledger' && req.method === 'GET') {
      res.end(
        JSON.stringify({
          ok: true,
          summary: { totalMails: ledgerIngested.length, sentMails: 1, recvMails: 0, totalFiles: 1 },
          mails: ledgerIngested,
        })
      );
      return;
    }

    if (reqUrl.pathname === '/api/bmail-ledger' && req.method === 'POST') {
      if (body.action === 'ingest' && Array.isArray(body.items)) {
        for (const item of body.items) {
          ledgerIngested.push(item as Record<string, unknown>);
        }
        res.statusCode = 201;
        res.end(
          JSON.stringify({
            ok: true,
            inserted: body.items.length,
            touchedDates: ['2026-11-22'],
          })
        );
        return;
      }
      if (body.action === 'push-to-weekly') {
        weeklyPushedDate = String(body.log_date || '');
        res.end(JSON.stringify({ ok: true, week: '2026-W47', project: 'H2710' }));
        return;
      }
    }

    res.statusCode = 404;
    res.end(JSON.stringify({ error: 'not_found' }));
  });

  await new Promise<void>((resolve) => {
    const tryListen = () => {
      server.listen(0, '127.0.0.1', () => {
        const a = server.address();
        const p = a && typeof a === 'object' ? a.port : 0;
        // Node fetch (Undici) blocks WHATWG unsafe ports (e.g. 6665-6669 IRC)
        if ([6665, 6666, 6667, 6668, 6669, 6697, 10080].includes(p)) {
          server.close(() => tryListen());
          return;
        }
        resolve();
      });
    };
    tryListen();
  });
  const addr = server.address();
  const port = addr && typeof addr === 'object' ? addr.port : 0;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const bridge = new SgwChatBridge(tmpDir, baseUrl);
    assert.equal(bridge.getStatus().connected, false);

    // Wrong password fails cleanly without storing session
    await assert.rejects(
      () => bridge.login({ email: 'kim@seastargo.com', password: 'wrong' }),
      (err: unknown) => err instanceof SgwBridgeError && err.status === 401
    );
    assert.equal(bridge.getStatus().connected, false);

    // Valid login stores encrypted session without plaintext JWT or password on disk
    const status = await bridge.login({ email: 'kim@seastargo.com', password: 'valid-pass' });
    assert.equal(status.connected, true);
    assert.equal(status.user?.sub, 'u-10203');

    const encRaw = fs.readFileSync(path.join(tmpDir, 'sgw-chat-session.enc'), 'utf8');
    assert.equal(encRaw.includes('jwt_test_secret_token_value'), false);
    assert.equal(encRaw.includes('valid-pass'), false);

    // Reload from disk with a new SgwChatBridge instance
    const bridgeReloaded = new SgwChatBridge(tmpDir, baseUrl);
    assert.equal(bridgeReloaded.getStatus().user?.email, 'kim@seastargo.com');

    // Fetch rooms & participant profiles
    const { rooms, unreadTotal } = await bridgeReloaded.listRooms();
    assert.equal(rooms.length, 1);
    assert.equal(unreadTotal, 2);
    assert.equal(rooms[0].participantProfiles?.length, 2);

    // Share mail to room + mark read
    const shareText = buildMailShareText({
      id: 'pst-abc-123',
      subject: '[H2710] Switchboard Drawing Rev.2',
      fromName: '박선급',
      fromEmail: 'park@krs.co.kr',
      date: '2026-11-22',
      snippet: '검토 의견서 회신드립니다.',
    });
    assert.ok(shareText.includes('bmail://open?id=pst-abc-123'));
    const sentMsg = await bridgeReloaded.sendRoomMessage('room-1', {
      content: shareText,
      fileName: 'note.txt',
      fileData: `data:text/plain;base64,${Buffer.from('hello').toString('base64')}`,
    });
    assert.equal(sentMsg.ok, true);
    assert.equal(capturedFileData, Buffer.from('hello').toString('base64'));
    const roomMsgs = await bridgeReloaded.getRoomMessages('room-1');
    assert.equal(roomMsgs.messages.length, 1);
    assert.equal(roomMsgs.messages[0].hasFile, true);

    await bridgeReloaded.markRoomRead('room-1', sentMsg.id);
    assert.equal(readMarkedRoomId, 'room-1');

    // Direct message roundtrip
    const dmRes = await bridgeReloaded.sendDirectMessage({
      receiverId: 'u-10204',
      receiverName: '이선우',
      content: '확인 부탁드립니다.',
    });
    assert.equal(dmRes.ok, true);
    const dmList = await bridgeReloaded.listDirectMessages();
    assert.equal(dmList.messages.length, 1);
    await bridgeReloaded.markDmRead('u-10204');
    assert.equal(readMarkedDmOtherId, 'u-10204');

    // Mail ledger sync + dashboard + weekly report push
    const syncRes = await bridgeReloaded.syncEmailsToLedger([
      {
        id: 'pst-ledger-1',
        folder: 'sent',
        subject: '[H2710] E1001 Rev.2 승인도 출도',
        fromName: '김봉정',
        fromEmail: 'kim@seastargo.com',
        to: ['approval@yard.test'],
        date: '2026-11-22',
        timestamp: '2026-11-22T09:00:00.000Z',
        timeLabel: '09:00',
        snippet: '도면 송부드립니다.',
        body: '첨부 도면 확인 부탁드립니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'att-ledger-1',
            filename: 'H2710-E1001_Rev2.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 12000,
            kind: 'pdf',
          },
        ],
        isRead: true,
        isFlagged: false,
        source: 'pst-import',
      },
    ]);
    assert.equal(syncRes.ok, true);
    assert.equal(syncRes.inserted, 1);
    const ledgerDash = await bridgeReloaded.getLedgerDashboard({ project: 'H2710' });
    assert.equal((ledgerDash.summary as { totalMails: number }).totalMails, 1);
    const weeklyRes = await bridgeReloaded.pushDailyLogToWeekly('2026-11-22');
    assert.equal(weeklyRes.ok, true);
    assert.equal(weeklyPushedDate, '2026-11-22');

    // Attachment > 7MB base64 rejected locally before HTTP call
    assert.throws(
      () => normalizeBase64Payload('A'.repeat(7 * 1024 * 1024 + 10)),
      (err: unknown) => err instanceof SgwBridgeError && err.status === 413
    );

    // 401 session expiry triggers automatic local cleanup
    rejectNextWith401 = true;
    await assert.rejects(
      () => bridgeReloaded.listRooms(),
      (err: unknown) => err instanceof SgwBridgeError && err.status === 401
    );
    assert.equal(bridgeReloaded.getStatus().connected, false);

    // Interactive SGW login-window session binding via Bearer token + cookieHeader (/api/auth/me)
    rejectNextWith401 = false;
    const sessionConnected = await bridgeReloaded.connectWithSession({
      sgwBaseUrl: baseUrl,
      token: 'jwt_test_secret_token_value',
      cookieHeader: 'sgw_sess=ok',
    });
    assert.equal(sessionConnected.connected, true);
    assert.equal(sessionConnected.user?.sub, 'u-10203');
    const bridgeReloadedSession = new SgwChatBridge(tmpDir, baseUrl);
    assert.equal(bridgeReloadedSession.getStatus().connected, true);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
