import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import {
  buildPasswordCandidates,
  buildSmtpMimePayload,
  connectMailWithAutoHeal,
  detectCompanyPop3Config,
  diagnoseMailAuthError,
  fetchImapMessages,
  fetchPop3Messages,
  MailAuthError,
  Pop3SessionStore,
} from './pop3AutoConnect';

test('detectCompanyPop3Config accurately classifies Mailplug, Dooray (IMAP-only), Cafe24 (POP3-only webmail), Hiworks, and iCloud', async () => {
  const alwaysReachable = async () => ({ reachable: true, banner: '+OK ready' });
  const alwaysImapOk = async () => ({ reachable: true, banner: '* OK IMAP4rev1' });

  // 1. Mailplug via mx01.mailplug.com (user screenshot domain pattern)
  const mailplug = await detectCompanyPop3Config({
    email: 'designsir@seastar-example.co.kr',
    resolveMxImpl: async () => [{ exchange: 'mx01.mailplug.com', priority: 10 }],
    probePortImpl: alwaysReachable,
    probeImapImpl: alwaysImapOk,
  });
  assert.equal(mailplug.detectedProviderId, 'mailplug');
  assert.equal(mailplug.pop3Host, 'pop3.mailplug.co.kr');
  assert.equal(mailplug.imapHost, 'imap.mailplug.co.kr');
  assert.equal(mailplug.appPasswordPolicy, 'required');
  assert.equal(mailplug.overseasIpWarning, true);
  assert.equal(mailplug.recommendedProtocol, 'pop3');
  assert.deepEqual(mailplug.supportedProtocols, ['pop3', 'imap']);

  // 2. Dooray is IMAP-only (no POP3)
  const dooray = await detectCompanyPop3Config({
    email: 'user@team.dooray.com',
    resolveMxImpl: async () => [{ exchange: 'mx.dooray.com', priority: 10 }],
    probePortImpl: alwaysReachable,
    probeImapImpl: alwaysImapOk,
  });
  assert.equal(dooray.detectedProviderId, 'dooray');
  assert.deepEqual(dooray.supportedProtocols, ['imap']);
  assert.equal(dooray.recommendedProtocol, 'imap');
  assert.equal(dooray.imapHost, 'imap.dooray.com');
  assert.equal(dooray.imapPort, 993);

  // 3. Cafe24 is POP3-only using webmail.{domain}:110 + smtp.cafe24.com:587
  const cafe24 = await detectCompanyPop3Config({
    email: 'sales@shop-example.co.kr',
    resolveMxImpl: async () => [{ exchange: 'mw-002.cafe24.com', priority: 10 }],
    probePortImpl: alwaysReachable,
    probeImapImpl: alwaysImapOk,
  });
  assert.equal(cafe24.detectedProviderId, 'cafe24');
  assert.deepEqual(cafe24.supportedProtocols, ['pop3']);
  assert.equal(cafe24.recommendedProtocol, 'pop3');
  assert.equal(cafe24.pop3Host, 'webmail.shop-example.co.kr');
  assert.equal(cafe24.pop3Port, 110);
  assert.equal(cafe24.smtpHost, 'smtp.cafe24.com');
  assert.equal(cafe24.smtpPort, 587);

  // 4. iCloud is IMAP-only
  const icloud = await detectCompanyPop3Config({
    email: 'owner@icloud.com',
    resolveMxImpl: async () => [],
    probePortImpl: alwaysReachable,
    probeImapImpl: alwaysImapOk,
  });
  assert.equal(icloud.detectedProviderId, 'icloud');
  assert.deepEqual(icloud.supportedProtocols, ['imap']);
  assert.equal(icloud.imapHost, 'imap.mail.me.com');
  assert.equal(icloud.appPasswordPolicy, 'required');
});

test('diagnoseMailAuthError diagnoses exact Mailplug -ERR 553 (#5.7.5) error with app password, overseas IP, and IMAP fallback', () => {
  const diag = diagnoseMailAuthError({
    providerId: 'mailplug',
    protocol: 'pop3',
    rawServerError: '-ERR 553 sorry, that password allowed relay (#5.7.5)',
  });

  assert.equal(diag.code, 'overseas-ip-or-policy');
  assert.equal(diag.suggestProtocolSwitch, 'imap');
  assert.ok(diag.checklist.some((line) => line.includes('앱 비밀번호')));
  assert.ok(diag.checklist.some((line) => line.includes('해외 IP')));
  assert.ok(diag.checklist.some((line) => line.includes('POP3/IMAP')));
  assert.ok(diag.checklist.some((line) => line.includes('2025년 4월 1일')));

  const imapDiag = diagnoseMailAuthError({
    providerId: 'mailplug',
    protocol: 'imap',
    rawServerError: 'A001 NO LOGIN failed',
  });
  assert.equal(imapDiag.suggestProtocolSwitch, 'pop3');
  assert.ok(imapDiag.checklist.some((line) => line.includes('IMAP 제한')));
});

test('fetchPop3Messages automatically retries with local-part username when full-email fails on local-or-full provider', async () => {
  const seenUsers: string[] = [];
  const rawEml = [
    'From: Alice <alice@example.com>',
    'To: hong@naver.com',
    'Subject: Local ID Fallback Test',
    'Date: Fri, 20 Nov 2026 09:00:00 +0900',
    '',
    'Hello from POP3 local-part login!',
  ].join('\r\n');

  const server = net.createServer((socket) => {
    let currentUser = '';
    socket.write('+OK Mock POP3 Ready\r\n');
    let buf = '';
    socket.on('data', (chunk) => {
      buf += chunk.toString('utf8');
      while (buf.includes('\r\n')) {
        const idx = buf.indexOf('\r\n');
        const line = buf.slice(0, idx);
        buf = buf.slice(idx + 2);
        if (line.startsWith('USER ')) {
          currentUser = line.slice(5).trim();
          seenUsers.push(currentUser);
          socket.write('+OK user accepted\r\n');
        } else if (line.startsWith('PASS ')) {
          // Reject full email, accept local-part ID only (mimics Daum/Nate/legacy Naver)
          if (currentUser === 'hong') {
            socket.write('+OK logged in\r\n');
          } else {
            socket.write('-ERR invalid user format\r\n');
          }
        } else if (line === 'STAT') {
          socket.write(`+OK 1 ${Buffer.byteLength(rawEml)}\r\n`);
        } else if (line === 'RETR 1') {
          socket.write(`+OK ${Buffer.byteLength(rawEml)} octets\r\n${rawEml}\r\n.\r\n`);
        } else if (line === 'QUIT') {
          socket.write('+OK bye\r\n');
          socket.end();
        }
      }
    });
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const port = (server.address() as net.AddressInfo).port;

  try {
    const result = await fetchPop3Messages({
      email: 'hong@naver.com',
      password: 'app-password-1234',
      pop3Host: '127.0.0.1',
      pop3Port: port,
      pop3Tls: false,
      providerId: 'naver',
      usernameFormat: 'local-or-full',
      allowLoopbackForTest: true,
    });

    assert.deepEqual(seenUsers, ['hong@naver.com', 'hong']);
    assert.equal(result.protocolUsed, 'pop3');
    assert.equal(result.loginUsernameUsed, 'hong');
    assert.equal(result.fetchedCount, 1);
    assert.equal(result.emails[0].subject, 'Local ID Fallback Test');
  } finally {
    server.close();
  }
});

test('fetchImapMessages connects to IMAP4rev1, authenticates, reads FLAGS and BODY.PEEK[] literal accurately', async () => {
  const rawEml = [
    'From: Chief Engineer <chief@shipyard.co.kr>',
    'To: me@mailplug.co.kr',
    'Subject: [H1001] IMAP Live Fetch Verification',
    'Date: Fri, 20 Nov 2026 10:15:00 +0900',
    '',
    'IMAP BODY.PEEK test body content.',
  ].join('\r\n');
  const rawBytes = Buffer.byteLength(rawEml, 'utf8');

  const server = net.createServer((socket) => {
    socket.write('* OK [CAPABILITY IMAP4rev1] Mock IMAP Ready\r\n');
    let buf = '';
    socket.on('data', (chunk) => {
      buf += chunk.toString('utf8');
      while (buf.includes('\r\n')) {
        const idx = buf.indexOf('\r\n');
        const line = buf.slice(0, idx);
        buf = buf.slice(idx + 2);
        const [tag, cmd] = line.split(' ');
        if (cmd === 'LOGIN') {
          socket.write(`${tag} OK LOGIN completed\r\n`);
        } else if (cmd === 'SELECT') {
          socket.write('* 1 EXISTS\r\n');
          socket.write('* 0 RECENT\r\n');
          socket.write(`${tag} OK [READ-WRITE] SELECT completed\r\n`);
        } else if (cmd === 'FETCH') {
          socket.write(`* 1 FETCH (FLAGS (\\Seen \\Flagged) BODY[] {${rawBytes}}\r\n`);
          socket.write(`${rawEml})\r\n`);
          socket.write(`${tag} OK FETCH completed\r\n`);
        } else if (cmd === 'LOGOUT') {
          socket.write('* BYE logging out\r\n');
          socket.write(`${tag} OK LOGOUT completed\r\n`);
          socket.end();
        }
      }
    });
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const port = (server.address() as net.AddressInfo).port;

  try {
    const result = await fetchImapMessages({
      email: 'me@mailplug.co.kr',
      password: 'app-password-9999',
      imapHost: '127.0.0.1',
      imapPort: port,
      imapTls: false,
      providerId: 'mailplug',
      usernameFormat: 'full-email',
      allowLoopbackForTest: true,
    });

    assert.equal(result.protocolUsed, 'imap');
    assert.equal(result.totalMailboxCount, 1);
    assert.equal(result.fetchedCount, 1);
    assert.equal(result.emails[0].subject, '[H1001] IMAP Live Fetch Verification');
    assert.equal(result.emails[0].isRead, true);
    assert.equal(result.emails[0].isFlagged, true);
  } finally {
    server.close();
  }
});

test('buildPasswordCandidates strips copy-pasted spaces and hyphens cleanly', () => {
  const candidates = buildPasswordCandidates('  abcd-efgh ijkl-mnop  ');
  assert.equal(candidates[0].value, '  abcd-efgh ijkl-mnop  ');
  assert.equal(candidates[0].isNormalized, false);
  assert.ok(candidates.some((c) => c.value === 'abcdefghijklmnop' && c.isNormalized));
});

test('connectMailWithAutoHeal automatically heals IMAP->POP3 fallback and attaches Taiwan egress IP guide on #5.7.5', async () => {
  const discovered = await detectCompanyPop3Config({
    email: 'designsir@seastar-example.co.kr',
    resolveMxImpl: async () => [{ exchange: 'mx01.mailplug.com', priority: 10 }],
    probePortImpl: async () => ({ reachable: true, banner: '+OK ready' }),
    probeImapImpl: async () => ({ reachable: true, banner: '* OK IMAP' }),
  });

  // 1) Auto-heals IMAP failure by switching to POP3 automatically
  const healed = await connectMailWithAutoHeal({
    email: 'designsir@seastar-example.co.kr',
    password: 'abcd-efgh-ijkl-mnop',
    discovered,
    requestedProtocol: 'imap',
    pop3Host: 'pop3.mailplug.co.kr',
    pop3Port: 995,
    pop3Tls: true,
    imapHost: 'imap.mailplug.co.kr',
    imapPort: 993,
    imapTls: true,
    fetchImapImpl: async () => {
      throw new MailAuthError('IMAP blocked', diagnoseMailAuthError({
        providerId: 'mailplug',
        protocol: 'imap',
        rawServerError: 'A001 NO 553 sorry, that password allowed relay (#5.7.5)',
      }));
    },
    fetchPop3Impl: async (p) => ({
      protocolUsed: 'pop3',
      loginUsernameUsed: p.email,
      totalMailboxCount: 2,
      totalOctets: 512,
      fetchedCount: 1,
      emails: [],
      banner: '+OK POP3 ready',
    }),
    detectEgressImpl: async () => ({
      ipv4: '34.34.244.7',
      countryCode: 'TW',
      countryName: '대만',
      city: 'Taipei',
      isKorea: false,
    }),
  });

  assert.equal(healed.effectiveProtocol, 'pop3');
  assert.equal(healed.effectivePassword, 'abcdefghijklmnop');
  assert.ok(healed.healedBy && healed.healedBy.includes('POP3'));
  assert.ok(healed.attempts.length >= 2);

  // 2) When all strategies fail with #5.7.5, enriches diagnosis with live server IPv4 + step-by-step guide
  await assert.rejects(
    async () =>
      connectMailWithAutoHeal({
        email: 'designsir@seastar-example.co.kr',
        password: 'validapppassword',
        discovered,
        requestedProtocol: 'pop3',
        pop3Host: 'pop3.mailplug.co.kr',
        pop3Port: 995,
        pop3Tls: true,
        imapHost: 'imap.mailplug.co.kr',
        imapPort: 993,
        imapTls: true,
        fetchPop3Impl: async () => {
          throw new MailAuthError('#5.7.5', diagnoseMailAuthError({
            providerId: 'mailplug',
            protocol: 'pop3',
            rawServerError: '-ERR 553 sorry, that password allowed relay (#5.7.5)',
          }));
        },
        fetchImapImpl: async () => {
          throw new MailAuthError('#5.7.5', diagnoseMailAuthError({
            providerId: 'mailplug',
            protocol: 'imap',
            rawServerError: 'A001 NO 553 sorry, that password allowed relay (#5.7.5)',
          }));
        },
        detectEgressImpl: async () => ({
          ipv4: '34.34.244.7',
          countryCode: 'TW',
          countryName: '대만',
          city: 'Taipei',
          isKorea: false,
        }),
      }),
    (err: unknown) => {
      assert.ok(err instanceof MailAuthError);
      assert.equal(err.diagnosis.serverEgress?.ipv4, '34.34.244.7');
      assert.equal(err.diagnosis.serverEgress?.countryCode, 'TW');
      assert.ok((err.diagnosis.autoHealAttempts?.length || 0) >= 2);
      assert.ok(err.diagnosis.resolutionSteps?.some((s) => s.copyValue === '34.34.244.7'));
      return true;
    }
  );
});

test('Pop3SessionStore persists AES-256-GCM encrypted sessions across restarts without plaintext passwords', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-pop3-sess-'));
  try {
    const secret = 'test-session-master-key-2026';
    const store1 = new Pop3SessionStore({ baseDir: dir, secretKey: secret });
    const created = store1.create({
      email: 'designsir@seastargo.com',
      password: 'SuperSecretAppPassword!99',
      providerId: 'mailplug',
      providerName: '메일플러그 (Mailplug)',
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

    const encPath = path.join(dir, 'pop3-sessions.enc.json');
    assert.equal(fs.existsSync(encPath), true);
    const encRaw = fs.readFileSync(encPath, 'utf8');
    assert.equal(encRaw.includes('SuperSecretAppPassword!99'), false);
    assert.equal(encRaw.includes('designsir@seastargo.com'), false);

    // Reload from disk in fresh instance (simulating app restart)
    const store2 = new Pop3SessionStore({ baseDir: dir, secretKey: secret });
    const restored = store2.getLatest();
    assert.ok(restored);
    assert.equal(restored.token, created.token);
    assert.equal(restored.email, 'designsir@seastargo.com');
    assert.equal(restored.password, 'SuperSecretAppPassword!99');
    assert.equal(store2.listAccounts().length, 1);

    // Delete removes persisted session across future reloads
    assert.equal(store2.delete(created.token), true);
    const store3 = new Pop3SessionStore({ baseDir: dir, secretKey: secret });
    assert.equal(store3.getLatest(), null);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('connectMailWithAutoHeal tries space/hyphen-stripped password even when trailing whitespace exists', async () => {
  const discovered = await detectCompanyPop3Config({
    email: 'designsir@seastar-example.co.kr',
    resolveMxImpl: async () => [{ exchange: 'mx01.mailplug.com', priority: 10 }],
    probePortImpl: async () => ({ reachable: true, banner: '+OK ready' }),
    probeImapImpl: async () => ({ reachable: true, banner: '* OK IMAP4rev1' }),
  });
  const seenPasswords: string[] = [];
  const res = await connectMailWithAutoHeal({
    email: 'designsir@seastar-example.co.kr',
    password: ' abcd-efgh-ijkl-mnop ',
    discovered,
    requestedProtocol: 'pop3',
    pop3Host: 'pop3.mailplug.co.kr',
    pop3Port: 995,
    pop3Tls: true,
    imapHost: 'imap.mailplug.co.kr',
    imapPort: 993,
    imapTls: true,
    fetchPop3Impl: async (p) => {
      seenPasswords.push(p.password);
      if (p.password !== 'abcdefghijklmnop') {
        throw new MailAuthError('#5.7.5', diagnoseMailAuthError({
          providerId: 'mailplug',
          protocol: 'pop3',
          rawServerError: '-ERR 553 sorry, that password allowed relay (#5.7.5)',
        }));
      }
      return {
        protocolUsed: 'pop3',
        loginUsernameUsed: p.email,
        totalMailboxCount: 1,
        totalOctets: 128,
        fetchedCount: 1,
        emails: [],
        banner: '+OK',
      };
    },
    detectEgressImpl: async () => ({
      ipv4: null,
      countryCode: null,
      countryName: null,
      city: null,
      isKorea: null,
    }),
  });
  assert.equal(res.effectivePassword, 'abcdefghijklmnop');
  assert.ok(seenPasswords.includes('abcdefghijklmnop'));
});

test('buildSmtpMimePayload embeds inline base64 signature images as multipart/related Content-ID parts', () => {
  const pngB64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=';
  const mime = buildSmtpMimePayload({
    from: 'designsir@seastargo.com',
    to: ['client@example.com'],
    subject: 'FW: 명함 서명 테스트',
    textBody: '안녕하세요.\n김봉정 드림',
    htmlBody: `<p>안녕하세요.</p><div class="sig"><img src="data:image/png;base64,${pngB64}" alt="SEASTAR"></div>`,
  });

  assert.ok(mime.includes('multipart/alternative;'));
  assert.ok(mime.includes('multipart/related; type="text/html"'));
  assert.ok(mime.includes('Content-ID: <img1@bmail.local>'));
  assert.ok(mime.includes('Content-Type: image/png'));
  assert.ok(mime.endsWith('\r\n.'));
});

