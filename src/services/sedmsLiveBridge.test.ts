import assert from 'node:assert/strict';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { SedmsBridgeError, SedmsLiveBridge } from './sedmsLiveBridge';

test('SedmsLiveBridge encrypts token at rest, checks health, syncs contacts, and scans drawing refs', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-sedms-live-'));
  let capturedAuth = '';
  let previewCount = 0;
  let commitCount = 0;

  const server = http.createServer(async (req, res) => {
    const reqUrl = new URL(req.url || '/', 'http://127.0.0.1');
    const chunks: Buffer[] = [];
    for await (const c of req) chunks.push(Buffer.isBuffer(c) ? c : Buffer.from(c));
    const bodyText = Buffer.concat(chunks).toString('utf8');
    const body = bodyText ? (JSON.parse(bodyText) as Record<string, unknown>) : {};

    res.setHeader('Content-Type', 'application/json');

    if (reqUrl.pathname === '/api/health' && req.method === 'GET') {
      res.end(JSON.stringify({ ok: true }));
      return;
    }

    capturedAuth = String(req.headers.authorization || '');
    if (capturedAuth !== 'Bearer sedms_secret_token_999') {
      res.statusCode = 403;
      res.end(JSON.stringify({ ok: false, error: 'forbidden' }));
      return;
    }

    if (reqUrl.pathname === '/api/recipients/registry/import/preview' && req.method === 'POST') {
      const rows = Array.isArray(body.rows) ? body.rows : [];
      previewCount = rows.length;
      res.end(JSON.stringify({ ok: true, validCount: rows.length, duplicateCount: 0 }));
      return;
    }

    if (reqUrl.pathname === '/api/recipients/registry/import/commit' && req.method === 'POST') {
      const rows = Array.isArray(body.rows) ? body.rows : [];
      commitCount = rows.length;
      res.end(JSON.stringify({ ok: true, insertedCount: rows.length, skippedCount: 0, invalidCount: 0 }));
      return;
    }

    res.statusCode = 404;
    res.end(JSON.stringify({ ok: false, error: 'not_found' }));
  });

  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', () => resolve()));
  const addr = server.address();
  const port = addr && typeof addr === 'object' ? addr.port : 0;
  const baseUrl = `http://127.0.0.1:${port}`;

  try {
    const bridge = new SedmsLiveBridge(tmpDir, baseUrl);
    const health = await bridge.checkHealth();
    assert.equal(health.ok, true);
    assert.equal(health.status, 200);

    // Fails closed before token configured
    await assert.rejects(
      () =>
        bridge.previewContactsImport([
          { id: 'c1', fullName: '홍길동', email: 'hong@example.test', source: 'outlook-pc' },
        ]),
      (err: unknown) => err instanceof SedmsBridgeError && err.status === 401
    );

    // Stores encrypted token without plaintext leakage
    const status = bridge.saveConfig({
      sedmsBaseUrl: baseUrl,
      bearerToken: 'sedms_secret_token_999',
      tenantId: 'SEASTAR',
    });
    assert.equal(status.hasToken, true);
    assert.equal(status.tenantId, 'SEASTAR');

    const encRaw = fs.readFileSync(path.join(tmpDir, 'sedms-bridge.enc'), 'utf8');
    assert.equal(encRaw.includes('sedms_secret_token_999'), false);

    // Reload from encrypted disk store
    const reloaded = new SedmsLiveBridge(tmpDir);
    const contacts: Array<{
      id: string;
      fullName: string;
      email: string;
      company?: string;
      source: 'pst-derived';
    }> = [
      { id: 'c1', fullName: '박담당', email: 'park@yard.test', company: 'TestYard', source: 'pst-derived' },
      { id: 'c2', fullName: '최심사', email: 'choi@class.test', source: 'pst-derived' },
    ];
    const previewRes = await reloaded.previewContactsImport(contacts);
    assert.equal(previewRes.ok, true);
    assert.equal(previewCount, 2);

    const commitRes = await reloaded.commitContactsImport(contacts);
    assert.equal(commitRes.ok, true);
    assert.equal(commitCount, 2);
    assert.equal(capturedAuth, 'Bearer sedms_secret_token_999');

    // Deterministic offline email scan
    const scanReport = reloaded.scanLocalEmailsForSedms([
      {
        id: 'mail-scan-1',
        folder: 'sent',
        folderName: 'H2710',
        subject: '[H2710] DWG-E1001 Rev.2 전장 승인도 출도',
        fromName: '김설계',
        fromEmail: 'kim@seastargo.com',
        to: ['park@yard.test'],
        date: '2026-11-22',
        timestamp: '2026-11-22T10:00:00.000Z',
        timeLabel: '10:00',
        snippet: 'DWG-E1001 REV-2 출도 송부드립니다.',
        body: 'H2710 DWG-E1001 검토 요청합니다.',
        hasAttachment: true,
        attachments: [
          {
            id: 'att-1',
            filename: 'H2710-DWG-E1001_Rev2.pdf',
            mimeType: 'application/pdf',
            sizeBytes: 51200,
            kind: 'pdf',
          },
        ],
        isRead: true,
        isFlagged: false,
        source: 'pst-import',
      },
    ]);
    assert.equal(scanReport.totalScanned, 1);
    assert.equal(scanReport.drawingMatchedCount, 1);
    assert.equal(scanReport.autoLinkedCount, 1);
    assert.equal(scanReport.topProjects[0]?.projectCode, 'H2710');

    const cleared = reloaded.clearToken();
    assert.equal(cleared.hasToken, false);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
