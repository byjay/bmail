import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import zlib from 'node:zlib';
import { LocalBackupScheduler } from './localBackupScheduler';

test('LocalBackupScheduler creates compressed .sqlite.gz snapshots with SHA-256 and enforces retention', async () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-backup-test-'));
  const fakeDbPath = path.join(tmpDir, 'bmail-local.sqlite');
  const samplePayload = Buffer.from('SQLite format 3\0' + 'BMAIL-TEST-ROW-'.repeat(200), 'utf8');
  fs.writeFileSync(fakeDbPath, samplePayload);

  try {
    const scheduler = new LocalBackupScheduler(tmpDir, undefined, { maxRetention: 2 });
    const first = scheduler.runSnapshotNow('manual');
    assert.ok(first);
    assert.ok(fs.existsSync(first.fullPath));
    assert.ok(first.sha256.length === 64);

    // Verify gunzip round-trip recovers exact original bytes
    const decompressed = zlib.gunzipSync(fs.readFileSync(first.fullPath));
    assert.equal(decompressed.equals(samplePayload), true);

    await new Promise((r) => setTimeout(r, 1100));
    const second = scheduler.runSnapshotNow('manual');
    assert.ok(second);

    await new Promise((r) => setTimeout(r, 1100));
    const third = scheduler.runSnapshotNow('manual');
    assert.ok(third);

    // Retention = 2 -> oldest pruned cleanly
    const status = scheduler.getStatus();
    assert.equal(status.dbExists, true);
    assert.equal(status.snapshots.length, 2);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
});
