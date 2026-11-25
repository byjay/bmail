import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'fs';
import os from 'os';
import path from 'path';
import zlib from 'zlib';
import {
  buildSafeSourceArchive,
  buildSafeSourceZipArchive,
  shouldExcludeArchivePath,
} from './sourceArchiveBuilder';

test('shouldExcludeArchivePath: blocks secrets, vault files, node_modules, dist, and .git while allowing .env.example', () => {
  assert.equal(shouldExcludeArchivePath('.env', false), true);
  assert.equal(shouldExcludeArchivePath('.env.local', false), true);
  assert.equal(shouldExcludeArchivePath('.env.production', false), true);
  assert.equal(shouldExcludeArchivePath('.env.example', false), false);

  assert.equal(shouldExcludeArchivePath('data', true), true);
  assert.equal(shouldExcludeArchivePath('data/emv-vault.enc.json', false), true);
  assert.equal(shouldExcludeArchivePath('node_modules', true), true);
  assert.equal(shouldExcludeArchivePath('node_modules/d3/index.js', false), true);
  assert.equal(shouldExcludeArchivePath('dist/index.html', false), true);
  assert.equal(shouldExcludeArchivePath('.git/config', false), true);
  assert.equal(shouldExcludeArchivePath('certs/server.key', false), true);
  assert.equal(shouldExcludeArchivePath('backup/archive.pst', false), true);
  assert.equal(shouldExcludeArchivePath('../outside.txt', false), true);

  assert.equal(shouldExcludeArchivePath('src/App.tsx', false), false);
  assert.equal(shouldExcludeArchivePath('HANDOFF.md', false), false);
  assert.equal(shouldExcludeArchivePath('SETUP_MANUAL.md', false), false);
  assert.equal(shouldExcludeArchivePath('wrangler.toml', false), false);
  assert.equal(shouldExcludeArchivePath('cloudflare/worker.ts', false), false);
});

test('buildSafeSourceArchive: builds valid gzipped tar archive excluding secret files', () => {
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-archive-test-'));
  try {
    fs.mkdirSync(path.join(tmpRoot, 'src'), { recursive: true });
    fs.mkdirSync(path.join(tmpRoot, 'data'), { recursive: true });
    fs.mkdirSync(path.join(tmpRoot, 'node_modules', 'pkg'), { recursive: true });

    fs.writeFileSync(path.join(tmpRoot, 'src', 'index.ts'), 'export const ok = true;\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, '.env.example'), 'GEMINI_API_KEY=\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, '.env'), 'GEMINI_API_KEY=SECRET_DO_NOT_EXPORT\n', 'utf8');
    fs.writeFileSync(
      path.join(tmpRoot, 'data', 'emv-vault.enc.json'),
      '{"secret":"NEVER"}\n',
      'utf8'
    );
    fs.writeFileSync(path.join(tmpRoot, 'node_modules', 'pkg', 'a.js'), 'module.exports={};', 'utf8');

    const result = buildSafeSourceArchive(tmpRoot, 'bmail-test');
    assert.equal(result.fileCount, 2);
    assert.deepEqual(result.includedPaths, ['.env.example', 'src/index.ts']);

    const rawTar = zlib.gunzipSync(result.buffer).toString('utf8');
    assert.ok(rawTar.includes('bmail-test/src/index.ts'));
    assert.ok(rawTar.includes('bmail-test/.env.example'));
    assert.equal(rawTar.includes('SECRET_DO_NOT_EXPORT'), false);
    assert.equal(rawTar.includes('emv-vault.enc.json'), false);
    assert.equal(rawTar.includes('node_modules'), false);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});

test('buildSafeSourceZipArchive: builds valid PKZIP archive including SETUP_MANUAL.md and excluding secrets', () => {
  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'bmail-zip-test-'));
  try {
    fs.mkdirSync(path.join(tmpRoot, 'src'), { recursive: true });
    fs.mkdirSync(path.join(tmpRoot, 'data'), { recursive: true });
    fs.mkdirSync(path.join(tmpRoot, 'node_modules', 'pkg'), { recursive: true });

    fs.writeFileSync(path.join(tmpRoot, 'SETUP_MANUAL.md'), '# 설정 매뉴얼\n내용 테스트\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, 'src', 'index.ts'), 'export const zipOk = true;\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, '.env.example'), 'GEMINI_API_KEY=\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, '.env'), 'GEMINI_API_KEY=ZIP_SECRET_NEVER\n', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, 'data', 'emv-vault.enc.json'), '{"secret":"NO"}', 'utf8');
    fs.writeFileSync(path.join(tmpRoot, 'node_modules', 'pkg', 'x.js'), 'module.exports={};', 'utf8');

    const result = buildSafeSourceZipArchive(tmpRoot, 'bmail-zip');
    assert.equal(result.fileCount, 3);
    assert.deepEqual(result.includedPaths, ['.env.example', 'SETUP_MANUAL.md', 'src/index.ts']);

    // Verify PKZIP local file header (PK\x03\x04) and EOCD signature (PK\x05\x06)
    assert.equal(result.buffer.readUInt32LE(0), 0x04034b50);
    assert.ok(result.buffer.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])) > 0);

    const utf8Scan = result.buffer.toString('utf8');
    assert.ok(utf8Scan.includes('bmail-zip/SETUP_MANUAL.md'));
    assert.ok(utf8Scan.includes('bmail-zip/src/index.ts'));
    assert.ok(utf8Scan.includes('bmail-zip/.env.example'));
    assert.equal(utf8Scan.includes('ZIP_SECRET_NEVER'), false);
    assert.equal(utf8Scan.includes('emv-vault.enc.json'), false);
    assert.equal(utf8Scan.includes('node_modules'), false);
  } finally {
    fs.rmSync(tmpRoot, { recursive: true, force: true });
  }
});
