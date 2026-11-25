#!/usr/bin/env node
/**
 * Publish Bmail Windows Installer + SHA-256 manifest to Cloudflare R2 (`bmail-releases`).
 * Usage: node scripts/publish-release.mjs [--notes "업데이트 설명"]
 */
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, '..');

const pkgPath = path.join(ROOT, 'package.json');
const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
const version = String(pkg.version || '1.0.0').trim();
const fileName = `Bmail-Outlook-AI-Setup-${version}.exe`;
const installerPath = path.join(ROOT, 'release-exe', fileName);

if (!fs.existsSync(installerPath)) {
  console.error(`[publish-release] Missing installer: ${installerPath}`);
  console.error('Run `npm run build:exe` first.');
  process.exit(1);
}

const notesArgIdx = process.argv.indexOf('--notes');
const releaseNotes =
  notesArgIdx !== -1 && process.argv[notesArgIdx + 1]
    ? process.argv[notesArgIdx + 1]
    : `Bmail Outlook AI v${version} — 1클릭 자동 업데이트·SGW 실시간 채팅 연동·경량 설치형`;

const stat = fs.statSync(installerPath);
const hash = crypto.createHash('sha256');
const fd = fs.openSync(installerPath, 'r');
try {
  const buf = Buffer.allocUnsafe(1024 * 1024);
  let bytesRead = 0;
  while ((bytesRead = fs.readSync(fd, buf, 0, buf.length, null)) > 0) {
    hash.update(buf.subarray(0, bytesRead));
  }
} finally {
  fs.closeSync(fd);
}
const sha256 = hash.digest('hex');

const baseUrl = (process.env.BMAIL_RELEASE_BASE_URL || 'https://mail.seastar.work/releases').replace(/\/+$/, '');
const manifest = {
  version,
  fileName,
  downloadUrl: `${baseUrl}/${fileName}`,
  latestDownloadUrl: `${baseUrl}/Bmail-Outlook-AI-Setup.exe`,
  sha256,
  sizeBytes: stat.size,
  publishedAt: new Date().toISOString(),
  releaseNotes,
};

const manifestPath = path.join(ROOT, 'release-exe', 'latest.json');
fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
console.log(`[publish-release] Generated manifest: ${manifestPath}`);
console.log(JSON.stringify(manifest, null, 2));

const putObject = (key, filePath, contentType) => {
  console.log(`[publish-release] Uploading -> bmail-releases/${key} (${fs.statSync(filePath).size} bytes)...`);
  execFileSync(
    'npx',
    [
      'wrangler',
      'r2',
      'object',
      'put',
      `bmail-releases/${key}`,
      '--file',
      filePath,
      '--content-type',
      contentType,
      '--remote',
    ],
    {
      cwd: ROOT,
      shell: true,
      stdio: 'inherit',
    }
  );
};

putObject(`releases/${fileName}`, installerPath, 'application/vnd.microsoft.portable-executable');
putObject('releases/Bmail-Outlook-AI-Setup.exe', installerPath, 'application/vnd.microsoft.portable-executable');

const blockmapPath = `${installerPath}.blockmap`;
if (fs.existsSync(blockmapPath)) {
  putObject(`releases/${fileName}.blockmap`, blockmapPath, 'application/octet-stream');
}

putObject('releases/latest.json', manifestPath, 'application/json');
console.log('[publish-release] Done! Live release ready at https://mail.seastar.work/releases/latest.json');
