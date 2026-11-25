import fs from 'fs';
import path from 'path';
import zlib from 'zlib';

/**
 * Directories strictly excluded from full-source code downloads
 * (runtime secrets, encrypted EMV vault, dependencies, build output, git/test artifacts).
 */
const EXCLUDED_DIR_NAMES = new Set([
  'node_modules',
  'dist',
  'dist-server',
  '.git',
  'data',
  '.playwright-cli',
  '.aistudio',
  '.wrangler',
  'release',
  'release-exe',
  'out',
  'tmp',
]);

/**
 * Sensitive file extensions never allowed inside a downloadable source archive.
 */
const EXCLUDED_EXTENSIONS = new Set([
  '.pem',
  '.key',
  '.p12',
  '.pfx',
  '.pst',
  '.ost',
  '.db',
  '.sqlite',
  '.sqlite3',
  '.log',
]);

export interface SourceArchiveBuildResult {
  buffer: Buffer;
  fileCount: number;
  uncompressedBytes: number;
  includedPaths: string[];
}

export function shouldExcludeArchivePath(relativePosixPath: string, isDirectory: boolean): boolean {
  const normalized = relativePosixPath.replace(/\\/g, '/').replace(/^\/+/, '');
  if (!normalized) return true;

  const segments = normalized.split('/');
  const baseName = segments[segments.length - 1];

  // Never allow path traversal segments
  if (segments.some((s) => s === '..' || s === '.')) {
    return true;
  }

  // Exclude sensitive/generated directories at any depth
  if (isDirectory && EXCLUDED_DIR_NAMES.has(baseName)) {
    return true;
  }
  if (segments.slice(0, -1).some((s) => EXCLUDED_DIR_NAMES.has(s))) {
    return true;
  }

  if (isDirectory) return false;

  // Exclude all .env files EXCEPT .env.example
  if (baseName === '.env' || (baseName.startsWith('.env.') && baseName !== '.env.example')) {
    return true;
  }

  // Exclude encrypted vaults or private credential JSON dumps
  if (
    baseName.includes('emv-vault') ||
    baseName.endsWith('.enc.json') ||
    baseName === '.DS_Store' ||
    baseName === 'Thumbs.db'
  ) {
    return true;
  }

  const ext = path.extname(baseName).toLowerCase();
  if (EXCLUDED_EXTENSIONS.has(ext)) {
    return true;
  }

  return false;
}

function writeOctalField(target: Buffer, offset: number, length: number, value: number): void {
  const safeVal = Math.max(0, Math.floor(value));
  const octal = safeVal.toString(8).padStart(length - 1, '0');
  target.write(octal.slice(-(length - 1)), offset, length - 1, 'ascii');
  target[offset + length - 1] = 0;
}

/**
 * Encodes a single file entry into a POSIX ustar 512-byte header + padded data blocks.
 * Uses pure Node.js Buffer operations with zero external dependencies or shell commands.
 */
function buildUstarEntry(entryPath: string, content: Buffer, mtimeSeconds: number): Buffer {
  const header = Buffer.alloc(512, 0);
  const fullPathBytes = Buffer.from(entryPath, 'utf8');

  let nameField = entryPath;
  let prefixField = '';

  if (fullPathBytes.length > 100) {
    // Split on '/' so prefix <= 155 bytes and name <= 100 bytes
    const slashIdx = entryPath.lastIndexOf('/');
    if (slashIdx > 0) {
      const candidatePrefix = entryPath.slice(0, slashIdx);
      const candidateName = entryPath.slice(slashIdx + 1);
      if (
        Buffer.byteLength(candidatePrefix, 'utf8') <= 155 &&
        Buffer.byteLength(candidateName, 'utf8') <= 100
      ) {
        prefixField = candidatePrefix;
        nameField = candidateName;
      }
    }
  }

  header.write(nameField, 0, Math.min(100, Buffer.byteLength(nameField, 'utf8')), 'utf8');
  writeOctalField(header, 100, 8, 0o644); // mode
  writeOctalField(header, 108, 8, 0); // uid
  writeOctalField(header, 116, 8, 0); // gid
  writeOctalField(header, 124, 12, content.length); // size
  writeOctalField(header, 136, 12, mtimeSeconds); // mtime
  // Checksum field (148..155) filled with ASCII spaces during calculation
  header.fill(0x20, 148, 156);
  header[156] = 0x30; // '0' regular file
  header.write('ustar\0', 257, 6, 'ascii');
  header.write('00', 263, 2, 'ascii');
  if (prefixField) {
    header.write(prefixField, 345, Math.min(155, Buffer.byteLength(prefixField, 'utf8')), 'utf8');
  }

  let checksum = 0;
  for (let i = 0; i < 512; i += 1) {
    checksum += header[i];
  }
  const chksumStr = checksum.toString(8).padStart(6, '0');
  header.write(chksumStr, 148, 6, 'ascii');
  header[154] = 0;
  header[155] = 0x20;

  const remainder = content.length % 512;
  const padLength = remainder === 0 ? 0 : 512 - remainder;
  const padding = padLength > 0 ? Buffer.alloc(padLength, 0) : Buffer.alloc(0);

  return Buffer.concat([header, content, padding]);
}

interface CollectedSourceFile {
  relPosix: string;
  content: Buffer;
  mtimeMs: number;
}

function collectSafeSourceFiles(rootDir: string): CollectedSourceFile[] {
  const resolvedRoot = path.resolve(rootDir);
  const files: CollectedSourceFile[] = [];

  const walk = (currentDir: string) => {
    const entries = fs.readdirSync(currentDir, { withFileTypes: true });
    entries.sort((a, b) => a.name.localeCompare(b.name));

    for (const entry of entries) {
      // Skip symlinks to prevent escaping the project root
      if (entry.isSymbolicLink()) continue;

      const absPath = path.resolve(currentDir, entry.name);
      if (!absPath.startsWith(resolvedRoot + path.sep)) {
        continue;
      }

      const relPosix = path.relative(resolvedRoot, absPath).split(path.sep).join('/');
      if (shouldExcludeArchivePath(relPosix, entry.isDirectory())) {
        continue;
      }

      if (entry.isDirectory()) {
        walk(absPath);
      } else if (entry.isFile()) {
        const stat = fs.statSync(absPath);
        // Guard against unexpected multi-megabyte binary blobs (> 15 MB)
        if (stat.size > 15 * 1024 * 1024) continue;
        files.push({
          relPosix,
          content: fs.readFileSync(absPath),
          mtimeMs: stat.mtimeMs,
        });
      }
    }
  };

  walk(resolvedRoot);
  return files;
}

/**
 * Safely collects all project source files under rootDir (confining every path inside rootDir)
 * and builds a gzip-compressed `.tar.gz` archive excluding secrets and build artifacts.
 */
export function buildSafeSourceArchive(
  rootDir: string,
  archivePrefix = 'bmail-outlook-source'
): SourceArchiveBuildResult {
  const files = collectSafeSourceFiles(rootDir);
  const includedPaths: string[] = [];
  const chunks: Buffer[] = [];
  let uncompressedBytes = 0;

  for (const file of files) {
    const tarPath = `${archivePrefix}/${file.relPosix}`;
    const mtimeSec = Math.floor(file.mtimeMs / 1000);
    chunks.push(buildUstarEntry(tarPath, file.content, mtimeSec));
    includedPaths.push(file.relPosix);
    uncompressedBytes += file.content.length;
  }

  // Two 512-byte zero blocks terminate a POSIX tar archive
  chunks.push(Buffer.alloc(1024, 0));

  const tarBuffer = Buffer.concat(chunks);
  const gzBuffer = zlib.gzipSync(tarBuffer, { level: 6 });

  return {
    buffer: gzBuffer,
    fileCount: includedPaths.length,
    uncompressedBytes,
    includedPaths,
  };
}

// Precomputed standard IEEE 802.3 CRC-32 lookup table for PKZIP archives
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) !== 0 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function computeCrc32(buf: Buffer): number {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function toDosDateTime(mtimeMs: number): { dosTime: number; dosDate: number } {
  const d = new Date(mtimeMs);
  const year = Math.max(1980, Math.min(2107, d.getFullYear()));
  const month = d.getMonth() + 1;
  const day = d.getDate();
  const hours = d.getHours();
  const minutes = d.getMinutes();
  const seconds = Math.floor(d.getSeconds() / 2);
  const dosTime = ((hours & 0x1f) << 11) | ((minutes & 0x3f) << 5) | (seconds & 0x1f);
  const dosDate = (((year - 1980) & 0x7f) << 9) | ((month & 0x0f) << 5) | (day & 0x1f);
  return { dosTime, dosDate };
}

/**
 * Pure Node.js PKZIP (`.zip`) source archive builder with UTF-8 filename support
 * and zero external dependencies — ideal for 1-click extraction on Windows, Android, and macOS.
 */
export function buildSafeSourceZipArchive(
  rootDir: string,
  archivePrefix = 'bmail-outlook-source'
): SourceArchiveBuildResult {
  const files = collectSafeSourceFiles(rootDir);
  const localParts: Buffer[] = [];
  const centralParts: Buffer[] = [];
  const includedPaths: string[] = [];
  let uncompressedBytes = 0;
  let offset = 0;

  for (const file of files) {
    const entryName = `${archivePrefix}/${file.relPosix}`;
    const nameBuf = Buffer.from(entryName, 'utf8');
    const crc = computeCrc32(file.content);
    const compressed = zlib.deflateRawSync(file.content, { level: 6 });
    const useDeflate = compressed.length < file.content.length;
    const payload = useDeflate ? compressed : file.content;
    const method = useDeflate ? 8 : 0;
    const { dosTime, dosDate } = toDosDateTime(file.mtimeMs);
    const utf8Flag = 0x0800;

    // Local file header (30 bytes + filename)
    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4); // version needed: 2.0
    localHeader.writeUInt16LE(utf8Flag, 6);
    localHeader.writeUInt16LE(method, 8);
    localHeader.writeUInt16LE(dosTime, 10);
    localHeader.writeUInt16LE(dosDate, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(payload.length, 18);
    localHeader.writeUInt32LE(file.content.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28); // extra len

    localParts.push(localHeader, nameBuf, payload);

    // Central directory header (46 bytes + filename)
    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4); // version made by
    centralHeader.writeUInt16LE(20, 6); // version needed
    centralHeader.writeUInt16LE(utf8Flag, 8);
    centralHeader.writeUInt16LE(method, 10);
    centralHeader.writeUInt16LE(dosTime, 12);
    centralHeader.writeUInt16LE(dosDate, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(payload.length, 20);
    centralHeader.writeUInt32LE(file.content.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30); // extra len
    centralHeader.writeUInt16LE(0, 32); // comment len
    centralHeader.writeUInt16LE(0, 34); // disk start
    centralHeader.writeUInt16LE(0, 36); // internal attrs
    centralHeader.writeUInt32LE(0, 38); // external attrs
    centralHeader.writeUInt32LE(offset, 42);

    centralParts.push(centralHeader, nameBuf);

    offset += localHeader.length + nameBuf.length + payload.length;
    includedPaths.push(file.relPosix);
    uncompressedBytes += file.content.length;
  }

  const centralDirBuffer = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4); // disk number
  eocd.writeUInt16LE(0, 6); // central dir start disk
  eocd.writeUInt16LE(includedPaths.length, 8);
  eocd.writeUInt16LE(includedPaths.length, 10);
  eocd.writeUInt32LE(centralDirBuffer.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20); // comment length

  return {
    buffer: Buffer.concat([...localParts, centralDirBuffer, eocd]),
    fileCount: includedPaths.length,
    uncompressedBytes,
    includedPaths,
  };
}

