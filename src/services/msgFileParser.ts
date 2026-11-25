import crypto from 'node:crypto';
import fs from 'node:fs';
import type { AttachmentKind, MailAttachmentMeta, MailItem } from '../types/mail';
import { normalizeConversationSubject } from '../utils/deterministicAnalyzer';
import { stripHtmlToText } from '../utils/mailFileParsers';

const CFBF_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

interface CfbfDirectoryEntry {
  index: number;
  name: string;
  objectType: number; // 1 = storage, 2 = stream, 5 = root
  leftSiblingId: number;
  rightSiblingId: number;
  childId: number;
  startSector: number;
  size: number;
}

interface CfbfReader {
  entries: CfbfDirectoryEntry[];
  readEntryBuffer: (entry: CfbfDirectoryEntry) => Buffer;
}

function parseCfbf(buffer: Buffer): CfbfReader {
  if (buffer.length < 512 || !buffer.subarray(0, 8).equals(CFBF_MAGIC)) {
    throw new Error('유효한 Microsoft Outlook .msg (OLE2/CFBF) 바이너리 파일이 아닙니다.');
  }

  const minorVersion = buffer.readUInt16LE(24);
  const majorVersion = buffer.readUInt16LE(26);
  const sectorShift = buffer.readUInt16LE(30);
  const sectorSize = sectorShift === 12 || majorVersion === 4 ? 4096 : 512;
  const miniSectorShift = buffer.readUInt16LE(32);
  const miniSectorSize = 1 << (miniSectorShift || 6); // normally 64 bytes
  const firstDirSector = buffer.readUInt32LE(48);
  const miniStreamCutoff = buffer.readUInt32LE(56) || 4096;
  const firstMiniFatSector = buffer.readUInt32LE(60);
  const numMiniFatSectors = buffer.readUInt32LE(64);
  const firstDifatSector = buffer.readUInt32LE(68);
  const numDifatSectors = buffer.readUInt32LE(72);

  void minorVersion;

  const readRawSector = (sectorId: number): Buffer => {
    const offset = (sectorId + 1) * sectorSize;
    if (sectorId >= 0xfffffffa || offset < 0 || offset + sectorSize > buffer.length) {
      return Buffer.alloc(0);
    }
    return buffer.subarray(offset, offset + sectorSize);
  };

  // Collect FAT sector IDs from header DIFAT[0..108] + extra DIFAT chain if needed
  const fatSectorIds: number[] = [];
  for (let i = 0; i < 109; i += 1) {
    const sid = buffer.readUInt32LE(76 + i * 4);
    if (sid !== FREESECT && sid < 0xfffffffa) fatSectorIds.push(sid);
  }
  let difatSid = firstDifatSector;
  let difatGuard = 0;
  while (difatSid !== ENDOFCHAIN && difatSid !== FREESECT && difatSid < 0xfffffffa && difatGuard < numDifatSectors + 4) {
    const sec = readRawSector(difatSid);
    if (sec.length < sectorSize) break;
    const perSector = sectorSize / 4 - 1;
    for (let i = 0; i < perSector; i += 1) {
      const sid = sec.readUInt32LE(i * 4);
      if (sid !== FREESECT && sid < 0xfffffffa) fatSectorIds.push(sid);
    }
    difatSid = sec.readUInt32LE(sectorSize - 4);
    difatGuard += 1;
  }

  // Build FAT table
  const fat: number[] = [];
  for (const fatSid of fatSectorIds) {
    const sec = readRawSector(fatSid);
    if (sec.length < sectorSize) continue;
    for (let off = 0; off < sectorSize; off += 4) {
      fat.push(sec.readUInt32LE(off));
    }
  }

  const readSectorChain = (startSector: number, expectedSize?: number): Buffer => {
    if (startSector >= 0xfffffffa || startSector < 0) return Buffer.alloc(0);
    const chunks: Buffer[] = [];
    const seen = new Set<number>();
    let sid = startSector;
    let total = 0;
    const maxSectors = Math.max(fat.length, 4096);
    while (sid !== ENDOFCHAIN && sid !== FREESECT && sid < 0xfffffffa && !seen.has(sid) && seen.size < maxSectors) {
      seen.add(sid);
      const sec = readRawSector(sid);
      if (sec.length === 0) break;
      chunks.push(sec);
      total += sec.length;
      if (expectedSize !== undefined && expectedSize > 0 && total >= expectedSize) break;
      sid = sid < fat.length ? fat[sid] : ENDOFCHAIN;
    }
    const merged = Buffer.concat(chunks);
    return expectedSize !== undefined && expectedSize >= 0
      ? merged.subarray(0, Math.min(expectedSize, merged.length))
      : merged;
  };

  // Directory entries
  const dirBuffer = readSectorChain(firstDirSector);
  const entries: CfbfDirectoryEntry[] = [];
  const entryCount = Math.floor(dirBuffer.length / 128);
  for (let i = 0; i < entryCount; i += 1) {
    const base = i * 128;
    const objectType = dirBuffer.readUInt8(base + 66);
    if (objectType === 0) continue;
    const nameByteLen = Math.min(64, dirBuffer.readUInt16LE(base + 64));
    const rawName = nameByteLen >= 2
      ? dirBuffer.subarray(base, base + nameByteLen - 2).toString('utf16le')
      : '';
    const leftSiblingId = dirBuffer.readUInt32LE(base + 68);
    const rightSiblingId = dirBuffer.readUInt32LE(base + 72);
    const childId = dirBuffer.readUInt32LE(base + 76);
    const startSector = dirBuffer.readUInt32LE(base + 116);
    const sizeLow = dirBuffer.readUInt32LE(base + 120);
    const sizeHigh = majorVersion >= 4 ? dirBuffer.readUInt32LE(base + 124) : 0;
    const size = sizeHigh > 0 && sizeHigh < 0x1000 ? sizeHigh * 0x100000000 + sizeLow : sizeLow;

    entries.push({
      index: i,
      name: rawName,
      objectType,
      leftSiblingId,
      rightSiblingId,
      childId,
      startSector,
      size,
    });
  }

  // Build MiniFAT + MiniStream from root entry (index 0)
  const rootEntry = entries.find((e) => e.index === 0 || e.objectType === 5);
  const miniStream =
    rootEntry && rootEntry.startSector < 0xfffffffa && rootEntry.size > 0
      ? readSectorChain(rootEntry.startSector, rootEntry.size)
      : Buffer.alloc(0);

  const miniFat: number[] = [];
  if (numMiniFatSectors > 0 && firstMiniFatSector < 0xfffffffa) {
    const mfBuf = readSectorChain(firstMiniFatSector);
    for (let off = 0; off + 4 <= mfBuf.length; off += 4) {
      miniFat.push(mfBuf.readUInt32LE(off));
    }
  }

  const readMiniChain = (startSector: number, size: number): Buffer => {
    if (startSector >= 0xfffffffa || miniStream.length === 0) return Buffer.alloc(0);
    const chunks: Buffer[] = [];
    const seen = new Set<number>();
    let sid = startSector;
    let total = 0;
    while (sid !== ENDOFCHAIN && sid !== FREESECT && sid < 0xfffffffa && !seen.has(sid) && seen.size < miniFat.length + 64) {
      seen.add(sid);
      const off = sid * miniSectorSize;
      if (off < 0 || off >= miniStream.length) break;
      const slice = miniStream.subarray(off, Math.min(off + miniSectorSize, miniStream.length));
      chunks.push(slice);
      total += slice.length;
      if (size > 0 && total >= size) break;
      sid = sid < miniFat.length ? miniFat[sid] : ENDOFCHAIN;
    }
    const merged = Buffer.concat(chunks);
    return size > 0 ? merged.subarray(0, Math.min(size, merged.length)) : merged;
  };

  const readEntryBuffer = (entry: CfbfDirectoryEntry): Buffer => {
    if (entry.size === 0) return Buffer.alloc(0);
    if (entry.size < miniStreamCutoff && miniStream.length > 0 && miniFat.length > 0) {
      return readMiniChain(entry.startSector, entry.size);
    }
    return readSectorChain(entry.startSector, entry.size);
  };

  return { entries, readEntryBuffer };
}

function collectSubtreeEntries(entries: CfbfDirectoryEntry[], storageChildId: number): CfbfDirectoryEntry[] {
  const byIdx = new Map(entries.map((e) => [e.index, e]));
  const result: CfbfDirectoryEntry[] = [];
  const visited = new Set<number>();
  const stack: number[] = [storageChildId];
  while (stack.length > 0) {
    const id = stack.pop()!;
    if (id === FREESECT || visited.has(id)) continue;
    visited.add(id);
    const node = byIdx.get(id);
    if (!node) continue;
    result.push(node);
    if (node.leftSiblingId !== FREESECT) stack.push(node.leftSiblingId);
    if (node.rightSiblingId !== FREESECT) stack.push(node.rightSiblingId);
  }
  return result;
}

function decodeMapiStreamValue(name: string, buf: Buffer): string {
  const upper = name.toUpperCase();
  if (upper.endsWith('001F')) {
    return buf.toString('utf16le').replace(/\0+$/g, '').trim();
  }
  return buf.toString('utf8').replace(/\0+$/g, '').trim();
}

function inferAttachmentKind(filename: string, mimeType: string): AttachmentKind {
  const lower = filename.toLowerCase();
  const mime = mimeType.toLowerCase();
  if (lower.endsWith('.pdf') || mime.includes('pdf')) return 'pdf';
  if (/\.(xlsx?|csv)$/i.test(lower) || mime.includes('spreadsheet') || mime.includes('excel')) return 'excel';
  if (/\.(docx?)$/i.test(lower) || mime.includes('word')) return 'word';
  if (/\.(mp3|wav|m4a|webm|ogg|flac)$/i.test(lower) || mime.startsWith('audio/')) return 'audio';
  if (/\.(txt|md|json)$/i.test(lower) || mime.startsWith('text/')) return 'text';
  return 'other';
}

export function parseMsgBuffer(buffer: Buffer, originalFileName = 'message.msg'): MailItem {
  const cfbf = parseCfbf(buffer);
  const rootEntry = cfbf.entries.find((e) => e.index === 0 || e.objectType === 5);
  const rootChildren =
    rootEntry && rootEntry.childId !== FREESECT
      ? collectSubtreeEntries(cfbf.entries, rootEntry.childId)
      : cfbf.entries;

  const streamMap = new Map<string, Buffer>();
  const attachStorages: CfbfDirectoryEntry[] = [];

  for (const entry of rootChildren) {
    if (entry.objectType === 2) {
      streamMap.set(entry.name.toUpperCase(), cfbf.readEntryBuffer(entry));
    } else if (entry.objectType === 1 && /^__attach_version1\.0_/i.test(entry.name)) {
      attachStorages.push(entry);
    }
  }

  const readProp = (...propKeys: string[]): string => {
    for (const pk of propKeys) {
      const target = `__SUBSTG1.0_${pk.toUpperCase()}`;
      const buf = streamMap.get(target);
      if (buf && buf.length > 0) {
        const val = decodeMapiStreamValue(target, buf);
        if (val) return val;
      }
    }
    return '';
  };

  const subject =
    readProp('0037001F', '0037001E', '0E1D001F', '0E1D001E') ||
    originalFileName.replace(/\.msg$/i, '');
  const fromName =
    readProp('0C1A001F', '0C1A001E', '0042001F', '0042001E') || 'Outlook 발신자';
  const rawFromEmail = readProp('5D01001F', '5D01001E', '0C1F001F', '0C1F001E', '0065001F', '0065001E');
  const fromEmail = rawFromEmail.includes('@') ? rawFromEmail : 'msg-import@outlook.local';
  const displayTo = readProp('0E04001F', '0E04001E');
  const toList = displayTo
    .split(/[;,]/)
    .map((s) => s.trim())
    .filter(Boolean);

  const plainBody = readProp('1000001F', '1000001E');
  let htmlBody = '';
  const htmlBin = streamMap.get('__SUBSTG1.0_10130102') || streamMap.get('__SUBSTG1.0_1013001F');
  if (htmlBin && htmlBin.length > 0) {
    htmlBody = htmlBin.toString('utf8').replace(/\0+$/g, '').trim();
  }
  const body = (plainBody || (htmlBody ? stripHtmlToText(htmlBody) : '') || '(본문 없음)').slice(0, 12000);

  // Extract FILETIME from __properties_version1.0 if present, otherwise current timestamp
  let msgDate = new Date();
  const propsBuf = streamMap.get('__PROPERTIES_VERSION1.0');
  if (propsBuf && propsBuf.length >= 48) {
    for (let off = 32; off + 16 <= propsBuf.length; off += 16) {
      const propTag = propsBuf.readUInt32LE(off);
      const propType = propTag & 0xffff;
      if (propType === 0x0040) {
        const lo = propsBuf.readUInt32LE(off + 8);
        const hi = propsBuf.readUInt32LE(off + 12);
        const filetime100ns = BigInt(hi) * 4294967296n + BigInt(lo);
        const unixMs = Number((filetime100ns - 116444736000000000n) / 10000n);
        if (unixMs > 946684800000 && unixMs < 2524608000000) {
          msgDate = new Date(unixMs);
          break;
        }
      }
    }
  }

  const msgHash = crypto
    .createHash('sha1')
    .update(originalFileName)
    .update(subject)
    .update(buffer.subarray(0, Math.min(2048, buffer.length)))
    .digest('hex')
    .slice(0, 12);
  const id = `msg-${msgHash}`;

  const attachments: MailAttachmentMeta[] = [];
  attachStorages.forEach((storage, idx) => {
    const children =
      storage.childId !== FREESECT ? collectSubtreeEntries(cfbf.entries, storage.childId) : [];
    const attStreams = new Map<string, CfbfDirectoryEntry>();
    for (const child of children) {
      if (child.objectType === 2) {
        attStreams.set(child.name.toUpperCase(), child);
      }
    }
    const readAttStr = (...keys: string[]): string => {
      for (const k of keys) {
        const ent = attStreams.get(`__SUBSTG1.0_${k.toUpperCase()}`);
        if (ent) {
          const val = decodeMapiStreamValue(ent.name, cfbf.readEntryBuffer(ent));
          if (val) return val;
        }
      }
      return '';
    };

    const filename =
      readAttStr('3707001F', '3707001E', '3704001F', '3704001E', '3001001F') ||
      `attachment-${idx + 1}.bin`;
    const mimeType = readAttStr('370E001F', '370E001E') || 'application/octet-stream';
    const dataEnt = attStreams.get('__SUBSTG1.0_37010102');
    const cid = readAttStr('3712001F', '3712001E');
    const sizeBytes = dataEnt ? dataEnt.size : 0;

    attachments.push({
      id: `${id}-att-${idx}`,
      filename,
      mimeType,
      sizeBytes,
      kind: inferAttachmentKind(filename, mimeType),
      ...(cid ? { isInline: true, contentId: cid } : {}),
    });
  });

  const dateStr = msgDate.toISOString().slice(0, 10);
  const timeLabel = msgDate.toTimeString().slice(0, 5);

  return {
    id,
    folder: 'inbox',
    folderName: 'MSG 불러오기',
    subject,
    fromName,
    fromEmail,
    to: toList,
    date: dateStr,
    timestamp: msgDate.toISOString(),
    timeLabel,
    snippet: body.replace(/\s+/g, ' ').slice(0, 140),
    body,
    ...(htmlBody ? { bodyHtml: htmlBody } : {}),
    hasAttachment: attachments.some((a) => !a.isInline),
    ...(attachments.length > 0 ? { attachments } : {}),
    conversationKey: normalizeConversationSubject(subject),
    importance: 'normal',
    category: 'MSG 파일',
    isRead: false,
    isFlagged: false,
    source: 'pst-import',
  };
}

export function parseMsgFile(filePath: string, originalFileName: string): MailItem {
  const buf = fs.readFileSync(filePath);
  return parseMsgBuffer(buf, originalFileName);
}
