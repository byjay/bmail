import assert from 'node:assert/strict';
import test from 'node:test';
import { parseMsgBuffer } from './msgFileParser';

const CFBF_MAGIC = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

interface FixtureEntry {
  name: string;
  objectType: number;
  leftSiblingId?: number;
  rightSiblingId?: number;
  childId?: number;
  startSector?: number;
  size?: number;
}

function buildMinimalCfbfMsgFixture(): Buffer {
  const sectorSize = 512;
  // Layout:
  // Sector 0: FAT
  // Sector 1: Directory (up to 4 entries of 128 bytes -> use 2 sectors = 8 entries)
  // Sector 2: Directory cont.
  // Sector 3: Subject UTF-16LE (padded >= 4096 cutoff so regular FAT applies!)
  // Wait: let's simply put MiniStream cutoff = 0 so every stream lives in regular 512-byte sectors cleanly!
  // Let's verify how many sectors we need:
  // Sector 0: FAT
  // Sector 1..2: Directory (slots 0..7)
  // Sector 3..10: streams + attachment streams

  const streams: Array<{ name: string; buf: Buffer; startSector: number }> = [
    { name: '__substg1.0_0037001F', buf: Buffer.from('H2710 도면 MSG 직송 테스트\0', 'utf16le'), startSector: 3 },
    { name: '__substg1.0_0C1A001F', buf: Buffer.from('박붕근 대표\0', 'utf16le'), startSector: 4 },
    { name: '__substg1.0_5D01001F', buf: Buffer.from('park@example.com\0', 'utf16le'), startSector: 5 },
    { name: '__substg1.0_0E04001F', buf: Buffer.from('김봉정 부장\0', 'utf16le'), startSector: 6 },
    { name: '__substg1.0_1000001F', buf: Buffer.from('본문 테스트 내용입니다.\r\n확인 부탁드립니다.\0', 'utf16le'), startSector: 7 },
    { name: '__substg1.0_3707001F', buf: Buffer.from('H2710-SPEC-REV1.pdf\0', 'utf16le'), startSector: 8 },
    { name: '__substg1.0_370E001F', buf: Buffer.from('application/pdf\0', 'utf16le'), startSector: 9 },
    { name: '__substg1.0_37010102', buf: Buffer.from('PDF-BINARY-DUMMY-BYTES'), startSector: 10 },
  ];

  // Directory layout (indices 0..10):
  // 0: Root Entry -> childId = 1
  // 1..5: top streams chained via rightSiblingId, index 5 -> rightSiblingId = 6 (__attach_version1.0_#00000000)
  // 6: __attach_version1.0_#00000000 (storage, objectType=1) -> childId = 7
  // 7..9: attach streams chained via rightSiblingId
  const dirEntries: FixtureEntry[] = [
    { name: 'Root Entry', objectType: 5, childId: 1, startSector: ENDOFCHAIN, size: 0 },
    { name: streams[0].name, objectType: 2, rightSiblingId: 2, startSector: 3, size: streams[0].buf.length },
    { name: streams[1].name, objectType: 2, rightSiblingId: 3, startSector: 4, size: streams[1].buf.length },
    { name: streams[2].name, objectType: 2, rightSiblingId: 4, startSector: 5, size: streams[2].buf.length },
    { name: streams[3].name, objectType: 2, rightSiblingId: 5, startSector: 6, size: streams[3].buf.length },
    { name: streams[4].name, objectType: 2, rightSiblingId: 6, startSector: 7, size: streams[4].buf.length },
    { name: '__attach_version1.0_#00000000', objectType: 1, childId: 7, startSector: ENDOFCHAIN, size: 0 },
    { name: streams[5].name, objectType: 2, rightSiblingId: 8, startSector: 8, size: streams[5].buf.length },
    { name: streams[6].name, objectType: 2, rightSiblingId: 9, startSector: 9, size: streams[6].buf.length },
    { name: streams[7].name, objectType: 2, startSector: 10, size: streams[7].buf.length },
  ];

  const dirSectorsCount = 3; // 10 * 128 = 1280 bytes -> 3 sectors (sectors 1, 11, 12 or let's put dir in sectors 1,2,11)
  void dirSectorsCount;
  // Let's keep it clean: dir uses sectors 11, 12, 13? Easier: dir at sectors 1, 2, 11
  const totalSectors = 12;
  const header = Buffer.alloc(sectorSize, 0xff);
  CFBF_MAGIC.copy(header, 0);
  header.writeUInt16LE(0x003e, 24);
  header.writeUInt16LE(0x0003, 26); // v3
  header.writeUInt16LE(0xfffe, 28);
  header.writeUInt16LE(9, 30); // 512B sector
  header.writeUInt16LE(6, 32); // 64B mini sector
  header.writeUInt32LE(0, 40);
  header.writeUInt32LE(1, 44); // 1 FAT sector
  header.writeUInt32LE(1, 48); // firstDirSector = 1
  header.writeUInt32LE(0, 52);
  header.writeUInt32LE(1, 56); // miniStreamCutoff = 1 -> all streams >0 use regular FAT!
  header.writeUInt32LE(ENDOFCHAIN, 60);
  header.writeUInt32LE(0, 64);
  header.writeUInt32LE(ENDOFCHAIN, 68);
  header.writeUInt32LE(0, 72);
  header.writeUInt32LE(0, 76); // DIFAT[0] = sector 0

  const fatSector = Buffer.alloc(sectorSize, 0xff);
  fatSector.writeUInt32LE(0xfffffffd, 0 * 4); // FATSECT
  fatSector.writeUInt32LE(2, 1 * 4); // dir 1 -> 2
  fatSector.writeUInt32LE(11, 2 * 4); // dir 2 -> 11
  fatSector.writeUInt32LE(ENDOFCHAIN, 11 * 4);
  for (const s of streams) {
    fatSector.writeUInt32LE(ENDOFCHAIN, s.startSector * 4);
  }

  const dirTotal = Buffer.alloc(3 * sectorSize, 0);
  dirEntries.forEach((ent, idx) => {
    const off = idx * 128;
    const u16 = Buffer.from(`${ent.name}\0`, 'utf16le');
    u16.subarray(0, 64).copy(dirTotal, off);
    dirTotal.writeUInt16LE(Math.min(64, u16.length), off + 64);
    dirTotal.writeUInt8(ent.objectType, off + 66);
    dirTotal.writeUInt8(1, off + 67);
    dirTotal.writeUInt32LE(ent.leftSiblingId ?? FREESECT, off + 68);
    dirTotal.writeUInt32LE(ent.rightSiblingId ?? FREESECT, off + 72);
    dirTotal.writeUInt32LE(ent.childId ?? FREESECT, off + 76);
    dirTotal.writeUInt32LE(ent.startSector ?? ENDOFCHAIN, off + 116);
    dirTotal.writeUInt32LE(ent.size ?? 0, off + 120);
  });

  const bodySectors = Array.from({ length: totalSectors }, () => Buffer.alloc(sectorSize, 0));
  fatSector.copy(bodySectors[0]);
  dirTotal.subarray(0, 512).copy(bodySectors[1]);
  dirTotal.subarray(512, 1024).copy(bodySectors[2]);
  dirTotal.subarray(1024, 1536).copy(bodySectors[11]);
  for (const s of streams) {
    s.buf.copy(bodySectors[s.startSector]);
  }

  return Buffer.concat([header, ...bodySectors]);
}

test('parseMsgBuffer extracts MAPI UTF-16LE subject, sender, body, and attachments from CFBF .msg binary', () => {
  const fixture = buildMinimalCfbfMsgFixture();
  const mail = parseMsgBuffer(fixture, 'sample-h2710.msg');

  assert.equal(mail.subject, 'H2710 도면 MSG 직송 테스트');
  assert.equal(mail.fromName, '박붕근 대표');
  assert.equal(mail.fromEmail, 'park@example.com');
  assert.deepEqual(mail.to, ['김봉정 부장']);
  assert.ok(mail.body.includes('본문 테스트 내용입니다.'));
  assert.equal(mail.hasAttachment, true);
  assert.equal(mail.attachments?.length, 1);
  assert.equal(mail.attachments?.[0].filename, 'H2710-SPEC-REV1.pdf');
  assert.equal(mail.attachments?.[0].kind, 'pdf');
  assert.equal(mail.attachments?.[0].sizeBytes, 22);
});

test('parseMsgBuffer rejects non-OLE2 malformed buffers safely', () => {
  assert.throws(
    () => parseMsgBuffer(Buffer.from('NOT_A_MSG_FILE'), 'broken.msg'),
    /유효한 Microsoft Outlook \.msg/
  );
});
