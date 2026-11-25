import zlib from 'node:zlib';

const CRC_TABLE = new Uint32Array(256);
for (let n = 0; n < 256; n += 1) {
  let c = n;
  for (let k = 0; k < 8; k += 1) {
    c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  }
  CRC_TABLE[n] = c >>> 0;
}

function crc32(buf: Buffer): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function makeChunk(type: string, data: Buffer): Buffer {
  const typeBuf = Buffer.from(type, 'ascii');
  const lenBuf = Buffer.alloc(4);
  lenBuf.writeUInt32BE(data.length, 0);
  const crcInput = Buffer.concat([typeBuf, data]);
  const crcBuf = Buffer.alloc(4);
  crcBuf.writeUInt32BE(crc32(crcInput), 0);
  return Buffer.concat([lenBuf, typeBuf, data, crcBuf]);
}

function insideRoundedRect(
  x: number,
  y: number,
  x0: number,
  y0: number,
  w: number,
  h: number,
  r: number
): boolean {
  if (x < x0 || x > x0 + w || y < y0 || y > y0 + h) return false;
  const cx = x < x0 + r ? x0 + r : x > x0 + w - r ? x0 + w - r : x;
  const cy = y < y0 + r ? y0 + r : y > y0 + h - r ? y0 + h - r : y;
  const dx = x - cx;
  const dy = y - cy;
  return dx * dx + dy * dy <= r * r;
}

function distToSegment(
  px: number,
  py: number,
  x1: number,
  y1: number,
  x2: number,
  y2: number
): number {
  const vx = x2 - x1;
  const vy = y2 - y1;
  const wx = px - x1;
  const wy = py - y1;
  const c1 = vx * wx + vy * wy;
  if (c1 <= 0) return Math.hypot(px - x1, py - y1);
  const c2 = vx * vx + vy * vy;
  if (c2 <= c1) return Math.hypot(px - x2, py - y2);
  const b = c1 / c2;
  return Math.hypot(px - (x1 + b * vx), py - (y1 + b * vy));
}

/**
 * Generates a crisp, valid RGBA PNG buffer for Outlook Add-in Ribbon icons
 * (16x16, 32x32, 64x64, 80x80, 128x128) with zero external native image dependencies.
 */
export function createIconPngBuffer(size: number): Buffer {
  const clamped = Math.max(16, Math.min(256, Math.round(size)));
  const stride = clamped * 4 + 1;
  const raw = Buffer.alloc(stride * clamped);

  for (let py = 0; py < clamped; py += 1) {
    const rowStart = py * stride;
    raw[rowStart] = 0; // Filter type 0 (None)
    const ny = (py + 0.5) / clamped;

    for (let px = 0; px < clamped; px += 1) {
      const nx = (px + 0.5) / clamped;
      const idx = rowStart + 1 + px * 4;

      // Outer Cobalt Blue rounded square (#0F6CBD -> #0B4F8A subtle vertical tone)
      const inOuter = insideRoundedRect(nx, ny, 0.04, 0.04, 0.92, 0.92, 0.2);
      if (!inOuter) {
        raw[idx] = 0;
        raw[idx + 1] = 0;
        raw[idx + 2] = 0;
        raw[idx + 3] = 0;
        continue;
      }

      // Base cobalt gradient
      let r = Math.round(15 - ny * 4);
      let g = Math.round(115 - ny * 26);
      let b = Math.round(198 - ny * 38);
      let a = 255;

      // Crisp white envelope + document glyph
      const inSheet = insideRoundedRect(nx, ny, 0.21, 0.19, 0.58, 0.44, 0.05);
      const inEnvelopeBase = insideRoundedRect(nx, ny, 0.18, 0.42, 0.64, 0.39, 0.06);

      if (inSheet) {
        // Editorial lines inside the top document sheet
        const isLine1 = ny >= 0.28 && ny <= 0.32 && nx >= 0.31 && nx <= 0.69;
        const isLine2 = ny >= 0.37 && ny <= 0.41 && nx >= 0.31 && nx <= 0.58;
        if (isLine1 || isLine2) {
          r = 15;
          g = 108;
          b = 189;
        } else {
          r = 248;
          g = 250;
          b = 252;
        }
      }

      if (inEnvelopeBase) {
        const dLeft = distToSegment(nx, ny, 0.2, 0.45, 0.5, 0.64);
        const dRight = distToSegment(nx, ny, 0.8, 0.45, 0.5, 0.64);
        const strokeWidth = clamped <= 24 ? 0.055 : 0.036;
        if (dLeft < strokeWidth || dRight < strokeWidth) {
          r = 15;
          g = 108;
          b = 189;
        } else {
          r = 255;
          g = 255;
          b = 255;
        }
      }

      raw[idx] = r;
      raw[idx + 1] = g;
      raw[idx + 2] = b;
      raw[idx + 3] = a;
    }
  }

  const signature = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(clamped, 0);
  ihdr.writeUInt32BE(clamped, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // color type RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  const idat = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    signature,
    makeChunk('IHDR', ihdr),
    makeChunk('IDAT', idat),
    makeChunk('IEND', Buffer.alloc(0)),
  ]);
}
