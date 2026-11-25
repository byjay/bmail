import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import zlib from 'node:zlib';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DESKTOP_DIR = path.resolve(__dirname, '..', 'desktop');

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) {
    c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  }
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const t = Buffer.from(type, 'ascii');
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(Buffer.concat([t, data])), 0);
  return Buffer.concat([len, t, data, crc]);
}

function blendPixel(raw, stride, size, x, y, r, g, b, a) {
  if (x < 0 || y < 0 || x >= size || y >= size || a <= 0) return;
  const idx = y * stride + 1 + x * 4;
  const srcA = Math.min(255, Math.max(0, a)) / 255;
  const dstA = raw[idx + 3] / 255;
  const outA = srcA + dstA * (1 - srcA);
  if (outA <= 0) return;
  raw[idx] = Math.round((r * srcA + raw[idx] * dstA * (1 - srcA)) / outA);
  raw[idx + 1] = Math.round((g * srcA + raw[idx + 1] * dstA * (1 - srcA)) / outA);
  raw[idx + 2] = Math.round((b * srcA + raw[idx + 2] * dstA * (1 - srcA)) / outA);
  raw[idx + 3] = Math.round(outA * 255);
}

function distToSegment(px, py, x1, y1, x2, y2) {
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

function renderBmailIconPng(size) {
  const stride = size * 4 + 1;
  const raw = Buffer.alloc(stride * size);

  const margin = size * 0.06;
  const minBound = margin;
  const maxBound = size - margin;
  const cornerR = size * 0.22;

  // Envelope rect
  const envL = size * 0.19;
  const envR = size * 0.81;
  const envT = size * 0.29;
  const envB = size * 0.72;
  const envCorner = size * 0.055;
  const midX = size * 0.5;
  const flapY = size * 0.53;
  const strokeHalf = Math.max(0.85, size * 0.036);

  // Accent AI dot (top-right inside badge)
  const dotX = size * 0.77;
  const dotY = size * 0.23;
  const dotR = size * 0.095;

  for (let y = 0; y < size; y += 1) {
    raw[y * stride] = 0;
    for (let x = 0; x < size; x += 1) {
      const px = x + 0.5;
      const py = y + 0.5;

      // Signed distance to rounded badge rect
      const cx = Math.max(minBound + cornerR, Math.min(maxBound - cornerR, px));
      const cy = Math.max(minBound + cornerR, Math.min(maxBound - cornerR, py));
      const dx = Math.abs(px - cx);
      const dy = Math.abs(py - cy);
      const clampedX = Math.max(minBound + cornerR, Math.min(maxBound - cornerR, px));
      const clampedY = Math.max(minBound + cornerR, Math.min(maxBound - cornerR, py));
      const badgeDist =
        Math.hypot(px - clampedX, py - clampedY) - cornerR;

      if (badgeDist <= 0.8) {
        const alpha = badgeDist <= -0.5 ? 255 : Math.round((0.8 - badgeDist) * (255 / 1.3));
        const t = y / size;
        const r = Math.round(18 * (1 - t) + 11 * t);
        const g = Math.round(124 * (1 - t) + 92 * t);
        const b = Math.round(212 * (1 - t) + 171 * t);
        blendPixel(raw, stride, size, x, y, r, g, b, alpha);

        // Subtle inner highlight border
        if (badgeDist > -Math.max(1.2, size * 0.025) && badgeDist <= -0.1) {
          blendPixel(raw, stride, size, x, y, 255, 255, 255, 42);
        }
      }

      // White envelope fill
      const ecx = Math.max(envL + envCorner, Math.min(envR - envCorner, px));
      const ecy = Math.max(envT + envCorner, Math.min(envB - envCorner, py));
      const envDist = Math.hypot(px - ecx, py - ecy) - envCorner;
      if (envDist <= 0.6) {
        const alpha = envDist <= -0.4 ? 248 : Math.round((0.6 - envDist) * 248);
        blendPixel(raw, stride, size, x, y, 255, 255, 255, alpha);
      }

      // Envelope flap V lines etched cleanly in Fluent brand blue
      if (envDist <= -0.2) {
        const dLeft = distToSegment(px, py, envL + strokeHalf * 0.5, envT + strokeHalf * 0.5, midX, flapY);
        const dRight = distToSegment(px, py, envR - strokeHalf * 0.5, envT + strokeHalf * 0.5, midX, flapY);
        const flapD = Math.min(dLeft, dRight);
        if (flapD <= strokeHalf + 0.5) {
          const a = flapD <= strokeHalf - 0.3 ? 255 : Math.round(((strokeHalf + 0.5 - flapD) / 0.8) * 255);
          blendPixel(raw, stride, size, x, y, 15, 108, 189, a);
        }
      }

      // Bright sky-blue AI indicator dot (top right)
      const dotDist = Math.hypot(px - dotX, py - dotY);
      if (dotDist <= dotR + 1.4) {
        if (dotDist > dotR - 0.3) {
          const ringA = dotDist <= dotR + 0.5 ? 255 : Math.round(((dotR + 1.4 - dotDist) / 0.9) * 255);
          blendPixel(raw, stride, size, x, y, 15, 108, 189, ringA);
        } else {
          blendPixel(raw, stride, size, x, y, 56, 189, 248, 255);
        }
      }
      void dx;
      void dy;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

function buildIcoFromPngFrames(frames) {
  const count = frames.length;
  const headerSize = 6;
  const dirEntrySize = 16;
  let offset = headerSize + count * dirEntrySize;

  const header = Buffer.alloc(headerSize);
  header.writeUInt16LE(0, 0); // reserved
  header.writeUInt16LE(1, 2); // type 1 = ICO
  header.writeUInt16LE(count, 4);

  const entries = [];
  for (const frame of frames) {
    const entry = Buffer.alloc(dirEntrySize);
    entry[0] = frame.size >= 256 ? 0 : frame.size;
    entry[1] = frame.size >= 256 ? 0 : frame.size;
    entry[2] = 0; // palette colors
    entry[3] = 0; // reserved
    entry.writeUInt16LE(1, 4); // planes
    entry.writeUInt16LE(32, 6); // bits per pixel
    entry.writeUInt32LE(frame.png.length, 8);
    entry.writeUInt32LE(offset, 12);
    entries.push(entry);
    offset += frame.png.length;
  }

  return Buffer.concat([header, ...entries, ...frames.map((f) => f.png)]);
}

fs.mkdirSync(DESKTOP_DIR, { recursive: true });
const sizes = [16, 32, 48, 64, 128, 256];
const frames = sizes.map((size) => ({ size, png: renderBmailIconPng(size) }));
const png256 = frames[frames.length - 1].png;
const icoBuf = buildIcoFromPngFrames(frames);

const pngOut = path.join(DESKTOP_DIR, 'icon.png');
const icoOut = path.join(DESKTOP_DIR, 'icon.ico');
fs.writeFileSync(pngOut, png256);
fs.writeFileSync(icoOut, icoBuf);

console.log(
  JSON.stringify(
    {
      pngOut,
      pngBytes: png256.length,
      pngMagic: png256.subarray(0, 4).toString('hex'),
      icoOut,
      icoBytes: icoBuf.length,
      icoMagic: icoBuf.subarray(0, 4).toString('hex'),
      frameCount: frames.length,
    },
    null,
    2
  )
);
