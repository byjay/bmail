import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';

function clamp(v, lo = 0, hi = 255) {
  return Math.max(lo, Math.min(hi, Math.round(v)));
}

function lerp(a, b, t) {
  return a + (b - a) * t;
}

function blendPixel(buf, idx, r, g, b, alpha) {
  if (alpha <= 0) return;
  const a = Math.min(1, alpha);
  const inv = 1 - a;
  const dstA = buf[idx + 3] / 255;
  const outA = a + dstA * inv;
  if (outA <= 0) return;
  buf[idx] = clamp((buf[idx] * dstA * inv + r * a) / outA);
  buf[idx + 1] = clamp((buf[idx + 1] * dstA * inv + g * a) / outA);
  buf[idx + 2] = clamp((buf[idx + 2] * dstA * inv + b * a) / outA);
  buf[idx + 3] = clamp(outA * 255);
}

// Signed distance to a rounded squircle box (supersampled cleanly)
function sdRoundedBox(px, py, cx, cy, halfW, halfH, radius) {
  const qx = Math.abs(px - cx) - halfW + radius;
  const qy = Math.abs(py - cy) - halfH + radius;
  const outX = Math.max(qx, 0);
  const outY = Math.max(qy, 0);
  return Math.hypot(outX, outY) + Math.min(Math.max(qx, qy), 0) - radius;
}

// Distance to segment [ax,ay]-[bx,by]
function distToSegment(px, py, ax, ay, bx, by) {
  const vx = bx - ax;
  const vy = by - ay;
  const wx = px - ax;
  const wy = py - ay;
  const lenSq = vx * vx + vy * vy || 1;
  const t = Math.max(0, Math.min(1, (wx * vx + wy * vy) / lenSq));
  return Math.hypot(px - (ax + t * vx), py - (ay + t * vy));
}

function renderAppIconRgba(size) {
  const rgba = Buffer.alloc(size * size * 4, 0);
  const margin = size * 0.0625; // compact macOS/Win11 squircle margin
  const half = size * 0.5 - margin;
  const radius = size * 0.22;
  const cx = size * 0.5;
  const cy = size * 0.5;

  // 2x2 subpixel supersampling for crisp edges
  const samples = [
    [-0.25, -0.25],
    [0.25, -0.25],
    [-0.25, 0.25],
    [0.25, 0.25],
  ];

  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      let accR = 0;
      let accG = 0;
      let accB = 0;
      let accA = 0;

      for (const [ox, oy] of samples) {
        const px = x + 0.5 + ox;
        const py = y + 0.5 + oy;
        const u = px / size;
        const v = py / size;

        const dBox = sdRoundedBox(px, py, cx, cy, half, half, radius);
        let r = 0;
        let g = 0;
        let b = 0;
        let a = 0;

        if (dBox < 0.75) {
          const cov = clamp(0.75 - dBox, 0, 1);
          // Rich Fluent ocean gradient (#1B89EB top-left -> #0F6CBD mid -> #06356B bottom-right)
          const tDiag = Math.max(0, Math.min(1, u * 0.55 + v * 0.55));
          r = lerp(29, 7, tDiag);
          g = lerp(148, 56, tDiag);
          b = lerp(241, 118, tDiag);

          // Top subtle glass highlight arc
          const glossDist = Math.hypot((u - 0.36) * 1.1, (v - 0.16) * 1.8);
          if (glossDist < 0.55) {
            const gStrength = Math.pow(1 - glossDist / 0.55, 2) * 0.26;
            r = lerp(r, 255, gStrength);
            g = lerp(g, 255, gStrength);
            b = lerp(b, 255, gStrength);
          }

          // Subtle inner 1px specular border
          if (dBox > -Math.max(1.2, size * 0.018)) {
            r = lerp(r, 175, 0.32);
            g = lerp(g, 222, 0.32);
            b = lerp(b, 255, 0.32);
          }

          // Crisp White Mail Envelope centered slightly lower-left
          const envW = size * 0.26;
          const envH = size * 0.185;
          const envCx = size * 0.485;
          const envCy = size * 0.545;
          const envR = Math.max(1.5, size * 0.045);
          const dEnv = sdRoundedBox(px, py, envCx, envCy, envW, envH, envR);

          // Drop shadow under envelope
          const dEnvShadow = sdRoundedBox(px, py - size * 0.02, envCx, envCy, envW, envH, envR);
          if (dEnvShadow < 2 && dEnv > 0) {
            const sh = Math.max(0, 1 - Math.abs(dEnvShadow) / 2.5) * 0.22;
            r = lerp(r, 3, sh);
            g = lerp(g, 20, sh);
            b = lerp(b, 50, sh);
          }

          if (dEnv < 0.6) {
            const ec = Math.min(1, 0.6 - dEnv);
            // Soft bottom-edge envelope shading
            const shade = Math.max(0, (py - (envCy - envH)) / (2 * envH));
            const er = lerp(255, 232, shade * 0.55);
            const eg = lerp(255, 243, shade * 0.55);
            const eb = lerp(255, 255, shade * 0.55);
            r = lerp(r, er, ec);
            g = lerp(g, eg, ec);
            b = lerp(b, eb, ec);

            // Clean V-fold flap crease inside envelope
            const flapL = distToSegment(
              px,
              py,
              envCx - envW * 0.86,
              envCy - envH * 0.72,
              envCx,
              envCy + envH * 0.12
            );
            const flapR = distToSegment(
              px,
              py,
              envCx + envW * 0.86,
              envCy - envH * 0.72,
              envCx,
              envCy + envH * 0.12
            );
            const strokeW = Math.max(0.95, size * 0.036);
            const dCrease = Math.min(flapL, flapR);
            if (dCrease < strokeW) {
              const lc = Math.max(0, 1 - Math.max(0, dCrease - strokeW * 0.35) / (strokeW * 0.65));
              r = lerp(r, 15, lc);
              g = lerp(g, 108, lc);
              b = lerp(b, 189, lc);
            }
          }

          // Luminous Cyan/White 4-Point AI Sparkle Badge in Top-Right corner
          const spCx = size * 0.735;
          const spCy = size * 0.275;
          const spSize = size * 0.145;
          const sx = Math.abs(px - spCx) / spSize;
          const sy = Math.abs(py - spCy) / spSize;
          // Astroid / 4-point star curve: sx^0.65 + sy^0.65 <= 1
          const astroid = Math.pow(sx, 0.62) + Math.pow(sy, 0.62);
          if (astroid < 1.35) {
            const glow = Math.max(0, 1 - astroid / 1.35) * 0.45;
            r = lerp(r, 125, glow);
            g = lerp(g, 232, glow);
            b = lerp(b, 255, glow);
          }
          if (astroid < 1.0) {
            const sc = Math.min(1, (1.0 - astroid) * 3.5);
            const sr = lerp(110, 255, Math.max(0, 1 - Math.hypot(sx, sy)));
            const sg = lerp(235, 255, Math.max(0, 1 - Math.hypot(sx, sy)));
            r = lerp(r, sr, sc);
            g = lerp(g, sg, sc);
            b = lerp(b, 255, sc);
          }

          a = cov;
        }

        accR += r * a;
        accG += g * a;
        accB += b * a;
        accA += a;
      }

      const idx = (y * size + x) * 4;
      if (accA > 0) {
        rgba[idx] = clamp(accR / accA);
        rgba[idx + 1] = clamp(accG / accA);
        rgba[idx + 2] = clamp(accB / accA);
        rgba[idx + 3] = clamp((accA / samples.length) * 255);
      }
    }
  }
  return rgba;
}

const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n += 1) {
  let c = n;
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  crcTable[n] = c >>> 0;
}
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) c = crcTable[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
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
function encodeRgbaToPng(size, rgba) {
  const stride = size * 4 + 1;
  const raw = Buffer.alloc(stride * size);
  for (let y = 0; y < size; y += 1) {
    raw[y * stride] = 0;
    rgba.subarray(y * size * 4, (y + 1) * size * 4).copy(raw, y * stride + 1);
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  const compressed = zlib.deflateSync(raw, { level: 9 });
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', compressed),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

function encodeRgbaToBmpDib(size, rgba) {
  const headerSize = 40;
  const pixelBytes = size * size * 4;
  const maskRowBytes = Math.ceil(size / 32) * 4;
  const maskBytes = maskRowBytes * size;
  const buf = Buffer.alloc(headerSize + pixelBytes + maskBytes, 0);

  buf.writeUInt32LE(headerSize, 0);
  buf.writeInt32LE(size, 4);
  buf.writeInt32LE(size * 2, 8);
  buf.writeUInt16LE(1, 12);
  buf.writeUInt16LE(32, 14);
  buf.writeUInt32LE(0, 16);
  buf.writeUInt32LE(pixelBytes + maskBytes, 20);

  for (let y = 0; y < size; y += 1) {
    const srcRow = (size - 1 - y) * size * 4;
    const dstRow = headerSize + y * size * 4;
    for (let x = 0; x < size; x += 1) {
      const s = srcRow + x * 4;
      const d = dstRow + x * 4;
      buf[d] = rgba[s + 2]; // B
      buf[d + 1] = rgba[s + 1]; // G
      buf[d + 2] = rgba[s]; // R
      buf[d + 3] = rgba[s + 3]; // A
    }
  }
  return buf;
}

function buildIcoBuffer(sizes) {
  const images = sizes.map((sz) => {
    const rgba = renderAppIconRgba(sz);
    const data = sz >= 256 ? encodeRgbaToPng(sz, rgba) : encodeRgbaToBmpDib(sz, rgba);
    return { sz, data };
  });

  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(images.length, 4);

  const dirEntries = [];
  let offset = 6 + images.length * 16;
  for (const img of images) {
    const entry = Buffer.alloc(16);
    entry[0] = img.sz >= 256 ? 0 : img.sz;
    entry[1] = img.sz >= 256 ? 0 : img.sz;
    entry[2] = 0;
    entry[3] = 0;
    entry.writeUInt16LE(1, 4);
    entry.writeUInt16LE(32, 6);
    entry.writeUInt32LE(img.data.length, 8);
    entry.writeUInt32LE(offset, 12);
    dirEntries.push(entry);
    offset += img.data.length;
  }

  return Buffer.concat([header, ...dirEntries, ...images.map((i) => i.data)]);
}

const root = path.resolve(import.meta.dirname, '..');
const png256 = encodeRgbaToPng(256, renderAppIconRgba(256));
fs.writeFileSync(path.join(root, 'desktop', 'icon.png'), png256);

const icoBuf = buildIcoBuffer([16, 32, 48, 64, 128, 256]);
fs.writeFileSync(path.join(root, 'desktop', 'icon.ico'), icoBuf);

const svgIcon = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 256 256">
  <defs>
    <linearGradient id="bg" x1="12%" y1="8%" x2="90%" y2="92%">
      <stop offset="0%" stop-color="#1D94F1"/>
      <stop offset="52%" stop-color="#0F6CBD"/>
      <stop offset="100%" stop-color="#06356B"/>
    </linearGradient>
    <linearGradient id="spark" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FFFFFF"/>
      <stop offset="100%" stop-color="#7EE8FF"/>
    </linearGradient>
  </defs>
  <rect x="16" y="16" width="224" height="224" rx="56" fill="url(#bg)"/>
  <rect x="17" y="17" width="222" height="222" rx="55" fill="none" stroke="rgba(255,255,255,0.28)" stroke-width="2.5"/>
  <rect x="54" y="90" width="136" height="96" rx="14" fill="#FFFFFF" opacity="0.98"/>
  <path d="M66 102 L122 146 L178 102" fill="none" stroke="#0F6CBD" stroke-width="11" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="M188 34 C190 60 200 70 226 72 C200 74 190 84 188 110 C186 84 176 74 150 72 C176 70 186 60 188 34 Z" fill="url(#spark)"/>
</svg>
`;
fs.mkdirSync(path.join(root, 'public'), { recursive: true });
fs.writeFileSync(path.join(root, 'public', 'favicon.svg'), svgIcon, 'utf8');

console.log(JSON.stringify({
  iconPngBytes: png256.length,
  iconIcoBytes: icoBuf.length,
  faviconSvgBytes: Buffer.byteLength(svgIcon, 'utf8'),
}, null, 2));
