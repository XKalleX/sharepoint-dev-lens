#!/usr/bin/env node
/**
 * Generates icons/icon{16,48,128}.png using ONLY Node.js built-ins.
 * No npm, no external packages required.
 * Run once: node scripts/generate-icons.js
 */
import { deflateSync } from 'zlib';
import { writeFileSync, mkdirSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const iconsDir  = resolve(__dirname, '..', 'src', 'icons');

// ── CRC32 ────────────────────────────────────────────────────────────────────
const CRC = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  CRC[n] = c;
}
function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = CRC[(c ^ b) & 0xff] ^ (c >>> 8);
  return (~c) >>> 0;
}

// ── PNG chunk ────────────────────────────────────────────────────────────────
function chunk(type, data) {
  const t = Buffer.from(type, 'ascii');
  const l = Buffer.allocUnsafe(4); l.writeUInt32BE(data.length);
  const ci = Buffer.concat([t, data]);
  const cv = Buffer.allocUnsafe(4); cv.writeUInt32BE(crc32(ci));
  return Buffer.concat([l, t, data, cv]);
}

// ── Build PNG from RGBA buffer ────────────────────────────────────────────────
function buildPng(size, rgba) {
  const ihdr = Buffer.allocUnsafe(13);
  ihdr.writeUInt32BE(size, 0); ihdr.writeUInt32BE(size, 4);
  ihdr.writeUInt8(8, 8); ihdr.writeUInt8(6, 9); // RGBA
  ihdr.writeUInt8(0, 10); ihdr.writeUInt8(0, 11); ihdr.writeUInt8(0, 12);

  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = Buffer.allocUnsafe(1 + size * 4);
    row[0] = 0; // filter: None
    for (let x = 0; x < size; x++) {
      const s = (y * size + x) * 4;
      row.set(rgba.slice(s, s + 4), 1 + x * 4);
    }
    rows.push(row);
  }

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(Buffer.concat(rows), { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

// ── Drawing helpers ──────────────────────────────────────────────────────────
function setPixel(buf, s, x, y, r, g, b, a) {
  if (x < 0 || y < 0 || x >= s || y >= s) return;
  const i = (y * s + x) * 4;
  // Alpha compositing over current colour
  const aF = a / 255, aB = buf[i + 3] / 255;
  const aO = aF + aB * (1 - aF);
  if (aO === 0) return;
  buf[i]     = Math.round((r * aF + buf[i]     * aB * (1 - aF)) / aO);
  buf[i + 1] = Math.round((g * aF + buf[i + 1] * aB * (1 - aF)) / aO);
  buf[i + 2] = Math.round((b * aF + buf[i + 2] * aB * (1 - aF)) / aO);
  buf[i + 3] = Math.round(aO * 255);
}

function fillRect(buf, s, x0, y0, x1, y1, r, g, b, a) {
  for (let y = y0; y <= y1; y++)
    for (let x = x0; x <= x1; x++)
      setPixel(buf, s, x, y, r, g, b, a);
}

// Anti-aliased circle ring using SDF
function drawRing(buf, s, cx, cy, outerR, innerR, r, g, b) {
  const margin = 1.5;
  for (let y = Math.floor(cy - outerR - margin); y <= Math.ceil(cy + outerR + margin); y++) {
    for (let x = Math.floor(cx - outerR - margin); x <= Math.ceil(cx + outerR + margin); x++) {
      const d = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2);
      const distOuter = outerR - d;
      const distInner = d - innerR;
      const alpha = Math.min(Math.max(distOuter, 0), 1) * Math.min(Math.max(distInner, 0), 1);
      if (alpha > 0) setPixel(buf, s, x, y, r, g, b, Math.round(alpha * 255));
    }
  }
}

// Anti-aliased thick line
function drawLine(buf, s, x0, y0, x1, y1, thick, r, g, b) {
  const dx = x1 - x0, dy = y1 - y0;
  const len = Math.sqrt(dx * dx + dy * dy);
  if (len === 0) return;
  const nx = -dy / len, ny = dx / len;
  const margin = thick / 2 + 1.5;
  const minX = Math.floor(Math.min(x0, x1) - margin);
  const maxX = Math.ceil(Math.max(x0, x1) + margin);
  const minY = Math.floor(Math.min(y0, y1) - margin);
  const maxY = Math.ceil(Math.max(y0, y1) + margin);
  for (let y = minY; y <= maxY; y++) {
    for (let x = minX; x <= maxX; x++) {
      const px = x - x0, py = y - y0;
      const along = (px * dx + py * dy) / len;
      const perp  = Math.abs(px * nx + py * ny);
      const aAlong = Math.min(Math.max(along + .5, 0), 1) * Math.min(Math.max(len - along + .5, 0), 1);
      const aPerp  = Math.min(Math.max(thick / 2 - perp + .5, 0), 1);
      const alpha  = aAlong * aPerp;
      if (alpha > 0) setPixel(buf, s, x, y, r, g, b, Math.round(alpha * 255));
    }
  }
}

// Rounded-rect background
function drawBackground(buf, s) {
  const rr = Math.max(2, Math.round(s * 0.18));
  // Fill full rect, then clear corners
  fillRect(buf, s, 0, 0, s - 1, s - 1, 0, 120, 212, 255);
  // Clear corners using anti-aliased circle
  for (const [cx, cy] of [[rr, rr], [s-1-rr, rr], [rr, s-1-rr], [s-1-rr, s-1-rr]]) {
    for (let y = 0; y < s; y++) {
      for (let x = 0; x < s; x++) {
        // Only process the corner quadrants
        if (x > rr && x < s - 1 - rr) continue;
        if (y > rr && y < s - 1 - rr) continue;
        const d = Math.sqrt((x - cx) ** 2 + (y - cy) ** 2);
        const alpha = Math.min(Math.max(rr - d + .5, 0), 1);
        if (alpha < 1) {
          const i = (y * s + x) * 4;
          buf[i + 3] = Math.round(buf[i + 3] * alpha);
        }
      }
    }
  }
}

// ── Icon generator ────────────────────────────────────────────────────────────
function generateIcon(size) {
  const buf = new Uint8Array(size * size * 4); // transparent
  drawBackground(buf, size);

  if (size >= 48) {
    // Magnifier glass
    const cx  = size * 0.43, cy  = size * 0.43;
    const ouR = size * 0.27, inR = size * 0.18;
    const lw  = ouR - inR;
    drawRing(buf, size, cx, cy, ouR, inR, 255, 255, 255);

    // Handle
    const angle = Math.PI / 4;
    const hx0 = cx + (ouR - lw * .3) * Math.cos(angle);
    const hy0 = cy + (ouR - lw * .3) * Math.sin(angle);
    const hx1 = cx + (ouR + size * .20) * Math.cos(angle);
    const hy1 = cy + (ouR + size * .20) * Math.sin(angle);
    drawLine(buf, size, hx0, hy0, hx1, hy1, lw, 255, 255, 255);
  } else {
    // 16px: simple white dot
    drawRing(buf, size, size * .48, size * .45, size * .28, size * .16, 255, 255, 255);
    drawLine(buf, size,
      size * .65, size * .65,
      size * .82, size * .82, size * .10, 255, 255, 255);
  }

  return buildPng(size, buf);
}

// ── Write files ───────────────────────────────────────────────────────────────
mkdirSync(iconsDir, { recursive: true });
for (const size of [16, 48, 128]) {
  const path = resolve(iconsDir, `icon${size}.png`);
  writeFileSync(path, generateIcon(size));
  console.log(`  ✓  icon${size}.png`);
}
console.log('\n  Icons written to src/icons/');
