/* ---------------------------------------------------------------------------
 * make-icons.mjs — regenerate icons/icon{16,32,48,128}.png
 *
 * Draws the shield mark used by the extension: a heater shield filled with the
 * brand gradient (indigo -> violet), a soft top-left sheen and a white check.
 * Rasterised by hand with 4x4 supersampling, encoded as PNG with node:zlib —
 * no third-party dependency, so the icons stay reproducible.
 *
 *   node tools/make-icons.mjs
 * ------------------------------------------------------------------------- */

import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SIZES = [16, 32, 48, 128];
const SS = 4; // supersampling factor per axis

/* ------------------------------- geometry -------------------------------- */

const lerp = (a, b, t) => a + (b - a) * t;

function quad(p0, p1, p2, t) {
  const u = 1 - t;
  return [
    u * u * p0[0] + 2 * u * t * p1[0] + t * t * p2[0],
    u * u * p0[1] + 2 * u * t * p1[1] + t * t * p2[1],
  ];
}

function cubic(p0, p1, p2, p3, t) {
  const u = 1 - t;
  return [
    u ** 3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t ** 3 * p3[0],
    u ** 3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t ** 3 * p3[1],
  ];
}

/** Outline of the shield in unit space, clockwise, as a dense polygon. */
function shieldOutline(inset = 0) {
  const cx = 0.5;
  const top = 0.105 + inset;
  const side = 0.885 - inset;         // right edge
  const shoulder = 0.20 + inset;      // where the rounded corner ends
  const straightTo = 0.47;            // sides run straight down to here
  const tip = 0.945 - inset;
  const cornerX = side - 0.095;

  const right = [[cx, top], [cornerX, top]];

  // rounded top-right corner
  for (let i = 1; i <= 8; i++) {
    right.push(quad([cornerX, top], [side, top], [side, shoulder], i / 8));
  }
  right.push([side, straightTo]);
  // sweep down to the bottom tip
  for (let i = 1; i <= 26; i++) {
    right.push(cubic([side, straightTo], [side, 0.73], [0.745, 0.855], [cx, tip], i / 26));
  }

  const left = right
    .slice(0, -1)
    .reverse()
    .map(([x, y]) => [2 * cx - x, y]);

  return [...right, ...left];
}

function makeInsideTest(polygon) {
  return (x, y) => {
    let inside = false;
    for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
      const [xi, yi] = polygon[i];
      const [xj, yj] = polygon[j];
      if (yi > y !== yj > y && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) inside = !inside;
    }
    return inside;
  };
}

function distanceToPolyline(points, x, y) {
  let best = Infinity;
  for (let i = 0; i < points.length - 1; i++) {
    const [ax, ay] = points[i];
    const [bx, by] = points[i + 1];
    const dx = bx - ax;
    const dy = by - ay;
    const t = Math.max(0, Math.min(1, ((x - ax) * dx + (y - ay) * dy) / (dx * dx + dy * dy)));
    best = Math.min(best, Math.hypot(x - (ax + dx * t), y - (ay + dy * t)));
  }
  return best;
}

/* -------------------------------- palette -------------------------------- */

const GRAD_A = [99, 116, 255];   // indigo, top-left
const GRAD_B = [147, 92, 246];   // violet, bottom-right
const RIM = [186, 199, 255];     // light rim so the mark reads on dark toolbars
const CHECK = [255, 255, 255];

const CHECK_PATH = [[0.305, 0.505], [0.442, 0.645], [0.705, 0.362]];

/* ------------------------------ rasteriser ------------------------------- */

function renderIcon(size) {
  const outer = makeInsideTest(shieldOutline(0));
  const rimInset = size >= 96 ? 0.026 : size >= 48 ? 0.034 : size >= 32 ? 0.042 : 0.055;
  const inner = makeInsideTest(shieldOutline(rimInset));
  const checkWidth = size >= 96 ? 0.052 : size >= 48 ? 0.058 : size >= 32 ? 0.067 : 0.078;

  const pixels = Buffer.alloc(size * size * 4);
  const step = 1 / (size * SS);

  for (let py = 0; py < size; py++) {
    for (let px = 0; px < size; px++) {
      let cover = 0;     // shield body (inner)
      let rim = 0;       // outline ring between outer and inner
      let check = 0;
      let rSum = 0;
      let gSum = 0;
      let bSum = 0;

      for (let sy = 0; sy < SS; sy++) {
        for (let sx = 0; sx < SS; sx++) {
          const x = (px * SS + sx + 0.5) * step;
          const y = (py * SS + sy + 0.5) * step;
          if (!outer(x, y)) continue;

          const isInner = inner(x, y);
          if (isInner) cover++;
          else rim++;

          // diagonal gradient, plus a soft radial sheen near the top-left
          const t = Math.max(0, Math.min(1, (x * 0.62 + y * 0.75) * 0.92));
          const sheen = Math.max(0, 1 - Math.hypot(x - 0.34, y - 0.26) / 0.62) ** 2 * 0.30;
          rSum += lerp(GRAD_A[0], GRAD_B[0], t) + (255 - GRAD_A[0]) * sheen;
          gSum += lerp(GRAD_A[1], GRAD_B[1], t) + (255 - GRAD_A[1]) * sheen;
          bSum += lerp(GRAD_A[2], GRAD_B[2], t) + (255 - GRAD_A[2]) * sheen;

          if (isInner && distanceToPolyline(CHECK_PATH, x, y) < checkWidth) check++;
        }
      }

      const samples = SS * SS;
      const filled = cover + rim;
      const offset = (py * size + px) * 4;
      if (!filled) continue;

      const bodyR = rSum / filled;
      const bodyG = gSum / filled;
      const bodyB = bSum / filled;

      // blend: gradient body -> rim tint -> white check
      const rimShare = rim / filled;
      let r = lerp(bodyR, RIM[0], rimShare);
      let g = lerp(bodyG, RIM[1], rimShare);
      let b = lerp(bodyB, RIM[2], rimShare);

      const checkShare = check / filled;
      r = lerp(r, CHECK[0], checkShare);
      g = lerp(g, CHECK[1], checkShare);
      b = lerp(b, CHECK[2], checkShare);

      pixels[offset] = Math.round(r);
      pixels[offset + 1] = Math.round(g);
      pixels[offset + 2] = Math.round(b);
      pixels[offset + 3] = Math.round((filled / samples) * 255);
    }
  }
  return pixels;
}

/* -------------------------------- encoder -------------------------------- */

const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = -1;
  for (const byte of buf) c = CRC_TABLE[(c ^ byte) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const head = Buffer.alloc(4);
  head.writeUInt32BE(data.length, 0);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const tail = Buffer.alloc(4);
  tail.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([head, body, tail]);
}

function encodePng(size, pixels) {
  const stride = size * 4;
  const raw = Buffer.alloc((stride + 1) * size);
  for (let y = 0; y < size; y++) {
    raw[y * (stride + 1)] = 0; // filter: none
    pixels.copy(raw, y * (stride + 1) + 1, y * stride, (y + 1) * stride);
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;  // bit depth
  ihdr[9] = 6;  // colour type: RGBA
  ihdr[10] = 0; // compression
  ihdr[11] = 0; // filter
  ihdr[12] = 0; // interlace

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

/* --------------------------------- main ---------------------------------- */

fs.mkdirSync(path.join(ROOT, 'icons'), { recursive: true });
for (const size of SIZES) {
  const file = path.join(ROOT, 'icons', `icon${size}.png`);
  fs.writeFileSync(file, encodePng(size, renderIcon(size)));
  process.stdout.write(`wrote icons/icon${size}.png\n`);
}
