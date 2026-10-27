// Generates the placeholder Rabbit Office icon (4 colored rounded squares) as a 512x512 PNG.
// node scripts/make-icon.js <out.png>
const fs = require('fs');
const zlib = require('zlib');
const S = 512, out = process.argv[2] || 'icon.png';
const colors = [[0x2f, 0x62, 0xc9], [0x1d, 0x8a, 0x52], [0xd2, 0x55, 0x2d], [0x7b, 0x4f, 0xd6]];
const px = Buffer.alloc(S * S * 4);
const pad = 40, gap = 24, cell = (S - pad * 2 - gap) / 2, r = 44;
function inRounded(x, y, x0, y0, w) {
  const cx = Math.min(Math.max(x, x0 + r), x0 + w - r), cy = Math.min(Math.max(y, y0 + r), y0 + w - r);
  const d = Math.hypot(x - cx, y - cy); return Math.max(0, Math.min(1, r - d + 0.5));
}
for (let y = 0; y < S; y++) for (let x = 0; x < S; x++) {
  let a = 0, c = colors[0];
  for (let i = 0; i < 4; i++) { const x0 = pad + (i % 2) * (cell + gap), y0 = pad + Math.floor(i / 2) * (cell + gap); const v = inRounded(x + 0.5, y + 0.5, x0, y0, cell); if (v > a) { a = v; c = colors[i]; } }
  const o = (y * S + x) * 4; px[o] = c[0]; px[o + 1] = c[1]; px[o + 2] = c[2]; px[o + 3] = Math.round(a * 255);
}
const raw = Buffer.alloc((S * 4 + 1) * S);
for (let y = 0; y < S; y++) { raw[y * (S * 4 + 1)] = 0; px.copy(raw, y * (S * 4 + 1) + 1, y * S * 4, (y + 1) * S * 4); }
const crcT = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
const crc = (b) => { let c = 0xffffffff; for (const v of b) c = crcT[(c ^ v) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
const chunk = (t, d) => { const l = Buffer.alloc(4); l.writeUInt32BE(d.length); const td = Buffer.concat([Buffer.from(t), d]); const c = Buffer.alloc(4); c.writeUInt32BE(crc(td)); return Buffer.concat([l, td, c]); };
const ihdr = Buffer.alloc(13); ihdr.writeUInt32BE(S, 0); ihdr.writeUInt32BE(S, 4); ihdr[8] = 8; ihdr[9] = 6;
fs.writeFileSync(out, Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]));
console.log('icon written:', out);
