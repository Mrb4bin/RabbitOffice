// Rabbit Xel — spreadsheet
import { createShell } from '../core/shell.js';
import { icon } from '../core/icons.js';
import { h, $, $$, esc, toast, dialog, menu, popover, closeMenus, colorPicker, promptDlg, confirmDlg, formDlg, debounce, uid, clamp } from '../core/ui.js';
import { TEMPLATES } from '../core/templates.js';
import { extOf, printToPdf, native } from '../core/platform.js';
import { renderChart, CHART_TYPES, chartDataFromGrid } from '../core/charts.js';
import { Engine, MAX_ROWS, MAX_COLS, colName, refName, parseRange, rangeName, shiftFormula, adjustFormula, renameSheetInFormula, parseDateInput, serialToDate, dateToSerial, fmtGeneral, isErr, FN_HELP, ALL_FN_NAMES, ALIAS, tokenize } from './formula.js';
import { exportXlsx, importXlsx } from './xlsx.js';

const HW = 46, HH = 22, DEF_W = 92, DEF_H = 22;
const DEFAULT_FONT = 'Calibri';
const K = (r, c) => r * MAX_COLS + c;
const RC = (k) => [Math.floor(k / MAX_COLS), k % MAX_COLS];

/* ======================= model ======================= */
const book = { sheets: [], active: 0 };
const engine = new Engine(book);
function newSheet(name) { return { name, cells: new Map(), colW: {}, rowH: {}, merges: [], freeze: { r: 0, c: 0 }, charts: [], cf: [], filter: null, color: null, gridlines: true, validations: [], print: {}, protected: false }; }
const sh = () => book.sheets[book.active];
let sel = { ar: 0, ac: 0, fr: 0, fc: 0 };
const rng = () => ({ r1: Math.min(sel.ar, sel.fr), c1: Math.min(sel.ac, sel.fc), r2: Math.max(sel.ar, sel.fr), c2: Math.max(sel.ac, sel.fc) });
const getCell = (r, c, s = sh()) => s.cells.get(K(r, c));
let Z = 1; let hiddenRows = new Set();

/* ======================= geometry ======================= */
const colX = new Float64Array(MAX_COLS + 1); const rowY = new Float64Array(MAX_ROWS + 1);
const colW = (c) => sh().colW[c] ?? DEF_W;
const rowH = (r) => (hiddenRows.has(r) ? 0 : sh().rowH[r] ?? DEF_H);
function rebuildGeom() {
  computeHidden();
  for (let c = 0; c < MAX_COLS; c++) colX[c + 1] = colX[c] + colW(c);
  for (let r = 0; r < MAX_ROWS; r++) rowY[r + 1] = rowY[r] + rowH(r);
  spacer.style.width = (HW + colX[MAX_COLS]) * Z + 'px'; spacer.style.height = (HH + rowY[MAX_ROWS]) * Z + 'px';
}
function bsearch(arr, v, n) { let lo = 0, hi = n; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (arr[m] <= v) lo = m; else hi = m - 1; } return Math.min(lo, n - 1); }
const colAt = (x) => bsearch(colX, x, MAX_COLS);
const rowAt = (y) => bsearch(rowY, y, MAX_ROWS);
const scrollX = () => scroller.scrollLeft / Z; const scrollY = () => scroller.scrollTop / Z;
function frz() { const f = sh().freeze; return { fr: f.r, fc: f.c, fx: colX[f.c], fy: rowY[f.r] }; }
function hitTest(clientX, clientY) {
  const b = host.getBoundingClientRect(); const lx = (clientX - b.left) / Z, ly = (clientY - b.top) / Z;
  const { fx, fy } = frz();
  const sx = lx - HW < fx ? lx - HW : lx - HW + scrollX();
  const sy = ly - HH < fy ? ly - HH : ly - HH + scrollY();
  return { lx, ly, sx, sy, c: colAt(Math.max(0, sx)), r: rowAt(Math.max(0, sy)), inColHdr: ly < HH, inRowHdr: lx < HW };
}
function cellScreenRect(r, c) {
  const { fr, fc } = frz();
  const x = HW + colX[c] - (c >= fc ? scrollX() : 0); const y = HH + rowY[r] - (r >= fr ? scrollY() : 0);
  const m = mergeAt(r, c); const w = m ? colX[m.c2 + 1] - colX[m.c1] : colW(c); const hh = m ? rowY[m.r2 + 1] - rowY[m.r1] : rowH(r);
  return { x, y, w, h: hh };
}
function mergeAt(r, c, s = sh()) { for (const m of s.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }

/* ======================= DOM ======================= */
const nameBox = h('input', { class: 'namebox', spellcheck: 'false', 'aria-label': 'Имя ячейки' });
const fInput = h('input', { class: 'finput', spellcheck: 'false', 'aria-label': 'Строка формул' });
const fxBtn = h('button', { class: 'fx-btn', 'data-tip': 'Вставить функцию|Shift+F3', onclick: () => insertFunctionDialog() }, 'fx');
const fbar = h('div', { class: 'fbar' }, nameBox, h('span', { class: 'sep' }), fxBtn, fInput);
const host = h('div', { class: 'grid-host' });
const canvas = h('canvas', { class: 'grid-canvas' });
const scroller = h('div', { class: 'grid-scroll', tabindex: '0', 'aria-label': 'Таблица' });
const spacer = h('div', { class: 'grid-spacer' }); scroller.append(spacer);
const overlay = h('div', { class: 'grid-overlay' });
// Invisible textarea that owns keyboard focus: receives typing (incl. IME / insertText) and clipboard events.
const sink = h('textarea', { class: 'key-sink', 'aria-label': 'Ввод в ячейку', inputmode: 'none', autocomplete: 'off', spellcheck: 'false' });
host.append(canvas, scroller, overlay, sink);
scroller.addEventListener('focus', () => sink.focus({ preventScroll: true }));
sink.addEventListener('input', (e) => {
  if (e.isComposing) return;
  const v = sink.value; sink.value = '';
  if (v && !editing) { editSheet = book.active; startEdit(v, 'enter'); }
});
sink.addEventListener('compositionend', () => { const v = sink.value; sink.value = ''; if (v && !editing) { editSheet = book.active; startEdit(v, 'enter'); } });
const tabsBar = h('div', { class: 'sheet-tabs' });
const wrap = h('div', { class: 'xel-wrap' }, fbar, host, tabsBar);
const ctx = canvas.getContext('2d');

/* ======================= styles & display ======================= */
let colors = {};
function readColors() { const cs = getComputedStyle(document.documentElement); for (const k of ['cell-bg', 'cell-text', 'gl', 'hdr-bg', 'hdr-text', 'hdr-sel', 'hdr-line', 'accent']) colors[k] = cs.getPropertyValue('--' + k).trim(); }
const fontOf = (s) => `${s?.i ? 'italic ' : ''}${s?.b ? 'bold ' : ''}${((s?.fs || 11) * 96) / 72}px "${s?.ff || DEFAULT_FONT}", Carlito, "Segoe UI", Arial, sans-serif`;
const nf = (dec, group = true) => new Intl.NumberFormat('ru-RU', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: group });
const nfCache = new Map(); const NF = (dec, g) => { const k = dec + '|' + g; if (!nfCache.has(k)) nfCache.set(k, nf(dec, g)); return nfCache.get(k); };
function formatValue(v, s) {
  if (v == null) return '';
  if (isErr(v)) return v.e;
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  if (typeof v !== 'number') return String(v);
  const f = s?.fmt || 'general';
  const dec = s?.dec;
  switch (f) {
    case 'number': return NF(dec ?? 2, true).format(v);
    case 'currency': { const sym = s.cur || '₽'; const t = NF(dec ?? 2, true).format(Math.abs(v)); return (v < 0 ? '-' : '') + (sym === '₽' ? t + ' ₽' : sym + t); }
    case 'percent': return NF(dec ?? 0, true).format(v * 100) + '%';
    case 'sci': return v.toExponential(dec ?? 2).replace('e', 'E');
    case 'date': { const d = serialToDate(v); return `${String(d.getUTCDate()).padStart(2, '0')}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${d.getUTCFullYear()}`; }
    case 'datelong': return serialToDate(v).toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
    case 'time': { const d = serialToDate(v); return `${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`; }
    case 'datetime': { const d = serialToDate(v); return `${String(d.getUTCDate()).padStart(2, '0')}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${d.getUTCFullYear()} ${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`; }
    case 'text': return fmtGeneral(v);
    default: if (dec != null) return NF(dec, false).format(v); { const g = fmtGeneral(v); return g.includes('E') ? g : g.replace('.', ','); }
  }
}
function displayOf(r, c, s = sh(), si = book.active) { const cell = s.cells.get(K(r, c)); if (!cell || cell.v == null || cell.v === '') return ''; return formatValue(engine.value(si, r, c), cell.s); }
// Parse user input into raw cell value (+ optional auto format)
function parseInput(text) {
  if (text == null) return { v: null };
  const t = String(text);
  if (t === '') return { v: null };
  if (t[0] === '=') return { v: t.length > 1 ? normalizeFormula(t) : t };
  if (t[0] === "'") return { v: t.slice(1), fmt: 'text' };
  const s = t.trim();
  if (/^[-+]?(\d{1,3}(\s\d{3})+|\d+)([.,]\d+)?([eE][-+]?\d+)?$/.test(s)) return { v: parseFloat(s.replace(/\s/g, '').replace(',', '.')) };
  let m = /^([-+]?[\d\s]*[.,]?\d+)\s*%$/.exec(s); if (m) return { v: parseFloat(m[1].replace(/\s/g, '').replace(',', '.')) / 100, fmt: 'percent' };
  m = /^([-+]?[\d\s]*[.,]?\d+)\s*(₽|р\.?|руб\.?)$/i.exec(s); if (m) return { v: parseFloat(m[1].replace(/\s/g, '').replace(',', '.')), fmt: 'currency', cur: '₽' };
  m = /^([$€£])\s*([-+]?[\d\s]*[.,]?\d+)$/.exec(s); if (m) return { v: parseFloat(m[2].replace(/\s/g, '').replace(',', '.')), fmt: 'currency', cur: m[1] };
  if (/^(true|истина)$/i.test(s)) return { v: true }; if (/^(false|ложь)$/i.test(s)) return { v: false };
  const d = parseDateInput(s); if (d) return { v: d.v, fmt: d.fmt };
  return { v: t };
}
function normalizeFormula(f) { try { tokenize(f.slice(1)); } catch { return f; } return f.replace(/[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*(?=\s*\()/g, (n) => { const u = n.toUpperCase(); return ALIAS[u] ? u : u; }); }

/* ======================= undo / redo ======================= */
const undoStack = []; const redoStack = [];
let tx = null;
function begin(struct = false) { if (!guard()) throw new ProtectedSheet(); tx = { si: book.active, before: new Map(), struct: struct ? snapSheet(sh()) : null, sel: { ...sel } }; }
function recordCell(k) { if (tx && !tx.struct && !tx.before.has(k)) { const c = sh().cells.get(k); tx.before.set(k, c ? structuredClone(c) : undefined); } }
function commit() {
  if (!tx) return;
  const t = tx; tx = null;
  if (t.struct) { t.after = snapSheet(book.sheets[t.si]); }
  else { if (!t.before.size) return afterChange(); t.after = new Map([...t.before.keys()].map((k) => { const c = book.sheets[t.si].cells.get(k); return [k, c ? structuredClone(c) : undefined]; })); }
  undoStack.push(t); if (undoStack.length > 200) undoStack.shift(); redoStack.length = 0;
  afterChange();
}
function applyEntry(t, which) {
  book.active = t.si; const s = sh();
  if (t.struct) restoreSheet(s, which === 'before' ? t.struct : t.after);
  else for (const [k, c] of t[which]) { if (c === undefined) s.cells.delete(k); else s.cells.set(k, structuredClone(c)); }
  sel = { ...t.sel };
}
function undo() { if (editing) return cancelEdit(); const t = undoStack.pop(); if (!t) return; applyEntry(t, 'before'); redoStack.push(t); renderTabs(); rebuildGeom(); afterChange(); }
function redo() { const t = redoStack.pop(); if (!t) return; applyEntry(t, 'after'); undoStack.push(t); renderTabs(); rebuildGeom(); afterChange(); }
function snapSheet(s) { return JSON.stringify(sheetToJSON(s)); }
function restoreSheet(s, json) { const d = sheetFromJSON(JSON.parse(json)); Object.assign(s, d); }

function setCell(r, c, patch, s = sh()) {
  const k = K(r, c); recordCell(k);
  const cur = s.cells.get(k);
  if (patch === null) { if (cur?.s || cur?.n) { const n = { ...cur }; delete n.v; s.cells.set(k, n); } else s.cells.delete(k); return; }
  const next = { ...(cur || {}), ...patch };
  if ((next.v == null || next.v === '') && !next.s && !next.n) s.cells.delete(k); else s.cells.set(k, next);
}
function setStyle(r, c, patch) {
  const k = K(r, c); recordCell(k);
  const cur = sh().cells.get(k) || {}; const s = { ...(cur.s || {}), ...patch };
  for (const key of Object.keys(s)) if (s[key] == null || s[key] === false) delete s[key];
  const next = { ...cur, s: Object.keys(s).length ? s : undefined }; if (!next.s) delete next.s;
  if ((next.v == null || next.v === '') && !next.s && !next.n) sh().cells.delete(k); else sh().cells.set(k, next);
}
function inputCell(r, c, text) {
  const p = parseInput(text);
  const cur = getCell(r, c);
  setCell(r, c, { v: p.v });
  if (p.fmt && (!cur?.s?.fmt || cur.s.fmt === 'general')) setStyle(r, c, { fmt: p.fmt, cur: p.cur });
}

/* ======================= rendering ======================= */
let rafPending = false;
function draw() { if (!rafPending) { rafPending = true; requestAnimationFrame(() => { rafPending = false; paint(); }); } }
let cfCache = new Map();
function paint() {
  const dpr = window.devicePixelRatio || 1; const W = host.clientWidth, H = host.clientHeight;
  if (canvas.width !== Math.round(W * dpr) || canvas.height !== Math.round(H * dpr)) { canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr); canvas.style.width = W + 'px'; canvas.style.height = H + 'px'; }
  ctx.setTransform(dpr * Z, 0, 0, dpr * Z, 0, 0);
  const lw = W / Z, lh = H / Z;
  ctx.fillStyle = colors['cell-bg']; ctx.fillRect(0, 0, lw, lh);
  cfCache = new Map();
  const { fx, fy } = frz(); const sx = scrollX(), sy = scrollY();
  drawRegion(HW + fx, HH + fy, lw - HW - fx, lh - HH - fy, fx + sx, fy + sy);
  if (fy) drawRegion(HW + fx, HH, lw - HW - fx, fy, fx + sx, 0);
  if (fx) drawRegion(HW, HH + fy, fx, lh - HH - fy, 0, fy + sy);
  if (fx && fy) drawRegion(HW, HH, fx, fy, 0, 0);
  drawHeaders(lw, lh, fx, fy, sx, sy);
  if (fx) { ctx.strokeStyle = colors['hdr-line']; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(HW + fx, 0); ctx.lineTo(HW + fx, lh); ctx.stroke(); }
  if (fy) { ctx.strokeStyle = colors['hdr-line']; ctx.lineWidth = 2; ctx.beginPath(); ctx.moveTo(0, HH + fy); ctx.lineTo(lw, HH + fy); ctx.stroke(); }
  positionEditor(); positionCharts();
}
function cfStyle(r, c, v) {
  let out = null;
  for (const rule of sh().cf) {
    const g = rule.range; if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2) continue;
    const t = rule.type;
    if (t === 'scale' || t === 'bar') {
      if (typeof v !== 'number') continue;
      let mm = cfCache.get(rule.id);
      if (!mm) { let mn = Infinity, mx = -Infinity; for (let rr = g.r1; rr <= Math.min(g.r2, g.r1 + 5000); rr++) for (let cc = g.c1; cc <= g.c2; cc++) { const x = engine.value(book.active, rr, cc); if (typeof x === 'number') { mn = Math.min(mn, x); mx = Math.max(mx, x); } } mm = { mn, mx }; cfCache.set(rule.id, mm); }
      const p = mm.mx > mm.mn ? (v - mm.mn) / (mm.mx - mm.mn) : 1;
      if (t === 'scale') out = { ...out, bg: mix(rule.c1 || '#f8696b', rule.c2 || '#63be7b', p) };
      else out = { ...out, bar: p, barColor: rule.color || '#5b9bd5' };
      continue;
    }
    let hit = false;
    const n = typeof v === 'number' ? v : null;
    if (t === 'gt') hit = n != null && n > +rule.a; else if (t === 'lt') hit = n != null && n < +rule.a;
    else if (t === 'between') hit = n != null && n >= +rule.a && n <= +rule.b; else if (t === 'eq') hit = String(v ?? '') == String(rule.a);
    else if (t === 'text') hit = String(v ?? '').toLowerCase().includes(String(rule.a).toLowerCase());
    else if (t === 'dup') { let dups = cfCache.get(rule.id); if (!dups) { const cnt = new Map(); for (let rr = g.r1; rr <= Math.min(g.r2, g.r1 + 5000); rr++) for (let cc = g.c1; cc <= g.c2; cc++) { const x = engine.value(book.active, rr, cc); if (x != null && x !== '') cnt.set(String(x), (cnt.get(String(x)) || 0) + 1); } dups = cnt; cfCache.set(rule.id, dups); } hit = v != null && v !== '' && dups.get(String(v)) > 1; }
    if (hit) out = { ...out, ...rule.style };
  }
  return out;
}
function mix(a, b, p) { const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16)), pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16)); return '#' + pa.map((x, i) => Math.round(x + (pb[i] - x) * p).toString(16).padStart(2, '0')).join(''); }

function drawRegion(X, Y, w, hgt, sx0, sy0) {
  if (w <= 0 || hgt <= 0) return;
  const s = sh();
  ctx.save(); ctx.beginPath(); ctx.rect(X, Y, w, hgt); ctx.clip();
  const c0 = colAt(sx0), r0 = rowAt(sy0);
  let c1 = c0; while (c1 < MAX_COLS - 1 && colX[c1 + 1] < sx0 + w) c1++;
  let r1 = r0; while (r1 < MAX_ROWS - 1 && rowY[r1 + 1] < sy0 + hgt) r1++;
  const px = (c) => X + colX[c] - sx0, py = (r) => Y + rowY[r] - sy0;
  // fills
  for (let r = r0; r <= r1; r++) {
    if (!rowH(r)) continue;
    for (let c = c0; c <= c1; c++) {
      const cell = s.cells.get(K(r, c)); const v = cell && cell.v != null ? engine.value(book.active, r, c) : null;
      const cf = s.cf.length ? cfStyle(r, c, v) : null;
      const bg = cf?.bg || cell?.s?.bg;
      if (bg) { ctx.fillStyle = bg; ctx.fillRect(px(c), py(r), colW(c), rowH(r)); }
      if (cf?.bar != null) { ctx.fillStyle = cf.barColor + 'aa'; ctx.fillRect(px(c) + 2, py(r) + 3, Math.max(2, (colW(c) - 4) * cf.bar), rowH(r) - 6); }
    }
  }
  // gridlines
  if (s.gridlines !== false) {
    ctx.strokeStyle = colors.gl; ctx.lineWidth = 1 / Z; ctx.beginPath();
    for (let c = c0; c <= c1 + 1; c++) { const x = Math.round(px(c) * Z) / Z + 0.5 / Z; ctx.moveTo(x, Y); ctx.lineTo(x, Y + hgt); }
    for (let r = r0; r <= r1 + 1; r++) { if (!rowH(r) && r <= r1) continue; const y = Math.round(py(r) * Z) / Z + 0.5 / Z; ctx.moveTo(X, y); ctx.lineTo(X + w, y); }
    ctx.stroke();
  }
  // merged cells cover
  for (const m of s.merges) {
    if (m.r2 < r0 || m.r1 > r1 || m.c2 < c0 || m.c1 > c1) continue;
    const cell = s.cells.get(K(m.r1, m.c1));
    ctx.fillStyle = cell?.s?.bg || colors['cell-bg'];
    ctx.fillRect(px(m.c1) + 1 / Z, py(m.r1) + 1 / Z, colX[m.c2 + 1] - colX[m.c1] - 1 / Z, rowY[m.r2 + 1] - rowY[m.r1] - 1 / Z);
  }
  // text
  ctx.textBaseline = 'middle';
  const drawn = new Set();
  for (let r = r0; r <= r1; r++) {
    if (!rowH(r)) continue;
    for (let c = c0; c <= c1; c++) {
      let rr = r, cc = c; const m = s.merges.length ? mergeAt(r, c) : null;
      if (m) { if (drawn.has(m)) continue; drawn.add(m); rr = m.r1; cc = m.c1; }
      const cell = s.cells.get(K(rr, cc));
      if (cell?.n) { const x2 = px(m ? m.c2 + 1 : c + 1); ctx.fillStyle = '#e03c31'; ctx.beginPath(); ctx.moveTo(x2 - 6, py(rr)); ctx.lineTo(x2, py(rr)); ctx.lineTo(x2, py(rr) + 6); ctx.fill(); }
      if (!cell || cell.v == null || cell.v === '') continue;
      const v = engine.value(book.active, rr, cc);
      const st = cell.s || {};
      const cf = s.cf.length ? cfStyle(rr, cc, v) : null;
      let text = formatValue(v, st);
      if (text === '') continue;
      const cw = m ? colX[m.c2 + 1] - colX[m.c1] : colW(c); const ch = m ? rowY[m.r2 + 1] - rowY[m.r1] : rowH(r);
      const x0 = px(cc), y0 = py(rr);
      ctx.font = fontOf(st);
      const link = linkOf(cell, v);
      ctx.fillStyle = link ? '#1a73e8' : cf?.fc || st.fc || (isErr(v) ? '#c42b1c' : colors['cell-text']);
      const al = st.al || (typeof v === 'number' && st.fmt !== 'text' ? 'right' : typeof v === 'boolean' || isErr(v) ? 'center' : 'left');
      const pad = 4;
      if (st.wr) {
        const lines = wrapText(text, cw - pad * 2); const lhh = ((st.fs || 11) * 96 / 72) * 1.2;
        let ty = st.va === 'top' ? y0 + 2 + lhh / 2 : st.va === 'middle' ? y0 + ch / 2 - (lines.length - 1) * lhh / 2 : y0 + ch - 3 - (lines.length - 0.5) * lhh;
        ctx.save(); ctx.beginPath(); ctx.rect(x0, y0, cw, ch); ctx.clip();
        for (const ln of lines) { ctx.textAlign = al; ctx.fillText(ln, al === 'left' ? x0 + pad : al === 'right' ? x0 + cw - pad : x0 + cw / 2, ty); ty += lhh; }
        ctx.restore();
      } else {
        const tw = ctx.measureText(text).width;
        let clipW = cw;
        if (typeof v === 'number' && tw > cw - pad && st.fmt !== 'text') { text = '#'.repeat(Math.max(1, Math.floor((cw - pad) / ctx.measureText('#').width))); }
        else if (tw > cw - pad && al === 'left' && !m) { let nc = c + 1; while (nc < MAX_COLS && clipW < tw + pad * 2 && !(s.cells.get(K(r, nc))?.v != null && s.cells.get(K(r, nc)).v !== '')) { clipW += colW(nc); nc++; } }
        const ty = st.va === 'top' ? y0 + ((st.fs || 11) * 96 / 72) * 0.7 + 2 : st.va === 'middle' ? y0 + ch / 2 + 0.5 : y0 + ch - ((st.fs || 11) * 96 / 72) * 0.55 - 2;
        ctx.save(); ctx.beginPath(); ctx.rect(x0, y0, clipW, ch); ctx.clip();
        ctx.textAlign = al;
        const tx = al === 'left' ? x0 + pad : al === 'right' ? x0 + cw - pad : x0 + cw / 2;
        ctx.fillText(text, tx, ty);
        if (st.u || st.s || link) { const w2 = ctx.measureText(text).width; const lx = al === 'left' ? tx : al === 'right' ? tx - w2 : tx - w2 / 2; ctx.fillRect(lx, st.u || link ? ty + 6 : ty, w2, 1); }
        ctx.restore();
      }
    }
  }
  // borders
  ctx.strokeStyle = '#1f1f1f';
  for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
    const bd = s.cells.get(K(r, c))?.s?.bd; if (!bd) continue;
    const x = px(c), y = py(r), w2 = colW(c), h2 = rowH(r);
    ctx.strokeStyle = bd.c || (colors['cell-bg'].startsWith('#1') || colors['cell-bg'].startsWith('#2') ? '#d0d0d0' : '#1f1f1f');
    const ln = (wd, x1, y1, x2, y2) => { if (!wd) return; ctx.lineWidth = wd; ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke(); };
    ln(bd.t, x, y + 0.5, x + w2, y + 0.5); ln(bd.b, x, y + h2 - 0.5 + (bd.b > 1 ? 0.5 : 0), x + w2, y + h2 - 0.5 + (bd.b > 1 ? 0.5 : 0)); ln(bd.l, x + 0.5, y, x + 0.5, y + h2); ln(bd.r, x + w2 - 0.5, y, x + w2 - 0.5, y + h2);
  }
  // filter buttons
  const f = s.filter;
  if (f && f.r >= r0 && f.r <= r1) for (let c = Math.max(c0, f.c1); c <= Math.min(c1, f.c2); c++) {
    const bx = px(c + 1) - 18, by = py(f.r) + 3; const active = f.crit?.[c]?.length;
    ctx.fillStyle = active ? colors.accent : colors['hdr-bg']; ctx.strokeStyle = colors['hdr-line']; ctx.lineWidth = 1;
    ctx.fillRect(bx, by, 15, rowH(f.r) - 6); ctx.strokeRect(bx + 0.5, by + 0.5, 14, rowH(f.r) - 7);
    ctx.fillStyle = active ? '#fff' : colors['hdr-text']; ctx.beginPath(); const cy = by + (rowH(f.r) - 6) / 2; ctx.moveTo(bx + 4, cy - 2); ctx.lineTo(bx + 11, cy - 2); ctx.lineTo(bx + 7.5, cy + 2.5); ctx.fill();
  }
  // copy marquee
  if (clip && clip.si === book.active) {
    const g = clip; ctx.save(); ctx.setLineDash([4, 3]); ctx.lineDashOffset = -marchOffset; ctx.strokeStyle = colors.accent; ctx.lineWidth = 1.5;
    ctx.strokeRect(px(g.c1) + 1, py(g.r1) + 1, colX[g.c2 + 1] - colX[g.c1] - 2, rowY[g.r2 + 1] - rowY[g.r1] - 2); ctx.restore();
  }
  // formula reference highlights
  if (editing && editRefs.length) {
    editRefs.forEach((g, i) => { if (g.si !== book.active) return; const col = REF_COLORS[i % REF_COLORS.length]; ctx.strokeStyle = col; ctx.lineWidth = 1.5; ctx.strokeRect(px(g.c1) + 1, py(g.r1) + 1, colX[g.c2 + 1] - colX[g.c1] - 2, rowY[g.r2 + 1] - rowY[g.r1] - 2); ctx.fillStyle = col + '1f'; ctx.fillRect(px(g.c1), py(g.r1), colX[g.c2 + 1] - colX[g.c1], rowY[g.r2 + 1] - rowY[g.r1]); });
  }
  // selection
  let g = expandMerged(rng());
  const sx = px(g.c1), sy = py(g.r1), sw = colX[g.c2 + 1] - colX[g.c1], shh = rowY[g.r2 + 1] - rowY[g.r1];
  if (g.r1 !== g.r2 || g.c1 !== g.c2) {
    const a = cellRectSheet(sel.ar, sel.ac);
    const ax = px(a.c1), ay = py(a.r1), aw = colX[a.c2 + 1] - colX[a.c1], ah = rowY[a.r2 + 1] - rowY[a.r1];
    // tint the range except the active cell
    ctx.save(); ctx.beginPath(); ctx.rect(sx, sy, sw, shh); ctx.rect(ax + aw, ay, -aw, ah); ctx.clip('evenodd');
    ctx.fillStyle = colors.accent + '22'; ctx.fillRect(sx, sy, sw, shh); ctx.restore();
  }
  ctx.strokeStyle = colors.accent; ctx.lineWidth = 2; ctx.strokeRect(sx + 0.5, sy + 0.5, sw - 1, shh - 1);
  if (!editing) { ctx.fillStyle = colors.accent; ctx.fillRect(sx + sw - 4, sy + shh - 4, 7, 7); ctx.strokeStyle = colors['cell-bg']; ctx.lineWidth = 1; ctx.strokeRect(sx + sw - 4, sy + shh - 4, 7, 7); }
  // dropdown arrow for list validation on the active cell
  if (!editing && validationAt(sel.ar, sel.ac)?.type === 'list') { const a = cellRectSheet(sel.ar, sel.ac); const bx = px(a.c2 + 1) + 1, by = py(a.r1); const bh = rowY[a.r2 + 1] - rowY[a.r1]; ctx.fillStyle = colors['hdr-bg']; ctx.fillRect(bx, by, 16, bh); ctx.strokeStyle = colors['hdr-line']; ctx.strokeRect(bx + 0.5, by + 0.5, 15, bh - 1); ctx.fillStyle = colors['hdr-text']; ctx.beginPath(); ctx.moveTo(bx + 4, by + bh / 2 - 2); ctx.lineTo(bx + 12, by + bh / 2 - 2); ctx.lineTo(bx + 8, by + bh / 2 + 3); ctx.fill(); }
  if (fillDrag) { const fg = fillDrag.target; ctx.setLineDash([3, 3]); ctx.strokeStyle = colors['hdr-text']; ctx.lineWidth = 1; ctx.strokeRect(px(fg.c1) + 0.5, py(fg.r1) + 0.5, colX[fg.c2 + 1] - colX[fg.c1] - 1, rowY[fg.r2 + 1] - rowY[fg.r1] - 1); ctx.setLineDash([]); }
  ctx.restore();
}
function cellRectSheet(r, c) { const m = mergeAt(r, c); return m || { r1: r, c1: c, r2: r, c2: c }; }
function expandMerged(g) {
  if (!sh().merges.length) return g; let changed = true; g = { ...g };
  while (changed) { changed = false; for (const m of sh().merges) { if (m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2) continue; if (m.r1 < g.r1) { g.r1 = m.r1; changed = true; } if (m.c1 < g.c1) { g.c1 = m.c1; changed = true; } if (m.r2 > g.r2) { g.r2 = m.r2; changed = true; } if (m.c2 > g.c2) { g.c2 = m.c2; changed = true; } } }
  return g;
}
function wrapText(text, maxW) {
  const out = [];
  for (const para of String(text).split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) { const t = line + word; if (ctx.measureText(t).width > maxW && line.trim()) { out.push(line.trimEnd()); line = word.trimStart(); } else line = t; }
    out.push(line);
  }
  return out;
}
function drawHeaders(lw, lh, fx, fy, sx, sy) {
  const g = rng(); const s = sh();
  ctx.font = `12px "Segoe UI", system-ui, sans-serif`; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  // column header
  ctx.fillStyle = colors['hdr-bg']; ctx.fillRect(0, 0, lw, HH); ctx.fillRect(0, 0, HW, lh);
  const colHdr = (X, w, sx0) => {
    ctx.save(); ctx.beginPath(); ctx.rect(X, 0, w, HH); ctx.clip();
    for (let c = colAt(sx0); c < MAX_COLS && colX[c] < sx0 + w; c++) {
      const x = X + colX[c] - sx0; const cw = colW(c); if (!cw) continue;
      const on = c >= g.c1 && c <= g.c2;
      if (on) { ctx.fillStyle = colors['hdr-sel']; ctx.fillRect(x, 0, cw, HH); ctx.fillStyle = colors.accent; ctx.fillRect(x, HH - 2, cw, 2); }
      ctx.fillStyle = on ? colors.accent : colors['hdr-text']; ctx.fillText(colName(c), x + cw / 2, HH / 2 + 1);
      ctx.fillStyle = colors['hdr-line']; ctx.fillRect(x + cw - 1 / Z, 0, 1 / Z, HH);
    }
    ctx.restore();
  };
  const rowHdr = (Y, hgt, sy0) => {
    ctx.save(); ctx.beginPath(); ctx.rect(0, Y, HW, hgt); ctx.clip();
    for (let r = rowAt(sy0); r < MAX_ROWS && rowY[r] < sy0 + hgt; r++) {
      const y = Y + rowY[r] - sy0; const rh = rowH(r); if (!rh) continue;
      const on = r >= g.r1 && r <= g.r2;
      if (on) { ctx.fillStyle = colors['hdr-sel']; ctx.fillRect(0, y, HW, rh); ctx.fillStyle = colors.accent; ctx.fillRect(HW - 2, y, 2, rh); }
      ctx.fillStyle = on ? colors.accent : hiddenRowsBefore(r) ? colors.accent : colors['hdr-text']; ctx.fillText(String(r + 1), HW / 2, y + rh / 2 + 1);
      ctx.fillStyle = colors['hdr-line']; ctx.fillRect(0, y + rh - 1 / Z, HW, 1 / Z);
    }
    ctx.restore();
  };
  colHdr(HW + fx, lw - HW - fx, fx + sx); if (fx) colHdr(HW, fx, 0);
  rowHdr(HH + fy, lh - HH - fy, fy + sy); if (fy) rowHdr(HH, fy, 0);
  ctx.fillStyle = colors['hdr-bg']; ctx.fillRect(0, 0, HW, HH);
  ctx.fillStyle = colors['hdr-line']; ctx.fillRect(0, HH - 1 / Z, lw, 1 / Z); ctx.fillRect(HW - 1 / Z, 0, 1 / Z, lh);
  ctx.fillStyle = colors['hdr-text']; ctx.beginPath(); ctx.moveTo(HW - 5, HH - 5); ctx.lineTo(HW - 5, HH - 13); ctx.lineTo(HW - 13, HH - 5); ctx.fill();
}
const hiddenRowsBefore = (r) => hiddenRows.size && hiddenRows.has(r - 1);

/* ======================= filter ======================= */
function computeHidden() {
  hiddenRows = new Set();
  const s = sh(); const f = s.filter;
  for (const [r, v] of Object.entries(s.rowH)) if (v === 0) hiddenRows.add(+r);
  if (!f || !f.crit) return;
  const last = lastDataRow(f.c1, f.c2, f.r);
  for (let r = f.r + 1; r <= last; r++) {
    let hide = false;
    for (const [c, ex] of Object.entries(f.crit)) { if (!ex?.length) continue; const d = displayOf(r, +c); if (ex.includes(d)) { hide = true; break; } }
    if (!hide && f.cond) for (const [c, cond] of Object.entries(f.cond)) { if (cond && !condMatch(engine.value(book.active, r, +c), cond)) { hide = true; break; } }
    if (hide) hiddenRows.add(r);
  }
}
const FILTER_OPS = [['eq', 'равно'], ['ne', 'не равно'], ['gt', 'больше'], ['ge', 'больше или равно'], ['lt', 'меньше'], ['le', 'меньше или равно'], ['between', 'между'], ['contains', 'содержит'], ['notcontains', 'не содержит'], ['starts', 'начинается с'], ['ends', 'заканчивается на'], ['top', 'первые N наибольших'], ['above', 'выше среднего'], ['below', 'ниже среднего']];
function condMatch(v, c) {
  const num = (x) => +String(x).replace(',', '.'); const s = String(v ?? '').toLowerCase(); const a = String(c.a ?? '').toLowerCase();
  const isNum = typeof v === 'number' && c.a !== '' && isFinite(num(c.a));
  switch (c.op) {
    case 'eq': return isNum ? v === num(c.a) : s === a;
    case 'ne': return isNum ? v !== num(c.a) : s !== a;
    case 'gt': return typeof v === 'number' && v > num(c.a); case 'ge': return typeof v === 'number' && v >= num(c.a);
    case 'lt': return typeof v === 'number' && v < num(c.a); case 'le': return typeof v === 'number' && v <= num(c.a);
    case 'between': return typeof v === 'number' && v >= num(c.a) && v <= num(c.b);
    case 'contains': return s.includes(a); case 'notcontains': return !s.includes(a); case 'starts': return s.startsWith(a); case 'ends': return s.endsWith(a);
    case 'top': return typeof v === 'number' && v >= (c.threshold ?? -Infinity);
    case 'above': return typeof v === 'number' && v > (c.avg ?? 0); case 'below': return typeof v === 'number' && v < (c.avg ?? 0);
  }
  return true;
}
function lastDataRow(c1, c2, from = 0) { let last = from; for (const k of sh().cells.keys()) { const [r, c] = RC(k); if (c >= c1 && c <= c2 && r > last && sh().cells.get(k).v != null) last = r; } return last; }
function currentRegion(r, c) {
  const s = sh(); const has = (rr, cc) => rr >= 0 && cc >= 0 && s.cells.get(K(rr, cc))?.v != null && s.cells.get(K(rr, cc)).v !== '';
  let g = { r1: r, c1: c, r2: r, c2: c }; let grow = true;
  while (grow) { grow = false;
    for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) { if (has(g.r1 - 1, cc)) { g.r1--; grow = true; break; } } for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) { if (has(g.r2 + 1, cc)) { g.r2++; grow = true; break; } }
    for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) { if (has(rr, g.c1 - 1)) { g.c1--; grow = true; break; } } for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) { if (has(rr, g.c2 + 1)) { g.c2++; grow = true; break; } }
    if (g.r2 - g.r1 > 5000 || g.c2 - g.c1 > 200) break;
  }
  return g;
}
function toggleFilter() {
  const s = sh();
  begin(true);
  if (s.filter) { s.filter = null; }
  else { const g0 = rng(); const g = g0.r1 === g0.r2 && g0.c1 === g0.c2 ? currentRegion(sel.ar, sel.ac) : g0; s.filter = { r: g.r1, c1: g.c1, c2: g.c2, crit: {} }; }
  commit(); rebuildGeom(); draw();
}
function filterPopup(c, clientX, clientY) {
  const f = sh().filter; const last = lastDataRow(f.c1, f.c2, f.r);
  const vals = new Map(); for (let r = f.r + 1; r <= last; r++) { const d = displayOf(r, c); vals.set(d, (vals.get(d) || 0) + 1); }
  const ex = new Set(f.crit[c] || []);
  const box = h('div', { class: 'filter-pop' });
  const list = h('div', { class: 'vals' });
  const checks = [];
  const all = h('input', { type: 'checkbox', checked: !ex.size }); all.onchange = () => checks.forEach((x) => (x.el.checked = all.checked));
  list.append(h('label', { class: 'check' }, all, h('b', {}, '(Выделить все)')));
  for (const v of [...vals.keys()].sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }))) { const cb = h('input', { type: 'checkbox', checked: !ex.has(v) }); checks.push({ el: cb, v }); list.append(h('label', { class: 'check' }, cb, v === '' ? '(Пустые)' : v)); }
  box.append(h('button', { class: 'menu-item', html: icon('sort-asc') + '<span class="lbl">Сортировка от А до Я</span>', onclick: () => { closeMenus(); sortRange({ r1: f.r, c1: f.c1, r2: last, c2: f.c2 }, c, true, true); } }),
    h('button', { class: 'menu-item', html: icon('sort-desc') + '<span class="lbl">Сортировка от Я до А</span>', onclick: () => { closeMenus(); sortRange({ r1: f.r, c1: f.c1, r2: last, c2: f.c2 }, c, false, true); } }),
    h('button', { class: 'menu-item', html: icon('filter') + '<span class="lbl">Фильтр по условию…</span>', onclick: () => { closeMenus(); customFilter(c); } }),
    f.cond?.[c] ? h('button', { class: 'menu-item', html: icon('eraser') + '<span class="lbl">Снять условие</span>', onclick: () => { closeMenus(); begin(true); delete f.cond[c]; commit(); rebuildGeom(); draw(); } }) : null,
    h('div', { class: 'menu-sep' }), list,
    h('div', { class: 'row', style: { justifyContent: 'flex-end' } }, h('button', { class: 'btn sm', onclick: closeMenus }, 'Отмена'), h('button', { class: 'btn sm primary', onclick: () => { begin(true); f.crit[c] = checks.filter((x) => !x.el.checked).map((x) => x.v); commit(); closeMenus(); rebuildGeom(); draw(); } }, 'OK')));
  popover({ x: clientX, y: clientY }, box);
}

/* ======================= sort / dedupe ======================= */
function sortRange(g, col, asc, hasHeader) {
  begin(true);
  const s = sh(); const start = hasHeader ? g.r1 + 1 : g.r1;
  const rows = [];
  for (let r = start; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) cells.push(s.cells.get(K(r, c))); rows.push({ r, cells, key: engine.value(book.active, r, col) }); }
  rows.sort((a, b) => { const x = a.key, y = b.key; if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1; const cmp = typeof x === 'number' && typeof y === 'number' ? x - y : typeof x === 'number' ? -1 : typeof y === 'number' ? 1 : String(x).localeCompare(String(y), 'ru', { numeric: true }); return asc ? cmp : -cmp; });
  rows.forEach((row, i) => { const r = start + i; row.cells.forEach((cell, j) => { const k = K(r, g.c1 + j); if (!cell) s.cells.delete(k); else { const n = { ...cell }; if (typeof n.v === 'string' && n.v[0] === '=') n.v = shiftFormula(n.v, r - row.r, 0); s.cells.set(k, n); } }); });
  commit();
}
function sortDialog() {
  const g0 = rng(); const g = g0.r1 === g0.r2 ? currentRegion(sel.ar, sel.ac) : g0;
  const cols = []; for (let c = g.c1; c <= g.c2; c++) cols.push({ value: c, label: `Столбец ${colName(c)}${displayOf(g.r1, c) ? ' — ' + displayOf(g.r1, c) : ''}` });
  formDlg('Сортировка', [{ key: 'col', label: 'Сортировать по', type: 'select', options: cols, value: sel.ac }, { key: 'order', label: 'Порядок', type: 'select', options: [{ value: 'asc', label: 'По возрастанию (А → Я, 1 → 9)' }, { value: 'desc', label: 'По убыванию (Я → А, 9 → 1)' }] }, { key: 'hdr', label: 'Мои данные содержат заголовки', type: 'checkbox', value: true }], 'Сортировать')
    .then((r) => { if (r) sortRange(g, +r.col, r.order === 'asc', r.hdr); });
}
function removeDuplicates() {
  const g0 = rng(); const g = g0.r1 === g0.r2 ? currentRegion(sel.ar, sel.ac) : g0;
  begin(true);
  const seen = new Set(); const keep = []; const s = sh();
  for (let r = g.r1; r <= g.r2; r++) { const key = []; for (let c = g.c1; c <= g.c2; c++) key.push(displayOf(r, c)); const k = key.join('\u0001'); if (seen.has(k)) continue; seen.add(k); const cells = []; for (let c = g.c1; c <= g.c2; c++) cells.push(s.cells.get(K(r, c))); keep.push(cells); }
  const removed = g.r2 - g.r1 + 1 - keep.length;
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) s.cells.delete(K(r, c));
  keep.forEach((cells, i) => cells.forEach((cell, j) => { if (cell) s.cells.set(K(g.r1 + i, g.c1 + j), cell); }));
  commit(); toast(`Удалено дубликатов: ${removed}`, 'ok');
}

/* ======================= structure ops ======================= */
function shiftStructure(axis, at, count) {
  // count > 0 insert, < 0 delete
  const s = sh(); const name = s.name;
  const moved = new Map();
  for (const [k, cell] of s.cells) {
    let [r, c] = RC(k); const v = axis === 'row' ? r : c;
    if (count < 0 && v >= at && v < at - count) continue;
    if (v >= at) { if (axis === 'row') r += count; else c += count; }
    if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) continue;
    moved.set(K(r, c), cell);
  }
  s.cells = moved;
  const shiftObj = (o) => { const n = {}; for (const [k, v] of Object.entries(o)) { const i = +k; if (count < 0 && i >= at && i < at - count) continue; n[i >= at ? i + count : i] = v; } return n; };
  if (axis === 'row') s.rowH = shiftObj(s.rowH); else s.colW = shiftObj(s.colW);
  for (const sheet of book.sheets) for (const cell of sheet.cells.values()) if (typeof cell.v === 'string' && cell.v[0] === '=') cell.v = adjustFormula(cell.v, { sheet: name, isSameSheet: sheet === s, axis, at, count });
  const k1 = axis === 'row' ? 'r1' : 'c1', k2 = axis === 'row' ? 'r2' : 'c2';
  s.merges = s.merges.map((m) => ({ ...m, [k1]: m[k1] >= at ? m[k1] + count : m[k1], [k2]: m[k2] >= at ? m[k2] + count : m[k2] })).filter((m) => m[k2] >= m[k1] && m[k1] >= 0);
  s.cf.forEach((rule) => { const g = rule.range; if (g[k1] >= at) g[k1] = Math.max(at, g[k1] + count); if (g[k2] >= at) g[k2] = Math.max(g[k1], g[k2] + count); });
}
function insertRows(before = true) { const g = rng(); const n = g.r2 - g.r1 + 1; begin(true); shiftStructure('row', before ? g.r1 : g.r2 + 1, n); commit(); rebuildGeom(); }
function insertCols(before = true) { const g = rng(); const n = g.c2 - g.c1 + 1; begin(true); shiftStructure('col', before ? g.c1 : g.c2 + 1, n); commit(); rebuildGeom(); }
function deleteRows() { const g = rng(); begin(true); shiftStructure('row', g.r1, -(g.r2 - g.r1 + 1)); commit(); rebuildGeom(); }
function deleteCols() { const g = rng(); begin(true); shiftStructure('col', g.c1, -(g.c2 - g.c1 + 1)); commit(); rebuildGeom(); }
function mergeCells(mode = 'center') {
  const g = rng(); const s = sh();
  begin(true);
  const existing = s.merges.filter((m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2));
  if (existing.length) { s.merges = s.merges.filter((m) => !existing.includes(m)); commit(); return; }
  if (g.r1 === g.r2 && g.c1 === g.c2) { tx = null; return; }
  if (mode === 'across') { for (let r = g.r1; r <= g.r2; r++) s.merges.push({ r1: r, c1: g.c1, r2: r, c2: g.c2 }); }
  else { s.merges.push({ ...g }); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== g.r1 || c !== g.c1) { const cell = s.cells.get(K(r, c)); if (cell) { delete cell.v; if (!cell.s) s.cells.delete(K(r, c)); } } }
  if (mode === 'center') { const cell = s.cells.get(K(g.r1, g.c1)) || {}; cell.s = { ...(cell.s || {}), al: 'center', va: 'middle' }; s.cells.set(K(g.r1, g.c1), cell); }
  commit();
}

/* ======================= after change ======================= */
const statusSum = h('span'); const zoomLbl = h('button', { class: 'sb-btn', onclick: () => setZoom(1) }, '100%');
function afterChange() {
  engine.invalidate(); computeHiddenIfFilter(); draw(); syncBars(); renderCharts(); shell?.markDirty();
}
function computeHiddenIfFilter() { if (sh().filter || Object.values(sh().rowH).includes(0) || hiddenRows.size) rebuildGeom(); }
function syncBars() {
  const g = rng();
  if (document.activeElement !== nameBox) nameBox.value = g.r1 === g.r2 && g.c1 === g.c2 || isMergeExact(g) ? refName(sel.ar, sel.ac) : rangeName(g);
  const cell = getCell(sel.ar, sel.ac);
  if (!editing && document.activeElement !== fInput) fInput.value = rawToEdit(cell);
  // status
  let sum = 0, cnt = 0, nums = 0, mn = Infinity, mx = -Infinity;
  const area = (g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1);
  if (area > 1 && area < 200000) {
    for (const k of sh().cells.keys()) { const [r, c] = RC(k); if (r < g.r1 || r > g.r2 || c < g.c1 || c > g.c2 || hiddenRows.has(r)) continue; const v = engine.value(book.active, r, c); if (v == null || v === '') continue; cnt++; if (typeof v === 'number') { sum += v; nums++; mn = Math.min(mn, v); mx = Math.max(mx, v); } }
  }
  const f = (v) => formatValue(v, getCell(sel.ar, sel.ac)?.s?.fmt === 'percent' || getCell(sel.ar, sel.ac)?.s?.fmt === 'currency' ? getCell(sel.ar, sel.ac).s : { dec: Math.abs(v) < 1000 && !Number.isInteger(v) ? 2 : undefined });
  statusSum.textContent = cnt > 1 ? (nums ? `Среднее: ${f(sum / nums)}   Количество: ${cnt}   Сумма: ${f(sum)}   Мин: ${f(mn)}   Макс: ${f(mx)}` : `Количество: ${cnt}`) : '';
  // ribbon state
  const st = cell?.s || {};
  if (shell) {
    shell.ribbon.set('bold', { on: !!st.b }); shell.ribbon.set('italic', { on: !!st.i }); shell.ribbon.set('underline', { on: !!st.u }); shell.ribbon.set('strike', { on: !!st.s });
    shell.ribbon.set('al-l', { on: st.al === 'left' }); shell.ribbon.set('al-c', { on: st.al === 'center' }); shell.ribbon.set('al-r', { on: st.al === 'right' });
    shell.ribbon.set('va-t', { on: st.va === 'top' }); shell.ribbon.set('va-m', { on: st.va === 'middle' }); shell.ribbon.set('va-b', { on: !st.va || st.va === 'bottom' });
    shell.ribbon.set('wrap', { on: !!st.wr }); shell.ribbon.set('merge', { on: !!mergeAt(sel.ar, sel.ac) });
    shell.ribbon.set('font', { value: st.ff || DEFAULT_FONT }); shell.ribbon.set('size', { value: String(st.fs || 11) }); shell.ribbon.set('fmt', { value: st.fmt || 'general' });
    shell.ribbon.set('filter', { on: !!sh().filter }); shell.ribbon.set('gridlines', { on: sh().gridlines !== false }); shell.ribbon.set('protect', { on: !!sh().protected });
    shell.undoBtn.disabled = !undoStack.length; shell.redoBtn.disabled = !redoStack.length;
  }
}
const isMergeExact = (g) => { const m = mergeAt(g.r1, g.c1); return m && m.r1 === g.r1 && m.c1 === g.c1 && m.r2 === g.r2 && m.c2 === g.c2; };
function rawToEdit(cell) {
  if (!cell || cell.v == null) return '';
  const v = cell.v;
  if (typeof v === 'number') { const f = cell.s?.fmt; if (f === 'date') return formatValue(v, { fmt: 'date' }); if (f === 'datetime') return formatValue(v, { fmt: 'datetime' }); if (f === 'time') return formatValue(v, { fmt: 'time' }); if (f === 'percent') return fmtGeneral(+(v * 100).toPrecision(12)).replace('.', ',') + '%'; return fmtGeneral(v).replace('.', ','); }
  if (typeof v === 'boolean') return v ? 'ИСТИНА' : 'ЛОЖЬ';
  return String(v);
}

/* ======================= selection & navigation ======================= */
function select(r, c, extend = false) {
  r = clamp(r, 0, MAX_ROWS - 1); c = clamp(c, 0, MAX_COLS - 1);
  if (extend) { sel.fr = r; sel.fc = c; } else { const m = mergeAt(r, c); sel = m ? { ar: m.r1, ac: m.c1, fr: m.r2, fc: m.c2 } : { ar: r, ac: c, fr: r, fc: c }; }
  ensureVisible(extend ? sel.fr : sel.ar, extend ? sel.fc : sel.ac);
  draw(); syncBars(); selectChart(null);
}
function ensureVisible(r, c) {
  const { fr, fc, fx, fy } = frz(); const W = host.clientWidth / Z, H = host.clientHeight / Z;
  if (c >= fc) { const x = colX[c], x2 = colX[c + 1]; const vx = scrollX() + fx, vw = W - HW - fx - 14; if (x < vx) scroller.scrollLeft = (x - fx) * Z; else if (x2 > vx + vw) scroller.scrollLeft = (x2 - fx - vw) * Z; }
  if (r >= fr) { const y = rowY[r], y2 = rowY[r + 1]; const vy = scrollY() + fy, vh = H - HH - fy - 14; if (y < vy) scroller.scrollTop = (y - fy) * Z; else if (y2 > vy + vh) scroller.scrollTop = (y2 - fy - vh) * Z; }
}
function move(dr, dc, extend, jump) {
  let r = extend ? sel.fr : sel.ar, c = extend ? sel.fc : sel.ac;
  if (!extend) { const m = mergeAt(r, c); if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; } }
  if (jump) {
    const has = (rr, cc) => getCell(rr, cc)?.v != null && getCell(rr, cc).v !== '';
    const step = () => { r += dr; c += dc; };
    const inb = () => r + dr >= 0 && r + dr < MAX_ROWS && c + dc >= 0 && c + dc < MAX_COLS;
    if (has(r, c) && inb() && has(r + dr, c + dc)) { while (inb() && has(r + dr, c + dc)) step(); }
    else { step(); while (inb() && !has(r, c)) step(); if (!has(r, c) && !inb()) { /* edge */ } }
  } else { r += dr; c += dc; while (dr && r > 0 && r < MAX_ROWS - 1 && !rowH(r)) r += dr; }
  select(r, c, extend);
}
function selectAll() { sel = { ar: 0, ac: 0, fr: MAX_ROWS - 1, fc: MAX_COLS - 1 }; draw(); syncBars(); }
function selRangeForOps() { const g = rng(); if (g.r2 - g.r1 > 2000 || g.c2 - g.c1 > 100) { const used = engine.usedBounds(book.active); return { r1: g.r1, c1: g.c1, r2: Math.min(g.r2, used.mr + 50), c2: Math.min(g.c2, used.mc + 10) }; } return g; }

/* ======================= cell editor ======================= */
let editing = false; let editMode = 'enter'; let editor = null; let editRefs = []; let refInsert = null;
const REF_COLORS = ['#3b6fe0', '#e8702a', '#2fa36b', '#8a5cd6', '#d64545', '#1aa3b8'];
function startEdit(initial = null, mode = 'enter') {
  if (editing || !guard()) return;
  editing = true; editMode = mode;
  const cell = getCell(sel.ar, sel.ac);
  editor = h('textarea', { class: 'cell-editor', spellcheck: 'false', rows: 1 });
  editor.value = initial != null ? initial : rawToEdit(cell);
  const st = cell?.s || {};
  editor.style.font = fontOf(st).replace(/^([^\d]*)([\d.]+)px/, (m, a, n) => `${a}${n * Z}px`);
  editor.style.textAlign = st.al || 'left';
  overlay.append(editor);
  positionEditor();
  editor.focus();
  editor.setSelectionRange(editor.value.length, editor.value.length);
  editor.addEventListener('input', onEditInput);
  editor.addEventListener('keydown', onEditKey);
  editor.addEventListener('blur', () => setTimeout(() => { if (editing && document.activeElement !== editor && document.activeElement !== fInput && !acBox) commitEdit(); }, 120));
  onEditInput();
}
function positionEditor() {
  if (!editor) return;
  const r = cellScreenRect(sel.ar, sel.ac);
  editor.style.left = (r.x * Z - 1) + 'px'; editor.style.top = (r.y * Z - 1) + 'px';
  editor.style.minWidth = (r.w * Z + 2) + 'px'; editor.style.height = 'auto'; editor.style.minHeight = (r.h * Z + 2) + 'px';
  editor.style.width = 'auto';
  const meas = Math.max(r.w * Z + 2, Math.min(host.clientWidth - r.x * Z - 4, measureEditor() + 24));
  editor.style.width = meas + 'px';
  editor.style.height = Math.max(r.h * Z + 2, editor.scrollHeight + 2) + 'px';
}
function measureEditor() { ctx.save(); ctx.font = editor.style.font; const w = Math.max(...editor.value.split('\n').map((l) => ctx.measureText(l).width)); ctx.restore(); return w; }
function onEditInput() {
  fInput.value = editor.value;
  editor.style.height = 'auto'; positionEditor();
  refInsert = null;
  updateEditRefs(editor.value); autocomplete(editor);
}
function updateEditRefs(v) {
  editRefs = [];
  if (v[0] === '=') { try { for (const t of tokenize(v.slice(1))) if (t.type === 'ref') { const g = parseRange(t.ref.replace(/\$/g, '')); if (g) editRefs.push({ ...g, si: t.sheet ? book.sheets.findIndex((s) => s.name.toLowerCase() === t.sheet.toLowerCase()) : editSheet }); } } catch { /* partial formula */ } }
  draw();
}
let editSheet = 0;
function canInsertRef() {
  const src = editing ? editor : fInput; const v = src.value; if (v[0] !== '=') return false;
  if (refInsert) return true;
  const before = v.slice(0, src.selectionStart).trimEnd();
  return /[=(,;+\-*/^&<>:]$/.test(before);
}
function insertRefText(text) {
  const src = editing ? editor : fInput;
  const v = src.value; let s = src.selectionStart, e = src.selectionEnd;
  if (refInsert) { s = refInsert.start; e = refInsert.end; }
  src.value = v.slice(0, s) + text + v.slice(e);
  refInsert = { start: s, end: s + text.length };
  src.setSelectionRange(s + text.length, s + text.length);
  if (editing) { fInput.value = editor.value; positionEditor(); } else editorProxy();
  updateEditRefs(src.value);
  const keep = refInsert; refInsert = keep;
}
function onEditKey(e) {
  if (acBox && handleAcKey(e)) return;
  const k = e.key;
  if (k === 'Enter' && e.altKey) { e.preventDefault(); const p = editor.selectionStart; editor.value = editor.value.slice(0, p) + '\n' + editor.value.slice(editor.selectionEnd); editor.setSelectionRange(p + 1, p + 1); onEditInput(); return; }
  if (k === 'Enter') { e.preventDefault(); if (e.ctrlKey) { commitEdit(true); return; } if (commitEdit()) move(e.shiftKey ? -1 : 1, 0); return; }
  if (k === 'Tab') { e.preventDefault(); if (commitEdit()) move(0, e.shiftKey ? -1 : 1); return; }
  if (k === 'Escape') { e.preventDefault(); cancelEdit(); return; }
  if (editMode === 'enter' && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(k) && !(editor.value[0] === '=' && canInsertRef())) {
    e.preventDefault(); if (!commitEdit()) return; move(k === 'ArrowUp' ? -1 : k === 'ArrowDown' ? 1 : 0, k === 'ArrowLeft' ? -1 : k === 'ArrowRight' ? 1 : 0); return;
  }
  if (k === 'F4') { e.preventDefault(); toggleAbsolute(editor); }
}
function toggleAbsolute(src) {
  const v = src.value; const p = src.selectionStart;
  const re = /\$?[A-Za-z]{1,3}\$?\d+/g; let m;
  while ((m = re.exec(v))) { if (p >= m.index && p <= m.index + m[0].length) { const raw = m[0].replace(/\$/g, ''); const [, col, row] = /([A-Za-z]+)(\d+)/.exec(raw); const cur = m[0]; const cycle = [`$${col}$${row}`, `${col}$${row}`, `$${col}${row}`, `${col}${row}`]; const idx = cycle.indexOf(cur); const next = cycle[(idx + 1) % 4]; src.value = v.slice(0, m.index) + next + v.slice(m.index + cur.length); src.setSelectionRange(m.index + next.length, m.index + next.length); if (src === editor) onEditInput(); return; } }
}
function commitEdit(fillSelection = false) {
  if (!editing) return true;
  const val = editor.value; closeAc();
  const bad = validateInput(sel.ar, sel.ac, val);
  if (bad) { toast(bad, 'error', 4000); editor.focus(); editor.select(); return false; }
  editing = false; editor.remove(); editor = null; editRefs = []; refInsert = null;
  const prev = rawToEdit(getCell(sel.ar, sel.ac));
  if (val !== prev || fillSelection) {
    begin();
    if (fillSelection) { const g = rng(); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) inputCell(r, c, val[0] === '=' ? shiftFormula(val, r - sel.ar, c - sel.ac) : val); }
    else inputCell(sel.ar, sel.ac, val);
    const cell = getCell(sel.ar, sel.ac);
    if (cell && typeof cell.v === 'string' && cell.v.includes('\n')) setStyle(sel.ar, sel.ac, { wr: true });
    commit();
    autoRowHeight(sel.ar);
  }
  sink.focus({ preventScroll: true }); draw(); syncBars();
  return true;
}
function cancelEdit() { if (!editing) return; closeAc(); editing = false; editor.remove(); editor = null; editRefs = []; sink.focus({ preventScroll: true }); draw(); syncBars(); }
function autoRowHeight(r) {
  const s = sh(); let need = 0;
  for (const [k, cell] of s.cells) { const [rr, c] = RC(k); if (rr !== r || !cell.s?.wr || cell.v == null) continue; ctx.font = fontOf(cell.s); const lines = wrapText(formatValue(engine.value(book.active, rr, c), cell.s), colW(c) - 8); need = Math.max(need, lines.length * ((cell.s.fs || 11) * 96 / 72) * 1.2 + 6); }
  if (need > rowH(r)) { s.rowH[r] = Math.ceil(need); rebuildGeom(); draw(); }
}

/* formula bar editing */
fInput.addEventListener('focus', () => { editSheet = book.active; });
fInput.addEventListener('input', () => { refInsert = null; if (editing) { editor.value = fInput.value; positionEditor(); } updateEditRefs(fInput.value); autocomplete(fInput); });
fInput.addEventListener('keydown', (e) => {
  if (acBox && handleAcKey(e)) return;
  if (e.key === 'Enter') { e.preventDefault(); const bad = validateInput(sel.ar, sel.ac, fInput.value); if (bad) { toast(bad, 'error', 4000); return; } if (editing) { editor.value = fInput.value; commitEdit(); } else { if (!guard()) return; begin(); inputCell(sel.ar, sel.ac, fInput.value); commit(); sink.focus(); } move(1, 0); }
  if (e.key === 'Escape') { e.preventDefault(); if (editing) cancelEdit(); fInput.value = rawToEdit(getCell(sel.ar, sel.ac)); editRefs = []; sink.focus(); draw(); }
  if (e.key === 'F4') { e.preventDefault(); toggleAbsolute(fInput); }
});
fInput.addEventListener('blur', () => setTimeout(() => { if (!editing && document.activeElement !== fInput && fInput.value !== rawToEdit(getCell(sel.ar, sel.ac)) && !acBox && !pointerRef) { begin(); inputCell(sel.ar, sel.ac, fInput.value); commit(); } if (!editing) { editRefs = []; draw(); } }, 150));
function editorProxy() { /* formula bar editing without cell editor */ }
nameBox.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter') return; e.preventDefault();
  let v = nameBox.value.trim(); let target = book.active;
  const m = /^(.+)!(.+)$/.exec(v); if (m) { const i = book.sheets.findIndex((s) => s.name.toLowerCase() === m[1].replace(/'/g, '').toLowerCase()); if (i >= 0) target = i; v = m[2]; }
  const nm = book.names?.[v.toUpperCase()]; if (nm) { const i = book.sheets.findIndex((x) => x.name === nm.sheet); if (i >= 0) target = i; v = nm.ref; }
  const g = parseRange(v); if (!g) return toast('Неверная ссылка или имя: ' + nameBox.value, 'error');
  if (target !== book.active) switchSheet(target);
  sel = { ar: g.r1, ac: g.c1, fr: g.r2, fc: g.c2 }; ensureVisible(g.r1, g.c1); draw(); syncBars(); sink.focus();
});

/* autocomplete of function names */
let acBox = null, acItems = [], acIdx = 0, acSrc = null;
const FN_DESC = {}; for (const list of Object.values(FN_HELP)) for (const [n, sig, d] of list) FN_DESC[n] = [sig, d];
function autocomplete(src) {
  closeAc();
  const v = src.value; if (v[0] !== '=') return;
  const before = v.slice(0, src.selectionStart); const m = /([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)$/.exec(before);
  if (!m || /[\d$]/.test(m[1][0])) return;
  const pre = m[1].toUpperCase(); if (/^[A-Z]{1,3}\d/.test(pre)) return;
  acItems = ALL_FN_NAMES.filter((n) => n.startsWith(pre)).slice(0, 8); if (!acItems.length) return;
  acIdx = 0; acSrc = { src, start: src.selectionStart - m[1].length, end: src.selectionStart };
  acBox = h('div', { class: 'ac-list' });
  renderAc();
  document.body.append(acBox);
  const r = src.getBoundingClientRect(); acBox.style.left = r.left + 'px'; acBox.style.top = r.bottom + 2 + 'px';
}
function renderAc() { acBox.innerHTML = ''; acItems.forEach((n, i) => { const en = ALIAS[n] || n; const d = FN_DESC[en]; const it = h('div', { class: `ac-item ${i === acIdx ? 'on' : ''}`, html: `<b>${esc(n)}</b><span>${esc(d ? d[1] : '')}</span>` }); it.onmousedown = (e) => { e.preventDefault(); acIdx = i; acceptAc(); }; acBox.append(it); }); }
function handleAcKey(e) {
  if (e.key === 'ArrowDown') { e.preventDefault(); acIdx = (acIdx + 1) % acItems.length; renderAc(); return true; }
  if (e.key === 'ArrowUp') { e.preventDefault(); acIdx = (acIdx - 1 + acItems.length) % acItems.length; renderAc(); return true; }
  if (e.key === 'Tab' || (e.key === 'Enter' && acItems.length)) { e.preventDefault(); acceptAc(); return true; }
  if (e.key === 'Escape') { e.preventDefault(); closeAc(); return true; }
  return false;
}
function acceptAc() { const { src, start, end } = acSrc; const name = acItems[acIdx]; src.value = src.value.slice(0, start) + name + '(' + src.value.slice(end); src.setSelectionRange(start + name.length + 1, start + name.length + 1); closeAc(); if (src === editor) onEditInput(); else { fInput.dispatchEvent(new Event('input')); } }
function closeAc() { acBox?.remove(); acBox = null; }

/* ======================= mouse ======================= */
let dragSel = false, pointerRef = false, fillDrag = null, resizing = null, touchStart = null;
scroller.addEventListener('pointerdown', (e) => {
  if (e.button === 2) return;
  const p = hitTest(e.clientX, e.clientY);
  closeMenus(); hideTip();
  // header resize
  if (p.inColHdr && !p.inRowHdr) { const edge = nearColEdge(p.sx); if (edge != null) { resizing = { axis: 'col', i: edge, start: e.clientX, size: colW(edge) }; scroller.setPointerCapture(e.pointerId); e.preventDefault(); return; } }
  if (p.inRowHdr && !p.inColHdr) { const edge = nearRowEdge(p.sy); if (edge != null) { resizing = { axis: 'row', i: edge, start: e.clientY, size: rowH(edge) }; scroller.setPointerCapture(e.pointerId); e.preventDefault(); return; } }
  if (e.pointerType === 'touch') { touchStart = { x: e.clientX, y: e.clientY, p, t: Date.now() }; return; }
  e.preventDefault();
  // formula reference picking
  if ((editing || document.activeElement === fInput) && canInsertRef() && !p.inColHdr && !p.inRowHdr) {
    pointerRef = { r: p.r, c: p.c, si: book.active };
    insertRefText(refText(p.r, p.c, p.r, p.c));
    scroller.setPointerCapture(e.pointerId); return;
  }
  if (editing && !commitEdit()) return;
  if (p.inColHdr && p.inRowHdr) { selectAll(); return; }
  if (!p.inColHdr && !p.inRowHdr && onCellClick(p, e)) return;
  // filter button
  const f = sh().filter;
  if (f && p.r === f.r && p.c >= f.c1 && p.c <= f.c2 && !p.inColHdr && !p.inRowHdr) { const r = cellScreenRect(p.r, p.c); if (p.lx > r.x + r.w - 20) { filterPopup(p.c, e.clientX, e.clientY); return; } }
  // fill handle
  const g = expandMerged(rng()); const hr = cellScreenRect(g.r1, g.c1);
  const hx = hr.x + colX[g.c2 + 1] - colX[g.c1], hy = hr.y + rowY[g.r2 + 1] - rowY[g.r1];
  if (!p.inColHdr && !p.inRowHdr && Math.abs(p.lx - hx) < 6 && Math.abs(p.ly - hy) < 6) {
    fillDrag = { src: g, target: g }; scroller.setPointerCapture(e.pointerId); return;
  }
  if (p.inColHdr) { if (e.shiftKey) { sel.fc = p.c; sel.fr = MAX_ROWS - 1; } else sel = { ar: 0, ac: p.c, fr: MAX_ROWS - 1, fc: p.c }; dragSel = 'col'; }
  else if (p.inRowHdr) { if (e.shiftKey) { sel.fr = p.r; sel.fc = MAX_COLS - 1; } else sel = { ar: p.r, ac: 0, fr: p.r, fc: MAX_COLS - 1 }; dragSel = 'row'; }
  else { select(p.r, p.c, e.shiftKey); dragSel = 'cell'; }
  scroller.setPointerCapture(e.pointerId); sink.focus({ preventScroll: true });
  draw(); syncBars();
});
scroller.addEventListener('pointermove', (e) => {
  const p = hitTest(e.clientX, e.clientY);
  if (resizing) {
    const d = ((resizing.axis === 'col' ? e.clientX : e.clientY) - resizing.start) / Z; const v = Math.max(resizing.axis === 'col' ? 8 : 8, Math.round(resizing.size + d));
    if (resizing.axis === 'col') sh().colW[resizing.i] = v; else sh().rowH[resizing.i] = v;
    rebuildGeom(); draw(); return;
  }
  if (pointerRef) { const a = pointerRef; insertRefText(refText(a.r, a.c, p.r, p.c)); autoScroll(e); return; }
  if (fillDrag) {
    const g = fillDrag.src; let t = { ...g };
    const dr = p.r > g.r2 ? p.r - g.r2 : p.r < g.r1 ? p.r - g.r1 : 0; const dc = p.c > g.c2 ? p.c - g.c2 : p.c < g.c1 ? p.c - g.c1 : 0;
    if (Math.abs(dr) >= Math.abs(dc)) { if (dr > 0) t.r2 = p.r; else if (dr < 0) t.r1 = p.r; } else { if (dc > 0) t.c2 = p.c; else if (dc < 0) t.c1 = p.c; }
    fillDrag.target = t; draw(); autoScroll(e); return;
  }
  if (dragSel) {
    if (dragSel === 'col') { sel.fc = p.c; } else if (dragSel === 'row') { sel.fr = p.r; } else { sel.fr = p.r; sel.fc = p.c; }
    autoScroll(e); draw(); syncBars(); return;
  }
  // cursor
  if (p.inColHdr && !p.inRowHdr && nearColEdge(p.sx) != null) scroller.style.cursor = 'col-resize';
  else if (p.inRowHdr && !p.inColHdr && nearRowEdge(p.sy) != null) scroller.style.cursor = 'row-resize';
  else if (!p.inColHdr && !p.inRowHdr) {
    const g = expandMerged(rng()); const r = cellScreenRect(g.r1, g.c1); const ex = r.x + colX[g.c2 + 1] - colX[g.c1], ey = r.y + rowY[g.r2 + 1] - rowY[g.r1];
    scroller.style.cursor = Math.abs(p.lx - ex) < 6 && Math.abs(p.ly - ey) < 6 ? 'crosshair' : 'cell';
    const cell = getCell(p.r, p.c); if (cell?.n) showTip(cell.n, e.clientX, e.clientY); else hideTip();
  } else scroller.style.cursor = 'default';
});
scroller.addEventListener('pointerup', (e) => {
  if (touchStart) {
    const ts = touchStart; touchStart = null;
    if (Math.hypot(e.clientX - ts.x, e.clientY - ts.y) < 10) {
      const p = ts.p; if (editing) commitEdit();
      if (p.inColHdr && p.inRowHdr) selectAll();
      else if (p.inColHdr) sel = { ar: 0, ac: p.c, fr: MAX_ROWS - 1, fc: p.c };
      else if (p.inRowHdr) sel = { ar: p.r, ac: 0, fr: p.r, fc: MAX_COLS - 1 };
      else { const same = sel.ar === p.r && sel.ac === p.c && Date.now() - lastTap < 400; select(p.r, p.c); if (same) startEdit(null, 'edit'); lastTap = Date.now(); if (Date.now() - ts.t > 550) contextMenu(e); }
      draw(); syncBars();
    }
    return;
  }
  if (resizing) { const r = resizing; resizing = null; begin(true); tx.struct = snapSheetWith(r); commit(); rebuildGeom(); return; }
  if (pointerRef) { pointerRef = false; (editing ? editor : fInput).focus(); return; }
  if (fillDrag) { const f = fillDrag; fillDrag = null; doFill(f.src, f.target); return; }
  dragSel = false;
});
let lastTap = 0;
function snapSheetWith(r) { const s = sh(); const cur = r.axis === 'col' ? s.colW[r.i] : s.rowH[r.i]; if (r.axis === 'col') s.colW[r.i] = r.size; else s.rowH[r.i] = r.size; const snap = snapSheet(s); if (r.axis === 'col') s.colW[r.i] = cur; else s.rowH[r.i] = cur; return snap; }
scroller.addEventListener('dblclick', (e) => {
  const p = hitTest(e.clientX, e.clientY);
  if (p.inColHdr && !p.inRowHdr) { const edge = nearColEdge(p.sx); if (edge != null) { autoFitCol(edge); return; } }
  if (p.inColHdr || p.inRowHdr) return;
  const g = expandMerged(rng()); const r = cellScreenRect(g.r1, g.c1); const ex = r.x + colX[g.c2 + 1] - colX[g.c1], ey = r.y + rowY[g.r2 + 1] - rowY[g.r1];
  if (Math.abs(p.lx - ex) < 6 && Math.abs(p.ly - ey) < 6) { autoFillDown(); return; }
  startEdit(null, 'edit');
});
scroller.addEventListener('contextmenu', (e) => { e.preventDefault(); contextMenu(e); });
scroller.addEventListener('scroll', () => { draw(); hideTip(); });
host.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); setZoom(Z * (e.deltaY < 0 ? 1.1 : 0.9)); } }, { passive: false });
function nearColEdge(sx) { const c = colAt(sx); if (Math.abs(colX[c + 1] - sx) < 4 / Z) return c; if (c > 0 && Math.abs(colX[c] - sx) < 4 / Z) return c - 1; return null; }
function nearRowEdge(sy) { const r = rowAt(sy); if (Math.abs(rowY[r + 1] - sy) < 3 / Z) return r; if (r > 0 && Math.abs(rowY[r] - sy) < 3 / Z) return r - 1; return null; }
function refText(r1, c1, r2, c2) { const g = { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) }; const pre = book.active !== editSheet ? (/^[\wА-Яа-яЁё.]+$/.test(sh().name) ? sh().name : `'${sh().name}'`) + '!' : ''; return pre + rangeName(g); }
let autoScrollT = null;
function autoScroll(e) {
  const b = host.getBoundingClientRect(); const m = 24;
  const dx = e.clientX < b.left + HW * Z ? -20 : e.clientX > b.right - m ? 20 : 0; const dy = e.clientY < b.top + HH * Z ? -20 : e.clientY > b.bottom - m ? 20 : 0;
  if (dx || dy) { scroller.scrollLeft += dx; scroller.scrollTop += dy; }
}
function autoFitCol(c) {
  let w = 30; const s = sh();
  for (const [k, cell] of s.cells) { const [r, cc] = RC(k); if (cc !== c || cell.v == null) continue; ctx.font = fontOf(cell.s); w = Math.max(w, ctx.measureText(displayOf(r, c)).width + 12); }
  begin(true); s.colW[c] = Math.min(600, Math.ceil(w)); commit(); rebuildGeom();
}
let tipEl = null;
function showTip(text, x, y) { if (!tipEl) { tipEl = h('div', { class: 'cell-tip' }); document.body.append(tipEl); } tipEl.textContent = text; tipEl.style.left = x + 14 + 'px'; tipEl.style.top = y + 10 + 'px'; }
function hideTip() { tipEl?.remove(); tipEl = null; }

/* ======================= fill ======================= */
const LISTS = [['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'], ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'], ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'], ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'], ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'], ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'], ['q1', 'q2', 'q3', 'q4']];
function seriesFor(cells, n) {
  // cells: source cells in order; returns n new values
  const vals = cells.map((c) => c?.v);
  const out = [];
  if (vals.every((v) => typeof v === 'number') && vals.length >= 2) {
    const step = (vals[vals.length - 1] - vals[0]) / (vals.length - 1);
    for (let i = 1; i <= n; i++) out.push({ v: +(vals[vals.length - 1] + step * i).toPrecision(15) });
    return out;
  }
  for (let i = 0; i < n; i++) {
    const src = cells[i % cells.length]; const cycle = Math.floor(i / cells.length) + 1;
    let v = src?.v;
    if (typeof v === 'number' && vals.length === 1 && /date/.test(src.s?.fmt || '')) v = v + (i + 1);
    else if (typeof v === 'string' && v[0] !== '=') {
      const low = v.toLowerCase(); const list = LISTS.find((l) => l.includes(low));
      if (list) { const idx = (list.indexOf(low) + (i + 1)) % list.length; let nv = list[idx]; if (v[0] === v[0].toUpperCase()) nv = nv[0].toUpperCase() + nv.slice(1); if (v === v.toUpperCase()) nv = nv.toUpperCase(); v = nv; }
      else { const m = /^(.*?)(\d+)$/.exec(v); if (m && vals.length === 1) v = m[1] + (+m[2] + i + 1); }
    }
    out.push({ v, s: src?.s, cycle });
  }
  return out;
}
function doFill(src, t) {
  if (t.r1 === src.r1 && t.r2 === src.r2 && t.c1 === src.c1 && t.c2 === src.c2) return;
  begin();
  const s = sh();
  if (t.r2 > src.r2 || t.r1 < src.r1) {
    const down = t.r2 > src.r2; const n = down ? t.r2 - src.r2 : src.r1 - t.r1;
    for (let c = src.c1; c <= src.c2; c++) {
      let cells = []; for (let r = src.r1; r <= src.r2; r++) cells.push(s.cells.get(K(r, c)));
      if (!down) cells = cells.reverse();
      const vals = seriesFor(cells, n);
      vals.forEach((x, i) => {
        const r = down ? src.r2 + 1 + i : src.r1 - 1 - i; const srcR = down ? src.r1 + (i % cells.length) : src.r2 - (i % cells.length);
        const sc = s.cells.get(K(srcR, c)); let v = x.v;
        if (typeof sc?.v === 'string' && sc.v[0] === '=') v = shiftFormula(sc.v, r - srcR, 0);
        setCell(r, c, { v, s: sc?.s }); if (!sc?.s) { const cc = s.cells.get(K(r, c)); if (cc) delete cc.s; }
      });
    }
  } else {
    const right = t.c2 > src.c2; const n = right ? t.c2 - src.c2 : src.c1 - t.c1;
    for (let r = src.r1; r <= src.r2; r++) {
      let cells = []; for (let c = src.c1; c <= src.c2; c++) cells.push(s.cells.get(K(r, c)));
      if (!right) cells = cells.reverse();
      seriesFor(cells, n).forEach((x, i) => {
        const c = right ? src.c2 + 1 + i : src.c1 - 1 - i; const srcC = right ? src.c1 + (i % cells.length) : src.c2 - (i % cells.length);
        const sc = s.cells.get(K(r, srcC)); let v = x.v;
        if (typeof sc?.v === 'string' && sc.v[0] === '=') v = shiftFormula(sc.v, 0, c - srcC);
        setCell(r, c, { v, s: sc?.s });
      });
    }
  }
  commit();
  sel = { ar: Math.min(t.r1, src.r1), ac: Math.min(t.c1, src.c1), fr: Math.max(t.r2, src.r2), fc: Math.max(t.c2, src.c2) }; draw(); syncBars();
}
function autoFillDown() {
  const g = rng(); const c = g.c1 > 0 ? g.c1 - 1 : g.c2 + 1; let r = g.r2;
  while (r + 1 < MAX_ROWS && getCell(r + 1, c)?.v != null) r++;
  if (r > g.r2) doFill(g, { ...g, r2: r });
}
function fillDir(dir) {
  const g = rng(); begin();
  if (dir === 'down') for (let c = g.c1; c <= g.c2; c++) { const src = getCell(g.r1, c); for (let r = g.r1 + 1; r <= g.r2; r++) setCell(r, c, src ? { v: typeof src.v === 'string' && src.v[0] === '=' ? shiftFormula(src.v, r - g.r1, 0) : src.v, s: src.s } : null); }
  else for (let r = g.r1; r <= g.r2; r++) { const src = getCell(r, g.c1); for (let c = g.c1 + 1; c <= g.c2; c++) setCell(r, c, src ? { v: typeof src.v === 'string' && src.v[0] === '=' ? shiftFormula(src.v, 0, c - g.c1) : src.v, s: src.s } : null); }
  commit();
}

/* ======================= clipboard ======================= */
let clip = null; let marchOffset = 0;
setInterval(() => { if (clip && clip.si === book.active) { marchOffset = (marchOffset + 1) % 14; draw(); } }, 80);
function rangeTSV(g) { const rows = []; for (let r = g.r1; r <= g.r2; r++) { if (hiddenRows.has(r)) continue; const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(displayOf(r, c).replace(/\t|\n/g, ' ')); rows.push(row.join('\t')); } return rows.join('\n'); }
function rangeHTML(g) { let o = '<table>'; for (let r = g.r1; r <= g.r2; r++) { o += '<tr>'; for (let c = g.c1; c <= g.c2; c++) { const s = getCell(r, c)?.s || {}; o += `<td style="${s.b ? 'font-weight:bold;' : ''}${s.i ? 'font-style:italic;' : ''}${s.bg ? 'background:' + s.bg + ';' : ''}${s.fc ? 'color:' + s.fc + ';' : ''}">${esc(displayOf(r, c))}</td>`; } o += '</tr>'; } return o + '</table>'; }
function doCopy(e, cut = false) {
  const g = selRangeForOps();
  const cells = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(structuredClone(getCell(r, c) || null)); cells.push(row); }
  const tsv = rangeTSV(g);
  clip = { si: book.active, ...g, cells, cut, tsv };
  if (e?.clipboardData) { e.clipboardData.setData('text/plain', tsv); e.clipboardData.setData('text/html', rangeHTML(g)); e.preventDefault(); }
  else navigator.clipboard?.writeText(tsv).catch(() => {});
  draw();
}
function doPaste(text, html) {
  const r0 = sel.ar, c0 = sel.ac;
  begin();
  if (clip && (text == null || text.replace(/\r/g, '') === clip.tsv)) {
    const dr = r0 - clip.r1, dc = c0 - clip.c1; const src = book.sheets[clip.si];
    if (clip.cut) for (let r = clip.r1; r <= clip.r2; r++) for (let c = clip.c1; c <= clip.c2; c++) { if (clip.si === book.active) setCell(r, c, null); else src.cells.delete(K(r, c)); }
    clip.cells.forEach((row, i) => row.forEach((cell, j) => {
      const r = r0 + i, c = c0 + j;
      if (!cell) { setCell(r, c, null); const k = K(r, c); const cur = sh().cells.get(k); if (cur) sh().cells.delete(k); return; }
      const n = structuredClone(cell); if (typeof n.v === 'string' && n.v[0] === '=' && !clip.cut) n.v = shiftFormula(n.v, dr, dc);
      recordCell(K(r, c)); sh().cells.set(K(r, c), n);
    }));
    sel = { ar: r0, ac: c0, fr: r0 + clip.r2 - clip.r1, fc: c0 + clip.c2 - clip.c1 };
    if (clip.cut) clip = null;
  } else if (text != null) {
    let rows;
    if (html && /<table/i.test(html)) { const d = new DOMParser().parseFromString(html, 'text/html'); rows = [...d.querySelectorAll('tr')].map((tr) => [...tr.cells].map((td) => td.innerText.trim())); }
    else rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map((l) => l.split('\t'));
    rows.forEach((row, i) => row.forEach((v, j) => inputCell(r0 + i, c0 + j, v)));
    sel = { ar: r0, ac: c0, fr: r0 + rows.length - 1, fc: c0 + Math.max(...rows.map((r) => r.length)) - 1 };
  }
  commit();
}
document.addEventListener('copy', (e) => { if (isGridFocus()) doCopy(e); });
document.addEventListener('cut', (e) => { if (isGridFocus()) doCopy(e, true); });
document.addEventListener('paste', (e) => { if (!isGridFocus()) return; e.preventDefault(); doPaste(e.clipboardData.getData('text/plain'), e.clipboardData.getData('text/html')); });
const isGridFocus = () => document.activeElement === sink && !editing;

/* ======================= keyboard ======================= */
sink.addEventListener('keydown', (e) => {
  if (editing) return;
  const mod = e.ctrlKey || e.metaKey; const k = e.key;
  const nav = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
  if (nav) { e.preventDefault(); move(nav[0], nav[1], e.shiftKey, mod); return; }
  if (k === 'Enter') { e.preventDefault(); moveInSel(e.shiftKey ? -1 : 1, 0); return; }
  if (k === 'Tab') { e.preventDefault(); moveInSel(0, e.shiftKey ? -1 : 1); return; }
  if (k === 'Home') { e.preventDefault(); select(mod ? 0 : sel.ar, 0, e.shiftKey); return; }
  if (k === 'End' && mod) { e.preventDefault(); const u = engine.usedBounds(book.active); select(u.mr, u.mc, e.shiftKey); return; }
  if (k === 'PageDown' || k === 'PageUp') { e.preventDefault(); const rows = Math.floor(host.clientHeight / Z / DEF_H) - 2; move((k === 'PageDown' ? 1 : -1) * rows, 0, e.shiftKey); return; }
  if (k === 'F2') { e.preventDefault(); startEdit(null, 'edit'); return; }
  if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (k === 'Backspace' && rng().r1 === rng().r2 && rng().c1 === rng().c2) { startEdit('', 'enter'); return; } clearRange('content'); return; }
  if (k === 'Escape') { if (clip) { clip = null; draw(); } return; }
  if (mod) {
    const kk = k.toLowerCase();
    if (kk === 'a') { e.preventDefault(); const g = currentRegion(sel.ar, sel.ac); const cur = rng(); if (cur.r1 === g.r1 && cur.r2 === g.r2 && cur.c1 === g.c1 && cur.c2 === g.c2 || (g.r1 === g.r2 && g.c1 === g.c2)) selectAll(); else { sel = { ar: g.r1, ac: g.c1, fr: g.r2, fc: g.c2 }; draw(); syncBars(); } return; }
    if (kk === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (kk === 'y') { e.preventDefault(); redo(); return; }
    if (kk === 'b') { e.preventDefault(); toggleStyle('b'); return; }
    if (kk === 'i') { e.preventDefault(); toggleStyle('i'); return; }
    if (kk === 'u') { e.preventDefault(); toggleStyle('u'); return; }
    if (kk === 'd') { e.preventDefault(); fillDir('down'); return; }
    if (kk === 'r') { e.preventDefault(); fillDir('right'); return; }
    if (kk === '1') { e.preventDefault(); formatDialog(); return; }
    if (kk === 'k') { e.preventDefault(); insertLinkDialog(); return; }
    if (k === ';') { e.preventDefault(); begin(); const d = new Date(); setCell(sel.ar, sel.ac, { v: dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) }); setStyle(sel.ar, sel.ac, { fmt: 'date' }); commit(); return; }
    if (k === ' ') { e.preventDefault(); sel = { ar: 0, ac: sel.ac, fr: MAX_ROWS - 1, fc: sel.fc }; draw(); syncBars(); return; }
    if (kk === 'v' && !navigator.clipboard) return;
    return;
  }
  if (e.shiftKey && k === ' ') { e.preventDefault(); sel = { ar: sel.ar, ac: 0, fr: sel.fr, fc: MAX_COLS - 1 }; draw(); syncBars(); return; }
  if (e.altKey && k === '=') { e.preventDefault(); autoSum('SUM'); return; }
  if (e.shiftKey && k === 'F3') { e.preventDefault(); insertFunctionDialog(); return; }
  // printable characters arrive through the sink's input event
});
function moveInSel(dr, dc) {
  const g = rng();
  if (g.r1 === g.r2 && g.c1 === g.c2) { move(dr, dc); return; }
  let r = sel.ar + dr, c = sel.ac + dc;
  if (r > g.r2) { r = g.r1; c++; } if (r < g.r1) { r = g.r2; c--; } if (c > g.c2) { c = g.c1; r += dc ? 1 : 0; if (r > g.r2) r = g.r1; } if (c < g.c1) { c = g.c2; r -= dc ? 1 : 0; if (r < g.r1) r = g.r2; }
  sel.ar = clamp(r, g.r1, g.r2); sel.ac = clamp(c, g.c1, g.c2); draw(); syncBars();
}

/* ======================= formatting commands ======================= */
function forSel(fn) { const g = selRangeForOps(); begin(); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) fn(r, c); commit(); }
function toggleStyle(key) { const on = !getCell(sel.ar, sel.ac)?.s?.[key]; forSel((r, c) => setStyle(r, c, { [key]: on })); }
function style(patch) { forSel((r, c) => setStyle(r, c, patch)); sink.focus({ preventScroll: true }); }
function clearRange(what) {
  const g = selRangeForOps(); begin();
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
    const cell = getCell(r, c); if (!cell) continue;
    if (what === 'all') setCell(r, c, { v: null, s: undefined, n: undefined });
    else if (what === 'format') { recordCell(K(r, c)); const n = { ...cell }; delete n.s; if (n.v == null && !n.n) sh().cells.delete(K(r, c)); else sh().cells.set(K(r, c), n); }
    else setCell(r, c, null);
  }
  if (what === 'all') { sh().cf = sh().cf.filter((x) => !(x.range.r1 >= g.r1 && x.range.r2 <= g.r2 && x.range.c1 >= g.c1 && x.range.c2 <= g.c2)); }
  commit();
}
function changeDecimals(d) {
  const cell = getCell(sel.ar, sel.ac); const v = cell ? engine.value(book.active, sel.ar, sel.ac) : null;
  let cur = cell?.s?.dec;
  if (cur == null) { const f = cell?.s?.fmt; cur = f === 'number' || f === 'currency' ? 2 : f === 'percent' ? 0 : typeof v === 'number' ? Math.min(10, (fmtGeneral(v).split('.')[1] || '').length) : 0; }
  style({ dec: clamp(cur + d, 0, 12) });
}
function setBorders(kind) {
  const g = selRangeForOps(); begin();
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
    const cur = { ...(getCell(r, c)?.s?.bd || {}) }; let bd = cur;
    if (kind === 'none') bd = null;
    else if (kind === 'all') bd = { t: 1, r: 1, b: 1, l: 1 };
    else if (kind === 'outer' || kind === 'thick') { const w = kind === 'thick' ? 2 : 1; if (r === g.r1) bd.t = w; if (r === g.r2) bd.b = w; if (c === g.c1) bd.l = w; if (c === g.c2) bd.r = w; }
    else if (kind === 'bottom') { if (r === g.r2) bd.b = 1; } else if (kind === 'top') { if (r === g.r1) bd.t = 1; }
    else if (kind === 'left') { if (c === g.c1) bd.l = 1; } else if (kind === 'right') { if (c === g.c2) bd.r = 1; }
    else if (kind === 'double') { if (r === g.r2) bd.b = 3; }
    setStyle(r, c, { bd: bd && Object.keys(bd).length ? bd : null });
  }
  commit();
}
function autoSum(fn = 'SUM') {
  const g = rng();
  if (g.r1 === g.r2 && g.c1 === g.c2) {
    let r = sel.ar - 1; while (r >= 0 && typeof engine.value(book.active, r, sel.ac) === 'number') r--;
    let range;
    if (r < sel.ar - 1) range = { r1: r + 1, c1: sel.ac, r2: sel.ar - 1, c2: sel.ac };
    else { let c = sel.ac - 1; while (c >= 0 && typeof engine.value(book.active, sel.ar, c) === 'number') c--; if (c < sel.ac - 1) range = { r1: sel.ar, c1: c + 1, r2: sel.ar, c2: sel.ac - 1 }; }
    editSheet = book.active;
    startEdit(`=${fn}(${range ? rangeName(range) : ''})`, 'edit');
    if (!range) editor.setSelectionRange(editor.value.length - 1, editor.value.length - 1);
    return;
  }
  begin();
  const target = g.r2 + 1;
  for (let c = g.c1; c <= g.c2; c++) setCell(target, c, { v: `=${fn}(${rangeName({ r1: g.r1, c1: c, r2: g.r2, c2: c })})` });
  commit();
}

/* ======================= charts ======================= */
const chartLayer = h('div', { style: { position: 'absolute', inset: '0' } }); overlay.append(chartLayer);
let selChart = null;
function chartData(ch) {
  const g = ch.range; const si = ch.si ?? book.active; const grid = [];
  for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) { const row = []; for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 30); c++) row.push(engine.value(si, r, c)); grid.push(row); }
  return chartDataFromGrid(grid);
}
function renderCharts() {
  chartLayer.innerHTML = '';
  const dark = document.documentElement.dataset.theme.includes('dark');
  for (const ch of sh().charts) {
    const box = h('div', { class: `chart-box ${selChart === ch.id ? 'sel' : ''}`, 'data-id': ch.id });
    box.innerHTML = renderChart({ type: ch.type, title: ch.title, legend: ch.legend, ...chartData(ch) }, ch.w, ch.h, { text: dark ? '#c8ccd4' : '#444', grid: dark ? 'rgba(255,255,255,.12)' : 'rgba(0,0,0,.1)' });
    const rs = h('div', { class: 'rs' }); box.append(rs);
    box.addEventListener('pointerdown', (e) => {
      e.stopPropagation(); e.preventDefault(); if (editing) commitEdit(); selectChart(ch.id);
      const resize = e.target === rs; const sx = e.clientX, sy = e.clientY; const o = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
      const before = snapSheet(sh());
      const mv = (ev) => { const dx = (ev.clientX - sx) / Z, dy = (ev.clientY - sy) / Z; if (resize) { ch.w = Math.max(160, o.w + dx); ch.h = Math.max(110, o.h + dy); renderCharts(); } else { ch.x = Math.max(0, o.x + dx); ch.y = Math.max(0, o.y + dy); positionCharts(); } };
      const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); if (o.x !== ch.x || o.y !== ch.y || o.w !== ch.w || o.h !== ch.h) { undoStack.push({ si: book.active, struct: before, after: snapSheet(sh()), sel: { ...sel } }); redoStack.length = 0; shell.markDirty(); } };
      document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
    });
    box.addEventListener('dblclick', () => chartDialog(ch));
    box.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); selectChart(ch.id); menu({ x: e.clientX, y: e.clientY }, [{ label: 'Изменить диаграмму…', icon: 'edit', action: () => chartDialog(ch) }, { label: 'Удалить', icon: 'trash', danger: true, action: deleteChart }]); });
    chartLayer.append(box);
  }
  positionCharts();
}
function positionCharts() {
  for (const box of chartLayer.children) {
    const ch = sh().charts.find((c) => c.id === box.dataset.id); if (!ch) continue;
    box.style.left = (HW + ch.x - scrollX()) * Z + 'px'; box.style.top = (HH + ch.y - scrollY()) * Z + 'px';
    box.style.width = ch.w * Z + 'px'; box.style.height = ch.h * Z + 'px';
  }
}
function selectChart(id) { if (selChart === id) return; selChart = id; for (const b of chartLayer.children) b.classList.toggle('sel', b.dataset.id === id); shell?.ribbon.showTab('chart', !!id, true); }
function deleteChart() { if (!selChart) return; begin(true); sh().charts = sh().charts.filter((c) => c.id !== selChart); selChart = null; commit(); shell.ribbon.showTab('chart', false); }
function insertChart(type) {
  let g = rng(); if (g.r1 === g.r2 && g.c1 === g.c2) g = currentRegion(sel.ar, sel.ac);
  if (g.r1 === g.r2 && g.c1 === g.c2 && getCell(g.r1, g.c1)?.v == null) return toast('Выделите данные для диаграммы', 'error');
  begin(true);
  const ch = { id: uid('ch'), type, title: '', range: g, x: colX[g.c2 + 1] + 20, y: rowY[g.r1], w: 440, h: 270, legend: true };
  const hdr = displayOf(g.r1, g.c1 + 1); if (hdr && isNaN(+hdr.replace(',', '.'))) ch.title = g.c2 - g.c1 === 1 ? hdr : '';
  sh().charts.push(ch); commit(); selectChart(ch.id); renderCharts();
}
async function chartDialog(ch) {
  const r = await formDlg('Диаграмма', [{ key: 'type', label: 'Тип', type: 'select', options: CHART_TYPES.map((t) => ({ value: t.id, label: t.name })), value: ch.type }, { key: 'title', label: 'Заголовок', value: ch.title }, { key: 'range', label: 'Диапазон данных', value: rangeName(ch.range) }, { key: 'legend', label: 'Показывать легенду', type: 'checkbox', value: ch.legend !== false }], 'Применить');
  if (!r) return;
  const g = parseRange(r.range); begin(true); Object.assign(ch, { type: r.type, title: r.title, legend: r.legend, range: g || ch.range }); commit();
}
function setChartType(t) { const ch = sh().charts.find((c) => c.id === selChart); if (!ch) return; begin(true); ch.type = t; commit(); }

/* ======================= conditional formatting ======================= */
async function addCF(type) {
  const g = selRangeForOps(); const id = uid('cf');
  const styles = { red: { bg: '#ffc7ce', fc: '#9c0006' }, green: { bg: '#c6efce', fc: '#006100' }, yellow: { bg: '#ffeb9c', fc: '#9c5700' }, bold: { b: true } };
  const styleField = { key: 'style', label: 'Формат', type: 'select', options: [{ value: 'red', label: 'Светло-красная заливка и тёмно-красный текст' }, { value: 'green', label: 'Зелёная заливка и тёмно-зелёный текст' }, { value: 'yellow', label: 'Жёлтая заливка и тёмно-жёлтый текст' }] };
  let rule = null;
  if (type === 'gt' || type === 'lt' || type === 'eq' || type === 'text') {
    const r = await formDlg({ gt: 'Больше чем', lt: 'Меньше чем', eq: 'Равно', text: 'Текст содержит' }[type], [{ key: 'a', label: 'Значение', value: '' }, styleField]); if (!r) return;
    rule = { id, type, range: g, a: r.a.replace(',', '.'), style: styles[r.style] };
  } else if (type === 'between') {
    const r = await formDlg('Между', [{ key: 'a', label: 'От', value: '' }, { key: 'b', label: 'До', value: '' }, styleField]); if (!r) return;
    rule = { id, type, range: g, a: r.a.replace(',', '.'), b: r.b.replace(',', '.'), style: styles[r.style] };
  } else if (type === 'dup') rule = { id, type, range: g, style: styles.red };
  else if (type === 'scale') rule = { id, type, range: g, c1: '#f8696b', c2: '#63be7b' };
  else if (type === 'scale2') rule = { id, type: 'scale', range: g, c1: '#ffffff', c2: '#5a8ac6' };
  else if (type === 'bar') rule = { id, type, range: g, color: '#5b9bd5' };
  else if (type === 'bar2') rule = { id, type: 'bar', range: g, color: '#63be7b' };
  if (!rule) return;
  begin(true); sh().cf.push(rule); commit();
}
function clearCF(all) { const g = rng(); begin(true); sh().cf = all ? [] : sh().cf.filter((x) => x.range.r2 < g.r1 || x.range.r1 > g.r2 || x.range.c2 < g.c1 || x.range.c1 > g.c2); commit(); }

/* ======================= dialogs ======================= */
function insertFunctionDialog() {
  let cat = Object.keys(FN_HELP)[0]; let chosen = null;
  const cats = h('div', { class: 'cats' }); const fns = h('div', { class: 'fns' });
  const search = h('input', { class: 'input', placeholder: 'Поиск функции…', style: { marginBottom: '10px' } });
  const renderFns = () => {
    fns.innerHTML = ''; const q = search.value.trim().toLowerCase();
    const list = q ? Object.values(FN_HELP).flat().filter(([n, sig, d]) => (n + sig + d).toLowerCase().includes(q)) : FN_HELP[cat];
    for (const [n, sig, d] of list) { const it = h('div', { class: `fn ${chosen === n ? 'on' : ''}`, html: `<code>${esc(sig)}</code><div class="small muted">${esc(d)}</div>` }); it.onclick = () => { chosen = n; renderFns(); }; it.ondblclick = () => { chosen = n; api.close(true); }; fns.append(it); }
  };
  for (const c of Object.keys(FN_HELP)) { const b = h('button', { class: c === cat ? 'on' : '', onclick: () => { cat = c; search.value = ''; [...cats.children].forEach((x) => x.classList.toggle('on', x === b)); renderFns(); } }, c); cats.append(b); }
  search.oninput = renderFns; renderFns();
  let api;
  dialog({ title: 'Вставка функции', wide: true, body: h('div', {}, search, h('div', { class: 'fn-cats' }, cats, fns), h('p', { class: 'small muted' }, 'Можно писать функции по-русски (СУММ, ЕСЛИ, ВПР) или по-английски (SUM, IF, VLOOKUP). Аргументы разделяются «;» или «,».')), onOpen: (a) => { api = a; }, buttons: [{ label: 'Отмена', value: false }, { label: 'Вставить', value: true, primary: true }] })
    .then((ok) => { if (!ok || !chosen) return; const ru = Object.entries(ALIAS).find(([, en]) => en === chosen)?.[0]; editSheet = book.active; if (editing) { editor.value += (ru || chosen) + '('; onEditInput(); editor.focus(); } else startEdit('=' + (ru || chosen) + '(', 'edit'); });
}
async function formatDialog() {
  const st = getCell(sel.ar, sel.ac)?.s || {};
  const r = await formDlg('Формат ячеек', [
    { key: 'fmt', label: 'Числовой формат', type: 'select', value: st.fmt || 'general', options: FORMATS.map(([v, l]) => ({ value: v, label: l })) },
    { key: 'dec', label: 'Число десятичных знаков', type: 'number', value: st.dec ?? '', min: 0, max: 12 },
    { key: 'cur', label: 'Символ валюты', type: 'select', value: st.cur || '₽', options: ['₽', '$', '€', '£', '¥', '₸', '₴'] },
    { key: 'al', label: 'Выравнивание по горизонтали', type: 'select', value: st.al || '', options: [{ value: '', label: 'Обычное' }, { value: 'left', label: 'По левому краю' }, { value: 'center', label: 'По центру' }, { value: 'right', label: 'По правому краю' }] },
    { key: 'va', label: 'По вертикали', type: 'select', value: st.va || 'bottom', options: [{ value: 'top', label: 'По верхнему краю' }, { value: 'middle', label: 'По центру' }, { value: 'bottom', label: 'По нижнему краю' }] },
    { key: 'wr', label: 'Переносить текст', type: 'checkbox', value: st.wr },
    { key: 'ff', label: 'Шрифт', type: 'select', value: st.ff || DEFAULT_FONT, options: FONTS },
    { key: 'fs', label: 'Размер', type: 'number', value: st.fs || 11, min: 6, max: 96 },
  ], 'Применить');
  if (!r) return;
  style({ fmt: r.fmt === 'general' ? null : r.fmt, dec: r.dec === '' || r.dec == null || isNaN(r.dec) ? null : +r.dec, cur: r.fmt === 'currency' ? r.cur : null, al: r.al || null, va: r.va === 'bottom' ? null : r.va, wr: r.wr || null, ff: r.ff === DEFAULT_FONT ? null : r.ff, fs: +r.fs === 11 ? null : +r.fs });
}
function findDialog(replace = false) {
  const q = h('input', { class: 'input', placeholder: 'Что найти' }); const rep = h('input', { class: 'input', placeholder: 'Заменить на' });
  const mc = h('input', { type: 'checkbox' }); const whole = h('input', { type: 'checkbox' }); const all = h('input', { type: 'checkbox' });
  const info = h('div', { class: 'small muted', style: { minHeight: '18px' } });
  let lastKey = -1;
  const matches = () => {
    const out = []; const needle = mc.checked ? q.value : q.value.toLowerCase();
    const sheets = all.checked ? book.sheets.map((_, i) => i) : [book.active];
    for (const si of sheets) for (const [k, cell] of book.sheets[si].cells) { if (cell.v == null) continue; const [r, c] = RC(k); let t = typeof cell.v === 'string' && cell.v[0] === '=' ? cell.v : displayOf(r, c, book.sheets[si], si); if (!mc.checked) t = t.toLowerCase(); if (whole.checked ? t === needle : t.includes(needle)) out.push({ si, k, r, c }); }
    return out.sort((a, b) => a.si - b.si || a.r - b.r || a.c - b.c);
  };
  const findNext = () => { if (!q.value) return; const m = matches(); if (!m.length) { info.textContent = 'Ничего не найдено'; return; } const cur = book.active * 1e9 + K(sel.ar, sel.ac); let i = m.findIndex((x) => x.si * 1e9 + x.k > (lastKey >= 0 ? lastKey : cur - 1)); if (i < 0) i = 0; const t = m[i]; lastKey = t.si * 1e9 + t.k; if (t.si !== book.active) switchSheet(t.si); select(t.r, t.c); info.textContent = `${i + 1} из ${m.length}`; };
  const replaceAll = () => { const m = matches(); if (!m.length) return; begin(true); const re = new RegExp(q.value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), mc.checked ? 'g' : 'gi'); let n = 0; for (const t of m) { if (t.si !== book.active) continue; const cell = sh().cells.get(t.k); if (typeof cell.v === 'string') { cell.v = whole.checked ? rep.value : cell.v.replace(re, rep.value); n++; } else if (whole.checked || String(cell.v) === q.value) { const p = parseInput(rep.value); cell.v = p.v; n++; } } commit(); info.textContent = `Заменено: ${n}`; };
  const body = h('div', { style: { width: 'min(440px,100%)' } }, h('div', { class: 'field' }, h('label', {}, 'Найти'), q), replace ? h('div', { class: 'field' }, h('label', {}, 'Заменить на'), rep) : '', h('div', { class: 'row wrap', style: { gap: '14px' } }, h('label', { class: 'check' }, mc, 'Учитывать регистр'), h('label', { class: 'check' }, whole, 'Ячейка целиком'), h('label', { class: 'check' }, all, 'Во всей книге')), info);
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); findNext(); } });
  dialog({ title: replace ? 'Найти и заменить' : 'Найти', body, buttons: [...(replace ? [{ label: 'Заменить все', value: () => { replaceAll(); return undefined; } }] : []), { label: 'Найти далее', primary: true, value: () => { findNext(); return undefined; } }] });
}
async function noteDialog() {
  const cell = getCell(sel.ar, sel.ac);
  const r = await formDlg('Примечание к ячейке ' + refName(sel.ar, sel.ac), [{ key: 'n', label: 'Текст', type: 'textarea', value: cell?.n || '' }], 'Сохранить');
  if (r == null) return; begin(); setCell(sel.ar, sel.ac, { n: r.n.trim() || undefined }); const c = getCell(sel.ar, sel.ac); if (c && !c.n) { delete c.n; if (c.v == null && !c.s) sh().cells.delete(K(sel.ar, sel.ac)); } commit();
}
async function sizeDialog(axis) {
  const g = rng(); const cur = axis === 'col' ? colW(g.c1) : rowH(g.r1);
  const r = await promptDlg(axis === 'col' ? 'Ширина столбца' : 'Высота строки', 'В пикселях', String(cur), { type: 'number' }); if (!r) return;
  begin(true); if (axis === 'col') for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 300); c++) sh().colW[c] = clamp(+r, 0, 1000); else for (let rr = g.r1; rr <= Math.min(g.r2, g.r1 + 5000); rr++) sh().rowH[rr] = clamp(+r, 0, 600); commit(); rebuildGeom();
}
function hideRowsCols(axis, hide) { const g = rng(); begin(true); if (axis === 'col') { for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 300); c++) { if (hide) sh().colW[c] = 0; else if (sh().colW[c] === 0) delete sh().colW[c]; } if (!hide) for (let c = Math.max(0, g.c1 - 1); c <= g.c2 + 1; c++) if (sh().colW[c] === 0) delete sh().colW[c]; } else { for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 5000); r++) { if (hide) sh().rowH[r] = 0; else if (sh().rowH[r] === 0) delete sh().rowH[r]; } if (!hide) for (let r = Math.max(0, g.r1 - 1); r <= g.r2 + 1; r++) if (sh().rowH[r] === 0) delete sh().rowH[r]; } commit(); rebuildGeom(); }
function contextMenu(e) {
  const p = hitTest(e.clientX, e.clientY);
  const g = rng(); const inside = p.r >= g.r1 && p.r <= g.r2 && p.c >= g.c1 && p.c <= g.c2;
  if (!inside && !p.inColHdr && !p.inRowHdr) select(p.r, p.c);
  const isCol = p.inColHdr, isRow = p.inRowHdr;
  menu({ x: e.clientX, y: e.clientY }, [
    { label: 'Вырезать', icon: 'cut', hint: 'Ctrl+X', action: () => { doCopy(null, true); } },
    { label: 'Копировать', icon: 'copy', hint: 'Ctrl+C', action: () => doCopy(null) },
    { label: 'Вставить', icon: 'paste', hint: 'Ctrl+V', action: async () => { try { const t = await navigator.clipboard.readText(); doPaste(t); } catch { doPaste(null); } } },
    '-',
    isCol ? { label: 'Вставить столбцы', icon: 'insert-col', action: () => insertCols(true) } : isRow ? { label: 'Вставить строки', icon: 'insert-row', action: () => insertRows(true) } : { label: 'Вставить…', icon: 'plus', sub: [{ label: 'Строку выше', action: () => insertRows(true) }, { label: 'Строку ниже', action: () => insertRows(false) }, { label: 'Столбец слева', action: () => insertCols(true) }, { label: 'Столбец справа', action: () => insertCols(false) }] },
    isCol ? { label: 'Удалить столбцы', icon: 'delete-col', action: deleteCols } : isRow ? { label: 'Удалить строки', icon: 'delete-row', action: deleteRows } : { label: 'Удалить…', icon: 'trash', sub: [{ label: 'Строки', action: deleteRows }, { label: 'Столбцы', action: deleteCols }] },
    { label: 'Очистить содержимое', icon: 'eraser', hint: 'Delete', action: () => clearRange('content') },
    '-',
    isCol ? { label: 'Ширина столбца…', icon: 'columns', action: () => sizeDialog('col') } : null, isRow ? { label: 'Высота строки…', icon: 'menu', action: () => sizeDialog('row') } : null,
    isCol ? { label: 'Скрыть', icon: 'eye-off', action: () => hideRowsCols('col', true) } : null, isCol ? { label: 'Показать', icon: 'eye', action: () => hideRowsCols('col', false) } : null,
    isRow ? { label: 'Скрыть', icon: 'eye-off', action: () => hideRowsCols('row', true) } : null, isRow ? { label: 'Показать', icon: 'eye', action: () => hideRowsCols('row', false) } : null,
    { label: 'Сортировка', icon: 'sort-asc', sub: [{ label: 'По возрастанию', action: () => quickSort(true) }, { label: 'По убыванию', action: () => quickSort(false) }, { label: 'Настраиваемая…', action: sortDialog }] },
    { label: getCell(sel.ar, sel.ac)?.n ? 'Изменить примечание' : 'Вставить примечание', icon: 'comment', action: noteDialog },
    { label: 'Формат ячеек…', icon: 'settings', hint: 'Ctrl+1', action: formatDialog },
  ]);
}
function quickSort(asc) { const g0 = rng(); const g = g0.r1 === g0.r2 ? currentRegion(sel.ar, sel.ac) : g0; const hdr = g.r2 > g.r1 && typeof engine.value(book.active, g.r1, sel.ac) === 'string' && typeof engine.value(book.active, g.r1 + 1, sel.ac) === 'number'; sortRange(g, sel.ac, asc, hdr); }

/* ======================= sheets ======================= */
function renderTabs() {
  tabsBar.innerHTML = '';
  tabsBar.append(h('button', { class: 'btn-r', 'data-tip': 'Новый лист|Shift+F11', html: icon('plus'), onclick: addSheet }));
  book.sheets.forEach((s, i) => {
    const t = h('button', { class: `sheet-tab ${i === book.active ? 'on' : ''}` }, s.name);
    if (s.color) t.append(h('span', { class: 'tc', style: { background: s.color } }));
    t.onclick = () => switchSheet(i);
    t.ondblclick = () => renameSheet(i, t);
    t.oncontextmenu = (e) => { e.preventDefault(); switchSheet(i); menu({ x: e.clientX, y: e.clientY }, [
      { label: 'Переименовать', icon: 'edit', action: () => renameSheet(i, tabsBar.children[i + 1]) },
      { label: 'Дублировать', icon: 'copy', action: () => { const d = sheetFromJSON(JSON.parse(snapSheet(s))); d.name = uniqueName(s.name + ' (2)'); book.sheets.splice(i + 1, 0, d); switchSheet(i + 1); shell.markDirty(); } },
      { label: 'Переместить влево', icon: 'chevl', disabled: i === 0, action: () => { [book.sheets[i - 1], book.sheets[i]] = [book.sheets[i], book.sheets[i - 1]]; book.active = i - 1; renderTabs(); shell.markDirty(); } },
      { label: 'Переместить вправо', icon: 'chevr', disabled: i === book.sheets.length - 1, action: () => { [book.sheets[i + 1], book.sheets[i]] = [book.sheets[i], book.sheets[i + 1]]; book.active = i + 1; renderTabs(); shell.markDirty(); } },
      { label: 'Цвет ярлычка', icon: 'palette', action: () => colorPicker({ x: e.clientX, y: e.clientY }, (c) => { s.color = c; renderTabs(); shell.markDirty(); }, { allowNone: true }) },
      '-', { label: 'Удалить', icon: 'trash', danger: true, disabled: book.sheets.length === 1, action: async () => { if (await confirmDlg('Удалить лист', `Удалить лист «${s.name}»? Это действие нельзя отменить.`, 'Удалить', true)) { book.sheets.splice(i, 1); undoStack.length = 0; switchSheet(Math.max(0, i - 1)); afterChange(); } } },
    ]); };
    tabsBar.append(t);
  });
}
function uniqueName(base) { let n = base, i = 2; while (book.sheets.some((s) => s.name.toLowerCase() === n.toLowerCase())) n = base.replace(/\s*\(\d+\)$/, '') + ` (${i++})`; return n; }
function addSheet() { let i = book.sheets.length + 1; while (book.sheets.some((s) => s.name === 'Лист' + i)) i++; book.sheets.push(newSheet('Лист' + i)); switchSheet(book.sheets.length - 1); shell.markDirty(); }
function switchSheet(i) { if (editing && !(editor.value[0] === '=')) commitEdit(); book.active = i; sel = sh().lastSel || { ar: 0, ac: 0, fr: 0, fc: 0 }; selChart = null; rebuildGeom(); renderTabs(); renderCharts(); draw(); syncBars(); scroller.scrollTo(0, 0); }
function renameSheet(i, tabEl) {
  const s = book.sheets[i]; const inp = h('input', { value: s.name }); tabEl.innerHTML = ''; tabEl.append(inp); inp.focus(); inp.select();
  const done = (ok) => { const n = inp.value.trim().replace(/[\\/?*[\]:]/g, ''); if (ok && n && n !== s.name) { if (book.sheets.some((x, j) => j !== i && x.name.toLowerCase() === n.toLowerCase())) toast('Лист с таким именем уже есть', 'error'); else { const old = s.name; s.name = n; for (const sheet of book.sheets) for (const cell of sheet.cells.values()) if (typeof cell.v === 'string' && cell.v[0] === '=' && cell.v.includes('!')) cell.v = renameSheetInFormula(cell.v, old, n); engine.invalidate(); shell.markDirty(); } } renderTabs(); sink.focus(); };
  inp.onkeydown = (e) => { if (e.key === 'Enter') done(true); if (e.key === 'Escape') done(false); e.stopPropagation(); }; inp.onblur = () => done(true);
}

/* ======================= serialization ======================= */
function sheetToJSON(s) {
  const cells = []; for (const [k, c] of s.cells) { const [r, cc] = RC(k); cells.push([r, cc, c.v ?? null, c.s || 0, c.n || 0]); }
  return { name: s.name, cells, colW: s.colW, rowH: s.rowH, merges: s.merges, freeze: s.freeze, charts: s.charts, cf: s.cf, filter: s.filter, color: s.color, gridlines: s.gridlines, validations: s.validations, print: s.print, protected: s.protected };
}
function sheetFromJSON(d) {
  const s = newSheet(d.name);
  for (const [r, c, v, st, n] of d.cells || []) { const cell = {}; if (v != null) cell.v = v; if (st) cell.s = st; if (n) cell.n = n; s.cells.set(K(r, c), cell); }
  Object.assign(s, { colW: d.colW || {}, rowH: d.rowH || {}, merges: d.merges || [], freeze: d.freeze || { r: 0, c: 0 }, charts: d.charts || [], cf: d.cf || [], filter: d.filter || null, color: d.color || null, gridlines: d.gridlines !== false, validations: d.validations || [], print: d.print || {}, protected: !!d.protected });
  return s;
}
function serialize() {
  // dedupe styles
  const styles = []; const idx = new Map();
  const sheets = book.sheets.map((s) => { const j = sheetToJSON(s); j.cells = j.cells.map(([r, c, v, st, n]) => { if (st) { const key = JSON.stringify(st); if (!idx.has(key)) { idx.set(key, styles.length + 1); styles.push(st); } st = idx.get(key); } return n ? [r, c, v, st, n] : st ? [r, c, v, st] : [r, c, v]; }); return j; });
  return JSON.stringify({ format: 'rabbit-workbook', version: 1, active: book.active, styles, sheets, names: book.names || {} });
}
function loadJSON(d) {
  const styles = d.styles || [];
  book.sheets = d.sheets.map((sj) => { sj.cells = (sj.cells || []).map(([r, c, v, st, n]) => [r, c, v, typeof st === 'number' && st ? styles[st - 1] : st || 0, n || 0]); return sheetFromJSON(sj); });
  if (!book.sheets.length) book.sheets.push(newSheet('Лист1'));
  book.active = clamp(d.active || 0, 0, book.sheets.length - 1);
  book.names = d.names || {};
}
function parseCSV(text) {
  const firstLine = text.split('\n')[0]; const delim = [';', '\t', ','].map((d) => [d, firstLine.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
    if (ch === '"') q = true; else if (ch === delim) { row.push(cur); cur = ''; } else if (ch === '\n') { row.push(cur.replace(/\r$/, '')); rows.push(row); row = []; cur = ''; } else cur += ch;
  }
  if (cur || row.length) { row.push(cur.replace(/\r$/, '')); rows.push(row); }
  return rows;
}
function toCSV(delim = ',') {
  const u = engine.usedBounds(book.active); const lines = [];
  for (let r = 0; r <= u.mr; r++) { const row = []; for (let c = 0; c <= u.mc; c++) { let v = engine.value(book.active, r, c); v = v == null ? '' : typeof v === 'number' ? (delim === ';' ? String(v).replace('.', ',') : String(v)) : formatValue(v, getCell(r, c)?.s); if (/["\n\r]/.test(v) || v.includes(delim)) v = '"' + v.replace(/"/g, '""') + '"'; row.push(v); } lines.push(row.join(delim)); }
  return '﻿' + lines.join('\r\n');
}
async function load(file) {
  const ext = extOf(file.name); undoStack.length = 0; redoStack.length = 0;
  if (ext === 'rxel') loadJSON(JSON.parse(new TextDecoder().decode(file.bytes)));
  else if (ext === 'xlsx') { const d = await importXlsx(file.bytes); book.sheets = d.sheets.map((sj) => sheetFromJSON(sj)); book.active = 0; }
  else { const text = new TextDecoder().decode(file.bytes).replace(/^﻿/, ''); const rows = parseCSV(text); book.sheets = [newSheet(file.name.replace(/\.[^.]+$/, '').slice(0, 30) || 'Лист1')]; book.active = 0; rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') { const p = parseInput(v); const cell = { v: p.v }; if (p.fmt) cell.s = { fmt: p.fmt, cur: p.cur }; sh().cells.set(K(r, c), cell); } })); }
  sel = { ar: 0, ac: 0, fr: 0, fc: 0 }; engine.invalidate(); engine.astCache.clear();
  rebuildGeom(); renderTabs(); renderCharts(); draw(); syncBars();
}

/* ======================= templates ======================= */
function tpl(id) {
  const s = newSheet('Лист1'); book.sheets = [s]; book.active = 0;
  const put = (ref, v, st) => { const g = parseRange(ref); const cell = {}; if (v != null) cell.v = v; if (st) cell.s = st; s.cells.set(K(g.r1, g.c1), cell); };
  const H = { b: true, bg: '#1d8a52', fc: '#ffffff' };
  if (id === 'budget') {
    s.name = 'Бюджет'; s.colW = { 0: 200, 1: 110, 2: 110, 3: 110 }; s.rowH = { 0: 34 };
    put('A1', 'Личный бюджет', { b: true, fs: 18, fc: '#1d8a52' }); put('A2', 'Месяц: ' + new Date().toLocaleDateString('ru-RU', { month: 'long', year: 'numeric' }), { fc: '#777' });
    ['Статья', 'План', 'Факт', 'Разница'].forEach((t, i) => put(colName(i) + '4', t, H));
    const rows = [['Зарплата', 120000, 120000], ['Фриланс', 20000, 26500], ['Аренда', -45000, -45000], ['Продукты', -25000, -27300], ['Транспорт', -6000, -5200], ['Развлечения', -10000, -12800], ['Сбережения', -30000, -30000]];
    rows.forEach(([n, p, f], i) => { const r = 5 + i; put('A' + r, n); put('B' + r, p, { fmt: 'currency', dec: 0 }); put('C' + r, f, { fmt: 'currency', dec: 0 }); put('D' + r, `=C${r}-B${r}`, { fmt: 'currency', dec: 0 }); });
    const t = 5 + rows.length; put('A' + t, 'Итого', { b: true, bd: { t: 1 } }); ['B', 'C', 'D'].forEach((c) => put(c + t, `=SUM(${c}5:${c}${t - 1})`, { b: true, fmt: 'currency', dec: 0, bd: { t: 1 } }));
    s.cf.push({ id: uid('cf'), type: 'lt', range: { r1: 4, c1: 3, r2: t - 2, c2: 3 }, a: '0', style: { bg: '#ffc7ce', fc: '#9c0006' } });
    s.charts.push({ id: uid('ch'), type: 'bar', title: 'План и факт', range: { r1: 3, c1: 0, r2: t - 2, c2: 2 }, x: 540, y: 20, w: 460, h: 300, legend: true });
    s.freeze = { r: 4, c: 0 };
  } else if (id === 'invoice') {
    s.name = 'Счёт'; s.colW = { 0: 40, 1: 260, 2: 70, 3: 110, 4: 120 }; s.rowH = { 0: 30 };
    put('A1', 'Счёт на оплату № 1', { b: true, fs: 16 }); put('A2', 'от ' + new Date().toLocaleDateString('ru-RU'), { fc: '#777' });
    put('A4', 'Поставщик:', { b: true }); put('B4', 'ООО «Кролик»'); put('A5', 'Покупатель:', { b: true }); put('B5', 'ИП Иванов И. И.');
    ['№', 'Товар / услуга', 'Кол-во', 'Цена', 'Сумма'].forEach((t, i) => put(colName(i) + '7', t, { ...H, bg: '#d2552d', al: 'center' }));
    [['Разработка сайта', 1, 85000], ['Хостинг (12 мес.)', 12, 900], ['Поддержка, час', 10, 2500]].forEach(([n, q, p], i) => { const r = 8 + i; put('A' + r, i + 1, { al: 'center' }); put('B' + r, n); put('C' + r, q, { al: 'center' }); put('D' + r, p, { fmt: 'currency' }); put('E' + r, `=C${r}*D${r}`, { fmt: 'currency' }); });
    put('D12', 'Итого:', { b: true, al: 'right' }); put('E12', '=SUM(E8:E10)', { b: true, fmt: 'currency' });
    put('D13', 'НДС 20%:', { al: 'right' }); put('E13', '=ROUND(E12*0.2,2)', { fmt: 'currency' });
    put('D14', 'К оплате:', { b: true, al: 'right' }); put('E14', '=E12+E13', { b: true, fmt: 'currency', bg: '#fff2cc' });
    for (let r = 6; r <= 9; r++) for (let c = 0; c < 5; c++) { const cell = s.cells.get(K(r, c)) || {}; cell.s = { ...(cell.s || {}), bd: { t: 1, r: 1, b: 1, l: 1 } }; s.cells.set(K(r, c), cell); }
  } else if (id === 'schedule') {
    s.name = 'План'; s.colW = { 0: 240, 1: 120, 2: 100, 3: 100, 4: 90, 5: 110 };
    ['Задача', 'Ответственный', 'Начало', 'Конец', 'Дней', 'Статус'].forEach((t, i) => put(colName(i) + '1', t, { ...H, bg: '#7b4fd6' }));
    const today = dateToSerial(new Date().getFullYear(), new Date().getMonth() + 1, new Date().getDate());
    [['Исследование', 'Анна', 0, 5, 'Готово'], ['Прототип', 'Борис', 5, 12, 'В работе'], ['Дизайн', 'Анна', 10, 20, 'В работе'], ['Разработка', 'Вера', 15, 40, 'Не начато'], ['Тестирование', 'Глеб', 38, 48, 'Не начато'], ['Запуск', 'Все', 48, 50, 'Не начато']].forEach(([t, w, a, b, st], i) => { const r = 2 + i; put('A' + r, t); put('B' + r, w); put('C' + r, today + a, { fmt: 'date' }); put('D' + r, today + b, { fmt: 'date' }); put('E' + r, `=NETWORKDAYS(C${r},D${r})`, { al: 'center' }); put('F' + r, st); });
    s.cf.push({ id: uid('cf'), type: 'text', range: { r1: 1, c1: 5, r2: 50, c2: 5 }, a: 'Готово', style: { bg: '#c6efce', fc: '#006100' } }, { id: uid('cf'), type: 'text', range: { r1: 1, c1: 5, r2: 50, c2: 5 }, a: 'В работе', style: { bg: '#ffeb9c', fc: '#9c5700' } }, { id: uid('cf'), type: 'bar', range: { r1: 1, c1: 4, r2: 50, c2: 4 }, color: '#8a5cd6' });
    s.freeze = { r: 1, c: 0 }; s.filter = { r: 0, c1: 0, c2: 5, crit: {} };
  }
}
async function newDoc(t) {
  undoStack.length = 0; redoStack.length = 0;
  if (t) tpl(t); else { book.sheets = [newSheet('Лист1')]; book.active = 0; }
  shell.setTitle(t ? TEMPLATES.xel.find((x) => x.id === t)?.name || 'Книга' : 'Книга');
  sel = { ar: 0, ac: 0, fr: 0, fc: 0 }; engine.invalidate();
  rebuildGeom(); renderTabs(); renderCharts(); draw(); syncBars();
}

/* ======================= print ======================= */
async function doPrint(pdf = false) {
  const g0 = rng(); const used = engine.usedBounds(book.active); const ps = sh().print || {};
  const g = ps.area ? ps.area : g0.r1 !== g0.r2 || g0.c1 !== g0.c2 ? selRangeForOps() : { r1: 0, c1: 0, r2: used.mr, c2: used.mc };
  $('#print-area')?.remove(); $('#print-page')?.remove();
  const border = ps.grid ? 'border:1px solid #bbb;' : '';
  let html = '<table>';
  if (ps.headers) { html += '<tr><td></td>'; for (let c = g.c1; c <= g.c2; c++) if (colW(c)) html += `<td style="background:#eee;text-align:center;${border}">${colName(c)}</td>`; html += '</tr>'; }
  for (let r = g.r1; r <= g.r2; r++) {
    if (hiddenRows.has(r)) continue; html += '<tr>';
    if (ps.headers) html += `<td style="background:#eee;text-align:center;${border}">${r + 1}</td>`;
    for (let c = g.c1; c <= g.c2; c++) {
      if (!colW(c)) continue;
      const m = mergeAt(r, c); if (m && (m.r1 !== r || m.c1 !== c)) continue;
      const s = getCell(r, c)?.s || {}; const v = engine.value(book.active, r, c); const al = s.al || (typeof v === 'number' ? 'right' : 'left');
      const bd = s.bd ? ['t', 'r', 'b', 'l'].map((k, i) => (s.bd[k] ? `border-${['top', 'right', 'bottom', 'left'][i]}:${s.bd[k]}px ${s.bd[k] === 3 ? 'double' : 'solid'} #222;` : '')).join('') : '';
      html += `<td${m ? ` colspan="${m.c2 - m.c1 + 1}" rowspan="${m.r2 - m.r1 + 1}"` : ''} style="text-align:${al};${border}${bd}${s.b ? 'font-weight:bold;' : ''}${s.i ? 'font-style:italic;' : ''}${s.u ? 'text-decoration:underline;' : ''}${s.bg ? 'background:' + s.bg + ';' : ''}${s.fc ? 'color:' + s.fc + ';' : ''}${s.fs ? 'font-size:' + s.fs + 'pt;' : ''}${s.wr ? 'white-space:normal;' : ''}min-width:${colW(c) * 0.75}pt">${esc(displayOf(r, c))}</td>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  const charts = sh().charts.map((ch) => `<div style="margin-top:16px;width:${ch.w}px;height:${ch.h}px;break-inside:avoid">${renderChart({ type: ch.type, title: ch.title, ...chartData(ch) }, ch.w, ch.h)}</div>`).join('');
  const landscape = ps.orient ? ps.orient === 'landscape' : g.c2 - g.c1 > 6;
  document.head.append(h('style', { id: 'print-page' }, `@page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 12mm; } ${ps.fit ? '#print-area table { width: 100%; table-layout: auto; } #print-area { zoom: var(--fitz, 1); }' : ''}`));
  const area = h('div', { id: 'print-area', html: `<h3 style="font:600 12pt sans-serif">${esc(shell.title)} — ${esc(sh().name)}</h3>` + html + charts });
  document.body.append(area);
  if (ps.fit) { area.style.display = 'block'; area.style.position = 'absolute'; area.style.visibility = 'hidden'; const w = area.querySelector('table').scrollWidth; const pageW = (landscape ? 273 : 186) * 3.78; area.style.cssText = ''; area.style.setProperty('--fitz', Math.min(1, pageW / w).toFixed(3)); }
  if (pdf) { const p = await printToPdf(shell.title + '.pdf', { landscape, preferCSSPageSize: true }); if (p) toast('PDF сохранён', 'ok'); } else window.print();
  setTimeout(() => { $('#print-area')?.remove(); $('#print-page')?.remove(); }, 1000);
}

/* ======================= protection ======================= */
class ProtectedSheet extends Error { constructor() { super('protected'); this.protectedSheet = true; } }
window.addEventListener('error', (e) => { if (e.error?.protectedSheet) e.preventDefault(); });
window.addEventListener('unhandledrejection', (e) => { if (e.reason?.protectedSheet) e.preventDefault(); });
let protectWarnT = 0;
function guard() {
  if (!sh().protected) return true;
  if (Date.now() - protectWarnT > 1500) { toast('Лист защищён от изменений. Снять защиту: Рецензирование → Защита листа', 'error', 4000); protectWarnT = Date.now(); }
  return false;
}
function toggleProtect() { sh().protected = !sh().protected; shell.markDirty(); syncBars(); toast(sh().protected ? 'Лист защищён' : 'Защита снята', 'ok'); }

/* ======================= data validation ======================= */
const validationAt = (r, c) => sh().validations?.find((v) => r >= v.range.r1 && r <= v.range.r2 && c >= v.range.c1 && c <= v.range.c2) || null;
function listValues(v) {
  const src = String(v.list || '').trim();
  if (src.startsWith('=')) { const g = parseRange(src.slice(1).replace(/\$/g, '')); if (!g) return []; const out = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) for (let c = g.c1; c <= g.c2; c++) { const d = displayOf(r, c); if (d !== '') out.push(d); } return out; }
  return src.split(/[;,\n]/).map((x) => x.trim()).filter(Boolean);
}
function validateInput(r, c, text) {
  const v = validationAt(r, c); if (!v || text === '' || text == null || String(text)[0] === '=') return null;
  const msg = (def) => v.msg || def; const p = parseInput(text);
  if (v.type === 'list') { const vals = listValues(v).map((x) => x.toLowerCase()); return vals.includes(String(text).trim().toLowerCase()) ? null : msg('Значение должно быть из списка: ' + listValues(v).slice(0, 8).join(', ')); }
  if (v.type === 'number' || v.type === 'integer') {
    if (typeof p.v !== 'number') return msg('Введите число');
    if (v.type === 'integer' && !Number.isInteger(p.v)) return msg('Введите целое число');
    if (v.min !== '' && v.min != null && p.v < +v.min) return msg(`Число должно быть не меньше ${v.min}`);
    if (v.max !== '' && v.max != null && p.v > +v.max) return msg(`Число должно быть не больше ${v.max}`);
  }
  if (v.type === 'date' && !/date/.test(p.fmt || '')) return msg('Введите дату в формате ДД.ММ.ГГГГ');
  if (v.type === 'textlen') { const n = String(text).length; if ((v.min !== '' && n < +v.min) || (v.max !== '' && n > +v.max)) return msg(`Длина текста: от ${v.min || 0} до ${v.max || '∞'} символов`); }
  return null;
}
async function validationDialog() {
  const g = selRangeForOps(); const cur = validationAt(g.r1, g.c1) || {};
  const r = await formDlg('Проверка данных · ' + rangeName(g), [
    { key: 'type', label: 'Тип данных', type: 'select', value: cur.type || 'list', options: [{ value: 'list', label: 'Список (выпадающий)' }, { value: 'number', label: 'Число' }, { value: 'integer', label: 'Целое число' }, { value: 'date', label: 'Дата' }, { value: 'textlen', label: 'Длина текста' }, { value: 'none', label: 'Любое значение (удалить проверку)' }] },
    { key: 'list', label: 'Источник списка', value: cur.list || 'Да; Нет; Возможно', hint: 'Значения через «;» или ссылка на диапазон, например =$H$1:$H$10' },
    { key: 'min', label: 'Минимум', value: cur.min ?? '' }, { key: 'max', label: 'Максимум', value: cur.max ?? '' },
    { key: 'msg', label: 'Сообщение об ошибке (необязательно)', value: cur.msg || '' },
  ], 'Применить');
  if (!r) return;
  begin(true);
  sh().validations = (sh().validations || []).filter((v) => v.range.r2 < g.r1 || v.range.r1 > g.r2 || v.range.c2 < g.c1 || v.range.c1 > g.c2);
  if (r.type !== 'none') sh().validations.push({ range: g, type: r.type, list: r.list, min: r.min, max: r.max, msg: r.msg });
  commit(); toast(r.type === 'none' ? 'Проверка данных удалена' : 'Проверка данных установлена', 'ok');
}
function showListPicker(clientX, clientY) {
  const v = validationAt(sel.ar, sel.ac); if (!v) return;
  const vals = listValues(v);
  menu({ x: clientX, y: clientY }, vals.length ? vals.map((x) => ({ label: x, checked: displayOf(sel.ar, sel.ac) === x, action: () => { if (!guard()) return; begin(); inputCell(sel.ar, sel.ac, x); commit(); } })) : [{ label: 'Список пуст', disabled: true }]);
}

/* ======================= links ======================= */
function linkOf(cell, v) {
  if (!cell) return null;
  if (typeof cell.v === 'string' && /^=\s*(HYPERLINK|ГИПЕРССЫЛКА)\s*\(/i.test(cell.v)) { const m = /\(\s*"([^"]+)"/.exec(cell.v); return m ? m[1] : null; }
  if (typeof v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(v)) return v.startsWith('www.') ? 'https://' + v : v;
  return null;
}
async function insertLinkDialog() {
  const cell = getCell(sel.ar, sel.ac); const cur = linkOf(cell, engine.value(book.active, sel.ar, sel.ac));
  const r = await formDlg('Гиперссылка', [{ key: 'text', label: 'Текст', value: displayOf(sel.ar, sel.ac) || '' }, { key: 'url', label: 'Адрес', value: cur || 'https://' }], 'Вставить');
  if (!r || !r.url || !guard()) return;
  begin(); setCell(sel.ar, sel.ac, { v: `=HYPERLINK("${r.url.replace(/"/g, '')}","${(r.text || r.url).replace(/"/g, "'")}")` }); commit();
}
// clicks inside the grid that are not plain selection: list dropdown arrow, Ctrl+click on a link
function onCellClick(p, e) {
  const v = validationAt(sel.ar, sel.ac);
  if (v?.type === 'list') { const r = cellScreenRect(sel.ar, sel.ac); if (p.lx >= r.x + r.w && p.lx <= r.x + r.w + 16 && p.ly >= r.y && p.ly <= r.y + r.h) { showListPicker(e.clientX, e.clientY); return true; } }
  if (e.ctrlKey || e.metaKey) { const url = linkOf(getCell(p.r, p.c), engine.value(book.active, p.r, p.c)); if (url) { window.open(url, '_blank'); return true; } }
  return false;
}

/* ======================= names ======================= */
const validName = (n) => /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^(TRUE|FALSE|ИСТИНА|ЛОЖЬ)$/i.test(n);
async function defineName() {
  const g = selRangeForOps();
  const n = await promptDlg('Присвоить имя', `Имя для диапазона ${sh().name}!${rangeName(g)}`, '');
  if (!n) return; if (!validName(n.trim())) return toast('Имя должно начинаться с буквы и не совпадать с адресом ячейки', 'error');
  book.names ||= {}; book.names[n.trim().toUpperCase()] = { label: n.trim(), sheet: sh().name, ref: rangeName(g) };
  engine.invalidate(); afterChange(); toast(`Имя «${n.trim()}» создано — используйте его в формулах`, 'ok');
}
function namesDialog() {
  const body = h('div', { style: { width: 'min(520px,100%)' } });
  const draw = () => {
    body.innerHTML = '';
    const list = Object.entries(book.names || {});
    if (!list.length) body.append(h('p', { class: 'muted' }, 'Имён пока нет. Выделите диапазон и нажмите «Присвоить имя» — затем пишите в формулах, например =СУММ(Продажи).'));
    const t = h('table', { class: 'data' }, h('tr', {}, h('th', {}, 'Имя'), h('th', {}, 'Диапазон'), h('th', {})));
    for (const [k, v] of list) t.append(h('tr', {}, h('td', {}, h('b', {}, v.label || k)), h('td', {}, `${v.sheet}!${v.ref}`), h('td', {}, h('button', { class: 'btn sm', onclick: () => { const i = book.sheets.findIndex((s) => s.name === v.sheet); if (i >= 0) switchSheet(i); const g = parseRange(v.ref); sel = { ar: g.r1, ac: g.c1, fr: g.r2, fc: g.c2 }; ensureVisible(g.r1, g.c1); draw2(); } }, 'Перейти'), ' ', h('button', { class: 'btn sm danger', onclick: () => { delete book.names[k]; engine.invalidate(); afterChange(); draw(); } }, 'Удалить'))));
    if (list.length) body.append(t);
  };
  const draw2 = () => { draw(); syncBars(); };
  draw();
  dialog({ title: 'Диспетчер имён', body, buttons: [{ label: 'Создать из выделения…', value: () => { defineName().then(draw); return undefined; } }, { label: 'Закрыть', value: true, primary: true }] });
}

/* ======================= custom filter ======================= */
async function customFilter(c) {
  const s = sh(); if (!s.filter) { toggleFilter(); }
  const f = sh().filter; if (c < f.c1 || c > f.c2) c = f.c1;
  const cur = f.cond?.[c] || {};
  const cols = []; for (let cc = f.c1; cc <= f.c2; cc++) cols.push({ value: cc, label: `${colName(cc)} — ${displayOf(f.r, cc) || 'столбец'}` });
  const r = await formDlg('Фильтр по условию', [{ key: 'col', label: 'Столбец', type: 'select', options: cols, value: c }, { key: 'op', label: 'Условие', type: 'select', options: FILTER_OPS.map(([v, l]) => ({ value: v, label: l })), value: cur.op || 'gt' }, { key: 'a', label: 'Значение (для «между» — от, для «первые N» — N)', value: cur.a ?? '' }, { key: 'b', label: 'До (для «между»)', value: cur.b ?? '' }, { key: 'clear', label: 'Снять условие с этого столбца', type: 'checkbox' }], 'Применить');
  if (!r) return;
  begin(true); f.cond ||= {};
  const col = +r.col;
  if (r.clear) delete f.cond[col];
  else {
    const cond = { op: r.op, a: r.a, b: r.b };
    const last = lastDataRow(f.c1, f.c2, f.r); const nums = []; for (let rr = f.r + 1; rr <= last; rr++) { const v = engine.value(book.active, rr, col); if (typeof v === 'number') nums.push(v); }
    if (r.op === 'top') { const n = Math.max(1, +r.a || 10); cond.threshold = [...nums].sort((x, y) => y - x)[Math.min(n, nums.length) - 1]; }
    if (r.op === 'above' || r.op === 'below') cond.avg = nums.reduce((x, y) => x + y, 0) / (nums.length || 1);
    f.cond[col] = cond;
  }
  commit(); rebuildGeom(); draw(); syncBars();
}

/* ======================= pivot table ======================= */
async function pivotDialog() {
  const g0 = rng(); const g = g0.r1 === g0.r2 && g0.c1 === g0.c2 ? currentRegion(sel.ar, sel.ac) : selRangeForOps();
  if (g.r2 - g.r1 < 1) return toast('Выделите таблицу с заголовками и хотя бы одной строкой данных', 'error');
  const fields = []; for (let c = g.c1; c <= g.c2; c++) fields.push({ value: c, label: displayOf(g.r1, c) || colName(c) });
  const numField = fields.find((f) => typeof engine.value(book.active, g.r1 + 1, f.value) === 'number')?.value ?? fields[fields.length - 1].value;
  const r = await formDlg('Сводная таблица · ' + rangeName(g), [
    { key: 'rows', label: 'Строки', type: 'select', options: fields, value: fields[0].value },
    { key: 'cols', label: 'Столбцы (необязательно)', type: 'select', options: [{ value: '', label: '(нет)' }, ...fields] },
    { key: 'val', label: 'Значения', type: 'select', options: fields, value: numField },
    { key: 'agg', label: 'Операция', type: 'select', options: [{ value: 'sum', label: 'Сумма' }, { value: 'count', label: 'Количество' }, { value: 'avg', label: 'Среднее' }, { value: 'min', label: 'Минимум' }, { value: 'max', label: 'Максимум' }] },
  ], 'Создать');
  if (!r) return;
  const rc = +r.rows, cc = r.cols === '' ? null : +r.cols, vc = +r.val;
  const groups = new Map(); const colKeys = new Set();
  for (let rr = g.r1 + 1; rr <= g.r2; rr++) {
    if (hiddenRows.has(rr)) continue;
    const rk = displayOf(rr, rc) || '(пусто)'; const ck = cc == null ? 'Итого' : displayOf(rr, cc) || '(пусто)'; colKeys.add(ck);
    const v = engine.value(book.active, rr, vc);
    if (!groups.has(rk)) groups.set(rk, new Map()); const gm = groups.get(rk); if (!gm.has(ck)) gm.set(ck, []); gm.get(ck).push(v);
  }
  const agg = (arr) => { const nums = arr.filter((x) => typeof x === 'number'); if (r.agg === 'count') return arr.filter((x) => x != null && x !== '').length; if (!nums.length) return null; if (r.agg === 'sum') return nums.reduce((a, b) => a + b, 0); if (r.agg === 'avg') return nums.reduce((a, b) => a + b, 0) / nums.length; if (r.agg === 'min') return Math.min(...nums); return Math.max(...nums); };
  const cols = [...colKeys].sort((a, b) => a.localeCompare(b, 'ru', { numeric: true })); const rows = [...groups.keys()].sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }));
  let n = 1; while (book.sheets.some((s) => s.name === 'Сводная' + n)) n++;
  const ps = newSheet('Сводная' + n); const put = (rr, c, v, s) => { const cell = {}; if (v != null) cell.v = v; if (s) cell.s = s; ps.cells.set(K(rr, c), cell); };
  const aggName = { sum: 'Сумма', count: 'Количество', avg: 'Среднее', min: 'Минимум', max: 'Максимум' }[r.agg];
  const head = { b: true, bg: '#1d8a52', fc: '#ffffff' }; const numSt = r.agg === 'count' ? { fmt: 'number', dec: 0 } : { fmt: 'number', dec: 2 };
  put(0, 0, `${aggName} по полю «${fields.find((f) => f.value === vc).label}»`, { b: true, fs: 13 });
  put(2, 0, fields.find((f) => f.value === rc).label, head);
  const showTotalCol = cc != null;
  cols.forEach((k, i) => put(2, 1 + i, k, { ...head, al: 'center' })); if (showTotalCol) put(2, 1 + cols.length, 'Общий итог', { ...head, al: 'center' });
  rows.forEach((rk, i) => {
    const rr = 3 + i; put(rr, 0, rk, i % 2 ? { bg: '#f2f7f4' } : null); const all = [];
    cols.forEach((ck, j) => { const arr = groups.get(rk).get(ck) || []; all.push(...arr); put(rr, 1 + j, agg(arr), { ...numSt, ...(i % 2 ? { bg: '#f2f7f4' } : {}) }); });
    if (showTotalCol) put(rr, 1 + cols.length, agg(all), { ...numSt, b: true, ...(i % 2 ? { bg: '#f2f7f4' } : {}) });
  });
  const tr = 3 + rows.length; put(tr, 0, 'Общий итог', { b: true, bd: { t: 1 } });
  const colAll = (ck) => rows.flatMap((rk) => groups.get(rk).get(ck) || []);
  cols.forEach((ck, j) => put(tr, 1 + j, agg(colAll(ck)), { ...numSt, b: true, bd: { t: 1 } }));
  if (showTotalCol) put(tr, 1 + cols.length, agg(cols.flatMap(colAll)), { ...numSt, b: true, bd: { t: 1 } });
  ps.colW = { 0: 180 }; for (let j = 1; j <= cols.length + 1; j++) ps.colW[j] = 110;
  ps.freeze = { r: 3, c: 1 };
  book.sheets.push(ps); undoStack.length = 0; switchSheet(book.sheets.length - 1); shell.markDirty();
  toast('Сводная таблица создана на листе ' + ps.name, 'ok');
}

/* ======================= format as table & cell styles ======================= */
const TABLE_STYLES = [['Синий', '#2f62c9', '#e8eefb'], ['Зелёный', '#1d8a52', '#e6f4ec'], ['Оранжевый', '#d2552d', '#fbeae4'], ['Фиолетовый', '#7b4fd6', '#f0eafb'], ['Серый', '#4b5563', '#f1f2f4'], ['Тёмный', '#1f2937', '#e5e7eb']];
function formatAsTable([name, head, band]) {
  if (!guard()) return;
  const g0 = rng(); const g = g0.r1 === g0.r2 && g0.c1 === g0.c2 ? currentRegion(sel.ar, sel.ac) : selRangeForOps();
  begin(true);
  for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
    const cell = getCell(r, c) || {}; const s = { ...(cell.s || {}) };
    if (r === g.r1) Object.assign(s, { b: true, bg: head, fc: '#ffffff' }); else { s.bg = (r - g.r1) % 2 === 0 ? band : undefined; if (!s.bg) delete s.bg; }
    s.bd = { t: 1, b: 1, l: c === g.c1 ? 1 : undefined, r: c === g.c2 ? 1 : undefined, c: '#c9ccd3' };
    sh().cells.set(K(r, c), { ...cell, s });
  }
  sh().filter = { r: g.r1, c1: g.c1, c2: g.c2, crit: {} };
  commit(); rebuildGeom(); draw(); toast(`Таблица «${name}» оформлена, фильтр включён`, 'ok');
}
const CELL_STYLES = [['Обычный', null], ['Хороший', { bg: '#c6efce', fc: '#006100' }], ['Плохой', { bg: '#ffc7ce', fc: '#9c0006' }], ['Нейтральный', { bg: '#ffeb9c', fc: '#9c5700' }], ['Заголовок 1', { b: true, fs: 15, fc: '#1f3864', bd: { b: 2 } }], ['Заголовок 2', { b: true, fs: 13, fc: '#1f3864', bd: { b: 1 } }], ['Итог', { b: true, bd: { t: 1, b: 3 } }], ['Ввод', { bg: '#ffcc99', fc: '#3f3f76', bd: { t: 1, b: 1, l: 1, r: 1, c: '#7f7f7f' } }], ['Вывод', { bg: '#f2f2f2', fc: '#3f3f3f', b: true, bd: { t: 1, b: 1, l: 1, r: 1, c: '#3f3f3f' } }], ['Примечание', { bg: '#ffffcc', bd: { t: 1, b: 1, l: 1, r: 1, c: '#b2b2b2' } }], ['Акцент', { bg: '#1d8a52', fc: '#ffffff', b: true }]];
function applyCellStyle(st) {
  forSel((r, c) => { const cell = getCell(r, c); recordCell(K(r, c)); const keep = cell?.s ? { fmt: cell.s.fmt, dec: cell.s.dec, cur: cell.s.cur, al: cell.s.al, wr: cell.s.wr } : {}; Object.keys(keep).forEach((k) => keep[k] == null && delete keep[k]); const s = st ? { ...keep, ...structuredClone(st) } : keep; const n = { ...(cell || {}), s: Object.keys(s).length ? s : undefined }; if (!n.s) delete n.s; if (n.v == null && !n.s && !n.n) sh().cells.delete(K(r, c)); else sh().cells.set(K(r, c), n); });
}

/* ======================= goal seek ======================= */
async function goalSeek() {
  const r = await formDlg('Подбор параметра', [{ key: 'set', label: 'Установить в ячейке (с формулой)', value: refName(sel.ar, sel.ac) }, { key: 'to', label: 'Значение', value: '' }, { key: 'by', label: 'Изменяя значение ячейки', value: '' }], 'Подобрать');
  if (!r) return;
  const a = parseRange(r.set), b = parseRange(r.by), target = +String(r.to).replace(',', '.');
  if (!a || !b || !isFinite(target)) return toast('Проверьте адреса ячеек и значение', 'error');
  if (typeof getCell(b.r1, b.c1)?.v === 'string' && getCell(b.r1, b.c1).v[0] === '=') return toast('Изменяемая ячейка должна содержать число, а не формулу', 'error');
  const k = K(b.r1, b.c1); const orig = sh().cells.get(k) ? structuredClone(sh().cells.get(k)) : undefined;
  const f = (x) => { const c = sh().cells.get(k) || {}; sh().cells.set(k, { ...c, v: x }); engine.invalidate(); const v = engine.value(book.active, a.r1, a.c1); return typeof v === 'number' ? v - target : NaN; };
  let x0 = typeof orig?.v === 'number' ? orig.v : 0, x1 = x0 ? x0 * 1.1 : 1; let f0 = f(x0), f1 = f(x1); let ok = false;
  for (let i = 0; i < 100 && isFinite(f1); i++) { if (Math.abs(f1) < 1e-9) { ok = true; break; } const d = f1 - f0; if (!d) break; const x2 = x1 - (f1 * (x1 - x0)) / d; x0 = x1; f0 = f1; x1 = x2; f1 = f(x1); }
  if (orig === undefined) sh().cells.delete(k); else sh().cells.set(k, orig); engine.invalidate();
  if (!ok && !(Math.abs(f1) < 1e-6)) { draw(); return toast('Решение не найдено', 'error'); }
  begin(); setCell(b.r1, b.c1, { v: +x1.toPrecision(12) }); commit(); select(b.r1, b.c1);
  toast(`Найдено: ${refName(b.r1, b.c1)} = ${formatValue(+x1.toPrecision(10), getCell(b.r1, b.c1)?.s)}`, 'ok', 5000);
}

/* ======================= page setup ======================= */
async function pageSetup() {
  const ps = sh().print || {};
  const r = await formDlg('Параметры страницы', [{ key: 'orient', label: 'Ориентация', type: 'select', value: ps.orient || 'auto', options: [{ value: 'auto', label: 'Автоматически' }, { value: 'portrait', label: 'Книжная' }, { value: 'landscape', label: 'Альбомная' }] }, { key: 'fit', label: 'Вписать по ширине страницы', type: 'checkbox', value: ps.fit }, { key: 'grid', label: 'Печатать сетку', type: 'checkbox', value: ps.grid }, { key: 'headers', label: 'Печатать заголовки строк и столбцов', type: 'checkbox', value: ps.headers }], 'OK');
  if (!r) return; sh().print = { ...ps, orient: r.orient === 'auto' ? null : r.orient, fit: r.fit, grid: r.grid, headers: r.headers }; shell.markDirty();
}

/* ======================= zoom & freeze ======================= */
function setZoom(z) { const ox = scroller.scrollLeft / Z, oy = scroller.scrollTop / Z; Z = clamp(Math.round(z * 100) / 100, 0.4, 3); rebuildGeom(); scroller.scrollLeft = ox * Z; scroller.scrollTop = oy * Z; zoomLbl.textContent = Math.round(Z * 100) + '%'; zoomRange.value = Math.round(Z * 100); if (editing) cancelEdit(); renderCharts(); draw(); }
function freeze(kind) { begin(true); const s = sh(); if (kind === 'none') s.freeze = { r: 0, c: 0 }; else if (kind === 'row') s.freeze = { r: 1, c: 0 }; else if (kind === 'col') s.freeze = { r: 0, c: 1 }; else s.freeze = { r: sel.ar, c: sel.ac }; commit(); scroller.scrollTo(0, 0); draw(); }

/* ======================= ribbon ======================= */
const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Segoe UI', 'Roboto', 'Verdana', 'Tahoma', 'Georgia', 'Courier New', 'Consolas', 'PT Sans'];
const FORMATS = [['general', 'Общий'], ['number', 'Числовой'], ['currency', 'Денежный'], ['percent', 'Процентный'], ['date', 'Дата'], ['datelong', 'Длинная дата'], ['time', 'Время'], ['datetime', 'Дата и время'], ['sci', 'Экспоненциальный'], ['text', 'Текстовый']];
const b = (id, ic, label, action, extra = {}) => ({ type: 'btn', id, icon: ic, label, action, ...extra });
const big = (id, ic, label, action, extra = {}) => ({ type: 'big', id, icon: ic, label, action, ...extra });
const ribbon = [
  { id: 'home', label: 'Главная', groups: [
    { label: 'Буфер обмена', items: [big('paste', 'paste', 'Вставить', async () => { try { doPaste(await navigator.clipboard.readText()); } catch { if (clip) doPaste(null); else toast('Используйте Ctrl+V'); } }), { type: 'col', items: [b('cut', 'cut', 'Вырезать', () => doCopy(null, true), { key: 'Ctrl+X' }), b('copy', 'copy', 'Копировать', () => doCopy(null), { key: 'Ctrl+C' }), b('fmtpaint', 'brush', 'Копировать формат', () => formatPainter())] }] },
    { label: 'Шрифт', items: [{ type: 'col', items: [
      { type: 'row', items: [{ type: 'select', id: 'font', width: 130, options: FONTS, value: DEFAULT_FONT, action: (v) => style({ ff: v === DEFAULT_FONT ? null : v }) }, { type: 'select', id: 'size', editable: true, width: 50, options: [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48], value: '11', action: (v) => { style({ fs: +String(v).replace(',', '.') === 11 ? null : +String(v).replace(',', '.') }); autoRowHeightForFont(); } }] },
      { type: 'row', items: [b('bold', 'bold', 'Полужирный', () => toggleStyle('b'), { key: 'Ctrl+B' }), b('italic', 'italic', 'Курсив', () => toggleStyle('i'), { key: 'Ctrl+I' }), b('underline', 'underline', 'Подчёркнутый', () => toggleStyle('u'), { key: 'Ctrl+U' }), b('strike', 'strike', 'Зачёркнутый', () => toggleStyle('s')),
        b('borders', 'borders', 'Границы', null, { menu: () => [{ label: 'Все границы', icon: 'borders', action: () => setBorders('all') }, { label: 'Внешние границы', icon: 'border-outer', action: () => setBorders('outer') }, { label: 'Толстая внешняя граница', icon: 'border-outer', action: () => setBorders('thick') }, { label: 'Нижняя граница', icon: 'border-bottom', action: () => setBorders('bottom') }, { label: 'Двойная нижняя граница', icon: 'border-bottom', action: () => setBorders('double') }, { label: 'Верхняя граница', icon: 'border-bottom', action: () => setBorders('top') }, { label: 'Левая граница', icon: 'border-bottom', action: () => setBorders('left') }, { label: 'Правая граница', icon: 'border-bottom', action: () => setBorders('right') }, '-', { label: 'Без границ', icon: 'border-none', action: () => setBorders('none') }] }),
        { type: 'color', id: 'bg', icon: 'fill-color', tip: 'Цвет заливки', value: '#ffff00', allowNone: true, action: (c) => style({ bg: c }) },
        { type: 'color', id: 'fc', icon: 'font-color', tip: 'Цвет текста', value: '#c00000', allowNone: true, action: (c) => style({ fc: c }) }] }] }] },
    { label: 'Выравнивание', items: [{ type: 'col', items: [
      { type: 'row', items: [b('va-t', 'valign-top', 'По верхнему краю', () => style({ va: 'top' })), b('va-m', 'valign-middle', 'По середине', () => style({ va: 'middle' })), b('va-b', 'valign-bottom', 'По нижнему краю', () => style({ va: null })), b('wrap', 'wrap', 'Переносить текст', () => { toggleStyle('wr'); const g = rng(); for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 200); r++) autoRowHeight(r); }, { showLabel: true })] },
      { type: 'row', items: [b('al-l', 'align-left', 'По левому краю', () => style({ al: 'left' })), b('al-c', 'align-center', 'По центру', () => style({ al: 'center' })), b('al-r', 'align-right', 'По правому краю', () => style({ al: 'right' })), b('merge', 'merge', 'Объединить и поместить в центре', null, { showLabel: true, menu: () => [{ label: 'Объединить и поместить в центре', action: () => mergeCells('center') }, { label: 'Объединить по строкам', action: () => mergeCells('across') }, { label: 'Объединить ячейки', action: () => mergeCells('merge') }, { label: 'Отменить объединение', action: () => { const m = mergeAt(sel.ar, sel.ac); if (m) mergeCells(); } }] })] }] }] },
    { label: 'Число', items: [{ type: 'col', items: [
      { type: 'select', id: 'fmt', width: 140, options: FORMATS.map(([v, l]) => ({ value: v, label: l })), value: 'general', action: (v) => style({ fmt: v === 'general' ? null : v, dec: null }) },
      { type: 'row', items: [b('cur', 'currency', 'Денежный формат', null, { menu: () => ['₽', '$', '€', '£', '¥', '₸'].map((s) => ({ label: s + '  ' + { '₽': 'Рубль', $: 'Доллар США', '€': 'Евро', '£': 'Фунт', '¥': 'Иена/юань', '₸': 'Тенге' }[s], action: () => style({ fmt: 'currency', cur: s }) })) }), b('pct', 'percent', 'Процентный формат', () => style({ fmt: 'percent' })), b('thou', 'hash', 'Формат с разделителями', () => style({ fmt: 'number' })), b('decinc', 'dec-inc', 'Увеличить разрядность', () => changeDecimals(1)), b('decdec', 'dec-dec', 'Уменьшить разрядность', () => changeDecimals(-1))] }] }] },
    { label: 'Стили', items: [big('cf', 'cond-format', 'Условное форматирование', null, { menu: () => [{ title: 'Правила выделения ячеек' }, { label: 'Больше…', action: () => addCF('gt') }, { label: 'Меньше…', action: () => addCF('lt') }, { label: 'Между…', action: () => addCF('between') }, { label: 'Равно…', action: () => addCF('eq') }, { label: 'Текст содержит…', action: () => addCF('text') }, { label: 'Повторяющиеся значения', action: () => addCF('dup') }, { title: 'Гистограммы и шкалы' }, { label: 'Гистограмма (синяя)', action: () => addCF('bar') }, { label: 'Гистограмма (зелёная)', action: () => addCF('bar2') }, { label: 'Цветовая шкала: красный → зелёный', action: () => addCF('scale') }, { label: 'Цветовая шкала: белый → синий', action: () => addCF('scale2') }, '-', { label: 'Удалить правила из выделенных ячеек', action: () => clearCF(false) }, { label: 'Удалить правила со всего листа', action: () => clearCF(true) }] }),
      big('fmtcells', 'settings', 'Формат ячеек', formatDialog, { key: 'Ctrl+1' })] },
    { label: 'Ячейки', items: [{ type: 'col', items: [b('ins', 'insert-row', 'Вставить', null, { showLabel: true, menu: () => [{ label: 'Строки выше', action: () => insertRows(true) }, { label: 'Строки ниже', action: () => insertRows(false) }, { label: 'Столбцы слева', action: () => insertCols(true) }, { label: 'Столбцы справа', action: () => insertCols(false) }, { label: 'Лист', action: addSheet }] }), b('del', 'delete-row', 'Удалить', null, { showLabel: true, menu: () => [{ label: 'Строки', action: deleteRows }, { label: 'Столбцы', action: deleteCols }] }), b('size', 'columns', 'Формат', null, { showLabel: true, menu: () => [{ label: 'Ширина столбца…', action: () => sizeDialog('col') }, { label: 'Автоподбор ширины', action: () => { const g = rng(); for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 50); c++) autoFitCol(c); } }, { label: 'Высота строки…', action: () => sizeDialog('row') }, '-', { label: 'Скрыть строки', action: () => hideRowsCols('row', true) }, { label: 'Показать строки', action: () => hideRowsCols('row', false) }, { label: 'Скрыть столбцы', action: () => hideRowsCols('col', true) }, { label: 'Показать столбцы', action: () => hideRowsCols('col', false) }] })] }] },
    { label: 'Редактирование', items: [{ type: 'col', items: [b('autosum', 'autosum', 'Автосумма', null, { showLabel: true, key: 'Alt+=', menu: () => [['SUM', 'Сумма'], ['AVERAGE', 'Среднее'], ['COUNT', 'Число'], ['MAX', 'Максимум'], ['MIN', 'Минимум']].map(([f, l]) => ({ label: l, action: () => autoSum(f) })) }), b('fill', 'fill-down', 'Заполнить', null, { showLabel: true, menu: () => [{ label: 'Вниз', hint: 'Ctrl+D', action: () => fillDir('down') }, { label: 'Вправо', hint: 'Ctrl+R', action: () => fillDir('right') }] }), b('clear', 'eraser', 'Очистить', null, { showLabel: true, menu: () => [{ label: 'Очистить всё', action: () => clearRange('all') }, { label: 'Очистить форматы', action: () => clearRange('format') }, { label: 'Очистить содержимое', action: () => clearRange('content') }] })] },
      { type: 'col', items: [b('sortf', 'sort-asc', 'Сортировка и фильтр', null, { showLabel: true, menu: () => [{ label: 'Сортировка по возрастанию', icon: 'sort-asc', action: () => quickSort(true) }, { label: 'Сортировка по убыванию', icon: 'sort-desc', action: () => quickSort(false) }, { label: 'Настраиваемая сортировка…', action: sortDialog }, '-', { label: 'Фильтр', icon: 'filter', action: toggleFilter }] }), b('findb', 'search', 'Найти и выделить', null, { showLabel: true, menu: () => [{ label: 'Найти…', hint: 'Ctrl+F', action: () => findDialog(false) }, { label: 'Заменить…', hint: 'Ctrl+H', action: () => findDialog(true) }, { label: 'Перейти к ячейке…', action: () => { nameBox.focus(); nameBox.select(); } }] })] }] },
  ] },
  { id: 'insert', label: 'Вставка', groups: [
    { label: 'Диаграммы', items: CHART_TYPES.map((t) => big('ch-' + t.id, t.icon, t.name, () => insertChart(t.id))) },
    { label: 'Функции', items: [big('insfn', 'fx', 'Функция', insertFunctionDialog, { key: 'Shift+F3' })] },
    { label: 'Примечания', items: [big('note', 'comment-plus', 'Примечание', noteDialog)] },
  ] },
  { id: 'formulas', label: 'Формулы', groups: [
    { label: 'Библиотека функций', items: [big('fx2', 'fx', 'Вставить функцию', insertFunctionDialog), big('as2', 'autosum', 'Автосумма', () => autoSum('SUM')),
      ...Object.keys(FN_HELP).map((cat) => big('cat-' + cat, { 'Математические': 'sigma', 'Статистические': 'chart', 'Логические': 'help', 'Текстовые': 'type', 'Дата и время': 'calendar', 'Ссылки и поиск': 'search', 'Финансовые': 'currency' }[cat], cat, null, { menu: () => FN_HELP[cat].map(([n, sig, d]) => ({ label: Object.entries(ALIAS).find(([, en]) => en === n)?.[0] || n, hint: d, action: () => { editSheet = book.active; startEdit('=' + (Object.entries(ALIAS).find(([, en]) => en === n)?.[0] || n) + '(', 'edit'); } })) }))] },
    { label: 'Вычисление', items: [big('recalc', 'redo', 'Пересчёт', () => { engine.invalidate(); draw(); renderCharts(); toast('Пересчитано', 'ok', 1200); }, { key: 'F9' }), big('showf', 'code', 'Показать формулы', () => { showFormulas = !showFormulas; shell.ribbon.set('showf', { on: showFormulas }); draw(); })] },
  ] },
  { id: 'data', label: 'Данные', groups: [
    { label: 'Сортировка и фильтр', items: [big('sa', 'sort-asc', 'А → Я', () => quickSort(true)), big('sd', 'sort-desc', 'Я → А', () => quickSort(false)), big('sortd', 'sort-asc', 'Сортировка', sortDialog), big('filter', 'filter', 'Фильтр', toggleFilter), big('clrf', 'eraser', 'Очистить', () => { const f = sh().filter; if (f) { begin(true); f.crit = {}; commit(); rebuildGeom(); } })] },
    { label: 'Работа с данными', items: [big('dedupe', 'delete-row', 'Удалить дубликаты', removeDuplicates), big('t2c', 'columns', 'Текст по столбцам', textToColumns)] },
  ] },
  { id: 'review', label: 'Рецензирование', groups: [
    { label: 'Примечания', items: [big('note2', 'comment-plus', 'Примечание', noteDialog), big('delnote', 'trash', 'Удалить', () => { begin(); setCell(sel.ar, sel.ac, { n: undefined }); const c = getCell(sel.ar, sel.ac); if (c) { delete c.n; if (c.v == null && !c.s) sh().cells.delete(K(sel.ar, sel.ac)); } commit(); })] },
    { label: 'Статистика', items: [big('stats', 'info', 'Статистика книги', () => { let cells = 0, formulas = 0; for (const s of book.sheets) for (const c of s.cells.values()) if (c.v != null) { cells++; if (typeof c.v === 'string' && c.v[0] === '=') formulas++; } dialog({ title: 'Статистика книги', body: `<table class="data"><tr><td>Листов</td><td><b>${book.sheets.length}</b></td></tr><tr><td>Заполненных ячеек</td><td><b>${cells}</b></td></tr><tr><td>Формул</td><td><b>${formulas}</b></td></tr><tr><td>Диаграмм</td><td><b>${book.sheets.reduce((a, s) => a + s.charts.length, 0)}</b></td></tr></table>` }); })] },
  ] },
  { id: 'view', label: 'Вид', groups: [
    { label: 'Показать', items: [big('gridlines', 'grid', 'Сетка', () => { sh().gridlines = sh().gridlines === false; shell.markDirty(); draw(); syncBars(); })] },
    { label: 'Окно', items: [big('freeze', 'freeze', 'Закрепить области', null, { menu: () => [{ label: 'Закрепить области', hint: 'по текущей ячейке', action: () => freeze('cell') }, { label: 'Закрепить верхнюю строку', action: () => freeze('row') }, { label: 'Закрепить первый столбец', action: () => freeze('col') }, '-', { label: 'Снять закрепление', action: () => freeze('none') }] })] },
    { label: 'Масштаб', items: [big('z100', 'search', '100%', () => setZoom(1)), big('zsel', 'maximize', 'По выделению', () => { const g = rng(); const w = colX[g.c2 + 1] - colX[g.c1], hh = rowY[g.r2 + 1] - rowY[g.r1]; setZoom(Math.min((host.clientWidth - HW - 20) / w, (host.clientHeight - HH - 20) / hh)); ensureVisible(g.r1, g.c1); })] },
  ] },
  { id: 'chart', label: 'Диаграмма', contextual: true, groups: [
    { label: 'Тип', items: CHART_TYPES.map((t) => big('ct-' + t.id, t.icon, t.name, () => setChartType(t.id))) },
    { label: 'Данные', items: [big('chedit', 'edit', 'Изменить', () => { const ch = sh().charts.find((c) => c.id === selChart); if (ch) chartDialog(ch); }), big('chdel', 'trash', 'Удалить', deleteChart)] },
  ] },
];
/* extra ribbon controls */
{
  const tab = (id) => ribbon.find((t) => t.id === id);
  const grp = (tid, label) => tab(tid).groups.find((g) => g.label === label);
  grp('home', 'Стили').items.push(
    big('astable', 'table', 'Форматировать как таблицу', null, { menu: () => TABLE_STYLES.map((st) => ({ label: st[0], el: null, icon: 'table', action: () => formatAsTable(st) })) }),
    big('cellstyles', 'palette', 'Стили ячеек', null, { menu: () => CELL_STYLES.map(([n, st]) => ({ label: n, style: st ? `${st.bg ? 'background:' + st.bg + ';' : ''}${st.fc ? 'color:' + st.fc + ';' : ''}${st.b ? 'font-weight:700;' : ''}padding:1px 6px;border-radius:3px` : '', action: () => applyCellStyle(st) })) }));
  tab('insert').groups.unshift({ label: 'Таблицы', items: [big('pivot', 'table', 'Сводная таблица', pivotDialog)] });
  tab('insert').groups.push({ label: 'Ссылки', items: [big('hlink', 'link', 'Ссылка', insertLinkDialog, { key: 'Ctrl+K' })] });
  tab('formulas').groups.push({ label: 'Определённые имена', items: [big('names', 'hash', 'Диспетчер имён', namesDialog), big('defname', 'plus', 'Присвоить имя', defineName)] });
  grp('data', 'Сортировка и фильтр').items.push(big('cfilter', 'filter', 'Фильтр по условию', () => customFilter(sel.ac)));
  grp('data', 'Работа с данными').items.push(big('valid', 'check', 'Проверка данных', validationDialog));
  tab('data').groups.push({ label: 'Прогноз', items: [big('goal', 'search', 'Подбор параметра', goalSeek), big('pivot2', 'table', 'Сводная таблица', pivotDialog)] });
  tab('review').groups.push({ label: 'Защита', items: [big('protect', 'lock', 'Защитить лист', toggleProtect)] });
  ribbon.splice(ribbon.findIndex((t) => t.id === 'formulas'), 0, { id: 'pagelayout', label: 'Разметка страницы', groups: [
    { label: 'Параметры страницы', items: [big('psetup', 'page-size', 'Параметры', pageSetup), big('porient', 'orientation', 'Ориентация', null, { menu: () => [['auto', 'Автоматически'], ['portrait', 'Книжная'], ['landscape', 'Альбомная']].map(([v, l]) => ({ label: l, checked: (sh().print?.orient || 'auto') === v, action: () => { sh().print = { ...(sh().print || {}), orient: v === 'auto' ? null : v }; shell.markDirty(); } })) }),
      big('parea', 'maximize', 'Область печати', null, { menu: () => [{ label: 'Задать по выделению', action: () => { sh().print = { ...(sh().print || {}), area: selRangeForOps() }; shell.markDirty(); toast('Область печати: ' + rangeName(sh().print.area), 'ok'); } }, { label: 'Убрать', disabled: !sh().print?.area, action: () => { delete sh().print.area; shell.markDirty(); } }] })] },
    { label: 'Вписать', items: [big('pfit', 'fit', 'По ширине страницы', () => { sh().print = { ...(sh().print || {}), fit: !sh().print?.fit }; shell.markDirty(); toast(sh().print.fit ? 'Таблица будет вписана по ширине' : 'Обычный масштаб печати'); })] },
    { label: 'Печать', items: [big('pgrid', 'grid', 'Печатать сетку', () => { sh().print = { ...(sh().print || {}), grid: !sh().print?.grid }; shell.markDirty(); }), big('pprint', 'print', 'Печать', () => doPrint(false), { key: 'Ctrl+P' })] },
  ] });
}
let showFormulas = false;
let painterStyle = null;
function formatPainter() { painterStyle = structuredClone(getCell(sel.ar, sel.ac)?.s || {}); shell.ribbon.set('fmtpaint', { on: true }); toast('Выделите ячейки, чтобы применить формат'); }
scroller.addEventListener('pointerup', () => { if (painterStyle && !dragSel) { const st = painterStyle; painterStyle = null; shell.ribbon.set('fmtpaint', { on: false }); forSel((r, c) => { const cell = getCell(r, c); recordCell(K(r, c)); sh().cells.set(K(r, c), { ...(cell || {}), s: Object.keys(st).length ? structuredClone(st) : undefined }); }); } });
function autoRowHeightForFont() { const g = rng(); const fs = getCell(sel.ar, sel.ac)?.s?.fs || 11; const need = Math.ceil(fs * 96 / 72 * 1.45); if (need > DEF_H) { begin(true); for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 500); r++) if ((sh().rowH[r] ?? DEF_H) < need) sh().rowH[r] = need; commit(); rebuildGeom(); } }
async function textToColumns() {
  const g = rng(); const r = await formDlg('Текст по столбцам', [{ key: 'd', label: 'Разделитель', type: 'select', options: [{ value: ',', label: 'Запятая' }, { value: ';', label: 'Точка с запятой' }, { value: '\t', label: 'Табуляция' }, { value: ' ', label: 'Пробел' }] }]); if (!r) return;
  begin(); for (let rr = g.r1; rr <= Math.min(g.r2, g.r1 + 5000); rr++) { const v = getCell(rr, g.c1)?.v; if (typeof v !== 'string') continue; v.split(r.d).forEach((p, i) => inputCell(rr, g.c1 + i, p.trim())); } commit();
}

/* ======================= shell ======================= */
const zoomRange = h('input', { type: 'range', min: 40, max: 300, value: 100, class: 'hide-m' });
zoomRange.oninput = () => setZoom(zoomRange.value / 100);
const shell = createShell({
  app: 'xel', ribbon, defaultTitle: 'Книга', serialize, load, newDoc, templates: TEMPLATES.xel,
  isEmpty: () => book.sheets.length === 1 && !sh().cells.size,
  undo, redo, focusEditor: () => sink.focus({ preventScroll: true }),
  onSearch: () => findDialog(false), print: () => doPrint(false),
  stats: () => [['Листов', book.sheets.length], ['Ячеек с данными', book.sheets.reduce((a, s) => a + s.cells.size, 0)]],
  exporters: [
    { label: 'Книга Excel (.xlsx)', desc: 'Совместима с Microsoft Excel, LibreOffice, Google Таблицами', icon: 'table', run: async () => shell.export(shell.title + '.xlsx', await exportXlsx(book, engine), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') },
    { label: 'CSV (разделитель — запятая)', desc: 'Текущий лист', icon: 'file', run: () => shell.export(`${shell.title} - ${sh().name}.csv`, toCSV(','), 'text/csv') },
    { label: 'CSV (разделитель — точка с запятой)', desc: 'Для русской версии Excel', icon: 'file', run: () => shell.export(`${shell.title} - ${sh().name}.csv`, toCSV(';'), 'text/csv') },
    { label: 'PDF', desc: 'Текущий лист или выделение', icon: 'print', run: () => doPrint(true) },
    { label: 'Веб-страница (.html)', desc: 'Таблица текущего листа', icon: 'globe', run: () => { const u = engine.usedBounds(book.active); shell.export(shell.title + '.html', `<!doctype html><meta charset="utf-8"><title>${esc(shell.title)}</title><style>table{border-collapse:collapse;font:13px Calibri,Arial,sans-serif}td{border:1px solid #ccc;padding:3px 6px}</style>${rangeHTML({ r1: 0, c1: 0, r2: u.mr, c2: u.mc })}`, 'text/html'); } },
  ],
});
shell.workspace.append(wrap);
shell.statusbar.append(h('span', {}, 'Готово'), h('span', { class: 'grow' }), statusSum, h('button', { class: 'sb-btn', html: icon('minus'), onclick: () => setZoom(Z - 0.1) }), zoomRange, h('button', { class: 'sb-btn', html: icon('plus'), onclick: () => setZoom(Z + 0.1) }), zoomLbl);

// "show formulas" mode: return the formula text instead of its value for the active sheet
const _val = engine.value.bind(engine);
engine.value = (si, r, c) => { if (showFormulas && si === book.active) { const raw = book.sheets[si].cells.get(K(r, c))?.v; if (typeof raw === 'string' && raw[0] === '=') return raw; } return _val(si, r, c); };

document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); findDialog(false); }
  else if (mod && e.key.toLowerCase() === 'h') { e.preventDefault(); findDialog(true); }
  else if (e.key === 'F9') { e.preventDefault(); engine.invalidate(); draw(); }
  else if (e.shiftKey && e.key === 'F11') { e.preventDefault(); addSheet(); }
  else if (mod && e.key === 'PageDown') { e.preventDefault(); switchSheet((book.active + 1) % book.sheets.length); }
  else if (mod && e.key === 'PageUp') { e.preventDefault(); switchSheet((book.active - 1 + book.sheets.length) % book.sheets.length); }
  else if ((e.key === 'Delete' || e.key === 'Backspace') && selChart && document.activeElement !== fInput && !editing) { e.preventDefault(); deleteChart(); }
});
new ResizeObserver(() => draw()).observe(host);
window.addEventListener('rabbit-settings', (e) => { if (e.detail.k === 'theme') setTimeout(() => { readColors(); draw(); renderCharts(); }, 50); });
setInterval(() => { if (book.sheets.length) sh().lastSel = { ...sel }; }, 500);
readColors();
book.sheets = [newSheet('Лист1')];
rebuildGeom(); renderTabs(); draw();
shell.boot().then(() => { sink.focus({ preventScroll: true }); syncBars(); });
