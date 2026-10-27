// Rabbit Documents — word processor
import { createShell } from '../core/shell.js';
import { icon } from '../core/icons.js';
import { h, $, $$, esc, toast, dialog, menu, popover, closeMenus, colorPicker, promptDlg, confirmDlg, formDlg, pickImage, debounce, uid, clamp } from '../core/ui.js';
import { settings, profile } from '../core/store.js';
import { exportDocx, importDocx } from './docx.js';
import { TEMPLATES } from '../core/templates.js';
import { extOf, printToPdf, native } from '../core/platform.js';
import { renderChart, CHART_TYPES, chartDataFromGrid } from '../core/charts.js';
import { SHAPES } from '../point/render.js';

const MM = 96 / 25.4;
const PAGE_SIZES = { A4: [210, 297], A5: [148, 210], A3: [297, 420], Letter: [215.9, 279.4], Legal: [215.9, 355.6], B5: [176, 250] };
const MARGINS = { normal: ['Обычные', [25.4, 25.4, 25.4, 25.4]], narrow: ['Узкие', [12.7, 12.7, 12.7, 12.7]], moderate: ['Средние', [25.4, 19.1, 25.4, 19.1]], wide: ['Широкие', [25.4, 50.8, 25.4, 50.8]], gost: ['ГОСТ (30/15/20/20)', [20, 15, 20, 30]] };
const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Georgia', 'Cambria', 'Segoe UI', 'Roboto', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Garamond', 'Courier New', 'Consolas', 'PT Sans', 'PT Serif', 'Inter'];
const PT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const DOC_STYLES = {
  modern: { name: 'Современный', font: 'Calibri', head: 'Calibri', size: 11, lh: 1.5 },
  gost: { name: 'ГОСТ / деловой', font: 'Times New Roman', head: 'Times New Roman', size: 14, lh: 1.5 },
  elegant: { name: 'Элегантный', font: 'Georgia', head: 'Garamond', size: 12, lh: 1.6 },
  tech: { name: 'Технический', font: 'Segoe UI', head: 'Segoe UI', size: 10.5, lh: 1.45 },
  compact: { name: 'Компактный', font: 'Arial', head: 'Arial', size: 10, lh: 1.3 },
};
const BLOCK_STYLES = [
  { id: 'p', name: 'Обычный', tag: 'p', css: 'font-size:13px' },
  { id: 'title', name: 'Название', tag: 'h1', cls: 'doc-title', css: 'font-size:20px;font-weight:300;color:#14213d' },
  { id: 'subtitle', name: 'Подзаголовок', tag: 'h4', cls: 'doc-subtitle', css: 'font-size:13px;color:#5f6b85' },
  { id: 'h1', name: 'Заголовок 1', tag: 'h1', css: 'font-size:17px;font-weight:600;color:#1f3d7a' },
  { id: 'h2', name: 'Заголовок 2', tag: 'h2', css: 'font-size:15px;font-weight:600;color:#1f3d7a' },
  { id: 'h3', name: 'Заголовок 3', tag: 'h3', css: 'font-size:13px;font-weight:600;color:#2a3550' },
  { id: 'quote', name: 'Цитата', tag: 'blockquote', css: 'font-style:italic;color:#4a5266;border-left:3px solid #9fb3d9;padding-left:5px' },
  { id: 'code', name: 'Код', tag: 'pre', css: 'font-family:Consolas,monospace;font-size:11px;background:#f4f5f7' },
];

let doc = newDocState();
function newDocState() {
  return { page: { size: 'A4', w: 210, h: 297, landscape: false, margins: [...MARGINS.normal[1]], marginsPreset: 'normal', columns: 1, color: null },
    header: { text: '', align: 'center', show: false }, footer: { text: '', align: 'center', pageNumbers: false, show: false },
    style: { preset: 'modern', ...DOC_STYLES.modern }, comments: [], track: false, watermark: '', readonly: false, hyphens: false };
}

/* ======================= DOM ======================= */
const scroll = h('div', { class: 'doc-scroll' });
const stage = h('div', { class: 'doc-stage' });
const page = h('div', { class: 'page' });
const hfHead = h('div', { class: 'hf hf-head hidden-hf', contenteditable: 'plaintext-only', 'data-ph': 'Верхний колонтитул' });
const hfFoot = h('div', { class: 'hf hf-foot hidden-hf', contenteditable: 'plaintext-only', 'data-ph': 'Нижний колонтитул' });
const editor = h('div', { class: 'editor', contenteditable: 'true', spellcheck: String(settings.get('spellcheck')), role: 'textbox', 'aria-multiline': 'true', 'aria-label': 'Документ' });
const marks = h('div', { class: 'page-marks' });
const wmEl = h('div', { class: 'watermark hidden', 'aria-hidden': 'true' });
const roBanner = h('div', { class: 'ro-banner hidden no-print' }, icon('lock'), h('span', { class: 'grow' }, 'Документ защищён от изменений'), h('button', { class: 'btn sm', onclick: () => setReadonly(false) }, 'Разрешить редактирование'));
page.append(wmEl, hfHead, editor, hfFoot, marks);
stage.append(page); scroll.append(stage);
const navPane = h('aside', { class: 'panel nav-pane hidden' }, h('div', { class: 'panel-head' }, 'Навигация', h('span', { class: 'grow' }), h('button', { class: 'btn ghost icon sm', html: icon('close'), onclick: () => togglePane('nav') })), h('div', { class: 'panel-body', id: 'nav-list' }));
const cmtPane = h('aside', { class: 'panel cmt-pane hidden' }, h('div', { class: 'panel-head' }, 'Примечания', h('span', { class: 'grow' }), h('button', { class: 'btn ghost icon sm', html: icon('close'), onclick: () => togglePane('cmt') })), h('div', { class: 'panel-body', id: 'cmt-list' }));

/* ======================= selection helpers ======================= */
let savedRange = null;
document.addEventListener('selectionchange', () => {
  const s = getSelection();
  if (s.rangeCount && editor.contains(s.anchorNode)) { savedRange = s.getRangeAt(0).cloneRange(); updateState(); }
});
function restoreSel() {
  if (document.activeElement !== editor) editor.focus({ preventScroll: true });
  if (savedRange) { const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange); }
}
function cmd(name, val = null, css = true) {
  restoreSel();
  document.execCommand('styleWithCSS', false, css);
  document.execCommand(name, false, val);
  changed(); updateState();
}
function anchorEl() { const n = savedRange?.startContainer; if (!n) return null; return n.nodeType === 1 ? n : n.parentElement; }
const BLOCK_SEL = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,pre,div:not(.toc):not(.page-break):not(.editor)';
function selectedBlocks() {
  if (!savedRange) return [];
  const all = [...editor.querySelectorAll(BLOCK_SEL)].filter((el) => savedRange.intersectsNode(el));
  const inner = all.filter((el) => !all.some((o) => o !== el && el.contains(o)));
  if (!inner.length) { const a = anchorEl()?.closest(BLOCK_SEL); if (a && editor.contains(a)) return [a]; }
  return inner;
}
function ensureBlock() {
  if (!editor.firstChild || [...editor.childNodes].some((n) => n.nodeType === 3 && n.data.trim())) { restoreSel(); document.execCommand('formatBlock', false, 'p'); }
}

/* ======================= formatting ======================= */
let pendingPt = 11;
function applyFontSize(pt) {
  pt = clamp(parseFloat(String(pt).replace(',', '.')) || 11, 1, 400);
  pendingPt = pt;
  restoreSel();
  document.execCommand('styleWithCSS', false, true);
  document.execCommand('fontSize', false, '7');
  fixFontSizes();
  changed(); updateState();
}
function fixFontSizes() {
  editor.querySelectorAll('span[style*="xxx-large"], font[size="7"]').forEach((el) => {
    if (el.tagName === 'FONT') { const s = document.createElement('span'); s.style.fontSize = pendingPt + 'pt'; s.append(...el.childNodes); el.replaceWith(s); }
    else el.style.fontSize = pendingPt + 'pt';
  });
}
function growFont(d) {
  const cur = currentPt(); const i = PT_SIZES.findIndex((s) => s >= cur);
  const next = d > 0 ? PT_SIZES[Math.min(PT_SIZES.length - 1, (PT_SIZES[i] === cur ? i + 1 : i))] : PT_SIZES[Math.max(0, (i < 0 ? PT_SIZES.length : i) - 1)];
  applyFontSize(next || cur + d);
}
function currentPt() { const a = anchorEl(); if (!a) return doc.style.size; return Math.round(parseFloat(getComputedStyle(a).fontSize) * 0.75 * 2) / 2; }
function setBlockStyle(id) {
  const st = BLOCK_STYLES.find((s) => s.id === id); if (!st) return;
  restoreSel();
  document.execCommand('formatBlock', false, st.tag);
  const blocks = selectedBlocks();
  for (const b of blocks) { if (b.tagName.toLowerCase() === st.tag) { b.className = st.cls || ''; if (!b.className) b.removeAttribute('class'); } }
  changed(); updateState(); refreshNav();
}
function setLineHeight(v) { for (const b of selectedBlocks()) b.style.lineHeight = v; changed(); }
function setSpacing(before, after) { for (const b of selectedBlocks()) { if (before != null) b.style.marginTop = before + 'pt'; if (after != null) b.style.marginBottom = after + 'pt'; } changed(); }
function changeCase(mode) {
  restoreSel(); const s = getSelection(); if (!s.rangeCount || s.isCollapsed) return toast('Выделите текст');
  const r = s.getRangeAt(0);
  const walker = document.createTreeWalker(r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer, NodeFilter.SHOW_TEXT);
  let first = true;
  const tf = (t) => mode === 'upper' ? t.toLocaleUpperCase('ru') : mode === 'lower' ? t.toLocaleLowerCase('ru') : mode === 'title' ? t.toLocaleLowerCase('ru').replace(/(^|[\s\-«"(])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('ru')) :
    t.toLocaleLowerCase('ru').replace(/(^|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('ru')).replace(/^(\s*)(\p{L})/u, (m, a, b) => (first ? a + b.toLocaleUpperCase('ru') : m));
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!r.intersectsNode(n)) continue;
    const st = n === r.startContainer ? r.startOffset : 0; const en = n === r.endContainer ? r.endOffset : n.data.length;
    n.data = n.data.slice(0, st) + tf(n.data.slice(st, en)) + n.data.slice(en); first = false;
  }
  changed();
}
let painter = null;
function formatPainter() {
  const a = anchorEl(); if (!a) return;
  const cs = getComputedStyle(a);
  painter = { b: +cs.fontWeight >= 600, i: cs.fontStyle === 'italic', u: cs.textDecorationLine.includes('underline'), color: cs.color, font: cs.fontFamily.split(',')[0].replace(/["']/g, ''), pt: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2 };
  shell.ribbon.set('painter', { on: true }); editor.style.cursor = 'copy';
  toast('Выделите текст, чтобы применить формат');
}
editor.addEventListener('mouseup', () => {
  if (!painter) return;
  setTimeout(() => {
    const s = getSelection(); if (s.isCollapsed) return;
    const p = painter; painter = null; editor.style.cursor = ''; shell.ribbon.set('painter', { on: false });
    cmd('removeFormat');
    if (p.b) cmd('bold'); if (p.i) cmd('italic'); if (p.u) cmd('underline');
    cmd('foreColor', p.color); cmd('fontName', p.font); applyFontSize(p.pt);
  }, 0);
});

/* ======================= insert ======================= */
function insertHTML(html) { restoreSel(); document.execCommand('insertHTML', false, html); changed(); }
function insertTable(rows, cols) {
  const w = (100 / cols).toFixed(2);
  let html = '<table class="doc-table header-row"><tbody>';
  for (let r = 0; r < rows; r++) { html += '<tr>'; for (let c = 0; c < cols; c++) html += `<td style="width:${w}%"><br></td>`; html += '</tr>'; }
  insertHTML(html + '</tbody></table><p><br></p>');
}
function tablePicker(anchor) {
  const grid = h('div', { class: 'gridpick' }); const lbl = h('div', { class: 'small muted', style: { margin: '6px 0' } }, 'Вставка таблицы');
  const cells = [];
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const i = h('i'); i.onmouseenter = () => { cells.forEach((x) => x.el.classList.toggle('on', x.r <= r && x.c <= c)); lbl.textContent = `Таблица ${c + 1}×${r + 1}`; };
    i.onclick = () => { closeMenus(); insertTable(r + 1, c + 1); }; cells.push({ el: i, r, c }); grid.append(i);
  }
  const more = h('button', { class: 'btn sm ghost', style: { marginTop: '6px', width: '100%' }, onclick: async () => { closeMenus(); const r = await formDlg('Вставить таблицу', [{ key: 'cols', label: 'Число столбцов', type: 'number', value: 3, min: 1, max: 30 }, { key: 'rows', label: 'Число строк', type: 'number', value: 3, min: 1, max: 500 }]); if (r) insertTable(clamp(r.rows, 1, 500), clamp(r.cols, 1, 30)); } }, 'Вставить таблицу…');
  popover(anchor, h('div', {}, lbl, grid, more));
}
async function insertImage() {
  const src = await pickImage(); if (!src) return;
  const img = new Image(); img.src = src; await img.decode().catch(() => {});
  const maxW = editor.clientWidth || 600; const w = Math.min(img.naturalWidth || maxW, maxW);
  insertHTML(`<img src="${src}" style="width:${Math.round(w)}px" alt="">`);
}
async function insertLink() {
  const s = getSelection(); const sel = savedRange && !savedRange.collapsed ? savedRange.toString() : '';
  const existing = anchorEl()?.closest('a');
  const r = await formDlg('Гиперссылка', [{ key: 'text', label: 'Текст', value: existing?.textContent || sel }, { key: 'url', label: 'Адрес', value: existing?.getAttribute('href') || 'https://', placeholder: 'https://example.com' }], 'Вставить');
  if (!r || !r.url) return;
  if (existing) { existing.href = r.url; existing.textContent = r.text || r.url; changed(); return; }
  if (sel && sel === r.text) cmd('createLink', r.url);
  else insertHTML(`<a href="${esc(r.url)}">${esc(r.text || r.url)}</a>&nbsp;`);
}
function insertTOC() {
  const html = `<div class="toc" contenteditable="false">${tocInner()}</div><p><br></p>`;
  insertHTML(html);
}
function tocInner() {
  const hs = [...editor.querySelectorAll('h1:not(.doc-title),h2,h3')].filter((x) => !x.closest('.toc'));
  hs.forEach((x) => { if (!x.id) x.id = uid('h'); });
  return `<div class="toc-h">Оглавление <button data-act="toc-update">Обновить</button></div>` + (hs.length ? hs.map((x) => `<a href="#${x.id}" data-level="${x.tagName[1]}" style="padding-left:${(x.tagName[1] - 1) * 18}px">${esc(x.textContent)}</a>`).join('') : '<div style="color:#999">Добавьте заголовки (стили «Заголовок 1–3»), затем нажмите «Обновить»</div>');
}
const SYMBOLS = '© ® ™ § ¶ ° ± × ÷ ≠ ≈ ≤ ≥ ∞ √ ∑ ∏ ∫ ∂ ∆ π Ω α β γ δ ε λ μ σ φ ψ € ₽ $ £ ¥ ¢ … — – « » „ “ ” ‘ ’ • · ← → ↑ ↓ ↔ ⇒ ⇔ ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠ ½ ⅓ ¼ ¾ ² ³ ¹ ‰ № ∈ ∉ ∩ ∪ ⊂ ⊃ ∀ ∃ ¬ ∧ ∨'.split(' ');
function symbolPicker(anchor) {
  const g = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(10, 28px)', gap: '2px' } });
  for (const s of SYMBOLS) g.append(h('button', { class: 'btn-r', style: { justifyContent: 'center', fontSize: '15px', width: '28px' }, onclick: () => { closeMenus(); insertHTML(s); } }, s));
  popover(anchor, g);
}
async function insertChart() {
  const ta = h('textarea', { class: 'input mono', rows: 7 }); ta.value = 'Квартал;Продажи;Расходы\nI;120;80\nII;150;95\nIII;170;100\nIV;210;120';
  const type = h('select', { class: 'input' }, ...CHART_TYPES.map((t) => h('option', { value: t.id }, t.name)));
  const title = h('input', { class: 'input', value: 'Диаграмма' });
  const prev = h('div', { style: { height: '200px', border: '1px solid var(--border)', borderRadius: '6px', background: '#fff' } });
  const parse = () => chartDataFromGrid(ta.value.trim().split('\n').map((l) => l.split(/[;\t,]/).map((v) => (isFinite(+v) && v.trim() !== '' ? +v : v.trim()))));
  const draw = () => { prev.innerHTML = renderChart({ type: type.value, title: title.value, ...parse() }, 480, 200); };
  [ta, type, title].forEach((e) => e.addEventListener('input', draw)); type.onchange = draw;
  const body = h('div', { style: { width: 'min(560px,100%)' } }, h('div', { class: 'row' }, h('div', { class: 'field grow' }, h('label', {}, 'Тип'), type), h('div', { class: 'field grow' }, h('label', {}, 'Заголовок'), title)), h('div', { class: 'field' }, h('label', {}, 'Данные (через ; или табуляцию, первая строка — заголовки)'), ta), prev);
  setTimeout(draw, 20);
  const ok = await dialog({ title: 'Вставить диаграмму', body, buttons: [{ label: 'Отмена', value: false }, { label: 'Вставить', value: true, primary: true }] });
  if (!ok) return;
  const svg = renderChart({ type: type.value, title: title.value, ...parse() }, 600, 340, { bg: '#ffffff' });
  insertHTML(`<img src="data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}" style="width:${Math.min(600, editor.clientWidth)}px" alt="${esc(title.value)}">`);
}

/* ======================= tables ======================= */
function curCell() { return anchorEl()?.closest('td,th'); }
function tableOp(op) {
  const cell = curCell(); if (!cell) return;
  const row = cell.parentElement; const table = cell.closest('table'); const idx = cell.cellIndex;
  const emptyRow = () => { const tr = document.createElement('tr'); for (const c of row.cells) { const td = document.createElement('td'); td.colSpan = c.colSpan; td.style.width = c.style.width; td.innerHTML = '<br>'; tr.append(td); } return tr; };
  if (op === 'rowAbove') row.before(emptyRow());
  if (op === 'rowBelow') row.after(emptyRow());
  if (op === 'colLeft' || op === 'colRight') for (const r of table.rows) { const td = document.createElement('td'); td.innerHTML = '<br>'; const ref = r.cells[Math.min(idx, r.cells.length - 1)]; op === 'colLeft' ? ref.before(td) : ref.after(td); }
  if (op === 'delRow') { if (table.rows.length === 1) table.remove(); else row.remove(); }
  if (op === 'delCol') { if (row.cells.length === 1) table.remove(); else for (const r of [...table.rows]) r.cells[Math.min(idx, r.cells.length - 1)]?.remove(); }
  if (op === 'delTable') table.remove();
  if (op === 'merge') { const next = cell.nextElementSibling; if (next) { cell.colSpan = (cell.colSpan || 1) + (next.colSpan || 1); if (next.textContent.trim()) cell.append(' ', ...next.childNodes); next.remove(); } }
  if (op === 'split' && cell.colSpan > 1) { const n = cell.colSpan; cell.colSpan = 1; for (let i = 1; i < n; i++) { const td = document.createElement('td'); td.innerHTML = '<br>'; cell.after(td); } }
  if (op === 'borders') table.classList.toggle('no-borders');
  if (op === 'header') table.classList.toggle('header-row');
  if (op === 'banded') table.classList.toggle('banded');
  if (op === 'even') { const cols = Math.max(...[...table.rows].map((r) => r.cells.length)); for (const r of table.rows) for (const c of r.cells) c.style.width = ((100 / cols) * (c.colSpan || 1)).toFixed(2) + '%'; }
  changed(); updateState();
}
function cellShade(color) { const cell = curCell(); if (!cell) return; cell.style.backgroundColor = color || ''; changed(); }
editor.addEventListener('keydown', (e) => {
  if (e.key === 'Tab') {
    const cell = curCell();
    if (cell) {
      e.preventDefault();
      const cells = [...cell.closest('table').querySelectorAll('td,th')]; let i = cells.indexOf(cell) + (e.shiftKey ? -1 : 1);
      if (i >= cells.length) { tableOp('rowBelow'); i = cells.length; }
      const target = [...cell.closest('table').querySelectorAll('td,th')][Math.max(0, i)];
      const r = document.createRange(); r.selectNodeContents(target); r.collapse(false); const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      return;
    }
    if (anchorEl()?.closest('li')) { e.preventDefault(); cmd(e.shiftKey ? 'outdent' : 'indent'); return; }
    e.preventDefault(); insertHTML('&emsp;');
  }
});

/* ======================= images ======================= */
let selImg = null; let handle = null;
function selectImage(img) {
  selImg?.classList.remove('sel'); handle?.remove(); handle = null; selImg = img;
  shell.ribbon.showTab('picture', !!img, true);
  if (!img) return;
  img.classList.add('sel');
  handle = h('div', { class: 'img-handle' }); page.append(handle); placeHandle();
  handle.addEventListener('pointerdown', (e) => {
    e.preventDefault(); handle.setPointerCapture(e.pointerId);
    const sx = e.clientX, sw = img.offsetWidth; const z = zoom;
    const move = (ev) => { img.style.width = clamp(Math.round(sw + (ev.clientX - sx) / z), 24, editor.clientWidth) + 'px'; img.style.height = 'auto'; placeHandle(); };
    const up = () => { handle.removeEventListener('pointermove', move); handle.removeEventListener('pointerup', up); changed(); };
    handle.addEventListener('pointermove', move); handle.addEventListener('pointerup', up);
  });
}
function placeHandle() { if (!handle || !selImg) return; const pr = page.getBoundingClientRect(), r = selImg.getBoundingClientRect(); handle.style.left = (r.right - pr.left) / zoom - 6 + 'px'; handle.style.top = (r.bottom - pr.top) / zoom - 6 + 'px'; }
editor.addEventListener('click', (e) => {
  if (e.target.tagName === 'IMG') { selectImage(e.target); const r = document.createRange(); r.selectNode(e.target); getSelection().removeAllRanges(); getSelection().addRange(r); }
  else if (selImg) selectImage(null);
  const a = e.target.closest('a'); if (a && (e.ctrlKey || e.metaKey) && !a.closest('.toc')) { e.preventDefault(); window.open(a.href, '_blank'); }
  const tocA = e.target.closest('.toc a'); if (tocA) { e.preventDefault(); document.getElementById(tocA.getAttribute('href').slice(1))?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
  if (e.target.dataset?.act === 'toc-update') { e.target.closest('.toc').innerHTML = tocInner(); changed(); }
  const li = e.target.closest('ul.checklist > li'); if (li && e.offsetX < 18) { li.classList.toggle('done'); changed(); }
  const cm = e.target.closest('.cmt-anchor'); if (cm) focusComment(cm.dataset.cid);
  const fr = e.target.closest('sup.fn-ref'); if (fr) editor.querySelector(`p.fn[data-fn="${fr.dataset.fn}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  const fnum = e.target.closest('.fn-num'); if (fnum) editor.querySelector(`sup.fn-ref[data-fn="${fnum.parentElement.dataset.fn}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
});
function imgWrap(mode) { if (!selImg) return; selImg.classList.remove('fl', 'fr', 'ctr'); if (mode) selImg.classList.add(mode); changed(); placeHandle(); }
function imgWidth(pct) { if (!selImg) return; selImg.style.width = Math.round(editor.clientWidth * pct) + 'px'; changed(); placeHandle(); }

/* ======================= comments ======================= */
async function addComment() {
  restoreSel(); const s = getSelection();
  if (!s.rangeCount || s.isCollapsed) return toast('Выделите текст, к которому относится примечание');
  const r = s.getRangeAt(0).cloneRange();
  const startBlock = (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement).closest(BLOCK_SEL) || editor;
  if (!startBlock.contains(r.endContainer)) r.setEnd(startBlock, startBlock.childNodes.length);
  const text = await promptDlg('Новое примечание', 'Текст примечания', '');
  if (!text) return;
  const id = uid('c');
  const span = document.createElement('span'); span.className = 'cmt-anchor'; span.dataset.cid = id;
  span.append(r.extractContents()); r.insertNode(span);
  const p = profile();
  doc.comments.push({ id, author: p.name, color: p.color, time: Date.now(), text, quote: span.textContent.slice(0, 80), replies: [] });
  changed(); showPane('cmt', true); renderComments(); focusComment(id);
}
function renderComments() {
  const list = $('#cmt-list'); list.innerHTML = '';
  const order = [...editor.querySelectorAll('.cmt-anchor')].map((a) => a.dataset.cid);
  doc.comments = doc.comments.filter((c) => order.includes(c.id));
  if (!doc.comments.length) { list.append(h('div', { class: 'empty' }, 'Примечаний нет. Выделите текст и нажмите «Создать примечание».')); return; }
  doc.comments.sort((a, b) => order.indexOf(a.id) - order.indexOf(b.id));
  for (const c of doc.comments) {
    const card = h('div', { class: 'cmt-card', 'data-cid': c.id });
    const who = (n, col, t) => `<div class="who"><span class="dot" style="background:${col}">${esc(n[0] || '?')}</span>${esc(n)}<span class="when">${new Date(t).toLocaleString('ru-RU', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</span></div>`;
    card.innerHTML = who(c.author, c.color, c.time) + `<div class="quote">${esc(c.quote)}</div><div class="txt">${esc(c.text)}</div>` + c.replies.map((r) => `<div style="margin-left:12px;margin-top:6px">${who(r.author, r.color, r.time)}<div class="txt">${esc(r.text)}</div></div>`).join('');
    const acts = h('div', { class: 'acts' },
      h('button', { class: 'btn sm ghost', onclick: async () => { const t = await promptDlg('Ответ', '', ''); if (t) { const p = profile(); c.replies.push({ author: p.name, color: p.color, time: Date.now(), text: t }); changed(); renderComments(); } } }, 'Ответить'),
      h('button', { class: 'btn sm ghost', onclick: () => deleteComment(c.id) }, 'Решено'));
    card.append(acts);
    card.onclick = () => focusComment(c.id, false);
    list.append(card);
  }
}
function focusComment(id, scrollDoc = true) {
  showPane('cmt', true);
  $$('.cmt-anchor').forEach((a) => a.classList.toggle('active', a.dataset.cid === id));
  $$('.cmt-card').forEach((c) => c.classList.toggle('active', c.dataset.cid === id));
  if (scrollDoc) editor.querySelector(`.cmt-anchor[data-cid="${id}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  $(`.cmt-card[data-cid="${id}"]`)?.scrollIntoView({ block: 'nearest' });
}
function deleteComment(id) {
  const a = editor.querySelector(`.cmt-anchor[data-cid="${id}"]`);
  if (a) a.replaceWith(...a.childNodes);
  doc.comments = doc.comments.filter((c) => c.id !== id); changed(); renderComments();
}
function stepComment(d) {
  const as = [...editor.querySelectorAll('.cmt-anchor')]; if (!as.length) return toast('Примечаний нет');
  const cur = as.findIndex((a) => a.classList.contains('active')); const n = as[(cur + d + as.length) % as.length]; focusComment(n.dataset.cid);
}

/* ======================= find & replace ======================= */
let findBar = null; let matches = []; let mIdx = -1;
function openFind(replace = false) {
  if (findBar) { findBar.remove(); findBar = null; }
  const q = h('input', { class: 'input', placeholder: 'Найти в документе', value: savedRange && !savedRange.collapsed ? savedRange.toString().slice(0, 100) : '' });
  const rep = h('input', { class: 'input', placeholder: 'Заменить на' });
  const cnt = h('span', { class: 'cnt' });
  const cs = h('input', { type: 'checkbox' }); const ww = h('input', { type: 'checkbox' });
  const run = () => { search(q.value, cs.checked, ww.checked); cnt.textContent = matches.length ? `${mIdx + 1} из ${matches.length}` : q.value ? 'Нет' : ''; };
  const go = (d) => { if (!matches.length) return; mIdx = (mIdx + d + matches.length) % matches.length; showMatch(); cnt.textContent = `${mIdx + 1} из ${matches.length}`; };
  q.addEventListener('input', run); cs.onchange = run; ww.onchange = run;
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); go(e.shiftKey ? -1 : 1); } if (e.key === 'Escape') closeFind(); });
  rep.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeFind(); });
  const repRow = h('div', { class: 'row' + (replace ? '' : ' hidden') }, rep,
    h('button', { class: 'btn sm', onclick: () => { if (mIdx < 0) return; const r = matches[mIdx]; r.deleteContents(); r.insertNode(document.createTextNode(rep.value)); editor.normalize(); changed(); run(); } }, 'Заменить'),
    h('button', { class: 'btn sm', onclick: () => { const n = matches.length; for (let i = matches.length - 1; i >= 0; i--) { matches[i].deleteContents(); matches[i].insertNode(document.createTextNode(rep.value)); } editor.normalize(); changed(); run(); toast(`Заменено: ${n}`, 'ok'); } }, 'Все'));
  findBar = h('div', { class: 'find-bar' },
    h('div', { class: 'row' }, q, cnt, h('button', { class: 'btn ghost icon sm', html: icon('chevu'), onclick: () => go(-1) }), h('button', { class: 'btn ghost icon sm', html: icon('chevd'), onclick: () => go(1) }),
      h('button', { class: 'btn ghost icon sm', 'data-tip': 'Замена', html: icon('replace'), onclick: () => repRow.classList.toggle('hidden') }), h('button', { class: 'btn ghost icon sm', html: icon('close'), onclick: closeFind })),
    repRow, h('div', { class: 'row small' }, h('label', { class: 'check' }, cs, 'Учитывать регистр'), h('label', { class: 'check' }, ww, 'Слово целиком')));
  shell.workspace.append(findBar); q.focus(); q.select(); if (q.value) run();
}
function closeFind() { findBar?.remove(); findBar = null; matches = []; if (window.CSS?.highlights) { CSS.highlights.delete('find'); CSS.highlights.delete('find-cur'); } }
function search(q, caseSens, whole) {
  matches = []; mIdx = -1;
  if (q) {
    const esc_ = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(whole ? `(?<![\\p{L}\\p{N}])${esc_}(?![\\p{L}\\p{N}])` : esc_, caseSens ? 'gu' : 'giu');
    const w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) { re.lastIndex = 0; let m; while ((m = re.exec(n.data))) { const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + m[0].length); matches.push(r); if (!m[0].length) re.lastIndex++; } }
  }
  if (window.CSS?.highlights) { CSS.highlights.set('find', new Highlight(...matches)); CSS.highlights.delete('find-cur'); }
  if (matches.length) { mIdx = 0; showMatch(); }
}
function showMatch() {
  const r = matches[mIdx]; if (!r) return;
  if (window.CSS?.highlights) CSS.highlights.set('find-cur', new Highlight(r));
  const rect = r.getBoundingClientRect(); const sr = scroll.getBoundingClientRect();
  if (rect.top < sr.top + 40 || rect.bottom > sr.bottom - 40) scroll.scrollTop += rect.top - sr.top - sr.height / 3;
}

/* ======================= page layout ======================= */
let zoom = 1;
function applyPage() {
  const p = doc.page; const w = (p.landscape ? p.h : p.w) * MM, hh = (p.landscape ? p.w : p.h) * MM;
  const [mt, mr, mb, ml] = p.margins.map((m) => m * MM);
  Object.entries({ '--pw': w + 'px', '--ph': hh + 'px', '--mt': mt + 'px', '--mr': mr + 'px', '--mb': mb + 'px', '--ml': ml + 'px' }).forEach(([k, v]) => page.style.setProperty(k, v));
  page.style.setProperty('--page-bg', p.color || '#ffffff');
  editor.style.columnCount = p.columns > 1 ? p.columns : '';
  const st = doc.style;
  page.style.setProperty('--doc-font', `"${st.font}", ${st.font === 'Calibri' ? 'Carlito, ' : ''}system-ui, sans-serif`);
  page.style.setProperty('--doc-head-font', `"${st.head || st.font}", system-ui, sans-serif`);
  page.style.setProperty('--doc-size', st.size + 'pt');
  editor.style.lineHeight = st.lh || 1.5;
  hfHead.textContent = doc.header.text; hfFoot.textContent = doc.footer.text;
  hfHead.style.textAlign = doc.header.align; hfFoot.style.textAlign = doc.footer.align;
  hfHead.classList.toggle('hidden-hf', !(doc.header.show || doc.header.text));
  hfFoot.classList.toggle('hidden-hf', !(doc.footer.show || doc.footer.text || doc.footer.pageNumbers));
  hfFoot.dataset.ph = doc.footer.pageNumbers ? 'Нижний колонтитул · номер страницы' : 'Нижний колонтитул';
  wmEl.textContent = doc.watermark || ''; wmEl.classList.toggle('hidden', !doc.watermark);
  editor.style.hyphens = doc.hyphens ? 'auto' : '';
  editor.contentEditable = doc.readonly ? 'false' : 'true'; roBanner.classList.toggle('hidden', !doc.readonly);
  editor.classList.toggle('tracking', !!doc.track);
  if (shell) { shell.ribbon.set('track', { on: !!doc.track }); shell.ribbon.set('hyph', { on: !!doc.hyphens }); shell.ribbon.set('ro', { on: !!doc.readonly }); }
  updatePrintStyle();
  layoutMarks();
}
function updatePrintStyle() {
  const p = doc.page; const sz = `${p.landscape ? p.h : p.w}mm ${p.landscape ? p.w : p.h}mm`;
  const q = (s) => '"' + String(s).replace(/["\\]/g, '\\$&').replace(/\n/g, ' ') + '"';
  const box = (pos, txt) => `@${pos} { content: ${txt}; font: 9pt ${doc.style.font}, sans-serif; color: #666; }`;
  const hp = { left: 'top-left', center: 'top-center', right: 'top-right' }[doc.header.align];
  const fp = { left: 'bottom-left', center: 'bottom-center', right: 'bottom-right' }[doc.footer.align];
  const ft = [doc.footer.text ? q(doc.footer.text + (doc.footer.pageNumbers ? '   ' : '')) : '', doc.footer.pageNumbers ? 'counter(page)' : ''].filter(Boolean).join(' ');
  $('#print-style').textContent = `@page { size: ${sz}; margin: ${p.margins.map((m) => m + 'mm').join(' ')}; ${doc.header.text ? box(hp, q(doc.header.text)) : ''} ${ft ? box(fp, ft) : ''} }`;
}
function layoutMarks() {
  const p = doc.page; const ph = (p.landscape ? p.w : p.h) * MM; const [mt, , mb] = p.margins.map((m) => m * MM);
  page.style.minHeight = '';
  const contentH = editor.scrollHeight; const per = ph - mt - mb;
  const n = Math.max(1, Math.ceil((contentH - 2) / per));
  page.style.minHeight = n * ph + 'px';
  marks.innerHTML = '';
  for (let i = 1; i < n; i++) marks.append(h('div', { class: 'page-mark', style: { top: i * ph + 'px' } }, h('span', {}, `Стр. ${i + 1}`)));
  pageCount = n; updateStatus();
}
let pageCount = 1;
function setZoom(z) { zoom = clamp(Math.round(z * 100) / 100, 0.3, 3); stage.style.zoom = zoom; zoomLbl.textContent = Math.round(zoom * 100) + '%'; zoomRange.value = Math.round(zoom * 100); placeHandle(); }
function fitWidth() { const avail = scroll.clientWidth - 48; setZoom(avail / page.offsetWidth); }

/* ======================= panes ======================= */
function showPane(which, on) { (which === 'nav' ? navPane : cmtPane).classList.toggle('hidden', !on); shell.ribbon.set(which === 'nav' ? 'navpane' : 'cmtpane', { on }); if (which === 'nav' && on) refreshNav(); if (which === 'cmt' && on) renderComments(); }
function togglePane(which) { showPane(which, (which === 'nav' ? navPane : cmtPane).classList.contains('hidden')); }
const refreshNav = debounce(() => {
  if (navPane.classList.contains('hidden')) return;
  const list = $('#nav-list'); list.innerHTML = '';
  const hs = [...editor.querySelectorAll('h1,h2,h3')].filter((x) => !x.closest('.toc') && x.textContent.trim());
  if (!hs.length) list.append(h('div', { class: 'empty' }, 'Используйте стили заголовков, чтобы здесь появилась структура документа'));
  for (const x of hs) list.append(h('button', { class: `h-item l${x.classList.contains('doc-title') ? 0 : x.tagName[1]}`, onclick: () => x.scrollIntoView({ behavior: 'smooth', block: 'start' }) }, x.textContent));
}, 250);

/* ======================= state & status ======================= */
function updateState() {
  if (!shell) return;
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  for (const [id, c] of [['bold', 'bold'], ['italic', 'italic'], ['underline', 'underline'], ['strike', 'strikeThrough'], ['sub', 'subscript'], ['sup', 'superscript'], ['ul', 'insertUnorderedList'], ['ol', 'insertOrderedList'], ['jl', 'justifyLeft'], ['jc', 'justifyCenter'], ['jr', 'justifyRight'], ['jf', 'justifyFull']]) shell.ribbon.set(id, { on: q(c) });
  const a = anchorEl(); if (!a) return;
  const cs = getComputedStyle(a);
  shell.ribbon.set('font', { value: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim() });
  shell.ribbon.set('size', { value: String(Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2) });
  const blk = a.closest(BLOCK_SEL);
  const sid = blk ? (blk.classList.contains('doc-title') ? 'title' : blk.classList.contains('doc-subtitle') ? 'subtitle' : { H1: 'h1', H2: 'h2', H3: 'h3', BLOCKQUOTE: 'quote', PRE: 'code' }[blk.tagName] || 'p') : 'p';
  $$('.gallery-item').forEach((g) => g.style.borderColor = g.dataset.id === sid ? 'var(--accent)' : '');
  shell.ribbon.showTab('table', !!a.closest('td,th'));
}
const sbPage = h('span'); const sbWords = h('button', { class: 'sb-btn', onclick: () => wordCount() }); const sbLang = h('span', { class: 'hide-m' }, 'Русский');
const zoomLbl = h('button', { class: 'sb-btn', onclick: () => setZoom(1) }, '100%');
const zoomRange = h('input', { type: 'range', min: 30, max: 300, value: 100, class: 'hide-m' });
zoomRange.oninput = () => setZoom(zoomRange.value / 100);
function updateStatus() {
  const t = editor.innerText; const words = (t.match(/[\p{L}\p{N}]+(?:[-'’][\p{L}\p{N}]+)*/gu) || []).length;
  const ph = page.offsetHeight / pageCount * zoom; const cur = clamp(Math.floor((scroll.scrollTop + scroll.clientHeight / 3) / ph) + 1, 1, pageCount);
  sbPage.textContent = `Страница ${cur} из ${pageCount}`; sbWords.textContent = `Слов: ${words.toLocaleString('ru-RU')}`;
}
scroll.addEventListener('scroll', debounce(updateStatus, 80));
function wordCount() {
  const t = editor.innerText;
  const words = (t.match(/[\p{L}\p{N}]+(?:[-'’][\p{L}\p{N}]+)*/gu) || []).length;
  const chars = t.replace(/\n/g, '').length, noSp = t.replace(/\s/g, '').length;
  const paras = [...editor.querySelectorAll(BLOCK_SEL)].filter((p) => p.textContent.trim()).length;
  dialog({ title: 'Статистика', body: `<table class="data"><tr><td>Страниц</td><td><b>${pageCount}</b></td></tr><tr><td>Слов</td><td><b>${words}</b></td></tr><tr><td>Знаков (без пробелов)</td><td><b>${noSp}</b></td></tr><tr><td>Знаков (с пробелами)</td><td><b>${chars}</b></td></tr><tr><td>Абзацев</td><td><b>${paras}</b></td></tr><tr><td>Авторских листов</td><td><b>${(chars / 40000).toFixed(2)}</b></td></tr></table>` });
}

/* ======================= change tracking ======================= */
function changed() {
  shell?.markDirty();
  layoutMarksDeb(); refreshNav();
}
const layoutMarksDeb = debounce(layoutMarks, 150);
editor.addEventListener('input', () => { fixFontSizes(); if (!editor.innerHTML.trim() || editor.innerHTML === '<br>') editor.innerHTML = '<p><br></p>'; changed(); renumberDeb(); });
hfHead.addEventListener('input', () => { doc.header.text = hfHead.textContent; updatePrintStyle(); shell.markDirty(); });
hfFoot.addEventListener('input', () => { doc.footer.text = hfFoot.textContent; updatePrintStyle(); shell.markDirty(); });
new ResizeObserver(() => layoutMarksDeb()).observe(editor);

/* paste & drop */
const ALLOWED = new Set(['P', 'BR', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'SUB', 'SUP', 'A', 'UL', 'OL', 'LI', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'CODE', 'TABLE', 'THEAD', 'TBODY', 'TR', 'TD', 'TH', 'IMG', 'HR', 'SPAN', 'DIV', 'FONT']);
const STYLE_OK = ['font-weight', 'font-style', 'text-decoration', 'text-decoration-line', 'color', 'background-color', 'text-align', 'font-size', 'font-family', 'vertical-align'];
function sanitize(html) {
  const t = document.createElement('template'); t.innerHTML = html.replace(/<!--[\s\S]*?-->/g, '');
  const walk = (node) => {
    for (const el of [...node.children]) {
      if (['SCRIPT', 'STYLE', 'META', 'LINK', 'TITLE', 'XML', 'IFRAME', 'OBJECT', 'NOSCRIPT', 'svg'].includes(el.tagName) || el.tagName.includes(':')) { el.remove(); continue; }
      walk(el);
      if (!ALLOWED.has(el.tagName)) { el.replaceWith(...el.childNodes); continue; }
      const keep = {};
      if (el.tagName === 'A') keep.href = el.getAttribute('href');
      if (el.tagName === 'IMG') keep.src = el.getAttribute('src');
      if (el.colSpan > 1) keep.colspan = el.colSpan;
      const st = STYLE_OK.map((p) => [p, el.style.getPropertyValue(p)]).filter(([, v]) => v && !/^(windowtext|inherit|initial)$/.test(v)).map(([p, v]) => `${p}:${v}`).join(';');
      for (const a of [...el.attributes]) el.removeAttribute(a.name);
      for (const [k, v] of Object.entries(keep)) if (v != null && !(k === 'href' && /^\s*javascript:/i.test(v))) el.setAttribute(k, v);
      if (st) el.setAttribute('style', st);
    }
  };
  walk(t.content);
  return t.innerHTML;
}
editor.addEventListener('paste', async (e) => {
  const cd = e.clipboardData; if (!cd) return;
  const img = [...cd.files].find((f) => f.type.startsWith('image/'));
  if (img) { e.preventDefault(); const r = new FileReader(); r.onload = () => insertHTML(`<img src="${r.result}" style="max-width:100%">`); r.readAsDataURL(img); return; }
  const html = cd.getData('text/html');
  if (doc.track) { e.preventDefault(); trackedInsert(html ? sanitize(html) : esc(cd.getData('text/plain')).replace(/\n/g, '<br>')); return; }
  if (html) { e.preventDefault(); insertHTML(sanitize(html)); }
});
editor.addEventListener('drop', (e) => {
  const f = [...(e.dataTransfer?.files || [])].find((x) => x.type.startsWith('image/'));
  if (!f) return; e.preventDefault();
  const pos = document.caretRangeFromPoint?.(e.clientX, e.clientY); if (pos) savedRange = pos;
  const r = new FileReader(); r.onload = () => insertHTML(`<img src="${r.result}" style="max-width:100%">`); r.readAsDataURL(f);
});

/* ======================= serialization ======================= */
function cleanHTML() {
  const c = editor.cloneNode(true);
  c.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel'));
  c.querySelectorAll('.cmt-anchor.active').forEach((x) => x.classList.remove('active'));
  c.querySelectorAll('[class=""]').forEach((x) => x.removeAttribute('class'));
  return c.innerHTML;
}
function serialize() { return JSON.stringify({ format: 'rabbit-document', version: 1, ...doc, title: shell.title, html: cleanHTML() }); }
async function load(file) {
  const ext = extOf(file.name); const text = () => new TextDecoder().decode(file.bytes);
  doc = newDocState();
  if (ext === 'rdoc') {
    const d = JSON.parse(text());
    doc.page = { ...doc.page, ...d.page }; doc.header = { ...doc.header, ...d.header }; doc.footer = { ...doc.footer, ...d.footer };
    doc.style = { ...doc.style, ...d.style }; doc.comments = d.comments || [];
    doc.track = !!d.track; doc.watermark = d.watermark || ''; doc.readonly = !!d.readonly; doc.hyphens = !!d.hyphens;
    editor.innerHTML = sanitizeStored(d.html || '<p><br></p>');
  } else if (ext === 'docx') {
    const r = await importDocx(file.bytes);
    if (r.page.w) Object.assign(doc.page, { w: r.page.w, h: r.page.h, landscape: r.page.landscape, size: Object.entries(PAGE_SIZES).find(([, v]) => Math.abs(v[0] - r.page.w) < 2 && Math.abs(v[1] - r.page.h) < 2)?.[0] || 'custom' });
    if (r.page.margins) { doc.page.margins = r.page.margins; doc.page.marginsPreset = 'custom'; }
    editor.innerHTML = r.html || '<p><br></p>';
  } else if (ext === 'html' || ext === 'htm') {
    const d = new DOMParser().parseFromString(text(), 'text/html');
    editor.innerHTML = sanitize(d.body.innerHTML) || '<p><br></p>';
  } else if (ext === 'md') editor.innerHTML = mdToHtml(text());
  else editor.innerHTML = text().split(/\r?\n/).map((l) => `<p>${esc(l) || '<br>'}</p>`).join('');
  renumberFootnotes();
  applyPage(); renderComments(); refreshNav(); syncRibbon();
}
function sanitizeStored(html) { const t = document.createElement('template'); t.innerHTML = html; t.content.querySelectorAll('script,iframe,object').forEach((x) => x.remove()); t.content.querySelectorAll('*').forEach((el) => { for (const a of [...el.attributes]) if (/^on/i.test(a.name) || (a.name === 'href' && /^\s*javascript:/i.test(a.value))) el.removeAttribute(a.name); }); return t.innerHTML; }

/* ======================= markdown ======================= */
function mdInline(s) {
  return esc(s).replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img alt="$1" src="$2">').replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
    .replace(/`([^`]+)`/g, '<code>$1</code>').replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/__([^_]+)__/g, '<b>$1</b>').replace(/\*([^*]+)\*/g, '<i>$1</i>').replace(/(^|\W)_([^_]+)_(?=\W|$)/g, '$1<i>$2</i>').replace(/~~([^~]+)~~/g, '<s>$1</s>');
}
function mdToHtml(md) {
  const lines = md.replace(/\r/g, '').split('\n'); let out = ''; let i = 0;
  while (i < lines.length) {
    const l = lines[i];
    if (/^```/.test(l)) { let code = ''; i++; while (i < lines.length && !/^```/.test(lines[i])) code += lines[i++] + '\n'; i++; out += `<pre>${esc(code.replace(/\n$/, ''))}</pre>`; continue; }
    const hm = /^(#{1,6})\s+(.*)/.exec(l); if (hm) { const n = Math.min(3, hm[1].length); out += `<h${n}>${mdInline(hm[2])}</h${n}>`; i++; continue; }
    if (/^(\*{3,}|-{3,}|_{3,})\s*$/.test(l)) { out += '<hr>'; i++; continue; }
    if (/^>\s?/.test(l)) { let q = ''; while (i < lines.length && /^>\s?/.test(lines[i])) q += lines[i++].replace(/^>\s?/, '') + ' '; out += `<blockquote>${mdInline(q.trim())}</blockquote>`; continue; }
    if (/^\s*([-*+]|\d+[.)])\s+/.test(l)) {
      const ordered = /^\s*\d/.test(l); const tag = ordered ? 'ol' : 'ul'; let items = '', check = false;
      while (i < lines.length && /^\s*([-*+]|\d+[.)])\s+/.test(lines[i])) { let t = lines[i++].replace(/^\s*([-*+]|\d+[.)])\s+/, ''); const cm = /^\[( |x)\]\s+/i.exec(t); if (cm) { check = true; t = t.slice(cm[0].length); items += `<li${cm[1] !== ' ' ? ' class="done"' : ''}>${mdInline(t)}</li>`; } else items += `<li>${mdInline(t)}</li>`; }
      out += `<${tag}${check ? ' class="checklist"' : ''}>${items}</${tag}>`; continue;
    }
    if (/^\|.*\|\s*$/.test(l)) {
      const rows = []; while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) rows.push(lines[i++]);
      const cells = rows.filter((r) => !/^\|[\s:-|]+\|$/.test(r)).map((r) => r.slice(1, -1).split('|').map((c) => mdInline(c.trim())));
      out += `<table class="doc-table header-row"><tbody>${cells.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>`; continue;
    }
    if (!l.trim()) { i++; continue; }
    let p = l; i++; while (i < lines.length && lines[i].trim() && !/^(#|>|```|\s*([-*+]|\d+[.)])\s|\|)/.test(lines[i])) p += ' ' + lines[i++];
    out += `<p>${mdInline(p)}</p>`;
  }
  return out || '<p><br></p>';
}
function htmlToMd(root) {
  const inl = (n) => [...n.childNodes].map((c) => {
    if (c.nodeType === 3) return c.data.replace(/\u00a0/g, ' ');
    if (c.nodeType !== 1) return '';
    const t = c.tagName; const inner = inl(c);
    if (t === 'B' || t === 'STRONG' || +c.style.fontWeight >= 600 || c.style.fontWeight === 'bold') return inner.trim() ? `**${inner}**` : inner;
    if (t === 'I' || t === 'EM' || c.style.fontStyle === 'italic') return inner.trim() ? `*${inner}*` : inner;
    if (t === 'S' || t === 'STRIKE' || t === 'DEL') return `~~${inner}~~`;
    if (t === 'CODE') return '`' + inner + '`';
    if (t === 'A') return `[${inner}](${c.getAttribute('href')})`;
    if (t === 'IMG') return `![${c.alt || ''}](${c.src.startsWith('data:') ? 'image' : c.src})`;
    if (t === 'BR') return '  \n';
    return inner;
  }).join('');
  const blk = (n, depth = 0) => {
    if (n.nodeType === 3) return n.data.trim() ? n.data + '\n\n' : '';
    if (n.nodeType !== 1) return '';
    const t = n.tagName;
    if (n.classList.contains('page-break')) return '\n---\n\n';
    if (n.classList.contains('toc')) return '';
    if (/^H[1-6]$/.test(t)) return '#'.repeat(n.classList.contains('doc-title') ? 1 : +t[1]) + ' ' + inl(n).trim() + '\n\n';
    if (t === 'BLOCKQUOTE') return inl(n).split('\n').map((l) => '> ' + l).join('\n') + '\n\n';
    if (t === 'PRE') return '```\n' + n.textContent + '\n```\n\n';
    if (t === 'HR') return '---\n\n';
    if (t === 'UL' || t === 'OL') return [...n.children].map((li, i) => '  '.repeat(depth) + (n.classList.contains('checklist') ? `- [${li.classList.contains('done') ? 'x' : ' '}] ` : t === 'OL' ? `${i + 1}. ` : '- ') + inl([...li.childNodes].filter((c) => !/^(UL|OL)$/.test(c.tagName)).reduce((d, c) => (d.append(c.cloneNode(true)), d), document.createElement('div'))).trim() + '\n' + [...li.children].filter((c) => /^(UL|OL)$/.test(c.tagName)).map((c) => blk(c, depth + 1)).join('')).join('') + (depth ? '' : '\n');
    if (t === 'TABLE') { const rows = [...n.rows].map((r) => '| ' + [...r.cells].map((c) => inl(c).trim().replace(/\|/g, '\\|')).join(' | ') + ' |'); if (rows.length) rows.splice(1, 0, '|' + ' --- |'.repeat(n.rows[0].cells.length)); return rows.join('\n') + '\n\n'; }
    if (t === 'DIV' && [...n.children].some((c) => /^(P|DIV|H\d|UL|OL|TABLE)$/.test(c.tagName))) return [...n.childNodes].map((c) => blk(c, depth)).join('');
    const s = inl(n).trim(); return s ? s + '\n\n' : '';
  };
  return [...root.childNodes].map((n) => blk(n)).join('').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

/* ======================= export ======================= */
function standaloneHTML() {
  const css = `body{font-family:"${doc.style.font}",system-ui,sans-serif;font-size:${doc.style.size}pt;line-height:${doc.style.lh};max-width:${(doc.page.w - doc.page.margins[1] - doc.page.margins[3]).toFixed(0)}mm;margin:40px auto;padding:0 20px;color:#1f1f1f}
h1,h2,h3{color:#1f3d7a;line-height:1.25}h1.doc-title{font-size:2.4em;font-weight:300;color:#14213d}table{border-collapse:collapse;width:100%}td,th{border:1px solid #b9bfca;padding:5px 8px}
blockquote{border-left:3px solid #9fb3d9;margin:.8em 0;padding:.3em 1em;color:#4a5266;font-style:italic}pre{background:#f4f5f7;padding:.7em;border-radius:4px}img{max-width:100%}.cmt-anchor{background:none}.page-break{break-after:page}.toc button{display:none}img.fl{float:left;margin:4px 14px 8px 0}img.fr{float:right;margin:4px 0 8px 14px}img.ctr{display:block;margin:8px auto}`;
  return `<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>${esc(shell.title)}</title><style>${css}</style></head><body>${finalClone().innerHTML}</body></html>`;
}
async function doPrint(pdf = false) {
  updatePrintStyle(); selectImage(null); closeFind();
  if (pdf) { const p = await printToPdf(shell.title + '.pdf', { pageSize: doc.page.size in PAGE_SIZES ? doc.page.size : 'A4', landscape: doc.page.landscape }); if (p) toast('PDF сохранён', 'ok'); }
  else window.print();
}

/* ======================= templates ======================= */
const today = () => new Date().toLocaleDateString('ru-RU', { day: 'numeric', month: 'long', year: 'numeric' });
const TPL = {
  letter: () => `<p style="text-align:right"><b>ООО «Компания»</b><br>г. Москва, ул. Примерная, д. 1<br>+7 (000) 000-00-00 · info@company.ru</p><p><br></p><p>${today()}</p><p>Генеральному директору<br>ООО «Партнёр»<br>И. И. Иванову</p><p><br></p><h2 style="text-align:center">Уважаемый Иван Иванович!</h2><p>Настоящим письмом сообщаем вам о&nbsp;… Опишите суть обращения: что произошло, что вы предлагаете и какой ответ ожидаете.</p><p>Будем признательны за ответ до&nbsp;…</p><p><br></p><p>С уважением,<br>Генеральный директор<br><br>_______________ / П. П. Петров</p>`,
  report: () => `<h1 class="doc-title">Отчёт о работе за квартал</h1><h4 class="doc-subtitle">Подготовил: ${esc(profile().name)} · ${today()}</h4><div class="toc" contenteditable="false"></div><h1>1. Краткое резюме</h1><p>Основные результаты периода в двух–трёх предложениях.</p><h1>2. Ключевые показатели</h1><table class="doc-table header-row banded"><tbody><tr><td>Показатель</td><td>План</td><td>Факт</td><td>Отклонение</td></tr><tr><td>Выручка, млн ₽</td><td>12,0</td><td>13,4</td><td>+11,7%</td></tr><tr><td>Новые клиенты</td><td>150</td><td>142</td><td>−5,3%</td></tr><tr><td>NPS</td><td>45</td><td>51</td><td>+6</td></tr></tbody></table><h2>2.1 Анализ</h2><ul><li>Что сработало хорошо</li><li>Что не получилось и почему</li></ul><h1>3. Планы на следующий период</h1><ol><li>Первая задача</li><li>Вторая задача</li></ol><blockquote>Важный вывод или цитата, которую стоит выделить.</blockquote>`,
  resume: () => `<h1 class="doc-title">Анна Смирнова</h1><h4 class="doc-subtitle">Продуктовый дизайнер · Москва · anna@mail.ru · +7 900 000-00-00</h4><h2>О себе</h2><p>Дизайнер с 6-летним опытом в B2B и финтехе. Люблю исследования, системы и аккуратные интерфейсы.</p><h2>Опыт работы</h2><h3>Ведущий дизайнер — «Финтех Лаб» (2021 — н. в.)</h3><ul><li>Запустила дизайн-систему из 120 компонентов</li><li>Увеличила конверсию онбординга на 23%</li></ul><h3>Дизайнер интерфейсов — «Студия Север» (2018 — 2021)</h3><ul><li>Более 30 проектов для клиентов из ритейла и логистики</li></ul><h2>Навыки</h2><p>UX-исследования · Прототипирование · Дизайн-системы · HTML/CSS · Аналитика</p><h2>Образование</h2><p>НИУ ВШЭ, «Дизайн», бакалавриат, 2018</p>`,
  notes: () => `<h1 class="doc-title">Протокол встречи</h1><p><b>Дата:</b> ${today()}<br><b>Участники:</b> …<br><b>Ведущий:</b> ${esc(profile().name)}</p><h2>Повестка</h2><ol><li>Статус проекта</li><li>Риски</li><li>Следующие шаги</li></ol><h2>Решения</h2><ul><li>…</li></ul><h2>Задачи</h2><ul class="checklist"><li>Подготовить смету — Иванов, до пятницы</li><li>Согласовать макеты — Петрова</li><li class="done">Разослать протокол</li></ul>`,
};
async function newDoc(tpl) {
  doc = newDocState();
  editor.innerHTML = TPL[tpl] ? TPL[tpl]() : '<p><br></p>';
  editor.querySelectorAll('.toc').forEach((t) => (t.innerHTML = tocInner()));
  if (tpl === 'letter' || tpl === 'report') { doc.style = { preset: 'gost', ...DOC_STYLES.gost }; doc.page.margins = [...MARGINS.gost[1]]; doc.page.marginsPreset = 'gost'; }
  if (tpl === 'report') { doc.footer.pageNumbers = true; doc.footer.align = 'right'; }
  if (tpl) shell.setTitle(TEMPLATES.docs.find((t) => t.id === tpl)?.name || 'Документ'); else shell.setTitle('Документ');
  applyPage(); renderComments(); syncRibbon();
  setTimeout(() => { editor.focus(); const r = document.createRange(); r.setStart(editor.firstChild || editor, 0); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); }, 30);
}
function syncRibbon() { shell?.ribbon.set('docstyle', { value: doc.style.preset }); shell?.ribbon.set('pgnum', { on: doc.footer.pageNumbers }); }

/* ======================= text-node helpers ======================= */
// Split text nodes at the range edges and return the text nodes fully inside the range.
function isolate(range) {
  const root = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
  const sc = range.startContainer, so = range.startOffset, ec = range.endContainer, eo = range.endOffset;
  const nodes = []; const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let n = w.nextNode(); n; n = w.nextNode()) if (n.data.length && range.intersectsNode(n)) nodes.push(n);
  const parts = [];
  for (let n of nodes) {
    const s = n === sc ? so : 0, e = n === ec ? eo : n.data.length;
    if (s >= e) continue;
    if (e < n.data.length) n.splitText(e);
    if (s > 0) n = n.splitText(s);
    if (n.parentElement?.closest('[contenteditable="false"]')) continue; // toc, footnote marks
    parts.push(n);
  }
  return parts;
}
function styleSelection(css) {
  restoreSel(); const s = getSelection(); if (!s.rangeCount || s.isCollapsed) { toast('Выделите текст'); return; }
  const parts = isolate(s.getRangeAt(0));
  for (const n of parts) { const sp = document.createElement('span'); for (const [k, v] of Object.entries(css)) sp.style.setProperty(k, v); n.replaceWith(sp); sp.append(n); }
  if (parts.length) { const r = document.createRange(); r.setStartBefore(parts[0].parentNode); r.setEndAfter(parts[parts.length - 1].parentNode); s.removeAllRanges(); s.addRange(r); savedRange = r.cloneRange(); }
  changed();
}

/* ======================= track changes ======================= */
function trkEl(tag) {
  const p = profile(); const el = document.createElement(tag);
  el.className = 'trk'; el.dataset.a = p.name; el.dataset.t = String(Date.now()); el.style.setProperty('--tc', p.color);
  el.title = `${tag === 'ins' ? 'Вставлено' : 'Удалено'}: ${p.name}, ${new Date().toLocaleString('ru-RU')}`;
  return el;
}
const ownIns = (node) => { const el = (node.nodeType === 3 ? node.parentElement : node)?.closest?.('ins.trk'); return el && el.dataset.a === profile().name ? el : null; };
function markDeleted(range) {
  const out = [];
  for (const n of isolate(range)) {
    const ins = ownIns(n);
    if (ins) { n.remove(); if (!ins.textContent) ins.remove(); continue; } // deleting own fresh insertion removes it for real
    if (n.parentElement.closest('del.trk')) { out.push(n.parentElement.closest('del.trk')); continue; }
    const d = trkEl('del'); n.replaceWith(d); d.append(n); out.push(d);
  }
  // merge neighbouring deletions by the same author into one mark
  const me = profile().name;
  return [...new Set(out.map((d) => {
    for (const side of ['previousSibling', 'nextSibling']) {
      const sib = d.isConnected && d[side];
      if (sib?.nodeType === 1 && sib.matches('del.trk') && sib.dataset.a === me && d.dataset.a === me) { if (side === 'previousSibling') { sib.append(...d.childNodes); d.remove(); d = sib; } else { d.append(...sib.childNodes); sib.remove(); } }
    }
    return d;
  }))].filter((d) => d.isConnected);
}
function placeCaret(node, after) { const r = document.createRange(); if (after) r.setStartAfter(node); else r.setStartBefore(node); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); savedRange = r.cloneRange(); }
function trackedInsert(html) {
  if (document.activeElement !== editor || !getSelection().rangeCount) restoreSel(); // live selection is fresher than savedRange
  const s = getSelection(); if (!s.rangeCount) return;
  if (!s.isCollapsed) { const d = markDeleted(s.getRangeAt(0).cloneRange()); if (d.length) placeCaret(d[d.length - 1], true); else s.collapseToStart(); }
  const r = s.getRangeAt(0); const anc = r.startContainer.nodeType === 3 ? r.startContainer.parentElement : r.startContainer;
  const del = anc.closest?.('del.trk'); if (del) { r.setStartAfter(del); r.collapse(true); }
  const ins = trkEl('ins'); ins.innerHTML = html; r.insertNode(ins);
  // caret goes to the end of the inserted text so further typing extends the same <ins>
  const last = ins.lastChild;
  if (last?.nodeType === 3) { const c = document.createRange(); c.setStart(last, last.length); c.collapse(true); s.removeAllRanges(); s.addRange(c); savedRange = c.cloneRange(); } else placeCaret(ins, true);
  changed();
}
editor.addEventListener('beforeinput', (e) => {
  if (!doc.track) return;
  const s = getSelection(); if (!s.rangeCount) return;
  const t = e.inputType;
  if (t === 'insertText' || t === 'insertReplacementText') {
    // all typing is handled manually: Chrome's own editing moves text out of <ins> and rewrites spaces
    e.preventDefault();
    const text = e.data ?? e.dataTransfer?.getData('text/plain') ?? ''; if (!text) return;
    if (!s.isCollapsed) { const d = markDeleted(s.getRangeAt(0).cloneRange()); if (d.length) placeCaret(d[d.length - 1], true); }
    const r = getSelection().getRangeAt(0);
    let node = r.startContainer, off = r.startOffset;
    const prevEl = node.nodeType === 1 ? node.childNodes[off - 1] : off === 0 ? node.previousSibling : null;
    if (prevEl?.matches?.('ins.trk') && ownIns(prevEl)) { // caret just after own insertion: extend it
      let last = prevEl.lastChild; if (last?.nodeType !== 3) last = prevEl.appendChild(document.createTextNode(''));
      node = last; off = last.length;
    }
    if (node.nodeType === 3 && ownIns(node) && !node.parentElement.closest('del.trk')) {
      let ins = text;
      if (ins === ' ' && (off === node.length || node.data[off] === ' ')) ins = ' ';
      if (ins !== ' ' && off > 1 && node.data[off - 1] === ' ' && !/\s/.test(node.data[off - 2])) node.replaceData(off - 1, 1, ' ');
      node.insertData(off, ins);
      const c = document.createRange(); c.setStart(node, off + ins.length); c.collapse(true); s.removeAllRanges(); s.addRange(c); savedRange = c.cloneRange();
      changed(); return;
    }
    trackedInsert(esc(text).replace(/^ | $/g, '&nbsp;'));
    return;
  }
  if (t.startsWith('delete')) {
    const back = t.includes('Backward');
    if (s.isCollapsed) s.modify('extend', back ? 'backward' : 'forward', t.includes('Word') ? 'word' : t.includes('Line') || t.includes('Soft') ? 'lineboundary' : 'character');
    if (s.isCollapsed) return;
    const r = s.getRangeAt(0).cloneRange();
    const a = ownIns(r.startContainer), b = ownIns(r.endContainer);
    if (a && a === b) return; // native delete inside own insertion
    e.preventDefault();
    const dels = markDeleted(r);
    if (dels.length) placeCaret(back ? dels[0] : dels[dels.length - 1], !back); else { r.collapse(true); s.removeAllRanges(); s.addRange(r); }
    changed();
  }
});
function resolveTrk(el, accept) { if (el.tagName === 'INS') { if (accept) el.replaceWith(...el.childNodes); else el.remove(); } else if (accept) el.remove(); else el.replaceWith(...el.childNodes); }
function trkAll(accept) { const n = editor.querySelectorAll('ins.trk, del.trk').length; editor.querySelectorAll('ins.trk, del.trk').forEach((el) => resolveTrk(el, accept)); editor.normalize(); changed(); toast(`${accept ? 'Принято' : 'Отклонено'} исправлений: ${n}`, 'ok'); }
function trkStep(dir) {
  const all = [...editor.querySelectorAll('ins.trk, del.trk')]; if (!all.length) { toast('Исправлений нет'); return null; }
  let el;
  if (savedRange) el = dir > 0 ? all.find((x) => savedRange.comparePoint(x, 0) === 1) : [...all].reverse().find((x) => savedRange.comparePoint(x, 0) === -1);
  el ||= dir > 0 ? all[0] : all[all.length - 1];
  const r = document.createRange(); r.selectNodeContents(el); const s = getSelection(); s.removeAllRanges(); s.addRange(r); savedRange = r.cloneRange();
  el.scrollIntoView({ behavior: 'smooth', block: 'center' });
  return el;
}
function trkCurrent(accept) {
  const el = anchorEl()?.closest('ins.trk, del.trk') || trkStep(1); if (!el) return;
  resolveTrk(el, accept); editor.normalize(); changed();
  if (editor.querySelector('ins.trk, del.trk')) trkStep(1);
}
function setTrack(on) { doc.track = on; applyPage(); shell.markDirty(); toast(on ? 'Исправления записываются' : 'Запись исправлений выключена'); }
function finalClone() { const c = editor.cloneNode(true); c.querySelectorAll('del.trk').forEach((x) => x.remove()); c.querySelectorAll('ins.trk').forEach((x) => x.replaceWith(...x.childNodes)); return c; }
function finalText() { const c = finalClone(); c.style.cssText = 'position:fixed;left:-99999px;top:0;width:800px'; document.body.append(c); const t = c.innerText; c.remove(); return t; }

/* ======================= footnotes ======================= */
function insertFootnote() {
  if (doc.readonly) return;
  const id = uid('fn');
  insertHTML(`<sup class="fn-ref" contenteditable="false" data-fn="${id}">*</sup>`);
  let sec = editor.querySelector('section.footnotes');
  if (!sec) { sec = h('section', { class: 'footnotes' }); editor.append(sec); }
  const p = h('p', { class: 'fn', 'data-fn': id }, h('span', { class: 'fn-num', contenteditable: 'false' }), ' ');
  sec.append(p); renumberFootnotes();
  const r = document.createRange(); r.selectNodeContents(p); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); savedRange = r.cloneRange();
  p.scrollIntoView({ behavior: 'smooth', block: 'center' }); changed();
}
function renumberFootnotes() {
  const refs = [...editor.querySelectorAll('sup.fn-ref')].filter((r) => !r.closest('section.footnotes'));
  let sec = editor.querySelector('section.footnotes');
  if (!refs.length) { sec?.remove(); return; }
  if (!sec) { sec = h('section', { class: 'footnotes' }); editor.append(sec); }
  if (sec !== editor.lastElementChild) editor.append(sec);
  const notes = new Map([...sec.querySelectorAll('p.fn')].map((p) => [p.dataset.fn, p]));
  refs.forEach((ref, i) => {
    ref.textContent = String(i + 1);
    let p = notes.get(ref.dataset.fn);
    if (!p) p = h('p', { class: 'fn', 'data-fn': ref.dataset.fn }, h('span', { class: 'fn-num', contenteditable: 'false' }), ' ');
    let num = p.querySelector('.fn-num'); if (!num) { num = h('span', { class: 'fn-num', contenteditable: 'false' }); p.prepend(num); }
    num.textContent = String(i + 1); sec.append(p); notes.delete(ref.dataset.fn);
  });
  for (const p of notes.values()) p.remove();
}
const renumberDeb = debounce(renumberFootnotes, 400);

/* ======================= dialogs: font & paragraph ======================= */
async function fontDialog() {
  const a = anchorEl(); const cs = a ? getComputedStyle(a) : null;
  const r = await formDlg('Шрифт', [
    { key: 'font', label: 'Шрифт', type: 'select', options: FONTS, value: cs?.fontFamily.split(',')[0].replace(/["']/g, '').trim() || doc.style.font },
    { key: 'size', label: 'Размер, пт', type: 'number', value: currentPt(), min: 1, max: 400, step: 0.5 },
    { key: 'weight', label: 'Начертание', type: 'select', options: [{ value: '', label: 'Без изменений' }, { value: 'normal', label: 'Обычный' }, { value: 'bold', label: 'Полужирный' }, { value: 'italic', label: 'Курсив' }, { value: 'bolditalic', label: 'Полужирный курсив' }] },
    { key: 'color', label: 'Цвет текста', type: 'color', value: '#000000' },
    { key: 'effect', label: 'Видоизменение', type: 'select', options: [{ value: '', label: 'Нет' }, { value: 'small-caps', label: 'Малые прописные' }, { value: 'uppercase', label: 'Все прописные' }, { value: 'underline', label: 'Подчёркнутый' }, { value: 'line-through', label: 'Зачёркнутый' }, { value: 'shadow', label: 'С тенью' }] },
    { key: 'spacing', label: 'Межзнаковый интервал, пт (0 — обычный)', type: 'number', value: 0, step: 0.5, min: -5, max: 20 },
  ], 'Применить');
  if (!r) return;
  const css = { 'font-family': `"${r.font}"`, 'font-size': r.size + 'pt' };
  if (r.weight === 'normal') Object.assign(css, { 'font-weight': '400', 'font-style': 'normal' });
  if (r.weight === 'bold' || r.weight === 'bolditalic') css['font-weight'] = '700';
  if (r.weight === 'italic' || r.weight === 'bolditalic') css['font-style'] = 'italic';
  if (r.color && r.color !== '#000000') css.color = r.color;
  if (r.effect === 'small-caps') css['font-variant'] = 'small-caps';
  if (r.effect === 'uppercase') css['text-transform'] = 'uppercase';
  if (r.effect === 'underline' || r.effect === 'line-through') css['text-decoration'] = r.effect;
  if (r.effect === 'shadow') css['text-shadow'] = '1px 1px 2px rgba(0,0,0,.35)';
  if (+r.spacing) css['letter-spacing'] = r.spacing + 'pt';
  styleSelection(css);
}
async function paragraphDialog() {
  const blocks = selectedBlocks(); if (!blocks.length) { ensureBlock(); }
  const b0 = selectedBlocks()[0]; const cs = b0 ? getComputedStyle(b0) : null;
  const cm = (px) => Math.round((parseFloat(px) || 0) / 37.795 * 100) / 100; const pt = (px) => Math.round((parseFloat(px) || 0) * 0.75);
  const lhNow = cs ? (cs.lineHeight === 'normal' ? 1.2 : Math.round(parseFloat(cs.lineHeight) / parseFloat(cs.fontSize) * 100) / 100) : 1.5;
  const r = await formDlg('Абзац', [
    { key: 'align', label: 'Выравнивание', type: 'select', value: cs?.textAlign === 'start' ? 'left' : cs?.textAlign || 'left', options: [{ value: 'left', label: 'По левому краю' }, { value: 'center', label: 'По центру' }, { value: 'right', label: 'По правому краю' }, { value: 'justify', label: 'По ширине' }] },
    { key: 'left', label: 'Отступ слева, см', type: 'number', step: 0.25, value: cm(cs?.marginLeft) },
    { key: 'right', label: 'Отступ справа, см', type: 'number', step: 0.25, value: cm(cs?.marginRight) },
    { key: 'first', label: 'Первая строка, см (отрицательное — выступ)', type: 'number', step: 0.25, value: cm(cs?.textIndent) },
    { key: 'before', label: 'Интервал перед, пт', type: 'number', step: 1, value: pt(cs?.marginTop) },
    { key: 'after', label: 'Интервал после, пт', type: 'number', step: 1, value: pt(cs?.marginBottom) },
    { key: 'lh', label: 'Междустрочный множитель', type: 'number', step: 0.05, value: lhNow },
    { key: 'keep', label: 'Не разрывать абзац между страницами', type: 'checkbox', value: cs?.breakInside === 'avoid' },
  ], 'Применить');
  if (!r) return;
  for (const b of selectedBlocks()) {
    Object.assign(b.style, { textAlign: r.align, marginLeft: r.left ? r.left + 'cm' : '', marginRight: r.right ? r.right + 'cm' : '', textIndent: r.first ? r.first + 'cm' : '', marginTop: r.before + 'pt', marginBottom: r.after + 'pt', lineHeight: r.lh || '', breakInside: r.keep ? 'avoid' : '' });
    if (r.first < 0) b.style.paddingLeft = -r.first + 'cm'; else b.style.paddingLeft = '';
  }
  changed();
}

/* ======================= more insert features ======================= */
let shapeColor = '#2f62c9';
function insertShape(kind) {
  const sh = SHAPES[kind]; const line = sh.line;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" width="${line ? 240 : 180}" height="${line ? 24 : 120}">${sh.arrow ? `<defs><marker id="a" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="5" markerHeight="5" orient="auto"><path d="M0 0L10 5L0 10z" fill="${shapeColor}"/></marker></defs>` : ''}<path d="${sh.d}" fill="${line ? 'none' : shapeColor}" stroke="${line ? shapeColor : 'none'}" stroke-width="${line ? 4 : 0}" vector-effect="non-scaling-stroke" ${sh.arrow ? 'marker-end="url(#a)"' : ''}/></svg>`;
  insertHTML(`<img src="data:image/svg+xml;base64,${btoa(unescape(encodeURIComponent(svg)))}" alt="${esc(sh.n)}" style="width:${line ? 240 : 180}px">`);
}
const EQ_SYMS = [['√', 'sqrt()'], ['x²', '^2'], ['xᵢ', '_i'], ['∑', '\\sum'], ['∫', '\\int'], ['∞', '\\infty'], ['π', '\\pi'], ['α', '\\alpha'], ['β', '\\beta'], ['Δ', '\\Delta'], ['±', '+-'], ['≤', '<='], ['≥', '>='], ['≠', '!='], ['≈', '~='], ['·', '*'], ['→', '->'], ['∈', '\\in']];
const GREEK = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', lambda: 'λ', mu: 'μ', pi: 'π', sigma: 'σ', phi: 'φ', omega: 'ω', theta: 'θ', Delta: 'Δ', Sigma: 'Σ', Omega: 'Ω', Pi: 'Π', sum: '∑', int: '∫', infty: '∞', in: '∈', cdot: '·', times: '×', partial: '∂', nabla: '∇' };
function eqToHtml(src) {
  let s = esc(src).replace(/\\([A-Za-z]+)/g, (m, n) => GREEK[n] || m)
    .replace(/\+-/g, '±').replace(/&lt;=/g, '≤').replace(/&gt;=/g, '≥').replace(/!=/g, '≠').replace(/~=/g, '≈').replace(/-&gt;/g, '→').replace(/\*/g, '·');
  s = s.replace(/sqrt\(([^()]*)\)/g, '√<span style="border-top:1px solid">$1</span>');
  s = s.replace(/\(([^()]+)\)\/\(([^()]+)\)/g, '<span class="frac"><span>$1</span><span>$2</span></span>');
  s = s.replace(/\^\{([^}]*)\}|\^(\w+|\d+)/g, (m, a, b) => `<sup>${a ?? b}</sup>`).replace(/_\{([^}]*)\}|_(\w)/g, (m, a, b) => `<sub>${a ?? b}</sub>`);
  return s;
}
async function insertEquation() {
  const inp = h('input', { class: 'input mono', value: 'E = mc^2', placeholder: 'например: x_1 = (-b + sqrt(D))/(2a)' });
  const prev = h('div', { class: 'eq-preview' });
  const upd = () => { prev.innerHTML = `<span class="eq">${eqToHtml(inp.value)}</span>`; };
  inp.oninput = upd; upd();
  const pad = h('div', { class: 'row wrap', style: { gap: '4px', margin: '8px 0' } }, ...EQ_SYMS.map(([l, t]) => h('button', { class: 'btn sm', onclick: () => { const p = inp.selectionStart; inp.value = inp.value.slice(0, p) + t + inp.value.slice(inp.selectionEnd); inp.focus(); inp.setSelectionRange(p + t.length - (t.endsWith('()') ? 1 : 0), p + t.length - (t.endsWith('()') ? 1 : 0)); upd(); } }, l)));
  const ok = await dialog({ title: 'Формула', body: h('div', { style: { width: 'min(520px,100%)' } }, inp, pad, h('div', { class: 'small muted' }, 'Синтаксис: x^2 — степень, x_i — индекс, sqrt(x) — корень, (a)/(b) — дробь, \\alpha \\pi \\sum — символы.'), prev), buttons: [{ label: 'Отмена', value: false }, { label: 'Вставить', value: true, primary: true }] });
  if (ok && inp.value.trim()) insertHTML(`<span class="eq" data-src="${esc(inp.value)}">${eqToHtml(inp.value)}</span>&nbsp;`);
}
function toggleDropCap() { const b = selectedBlocks()[0]; if (!b || b.tagName !== 'P') return toast('Поставьте курсор в обычный абзац'); b.classList.toggle('dropcap'); changed(); }
function setWatermark(text) { doc.watermark = text; applyPage(); shell.markDirty(); }
function setReadonly(on) { doc.readonly = on; applyPage(); shell.markDirty(); if (on) toast('Документ защищён от изменений'); }
function imgRotate() { if (!selImg) return; const cur = +(/rotate\((\d+)/.exec(selImg.style.transform)?.[1] || 0); selImg.style.transform = `rotate(${(cur + 90) % 360}deg)`; changed(); placeHandle(); }
function imgStyle(kind) { if (!selImg) return; selImg.classList.remove('st-frame', 'st-shadow', 'st-round', 'st-circle'); if (kind) selImg.classList.add('st-' + kind); changed(); placeHandle(); }

/* table column resize by dragging cell borders */
let colDrag = false;
editor.addEventListener('pointermove', (e) => {
  if (colDrag) return; const td = e.target.closest?.('td,th');
  if (!td) { if (editor.style.cursor === 'col-resize') editor.style.cursor = ''; return; }
  const r = td.getBoundingClientRect(); editor.style.cursor = Math.abs(e.clientX - r.right) < 5 && td.nextElementSibling ? 'col-resize' : painter ? 'copy' : '';
});
editor.addEventListener('pointerdown', (e) => {
  const td = e.target.closest?.('td,th'); if (!td || !td.nextElementSibling) return;
  const r = td.getBoundingClientRect(); if (Math.abs(e.clientX - r.right) >= 5) return;
  e.preventDefault(); colDrag = true;
  const table = td.closest('table'); const idx = td.cellIndex; const tw = table.getBoundingClientRect().width;
  const w0 = r.width, nw0 = td.nextElementSibling.getBoundingClientRect().width; const sx = e.clientX;
  const mv = (ev) => { const d = clamp(ev.clientX - sx, 24 - w0, nw0 - 24); for (const row of table.rows) { const c = row.cells[idx], n = row.cells[idx + 1]; if (c && c.colSpan === 1) c.style.width = ((w0 + d) / tw * 100).toFixed(2) + '%'; if (n && n.colSpan === 1) n.style.width = ((nw0 - d) / tw * 100).toFixed(2) + '%'; } };
  const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); colDrag = false; editor.style.cursor = ''; changed(); };
  document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
});

/* ======================= ribbon ======================= */
const b = (id, ic, label, action, extra = {}) => ({ type: 'btn', id, icon: ic, label, action, ...extra });
const big = (id, ic, label, action, extra = {}) => ({ type: 'big', id, icon: ic, label, action, ...extra });
const gallery = h('div', { class: 'gallery' });
for (const s of BLOCK_STYLES) {
  const g = h('button', { class: 'gallery-item', 'data-id': s.id, html: `<div class="s" style="${s.css}">АаБбВв</div><div class="n">${s.name}</div>` });
  g.addEventListener('mousedown', (e) => e.preventDefault()); g.onclick = () => setBlockStyle(s.id);
  gallery.append(g);
}
const lineMenu = () => [1, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: String(v).replace('.', ','), action: () => setLineHeight(v) })).concat(['-', { label: 'Интервал перед абзацем: 0', action: () => setSpacing(0, null) }, { label: 'Интервал перед абзацем: 12 пт', action: () => setSpacing(12, null) }, { label: 'Интервал после абзаца: 0', action: () => setSpacing(null, 0) }, { label: 'Интервал после абзаца: 8 пт', action: () => setSpacing(null, 8) }]);
const setPageSize = (k) => { const [w, hh] = PAGE_SIZES[k]; Object.assign(doc.page, { size: k, w, h: hh }); applyPage(); shell.markDirty(); };
const setMargins = (k) => { doc.page.margins = [...MARGINS[k][1]]; doc.page.marginsPreset = k; applyPage(); shell.markDirty(); };
async function customMargins() {
  const m = doc.page.margins;
  const r = await formDlg('Настраиваемые поля (мм)', [{ key: 't', label: 'Верхнее', type: 'number', value: m[0] }, { key: 'b', label: 'Нижнее', type: 'number', value: m[2] }, { key: 'l', label: 'Левое', type: 'number', value: m[3] }, { key: 'r', label: 'Правое', type: 'number', value: m[1] }]);
  if (r) { doc.page.margins = [r.t, r.r, r.b, r.l].map((v) => clamp(v, 0, 100)); doc.page.marginsPreset = 'custom'; applyPage(); shell.markDirty(); }
}
function applyDocStyle(k) { doc.style = { preset: k, ...DOC_STYLES[k] }; applyPage(); shell.markDirty(); }

const ribbon = [
  { id: 'home', label: 'Главная', groups: [
    { label: 'Буфер обмена', items: [big('paste', 'paste', 'Вставить', async () => { try { const t = await navigator.clipboard.readText(); insertHTML(esc(t).replace(/\n/g, '<br>')); } catch { toast('Используйте Ctrl+V для вставки'); } }),
      { type: 'col', items: [b('cut', 'cut', 'Вырезать', () => cmd('cut'), { key: 'Ctrl+X' }), b('copy', 'copy', 'Копировать', () => cmd('copy'), { key: 'Ctrl+C' }), b('painter', 'brush', 'Формат по образцу', formatPainter)] }] },
    { label: 'Шрифт', items: [{ type: 'col', items: [
      { type: 'row', items: [{ type: 'select', id: 'font', width: 140, options: FONTS.map((f) => ({ value: f, label: f })), value: 'Calibri', tip: 'Шрифт', action: (v) => cmd('fontName', v) },
        { type: 'select', id: 'size', editable: true, width: 52, options: PT_SIZES, value: '11', tip: 'Размер шрифта', action: (v) => applyFontSize(v) },
        b('grow', 'text-size', 'Увеличить размер', () => growFont(1), { key: 'Ctrl+]' }), b('shrink', 'type', 'Уменьшить размер', () => growFont(-1), { key: 'Ctrl+[' }),
        b('case', 'case', 'Регистр', null, { menu: () => [{ label: 'Как в предложениях.', action: () => changeCase('sentence') }, { label: 'все строчные', action: () => changeCase('lower') }, { label: 'ВСЕ ПРОПИСНЫЕ', action: () => changeCase('upper') }, { label: 'Начинать С Прописных', action: () => changeCase('title') }] }),
        b('clear', 'clear-format', 'Очистить формат', () => { cmd('removeFormat'); for (const bl of selectedBlocks()) bl.removeAttribute('style'); changed(); })] },
      { type: 'row', items: [b('bold', 'bold', 'Полужирный', () => cmd('bold'), { key: 'Ctrl+B' }), b('italic', 'italic', 'Курсив', () => cmd('italic'), { key: 'Ctrl+I' }), b('underline', 'underline', 'Подчёркнутый', () => cmd('underline'), { key: 'Ctrl+U' }), b('strike', 'strike', 'Зачёркнутый', () => cmd('strikeThrough')),
        b('sub', 'sub', 'Подстрочный', () => cmd('subscript'), { key: 'Ctrl+=' }), b('sup', 'sup', 'Надстрочный', () => cmd('superscript'), { key: 'Ctrl+Shift+=' }),
        { type: 'color', id: 'hl', icon: 'highlight', tip: 'Цвет выделения текста', value: '#ffff00', allowNone: true, action: (c) => cmd('hiliteColor', c || 'transparent') },
        { type: 'color', id: 'fc', icon: 'font-color', tip: 'Цвет текста', value: '#c00000', action: (c) => cmd('foreColor', c || '#000000') }] }] }] },
    { label: 'Абзац', items: [{ type: 'col', items: [
      { type: 'row', items: [b('ul', 'list', 'Маркеры', () => cmd('insertUnorderedList'), { key: 'Ctrl+Shift+L' }), b('ol', 'list-ordered', 'Нумерация', () => cmd('insertOrderedList')),
        b('check', 'checklist', 'Список задач', () => { cmd('insertUnorderedList'); const ul = anchorEl()?.closest('ul'); if (ul) { ul.classList.toggle('checklist'); changed(); } }),
        b('outdent', 'outdent', 'Уменьшить отступ', () => cmd('outdent')), b('indent', 'indent', 'Увеличить отступ', () => cmd('indent')), b('marks', 'pilcrow', 'Отобразить все знаки', () => { editor.classList.toggle('show-marks'); shell.ribbon.set('marks', { on: editor.classList.contains('show-marks') }); })] },
      { type: 'row', items: [b('jl', 'align-left', 'По левому краю', () => cmd('justifyLeft'), { key: 'Ctrl+L' }), b('jc', 'align-center', 'По центру', () => cmd('justifyCenter'), { key: 'Ctrl+E' }), b('jr', 'align-right', 'По правому краю', () => cmd('justifyRight'), { key: 'Ctrl+R' }), b('jf', 'align-justify', 'По ширине', () => cmd('justifyFull'), { key: 'Ctrl+J' }),
        b('lh', 'line-spacing', 'Междустрочный интервал', null, { menu: lineMenu }), b('quote', 'quote', 'Цитата', () => setBlockStyle('quote'))] }] }] },
    { label: 'Стили', items: [{ type: 'custom', el: gallery }] },
    { label: 'Редактирование', items: [{ type: 'col', items: [b('find', 'search', 'Найти', () => openFind(false), { key: 'Ctrl+F', showLabel: true }), b('replace', 'replace', 'Заменить', () => openFind(true), { key: 'Ctrl+H', showLabel: true }), b('selall', 'maximize', 'Выделить всё', () => { editor.focus(); cmd('selectAll'); }, { showLabel: true })] }] },
  ] },
  { id: 'insert', label: 'Вставка', groups: [
    { label: 'Страницы', items: [big('pb', 'page-break', 'Разрыв страницы', () => insertHTML('<div class="page-break" contenteditable="false"></div><p><br></p>'), { key: 'Ctrl+Enter' }), big('toc', 'list-ordered', 'Оглавление', insertTOC)] },
    { label: 'Таблицы', items: [big('table', 'table', 'Таблица', (e, el) => tablePicker(el))] },
    { label: 'Иллюстрации', items: [big('img', 'image', 'Рисунок', insertImage), big('chart', 'chart', 'Диаграмма', insertChart)] },
    { label: 'Ссылки', items: [big('link', 'link', 'Ссылка', insertLink, { key: 'Ctrl+K' }), big('cmt', 'comment-plus', 'Примечание', addComment)] },
    { label: 'Колонтитулы', items: [{ type: 'col', items: [
      b('hdr', 'header', 'Верхний колонтитул', () => { doc.header.show = true; applyPage(); hfHead.focus(); }, { showLabel: true }),
      b('ftr', 'footer', 'Нижний колонтитул', () => { doc.footer.show = true; applyPage(); hfFoot.focus(); }, { showLabel: true }),
      b('pgnum', 'hash', 'Номера страниц', null, { showLabel: true, menu: () => [{ label: 'Внизу слева', action: () => pgn('left') }, { label: 'Внизу по центру', action: () => pgn('center') }, { label: 'Внизу справа', action: () => pgn('right') }, '-', { label: 'Убрать номера', action: () => { doc.footer.pageNumbers = false; applyPage(); syncRibbon(); shell.markDirty(); } }] })] }] },
    { label: 'Текст', items: [{ type: 'col', items: [b('hr', 'hr', 'Горизонтальная линия', () => insertHTML('<hr><p><br></p>'), { showLabel: true }), b('date', 'calendar', 'Дата и время', () => menu(document.activeElement, []) && null, { showLabel: true, menu: () => [today(), new Date().toLocaleDateString('ru-RU'), new Date().toLocaleString('ru-RU'), new Date().toISOString().slice(0, 10)].map((s) => ({ label: s, action: () => insertHTML(esc(s)) })) }),
      b('sym', 'omega', 'Символ', (e, el) => symbolPicker(el), { showLabel: true })] },
      { type: 'col', items: [b('code', 'code', 'Блок кода', () => setBlockStyle('code'), { showLabel: true })] }] },
  ] },
  { id: 'layout', label: 'Макет', groups: [
    { label: 'Параметры страницы', items: [
      big('margins', 'margins', 'Поля', null, { menu: () => [...Object.entries(MARGINS).map(([k, [n, m]]) => ({ label: n, hint: `${m[0]} · ${m[3]} мм`, checked: doc.page.marginsPreset === k, action: () => setMargins(k) })), '-', { label: 'Настраиваемые поля…', action: customMargins }] }),
      big('orient', 'orientation', 'Ориентация', null, { menu: () => [{ label: 'Книжная', checked: !doc.page.landscape, action: () => { doc.page.landscape = false; applyPage(); shell.markDirty(); } }, { label: 'Альбомная', checked: doc.page.landscape, action: () => { doc.page.landscape = true; applyPage(); shell.markDirty(); } }] }),
      big('psize', 'page-size', 'Размер', null, { menu: () => Object.entries(PAGE_SIZES).map(([k, [w, hh]]) => ({ label: k, hint: `${w} × ${hh} мм`, checked: doc.page.size === k, action: () => setPageSize(k) })) }),
      big('cols', 'columns', 'Колонки', null, { menu: () => [1, 2, 3].map((n) => ({ label: ['Одна', 'Две', 'Три'][n - 1], checked: doc.page.columns === n, action: () => { doc.page.columns = n; applyPage(); shell.markDirty(); } })) })] },
    { label: 'Абзац', items: [{ type: 'col', items: [b('sp0', 'line-spacing', 'Интервалы', null, { showLabel: true, menu: lineMenu })] }] },
    { label: 'Фон страницы', items: [{ type: 'color', id: 'pagecolor', icon: 'fill-color', tip: 'Цвет страницы', value: '#fff8e7', allowNone: true, action: (c) => { doc.page.color = c; applyPage(); shell.markDirty(); } }] },
  ] },
  { id: 'design', label: 'Конструктор', groups: [
    { label: 'Стиль документа', items: [{ type: 'col', items: [
      { type: 'select', id: 'docstyle', width: 170, options: Object.entries(DOC_STYLES).map(([k, s]) => ({ value: k, label: s.name })), value: 'modern', tip: 'Набор шрифтов и интервалов', action: applyDocStyle },
      { type: 'row', items: [{ type: 'select', id: 'basefont', width: 120, options: FONTS, tip: 'Основной шрифт документа', value: 'Calibri', action: (v) => { doc.style.font = v; doc.style.head = v; applyPage(); shell.markDirty(); } }, { type: 'select', id: 'basesize', width: 60, options: [9, 10, 10.5, 11, 12, 13, 14, 16], value: '11', tip: 'Основной размер', action: (v) => { doc.style.size = +v; applyPage(); shell.markDirty(); } }] }] }] },
  ] },
  { id: 'review', label: 'Рецензирование', groups: [
    { label: 'Правописание', items: [big('spell', 'spell', 'Орфография', () => { const on = editor.spellcheck = !editor.spellcheck; settings.set('spellcheck', on); shell.ribbon.set('spell', { on }); editor.blur(); editor.focus(); toast(on ? 'Проверка орфографии включена' : 'Проверка орфографии выключена'); }), big('wc', 'word-count', 'Статистика', wordCount)] },
    { label: 'Примечания', items: [big('newcmt', 'comment-plus', 'Создать примечание', addComment, { key: 'Ctrl+Alt+M' }), { type: 'col', items: [b('prevc', 'chevu', 'Предыдущее', () => stepComment(-1), { showLabel: true }), b('nextc', 'chevd', 'Следующее', () => stepComment(1), { showLabel: true }), b('delc', 'trash', 'Удалить все', async () => { if (await confirmDlg('Примечания', 'Удалить все примечания?', 'Удалить', true)) { [...doc.comments].forEach((c) => deleteComment(c.id)); } }, { showLabel: true })] }, big('cmtpane', 'comment', 'Область примечаний', () => togglePane('cmt'))] },
  ] },
  { id: 'view', label: 'Вид', groups: [
    { label: 'Показ', items: [big('navpane', 'sidebar', 'Навигация', () => togglePane('nav')), big('focusmode', 'focus', 'Режим фокуса', () => focusMode(true))] },
    { label: 'Масштаб', items: [big('z100', 'search', '100%', () => setZoom(1)), big('zw', 'maximize', 'По ширине', fitWidth), { type: 'col', items: [b('zin', 'zoom-in', 'Увеличить', () => setZoom(zoom + 0.1), { showLabel: true }), b('zout', 'zoom-out', 'Уменьшить', () => setZoom(zoom - 0.1), { showLabel: true })] }] },
  ] },
  { id: 'table', label: 'Макет таблицы', contextual: true, groups: [
    { label: 'Строки и столбцы', items: [{ type: 'col', items: [b('ra', 'insert-row', 'Вставить сверху', () => tableOp('rowAbove'), { showLabel: true }), b('rb', 'insert-row', 'Вставить снизу', () => tableOp('rowBelow'), { showLabel: true })] }, { type: 'col', items: [b('cl', 'insert-col', 'Вставить слева', () => tableOp('colLeft'), { showLabel: true }), b('cr', 'insert-col', 'Вставить справа', () => tableOp('colRight'), { showLabel: true })] },
      big('tdel', 'trash', 'Удалить', null, { menu: () => [{ label: 'Удалить строку', action: () => tableOp('delRow') }, { label: 'Удалить столбец', action: () => tableOp('delCol') }, { label: 'Удалить таблицу', danger: true, action: () => tableOp('delTable') }] })] },
    { label: 'Объединение', items: [{ type: 'col', items: [b('merge', 'merge', 'Объединить с соседней', () => tableOp('merge'), { showLabel: true }), b('split', 'columns', 'Разделить ячейку', () => tableOp('split'), { showLabel: true }), b('even', 'dist-h', 'Выровнять ширину', () => tableOp('even'), { showLabel: true })] }] },
    { label: 'Стиль таблицы', items: [{ type: 'col', items: [b('thdr', 'heading', 'Строка заголовка', () => tableOp('header'), { showLabel: true }), b('tband', 'menu', 'Чередование строк', () => tableOp('banded'), { showLabel: true }), b('tbord', 'borders', 'Границы', () => tableOp('borders'), { showLabel: true })] },
      { type: 'color', id: 'shade', icon: 'fill-color', tip: 'Заливка ячейки', value: '#dbe7ff', allowNone: true, action: cellShade }] },
  ] },
  { id: 'picture', label: 'Рисунок', contextual: true, groups: [
    { label: 'Обтекание', items: [big('wi', 'align-left', 'В тексте', () => imgWrap(null)), big('wl', 'align-left', 'Слева', () => imgWrap('fl')), big('wc2', 'align-center', 'По центру', () => imgWrap('ctr')), big('wr', 'align-right', 'Справа', () => imgWrap('fr'))] },
    { label: 'Размер', items: [{ type: 'col', items: [b('w25', 'minus', '25%', () => imgWidth(0.25), { showLabel: true }), b('w50', 'minus', '50%', () => imgWidth(0.5), { showLabel: true }), b('w100', 'maximize', '100%', () => imgWidth(1), { showLabel: true })] }] },
    { label: 'Изображение', items: [big('alt', 'edit', 'Замещающий текст', async () => { if (!selImg) return; const t = await promptDlg('Замещающий текст', 'Описание изображения', selImg.alt); if (t != null) { selImg.alt = t; changed(); } }), big('imgdel', 'trash', 'Удалить', () => { selImg?.remove(); selectImage(null); changed(); })] },
  ] },
];
/* extra ribbon controls */
{
  const tab = (id) => ribbon.find((t) => t.id === id);
  const grp = (tid, label) => tab(tid).groups.find((g) => g.label === label);
  grp('home', 'Шрифт').items[0].items[0].items.push(b('fontdlg', 'settings', 'Шрифт: все параметры…', fontDialog, { key: 'Ctrl+D' }));
  grp('home', 'Абзац').items[0].items[1].items.push(b('paradlg', 'settings', 'Абзац: отступы и интервалы…', paragraphDialog));
  grp('insert', 'Иллюстрации').items.push(big('shapes', 'shapes', 'Фигуры', null, { menu: () => Object.entries(SHAPES).map(([k, s]) => ({ label: s.n, action: () => insertShape(k) })) }), { type: 'color', id: 'shapecolor', icon: 'fill-color', tip: 'Цвет новых фигур', value: shapeColor, action: (c) => { if (c) shapeColor = c; } });
  grp('insert', 'Ссылки').items.push(big('fn', 'hash', 'Сноска', insertFootnote, { key: 'Ctrl+Alt+F' }));
  grp('insert', 'Текст').items.push({ type: 'col', items: [b('eq', 'sigma', 'Формула', insertEquation, { showLabel: true }), b('dropcap', 'heading', 'Буквица', toggleDropCap, { showLabel: true })] });
  grp('layout', 'Параметры страницы').items.push(big('hyph', 'minus', 'Переносы', () => { doc.hyphens = !doc.hyphens; applyPage(); shell.markDirty(); }));
  tab('design').groups.push({ label: 'Фон страницы', items: [big('wm', 'type', 'Подложка', null, { menu: () => [...['ЧЕРНОВИК', 'ОБРАЗЕЦ', 'КОНФИДЕНЦИАЛЬНО', 'КОПИЯ', 'НЕ КОПИРОВАТЬ'].map((t) => ({ label: t, checked: doc.watermark === t, action: () => setWatermark(t) })), { label: 'Свой текст…', action: async () => { const t = await promptDlg('Подложка', 'Текст подложки', doc.watermark || ''); if (t != null) setWatermark(t.trim()); } }, '-', { label: 'Удалить подложку', action: () => setWatermark('') }] })] });
  tab('review').groups.splice(1, 0, { label: 'Запись исправлений', items: [
    big('track', 'edit', 'Исправления', () => setTrack(!doc.track), { key: 'Ctrl+Shift+E' }),
    { type: 'col', items: [b('acc', 'check', 'Принять', null, { showLabel: true, menu: () => [{ label: 'Принять это исправление', action: () => trkCurrent(true) }, { label: 'Принять все исправления', action: () => trkAll(true) }] }), b('rej', 'close', 'Отклонить', null, { showLabel: true, menu: () => [{ label: 'Отклонить это исправление', action: () => trkCurrent(false) }, { label: 'Отклонить все исправления', action: () => trkAll(false) }] }), b('trkview', 'eye', 'Показ', null, { showLabel: true, menu: () => [{ label: 'Все исправления', checked: !editor.classList.contains('hide-trk'), action: () => editor.classList.remove('hide-trk') }, { label: 'Итоговый вариант', checked: editor.classList.contains('hide-trk'), action: () => editor.classList.add('hide-trk') }] })] },
    { type: 'col', items: [b('trkprev', 'chevu', 'Предыдущее', () => trkStep(-1), { showLabel: true }), b('trknext', 'chevd', 'Следующее', () => trkStep(1), { showLabel: true })] }] });
  tab('review').groups.push({ label: 'Защита', items: [big('ro', 'lock', 'Ограничить редактирование', () => setReadonly(!doc.readonly))] });
  grp('picture', 'Изображение').items.push(big('rot', 'rotate', 'Повернуть', imgRotate), big('imgst', 'palette', 'Стиль', null, { menu: () => [['frame', 'Рамка'], ['shadow', 'Тень'], ['round', 'Скруглённые углы'], ['circle', 'Круг'], ['', 'Без стиля']].map(([k, n]) => ({ label: n, action: () => imgStyle(k) })) }));
}
function pgn(align) { doc.footer.pageNumbers = true; doc.footer.align = align; doc.footer.show = true; applyPage(); syncRibbon(); shell.markDirty(); toast('Номера страниц появятся при печати и в PDF/DOCX'); }
function focusMode(on) {
  document.body.classList.toggle('focus-mode', on);
  $('.focus-exit')?.remove();
  if (on) document.body.append(h('button', { class: 'btn focus-exit', onclick: () => focusMode(false), html: icon('close') + 'Выйти из режима фокуса' }));
}

/* ======================= shell ======================= */
const shell = createShell({
  app: 'docs', ribbon, defaultTitle: 'Документ',
  serialize, load, newDoc, templates: TEMPLATES.docs,
  isEmpty: () => !editor.textContent.trim() && !editor.querySelector('img,table'),
  undo: () => cmd('undo'), redo: () => cmd('redo'),
  focusEditor: () => restoreSel(),
  onSearch: () => openFind(false),
  print: () => doPrint(false),
  stats: () => { const t = editor.innerText; return [['Страниц', pageCount], ['Слов', (t.match(/[\p{L}\p{N}]+/gu) || []).length], ['Знаков', t.length], ['Размер страницы', `${doc.page.size}${doc.page.landscape ? ', альбомная' : ''}`]]; },
  exporters: [
    { label: 'Документ Word (.docx)', desc: 'Совместим с Microsoft Word, LibreOffice, Google Docs', icon: 'file', run: async () => shell.export(shell.title + '.docx', await exportDocx(editor, { ...doc, title: shell.title, author: profile().name }), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') },
    { label: 'PDF', desc: native ? 'Сохранить как PDF с колонтитулами' : 'Через диалог печати → «Сохранить как PDF»', icon: 'print', run: () => doPrint(true) },
    { label: 'Веб-страница (.html)', desc: 'Один файл со стилями и изображениями', icon: 'globe', run: () => shell.export(shell.title + '.html', standaloneHTML(), 'text/html') },
    { label: 'Markdown (.md)', desc: 'Для заметок, GitHub, вики', icon: 'code', run: () => shell.export(shell.title + '.md', htmlToMd(finalClone()), 'text/markdown') },
    { label: 'Обычный текст (.txt)', desc: 'Только текст без форматирования', icon: 'type', run: () => shell.export(shell.title + '.txt', finalText(), 'text/plain') },
  ],
});
shell.workspace.append(navPane, scroll, cmtPane);
scroll.prepend(roBanner);
shell.statusbar.append(sbPage, sbWords, sbLang, h('span', { class: 'grow' }),
  h('button', { class: 'sb-btn hide-m', 'data-tip': 'Режим фокуса', html: icon('focus'), onclick: () => focusMode(true) }),
  h('button', { class: 'sb-btn', html: icon('minus'), onclick: () => setZoom(zoom - 0.1) }), zoomRange, h('button', { class: 'sb-btn', html: icon('plus'), onclick: () => setZoom(zoom + 0.1) }), zoomLbl);
shell.ribbon.set('spell', { on: settings.get('spellcheck') });

document.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey; const k = e.key.toLowerCase();
  if (mod && k === 'f') { e.preventDefault(); openFind(false); }
  else if (mod && k === 'h') { e.preventDefault(); openFind(true); }
  else if (mod && e.shiftKey && k === 'e') { e.preventDefault(); setTrack(!doc.track); }
  else if (mod && e.altKey && k === 'f') { e.preventDefault(); insertFootnote(); }
  else if (mod && !e.shiftKey && k === 'd' && (editor.contains(document.activeElement) || document.activeElement === editor)) { e.preventDefault(); fontDialog(); }
  else if (!editor.contains(document.activeElement) && document.activeElement !== editor) return;
  else if (mod && k === 'e') { e.preventDefault(); cmd('justifyCenter'); }
  else if (mod && k === 'l' && !e.shiftKey) { e.preventDefault(); cmd('justifyLeft'); }
  else if (mod && k === 'l' && e.shiftKey) { e.preventDefault(); cmd('insertUnorderedList'); }
  else if (mod && k === 'r') { e.preventDefault(); cmd('justifyRight'); }
  else if (mod && k === 'j') { e.preventDefault(); cmd('justifyFull'); }
  else if (mod && k === 'k') { e.preventDefault(); insertLink(); }
  else if (mod && e.key === 'Enter') { e.preventDefault(); insertHTML('<div class="page-break" contenteditable="false"></div><p><br></p>'); }
  else if (mod && e.altKey && ['1', '2', '3'].includes(e.key)) { e.preventDefault(); setBlockStyle('h' + e.key); }
  else if (mod && e.altKey && k === 'm') { e.preventDefault(); addComment(); }
  else if (mod && (e.key === ']' || e.key === '>')) { e.preventDefault(); growFont(1); }
  else if (mod && (e.key === '[' || e.key === '<')) { e.preventDefault(); growFont(-1); }
  else if (mod && e.key === '=') { e.preventDefault(); cmd(e.shiftKey ? 'superscript' : 'subscript'); }
  else if (mod && k === 'z' && !e.shiftKey) { /* native undo */ }
  else if (e.key === 'Escape' && document.body.classList.contains('focus-mode')) focusMode(false);
});
scroll.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); setZoom(zoom * (e.deltaY < 0 ? 1.1 : 0.9)); } }, { passive: false });
window.addEventListener('resize', debounce(() => { if (innerWidth < 760) setZoom(1); placeHandle(); }, 100));
editor.addEventListener('focus', () => { if (selImg && document.activeElement === editor) placeHandle(); });

shell.boot().then(() => { applyPage(); updateStatus(); if (innerWidth > 1100 && scroll.clientWidth < page.offsetWidth + 48) fitWidth(); });
