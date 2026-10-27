// Rabbit Desk — infinite whiteboard with real-time collaboration (direct IP, LAN, server)
import { createShell } from '../core/shell.js';
import { icon } from '../core/icons.js';
import { h, $, $$, esc, toast, dialog, menu, popover, closeMenus, colorPicker, promptDlg, pickImage, debounce, clamp } from '../core/ui.js';
import { profile, settings } from '../core/store.js';
import { accounts, boards as cloudBoards, login } from '../core/cloud.js';
import { native, extOf } from '../core/platform.js';
import { TEMPLATES } from '../core/templates.js';
import { loginDialog } from '../core/shell.js';

const uid = () => Math.random().toString(36).slice(2, 10);
const STICKY = ['#fff3a3', '#ffd6a5', '#ffc4d6', '#c8f0c8', '#bde0fe', '#e0c3fc', '#ffffff', '#dcdcdc'];
const INK = ['#1f1f1f', '#e5484d', '#f76b15', '#30a46c', '#0090ff', '#8e4ec6'];
let objs = new Map(); // id -> object
let view = { x: 0, y: 0, z: 1 };
let tool = 'select'; let sel = new Set(); let editingId = null;
let clock = 0; const me = profile(); const myId = uid();

/* ---------- DOM ---------- */
const board = h('div', { class: 'board' });
const world = h('div', { class: 'world' });
const ink = document.createElementNS('http://www.w3.org/2000/svg', 'svg'); ink.classList.add('ink'); ink.setAttribute('width', '1'); ink.setAttribute('height', '1');
world.append(ink); board.append(world);
const cursorsLayer = h('div'); board.append(cursorsLayer);
const TOOLS = [['select', 'cursor', 'Выбор (V)'], ['hand', 'hand', 'Рука (H)'], ['sticky', 'sticky', 'Стикер (S)'], ['text', 'type', 'Текст (T)'], ['shape', 'shapes', 'Фигура (R)'], ['connector', 'connector', 'Связь между объектами (C)'], ['line', 'arrow-line', 'Линия или стрелка (L)'], ['pen', 'pen', 'Карандаш (P)'], ['marker', 'marker', 'Маркер (M)'], ['eraser', 'eraser', 'Ластик (E)'], ['frame', 'frame', 'Секция (F)'], ['comment', 'comment', 'Комментарий (K)']];
const dock = h('div', { class: 'dock' });
for (const [id, ic, tip] of TOOLS) { dock.append(h('button', { 'data-tool': id, 'data-tip': tip, html: icon(ic), onclick: (e) => (id === 'shape' ? shapeMenu(e.currentTarget) : setTool(id)) })); if (id === 'hand' || id === 'line' || id === 'eraser') dock.append(h('span', { class: 'sep' })); }
function shapeMenu(anchor) { menu(anchor, Object.entries(SHAPE_NAMES).map(([k, [n]]) => ({ label: n, checked: lastShape === k, action: () => { lastShape = k; setTool('shape'); } }))); }
dock.append(h('button', { 'data-tip': 'Изображение', html: icon('image'), onclick: insertImage }), h('button', { 'data-tip': 'Реакция', html: icon('stamp'), onclick: (e) => menu(e.currentTarget, ['👍', '❤️', '🎉', '🔥', '⭐', '✅', '❓', '💡'].map((s) => ({ label: s, action: () => put({ id: uid(), type: 'stamp', x: center().x - 24, y: center().y - 24, w: 48, h: 48, text: s }) }))) }));
const zoomLbl = h('button', { onclick: () => { view.z = 1; applyView(); } }, '100%');
const zoomCtl = h('div', { class: 'zoom-ctl' }, h('button', { html: icon('minus'), onclick: () => zoomAt(0.85) }), zoomLbl, h('button', { html: icon('plus'), onclick: () => zoomAt(1.18) }), h('button', { html: icon('fit'), 'data-tip': 'Показать всё', onclick: fitAll }));
board.append(dock, zoomCtl);
let ctxBar = null;
function setTool(t) { tool = t; $$('.dock button[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === t)); board.style.cursor = t === 'hand' ? 'grab' : t === 'select' ? 'default' : 'crosshair'; }

/* ---------- view ---------- */
function applyView() { world.style.transform = `translate(${view.x}px,${view.y}px) scale(${view.z})`; board.style.backgroundSize = `${24 * view.z}px ${24 * view.z}px`; board.style.backgroundPosition = `${view.x}px ${view.y}px`; zoomLbl.textContent = Math.round(view.z * 100) + '%'; placeCtx(); drawMini(); }
function toWorld(cx, cy) { const r = board.getBoundingClientRect(); return { x: (cx - r.left - view.x) / view.z, y: (cy - r.top - view.y) / view.z }; }
function center() { return toWorld(board.getBoundingClientRect().left + board.clientWidth / 2, board.getBoundingClientRect().top + board.clientHeight / 2); }
function zoomAt(f, cx, cy) { const r = board.getBoundingClientRect(); cx ??= r.left + r.width / 2; cy ??= r.top + r.height / 2; const p = toWorld(cx, cy); view.z = clamp(view.z * f, 0.1, 5); view.x = cx - r.left - p.x * view.z; view.y = cy - r.top - p.y * view.z; applyView(); }
function fitAll() { const list = [...objs.values()].filter((o) => !o.deleted && o.type !== 'connector'); if (!list.length) { view = { x: board.clientWidth / 2, y: board.clientHeight / 2, z: 1 }; return applyView(); } const b = bbox(list); const z = clamp(Math.min((board.clientWidth - 120) / (b.w || 1), (board.clientHeight - 160) / (b.h || 1)), 0.1, 2); view = { z, x: board.clientWidth / 2 - (b.x + b.w / 2) * z, y: board.clientHeight / 2 - (b.y + b.h / 2) * z - 20 }; applyView(); }
function bbox(list) { let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity; for (const o of list) { const b = objBox(o); x1 = Math.min(x1, b.x); y1 = Math.min(y1, b.y); x2 = Math.max(x2, b.x + b.w); y2 = Math.max(y2, b.y + b.h); } return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }; }
function objBox(o) { if (o.type === 'path') { const xs = o.pts.filter((_, i) => i % 2 === 0), ys = o.pts.filter((_, i) => i % 2); return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) }; } return { x: o.x, y: o.y, w: o.w || 0, h: o.h || 0 }; }

/* ---------- model ops (CRDT-lite: last-writer-wins per object) ---------- */
const undoSt = []; const redoSt = [];
function put(o, { local = true, record = true } = {}) {
  const prev = objs.get(o.id);
  if (local) { o.v = Date.now() * 1000 + (clock++ % 1000); o.by = myId; if (record) { undoSt.push({ before: prev ? structuredClone(prev) : { id: o.id, deleted: true }, after: structuredClone(o) }); redoSt.length = 0; } }
  else if (prev && prev.v > o.v) return;
  objs.set(o.id, o);
  renderObj(o);
  if (local) { net.send({ t: 'ops', ops: [o] }); shell?.markDirty(); }
}
function del(id) { const o = objs.get(id); if (!o) return; put({ ...o, deleted: true }); for (const c of objs.values()) if (c.type === 'connector' && !c.deleted && (c.from === id || c.to === id)) put({ ...c, deleted: true }); }
function undo() { const e = undoSt.pop(); if (!e) return; redoSt.push(e); put(structuredClone(e.before), { record: false }); }
function redo() { const e = redoSt.pop(); if (!e) return; undoSt.push(e); put(structuredClone(e.after), { record: false }); }

/* ---------- rendering ---------- */
const nodes = new Map();
const SHAPE_SVG = { rect: '<rect x="1" y="1" width="98" height="98" rx="6"/>', round: '<rect x="1" y="1" width="98" height="98" rx="24"/>', ellipse: '<ellipse cx="50" cy="50" rx="49" ry="49"/>', diamond: '<path d="M50 1L99 50L50 99L1 50Z"/>', triangle: '<path d="M50 2L98 98H2Z"/>', hexagon: '<path d="M25 2H75L98 50L75 98H25L2 50Z"/>', star: '<path d="M50 2L62 36H98L69 58L80 94L50 72L20 94L31 58L2 36H38Z"/>', arrow: '<path d="M2 30H62V4L98 50L62 96V70H2Z"/>', parallelogram: '<path d="M22 2H98L78 98H2Z"/>', bubble: '<path d="M8 2H92Q98 2 98 8V66Q98 74 92 74H40L20 96V74H8Q2 74 2 66V8Q2 2 8 2Z"/>', cross: '<path d="M35 2H65V35H98V65H65V98H35V65H2V35H35Z"/>' };
const SHAPE_NAMES = { rect: ['Прямоугольник', '#bde0fe'], round: ['Скруглённый', '#c8f0c8'], ellipse: ['Эллипс', '#c8f0c8'], diamond: ['Ромб', '#ffd6a5'], triangle: ['Треугольник', '#ffc4d6'], hexagon: ['Шестиугольник', '#e0c3fc'], star: ['Звезда', '#fff3a3'], arrow: ['Стрелка', '#bde0fe'], parallelogram: ['Параллелограмм', '#ffd6a5'], bubble: ['Выноска', '#ffffff'], cross: ['Крест', '#ffc4d6'] };
let lastShape = 'rect';
const markers = new Set();
// one arrowhead marker per stroke colour
function arrowMarker(color) {
  const id = 'arr-' + String(color || '#555').replace(/[^a-z0-9]/gi, '');
  if (!markers.has(id)) {
    markers.add(id);
    const m = document.createElementNS('http://www.w3.org/2000/svg', 'marker'); m.id = id;
    for (const [k, v] of Object.entries({ viewBox: '0 0 10 10', refX: 8, refY: 5, markerWidth: 6, markerHeight: 6, orient: 'auto-start-reverse' })) m.setAttribute(k, v);
    m.innerHTML = '<path d="M0 0L10 5L0 10z" fill="' + color + '"/>'; ink.querySelector('defs').append(m);
  }
  return 'url(#' + id + ')';
}
function renderObj(o) {
  let n = nodes.get(o.id);
  if (o.deleted) { n?.remove(); nodes.delete(o.id); sel.delete(o.id); renderConnectorsFor(o.id); return; }
  if (o.type === 'path' || o.type === 'connector') {
    if (!n) { n = document.createElementNS('http://www.w3.org/2000/svg', 'path'); n.dataset.id = o.id; n.classList.add('hit'); o.type === 'connector' ? ink.prepend(n) : ink.append(n); nodes.set(o.id, n); }
    if (o.type === 'path') { n.setAttribute('d', pathD(o.pts)); n.setAttribute('stroke', o.color); n.setAttribute('stroke-width', o.width); n.setAttribute('stroke-opacity', o.marker ? 0.4 : 1); n.setAttribute('marker-end', o.arrow ? arrowMarker(o.color) : ''); n.setAttribute('stroke-dasharray', o.dash ? '10 8' : ''); }
    else { const d = connD(o); n.setAttribute('d', d || ''); n.setAttribute('stroke', o.color || '#555'); n.setAttribute('stroke-width', 2.5); n.setAttribute('marker-end', arrowMarker(o.color || '#555')); n.setAttribute('stroke-dasharray', o.dash ? '10 8' : ''); }
    n.setAttribute('fill', 'none'); n.setAttribute('stroke-linecap', 'round'); n.setAttribute('stroke-linejoin', 'round');
    n.style.outline = ''; n.setAttribute('filter', sel.has(o.id) ? 'drop-shadow(0 0 3px #7b4fd6)' : '');
    return;
  }
  if (!n) { n = h('div', { class: 'obj ' + o.type, 'data-id': o.id }); nodes.set(o.id, n); if (o.type === 'frame') world.insertBefore(n, ink); else world.append(n); }
  Object.assign(n.style, { left: o.x + 'px', top: o.y + 'px', width: o.w ? o.w + 'px' : '', height: o.h ? o.h + 'px' : '', zIndex: o.type === 'frame' ? 0 : 1 });
  n.className = 'obj ' + o.type + (sel.has(o.id) ? ' sel' : '');
  if (editingId === o.id) return;
  if (o.type === 'sticky') { n.style.background = o.color; n.innerHTML = `<div class="tx">${esc(o.text || '')}</div>${o.author ? `<div class="by">${esc(o.author)}</div>` : ''}`; n.style.fontSize = (o.fs || 18) + 'px'; }
  else if (o.type === 'text') { n.innerHTML = `<div class="tx">${esc(o.text || '')}</div>`; n.style.fontSize = (o.fs || 20) + 'px'; n.style.color = o.color || ''; n.style.height = ''; }
  else if (o.type === 'shape') { n.innerHTML = `<svg viewBox="0 0 100 100" preserveAspectRatio="none"><g fill="${o.color}" stroke="${o.stroke || '#333'}" stroke-width="2" vector-effect="non-scaling-stroke">${SHAPE_SVG[o.shape]}</g></svg><div class="tx">${esc(o.text || '')}</div>`; n.querySelectorAll('rect,ellipse,path').forEach((p) => p.setAttribute('vector-effect', 'non-scaling-stroke')); }
  else if (o.type === 'frame') n.innerHTML = `<div class="ft">${esc(o.text || 'Секция')}</div>`;
  else if (o.type === 'image') n.innerHTML = `<img src="${o.src}" alt="">`;
  else if (o.type === 'stamp') n.textContent = o.text;
  else if (o.type === 'comment') { const last = o.thread?.[o.thread.length - 1]; n.innerHTML = `<span class="pin" style="background:${o.color || '#7b4fd6'}">${esc((o.author || '?')[0].toUpperCase())}</span>${(o.thread?.length || 0) > 1 ? `<span class="cnt">${o.thread.length}</span>` : ''}`; n.title = last ? `${last.a}: ${last.text}` : ''; }
  const tx = n.querySelector(':scope > .tx'); if (tx) { tx.style.fontWeight = o.bold ? '700' : ''; tx.style.textAlign = o.align || ''; tx.style.fontStyle = o.italic ? 'italic' : ''; }
  const votes = Object.values(o.votes || {}).filter(Boolean).length; if (votes) n.append(h('div', { class: 'votes' + (o.votes?.[me.name] ? ' mine' : '') }, '👍 ' + votes));
  if (o.type !== 'comment') n.append(h('div', { class: 'rs' }));
  if (['sticky', 'shape', 'text', 'image', 'frame'].includes(o.type)) for (const [ax, ay] of [[50, 0], [100, 50], [50, 100], [0, 50]]) n.append(h('div', { class: 'anc', style: { left: ax + '%', top: ay + '%' } }));
  renderConnectorsFor(o.id);
}
function renderConnectorsFor(id) { for (const c of objs.values()) if (c.type === 'connector' && !c.deleted && (c.from === id || c.to === id)) { const n = nodes.get(c.id); if (n) n.setAttribute('d', connD(c) || ''); } }
function pathD(p) { if (p.length < 4) return `M${p[0]} ${p[1]}l0.1 0`; let d = `M${p[0]} ${p[1]}`; for (let i = 2; i < p.length - 2; i += 2) { const mx = (p[i] + p[i + 2]) / 2, my = (p[i + 1] + p[i + 3]) / 2; d += `Q${p[i]} ${p[i + 1]} ${mx} ${my}`; } return d + `L${p[p.length - 2]} ${p[p.length - 1]}`; }
function edgePoint(o, tx, ty) { const b = objBox(o); const cx = b.x + b.w / 2, cy = b.y + b.h / 2; const dx = tx - cx, dy = ty - cy; const s = Math.min(Math.abs((b.w / 2 + 6) / (dx || 1e-6)), Math.abs((b.h / 2 + 6) / (dy || 1e-6))); return [cx + dx * s, cy + dy * s]; }
function connD(c) {
  const a = objs.get(c.from), b = c.to ? objs.get(c.to) : null; if (!a || a.deleted || (c.to && (!b || b.deleted))) return null;
  const bt = b ? [objBox(b).x + objBox(b).w / 2, objBox(b).y + objBox(b).h / 2] : [c.tx, c.ty];
  const [x1, y1] = edgePoint(a, bt[0], bt[1]); const [x2, y2] = b ? edgePoint(b, x1, y1) : bt;
  const mx = (x1 + x2) / 2; return Math.abs(x2 - x1) > Math.abs(y2 - y1) ? `M${x1} ${y1}C${mx} ${y1} ${mx} ${y2} ${x2} ${y2}` : `M${x1} ${y1}C${x1} ${(y1 + y2) / 2} ${x2} ${(y1 + y2) / 2} ${x2} ${y2}`;
}
ink.innerHTML = '<defs><marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="#555"/></marker></defs>';
function renderAll() { for (const n of nodes.values()) n.remove(); nodes.clear(); const list = [...objs.values()].sort((a, b) => (a.type === 'frame' ? -1 : 0) - (b.type === 'frame' ? -1 : 0)); for (const o of list) renderObj(o); }
function setSel(ids) { const old = sel; sel = new Set(ids); for (const id of new Set([...old, ...sel])) { const o = objs.get(id); if (o) renderObj(o); } placeCtx(); }

/* ---------- context toolbar ---------- */
function placeCtx() {
  ctxBar?.remove(); ctxBar = null;
  const list = [...sel].map((id) => objs.get(id)).filter((o) => o && !o.deleted); if (!list.length || editingId) return;
  const b = bbox(list); const r = board.getBoundingClientRect();
  ctxBar = h('div', { class: 'ctx-bar' });
  const colorable = list.filter((o) => ['sticky', 'shape', 'path', 'text', 'connector'].includes(o.type));
  const palette = list.every((o) => o.type === 'sticky' || o.type === 'shape') ? STICKY : INK;
  if (colorable.length) for (const c of palette) ctxBar.append(h('button', { class: `sw ${list[0].color === c ? 'on' : ''}`, style: { background: c }, onclick: () => { for (const o of colorable) put({ ...o, color: c }); placeCtx(); } }));
  if (list.some((o) => o.text != null && o.type !== 'stamp')) ctxBar.append(h('button', { class: 'b', html: icon('text-size'), 'data-tip': 'Размер текста', onclick: (e) => menu(e.currentTarget, [14, 18, 24, 32, 48, 64].map((s) => ({ label: s + ' px', action: () => list.forEach((o) => put({ ...o, fs: s })) }))) }));
  const texty = list.filter((o) => ['sticky', 'text', 'shape'].includes(o.type));
  if (texty.length) ctxBar.append(
    h('button', { class: 'b' + (texty[0].bold ? ' on' : ''), html: icon('bold'), 'data-tip': 'Полужирный', onclick: () => { const v = !texty[0].bold; texty.forEach((o) => put({ ...o, bold: v })); placeCtx(); } }),
    h('button', { class: 'b', html: icon(texty[0].align === 'left' ? 'align-left' : texty[0].align === 'right' ? 'align-right' : 'align-center'), 'data-tip': 'Выравнивание текста', onclick: (e) => menu(e.currentTarget, [['left', 'По левому краю', 'align-left'], ['center', 'По центру', 'align-center'], ['right', 'По правому краю', 'align-right']].map(([a, l, ic]) => ({ label: l, icon: ic, action: () => texty.forEach((o) => put({ ...o, align: a })) }))) }));
  if (list.length === 1 && list[0].type === 'shape') ctxBar.append(h('button', { class: 'b', html: icon('shapes'), 'data-tip': 'Сменить фигуру', onclick: (e) => menu(e.currentTarget, Object.entries(SHAPE_NAMES).map(([k, [nm]]) => ({ label: nm, checked: list[0].shape === k, action: () => put({ ...list[0], shape: k }) }))) }));
  const lines = list.filter((o) => (o.type === 'path' && o.pts.length === 4) || o.type === 'connector');
  if (lines.length) ctxBar.append(h('button', { class: 'b', html: icon('arrow-line'), 'data-tip': 'Стрелка на конце', onclick: () => { lines.filter((o) => o.type === 'path').forEach((o) => put({ ...o, arrow: !o.arrow })); placeCtx(); } }), h('button', { class: 'b', html: icon('more-h'), 'data-tip': 'Пунктир', onclick: () => { lines.forEach((o) => put({ ...o, dash: !o.dash })); placeCtx(); } }));
  const stickies = list.filter((o) => o.type === 'sticky');
  if (stickies.length) ctxBar.append(h('button', { class: 'b' + (stickies[0].votes?.[me.name] ? ' on' : ''), 'data-tip': 'Голос «за»', onclick: () => { const mine = !stickies[0].votes?.[me.name]; stickies.forEach((o) => put({ ...o, votes: { ...(o.votes || {}), [me.name]: mine } })); placeCtx(); } }, '👍'));
  if (list.length > 1) ctxBar.append(h('button', { class: 'b', html: icon('align-objects'), 'data-tip': 'Выровнять', onclick: (e) => menu(e.currentTarget, [['l', 'По левому краю'], ['c', 'По центру'], ['r', 'По правому краю'], ['t', 'По верху'], ['m', 'По середине'], ['b', 'По низу'], ['dh', 'Распределить по горизонтали'], ['dv', 'Распределить по вертикали'], ['grid', 'Разложить сеткой']].map(([k, l]) => ({ label: l, action: () => alignSel(k) }))) }));
  ctxBar.append(h('button', { class: 'b', html: icon('copy'), 'data-tip': 'Дублировать', onclick: duplicate }), h('button', { class: 'b', html: icon(list[0].locked ? 'unlock' : 'lock2'), 'data-tip': 'Закрепить', onclick: () => { list.forEach((o) => put({ ...o, locked: !o.locked })); placeCtx(); } }), h('button', { class: 'b', html: icon('trash'), 'data-tip': 'Удалить', onclick: () => { [...sel].forEach(del); setSel([]); } }));
  board.append(ctxBar);
  const x = view.x + (b.x + b.w / 2) * view.z - ctxBar.offsetWidth / 2, y = view.y + b.y * view.z - ctxBar.offsetHeight - 14;
  ctxBar.style.left = clamp(x, 8, board.clientWidth - ctxBar.offsetWidth - 8) + 'px'; ctxBar.style.top = Math.max(8, y) + 'px';
}
function duplicate() { const ids = []; for (const id of sel) { const o = objs.get(id); if (!o || o.type === 'connector') continue; const n = structuredClone(o); n.id = uid(); if (n.type === 'path') n.pts = n.pts.map((v) => v + 30); else { n.x += 30; n.y += 30; } put(n); ids.push(n.id); } setSel(ids); }

function alignSel(kind) {
  const list = [...sel].map((id) => objs.get(id)).filter((o) => o && !o.deleted && o.type !== 'connector' && o.type !== 'path');
  if (list.length < 2) return;
  const b = bbox(list);
  if (kind === 'grid') {
    const cols = Math.ceil(Math.sqrt(list.length)); const cw = Math.max(...list.map((o) => o.w || 0)) + 24, ch = Math.max(...list.map((o) => o.h || 40)) + 24;
    list.sort((p, q) => p.y - q.y || p.x - q.x).forEach((o, i) => put({ ...o, x: Math.round(b.x + (i % cols) * cw), y: Math.round(b.y + Math.floor(i / cols) * ch) }));
  } else if (kind === 'dh' || kind === 'dv') {
    const hor = kind === 'dh'; const s = [...list].sort((p, q) => (hor ? p.x - q.x : p.y - q.y));
    const total = s.reduce((a, o) => a + (hor ? o.w : o.h || 40), 0); const gap = ((hor ? b.w : b.h) - total) / (s.length - 1); let pos = hor ? b.x : b.y;
    for (const o of s) { put(hor ? { ...o, x: Math.round(pos) } : { ...o, y: Math.round(pos) }); pos += (hor ? o.w : o.h || 40) + gap; }
  } else for (const o of list) {
    const w = o.w || 0, hh = o.h || 40;
    put({ ...o, ...(kind === 'l' && { x: b.x }), ...(kind === 'c' && { x: Math.round(b.x + b.w / 2 - w / 2) }), ...(kind === 'r' && { x: b.x + b.w - w }), ...(kind === 't' && { y: b.y }), ...(kind === 'm' && { y: Math.round(b.y + b.h / 2 - hh / 2) }), ...(kind === 'b' && { y: b.y + b.h - hh }) });
  }
  placeCtx();
}

/* ---------- comment threads ---------- */
function openThread(o, isNew = false) {
  closeMenus();
  const wrap = h('div', { class: 'thread' });
  const list = h('div', { class: 'thread-list' });
  const draw = () => { list.innerHTML = ''; for (const m of o.thread || []) list.append(h('div', { class: 'msg' }, h('div', { class: 'who' }, h('span', { class: 'dot', style: { background: m.c } }, (m.a || '?')[0]), h('b', {}, m.a), h('span', { class: 'when' }, new Date(m.t).toLocaleString('ru-RU', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' }))), h('div', { class: 'txt' }, m.text))); };
  draw();
  const inp = h('textarea', { class: 'input', rows: 2, placeholder: isNew ? 'Напишите комментарий…' : 'Ответить…' });
  const send = () => { const t = inp.value.trim(); if (!t) return; o = { ...(objs.get(o.id) || o), thread: [...(objs.get(o.id)?.thread || o.thread || []), { a: me.name, c: me.color, t: Date.now(), text: t }] }; put(o); inp.value = ''; draw(); isNew = false; };
  inp.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); send(); } e.stopPropagation(); });
  wrap.append(list, inp, h('div', { class: 'row', style: { justifyContent: 'flex-end', marginTop: '6px' } },
    !isNew ? h('button', { class: 'btn sm', onclick: () => { closeMenus(); del(o.id); } }, 'Решено') : null,
    h('button', { class: 'btn sm primary', onclick: send }, 'Отправить')));
  const r = board.getBoundingClientRect();
  popover({ x: r.left + view.x + (o.x + 36) * view.z, y: r.top + view.y + o.y * view.z }, wrap);
  setTimeout(() => inp.focus(), 30);
}

/* ---------- shared timer (synced through board meta) ---------- */
let timer = null; let timerTick = null;
const timerPill = h('div', { class: 'timer-pill hidden' });
function applyMeta(meta) { if ('timer' in meta) { timer = meta.timer; runTimer(); } }
function setTimer(sec) { timer = sec ? { end: Date.now() + sec * 1000, total: sec, by: me.name } : null; net.send({ t: 'meta', meta: { timer } }); runTimer(); }
function runTimer() {
  clearInterval(timerTick); timerPill.classList.toggle('hidden', !timer);
  if (!timer) return;
  const tick = () => {
    const left = Math.max(0, Math.round((timer.end - Date.now()) / 1000));
    timerPill.innerHTML = `${icon('timer')}<b>${String(Math.floor(left / 60)).padStart(2, '0')}:${String(left % 60).padStart(2, '0')}</b>`;
    timerPill.classList.toggle('done', left === 0);
    if (left === 0) { clearInterval(timerTick); beep(); toast('Время вышло!', 'info', 4000); }
  };
  tick(); timerTick = setInterval(tick, 500);
}
function beep() { try { const a = new AudioContext(); [0, 0.35, 0.7].forEach((t) => { const o = a.createOscillator(), g = a.createGain(); o.frequency.value = 880; o.connect(g); g.connect(a.destination); g.gain.setValueAtTime(0.0001, a.currentTime + t); g.gain.exponentialRampToValueAtTime(0.3, a.currentTime + t + 0.02); g.gain.exponentialRampToValueAtTime(0.0001, a.currentTime + t + 0.25); o.start(a.currentTime + t); o.stop(a.currentTime + t + 0.3); }); } catch { /* no audio */ } }
function timerMenu(anchor) { menu(anchor, [{ title: 'Таймер для всех участников' }, ...[1, 3, 5, 10, 15, 30].map((m) => ({ label: `${m} мин`, icon: 'timer', action: () => setTimer(m * 60) })), { label: 'Своё время…', action: async () => { const v = await promptDlg('Таймер', 'Минут', '7', { type: 'number' }); if (+v > 0) setTimer(Math.round(+v * 60)); } }, '-', { label: 'Остановить', icon: 'stop', disabled: !timer, action: () => setTimer(0) }]); }
timerPill.onclick = (e) => timerMenu(e.currentTarget);

/* ---------- minimap ---------- */
const mini = h('canvas', { class: 'minimap hidden', width: 200, height: 130 });
let miniOn = false, miniRaf = 0;
function drawMini() {
  if (!miniOn) return; cancelAnimationFrame(miniRaf);
  miniRaf = requestAnimationFrame(() => {
    const g = mini.getContext('2d'); const W = mini.width, H = mini.height; g.clearRect(0, 0, W, H);
    const list = [...objs.values()].filter((o) => !o.deleted && o.type !== 'connector');
    const vw = { x: -view.x / view.z, y: -view.y / view.z, w: board.clientWidth / view.z, h: board.clientHeight / view.z };
    const b = list.length ? bbox(list) : vw; const x1 = Math.min(b.x, vw.x), y1 = Math.min(b.y, vw.y), x2 = Math.max(b.x + b.w, vw.x + vw.w), y2 = Math.max(b.y + b.h, vw.y + vw.h);
    const k = Math.min((W - 10) / (x2 - x1 || 1), (H - 10) / (y2 - y1 || 1)); const ox = 5 - x1 * k, oy = 5 - y1 * k;
    mini.dataset.t = JSON.stringify({ k, ox, oy });
    for (const o of list) { const bb = objBox(o); g.fillStyle = o.type === 'frame' ? 'rgba(120,120,140,.18)' : o.color || '#999'; g.fillRect(ox + bb.x * k, oy + bb.y * k, Math.max(2, bb.w * k), Math.max(2, (bb.h || 20) * k)); }
    g.strokeStyle = '#7b4fd6'; g.lineWidth = 1.5; g.strokeRect(ox + vw.x * k, oy + vw.y * k, vw.w * k, vw.h * k);
  });
}
mini.addEventListener('pointerdown', (e) => {
  e.stopPropagation();
  const go = (ev) => { const t = JSON.parse(mini.dataset.t || '{}'); const r = mini.getBoundingClientRect(); const wx = (ev.clientX - r.left - t.ox) / t.k, wy = (ev.clientY - r.top - t.oy) / t.k; view.x = board.clientWidth / 2 - wx * view.z; view.y = board.clientHeight / 2 - wy * view.z; applyView(); };
  go(e); mini.setPointerCapture(e.pointerId); mini.onpointermove = (ev) => { if (ev.buttons) go(ev); };
});
function toggleMini() { miniOn = !miniOn; mini.classList.toggle('hidden', !miniOn); drawMini(); }
setInterval(() => drawMini(), 1000);

/* ---------- editing text ---------- */
let finishEdit = null;
function editText(o) {
  if (!['sticky', 'text', 'shape', 'frame'].includes(o.type)) return;
  if (o.type === 'frame') { promptDlg('Название секции', '', o.text || '').then((t) => { if (t != null) put({ ...o, text: t }); }); return; }
  editingId = o.id; placeCtx();
  const n = nodes.get(o.id); const tx = n.querySelector('.tx'); tx.contentEditable = 'plaintext-only'; tx.focus();
  const r = document.createRange(); r.selectNodeContents(tx); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r);
  const done = () => { if (finishEdit !== done) return; finishEdit = null; tx.removeEventListener('blur', done); tx.contentEditable = 'false'; editingId = null; const t = tx.innerText.replace(/\n$/, ''); if (o.type === 'text' && !t.trim()) del(o.id); else if (t !== (o.text || '')) { const next = { ...objs.get(o.id), text: t }; put(next, { record: false }); undoSt.push({ before: structuredClone(o), after: structuredClone(next) }); redoSt.length = 0; } else renderObj(objs.get(o.id)); placeCtx(); };
  finishEdit = done; // explicit finish: blur does not fire when the window itself is unfocused
  tx.addEventListener('blur', done);
  tx.addEventListener('keydown', (e) => { if (e.key === 'Escape' || (e.key === 'Enter' && (e.ctrlKey || e.metaKey))) { e.preventDefault(); done(); tx.blur(); } e.stopPropagation(); });
  // stream text to collaborators while typing (undo records only the final value on blur)
  let lastSent = 0, pendingT = null;
  const stream = () => { lastSent = Date.now(); const cur = objs.get(o.id); if (cur && editingId === o.id) put({ ...cur, text: tx.innerText.replace(/\n$/, '') }, { record: false }); };
  tx.addEventListener('input', () => { clearTimeout(pendingT); if (Date.now() - lastSent > 150) stream(); else pendingT = setTimeout(stream, 150); });
}

/* ---------- pointer ---------- */
const pointers = new Map(); let gesture = null;
board.addEventListener('pointerdown', (e) => {
  if (e.target.closest('.dock,.zoom-ctl,.ctx-bar')) return;
  if (editingId && e.target.closest(`.obj[data-id="${editingId}"]`)) return;
  if (editingId) finishEdit?.();
  closeMenus();
  pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
  board.setPointerCapture(e.pointerId);
  if (pointers.size === 2) { const [a, b] = [...pointers.values()]; gesture = { type: 'pinch', d: Math.hypot(a.x - b.x, a.y - b.y), cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2, view: { ...view } }; return; }
  const p = toWorld(e.clientX, e.clientY);
  const objEl = e.target.closest('.obj'); const pathEl = e.target.closest('path.hit');
  const hitId = objEl?.dataset.id || pathEl?.dataset.id; const hit = hitId ? objs.get(hitId) : null;
  if (tool === 'hand' || e.button === 1 || (e.button === 0 && spaceDown)) { gesture = { type: 'pan', sx: e.clientX, sy: e.clientY, view: { ...view } }; board.style.cursor = 'grabbing'; return; }
  if (e.target.classList.contains('anc') && hit) { const c = { id: uid(), type: 'connector', from: hit.id, to: null, tx: p.x, ty: p.y, color: '#555' }; gesture = { type: 'connect', c }; put(c, { record: false }); return; }
  if (tool === 'select') {
    if (e.target.classList.contains('rs') && hit) { gesture = { type: 'resize', o: structuredClone(hit), p }; return; }
    if (hit && hit.locked) { setSel([hit.id]); return; }
    if (hit) { if (e.shiftKey) setSel(sel.has(hit.id) ? [...sel].filter((x) => x !== hit.id) : [...sel, hit.id]); else if (!sel.has(hit.id)) setSel([hit.id]);
      const moving = [...sel].map((id) => objs.get(id)).filter(Boolean);
      // frames carry contained objects
      for (const f of moving.filter((o) => o.type === 'frame')) for (const o of objs.values()) if (!o.deleted && o.type !== 'frame' && o.type !== 'connector' && !sel.has(o.id)) { const b = objBox(o); if (b.x >= f.x && b.y >= f.y && b.x + b.w <= f.x + f.w && b.y + b.h <= f.y + f.h) moving.push(o); }
      gesture = { type: 'move', p, orig: moving.map((o) => structuredClone(o)), moved: false }; return; }
    setSel([]); gesture = { type: 'marquee', p, el: h('div', { class: 'marq' }) }; board.append(gesture.el); return;
  }
  if (tool === 'eraser') { gesture = { type: 'erase' }; if (hit) del(hit.id); return; }
  if (tool === 'line') { const o = { id: uid(), type: 'path', pts: [Math.round(p.x), Math.round(p.y), Math.round(p.x), Math.round(p.y)], color: lastInk, width: 3, arrow: true }; gesture = { type: 'line', o }; put(o, { record: false }); return; }
  if (tool === 'comment') { if (hit?.type === 'comment') { openThread(hit); return; } const o = { id: uid(), type: 'comment', x: Math.round(p.x - 16), y: Math.round(p.y - 16), w: 32, h: 32, author: me.name, color: me.color, thread: [] }; setTool('select'); openThread(o, true); return; }
  if (tool === 'pen' || tool === 'marker') { const o = { id: uid(), type: 'path', pts: [Math.round(p.x), Math.round(p.y)], color: tool === 'marker' ? '#ffd400' : lastInk, width: tool === 'marker' ? 18 : 3, marker: tool === 'marker' }; gesture = { type: 'draw', o }; put(o, { record: false }); return; }
  if (tool === 'connector' && hit) { const c = { id: uid(), type: 'connector', from: hit.id, to: null, tx: p.x, ty: p.y, color: '#555' }; gesture = { type: 'connect', c }; put(c, { record: false }); return; }
  const BIG = ['ellipse', 'star', 'hexagon', 'triangle', 'diamond', 'cross'];
  const make = { shape: () => ({ type: 'shape', shape: lastShape, w: BIG.includes(lastShape) ? 160 : 180, h: BIG.includes(lastShape) ? 160 : 120, color: SHAPE_NAMES[lastShape][1] }), sticky: () => ({ type: 'sticky', w: 180, h: 180, color: lastSticky, text: '', author: me.name }), text: () => ({ type: 'text', w: 220, text: '', color: '' }), rect: () => ({ type: 'shape', shape: 'rect', w: 180, h: 120, color: '#bde0fe' }), ellipse: () => ({ type: 'shape', shape: 'ellipse', w: 160, h: 160, color: '#c8f0c8' }), diamond: () => ({ type: 'shape', shape: 'diamond', w: 160, h: 160, color: '#ffd6a5' }), frame: () => ({ type: 'frame', w: 640, h: 420, text: 'Секция' }) }[tool];
  if (make) { const o = { id: uid(), ...make() }; o.x = Math.round(p.x - (o.w || 0) / 2); o.y = Math.round(p.y - (o.h || 40) / 2); put(o); setSel([o.id]); setTool('select'); if (o.type !== 'frame') setTimeout(() => editText(o), 0); }
});
let lastSticky = STICKY[0], lastInk = INK[0], spaceDown = false;
board.addEventListener('pointermove', (e) => {
  const p = toWorld(e.clientX, e.clientY);
  sendCursor(p);
  if (!pointers.has(e.pointerId)) return; pointers.set(e.pointerId, { x: e.clientX, y: e.clientY });
  if (!gesture) return;
  if (gesture.type === 'pinch' && pointers.size >= 2) { const [a, b] = [...pointers.values()]; const d = Math.hypot(a.x - b.x, a.y - b.y); const r = board.getBoundingClientRect(); const z = clamp(gesture.view.z * (d / gesture.d), 0.1, 5); const cx = (a.x + b.x) / 2 - r.left, cy = (a.y + b.y) / 2 - r.top; const wx = (gesture.cx - r.left - gesture.view.x) / gesture.view.z, wy = (gesture.cy - r.top - gesture.view.y) / gesture.view.z; view = { z, x: cx - wx * z, y: cy - wy * z }; applyView(); return; }
  if (gesture.type === 'pan') { view.x = gesture.view.x + e.clientX - gesture.sx; view.y = gesture.view.y + e.clientY - gesture.sy; applyView(); return; }
  if (gesture.type === 'move') { const dx = p.x - gesture.p.x, dy = p.y - gesture.p.y; if (Math.hypot(dx, dy) * view.z > 2) gesture.moved = true; for (const o of gesture.orig) { const n = objs.get(o.id); if (o.type === 'path') n.pts = o.pts.map((v, i) => v + (i % 2 ? dy : dx)); else { n.x = Math.round(o.x + dx); n.y = Math.round(o.y + dy); } renderObj(n); } placeCtx(); throttledSend(gesture.orig.map((o) => objs.get(o.id))); return; }
  if (gesture.type === 'resize') { const o = objs.get(gesture.o.id); o.w = Math.max(30, Math.round(gesture.o.w + p.x - gesture.p.x)); o.h = o.type === 'text' ? undefined : Math.max(30, Math.round(gesture.o.h + p.y - gesture.p.y)); if (o.type === 'stamp') { o.h = o.w; } renderObj(o); return; }
  if (gesture.type === 'line') { const o = gesture.o; let x = p.x, y = p.y; if (e.shiftKey) { const dx = x - o.pts[0], dy = y - o.pts[1]; const a = Math.round(Math.atan2(dy, dx) / (Math.PI / 4)) * (Math.PI / 4); const len = Math.hypot(dx, dy); x = o.pts[0] + Math.cos(a) * len; y = o.pts[1] + Math.sin(a) * len; } o.pts = [o.pts[0], o.pts[1], Math.round(x), Math.round(y)]; renderObj(o); return; }
  if (gesture.type === 'draw') { const o = gesture.o; o.pts.push(Math.round(p.x), Math.round(p.y)); renderObj(o); return; }
  if (gesture.type === 'connect') { const c = gesture.c; c.tx = p.x; c.ty = p.y; renderObj(c); return; }
  if (gesture.type === 'erase') { const t = document.elementFromPoint(e.clientX, e.clientY); const id = t?.closest?.('.obj,path.hit')?.dataset?.id; if (id && objs.get(id)) del(id); return; }
  if (gesture.type === 'marquee') { const a = gesture.p; const x = Math.min(a.x, p.x), y = Math.min(a.y, p.y), w = Math.abs(p.x - a.x), hh = Math.abs(p.y - a.y); Object.assign(gesture.el.style, { left: view.x + x * view.z + 'px', top: view.y + y * view.z + 'px', width: w * view.z + 'px', height: hh * view.z + 'px' }); gesture.box = { x, y, w, h: hh }; }
});
const throttledSend = (() => { let t = 0; return (list) => { if (Date.now() - t > 50) { t = Date.now(); net.send({ t: 'ops', ops: list.map((o) => ({ ...o, v: Date.now() * 1000 })) }); } }; })();
board.addEventListener('pointerup', (e) => {
  pointers.delete(e.pointerId); if (pointers.size) return;
  const g = gesture; gesture = null; board.style.cursor = tool === 'hand' ? 'grab' : '';
  if (!g) return;
  if (g.type === 'move' && g.moved) for (const o of g.orig) put({ ...objs.get(o.id) });
  if (g.type === 'move' && !g.moved && e.detail >= 2) { const o = objs.get(g.orig[0]?.id); if (o) editText(o); }
  if (g.type === 'resize') put({ ...objs.get(g.o.id) });
  if (g.type === 'line') { if (Math.hypot(g.o.pts[2] - g.o.pts[0], g.o.pts[3] - g.o.pts[1]) < 4) g.o.pts[2] += 120; put({ ...g.o }); setTool('select'); setSel([g.o.id]); }
  if (g.type === 'move' && !g.moved && g.orig.length === 1 && g.orig[0].type === 'comment' && e.detail < 2) openThread(objs.get(g.orig[0].id));
  if (g.type === 'draw') { if (g.o.pts.length < 4) g.o.pts.push(g.o.pts[0] + 0.5, g.o.pts[1]); put({ ...g.o }); }
  if (g.type === 'connect') { const t = document.elementFromPoint(e.clientX, e.clientY)?.closest?.('.obj'); if (t && t.dataset.id !== g.c.from) put({ ...g.c, to: t.dataset.id }); else { objs.delete(g.c.id); nodes.get(g.c.id)?.remove(); nodes.delete(g.c.id); net.send({ t: 'ops', ops: [{ ...g.c, deleted: true, v: Date.now() * 1000 }] }); } }
  if (g.type === 'marquee') { g.el.remove(); if (g.box) setSel([...objs.values()].filter((o) => !o.deleted && o.type !== 'connector').filter((o) => { const b = objBox(o); return b.x >= g.box.x && b.y >= g.box.y && b.x + b.w <= g.box.x + g.box.w && b.y + b.h <= g.box.y + g.box.h; }).map((o) => o.id)); }
  placeCtx();
});
board.addEventListener('dblclick', (e) => { const n = e.target.closest('.obj'); if (n) { editText(objs.get(n.dataset.id)); return; } if (tool === 'select' && !e.target.closest('.dock,.zoom-ctl,.ctx-bar')) { const p = toWorld(e.clientX, e.clientY); const o = { id: uid(), type: 'sticky', x: Math.round(p.x - 90), y: Math.round(p.y - 90), w: 180, h: 180, color: lastSticky, text: '' }; put(o); setSel([o.id]); editText(o); } });
board.addEventListener('wheel', (e) => { e.preventDefault(); if (e.ctrlKey || e.metaKey) zoomAt(Math.exp(-e.deltaY * 0.0025), e.clientX, e.clientY); else { view.x -= e.deltaX; view.y -= e.deltaY; applyView(); } }, { passive: false });
board.addEventListener('contextmenu', (e) => { e.preventDefault(); const n = e.target.closest('.obj,path.hit'); if (n && !sel.has(n.dataset.id)) setSel([n.dataset.id]); menu({ x: e.clientX, y: e.clientY }, sel.size ? [{ label: 'Дублировать', icon: 'copy', hint: 'Ctrl+D', action: duplicate }, { label: 'На передний план', icon: 'bring-front', action: () => { for (const id of sel) { const o = objs.get(id); objs.delete(id); put({ ...o }); } renderAll(); } }, { label: 'Удалить', icon: 'trash', danger: true, action: () => { [...sel].forEach(del); setSel([]); } }] : [{ label: 'Стикер', icon: 'sticky', action: () => setTool('sticky') }, { label: 'Показать всё', icon: 'fit', action: fitAll }]); });
document.addEventListener('keydown', (e) => {
  if (editingId || /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) || document.querySelector('.overlay')) return;
  const mod = e.ctrlKey || e.metaKey; const k = e.key.toLowerCase();
  if (e.key === ' ') { spaceDown = true; board.style.cursor = 'grab'; e.preventDefault(); return; }
  if (mod && k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
  if (mod && k === 'y') { e.preventDefault(); redo(); return; }
  if (mod && k === 'd') { e.preventDefault(); duplicate(); return; }
  if (mod && k === 'a') { e.preventDefault(); setSel([...objs.values()].filter((o) => !o.deleted && o.type !== 'connector').map((o) => o.id)); return; }
  if (mod && (k === '=' || k === '+')) { e.preventDefault(); zoomAt(1.2); return; } if (mod && k === '-') { e.preventDefault(); zoomAt(0.83); return; } if (mod && k === '0') { e.preventDefault(); fitAll(); return; }
  if (mod) return;
  if (e.key === 'Delete' || e.key === 'Backspace') { [...sel].forEach(del); setSel([]); return; }
  if (e.key === 'Escape') { setSel([]); setTool('select'); return; }
  if (e.key === 'Enter' && sel.size === 1) { e.preventDefault(); editText(objs.get([...sel][0])); return; }
  if (k === 'r' || k === 'к') { lastShape = 'rect'; setTool('shape'); return; }
  if (k === 'o' || k === 'щ') { lastShape = 'ellipse'; setTool('shape'); return; }
  const map = { v: 'select', h: 'hand', s: 'sticky', t: 'text', c: 'connector', l: 'line', p: 'pen', m: 'marker', e: 'eraser', f: 'frame', k: 'comment', м: 'select', р: 'hand', ы: 'sticky', е: 'text', с: 'connector', д: 'line', з: 'pen', ь: 'marker', у: 'eraser', а: 'frame', л: 'comment' };
  if (map[k]) setTool(map[k]);
});
document.addEventListener('keyup', (e) => { if (e.key === ' ') { spaceDown = false; board.style.cursor = ''; } });
document.addEventListener('paste', async (e) => { if (editingId || /INPUT|TEXTAREA/.test(document.activeElement.tagName)) return; const f = [...e.clipboardData.files].find((x) => x.type.startsWith('image/')); if (f) { const r = new FileReader(); r.onload = () => addImage(r.result); r.readAsDataURL(f); return; } const t = e.clipboardData.getData('text/plain'); if (t) { const c = center(); const o = { id: uid(), type: 'sticky', x: c.x - 90, y: c.y - 90, w: 180, h: 180, color: lastSticky, text: t.slice(0, 500) }; put(o); } });
async function insertImage() { const src = await pickImage(); if (src) addImage(src); }
async function addImage(src) { const img = new Image(); img.src = src; await img.decode().catch(() => {}); let w = img.naturalWidth || 400, hh = img.naturalHeight || 300; const k = Math.min(1, 480 / w, 480 / hh); const c = center(); const o = { id: uid(), type: 'image', src: await downscale(src, img), x: Math.round(c.x - (w * k) / 2), y: Math.round(c.y - (hh * k) / 2), w: Math.round(w * k), h: Math.round(hh * k) }; put(o); }
async function downscale(src, img) { if (src.length < 600000) return src; const c = document.createElement('canvas'); const k = Math.min(1, 1600 / img.naturalWidth); c.width = img.naturalWidth * k; c.height = img.naturalHeight * k; c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); return c.toDataURL('image/jpeg', 0.85); }

/* =================== collaboration =================== */
const peers = new Map(); const cursorEls = new Map();
const peersEl = h('div', { class: 'peers no-drag' }); const connBadge = h('span', { class: 'conn-badge hidden' });
const net = {
  ws: null, mode: null, info: null,
  send(m) { if (this.ws?.readyState === 1) this.ws.send(JSON.stringify(m)); },
  connect(url, { mode, pin, label, seed = false }) {
    this.close(); this.mode = mode; this.info = { url, label };
    let ws; try { ws = new WebSocket(url); } catch (e) { toast('Неверный адрес: ' + e.message, 'error'); return; }
    this.ws = ws;
    connBadge.classList.remove('hidden'); connBadge.innerHTML = '<span class="spinner" style="width:10px;height:10px"></span> Подключение…';
    ws.onopen = () => ws.send(JSON.stringify({ t: 'hello', name: me.name, color: me.color, pin, client: myId }));
    ws.onmessage = (ev) => {
      let m; try { m = JSON.parse(ev.data); } catch { return; }
      if (m.t === 'welcome') {
        connBadge.innerHTML = `<i></i>${esc(label)}`;
        const remote = m.objects || [];
        if (seed && !remote.length) { this.send({ t: 'ops', ops: [...objs.values()] }); }
        else { if (!seed) objs = new Map(); for (const o of remote) { const cur = objs.get(o.id); if (!cur || cur.v <= o.v) objs.set(o.id, o); } renderAll(); if (seed) this.send({ t: 'ops', ops: [...objs.values()].filter((o) => !remote.find((r) => r.id === o.id)) }); }
        if (m.meta?.title && mode === 'server') shell.setTitle(m.meta.title);
        if (m.meta) applyMeta(m.meta);
        peers.clear(); for (const p of m.peers || []) peers.set(p.id, p); renderPeers();
        toast('Подключено: ' + label, 'ok'); fitAll();
      }
      if (m.t === 'ops') for (const o of m.ops) put(o, { local: false });
      if (m.t === 'peer-join') { peers.set(m.peer.id, m.peer); renderPeers(); toast(`${m.peer.name} присоединился`); }
      if (m.t === 'peer-leave') { peers.delete(m.id); cursorEls.get(m.id)?.remove(); cursorEls.delete(m.id); renderPeers(); }
      if (m.t === 'cursor') moveCursor(m);
      if (m.t === 'meta' && m.meta) applyMeta(m.meta);
      if (m.t === 'error') { toast(m.message, 'error', 5000); ws.fatal = true; } // wrong PIN, host stopped, board deleted: don't retry
    };
    ws.onclose = (ev) => {
      if (this.ws !== ws) return;
      connBadge.innerHTML = '<i style="background:#f87171"></i>Нет соединения'; peers.clear(); renderPeers(); for (const c of cursorEls.values()) c.remove(); cursorEls.clear();
      if (!this.closing && !ws.fatal && ev.code !== 4003) setTimeout(() => { if (this.ws === ws) this.connect(url, { mode, pin, label, seed: true }); }, 3000);
    };
  },
  close() { this.closing = true; this.ws?.close(); this.ws = null; this.closing = false; connBadge.classList.add('hidden'); peers.clear(); renderPeers(); },
};
let lastCursor = 0;
function sendCursor(p) { if (!net.ws || Date.now() - lastCursor < 40) return; lastCursor = Date.now(); net.send({ t: 'cursor', x: Math.round(p.x), y: Math.round(p.y) }); }
function moveCursor(m) {
  const p = peers.get(m.id); if (!p) return;
  let c = cursorEls.get(m.id);
  if (!c) { c = h('div', { class: 'cursor-peer', html: `<svg width="18" height="20" viewBox="0 0 18 20"><path d="M1 1l6 17 2.5-7L17 8z" fill="${p.color}" stroke="#fff" stroke-width="1.5"/></svg><span style="background:${p.color}">${esc(p.name)}</span>` }); cursorsLayer.append(c); cursorEls.set(m.id, c); }
  c.style.transform = `translate(${view.x + m.x * view.z}px, ${view.y + m.y * view.z}px)`; c.dataset.x = m.x; c.dataset.y = m.y;
}
function renderPeers() { peersEl.innerHTML = ''; for (const p of [...peers.values()].slice(0, 6)) { const a = h('span', { class: 'avatar', style: { background: p.color }, 'data-tip': p.name + ' — показать', onclick: () => { const c = cursorEls.get(p.id); if (c) { view.x = board.clientWidth / 2 - c.dataset.x * view.z; view.y = board.clientHeight / 2 - c.dataset.y * view.z; applyView(); } } }, (p.name || '?')[0].toUpperCase()); peersEl.append(a); } }

async function collabDialog(tab = 'direct') {
  const port = settings.get('deskPort') || 47800;
  const body = h('div', { style: { width: 'min(560px,100%)' } });
  const tabs = h('div', { class: 'tabs-inline' }); const pane = h('div');
  const T = [['direct', 'По IP-адресу'], ['lan', 'Локальная сеть'], ['server', 'Через сервер']];
  let api;
  const show = async (id) => {
    [...tabs.children].forEach((b) => b.classList.toggle('on', b.dataset.id === id)); pane.innerHTML = '';
    if (id === 'direct' || id === 'lan') {
      if (native?.desk) {
        const st = await native.desk.status();
        if (st.hosting) {
          pane.append(h('p', {}, 'Вы проводите сессию. Соавторы подключаются по адресу:'), ...st.addresses.map((a) => h('div', { class: 'addr', style: { marginBottom: '6px' } }, `${a}:${st.port}`)), h('p', { class: 'small muted' }, `PIN-код: ${st.pin || 'не задан'}. ${st.lan ? 'Сессия видна в локальной сети.' : ''}`),
            h('button', { class: 'btn danger', onclick: async () => { await native.desk.stop(); net.close(); show(id); } }, 'Остановить сессию'));
        } else {
          const pin = h('input', { class: 'input', placeholder: 'Необязательно', maxlength: 8, style: { width: '140px' } });
          pane.append(h('h3', { style: { margin: '4px 0 8px' } }, 'Провести сессию'), h('p', { class: 'small muted' }, id === 'lan' ? 'Доска станет видна всем компьютерам Rabbit Office в вашей сети.' : 'Компьютер станет хостом: другие подключатся по вашему IP.'),
            h('div', { class: 'row' }, h('label', { class: 'small' }, 'PIN-код:'), pin, h('button', { class: 'btn primary', onclick: async () => { const r = await native.desk.host({ port, pin: pin.value.trim(), lan: id === 'lan', title: shell.title, name: me.name }); net.connect(`ws://127.0.0.1:${r.port}/desk`, { mode: id, pin: pin.value.trim(), label: id === 'lan' ? 'Локальная сеть' : 'Хост ' + (r.addresses[0] || ''), seed: true }); show(id); } }, 'Начать')));
        }
        pane.append(h('hr', { style: { border: 0, borderTop: '1px solid var(--border)', margin: '16px 0' } }));
      } else pane.append(h('p', { class: 'small muted' }, 'Проводить сессию можно в настольной версии (Windows, macOS, Linux). Здесь вы можете присоединиться к чужой сессии.'));
      if (id === 'lan') {
        const list = h('div', { class: 'found' }, h('div', { class: 'small muted' }, 'Поиск сессий в сети…'));
        pane.append(h('h3', { style: { margin: '4px 0' } }, 'Сессии в вашей сети'), list);
        const found = new Map();
        const draw = () => { list.innerHTML = ''; if (!found.size) list.append(h('div', { class: 'small muted' }, 'Пока ничего не найдено. Убедитесь, что хост запустил сессию в режиме «Локальная сеть».')); for (const s of found.values()) list.append(h('button', { class: 'list-item', onclick: async () => { const pin = s.needsPin ? await promptDlg('PIN-код', 'Введите PIN сессии', '') : ''; if (pin === null) return; net.connect(`ws://${s.ip}:${s.port}/desk`, { mode: 'lan', pin, label: s.title || s.ip }); api.close(); }, html: `${icon('wifi')}<div class="grow"><div class="name"><b>${esc(s.title || 'Доска')}</b></div><div class="meta">${esc(s.host || '')} · ${esc(s.ip)}:${s.port}${s.needsPin ? ' · PIN' : ''}</div></div>` })); };
        // one entry per host session even if beacons arrive over several network adapters
        if (native?.desk) { native.desk.discover((s) => { found.set(s.sid || s.ip + ':' + s.port, { ...s, ip: s.best || s.ip }); draw(); }); }
        else { // HTTP probe of common subnets
          const prefix = h('input', { class: 'input', value: localStorage.getItem('rabbit.lanPrefix') || '192.168.1', style: { width: '160px' } });
          pane.insertBefore(h('div', { class: 'row', style: { margin: '6px 0' } }, h('span', { class: 'small' }, 'Подсеть:'), prefix, h('button', { class: 'btn sm', onclick: () => { localStorage.setItem('rabbit.lanPrefix', prefix.value); scan(prefix.value); } }, 'Сканировать')), list);
          const scan = async (pre) => { found.clear(); draw(); const jobs = []; for (let i = 1; i < 255; i++) { const ip = `${pre}.${i}`; const ctl = new AbortController(); setTimeout(() => ctl.abort(), 1500); jobs.push(fetch(`http://${ip}:${port}/rabbit/info`, { signal: ctl.signal }).then((r) => r.json()).then((j) => { found.set(ip, { ...j, ip, port }); draw(); }).catch(() => {})); } await Promise.all(jobs); draw(); };
          scan(prefix.value);
        }
        draw();
      } else {
        const addr = h('input', { class: 'input', placeholder: '192.168.1.15 или 192.168.1.15:47800', value: localStorage.getItem('rabbit.lastHost') || '' });
        const pin = h('input', { class: 'input', placeholder: 'PIN (если задан)', style: { width: '160px' } });
        pane.append(h('h3', { style: { margin: '4px 0 8px' } }, 'Присоединиться по IP'), h('div', { class: 'row' }, addr, pin), h('div', { style: { marginTop: '10px' } }, h('button', { class: 'btn primary', onclick: () => { let a = addr.value.trim(); if (!a) return; localStorage.setItem('rabbit.lastHost', a); if (!/:\d+$/.test(a)) a += ':' + port; net.connect(`ws://${a}/desk`, { mode: 'direct', pin: pin.value.trim(), label: a }); api.close(); } }, 'Подключиться')));
      }
    } else {
      const list = accounts.list().filter((a) => a.mode !== 'storage');
      if (!list.length) { pane.append(h('p', {}, 'Подключите свой сервер Rabbit Office — доски будут храниться на нём и доступны всем пользователям сервера.'), h('button', { class: 'btn primary', onclick: async () => { if (await loginDialog()) show('server'); } }, 'Войти на сервер')); return; }
      const acc = accounts.active() && list.includes(accounts.active()) ? accounts.active() : list[0];
      pane.append(h('div', { class: 'row', style: { marginBottom: '8px' } }, h('b', {}, acc.serverName), h('span', { class: 'muted small' }, acc.displayName), h('span', { class: 'grow' }), h('button', { class: 'btn sm', html: icon('plus') + 'Новая доска', onclick: async () => { const t = await promptDlg('Новая доска', 'Название', shell.title); if (!t) return; try { const b = await cloudBoards.create(acc, t); openServerBoard(acc, b.id, b.title, true); api.close(); } catch (e) { toast(e.message, 'error'); } } })));
      const l = h('div', { class: 'found' }, h('div', { class: 'spinner' })); pane.append(l);
      try { const bl = await cloudBoards.list(acc); l.innerHTML = ''; if (!bl.length) l.append(h('div', { class: 'small muted' }, 'Досок пока нет — создайте первую (текущее содержимое будет загружено на сервер).')); for (const b of bl) l.append(h('button', { class: 'list-item', onclick: () => { openServerBoard(acc, b.id, b.title, false); api.close(); }, html: `${icon('grid')}<div class="grow"><div class="name"><b>${esc(b.title)}</b></div><div class="meta">${b.online ? b.online + ' онлайн · ' : ''}${new Date(b.updated).toLocaleString('ru-RU')}</div></div>` })); } catch (e) { l.innerHTML = ''; l.append(h('div', { class: 'small', style: { color: 'var(--danger)' } }, e.message)); }
    }
  };
  for (const [id, n] of T) tabs.append(h('button', { 'data-id': id, onclick: () => show(id) }, n));
  body.append(tabs, pane);
  if (net.ws) body.prepend(h('div', { class: 'banner', style: { marginTop: 0, marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '10px' } }, h('span', { class: 'grow' }, 'Подключено: ' + net.info.label), h('button', { class: 'btn sm', onclick: () => { net.close(); native?.desk?.stop(); api.close(); } }, 'Отключиться')));
  dialog({ title: 'Совместная работа', body, buttons: [], onOpen: (a) => { api = a; show(tab); } });
}
function openServerBoard(acc, id, title, seed) { shell.setTitle(title); shell.loc = null; net.connect(cloudBoards.wsUrl(acc, id), { mode: 'server', label: acc.serverName, seed }); }

/* =================== templates / files =================== */
function tpl(id) {
  objs = new Map(); const add = (o) => objs.set(o.id = uid(), { ...o, v: 1 });
  if (id === 'kanban') { ['Сделать', 'В работе', 'Готово'].forEach((t, i) => { add({ type: 'frame', x: i * 420, y: 0, w: 380, h: 640, text: t }); [0, 1].forEach((j) => add({ type: 'sticky', x: i * 420 + 100, y: 60 + j * 220, w: 180, h: 180, color: [STICKY[0], STICKY[3], STICKY[4]][i], text: ['Задача', 'Ещё задача'][j] })); }); }
  if (id === 'retro') { [['Что было хорошо', STICKY[3]], ['Что мешало', STICKY[2]], ['Что улучшим', STICKY[4]]].forEach(([t, c], i) => { add({ type: 'frame', x: i * 420, y: 0, w: 380, h: 560, text: t }); add({ type: 'sticky', x: i * 420 + 100, y: 70, w: 180, h: 180, color: c, text: 'Мысль…' }); }); }
  if (id === 'brainstorm') { const c = { id: 'center', type: 'shape', shape: 'ellipse', x: 0, y: 0, w: 220, h: 220, color: '#e0c3fc', text: 'Главная тема' }; objs.set('center', { ...c, v: 1 }); [[-420, -300], [420, -300], [-420, 260], [420, 260]].forEach(([x, y], i) => { const s = uid(); objs.set(s, { id: s, type: 'sticky', x: x + 20, y: y + 20, w: 180, h: 180, color: STICKY[i], text: 'Идея ' + (i + 1), v: 1 }); const cid = uid(); objs.set(cid, { id: cid, type: 'connector', from: 'center', to: s, color: '#888', v: 1 }); }); }
}
async function newDoc(t) { objs = new Map(); if (t) tpl(t); shell.setTitle(t ? TEMPLATES.desk.find((x) => x.id === t)?.name || 'Доска' : 'Доска'); renderAll(); requestAnimationFrame(fitAll); const q = new URLSearchParams(location.search); if (q.get('board')) { const acc = accounts.get(q.get('account')) || accounts.active(); if (acc) openServerBoard(acc, q.get('board'), 'Доска', false); } }
function serialize() { return JSON.stringify({ format: 'rabbit-board', version: 1, objects: [...objs.values()].filter((o) => !o.deleted), view }); }
async function load(file) { const d = JSON.parse(new TextDecoder().decode(file.bytes)); objs = new Map((d.objects || []).map((o) => [o.id, o])); renderAll(); if (d.view) { view = d.view; applyView(); } else requestAnimationFrame(fitAll); }
function exportSvg() {
  const list = [...objs.values()].filter((o) => !o.deleted); if (!list.length) return toast('Доска пуста');
  const b = bbox(list.filter((o) => o.type !== 'connector')); const pad = 40;
  let s = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${b.x - pad} ${b.y - pad} ${b.w + pad * 2} ${b.h + pad * 2}" font-family="Segoe UI, sans-serif"><rect x="${b.x - pad}" y="${b.y - pad}" width="${b.w + pad * 2}" height="${b.h + pad * 2}" fill="#fafafa"/><defs><marker id="a" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path d="M0 0L10 5L0 10z" fill="#555"/></marker></defs>`;
  const wrapT = (t, x, y, w, fs) => { const lines = []; for (const para of String(t || '').split('\n')) { let line = ''; for (const word of para.split(' ')) { if ((line + ' ' + word).length * fs * 0.55 > w && line) { lines.push(line); line = word; } else line = line ? line + ' ' + word : word; } lines.push(line); } return lines.map((l, i) => `<text x="${x}" y="${y + (i - (lines.length - 1) / 2) * fs * 1.25}" font-size="${fs}" text-anchor="middle" dominant-baseline="middle">${esc(l)}</text>`).join(''); };
  for (const o of list.filter((x) => x.type === 'frame')) s += `<rect x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" rx="10" fill="#fff" stroke="#bbb" stroke-width="2"/><text x="${o.x + 4}" y="${o.y - 10}" font-size="14" font-weight="600" fill="#555">${esc(o.text)}</text>`;
  for (const o of list) {
    if (o.type === 'connector') { const d = connD(o); if (d) s += `<path d="${d}" fill="none" stroke="#555" stroke-width="2.5" marker-end="url(#a)"/>`; }
    if (o.type === 'path') s += `<path d="${pathD(o.pts)}" fill="none" stroke="${o.color}" stroke-width="${o.width}" stroke-opacity="${o.marker ? 0.4 : 1}" stroke-linecap="round"${o.arrow ? ' marker-end="url(#a)"' : ''}${o.dash ? ' stroke-dasharray="10 8"' : ''}/>`;
    if (o.type === 'sticky') s += `<rect x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" fill="${o.color}"/>` + wrapT(o.text, o.x + o.w / 2, o.y + o.h / 2, o.w - 28, o.fs || 18);
    if (o.type === 'shape') s += `<svg x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" viewBox="0 0 100 100" preserveAspectRatio="none" overflow="visible"><g fill="${o.color}" stroke="#333" stroke-width="2" vector-effect="non-scaling-stroke">${(SHAPE_SVG[o.shape] || SHAPE_SVG.rect).replace(/\/>/g, ' vector-effect="non-scaling-stroke"/>')}</g></svg>` + wrapT(o.text, o.x + o.w / 2, o.y + o.h / 2, o.w - 20, o.fs || 16);
    if (o.type === 'text') s += `<text x="${o.x + 4}" y="${o.y + (o.fs || 20)}" font-size="${o.fs || 20}">${esc(o.text)}</text>`;
    if (o.type === 'image') s += `<image href="${o.src}" x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}"/>`;
    if (o.type === 'stamp') s += `<text x="${o.x + o.w / 2}" y="${o.y + o.h / 2}" font-size="${o.w * 0.8}" text-anchor="middle" dominant-baseline="middle">${o.text}</text>`;
  }
  return s + '</svg>';
}
async function exportPng() { const svg = exportSvg(); if (!svg) return; const img = new Image(); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg))); await img.decode(); const vb = /viewBox="[^ ]+ [^ ]+ ([\d.]+) ([\d.]+)"/.exec(svg); const w = +vb[1], hh = +vb[2]; const k = Math.min(2, 8000 / Math.max(w, hh)); const c = document.createElement('canvas'); c.width = w * k; c.height = hh * k; c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); const b = await new Promise((r) => c.toBlob(r, 'image/png')); shell.export(shell.title + '.png', new Uint8Array(await b.arrayBuffer()), 'image/png'); }

/* =================== shell =================== */
const shell = createShell({
  app: 'desk', defaultTitle: 'Доска', serialize, load, newDoc, templates: TEMPLATES.desk, undo, redo,
  isEmpty: () => ![...objs.values()].some((o) => !o.deleted),
  ribbon: [
    { id: 'home', label: 'Главная', groups: [
      { label: 'Инструменты', items: TOOLS.slice(0, 8).map(([id, ic, tip]) => ({ type: 'big', id: 'tool-' + id, icon: ic, label: tip.replace(/\s*\(.+\)/, ''), action: () => setTool(id) })) },
      { label: 'Совместная работа', items: [{ type: 'big', id: 'collab', icon: 'collab', label: 'Пригласить', action: () => collabDialog('direct') }, { type: 'big', id: 'lan', icon: 'wifi', label: 'Локальная сеть', action: () => collabDialog('lan') }, { type: 'big', id: 'srv', icon: 'server', label: 'Сервер', action: () => collabDialog('server') }] },
      { label: 'Вид', items: [{ type: 'big', id: 'fit', icon: 'fit', label: 'Показать всё', action: fitAll }] },
    ] },
  ],
  exporters: [
    { label: 'Изображение PNG', desc: 'Вся доска', icon: 'image', run: exportPng },
    { label: 'Векторное изображение SVG', desc: 'Для печати и редактирования', icon: 'shapes', run: () => { const s = exportSvg(); if (s) shell.export(shell.title + '.svg', s, 'image/svg+xml'); } },
  ],
  stats: () => [['Объектов', [...objs.values()].filter((o) => !o.deleted).length], ['Соединение', net.info?.label || 'Нет']],
});
shell.workspace.append(board);
const shareBtn = h('button', { class: 'tb-btn', html: icon('collab') + '<span class="hide-m">Поделиться</span>', onclick: () => collabDialog(native ? 'direct' : accounts.list().length ? 'server' : 'direct') });
shell.titlebarExtra.append(connBadge, peersEl, h('button', { class: 'tb-btn', 'data-tip': 'Таймер для всех', html: icon('timer'), onclick: (e) => timerMenu(e.currentTarget) }), shareBtn);
board.append(timerPill, mini);
zoomCtl.append(h('button', { html: icon('map'), 'data-tip': 'Мини-карта', onclick: toggleMini }));
shell.statusbar.append(h('span', {}, 'Пробел+перетаскивание — перемещение · Ctrl+колесо — масштаб · двойной клик — стикер'));
setTool('select');
view = { x: 0, y: 0, z: 1 };
new ResizeObserver(() => placeCtx()).observe(board);
shell.boot().then(() => { if (!objs.size) { view = { x: board.clientWidth / 2 - 300, y: board.clientHeight / 2 - 200, z: 1 }; applyView(); } });
