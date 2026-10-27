// Rabbit Point — presentations
import { createShell, applyTheme } from '../core/shell.js';
import { icon } from '../core/icons.js';
import { h, $, $$, esc, toast, dialog, menu, popover, closeMenus, colorPicker, promptDlg, confirmDlg, formDlg, pickImage, pickFile, readFileAs, debounce, clamp } from '../core/ui.js';
import { zipCreate } from '../core/zip.js';
import { TEMPLATES } from '../core/templates.js';
import { extOf, printToPdf, native } from '../core/platform.js';
import { CHART_TYPES } from '../core/charts.js';
import { SHAPES, THEMES, LAYOUTS, makeSlide, renderSlide, renderEl, slideBg, uidp, parseGrid, IMG_FILTERS } from './render.js';
import { exportPptx, importPptx } from './pptx.js';

const params = new URLSearchParams(location.search);
const channel = 'BroadcastChannel' in window ? new BroadcastChannel('rabbit-point-show') : null;
if (params.get('presenter')) presenterMode(); else editorMode();

/* =====================================================================
   EDITOR
   ===================================================================== */
function editorMode() {
  let pres = newPres();
  let cur = 0; let selIds = new Set(); let editingId = null; let scale = 1; let viewMode = 'normal'; let showGuides = true;
  const undoSt = []; const redoSt = []; let lastState = null;

  function newPres(themeId = 'office') { const p = { w: 1280, h: 720, theme: { ...THEMES.find((t) => t.id === themeId) }, slides: [] }; p.slides.push(makeSlide(p, 'title')); return p; }
  const slide = () => pres.slides[cur];
  const elById = (id) => slide().els.find((e) => e.id === id);
  const selected = () => slide().els.filter((e) => selIds.has(e.id));

  /* ---------- DOM ---------- */
  const thumbs = h('div', { class: 'panel-body' });
  const slidesPanel = h('aside', { class: 'panel pt-slides' }, h('div', { class: 'panel-head' }, 'Слайды', h('span', { class: 'grow' }), h('button', { class: 'btn ghost icon sm', 'data-tip': 'Новый слайд|Ctrl+M', html: icon('plus'), onclick: () => addSlide() })), thumbs);
  const stageWrap = h('div', { class: 'stage-wrap' });
  const stage = h('div', { class: 'stage' });
  stageWrap.append(stage);
  const notes = h('textarea', { placeholder: 'Заметки к слайду', spellcheck: 'true' });
  const notesPane = h('div', { class: 'notes-pane' }, notes);
  const main = h('div', { class: 'pt-main' }, stageWrap, notesPane);

  /* ---------- history ---------- */
  const snapshot = () => JSON.stringify(pres);
  function commit() {
    const s = snapshot(); if (s === lastState) return;
    if (lastState) { undoSt.push({ s: lastState, cur }); if (undoSt.length > 150) undoSt.shift(); }
    redoSt.length = 0; lastState = s; shell.markDirty(); updUndo();
  }
  function restore(entry, pushTo) { pushTo.push({ s: snapshot(), cur }); pres = JSON.parse(entry.s); lastState = entry.s; cur = clamp(entry.cur, 0, pres.slides.length - 1); selIds.clear(); editingId = null; renderAll(); shell.markDirty(); updUndo(); }
  function undo() { if (editingId) { document.execCommand('undo'); return; } const e = undoSt.pop(); if (e) restore(e, redoSt); }
  function redo() { if (editingId) { document.execCommand('redo'); return; } const e = redoSt.pop(); if (e) restore(e, undoSt); }
  function updUndo() { if (!shell) return; shell.undoBtn.disabled = !undoSt.length && !editingId; shell.redoBtn.disabled = !redoSt.length && !editingId; }

  /* ---------- rendering ---------- */
  let slideEl = null;
  function renderStage() {
    stage.innerHTML = '';
    slideEl = renderSlide(pres, slide(), { edit: true, showAnim: shell?.ribbon.active === 'anim' });
    stage.append(slideEl);
    stage.style.width = pres.w + 'px'; stage.style.height = pres.h + 'px';
    if (editingId) { const n = slideEl.querySelector(`.el[data-id="${editingId}"] .tx`); if (n) n.contentEditable = 'true'; }
    renderSel(); fit(); notes.value = slide().notes || '';
  }
  function fit() {
    const W = stageWrap.clientWidth, H = stageWrap.clientHeight; if (!W || !H) return;
    const pad = innerWidth < 760 ? 12 : 40;
    scale = Math.min((W - pad * 2) / pres.w, (H - pad * 2) / pres.h) * zoomMul;
    stage.style.transform = `scale(${scale})`;
    stage.style.left = Math.max(pad, (W - pres.w * scale) / 2) + 'px'; stage.style.top = Math.max(pad, (H - pres.h * scale) / 2) + 'px';
    renderSel();
  }
  let zoomMul = 1;
  function renderThumbs() {
    thumbs.innerHTML = '';
    const tw = 168;
    pres.slides.forEach((s, i) => {
      const t = h('div', { class: `thumb ${i === cur ? 'on' : ''} ${s.hidden ? 'hidden-slide' : ''}`, draggable: 'true', 'data-i': i });
      const sl = renderSlide(pres, s); sl.style.transform = `scale(${tw / pres.w})`; t.append(sl);
      if (s.transition?.type && s.transition.type !== 'none') t.append(h('span', { class: 'ico' }, '★'));
      t.onclick = () => goto(i);
      t.oncontextmenu = (e) => { e.preventDefault(); goto(i); slideMenu({ x: e.clientX, y: e.clientY }); };
      t.ondragstart = (e) => { e.dataTransfer.setData('text/x-slide', String(i)); e.dataTransfer.effectAllowed = 'move'; };
      t.ondragover = (e) => { e.preventDefault(); $$('.thumb', thumbs).forEach((x) => x.classList.remove('drop-before')); t.classList.add('drop-before'); };
      t.ondragleave = () => t.classList.remove('drop-before');
      t.ondrop = (e) => { e.preventDefault(); t.classList.remove('drop-before'); const from = +e.dataTransfer.getData('text/x-slide'); moveSlide(from, i); };
      const row = h('div', { class: 'thumb-row' }, h('span', { class: 'num' }, String(i + 1)), t);
      thumbs.append(row);
      if (viewMode === 'sorter') {}
    });
    // fix thumb inner size after layout
    requestAnimationFrame(() => $$('.thumb', thumbs).forEach((t) => { const sl = t.firstChild; sl.style.transform = `scale(${t.clientWidth / pres.w})`; }));
    if (viewMode === 'sorter') renderSorter();
  }
  const refreshThumb = debounce(() => {
    const t = thumbs.querySelector(`.thumb[data-i="${cur}"]`); if (!t) return renderThumbs();
    const sl = renderSlide(pres, slide()); sl.style.transform = `scale(${t.clientWidth / pres.w})`; t.replaceChild(sl, t.firstChild);
  }, 250);
  function renderAll() { renderStage(); renderThumbs(); syncRibbon(); }
  function goto(i) { exitEdit(); cur = clamp(i, 0, pres.slides.length - 1); selIds.clear(); renderAll(); thumbs.querySelector(`.thumb[data-i="${cur}"]`)?.scrollIntoView({ block: 'nearest' }); }

  /* ---------- selection box ---------- */
  let selBox = null;
  function renderSel() {
    selBox?.remove(); selBox = null; $$('.guide', stage).forEach((g) => g.remove());
    const els = selected(); if (!els.length) { syncRibbon(); return; }
    const hs = 11 / scale;
    selBox = h('div', { class: 'sel-box' });
    let box;
    if (els.length === 1) { const e = els[0]; box = { x: e.x, y: e.y, w: e.w, h: e.h, rot: e.rot || 0 }; }
    else { const x1 = Math.min(...els.map((e) => e.x)), y1 = Math.min(...els.map((e) => e.y)), x2 = Math.max(...els.map((e) => e.x + e.w)), y2 = Math.max(...els.map((e) => e.y + e.h)); box = { x: x1, y: y1, w: x2 - x1, h: y2 - y1, rot: 0 }; }
    Object.assign(selBox.style, { left: box.x + 'px', top: box.y + 'px', width: box.w + 'px', height: box.h + 'px', transform: `rotate(${box.rot}deg)`, borderWidth: 1.5 / scale + 'px' });
    if (!editingId) {
      const dirs = [[-1, -1], [0, -1], [1, -1], [1, 0], [1, 1], [0, 1], [-1, 1], [-1, 0]];
      const isLine = els.length === 1 && SHAPES[els[0].shape]?.line;
      for (const [dx, dy] of dirs) {
        if (isLine && dy !== 0) continue;
        const hd = h('div', { class: 'hd', style: { left: (dx + 1) * 50 + '%', top: (dy + 1) * 50 + '%', width: hs + 'px', height: hs + 'px', borderWidth: 1.5 / scale + 'px', cursor: cursorFor(dx, dy, box.rot) } });
        hd.addEventListener('pointerdown', (e) => startResize(e, dx, dy, box)); selBox.append(hd);
      }
      if (els.length === 1) {
        selBox.append(h('div', { class: 'rot-line', style: { top: -22 / scale + 'px', height: 22 / scale + 'px', width: 1 / scale + 'px' } }));
        const rh = h('div', { class: 'hd rot', style: { left: '50%', top: -22 / scale + 'px', width: hs + 'px', height: hs + 'px', borderWidth: 1.5 / scale + 'px' } });
        rh.addEventListener('pointerdown', (e) => startRotate(e, els[0])); selBox.append(rh);
      }
    }
    stage.append(selBox);
    syncRibbon();
  }
  function cursorFor(dx, dy, rot) { const a = ((Math.atan2(dy, dx) * 180) / Math.PI + rot + 360) % 180; return a < 22.5 || a >= 157.5 ? 'ew-resize' : a < 67.5 ? 'nwse-resize' : a < 112.5 ? 'ns-resize' : 'nesw-resize'; }
  function toSlide(e) { const r = stage.getBoundingClientRect(); return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale }; }

  /* ---------- pointer interactions ---------- */
  stageWrap.addEventListener('pointerdown', (e) => {
    if (e.button !== 0 && e.pointerType === 'mouse') return;
    const elNode = e.target.closest('.el'); const id = elNode?.dataset.id;
    if (editingId && id === editingId) return; // caret placement inside text
    if (e.target.closest('.hd')) return;
    exitEdit();
    closeMenus();
    if (!id) { // marquee
      if (!e.shiftKey) selIds.clear(); renderSel();
      const p0 = toSlide(e); const mq = h('div', { class: 'marquee' }); let moved = false;
      stage.append(mq);
      const mv = (ev) => { moved = true; const p = toSlide(ev); const x = Math.min(p0.x, p.x), y = Math.min(p0.y, p.y), w = Math.abs(p.x - p0.x), hh = Math.abs(p.y - p0.y); Object.assign(mq.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: hh + 'px', borderWidth: 1 / scale + 'px' }); selIds = expandGroups([...(e.shiftKey ? selIds : []), ...slide().els.filter((el) => el.x < x + w && el.x + el.w > x && el.y < y + hh && el.y + el.h > y).map((el) => el.id)]); renderSel(); };
      const up = () => { mq.remove(); document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); };
      document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
      return;
    }
    e.preventDefault();
    if (e.shiftKey || e.ctrlKey) { const grp = expandGroups([id]); if (selIds.has(id)) grp.forEach((x) => selIds.delete(x)); else grp.forEach((x) => selIds.add(x)); renderSel(); return; }
    if (!selIds.has(id)) { selIds = expandGroups([id]); renderSel(); }
    startMove(e);
  });
  stageWrap.addEventListener('dblclick', (e) => {
    const n = e.target.closest('.el'); if (!n) return;
    const el = elById(n.dataset.id); if (!el) return;
    if (el.type === 'text' || (el.type === 'shape' && !SHAPES[el.shape]?.line)) enterEdit(el.id, e);
    else if (el.type === 'table') enterTableEdit(el, e);
    else if (el.type === 'chart') chartDialog(el);
    else if (el.type === 'image') replaceImage(el);
  });
  stageWrap.addEventListener('contextmenu', (e) => {
    e.preventDefault(); const n = e.target.closest('.el');
    if (n && !selIds.has(n.dataset.id)) { selIds = new Set([n.dataset.id]); renderSel(); }
    const has = selIds.size > 0;
    menu({ x: e.clientX, y: e.clientY }, has ? [
      { label: 'Вырезать', icon: 'cut', hint: 'Ctrl+X', action: () => { copySel(); deleteSel(); } }, { label: 'Копировать', icon: 'copy', hint: 'Ctrl+C', action: copySel }, { label: 'Вставить', icon: 'paste', hint: 'Ctrl+V', action: pasteEls }, { label: 'Дублировать', icon: 'copy', hint: 'Ctrl+D', action: duplicateSel },
      '-', { label: 'На передний план', icon: 'bring-front', action: () => zorder('front') }, { label: 'На задний план', icon: 'send-back', action: () => zorder('back') },
      '-', { label: 'Анимация…', icon: 'animation', sub: ANIMS.map(([id, n]) => ({ label: n, action: () => setAnim(id) })) }, { label: 'Удалить', icon: 'trash', danger: true, hint: 'Del', action: deleteSel },
    ] : [{ label: 'Вставить', icon: 'paste', action: pasteEls }, { label: 'Новый слайд', icon: 'slide-new', action: () => addSlide() }, { label: 'Формат фона…', icon: 'background', action: bgDialog }, { label: 'Макет', icon: 'layout', sub: Object.entries(LAYOUTS).map(([k, l]) => ({ label: l.n, action: () => applyLayout(k) })) }]);
  });

  function snapLines(exclude) {
    const xs = [0, pres.w / 2, pres.w], ys = [0, pres.h / 2, pres.h];
    for (const el of slide().els) { if (exclude.has(el.id)) continue; xs.push(el.x, el.x + el.w / 2, el.x + el.w); ys.push(el.y, el.y + el.h / 2, el.y + el.h); }
    return { xs, ys };
  }
  function showGuidesAt(gx, gy) {
    $$('.guide', stage).forEach((g) => g.remove()); if (!showGuides) return;
    for (const x of gx) stage.append(h('div', { class: 'guide', style: { left: x + 'px', top: 0, width: 1 / scale + 'px', height: pres.h + 'px' } }));
    for (const y of gy) stage.append(h('div', { class: 'guide', style: { top: y + 'px', left: 0, height: 1 / scale + 'px', width: pres.w + 'px' } }));
  }
  function startMove(e) {
    const p0 = toSlide(e); const els = selected(); const orig = els.map((el) => ({ el, x: el.x, y: el.y }));
    const bb = { x: Math.min(...els.map((el) => el.x)), y: Math.min(...els.map((el) => el.y)), r: Math.max(...els.map((el) => el.x + el.w)), b: Math.max(...els.map((el) => el.y + el.h)) };
    const lines = snapLines(selIds); let moved = false;
    const mv = (ev) => {
      const p = toSlide(ev); let dx = p.x - p0.x, dy = p.y - p0.y;
      if (!moved && Math.hypot(dx, dy) * scale < 3) return; moved = true;
      if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
      const th = 6 / scale; const gx = [], gy = [];
      if (!ev.altKey) {
        const cand = [bb.x + dx, (bb.x + bb.r) / 2 + dx, bb.r + dx]; let best = null;
        for (const c of cand) for (const l of lines.xs) { const d = l - c; if (Math.abs(d) < th && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, l }; }
        if (best) { dx += best.d; gx.push(best.l); }
        const candY = [bb.y + dy, (bb.y + bb.b) / 2 + dy, bb.b + dy]; best = null;
        for (const c of candY) for (const l of lines.ys) { const d = l - c; if (Math.abs(d) < th && (!best || Math.abs(d) < Math.abs(best.d))) best = { d, l }; }
        if (best) { dy += best.d; gy.push(best.l); }
      }
      for (const o of orig) { o.el.x = Math.round(o.x + dx); o.el.y = Math.round(o.y + dy); const n = slideEl.querySelector(`.el[data-id="${o.el.id}"]`); if (n) { n.style.left = o.el.x + 'px'; n.style.top = o.el.y + 'px'; } }
      renderSel(); showGuidesAt(gx, gy);
    };
    const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); $$('.guide', stage).forEach((g) => g.remove()); if (moved) { commit(); refreshThumb(); } };
    document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
  }
  function startResize(e, dx, dy, box) {
    e.preventDefault(); e.stopPropagation();
    const els = selected(); const p0 = toSlide(e);
    const orig = els.map((el) => ({ el, x: el.x, y: el.y, w: el.w, h: el.h, size: el.text?.size }));
    const single = els.length === 1; const rot = single ? ((els[0].rot || 0) * Math.PI) / 180 : 0;
    const keepAspect = single && els[0].type === 'image';
    const mv = (ev) => {
      const p = toSlide(ev); const wx = p.x - p0.x, wy = p.y - p0.y;
      const lx = wx * Math.cos(-rot) - wy * Math.sin(-rot), ly = wx * Math.sin(-rot) + wy * Math.cos(-rot);
      let nw = Math.max(10, box.w + dx * lx), nh = Math.max(dy === 0 && SHAPES[els[0]?.shape]?.line ? 1 : 10, box.h + dy * ly);
      if ((ev.shiftKey !== keepAspect) && dx && dy) { const r = box.w / box.h; if (nw / nh > r) nh = nw / r; else nw = nh * r; }
      if (single) {
        const o = orig[0]; const cx0 = o.x + o.w / 2, cy0 = o.y + o.h / 2;
        const lcx = (dx * (nw - o.w)) / 2, lcy = (dy * (nh - o.h)) / 2;
        const cx = cx0 + lcx * Math.cos(rot) - lcy * Math.sin(rot), cy = cy0 + lcx * Math.sin(rot) + lcy * Math.cos(rot);
        Object.assign(o.el, { w: Math.round(nw), h: Math.round(nh), x: Math.round(cx - nw / 2), y: Math.round(cy - nh / 2) });
      } else {
        const sx = nw / box.w, sy = nh / box.h; const ox = dx < 0 ? box.x + box.w : box.x, oy = dy < 0 ? box.y + box.h : box.y;
        for (const o of orig) { o.el.x = Math.round(ox + (o.x - ox) * (dx ? sx : 1)); o.el.y = Math.round(oy + (o.y - oy) * (dy ? sy : 1)); o.el.w = Math.round(o.w * (dx ? sx : 1)); o.el.h = Math.round(o.h * (dy ? sy : 1)); }
      }
      for (const o of orig) { const n = slideEl.querySelector(`.el[data-id="${o.el.id}"]`); if (n) { const nn = renderEl(pres, o.el, { edit: true }); n.replaceWith(nn); } }
      slideEl = stage.querySelector('.slide'); renderSel();
    };
    const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); commit(); renderStage(); refreshThumb(); };
    document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
  }
  function startRotate(e, el) {
    e.preventDefault(); e.stopPropagation();
    const cx = el.x + el.w / 2, cy = el.y + el.h / 2;
    const mv = (ev) => { const p = toSlide(ev); let a = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90; if (ev.shiftKey) a = Math.round(a / 15) * 15; else { const s = Math.round(a / 90) * 90; if (Math.abs(a - s) < 4) a = s; } el.rot = Math.round(((a % 360) + 360) % 360); const n = slideEl.querySelector(`.el[data-id="${el.id}"]`); n.style.transform = `rotate(${el.rot}deg)`; renderSel(); };
    const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); commit(); refreshThumb(); };
    document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up);
  }

  /* ---------- text editing ---------- */
  function enterEdit(id, ev) {
    exitEdit(); const el = elById(id); if (!el) return;
    if (!el.text) el.text = { html: '', size: 24, align: 'center' };
    editingId = id; selIds = new Set([id]);
    const node = slideEl.querySelector(`.el[data-id="${id}"]`); const tx = node.querySelector('.tx');
    node.classList.remove('placeholder-frame');
    if (!el.text.html && el.text.list) tx.innerHTML = '<ul><li><br></li></ul>';
    tx.contentEditable = 'true'; tx.focus();
    const r = document.createRange();
    if (ev && document.caretRangeFromPoint) { const cr = document.caretRangeFromPoint(ev.clientX, ev.clientY); if (cr && tx.contains(cr.startContainer)) { getSelection().removeAllRanges(); getSelection().addRange(cr); } else { r.selectNodeContents(tx); r.collapse(false); getSelection().removeAllRanges(); getSelection().addRange(r); } }
    else { r.selectNodeContents(tx); getSelection().removeAllRanges(); getSelection().addRange(r); }
    tx.addEventListener('input', onTextInput);
    renderSel(); updUndo();
  }
  function onTextInput() { const el = elById(editingId); if (!el) return; const tx = slideEl.querySelector(`.el[data-id="${editingId}"] .tx`); el.text.html = cleanHtml(tx.innerHTML); shell.markDirty(); refreshThumb(); }
  function cleanHtml(html) { const t = html.replace(/<br>$/, ''); return /^(<ul><li>(<br>)?<\/li><\/ul>|<br>|<div><br><\/div>)$/.test(t) ? '' : t; }
  function exitEdit() {
    if (!editingId) return;
    const el = elById(editingId); const tx = slideEl?.querySelector(`.el[data-id="${editingId}"] .tx`);
    if (el && tx) { el.text.html = cleanHtml(tx.innerHTML); tx.contentEditable = 'false'; tx.removeEventListener('input', onTextInput); }
    editingId = null; getSelection().removeAllRanges(); commit(); renderStage(); refreshThumb();
  }
  function enterTableEdit(el, ev) {
    const node = slideEl.querySelector(`.el[data-id="${el.id}"]`); const td = ev.target.closest('td'); if (!td) return;
    editingId = el.id;
    node.querySelectorAll('td').forEach((c) => (c.contentEditable = 'true'));
    td.focus();
    const save = () => { const rows = [...node.querySelectorAll('tr')].map((tr) => [...tr.cells].map((c) => c.innerHTML)); el.table.rows = rows; shell.markDirty(); refreshThumb(); };
    node.addEventListener('input', save);
    const tx = { contentEditable: '' };
    // when leaving, exitEdit renders stage from model
    const origExit = exitEdit;
    node.dataset.tableEdit = '1';
  }

  /* text formatting: inline when editing, element-level otherwise */
  function txt(fn, inline) {
    if (editingId && inline) { document.execCommand('styleWithCSS', false, true); inline(); onTextInput(); return; }
    const els = selected().filter((e) => e.text || e.type === 'table'); if (!els.length) return;
    for (const e of els) { if (!e.text && e.type !== 'table') continue; fn(e); }
    commit(); renderStage(); refreshThumb();
  }
  function setFontSize(px) {
    px = clamp(parseFloat(String(px).replace(',', '.')) || 24, 6, 400);
    if (editingId) {
      const s = getSelection(); if (s.rangeCount && !s.isCollapsed) { document.execCommand('styleWithCSS', false, true); document.execCommand('fontSize', false, '7'); slideEl.querySelectorAll('.tx span[style*="xxx-large"], .tx font[size="7"]').forEach((n) => { if (n.tagName === 'FONT') { const sp = document.createElement('span'); sp.style.fontSize = px + 'px'; sp.append(...n.childNodes); n.replaceWith(sp); } else n.style.fontSize = px + 'px'; }); onTextInput(); return; }
    }
    const was = editingId; if (was) exitEdit();
    for (const e of selected()) { if (e.text) e.text.size = px; if (e.type === 'table') e.size = px; }
    commit(); renderStage(); refreshThumb();
  }
  function bump(d) { const e = selected()[0]; if (!e) return; const s = e.text?.size || e.size || 24; const steps = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 88, 96, 120]; const i = steps.findIndex((x) => x >= s); setFontSize(d > 0 ? steps[Math.min(steps.length - 1, steps[i] === s ? i + 1 : i)] : steps[Math.max(0, (i < 0 ? steps.length : i) - 1)]); }

  /* ---------- element ops ---------- */
  function addEl(el) { exitEdit(); slide().els.push(el); selIds = new Set([el.id]); commit(); renderStage(); refreshThumb(); return el; }
  function addText() { const el = addEl({ id: uidp(), type: 'text', x: pres.w / 2 - 220, y: pres.h / 2 - 40, w: 440, h: 80, text: { html: '', size: 28, align: 'left' } }); enterEdit(el.id); }
  function addShape(kind) { const sh = SHAPES[kind]; const line = sh.line; addEl({ id: uidp(), type: 'shape', shape: kind, x: pres.w / 2 - (line ? 150 : 110), y: pres.h / 2 - (line ? 2 : 110), w: line ? 300 : 220, h: line ? 4 : 220, fill: line ? undefined : pres.theme.accent, stroke: line ? pres.theme.fg : undefined, strokeW: line ? 3 : 0, text: line ? undefined : { html: '', size: 22, align: 'center', valign: 'middle', color: '#ffffff' } }); }
  async function addImage(src) {
    src = src || (await pickImage()); if (!src) return;
    const img = new Image(); img.src = src; await img.decode().catch(() => {});
    let w = img.naturalWidth || 600, hh = img.naturalHeight || 400; const k = Math.min(1, (pres.w * 0.7) / w, (pres.h * 0.7) / hh); w *= k; hh *= k;
    addEl({ id: uidp(), type: 'image', src, x: Math.round((pres.w - w) / 2), y: Math.round((pres.h - hh) / 2), w: Math.round(w), h: Math.round(hh) });
  }
  async function replaceImage(el) { const src = await pickImage(); if (!src) return; el.src = src; commit(); renderStage(); refreshThumb(); }
  function addTable(rows = 3, cols = 4) { addEl({ id: uidp(), type: 'table', x: 120, y: 200, w: pres.w - 240, h: rows * 56, size: 20, table: { header: true, band: true, rows: Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? `Столбец ${c + 1}` : ''))) } }); }
  function addChart(type = 'column') { addEl({ id: uidp(), type: 'chart', x: 200, y: 170, w: 880, h: 480, chart: { type, title: 'Диаграмма', data: 'Категория;Ряд 1;Ряд 2\nA;4,3;2,4\nB;2,5;4,4\nC;3,5;1,8\nD;4,5;2,8' } }); }
  async function chartDialog(el) {
    const ta = h('textarea', { class: 'input mono', rows: 8 }); ta.value = el.chart.data;
    const r = await dialog({ title: 'Данные диаграммы', body: h('div', { style: { width: 'min(520px,100%)' } }, h('div', { class: 'row' }, h('div', { class: 'field grow' }, h('label', {}, 'Тип'), (() => { const s = h('select', { class: 'input', id: 'ct' }, ...CHART_TYPES.map((t) => h('option', { value: t.id, selected: t.id === el.chart.type }, t.name))); return s; })()), h('div', { class: 'field grow' }, h('label', {}, 'Заголовок'), h('input', { class: 'input', id: 'ctitle', value: el.chart.title || '' }))), h('div', { class: 'field' }, h('label', {}, 'Данные: первая строка — названия рядов, первый столбец — категории. Разделитель «;»'), ta)), buttons: [{ label: 'Отмена', value: null }, { label: 'Применить', primary: true, value: () => ({ type: $('#ct').value, title: $('#ctitle').value, data: ta.value }) }] });
    if (!r) return; Object.assign(el.chart, r); commit(); renderStage(); refreshThumb();
  }
  function deleteSel() { if (!selIds.size) return; slide().els = slide().els.filter((e) => !selIds.has(e.id)); selIds.clear(); commit(); renderStage(); refreshThumb(); }
  function copySel() { const els = selected(); if (!els.length) return; localStorage.setItem('rabbit.point.clip', JSON.stringify(els)); }
  function pasteEls() { let els; try { els = JSON.parse(localStorage.getItem('rabbit.point.clip') || '[]'); } catch { els = []; } if (!els.length) return; exitEdit(); const ids = []; for (const e of els) { const n = { ...structuredClone(e), id: uidp(), x: e.x + 24, y: e.y + 24 }; slide().els.push(n); ids.push(n.id); } localStorage.setItem('rabbit.point.clip', JSON.stringify(els.map((e) => ({ ...e, x: e.x + 24, y: e.y + 24 })))); selIds = new Set(ids); commit(); renderStage(); refreshThumb(); }
  function duplicateSel() { copySel(); pasteEls(); }
  function zorder(dir) { const els = slide().els; const sel = els.filter((e) => selIds.has(e.id)); const rest = els.filter((e) => !selIds.has(e.id)); if (dir === 'front') slide().els = [...rest, ...sel]; else if (dir === 'back') slide().els = [...sel, ...rest]; else { for (const e of sel) { const i = slide().els.indexOf(e); const j = clamp(i + (dir === 'up' ? 1 : -1), 0, els.length - 1); slide().els.splice(i, 1); slide().els.splice(j, 0, e); } } commit(); renderStage(); refreshThumb(); }
  function align(kind) {
    const els = selected(); if (!els.length) return;
    const multi = els.length > 1; const bb = multi ? { x: Math.min(...els.map((e) => e.x)), y: Math.min(...els.map((e) => e.y)), r: Math.max(...els.map((e) => e.x + e.w)), b: Math.max(...els.map((e) => e.y + e.h)) } : { x: 0, y: 0, r: pres.w, b: pres.h };
    for (const e of els) { if (kind === 'l') e.x = bb.x; if (kind === 'c') e.x = Math.round((bb.x + bb.r) / 2 - e.w / 2); if (kind === 'r') e.x = bb.r - e.w; if (kind === 't') e.y = bb.y; if (kind === 'm') e.y = Math.round((bb.y + bb.b) / 2 - e.h / 2); if (kind === 'b') e.y = bb.b - e.h; }
    if (kind === 'dh' && els.length > 2) { const s = [...els].sort((a, b) => a.x - b.x); const total = s.reduce((a, e) => a + e.w, 0); const gap = (bb.r - bb.x - total) / (s.length - 1); let x = bb.x; for (const e of s) { e.x = Math.round(x); x += e.w + gap; } }
    if (kind === 'dv' && els.length > 2) { const s = [...els].sort((a, b) => a.y - b.y); const total = s.reduce((a, e) => a + e.h, 0); const gap = (bb.b - bb.y - total) / (s.length - 1); let y = bb.y; for (const e of s) { e.y = Math.round(y); y += e.h + gap; } }
    commit(); renderStage(); refreshThumb();
  }
  function styleEls(fn) { exitEdit(); for (const e of selected()) fn(e); commit(); renderStage(); refreshThumb(); }
  const ANIMS = [['none', 'Нет'], ['fade', 'Появление'], ['fly', 'Вылет снизу'], ['flyl', 'Вылет слева'], ['zoom', 'Увеличение'], ['wipe', 'Стирание'], ['bounce', 'Выскакивание'], ['spin', 'Вращение']];
  function setAnim(type) { const els = selected(); if (!els.length) return toast('Выберите объект на слайде'); let max = Math.max(0, ...slide().els.map((e) => e.anim?.order || 0)); for (const e of els) { e.anim = type === 'none' ? undefined : { type, order: e.anim?.order || ++max, dur: e.anim?.dur || 0.6 }; } commit(); renderStage(); }

  /* ---------- slides ops ---------- */
  function addSlide(layout) { exitEdit(); const lay = layout || (slide()?.layout === 'title' ? 'content' : slide()?.layout || 'content'); pres.slides.splice(cur + 1, 0, makeSlide(pres, lay)); cur++; selIds.clear(); commit(); renderAll(); }
  function dupSlide() { const c = structuredClone(slide()); c.id = uidp(); c.els.forEach((e) => (e.id = uidp())); pres.slides.splice(cur + 1, 0, c); cur++; commit(); renderAll(); }
  async function delSlide() { if (pres.slides.length === 1) { pres.slides[0] = makeSlide(pres, 'blank'); } else pres.slides.splice(cur, 1); cur = clamp(cur, 0, pres.slides.length - 1); selIds.clear(); commit(); renderAll(); }
  function moveSlide(from, to) { if (from === to) return; const [s] = pres.slides.splice(from, 1); pres.slides.splice(to > from ? to - 1 : to, 0, s); cur = pres.slides.indexOf(s); commit(); renderAll(); }
  function applyLayout(k) {
    const L = LAYOUTS[k]; const old = slide(); const tmp = makeSlide(pres, k);
    const texts = old.els.filter((e) => e.ph); const others = old.els.filter((e) => !e.ph);
    tmp.els.forEach((e) => { const prev = texts.find((t) => t.ph === e.ph && !t.used); if (prev) { prev.used = true; e.text.html = prev.text.html; } });
    old.els = [...tmp.els, ...others]; old.layout = k; commit(); renderAll();
  }
  function slideMenu(at) {
    menu(at, [{ label: 'Создать слайд', icon: 'slide-new', action: () => addSlide() }, { label: 'Дублировать слайд', icon: 'copy', action: dupSlide }, { label: 'Удалить слайд', icon: 'trash', danger: true, action: delSlide }, '-',
      { label: slide().hidden ? 'Показывать слайд' : 'Скрыть слайд', icon: 'eye-off', action: () => { slide().hidden = !slide().hidden; commit(); renderThumbs(); } },
      { label: 'Макет', icon: 'layout', sub: Object.entries(LAYOUTS).map(([k, l]) => ({ label: l.n, checked: slide().layout === k, action: () => applyLayout(k) })) }, { label: 'Формат фона…', icon: 'background', action: bgDialog }]);
  }
  async function bgDialog() {
    const s = slide(); let mode = s.bg?.image ? 'image' : s.bg?.gradient ? 'gradient' : s.bg?.color ? 'color' : 'theme';
    const c1 = h('input', { type: 'color', value: (s.bg?.color && s.bg.color.startsWith('#')) ? s.bg.color : '#ffffff' });
    const g1 = h('input', { type: 'color', value: '#ff8a5b' }), g2 = h('input', { type: 'color', value: '#7b4fd6' });
    const sel = h('select', { class: 'input' }, ...[['theme', 'Как в теме'], ['color', 'Сплошная заливка'], ['gradient', 'Градиент'], ['image', 'Рисунок']].map(([v, l]) => h('option', { value: v, selected: v === mode }, l)));
    const all = h('input', { type: 'checkbox' });
    let img = s.bg?.image || null;
    const imgBtn = h('button', { class: 'btn sm', onclick: async () => { img = await pickImage(); if (img) sel.value = 'image'; } }, 'Выбрать рисунок…');
    const r = await dialog({ title: 'Формат фона', body: h('div', { style: { width: 'min(380px,100%)' } }, h('div', { class: 'field' }, h('label', {}, 'Заливка'), sel), h('div', { class: 'row', style: { marginBottom: '10px' } }, h('span', { class: 'small muted' }, 'Цвет:'), c1, h('span', { class: 'small muted', style: { marginLeft: '12px' } }, 'Градиент:'), g1, g2), imgBtn, h('br'), h('br'), h('label', { class: 'check' }, all, 'Применить ко всем слайдам')), buttons: [{ label: 'Отмена', value: false }, { label: 'Применить', value: true, primary: true }] });
    if (!r) return;
    const bg = sel.value === 'theme' ? null : sel.value === 'color' ? { color: c1.value } : sel.value === 'gradient' ? { gradient: `linear-gradient(135deg, ${g1.value}, ${g2.value})` } : img ? { image: img } : null;
    for (const sl of all.checked ? pres.slides : [s]) sl.bg = bg;
    commit(); renderAll();
  }
  function setTheme(id) { const t = THEMES.find((x) => x.id === id); pres.theme = { ...t }; for (const s of pres.slides) for (const e of s.els) if (e.type === 'shape' && !SHAPES[e.shape]?.line && e.fillFromTheme !== false) e.fill = t.accent; commit(); renderAll(); }
  function setSize(ratio) { const nh = ratio === '4:3' ? 960 : 720; const sy = nh / pres.h; pres.h = nh; for (const s of pres.slides) for (const e of s.els) { e.y = Math.round(e.y * sy); if (e.type !== 'text') e.h = Math.round(e.h * sy); } commit(); renderAll(); }

  /* ---------- keyboard ---------- */
  document.addEventListener('keydown', (e) => {
    if (show || document.querySelector('.overlay')) return;
    const mod = e.ctrlKey || e.metaKey; const k = e.key;
    const inField = /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) || document.activeElement.isContentEditable;
    if (k === 'F5') { e.preventDefault(); startShow(e.shiftKey ? cur : 0); return; }
    if (mod && k.toLowerCase() === 'm') { e.preventDefault(); addSlide(); return; }
    if (editingId) { if (k === 'Escape') { e.preventDefault(); const id = editingId; exitEdit(); selIds = new Set([id]); renderSel(); } return; }
    if (inField) return;
    if ((k === 'Delete' || k === 'Backspace')) { e.preventDefault(); if (selIds.size) deleteSel(); else if (document.activeElement === document.body || thumbs.contains(document.activeElement)) delSlide(); return; }
    if (k.startsWith('Arrow') && selIds.size) { e.preventDefault(); const d = e.shiftKey ? 10 : 1; for (const el of selected()) { if (k === 'ArrowLeft') el.x -= d; if (k === 'ArrowRight') el.x += d; if (k === 'ArrowUp') el.y -= d; if (k === 'ArrowDown') el.y += d; } renderStage(); debCommit(); return; }
    if ((k === 'PageDown' || (k === 'ArrowDown' && !selIds.size))) { e.preventDefault(); goto(cur + 1); return; }
    if ((k === 'PageUp' || (k === 'ArrowUp' && !selIds.size))) { e.preventDefault(); goto(cur - 1); return; }
    if (k === 'Escape') { selIds.clear(); renderSel(); return; }
    if ((k === 'Enter' || k === 'F2') && selIds.size === 1) { const el = selected()[0]; if (el.text) { e.preventDefault(); enterEdit(el.id); } return; }
    if (mod) {
      const kk = k.toLowerCase();
      if (kk === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
      else if (kk === 'y') { e.preventDefault(); redo(); }
      else if (kk === 'c') { if (selIds.size) { e.preventDefault(); copySel(); } }
      else if (kk === 'x') { if (selIds.size) { e.preventDefault(); copySel(); deleteSel(); } }
      else if (kk === 'g') { e.preventDefault(); e.shiftKey ? ungroupSel() : groupSel(); }
      else if (kk === 'd') { e.preventDefault(); selIds.size ? duplicateSel() : dupSlide(); }
      else if (kk === 'a') { e.preventDefault(); selIds = new Set(slide().els.map((x) => x.id)); renderSel(); }
      else if (kk === 'b') { e.preventDefault(); txt((x) => (x.text.bold = !x.text.bold)); }
      else if (kk === 'i') { e.preventDefault(); txt((x) => (x.text.italic = !x.text.italic)); }
      return;
    }
    if (k.length === 1 && selIds.size === 1) { const el = selected()[0]; if (el.text && !el.text.html) { enterEdit(el.id); } }
  });
  const debCommit = debounce(() => { commit(); refreshThumb(); }, 400);
  document.addEventListener('paste', async (e) => {
    if (editingId || /INPUT|TEXTAREA/.test(document.activeElement.tagName) || show) {
      if (editingId && e.clipboardData.getData('text/html')) { e.preventDefault(); document.execCommand('insertText', false, e.clipboardData.getData('text/plain')); }
      return;
    }
    const f = [...e.clipboardData.files].find((x) => x.type.startsWith('image/'));
    if (f) { e.preventDefault(); const r = new FileReader(); r.onload = () => addImage(r.result); r.readAsDataURL(f); return; }
    const local = localStorage.getItem('rabbit.point.clip');
    const text = e.clipboardData.getData('text/plain');
    if (local && (!text || text.trim() === '')) { e.preventDefault(); pasteEls(); return; }
    if (text) { e.preventDefault(); const el = addEl({ id: uidp(), type: 'text', x: 160, y: 200, w: 960, h: 120, text: { html: esc(text).replace(/\n/g, '<br>'), size: 24, align: 'left' } }); }
  });
  document.addEventListener('copy', (e) => { if (!editingId && selIds.size && !/INPUT|TEXTAREA/.test(document.activeElement.tagName)) { copySel(); e.clipboardData.setData('text/plain', ' '); e.preventDefault(); } });
  stageWrap.addEventListener('dragover', (e) => e.preventDefault());
  stageWrap.addEventListener('drop', (e) => { e.preventDefault(); const f = [...e.dataTransfer.files].find((x) => x.type.startsWith('image/')); if (f) { const r = new FileReader(); r.onload = () => addImage(r.result); r.readAsDataURL(f); } });
  notes.addEventListener('input', () => { slide().notes = notes.value; shell.markDirty(); debCommit(); });

  /* ---------- view: sorter ---------- */
  let sorterEl = null;
  function renderSorter() {
    sorterEl?.remove(); if (viewMode !== 'sorter') return;
    sorterEl = h('div', { class: 'sorter' });
    pres.slides.forEach((s, i) => {
      const t = h('div', { class: `thumb ${i === cur ? 'on' : ''} ${s.hidden ? 'hidden-slide' : ''}`, draggable: 'true' });
      const sl = renderSlide(pres, s); t.append(sl);
      t.onclick = () => { cur = i; renderThumbs(); }; t.ondblclick = () => { viewMode = 'normal'; goto(i); };
      t.ondragstart = (e) => e.dataTransfer.setData('text/x-slide', String(i)); t.ondragover = (e) => e.preventDefault(); t.ondrop = (e) => { e.preventDefault(); moveSlide(+e.dataTransfer.getData('text/x-slide'), i); };
      sorterEl.append(h('div', { class: 'thumb-row' }, t, h('span', { class: 'small muted' }, `${i + 1}. ${LAYOUTS[s.layout]?.n || ''}`)));
    });
    stageWrap.append(sorterEl);
    requestAnimationFrame(() => $$('.thumb', sorterEl).forEach((t) => { t.firstChild.style.transform = `scale(${t.clientWidth / pres.w})`; }));
  }

  /* =================== SLIDESHOW =================== */
  let show = null;
  function startShow(from = 0) {
    exitEdit();
    const order = pres.slides.map((s, i) => i).filter((i) => !pres.slides[i].hidden || i === from);
    let idx = Math.max(0, order.indexOf(from)); let step = 0;
    const root = h('div', { class: 'show' }); const frame = h('div', { class: 'frame' }); const counter = h('div', { class: 'counter' });
    const bar = h('div', { class: 'bar' }, h('button', { html: icon('chevl'), onclick: (e) => { e.stopPropagation(); prev(); } }), h('button', { html: icon('chevr'), onclick: (e) => { e.stopPropagation(); next(); } }), h('button', { html: icon('presenter'), onclick: (e) => { e.stopPropagation(); openPresenter(); } }), h('button', { html: icon('close'), onclick: (e) => { e.stopPropagation(); end(); } }));
    root.append(frame, bar, counter); document.body.append(root);
    let curEl = null; let anims = [];
    const fitShow = () => { const k = Math.min(innerWidth / pres.w, innerHeight / pres.h); frame.style.transform = `scale(${k})`; frame.style.left = (innerWidth - pres.w * k) / 2 + 'px'; frame.style.top = (innerHeight - pres.h * k) / 2 + 'px'; frame.style.width = pres.w + 'px'; frame.style.height = pres.h + 'px'; };
    fitShow(); window.addEventListener('resize', fitShow);
    function render(i, dir = 1, revealAll = false) {
      const s = pres.slides[order[i]]; const el = renderSlide(pres, s);
      anims = s.els.filter((e) => e.anim && e.anim.type !== 'none').sort((a, b) => a.anim.order - b.anim.order);
      step = revealAll ? anims.length : 0;
      anims.forEach((a, j) => { if (j >= step) el.querySelector(`.el[data-id="${a.id}"]`)?.classList.add('anim-hidden'); });
      const tr = s.transition || { type: 'none' }; const dur = (tr.dur || 0.5) + 's';
      if (curEl && tr.type !== 'none' && dir > 0) {
        const anim = { fade: 't-fade', push: 't-push', wipe: 't-wipe', zoom: 't-zoom', cover: 't-cover', flip: 't-flip' }[tr.type];
        el.style.animation = `${anim} ${dur} cubic-bezier(.2,.8,.2,1) both`;
        if (tr.type === 'push') curEl.style.animation = `t-push-out ${dur} cubic-bezier(.2,.8,.2,1) both`;
        const old = curEl; setTimeout(() => old.remove(), (tr.dur || 0.5) * 1000 + 50);
      } else curEl?.remove();
      frame.append(el); curEl = el;
      counter.textContent = `${order[i] + 1} / ${pres.slides.length}`;
      channel?.postMessage({ type: 'state', pres, index: order[i], next: order[i + 1] ?? null, step });
      armAuto(); penClear();
      curEl.querySelectorAll('[data-autoplay]').forEach((m) => m.play().catch(() => {}));
    }
    function next() {
      if (step < anims.length) { const a = anims[step++]; const n = curEl.querySelector(`.el[data-id="${a.id}"]`); if (n) { n.classList.remove('anim-hidden'); const kf = 'a-' + a.anim.type; n.animate ? n.style.setProperty('animation', `${kf} ${a.anim.dur || 0.6}s cubic-bezier(.2,.8,.2,1) both`) : 0; } channel?.postMessage({ type: 'state', pres, index: order[idx], next: order[idx + 1] ?? null, step }); armAuto(); return; }
      if (idx < order.length - 1) { idx++; render(idx, 1); } else if (pres.show?.loop) { idx = 0; render(0, 1); } else end();
    }
    function prev() { if (idx > 0) { idx--; render(idx, -1, true); } }
    function goTo(n) { const i = order.indexOf(n); if (i >= 0) { idx = i; render(idx, 1); } }
    let blank = null; let laser = null; let typed = '';
    let autoT = null;
    const armAuto = () => { clearTimeout(autoT); const sl = pres.slides[order[idx]]; const sec = sl?.advance ?? pres.show?.auto; if (sec > 0) autoT = setTimeout(() => next(), sec * 1000); };
    let pen = null, penCtx = null, drawing = false;
    const penClear = () => { if (penCtx) penCtx.clearRect(0, 0, pen.width, pen.height); };
    const togglePen = () => { if (pen) { pen.remove(); pen = null; penCtx = null; root.style.cursor = ''; return; } pen = h('canvas', { style: { position: 'absolute', inset: '0', zIndex: 12, cursor: 'crosshair' } }); pen.width = innerWidth * devicePixelRatio; pen.height = innerHeight * devicePixelRatio; pen.style.width = '100%'; pen.style.height = '100%'; root.append(pen); penCtx = pen.getContext('2d'); penCtx.scale(devicePixelRatio, devicePixelRatio); penCtx.strokeStyle = '#ff2a2a'; penCtx.lineWidth = 4; penCtx.lineCap = penCtx.lineJoin = 'round';
      pen.addEventListener('pointerdown', (e) => { e.stopPropagation(); drawing = true; penCtx.beginPath(); penCtx.moveTo(e.clientX, e.clientY); });
      pen.addEventListener('pointermove', (e) => { if (!drawing) return; penCtx.lineTo(e.clientX, e.clientY); penCtx.stroke(); });
      pen.addEventListener('pointerup', (e) => { e.stopPropagation(); drawing = false; }); };
    const key = (e) => {
      const k = e.key; e.stopPropagation();
      if ((e.ctrlKey || e.metaKey) && (k === 'p' || k === 'з')) { e.preventDefault(); togglePen(); return; }
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n'].includes(k) && !typed) { e.preventDefault(); if (blank) { blank.remove(); blank = null; } else next(); }
      else if (k === 'Enter' && typed) { goTo(+typed - 1); typed = ''; }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(k)) { e.preventDefault(); prev(); }
      else if (k === 'Escape') end();
      else if (k === 'Home') { idx = 0; render(0); } else if (k === 'End') { idx = order.length - 1; render(idx, 1, true); }
      else if (k === 'b' || k === 'w' || k === 'и' || k === 'ц') { if (blank) { blank.remove(); blank = null; } else { blank = h('div', { class: 'blank', style: { background: k === 'b' || k === 'и' ? '#000' : '#fff' } }); root.append(blank); } }
      else if (k === 'e' || k === 'у') penClear();
      else if (k === 'l' || k === 'д') { if (laser) { laser.remove(); laser = null; } else { laser = h('div', { class: 'laser' }); root.append(laser); } }
      else if (/\d/.test(k)) typed += k;
    };
    let hideT;
    const mm = (e) => { root.classList.add('cursor'); clearTimeout(hideT); hideT = setTimeout(() => root.classList.remove('cursor'), 2000); if (laser) { laser.style.left = e.clientX + 'px'; laser.style.top = e.clientY + 'px'; } };
    let tx0 = null;
    root.addEventListener('pointerdown', (e) => { tx0 = { x: e.clientX, y: e.clientY }; });
    root.addEventListener('pointerup', (e) => { if (!tx0 || e.target.closest('button') || e.target.closest('video,audio') || pen) return; const dx = e.clientX - tx0.x; if (Math.abs(dx) > 60) { dx < 0 ? next() : prev(); } else if (e.button === 0) next(); tx0 = null; });
    root.addEventListener('contextmenu', (e) => { e.preventDefault(); prev(); });
    document.addEventListener('keydown', key, true); root.addEventListener('mousemove', mm);
    const onMsg = (m) => { const d = m.data; if (d.type === 'cmd') { if (d.cmd === 'next') next(); if (d.cmd === 'prev') prev(); if (d.cmd === 'end') end(); if (d.cmd === 'hello') render(idx, 0, step > 0); } };
    channel?.addEventListener('message', onMsg);
    function end() {
      clearTimeout(autoT);
      document.removeEventListener('keydown', key, true); window.removeEventListener('resize', fitShow); channel?.removeEventListener('message', onMsg);
      root.remove(); show = null; channel?.postMessage({ type: 'end' });
      if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
      goto(order[idx]);
    }
    show = { end };
    document.documentElement.requestFullscreen?.().catch(() => {});
    document.addEventListener('fullscreenchange', function fc() { if (!document.fullscreenElement && show) { document.removeEventListener('fullscreenchange', fc); } });
    render(idx);
  }
  function openPresenter() { window.open(location.pathname + '?presenter=1', 'rabbit-presenter', 'width=1100,height=700'); }

  /* =================== EXPORT =================== */
  function printSlides(pdf) {
    $('#print-slides')?.remove();
    const box = h('div', { id: 'print-slides' });
    for (const s of pres.slides) if (!s.hidden) box.append(renderSlide(pres, s));
    document.body.append(box);
    $('#print-page').textContent = `@page { size: ${pres.w}px ${pres.h}px; margin: 0; } @media print { html, body { width: ${pres.w}px; } .slide { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }`;
    const done = () => setTimeout(() => box.remove(), 1500);
    if (pdf) printToPdf(shell.title + '.pdf', { preferCSSPageSize: true, printBackground: true }).then((p) => { if (p) toast('PDF сохранён', 'ok'); done(); });
    else { window.print(); done(); }
  }
  function exportHtml() {
    const slidesHtml = pres.slides.filter((s) => !s.hidden).map((s) => { const el = renderSlide(pres, s); el.dataset.tr = s.transition?.type || 'none'; el.dataset.dur = s.transition?.dur || 0.5; el.dataset.anims = JSON.stringify(s.els.filter((e) => e.anim && e.anim.type !== 'none').sort((a, b) => a.anim.order - b.anim.order).map((e) => [e.id, e.anim.type, e.anim.dur || 0.6])); return el.outerHTML; }).join('');
    const css = [...document.styleSheets].filter((ss) => ss.href?.includes('point.css')).flatMap((ss) => { try { return [...ss.cssRules].map((r) => r.cssText); } catch { return []; } }).filter((t) => /^(\.slide|@keyframes)/.test(t)).join('\n');
    const player = `const S=[...document.querySelectorAll('.slide')];let i=0,st=0;const fit=()=>{const k=Math.min(innerWidth/${pres.w},innerHeight/${pres.h});S.forEach(s=>{s.style.transform='translate('+(innerWidth-${pres.w}*k)/2+'px,'+(innerHeight-${pres.h}*k)/2+'px) scale('+k+')'})};
const an=s=>JSON.parse(s.dataset.anims||'[]');function show(n,dir){S.forEach((s,j)=>s.style.display=j===n?'block':'none');const s=S[n];st=dir<0?an(s).length:0;an(s).forEach((a,j)=>{const e=s.querySelector('[data-id="'+a[0]+'"]');if(e){e.style.visibility=j<st?'visible':'hidden';e.style.animation=''}});const t={fade:'t-fade',push:'t-push',wipe:'t-wipe',zoom:'t-zoom',cover:'t-cover',flip:'t-flip'}[s.dataset.tr];s.style.animation=dir>0&&t?t+' '+s.dataset.dur+'s cubic-bezier(.2,.8,.2,1) both':'';i=n}
function next(){const a=an(S[i]);if(st<a.length){const x=a[st++];const e=S[i].querySelector('[data-id="'+x[0]+'"]');e.style.visibility='visible';e.style.animation='a-'+x[1]+' '+x[2]+'s cubic-bezier(.2,.8,.2,1) both';return}if(i<S.length-1)show(i+1,1)}function prev(){if(i>0)show(i-1,-1)}
addEventListener('keydown',e=>{if(['ArrowRight','ArrowDown',' ','PageDown','Enter'].includes(e.key))next();if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(e.key))prev();if(e.key==='f')document.documentElement.requestFullscreen()});addEventListener('click',next);addEventListener('contextmenu',e=>{e.preventDefault();prev()});addEventListener('resize',fit);fit();show(0,0);`;
    const doc = `<!doctype html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(shell.title)}</title><style>html,body{margin:0;height:100%;background:#000;overflow:hidden}.slide{position:absolute;left:0;top:0;transform-origin:0 0}.el{position:absolute;box-sizing:border-box}.shape-svg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}.el-text{position:absolute;inset:0;display:flex;flex-direction:column;padding:10px 14px;overflow:hidden}.tx{line-height:1.2}.tx p,.tx div{margin:0}.tx ul,.tx ol{margin:0;padding-left:1.2em}.el img{width:100%;height:100%;display:block}.el table{width:100%;height:100%;border-collapse:collapse}.el td{border:1px solid rgba(0,0,0,.18);padding:6px 10px}.chart-holder{position:absolute;inset:0}${css}</style></head><body>${slidesHtml}<script>${player}<\/script></body></html>`;
    shell.export(shell.title + '.html', doc, 'text/html');
  }

  /* =================== ribbon =================== */
  const FONTS = ['Segoe UI', 'Calibri', 'Arial', 'Georgia', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Roboto', 'Garamond', 'Courier New', 'Impact'];
  const SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 88, 96];
  const b = (id, ic, label, action, extra = {}) => ({ type: 'btn', id, icon: ic, label, action, ...extra });
  const big = (id, ic, label, action, extra = {}) => ({ type: 'big', id, icon: ic, label, action, ...extra });
  const shapeMenu = () => Object.entries(SHAPES).map(([k, s]) => ({ label: s.n, action: () => addShape(k) }));
  const layoutMenu = (fn) => () => Object.entries(LAYOUTS).map(([k, l]) => ({ label: l.n, action: () => fn(k) }));
  const themeGallery = h('div', { class: 'theme-gallery' });
  function drawThemeGallery() { themeGallery.innerHTML = ''; for (const t of THEMES) { const bt = h('button', { class: pres.theme.id === t.id ? 'on' : '', 'data-tip': t.name, style: { background: t.bg }, html: `<b style="color:${t.fg};font-family:${t.title}">Аа</b><i style="background:${t.accent}"></i>` }); bt.addEventListener('mousedown', (e) => e.preventDefault()); bt.onclick = () => { setTheme(t.id); drawThemeGallery(); }; themeGallery.append(bt); } }
  const tableOp = (op) => { const el = selected().find((e) => e.type === 'table'); if (!el) return; exitEdit(); const rows = el.table.rows; const cols = rows[0].length; if (op === 'row') { rows.push(Array(cols).fill('')); el.h += 50; } if (op === 'col') rows.forEach((r) => r.push('')); if (op === 'delrow' && rows.length > 1) { rows.pop(); el.h = Math.max(40, el.h - 50); } if (op === 'delcol' && cols > 1) rows.forEach((r) => r.pop()); if (op === 'hdr') el.table.header = !el.table.header; if (op === 'band') el.table.band = !el.table.band; commit(); renderStage(); refreshThumb(); };
  const ribbon = [
    { id: 'home', label: 'Главная', groups: [
      { label: 'Буфер обмена', items: [big('paste', 'paste', 'Вставить', pasteEls), { type: 'col', items: [b('cut', 'cut', 'Вырезать', () => { copySel(); deleteSel(); }, { key: 'Ctrl+X' }), b('copy', 'copy', 'Копировать', copySel, { key: 'Ctrl+C' }), b('dup', 'copy', 'Дублировать', duplicateSel, { key: 'Ctrl+D' })] }] },
      { label: 'Слайды', items: [big('newslide', 'slide-new', 'Создать слайд', null, { menu: layoutMenu((k) => addSlide(k)) }), { type: 'col', items: [b('layout', 'layout', 'Макет', null, { showLabel: true, menu: layoutMenu(applyLayout) }), b('dupslide', 'copy', 'Дублировать', dupSlide, { showLabel: true }), b('delslide', 'trash', 'Удалить', delSlide, { showLabel: true })] }] },
      { label: 'Шрифт', items: [{ type: 'col', items: [
        { type: 'row', items: [{ type: 'select', id: 'font', width: 130, options: FONTS, value: 'Segoe UI', action: (v) => txt((e) => (e.text.font = v), () => document.execCommand('fontName', false, v)) }, { type: 'select', id: 'size', editable: true, width: 52, options: SIZES, value: '24', action: setFontSize }, b('grow', 'text-size', 'Увеличить', () => bump(1)), b('shrink', 'type', 'Уменьшить', () => bump(-1))] },
        { type: 'row', items: [b('bold', 'bold', 'Полужирный', () => txt((e) => (e.text.bold = !e.text.bold), () => document.execCommand('bold')), { key: 'Ctrl+B' }), b('italic', 'italic', 'Курсив', () => txt((e) => (e.text.italic = !e.text.italic), () => document.execCommand('italic')), { key: 'Ctrl+I' }), b('underline', 'underline', 'Подчёркнутый', () => txt((e) => (e.text.html = e.text.html ? `<u>${e.text.html}</u>` : e.text.html), () => document.execCommand('underline'))), b('strike', 'strike', 'Зачёркнутый', () => txt(() => {}, () => document.execCommand('strikeThrough'))), b('clearf', 'clear-format', 'Очистить формат', () => txt((e) => { e.text.color = undefined; e.text.bold = false; e.text.italic = false; e.text.font = undefined; const d = document.createElement('div'); d.innerHTML = e.text.html; e.text.html = d.innerText.replace(/\n/g, '<br>'); }, () => document.execCommand('removeFormat'))),
          { type: 'color', id: 'tc', icon: 'font-color', tip: 'Цвет текста', value: '#d2552d', action: (c) => txt((e) => { e.text.color = c; if (e.type === 'table') e.color = c; }, () => document.execCommand('foreColor', false, c || '#000')) }] }] }] },
      { label: 'Абзац', items: [{ type: 'col', items: [
        { type: 'row', items: [b('ul', 'list', 'Маркеры', () => txt((e) => { e.text.html = toggleList(e.text.html, 'ul'); }, () => document.execCommand('insertUnorderedList'))), b('ol', 'list-ordered', 'Нумерация', () => txt((e) => { e.text.html = toggleList(e.text.html, 'ol'); }, () => document.execCommand('insertOrderedList'))), b('lh', 'line-spacing', 'Междустрочный интервал', null, { menu: () => [1, 1.15, 1.3, 1.5, 2].map((v) => ({ label: String(v).replace('.', ','), action: () => txt((e) => (e.text.lh = v)) })) })] },
        { type: 'row', items: [b('al', 'align-left', 'По левому краю', () => txt((e) => (e.text.align = 'left'), () => document.execCommand('justifyLeft'))), b('ac', 'align-center', 'По центру', () => txt((e) => (e.text.align = 'center'), () => document.execCommand('justifyCenter'))), b('ar', 'align-right', 'По правому краю', () => txt((e) => (e.text.align = 'right'), () => document.execCommand('justifyRight'))), b('aj', 'align-justify', 'По ширине', () => txt((e) => (e.text.align = 'justify'), () => document.execCommand('justifyFull'))),
          b('vt', 'valign-top', 'Сверху', () => txt((e) => (e.text.valign = 'top'))), b('vm', 'valign-middle', 'По центру', () => txt((e) => (e.text.valign = 'middle'))), b('vb', 'valign-bottom', 'Снизу', () => txt((e) => (e.text.valign = 'bottom')))] }] }] },
      { label: 'Рисование', items: [big('shapes', 'shapes', 'Фигуры', null, { menu: shapeMenu }), big('textbox', 'type', 'Надпись', addText), { type: 'col', items: [b('arr', 'bring-front', 'Упорядочить', null, { showLabel: true, menu: () => [{ label: 'На передний план', icon: 'bring-front', action: () => zorder('front') }, { label: 'Переместить вперёд', action: () => zorder('up') }, { label: 'Переместить назад', action: () => zorder('down') }, { label: 'На задний план', icon: 'send-back', action: () => zorder('back') }, '-', { title: 'Выровнять' }, { label: 'По левому краю', action: () => align('l') }, { label: 'По центру', action: () => align('c') }, { label: 'По правому краю', action: () => align('r') }, { label: 'По верхнему краю', action: () => align('t') }, { label: 'По середине', action: () => align('m') }, { label: 'По нижнему краю', action: () => align('b') }, { label: 'Распределить по горизонтали', action: () => align('dh') }, { label: 'Распределить по вертикали', action: () => align('dv') }] }),
        { type: 'color', id: 'fill', icon: 'fill-color', tip: 'Заливка фигуры', value: '#2f62c9', allowNone: true, action: (c) => styleEls((e) => { if (e.type === 'shape') { e.fill = c || 'none'; e.fillFromTheme = false; } }) },
        { type: 'color', id: 'stroke', icon: 'edit', tip: 'Контур фигуры', value: '#1f1f1f', allowNone: true, action: (c) => styleEls((e) => { e.stroke = c || 'none'; if (!e.strokeW) e.strokeW = 2; }) }] }] },
    ] },
    { id: 'insert', label: 'Вставка', groups: [
      { label: 'Слайды', items: [big('ns2', 'slide-new', 'Создать слайд', null, { menu: layoutMenu((k) => addSlide(k)) })] },
      { label: 'Таблицы', items: [big('tbl', 'table', 'Таблица', null, { menu: () => [[2, 2], [3, 3], [3, 4], [4, 4], [5, 3], [6, 5]].map(([r, c]) => ({ label: `${c} × ${r}`, action: () => addTable(r, c) })) })] },
      { label: 'Изображения', items: [big('img', 'image', 'Рисунок', () => addImage())] },
      { label: 'Иллюстрации', items: [big('shapes2', 'shapes', 'Фигуры', null, { menu: shapeMenu }), big('chart', 'chart', 'Диаграмма', null, { menu: () => CHART_TYPES.map((t) => ({ label: t.name, action: () => addChart(t.id) })) })] },
      { label: 'Текст', items: [big('tb2', 'type', 'Надпись', addText), big('date', 'calendar', 'Дата', () => { const el = addEl({ id: uidp(), type: 'text', x: 60, y: pres.h - 60, w: 300, h: 40, text: { html: new Date().toLocaleDateString('ru-RU'), size: 16, align: 'left' } }); }), big('num', 'hash', 'Номера слайдов', () => { for (const [i, s] of pres.slides.entries()) { s.els = s.els.filter((e) => !e.slideNum); s.els.push({ id: uidp(), type: 'text', slideNum: true, x: pres.w - 140, y: pres.h - 56, w: 100, h: 36, text: { html: String(i + 1), size: 16, align: 'right' } }); } commit(); renderAll(); })] },
    ] },
    { id: 'design', label: 'Конструктор', groups: [
      { label: 'Темы', items: [{ type: 'custom', el: themeGallery }] },
      { label: 'Настроить', items: [big('size', 'page-size', 'Размер слайда', null, { menu: () => [{ label: 'Широкоэкранный (16:9)', checked: pres.h === 720, action: () => setSize('16:9') }, { label: 'Стандартный (4:3)', checked: pres.h === 960, action: () => setSize('4:3') }] }), big('bg', 'background', 'Формат фона', bgDialog),
        { type: 'col', items: [{ type: 'color', id: 'accent', icon: 'palette', tip: 'Цвет акцента темы', value: '#d2552d', action: (c) => { pres.theme.accent = c; commit(); renderAll(); } }, { type: 'color', id: 'fgc', icon: 'font-color', tip: 'Цвет текста темы', value: '#262a33', action: (c) => { pres.theme.fg = c; commit(); renderAll(); } }] }] },
    ] },
    { id: 'trans', label: 'Переходы', groups: [
      { label: 'Переход к этому слайду', items: [['none', 'Нет', 'close'], ['fade', 'Выцветание', 'transition'], ['push', 'Сдвиг', 'arrow-right'], ['wipe', 'Стирание', 'eraser'], ['zoom', 'Масштаб', 'zoom-in'], ['cover', 'Наплыв', 'arrow-up'], ['flip', 'Поворот', 'flip']].map(([id, n, ic]) => big('tr-' + id, ic, n, () => { slide().transition = { ...(slide().transition || {}), type: id }; commit(); renderThumbs(); syncRibbon(); toast(`Переход: ${n}`, 'ok', 1000); })) },
      { label: 'Время', items: [{ type: 'col', items: [{ type: 'select', id: 'trdur', width: 110, options: [{ value: '0.3', label: 'Быстро 0,3 с' }, { value: '0.5', label: 'Обычно 0,5 с' }, { value: '0.8', label: 'Медленно 0,8 с' }, { value: '1.2', label: 'Очень медл. 1,2 с' }], value: '0.5', action: (v) => { slide().transition = { ...(slide().transition || { type: 'fade' }), dur: +v }; commit(); } }, b('trall', 'copy', 'Применить ко всем', () => { for (const s of pres.slides) s.transition = { ...slide().transition }; commit(); renderThumbs(); toast('Переход применён ко всем слайдам', 'ok'); }, { showLabel: true })] }, big('trprev', 'play', 'Просмотр', () => startShow(cur))] },
    ] },
    { id: 'anim', label: 'Анимация', groups: [
      { label: 'Анимация входа', items: ANIMS.map(([id, n]) => big('an-' + id, id === 'none' ? 'close' : 'animation', n, () => setAnim(id))) },
      { label: 'Порядок', items: [{ type: 'col', items: [b('anup', 'chevu', 'Раньше', () => reorderAnim(-1), { showLabel: true }), b('andn', 'chevd', 'Позже', () => reorderAnim(1), { showLabel: true })] }, { type: 'select', id: 'andur', width: 110, options: [{ value: '0.4', label: '0,4 с' }, { value: '0.6', label: '0,6 с' }, { value: '1', label: '1 с' }, { value: '1.5', label: '1,5 с' }], value: '0.6', action: (v) => styleEls((e) => { if (e.anim) e.anim.dur = +v; }) }] },
    ] },
    { id: 'show', label: 'Слайд-шоу', groups: [
      { label: 'Начать', items: [big('fromstart', 'play', 'С начала', () => startShow(0), { key: 'F5' }), big('fromcur', 'play-circle', 'С текущего слайда', () => startShow(cur), { key: 'Shift+F5' }), big('presenter', 'presenter', 'Режим докладчика', () => { openPresenter(); startShow(cur); })] },
      { label: 'Настройка', items: [big('hide', 'eye-off', 'Скрыть слайд', () => { slide().hidden = !slide().hidden; commit(); renderThumbs(); syncRibbon(); })] },
      { label: 'Подсказки', items: [{ type: 'custom', el: h('div', { class: 'small muted', style: { maxWidth: '260px', lineHeight: '1.5', padding: '4px' }, html: '→ / пробел — далее, ← — назад, B — чёрный экран, W — белый, L — лазерная указка, Ctrl+P — перо, E — стереть рисунок, номер + Enter — переход, Esc — выход' }) }] },
    ] },
    { id: 'view', label: 'Вид', groups: [
      { label: 'Режимы', items: [big('vnormal', 'slides', 'Обычный', () => { viewMode = 'normal'; sorterEl?.remove(); sorterEl = null; syncRibbon(); }), big('vsorter', 'grid', 'Сортировщик', () => { viewMode = 'sorter'; renderSorter(); syncRibbon(); }), big('vnotes', 'notes', 'Заметки', () => { notesPane.classList.toggle('hidden'); syncRibbon(); setTimeout(fit, 0); })] },
      { label: 'Показ', items: [big('vguides', 'ruler', 'Направляющие', () => { showGuides = !showGuides; syncRibbon(); })] },
      { label: 'Масштаб', items: [big('zfit', 'fit', 'По размеру окна', () => { zoomMul = 1; fit(); }), { type: 'col', items: [b('zin', 'zoom-in', 'Увеличить', () => { zoomMul = Math.min(3, zoomMul * 1.15); fit(); }, { showLabel: true }), b('zout', 'zoom-out', 'Уменьшить', () => { zoomMul = Math.max(0.3, zoomMul / 1.15); fit(); }, { showLabel: true })] }] },
    ] },
    { id: 'shapefmt', label: 'Формат фигуры', contextual: true, groups: [
      { label: 'Стили фигур', items: [{ type: 'custom', el: (() => { const g = h('div', { class: 'theme-gallery' }); const presets = [['#2f62c9', null], ['#d2552d', null], ['#1d8a52', null], ['#7b4fd6', null], ['transparent', '#262a33'], ['#ffffff', '#c9ccd3'], ['#262a33', null]]; for (const [f, s] of presets) { const bt = h('button', { style: { background: f === 'transparent' ? 'var(--surface)' : f, boxShadow: s ? `inset 0 0 0 2px ${s}` : '' } }); bt.addEventListener('mousedown', (e) => e.preventDefault()); bt.onclick = () => styleEls((e) => { if (e.type !== 'shape') return; e.fill = f === 'transparent' ? 'none' : f; e.fillFromTheme = false; e.stroke = s || 'none'; e.strokeW = s ? 2 : 0; if (e.text) e.text.color = f === '#ffffff' || f === 'transparent' ? '#262a33' : '#ffffff'; }); g.append(bt); } return g; })() }] },
      { label: 'Контур', items: [{ type: 'col', items: [{ type: 'select', id: 'sw', width: 100, options: [0, 1, 2, 3, 4, 6, 8, 12].map((v) => ({ value: v, label: v ? v + ' пт' : 'Нет' })), value: '0', action: (v) => styleEls((e) => { e.strokeW = +v; if (+v && (!e.stroke || e.stroke === 'none')) e.stroke = pres.theme.fg; }) }, b('shadow', 'square', 'Тень', () => styleEls((e) => (e.shadow = !e.shadow)), { showLabel: true }), { type: 'select', id: 'opacity', width: 100, options: [100, 90, 75, 50, 25].map((v) => ({ value: v / 100, label: `Непрозр. ${v}%` })), value: '1', action: (v) => styleEls((e) => (e.opacity = +v)) }] }] },
      { label: 'Упорядочение', items: [big('front', 'bring-front', 'Вперёд', () => zorder('front')), big('back', 'send-back', 'Назад', () => zorder('back')), { type: 'col', items: [b('rot90', 'rotate', 'Повернуть на 90°', () => styleEls((e) => (e.rot = ((e.rot || 0) + 90) % 360)), { showLabel: true }), b('fliph', 'flip', 'Отразить', () => styleEls((e) => (e.flipH = !e.flipH)), { showLabel: true }), b('alignc', 'align-objects', 'Выровнять', null, { showLabel: true, menu: () => [['l', 'По левому краю'], ['c', 'По центру'], ['r', 'По правому краю'], ['t', 'По верху'], ['m', 'По середине'], ['b', 'По низу'], ['dh', 'Распределить по горизонтали'], ['dv', 'Распределить по вертикали']].map(([k, n]) => ({ label: n, action: () => align(k) })) })] }] },
      { label: 'Объект', items: [big('chdata', 'chart', 'Данные диаграммы', () => { const e = selected().find((x) => x.type === 'chart'); if (e) chartDialog(e); }), big('imgrep', 'image', 'Заменить рисунок', () => { const e = selected().find((x) => x.type === 'image'); if (e) replaceImage(e); }), big('delel', 'trash', 'Удалить', deleteSel)] },
      { label: 'Таблица', items: [{ type: 'col', items: [b('trow', 'insert-row', 'Добавить строку', () => tableOp('row'), { showLabel: true }), b('tcol', 'insert-col', 'Добавить столбец', () => tableOp('col'), { showLabel: true }), b('thdr', 'heading', 'Строка заголовка', () => tableOp('hdr'), { showLabel: true })] }, { type: 'col', items: [b('tdr', 'delete-row', 'Удалить строку', () => tableOp('delrow'), { showLabel: true }), b('tdc', 'delete-col', 'Удалить столбец', () => tableOp('delcol'), { showLabel: true }), b('tband', 'menu', 'Чередование', () => tableOp('band'), { showLabel: true })] }] },
    ] },
  ];
  /* ---------- groups ---------- */
  function expandGroups(ids) { const set = new Set(ids); for (const id of ids) { const g = elById(id)?.group; if (g) for (const e of slide().els) if (e.group === g) set.add(e.id); } return set; }
  function groupSel() { const els = selected(); if (els.length < 2) return toast('Выделите несколько объектов (Shift+клик)'); const g = uidp(); els.forEach((e) => (e.group = g)); commit(); renderSel(); toast('Объекты сгруппированы', 'ok', 1200); }
  function ungroupSel() { const els = selected(); if (!els.some((e) => e.group)) return toast('В выделении нет групп'); els.forEach((e) => delete e.group); commit(); renderSel(); toast('Группа разделена', 'ok', 1200); }

  /* ---------- media & icons ---------- */
  async function addMedia(kind) {
    const f = await pickFile(kind === 'audio' ? 'audio/*' : 'video/*'); if (!f) return;
    if (f.size > 60e6) return toast('Файл больше 60 МБ — сожмите его перед вставкой', 'error');
    const src = await readFileAs(f, 'dataurl');
    addEl(kind === 'audio' ? { id: uidp(), type: 'media', kind, src, x: pres.w - 150, y: pres.h - 140, w: 100, h: 100, autoplay: true } : { id: uidp(), type: 'media', kind, src, x: 240, y: 120, w: 800, h: 450, autoplay: false });
    toast(kind === 'audio' ? 'Звук будет играть при показе слайда' : 'Видео воспроизводится в режиме показа', 'ok');
  }
  const ICONS = ['⭐', '✅', '❌', '⚡', '🔥', '💡', '🚀', '🎯', '📈', '📉', '📊', '💰', '🛒', '📦', '🔒', '🔔', '📌', '📎', '🧩', '⚙️', '🏆', '❤️', '👍', '👥', '🌍', '📱', '💻', '☁️', '🕒', '📅', '✉️', '☎️', '🎓', '🏠', '🐇'];
  function addIcon(ch) { addEl({ id: uidp(), type: 'text', x: pres.w / 2 - 70, y: pres.h / 2 - 70, w: 140, h: 140, text: { html: ch, size: 96, align: 'center', valign: 'middle' } }); }
  function iconPicker(anchor) {
    const g = h('div', { style: { display: 'grid', gridTemplateColumns: 'repeat(7, 36px)', gap: '4px' } });
    for (const ch of ICONS) g.append(h('button', { class: 'btn-r', style: { fontSize: '22px', width: '36px', height: '36px', justifyContent: 'center' }, onclick: () => { closeMenus(); addIcon(ch); } }, ch));
    popover(anchor, g);
  }

  /* ---------- shape & picture effects ---------- */
  function gradientMenu() {
    const presets = [['Закат', '#ff8a5b', '#d2552d'], ['Океан', '#3b82f6', '#06b6d4'], ['Лес', '#22c55e', '#0f766e'], ['Лаванда', '#a78bfa', '#ec4899'], ['Графит', '#4b5563', '#111827'], ['Золото', '#fde68a', '#f59e0b']];
    return [...presets.map(([n, a, b]) => ({ label: n, action: () => styleEls((e) => { if (e.type === 'shape') { e.grad = [a, b, 'v']; e.fillFromTheme = false; } }) })), { label: 'Свои цвета…', action: async () => { const r = await formDlg('Градиент', [{ key: 'a', label: 'Цвет 1', type: 'color', value: '#ff8a5b' }, { key: 'b', label: 'Цвет 2', type: 'color', value: '#7b4fd6' }, { key: 'dir', label: 'Направление', type: 'select', options: [{ value: 'v', label: 'Сверху вниз' }, { value: 'h', label: 'Слева направо' }] }]); if (r) styleEls((e) => { if (e.type === 'shape') e.grad = [r.a, r.b, r.dir]; }); } }, '-', { label: 'Сплошная заливка', action: () => styleEls((e) => delete e.grad) }];
  }
  const FILTER_NAMES = [['none', 'Без эффекта'], ['grayscale', 'Чёрно-белый'], ['sepia', 'Сепия'], ['bright', 'Ярче'], ['dark', 'Темнее'], ['contrast', 'Контрастнее'], ['saturate', 'Насыщеннее'], ['blur', 'Размытие']];

  /* ---------- PNG export (slide DOM → SVG foreignObject → canvas) ---------- */
  function slideCss() {
    const base = '.slide{position:relative;overflow:hidden}.el{position:absolute;box-sizing:border-box}.shape-svg{position:absolute;inset:0;width:100%;height:100%;overflow:visible}.el-text{position:absolute;inset:0;display:flex;flex-direction:column;padding:10px 14px;overflow:hidden}.t-shape .el-text{padding:12px 18px}.tx{line-height:1.2;overflow-wrap:break-word}.tx p,.tx div{margin:0}.tx ul,.tx ol{margin:0;padding-left:1.2em}.tx li{margin:.25em 0}.el img{width:100%;height:100%;display:block}.el table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}.el td{border:1px solid rgba(0,0,0,.18);padding:6px 10px;vertical-align:middle}.chart-holder{position:absolute;inset:0}table.hdr tr:first-child td{font-weight:700;color:#fff;background:var(--accent-c)}';
    return base;
  }
  async function slidePng(s, k = 1.5) {
    const el = renderSlide(pres, s);
    el.querySelectorAll('video').forEach((v) => v.replaceWith(h('div', { style: { position: 'absolute', inset: '0', background: '#000', display: 'grid', placeItems: 'center', color: '#fff', fontSize: '64px' } }, '▶')));
    const xml = new XMLSerializer().serializeToString(el);
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${pres.w * k}" height="${pres.h * k}"><foreignObject x="0" y="0" width="${pres.w * k}" height="${pres.h * k}"><div xmlns="http://www.w3.org/1999/xhtml" style="transform:scale(${k});transform-origin:0 0;width:${pres.w}px;height:${pres.h}px"><style>${slideCss()}</style>${xml}</div></foreignObject></svg>`;
    const img = new Image(); img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); await img.decode();
    const c = document.createElement('canvas'); c.width = Math.round(pres.w * k); c.height = Math.round(pres.h * k); c.getContext('2d').drawImage(img, 0, 0);
    const blob = await new Promise((r) => c.toBlob(r, 'image/png')); return new Uint8Array(await blob.arrayBuffer());
  }
  async function exportPng(all) {
    exitEdit(); toast('Готовлю изображения…', 'info', 1500);
    if (!all) return shell.export(`${shell.title} - слайд ${cur + 1}.png`, await slidePng(slide()), 'image/png');
    const files = []; for (const [i, s] of pres.slides.entries()) if (!s.hidden) files.push({ name: `Слайд ${String(i + 1).padStart(2, '0')}.png`, data: await slidePng(s) });
    shell.export(shell.title + ' (слайды).zip', await zipCreate(files), 'application/zip');
  }

  /* ---------- show settings ---------- */
  async function showSettings() {
    const sh = pres.show || {};
    const r = await formDlg('Настройка показа', [{ key: 'loop', label: 'Непрерывный цикл до нажатия Esc (киоск)', type: 'checkbox', value: sh.loop }, { key: 'auto', label: 'Автоматическая смена слайдов, секунд (0 — по щелчку)', type: 'number', value: sh.auto || 0, min: 0, max: 600 }], 'OK');
    if (!r) return; pres.show = { loop: r.loop, auto: Math.max(0, +r.auto || 0) }; commit(); toast('Настройки показа сохранены', 'ok');
  }

  /* ---------- extra ribbon controls ---------- */
  {
    const tab = (id) => ribbon.find((t) => t.id === id);
    const grp = (tid, label) => tab(tid).groups.find((g) => g.label === label);
    grp('insert', 'Изображения').items.push(big('icons', 'star', 'Значки', (e, el) => iconPicker(el)));
    tab('insert').groups.push({ label: 'Мультимедиа', items: [big('video', 'play-circle', 'Видео', () => addMedia('video')), big('audio', 'bubble', 'Звук', () => addMedia('audio'))] });
    grp('shapefmt', 'Стили фигур').items.push({ type: 'col', items: [b('grad', 'palette', 'Градиент', null, { showLabel: true, menu: gradientMenu }), b('imgfx', 'image', 'Эффекты рисунка', null, { showLabel: true, menu: () => FILTER_NAMES.map(([k, n]) => ({ label: n, action: () => styleEls((e) => { if (e.type === 'image') e.filter = k === 'none' ? undefined : k; }) })) }), b('imgfit', 'fit', 'Заполнение рисунка', null, { showLabel: true, menu: () => [['fill', 'Растянуть'], ['cover', 'Заполнить с обрезкой'], ['contain', 'Вписать']].map(([k, n]) => ({ label: n, action: () => styleEls((e) => { if (e.type === 'image') e.fit = k === 'fill' ? undefined : k; }) })) })] });
    grp('shapefmt', 'Упорядочение').items.push({ type: 'col', items: [b('grp', 'group', 'Группировать', groupSel, { showLabel: true, key: 'Ctrl+G' }), b('ungrp', 'grid', 'Разгруппировать', ungroupSel, { showLabel: true, key: 'Ctrl+Shift+G' })] });
    grp('show', 'Настройка').items.push(big('showset', 'settings', 'Настройка показа', showSettings));
    grp('trans', 'Время').items.push({ type: 'select', id: 'advance', width: 150, options: [{ value: '', label: 'Смена по щелчку' }, ...[3, 5, 10, 15, 30, 60].map((v) => ({ value: String(v), label: `Через ${v} с` }))], value: '', tip: 'Автоматическая смена этого слайда', action: (v) => { if (v) slide().advance = +v; else delete slide().advance; commit(); } });
    grp('home', 'Рисование').items.push(big('grp2', 'group', 'Группа', null, { menu: () => [{ label: 'Группировать', hint: 'Ctrl+G', action: groupSel }, { label: 'Разгруппировать', hint: 'Ctrl+Shift+G', action: ungroupSel }] }));
  }

  function reorderAnim(d) { const el = selected()[0]; if (!el?.anim) return; const list = slide().els.filter((e) => e.anim).sort((a, b) => a.anim.order - b.anim.order); const i = list.indexOf(el); const j = clamp(i + d, 0, list.length - 1); list.splice(i, 1); list.splice(j, 0, el); list.forEach((e, k) => (e.anim.order = k + 1)); commit(); renderStage(); }
  function toggleList(html, tag) { const d = document.createElement('div'); d.innerHTML = html; const list = d.querySelector('ul,ol'); if (list && list.tagName.toLowerCase() === tag) { return [...list.children].map((li) => li.innerHTML).join('<br>'); } const lines = (list ? [...list.children].map((li) => li.innerHTML) : d.innerHTML.split(/<br\s*\/?>|<\/div><div>|<\/p><p>/i)).map((l) => l.replace(/^<(div|p)>|<\/(div|p)>$/g, '')).filter((l) => l.trim()); return `<${tag}>${(lines.length ? lines : ['']).map((l) => `<li>${l || '<br>'}</li>`).join('')}</${tag}>`; }
  function syncRibbon() {
    if (!shell) return;
    const els = selected(); const one = els[0];
    shell.ribbon.showTab('shapefmt', els.length > 0 && !editingId, false);
    if (one?.text) { shell.ribbon.set('size', { value: String(one.text.size || 24) }); shell.ribbon.set('font', { value: one.text.font || (one.ph === 'title' ? pres.theme.title : pres.theme.body) }); shell.ribbon.set('bold', { on: !!one.text.bold }); shell.ribbon.set('italic', { on: !!one.text.italic }); for (const [id, a] of [['al', 'left'], ['ac', 'center'], ['ar', 'right'], ['aj', 'justify']]) shell.ribbon.set(id, { on: one.text.align === a }); }
    const tr = slide()?.transition?.type || 'none'; for (const t of ['none', 'fade', 'push', 'wipe', 'zoom', 'cover', 'flip']) shell.ribbon.set('tr-' + t, { on: tr === t });
    shell.ribbon.set('trdur', { value: String(slide()?.transition?.dur || 0.5) });
    shell.ribbon.set('advance', { value: String(slide()?.advance || '') });
    for (const [id] of ANIMS) shell.ribbon.set('an-' + id, { on: one ? (one.anim?.type || 'none') === id : false });
    shell.ribbon.set('hide', { on: !!slide()?.hidden }); shell.ribbon.set('vnotes', { on: !notesPane.classList.contains('hidden') }); shell.ribbon.set('vguides', { on: showGuides });
    shell.ribbon.set('vnormal', { on: viewMode === 'normal' }); shell.ribbon.set('vsorter', { on: viewMode === 'sorter' });
    drawThemeGallery();
    updUndo();
  }

  /* =================== templates =================== */
  function buildTemplate(id) {
    const themeId = { midnight: 'midnight', pitch: 'pitch', minimal: 'minimal' }[id] || 'office';
    pres = newPres(themeId);
    const T = (s, i, html) => { s.els[i].text.html = html; };
    if (id === 'pitch') {
      T(pres.slides[0], 0, 'Кролик&nbsp;Технологии'); T(pres.slides[0], 1, 'Делаем офисную работу быстрее · Раунд A 2026');
      const add = (lay, title, body) => { const s = makeSlide(pres, lay); T(s, 0, title); if (body && s.els[1]) T(s, 1, body); pres.slides.push(s); return s; };
      add('content', 'Проблема', '<ul><li>Офисный софт дорогой и требует подписки</li><li>Данные компании хранятся в чужом облаке</li><li>Совместная работа неудобна</li></ul>');
      add('content', 'Решение', '<ul><li>Бесплатный офис на всех платформах</li><li>Собственный сервер за 5 минут</li><li>Доски для совместной работы в реальном времени</li></ul>');
      const m = add('titleonly', 'Рынок'); m.els.push({ id: uidp(), type: 'chart', x: 140, y: 180, w: 1000, h: 480, chart: { type: 'column', title: 'Пользователи, млн', data: 'Год;Пользователи\n2024;1,2\n2025;3,4\n2026;7,9\n2027;14' } });
      add('two', 'Бизнес-модель', '<ul><li>Бесплатно для всех</li><li>Платная поддержка</li></ul>');
      add('section', 'Команда', 'Опытные инженеры и дизайнеры');
      add('quote', '«Лучший офис — тот, который вам принадлежит»', '— Команда Rabbit');
      add('title', 'Спасибо!', 'hello@rabbit.example');
      pres.slides.forEach((s) => (s.transition = { type: 'push', dur: 0.6 }));
    } else if (id === 'midnight' || id === 'minimal') {
      T(pres.slides[0], 0, id === 'midnight' ? 'Ночная презентация' : 'Простота — это сила'); T(pres.slides[0], 1, 'Подзаголовок презентации');
      const s = makeSlide(pres, 'content'); T(s, 0, 'Основные тезисы'); T(s, 1, '<ul><li>Первый тезис</li><li>Второй тезис</li><li>Третий тезис</li></ul>'); pres.slides.push(s);
      const s2 = makeSlide(pres, 'section'); T(s2, 0, 'Следующий раздел'); pres.slides.push(s2);
    }
  }
  async function newDoc(tpl) {
    if (tpl) buildTemplate(tpl); else pres = newPres();
    cur = 0; selIds.clear(); undoSt.length = 0; redoSt.length = 0; lastState = snapshot();
    shell.setTitle(tpl ? TEMPLATES.point.find((t) => t.id === tpl)?.name || 'Презентация' : 'Презентация');
    renderAll();
  }
  async function load(file) {
    const ext = extOf(file.name);
    if (ext === 'rpt') { const d = JSON.parse(new TextDecoder().decode(file.bytes)); pres = { w: d.w || 1280, h: d.h || 720, theme: d.theme || { ...THEMES[0] }, slides: d.slides || [] }; }
    else if (ext === 'pptx') pres = await importPptx(file.bytes);
    if (!pres.slides.length) pres.slides.push(makeSlide(pres, 'title'));
    cur = 0; selIds.clear(); undoSt.length = 0; redoSt.length = 0; lastState = snapshot(); renderAll();
  }
  const serialize = () => { exitEditSilently(); return JSON.stringify({ format: 'rabbit-presentation', version: 1, ...pres }); };
  function exitEditSilently() { if (!editingId) return; const el = elById(editingId); const tx = slideEl?.querySelector(`.el[data-id="${editingId}"] .tx`); if (el && tx) el.text.html = cleanHtml(tx.innerHTML); }

  /* =================== shell =================== */
  const shell = createShell({
    app: 'point', ribbon, defaultTitle: 'Презентация', serialize, load, newDoc, templates: TEMPLATES.point,
    isEmpty: () => pres.slides.length === 1 && pres.slides[0].els.every((e) => !e.text?.html && e.ph),
    undo, redo, print: () => printSlides(false),
    stats: () => [['Слайдов', pres.slides.length], ['Тема', pres.theme.name], ['Формат', pres.h === 720 ? '16:9' : '4:3']],
    exporters: [
      { label: 'Презентация PowerPoint (.pptx)', desc: 'Совместима с PowerPoint, Keynote, LibreOffice Impress', icon: 'slides', run: async () => shell.export(shell.title + '.pptx', await exportPptx(pres), 'application/vnd.openxmlformats-officedocument.presentationml.presentation') },
      { label: 'PDF', desc: 'Каждый слайд — отдельная страница', icon: 'print', run: () => printSlides(true) },
      { label: 'Веб-презентация (.html)', desc: 'Открывается в любом браузере, с переходами и анимацией', icon: 'globe', run: exportHtml },
      { label: 'Текущий слайд (.png)', desc: 'Картинка 1920×1080', icon: 'image', run: () => exportPng(false) },
      { label: 'Все слайды (.png в архиве .zip)', desc: 'Каждый слайд — отдельная картинка', icon: 'image', run: () => exportPng(true) },
    ],
  });
  shell.workspace.append(slidesPanel, main);
  const sbSlide = h('span'); const sbTheme = h('span', { class: 'hide-m' });
  shell.statusbar.append(sbSlide, sbTheme, h('span', { class: 'grow' }), h('button', { class: 'sb-btn', html: icon('notes') + '<span class="hide-m">Заметки</span>', onclick: () => { notesPane.classList.toggle('hidden'); setTimeout(fit, 0); } }), h('button', { class: 'sb-btn', html: icon('grid'), 'data-tip': 'Сортировщик', onclick: () => { viewMode = viewMode === 'sorter' ? 'normal' : 'sorter'; if (viewMode === 'sorter') renderSorter(); else { sorterEl?.remove(); sorterEl = null; } } }), h('button', { class: 'sb-btn', html: icon('play'), 'data-tip': 'Слайд-шоу|F5', onclick: () => startShow(cur) }));
  const origRenderThumbs = renderThumbs;
  const updStatus = () => { sbSlide.textContent = `Слайд ${cur + 1} из ${pres.slides.length}`; sbTheme.textContent = `Тема: ${pres.theme.name}`; };
  setInterval(updStatus, 300);
  new ResizeObserver(() => fit()).observe(stageWrap);
  stageWrap.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); zoomMul = clamp(zoomMul * (e.deltaY < 0 ? 1.1 : 0.9), 0.3, 3); fit(); } else if (!editingId && zoomMul <= 1 && Math.abs(e.deltaY) > 30) { goto(cur + Math.sign(e.deltaY)); } }, { passive: false });
  const tabObserver = setInterval(() => { if (shell.ribbon.active !== lastTab) { lastTab = shell.ribbon.active; renderStage(); } }, 300);
  let lastTab = null;
  window.addEventListener('rabbit-settings', (e) => { if (e.detail.k === 'theme') applyTheme(); });
  channel?.addEventListener('message', (m) => { if (m.data.type === 'cmd' && m.data.cmd === 'hello' && !show) channel.postMessage({ type: 'state', pres, index: cur, next: cur + 1 < pres.slides.length ? cur + 1 : null, step: 0 }); });
  shell.boot().then(() => { lastState = snapshot(); updUndo(); });
}

/* =====================================================================
   PRESENTER VIEW (separate window)
   ===================================================================== */
function presenterMode() {
  applyTheme('dark');
  document.title = 'Режим докладчика — Rabbit Point';
  const curBox = h('div', { class: 'cur' }), nextBox = h('div', { class: 'nxt' }), notesBox = h('div', { class: 'notes' }, 'Ожидание показа…');
  const timer = h('span', { class: 'timer' }, '00:00'); const clock = h('span', { class: 'muted' });
  const ctl = h('div', { class: 'ctl' }, h('button', { class: 'btn', html: icon('chevl') + 'Назад', onclick: () => channel.postMessage({ type: 'cmd', cmd: 'prev' }) }), h('button', { class: 'btn primary', html: 'Далее' + icon('chevr'), onclick: () => channel.postMessage({ type: 'cmd', cmd: 'next' }) }), timer, h('button', { class: 'btn sm', onclick: () => { t0 = Date.now(); } }, 'Сброс'), h('span', { class: 'grow' }), clock, h('button', { class: 'btn', onclick: () => channel.postMessage({ type: 'cmd', cmd: 'end' }) }, 'Завершить показ'));
  const root = h('div', { class: 'presenter' }, curBox, nextBox, notesBox, ctl);
  document.body.append(root);
  let t0 = Date.now();
  setInterval(() => { const s = Math.floor((Date.now() - t0) / 1000); timer.textContent = `${String(Math.floor(s / 60)).padStart(2, '0')}:${String(s % 60).padStart(2, '0')}`; clock.textContent = new Date().toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' }); }, 500);
  const put = (box, pres, i) => { box.innerHTML = ''; if (i == null) { box.append(h('div', { style: { color: '#888', padding: '20px' } }, 'Конец презентации')); return; } const s = renderSlide(pres, pres.slides[i]); s.style.position = 'absolute'; s.style.transformOrigin = '0 0'; const k = Math.min(box.clientWidth / pres.w, box.clientHeight / pres.h); s.style.transform = `scale(${k})`; s.style.left = (box.clientWidth - pres.w * k) / 2 + 'px'; s.style.top = (box.clientHeight - pres.h * k) / 2 + 'px'; box.append(s); };
  channel?.addEventListener('message', (m) => { const d = m.data; if (d.type === 'state') { put(curBox, d.pres, d.index); put(nextBox, d.pres, d.next); notesBox.textContent = d.pres.slides[d.index].notes || 'Нет заметок'; } if (d.type === 'end') notesBox.textContent = 'Показ завершён'; });
  document.addEventListener('keydown', (e) => { if (['ArrowRight', ' ', 'PageDown'].includes(e.key)) channel.postMessage({ type: 'cmd', cmd: 'next' }); if (['ArrowLeft', 'PageUp'].includes(e.key)) channel.postMessage({ type: 'cmd', cmd: 'prev' }); });
  channel?.postMessage({ type: 'cmd', cmd: 'hello' });
}
