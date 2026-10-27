// Rabbit Office — small UI toolkit (dialogs, menus, toasts, pickers)
import { icon } from './icons.js';

export const $ = (s, r = document) => r.querySelector(s);
export const $$ = (s, r = document) => [...r.querySelectorAll(s)];
export const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
export const uid = (p = '') => p + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
export const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
export const debounce = (fn, ms = 300) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

export function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'style' && typeof v === 'object') { for (const [sk, sv] of Object.entries(v)) { if (sk.startsWith('--')) el.style.setProperty(sk, sv); else el.style[sk] = sv; } }
    else if (k === 'html') el.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) el.append(k instanceof Node ? k : document.createTextNode(k));
  return el;
}

/* ---------- toasts ---------- */
let toastBox;
export function toast(msg, type = 'info', ms = 2600) {
  if (!toastBox) { toastBox = h('div', { class: 'toasts' }); document.body.append(toastBox); }
  const ic = type === 'error' ? 'warn' : type === 'ok' ? 'check' : 'info';
  const t = h('div', { class: `toast ${type}`, html: icon(ic) + `<span>${esc(msg)}</span>` });
  toastBox.append(t);
  setTimeout(() => { t.style.transition = 'opacity .3s'; t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, ms);
}

/* ---------- dialogs ---------- */
export function dialog({ title = '', body = '', buttons = [{ label: 'OK', value: true, primary: true }], wide = false, onOpen, closable = true, cls = '' } = {}) {
  return new Promise((resolve) => {
    const ov = h('div', { class: 'overlay' });
    const dlg = h('div', { class: `dialog ${wide ? 'wide' : ''} ${cls}`, role: 'dialog' });
    const head = h('div', { class: 'dialog-head' }, h('span', { class: 'grow' }, title));
    if (closable) head.append(h('button', { class: 'btn ghost icon', html: icon('close'), onclick: () => done(null), 'aria-label': 'Закрыть' }));
    const bodyEl = h('div', { class: 'dialog-body' });
    if (typeof body === 'string') bodyEl.innerHTML = body; else if (body) bodyEl.append(body);
    const foot = h('div', { class: 'dialog-foot' });
    const api = { el: dlg, body: bodyEl, close: (v) => done(v) };
    for (const b of buttons) {
      const btn = h('button', { class: `btn ${b.primary ? 'primary' : ''} ${b.danger ? 'danger' : ''}` }, b.label);
      btn.onclick = async () => {
        const v = typeof b.value === 'function' ? await b.value(api) : b.value;
        if (v === undefined && typeof b.value === 'function') return; // handler keeps dialog open
        done(v);
      };
      foot.append(btn);
    }
    dlg.append(head, bodyEl);
    if (buttons.length) dlg.append(foot);
    ov.append(dlg);
    document.body.append(ov);
    const onKey = (e) => {
      if (e.key === 'Escape' && closable) { e.stopPropagation(); done(null); }
      if (e.key === 'Enter' && !e.shiftKey && e.target.tagName !== 'TEXTAREA') {
        const p = buttons.find((b) => b.primary); if (p) { e.preventDefault(); foot.querySelector('.primary')?.click(); }
      }
    };
    ov.addEventListener('keydown', onKey);
    ov.addEventListener('mousedown', (e) => { if (e.target === ov && closable) done(null); });
    function done(v) { ov.remove(); resolve(v); }
    setTimeout(() => { (dlg.querySelector('[autofocus]') || dlg.querySelector('input,select,textarea') || foot.querySelector('.primary'))?.focus(); onOpen?.(api); }, 10);
  });
}

export function alertDlg(title, text) {
  return dialog({ title, body: h('div', {}, text), buttons: [{ label: 'OK', value: true, primary: true }] });
}

export function confirmDlg(title, text, okLabel = 'OK', danger = false) {
  return dialog({ title, body: h('div', {}, text), buttons: [{ label: 'Отмена', value: false }, { label: okLabel, value: true, primary: !danger, danger }] }).then((v) => !!v);
}

export function promptDlg(title, label, value = '', { type = 'text', placeholder = '' } = {}) {
  const inp = h('input', { class: 'input', type, value, placeholder, autofocus: true });
  const body = h('div', {}, h('div', { class: 'field' }, label ? h('label', {}, label) : null, inp));
  setTimeout(() => inp.select(), 20);
  return dialog({ title, body, buttons: [{ label: 'Отмена', value: null }, { label: 'OK', primary: true, value: () => inp.value }] });
}

/* form dialog: fields [{key,label,type,value,options}] */
export function formDlg(title, fields, okLabel = 'OK') {
  const els = {};
  const body = h('div');
  for (const f of fields) {
    let el;
    if (f.type === 'select') {
      el = h('select', { class: 'input' }, ...f.options.map((o) => h('option', { value: o.value ?? o, selected: (o.value ?? o) == f.value }, o.label ?? o)));
    } else if (f.type === 'checkbox') {
      el = h('input', { type: 'checkbox', checked: !!f.value });
      body.append(h('label', { class: 'check', style: { marginBottom: '12px' } }, el, f.label));
      els[f.key] = el; continue;
    } else if (f.type === 'textarea') {
      el = h('textarea', { class: 'input', rows: f.rows || 4, placeholder: f.placeholder || '' }); el.value = f.value ?? '';
    } else {
      el = h('input', { class: 'input', type: f.type || 'text', value: f.value ?? '', placeholder: f.placeholder || '', min: f.min, max: f.max, step: f.step });
    }
    els[f.key] = el;
    body.append(h('div', { class: 'field' }, h('label', {}, f.label), el, f.hint ? h('div', { class: 'small muted' }, f.hint) : null));
  }
  return dialog({
    title, body, buttons: [{ label: 'Отмена', value: null }, {
      label: okLabel, primary: true, value: () => {
        const out = {};
        for (const f of fields) { const el = els[f.key]; out[f.key] = f.type === 'checkbox' ? el.checked : f.type === 'number' ? Number(el.value) : el.value; }
        return out;
      },
    }],
  });
}

/* ---------- menus ---------- */
let openMenu = null;
export function closeMenus() { if (openMenu) { openMenu.remove(); openMenu = null; } }
document.addEventListener('mousedown', (e) => { if (openMenu && !openMenu.contains(e.target)) closeMenus(); }, true);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && openMenu) { closeMenus(); } }, true);
window.addEventListener('blur', closeMenus);

function place(el, anchor) {
  document.body.append(el);
  const r = anchor instanceof Element ? anchor.getBoundingClientRect() : { left: anchor.x, right: anchor.x, top: anchor.y, bottom: anchor.y };
  const w = el.offsetWidth, hh = el.offsetHeight;
  let x = r.left, y = r.bottom + 2;
  if (x + w > innerWidth - 6) x = Math.max(6, (anchor instanceof Element ? r.right : r.left) - w);
  if (y + hh > innerHeight - 6) y = Math.max(6, r.top - hh - 2);
  el.style.left = x + 'px'; el.style.top = y + 'px';
}

/* items: {label, icon, hint, action, checked, disabled, danger, sub:[...]} | '-' | {title} */
export function menu(anchor, items) {
  closeMenus();
  const m = h('div', { class: 'menu', role: 'menu' });
  for (const it of items) {
    if (!it) continue;
    if (it === '-') { m.append(h('div', { class: 'menu-sep' })); continue; }
    if (it.title) { m.append(h('div', { class: 'menu-title' }, it.title)); continue; }
    if (it.el) { m.append(it.el); continue; }
    const b = h('button', { class: `menu-item ${it.checked ? 'checked' : ''} ${it.danger ? 'danger' : ''}`, disabled: it.disabled, role: 'menuitem' });
    b.innerHTML = (it.icon ? icon(it.checked ? 'check' : it.icon) : it.checked ? icon('check') : '<span style="width:16px"></span>') +
      `<span class="lbl" ${it.style ? `style="${esc(it.style)}"` : ''}>${esc(it.label)}</span>` + (it.hint ? `<span class="hint">${esc(it.hint)}</span>` : '') + (it.sub ? icon('chevr') : '');
    b.onclick = (e) => {
      if (it.sub) { e.stopPropagation(); const sm = menu(b.getBoundingClientRect().right > innerWidth - 200 ? b : { x: b.getBoundingClientRect().right, y: b.getBoundingClientRect().top - 4 }, it.sub); return sm; }
      closeMenus(); it.action?.();
    };
    m.append(b);
  }
  place(m, anchor);
  openMenu = m;
  return m;
}

export function popover(anchor, content, cls = '') {
  closeMenus();
  const m = h('div', { class: `menu popover ${cls}` });
  m.append(content);
  place(m, anchor);
  openMenu = m;
  return m;
}

/* ---------- color picker ---------- */
export const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
  '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
  '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
  '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
  '#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
  '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130',
];
export function colorPicker(anchor, onPick, { allowNone = false, noneLabel = 'Нет цвета', current } = {}) {
  const wrap = h('div');
  const grid = h('div', { class: 'palette' });
  for (const c of PALETTE) grid.append(h('button', { class: 'sw', style: { background: c }, title: c, onclick: () => { closeMenus(); onPick(c); } }));
  wrap.append(grid);
  const custom = h('input', { type: 'color', value: current && current.startsWith('#') && current.length === 7 ? current : '#3b6fe0', style: { width: '36px', height: '26px', border: 0, padding: 0, background: 'none' } });
  custom.onchange = () => { closeMenus(); onPick(custom.value); };
  const row = h('div', { class: 'palette-row' }, h('span', { class: 'small muted' }, 'Другой:'), custom);
  if (allowNone) row.append(h('button', { class: 'btn sm ghost', onclick: () => { closeMenus(); onPick(null); } }, noneLabel));
  wrap.append(row);
  return popover(anchor, wrap);
}

/* ---------- tooltips ---------- */
let tipEl, tipT;
document.addEventListener('pointerover', (e) => {
  const t = e.target.closest?.('[data-tip]');
  clearTimeout(tipT); tipEl?.remove(); tipEl = null;
  if (!t || e.pointerType === 'touch') return;
  tipT = setTimeout(() => {
    const [title, key] = t.dataset.tip.split('|');
    tipEl = h('div', { class: 'tooltip', html: `${esc(title)}${key ? `<span class="kbd">${esc(key)}</span>` : ''}` });
    document.body.append(tipEl);
    const r = t.getBoundingClientRect();
    let x = r.left + r.width / 2 - tipEl.offsetWidth / 2;
    x = clamp(x, 6, innerWidth - tipEl.offsetWidth - 6);
    let y = r.bottom + 6; if (y + tipEl.offsetHeight > innerHeight) y = r.top - tipEl.offsetHeight - 6;
    tipEl.style.left = x + 'px'; tipEl.style.top = y + 'px';
  }, 450);
});
document.addEventListener('pointerdown', () => { clearTimeout(tipT); tipEl?.remove(); tipEl = null; }, true);

/* ---------- misc helpers ---------- */
export function readFileAs(file, as = 'text') {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result); r.onerror = rej;
    if (as === 'dataurl') r.readAsDataURL(file); else if (as === 'buffer') r.readAsArrayBuffer(file); else r.readAsText(file);
  });
}
export function pickFile(accept = '*', multiple = false) {
  return new Promise((res) => {
    const inp = h('input', { type: 'file', accept, multiple, style: { display: 'none' } });
    inp.onchange = () => { res(multiple ? [...inp.files] : inp.files[0] || null); inp.remove(); };
    document.body.append(inp); inp.click();
  });
}
export async function pickImage() {
  const f = await pickFile('image/*');
  if (!f) return null;
  return readFileAs(f, 'dataurl');
}
export function fmtDate(ts) {
  const d = new Date(ts); const now = new Date();
  if (d.toDateString() === now.toDateString()) return d.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' });
  return d.toLocaleDateString('ru-RU', { day: 'numeric', month: 'short', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric' });
}
export function fmtSize(b) {
  if (b < 1024) return b + ' Б'; if (b < 1048576) return (b / 1024).toFixed(1) + ' КБ';
  if (b < 1073741824) return (b / 1048576).toFixed(1) + ' МБ'; return (b / 1073741824).toFixed(2) + ' ГБ';
}
