// Rabbit Office — application shell: title bar, ribbon, backstage (File), save/open flows, themes.
import { icon } from './icons.js';
import { h, $, $$, esc, toast, dialog, menu, colorPicker, confirmDlg, promptDlg, debounce, fmtDate, fmtSize, uid } from './ui.js';
import { settings, recent, library, profile, kv } from './store.js';
import { accounts, files as cloudFiles, login, logout, ApiError, changePassword } from './cloud.js';
import { APPS, native, platform, openLocalFile, readFromLoc, writeToLoc, locLabel, locFromRecent, extOf, baseName, exportFile, takeHandoff, goHome, openApp, setTitleBarColors, appForFile, handoff, isAndroid } from './platform.js';

export const THEMES = [
  { id: 'standard', name: 'Стандартная', bar: 'var(--accent)', bg: '#f3f3f3', rb: '#ffffff' },
  { id: 'dark', name: 'Стандартная тёмная', bar: '#1f1f20', bg: '#1b1b1c', rb: '#2b2b2c' },
  { id: 'glass', name: 'Liquid Glass', bar: 'transparent', bg: 'linear-gradient(135deg,#dfe6f3,#ece4f3)', rb: 'rgba(255,255,255,.65)' },
  { id: 'glass-dark', name: 'Liquid Glass тёмная', bar: 'transparent', bg: 'linear-gradient(135deg,#1b2233,#21182c)', rb: 'rgba(60,64,80,.6)' },
];

export function applyTheme(t = settings.get('theme')) {
  document.documentElement.dataset.theme = t;
  requestAnimationFrame(() => {
    const cs = getComputedStyle(document.documentElement);
    const glass = t.startsWith('glass');
    const bg = glass ? '#00000000' : cs.getPropertyValue('--titlebar-bg').trim();
    const fg = cs.getPropertyValue('--titlebar-fg').trim();
    setTitleBarColors(toHex(bg), toHex(fg));
    const meta = document.querySelector('meta[name=theme-color]');
    if (meta) meta.content = glass ? (t === 'glass' ? '#e9edf3' : '#12141a') : toHex(bg);
  });
}
function toHex(c) {
  if (!c || c.startsWith('#')) return c || '#000000';
  const cv = document.createElement('canvas').getContext('2d'); cv.fillStyle = c; return cv.fillStyle;
}
window.addEventListener('rabbit-settings', (e) => { if (e.detail.k === 'theme') applyTheme(); });
window.addEventListener('storage', (e) => { if (e.key === 'rabbit.settings') applyTheme(); });

export function themeMenu(anchor) {
  const cur = settings.get('theme');
  menu(anchor, [{ title: 'Оформление' }, ...THEMES.map((t) => ({ label: t.name, icon: t.id.includes('dark') ? 'moon' : 'sun', checked: cur === t.id, action: () => settings.set('theme', t.id) }))]);
}

export function themeCards() {
  const wrap = h('div', { class: 'theme-cards' });
  const render = () => {
    wrap.innerHTML = '';
    const cur = settings.get('theme');
    for (const t of THEMES) {
      const glass = t.id.startsWith('glass');
      const c = h('button', { class: `theme-card ${cur === t.id ? 'on' : ''}`, onclick: () => { settings.set('theme', t.id); render(); } });
      c.innerHTML = `<div class="prev" style="background:${t.bg}">
        <div class="bar" style="background:${t.bar}"></div>
        <div class="rb" style="background:${t.rb};${glass ? 'backdrop-filter:blur(8px);box-shadow:inset 0 1px 0 rgba(255,255,255,.5),0 2px 8px rgba(0,0,0,.08);border-radius:7px' : ''}"></div>
        <div class="pg" style="${glass ? 'box-shadow:0 6px 18px rgba(0,0,0,.12)' : ''}"></div></div>
        <div class="nm">${esc(t.name)}</div>`;
      wrap.append(c);
    }
  };
  render();
  return wrap;
}

/* =================== Ribbon =================== */
class Ribbon {
  constructor(shell, tabs) {
    this.shell = shell; this.tabs = tabs; this.items = new Map(); this.tabEls = new Map(); this.panels = new Map();
    this.el = h('nav', { class: 'ribbon no-print' });
    this.tabsEl = h('div', { class: 'ribbon-tabs', role: 'tablist' });
    this.panelEl = h('div', { class: 'ribbon-panel' });
    const fileTab = h('button', { class: 'ribbon-tab file', onclick: () => shell.openBackstage() }, 'Файл');
    this.tabsEl.append(fileTab);
    for (const t of tabs) {
      const b = h('button', { class: `ribbon-tab ${t.contextual ? 'contextual hidden' : ''}`, role: 'tab', onclick: () => this.activate(t.id, true) }, t.label);
      this.tabEls.set(t.id, b); this.tabsEl.append(b);
      this.panels.set(t.id, this.buildPanel(t));
    }
    this.tabsEl.append(h('div', { class: 'spacer' }));
    this.collapseBtn = h('button', { class: 'btn-r ribbon-collapse', 'data-tip': 'Свернуть/развернуть ленту|Ctrl+F1', html: icon('chevu'), onclick: () => this.toggleCollapsed() });
    this.tabsEl.append(this.collapseBtn);
    this.el.append(this.tabsEl, this.panelEl);
    if (settings.get('ribbonCollapsed')) this.el.classList.add('collapsed');
    this.updateCollapseIcon();
    this.activate(tabs.find((t) => !t.contextual)?.id);
    document.addEventListener('mousedown', (e) => { if (this.el.classList.contains('peek') && !this.el.contains(e.target) && !e.target.closest('.menu')) this.el.classList.remove('peek'); });
  }
  toggleCollapsed() { this.el.classList.toggle('collapsed'); this.el.classList.remove('peek'); settings.set('ribbonCollapsed', this.el.classList.contains('collapsed')); this.updateCollapseIcon(); window.dispatchEvent(new Event('resize')); }
  updateCollapseIcon() { this.collapseBtn.innerHTML = icon(this.el.classList.contains('collapsed') ? 'chevd' : 'chevu'); }
  activate(id, user) {
    if (!id) return;
    if (user && this.el.classList.contains('collapsed')) this.el.classList.toggle('peek', !(this.active === id && this.el.classList.contains('peek')));
    this.active = id;
    for (const [tid, b] of this.tabEls) b.classList.toggle('active', tid === id);
    this.panelEl.innerHTML = ''; this.panelEl.append(this.panels.get(id));
  }
  showTab(id, visible, activate = false) {
    const b = this.tabEls.get(id); if (!b) return;
    const was = !b.classList.contains('hidden');
    b.classList.toggle('hidden', !visible);
    if (visible && activate && !was) this.activate(id);
    if (!visible && this.active === id) this.activate(this.tabs.find((t) => !t.contextual)?.id);
  }
  buildPanel(tab) {
    const frag = h('div', { style: { display: 'contents' } });
    for (const g of tab.groups) {
      const body = h('div', { class: 'r-group-body' });
      for (const it of g.items) body.append(this.buildItem(it));
      frag.append(h('div', { class: 'r-group' }, body, h('div', { class: 'r-group-label' }, g.label || '')));
    }
    return frag;
  }
  buildItem(it) {
    if (it.type === 'row' || it.type === 'col') { const c = h('div', { class: it.type === 'row' ? 'r-row' : 'r-col' }); for (const x of it.items) c.append(this.buildItem(x)); return c; }
    if (it.type === 'custom') { if (it.id) this.items.set(it.id, { el: it.el, spec: it }); return it.el; }
    if (it.type === 'select') {
      let el;
      if (it.editable) {
        el = h('input', { class: 'r-select', style: { width: (it.width || 60) + 'px' }, value: it.value ?? '', 'data-tip': it.tip || '' });
        const dl = h('datalist', { id: 'dl-' + it.id }, ...it.options.map((o) => h('option', { value: o.value ?? o })));
        el.setAttribute('list', 'dl-' + it.id);
        el.addEventListener('change', () => it.action(el.value));
        el.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); it.action(el.value); this.shell.focusEditor?.(); } });
        const wrap = h('span', { style: { display: 'contents' } }, el, dl);
        this.items.set(it.id, { el, spec: it });
        return wrap;
      }
      el = h('select', { class: 'r-select', style: { width: (it.width || 120) + 'px' }, 'data-tip': it.tip || '' },
        ...it.options.map((o) => h('option', { value: o.value ?? o, style: o.style ? undefined : undefined }, o.label ?? o)));
      if (it.value != null) el.value = it.value;
      el.addEventListener('change', () => { it.action(el.value); this.shell.focusEditor?.(); });
      this.items.set(it.id, { el, spec: it });
      return el;
    }
    const big = it.type === 'big';
    const b = h('button', { class: `btn-r ${big ? 'big' : ''} ${it.type === 'color' ? 'split' : ''}`, 'data-tip': it.tip ? `${it.tip}${it.key ? '|' + it.key : ''}` : (it.label && !big && !it.showLabel ? '' : '') });
    if (!it.tip && it.label && !it.showLabel && !big) b.dataset.tip = it.label + (it.key ? '|' + it.key : '');
    b.innerHTML = icon(it.icon || 'file') + (big || it.showLabel ? `<span>${esc(it.label || '')}</span>` : '') + (it.menu || it.type === 'color' ? icon('chevd', 'caret') : '');
    b.addEventListener('mousedown', (e) => { if (!it.allowFocus) e.preventDefault(); }); // keep editor selection
    if (it.type === 'color') {
      const bar = h('span', { class: 'color-bar', style: { background: it.value || '#000' } }); b.append(bar);
      b.onclick = (e) => {
        const rect = b.getBoundingClientRect();
        if (e.clientX > rect.right - 14) colorPicker(b, (c) => { it.value = c; bar.style.background = c || 'transparent'; it.action(c); }, { allowNone: it.allowNone, current: it.value });
        else it.action(it.value);
      };
      this.items.set(it.id, { el: b, spec: it, bar });
    } else {
      b.onclick = (e) => { if (it.menu) { const items = typeof it.menu === 'function' ? it.menu(b) : it.menu; if (items) menu(b, items); } else it.action?.(e, b); };
      if (it.id) this.items.set(it.id, { el: b, spec: it });
    }
    return b;
  }
  set(id, { on, disabled, value } = {}) {
    const r = this.items.get(id); if (!r) return;
    if (on !== undefined) r.el.classList.toggle('on', !!on);
    if (disabled !== undefined) r.el.disabled = !!disabled;
    if (value !== undefined && (r.el.tagName === 'SELECT' || r.el.tagName === 'INPUT') && document.activeElement !== r.el) {
      if (r.el.tagName === 'SELECT' && value !== '' && ![...r.el.options].some((o) => o.value === String(value))) r.el.append(h('option', { value }, value));
      r.el.value = value;
    }
  }
}

/* =================== Shell =================== */
export function createShell(opts) {
  const app = APPS[opts.app];
  document.documentElement.dataset.app = opts.app;
  applyTheme();
  document.title = app.name;
  const shell = {
    app, opts, loc: null, dirty: false, title: opts.defaultTitle || 'Без названия',
    focusEditor: opts.focusEditor,
  };

  /* ---- title bar ---- */
  const tb = h('header', { class: 'titlebar no-print' });
  const logo = h('span', { class: 'logo single', style: { width: '18px', height: '18px' } });
  const homeBtn = h('button', { class: 'tb-btn', 'data-tip': 'Rabbit Office — главная', onclick: () => goHome() }, logo);
  const saveBtn = h('button', { class: 'tb-btn', 'data-tip': 'Сохранить|Ctrl+S', html: icon('save'), onclick: () => shell.save() });
  const undoBtn = h('button', { class: 'tb-btn', 'data-tip': 'Отменить|Ctrl+Z', html: icon('undo'), onclick: () => opts.undo?.() });
  const redoBtn = h('button', { class: 'tb-btn', 'data-tip': 'Повторить|Ctrl+Y', html: icon('redo'), onclick: () => opts.redo?.() });
  [undoBtn, redoBtn].forEach((b) => b.addEventListener('mousedown', (e) => e.preventDefault()));
  const autosaveBtn = h('button', { class: 'tb-btn hide-m', 'data-tip': 'Автосохранение', onclick: () => { settings.set('autosave', !settings.get('autosave')); updAuto(); if (settings.get('autosave') && shell.dirty && shell.loc) shell.save(); } });
  const updAuto = () => { autosaveBtn.innerHTML = `<span style="font-size:12px">Автосохр.</span><span style="display:inline-block;width:26px;height:14px;border-radius:8px;background:${settings.get('autosave') ? 'currentColor' : 'transparent'};border:1.5px solid currentColor;position:relative;opacity:.9"><i style="position:absolute;top:1.5px;${settings.get('autosave') ? 'right:2px' : 'left:2px'};width:8px;height:8px;border-radius:50%;background:${settings.get('autosave') ? 'var(--titlebar-bg)' : 'currentColor'}"></i></span>`; };
  updAuto();
  const titleInput = h('input', { class: 'doc-title', value: shell.title, spellcheck: 'false', 'aria-label': 'Название документа' });
  titleInput.addEventListener('change', () => { shell.title = titleInput.value.trim() || 'Без названия'; titleInput.value = shell.title; shell.markDirty(); updateDocTitle(); });
  titleInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { titleInput.blur(); opts.focusEditor?.(); } });
  const saveState = h('span', { class: 'save-state' });
  const center = h('div', { class: 'center' }, titleInput, saveState);
  const extra = h('div', { class: 'no-drag', style: { display: 'flex', gap: '4px', alignItems: 'center' } });
  const searchBtn = opts.onSearch ? h('button', { class: 'tb-btn', 'data-tip': 'Найти|Ctrl+F', html: icon('search'), onclick: () => opts.onSearch() }) : null;
  const themeBtn = h('button', { class: 'tb-btn', 'data-tip': 'Оформление', html: icon('palette'), onclick: (e) => themeMenu(e.currentTarget) });
  const acctBtn = h('button', { class: 'tb-btn', 'data-tip': 'Учётная запись и сервер' });
  const updAcct = () => {
    const p = profile(); const a = accounts.active();
    acctBtn.innerHTML = `<span class="avatar" style="background:${p.color}">${esc((a?.displayName || p.name).slice(0, 1).toUpperCase())}</span>`;
  };
  updAcct(); window.addEventListener('rabbit-accounts', updAcct); window.addEventListener('rabbit-settings', updAcct);
  acctBtn.onclick = () => accountMenu(acctBtn, shell);
  tb.append(homeBtn, h('span', { class: 'app-name' }, app.name), saveBtn, undoBtn, redoBtn, autosaveBtn, center, h('div', { class: 'grow' }), extra);
  if (searchBtn) tb.append(searchBtn);
  tb.append(themeBtn, acctBtn);
  shell.titlebarExtra = extra;
  shell.undoBtn = undoBtn; shell.redoBtn = redoBtn;

  /* ---- ribbon ---- */
  shell.ribbon = new Ribbon(shell, opts.ribbon || []);
  shell.workspace = h('main', { class: 'workspace' });
  shell.statusbar = h('footer', { class: 'statusbar no-print' });
  document.body.className = 'app';
  document.body.append(tb, shell.ribbon.el, shell.workspace, shell.statusbar);

  function updateDocTitle() {
    document.title = `${shell.title}${shell.dirty ? ' •' : ''} — ${app.name}`;
    saveState.textContent = shell.saving ? 'Сохранение…' : shell.dirty ? (shell.loc ? '• Изменено' : '• Не сохранено') : shell.loc ? 'Сохранено' : '';
    saveState.title = locLabel(shell.loc);
    native?.setDirty?.(shell.dirty);
  }
  shell.setTitle = (t) => { shell.title = t; titleInput.value = t; updateDocTitle(); };
  const autosave = debounce(() => { if (settings.get('autosave') && shell.loc && shell.dirty) shell.save({ silent: true }); }, 2500);
  /* crash recovery: unsaved work is mirrored to IndexedDB and offered back on next start */
  const draftKey = `draft:${app.id}:${uid()}`;
  const saveDraft = debounce(async () => {
    if (!shell.dirty) return;
    try { await kv.set(draftKey, { app: app.id, title: shell.title, data: await opts.serialize(), loc: shell.loc, time: Date.now() }); drafts.add(draftKey); } catch { /* storage full or serializer busy */ }
  }, 4000);
  const dropDraft = () => { kv.del(draftKey).catch(() => {}); drafts.remove(draftKey); };
  setInterval(() => drafts.alive(draftKey), 5000); drafts.alive(draftKey);
  window.addEventListener('pagehide', () => { if (!shell.dirty) dropDraft(); });
  shell.markDirty = () => { shell.dirty = true; updateDocTitle(); autosave(); saveDraft(); opts.onDirty?.(); };
  shell.markClean = () => { shell.dirty = false; updateDocTitle(); dropDraft(); };

  /* ---- save / open ---- */
  async function serialize() { return opts.serialize(); }
  shell.save = async ({ silent = false } = {}) => {
    if (!shell.loc) return shell.saveAs();
    if (extOf(shell.loc.path || '') && extOf(shell.loc.path) !== app.ext) return shell.saveAs(); // imported foreign format
    try {
      shell.saving = true; updateDocTitle();
      await writeToLoc(shell.loc, shell.fileName(), await serialize(), app.id);
      shell.saving = false; shell.markClean();
      if (!silent) toast('Сохранено', 'ok', 1400);
    } catch (e) { shell.saving = false; updateDocTitle(); toast('Ошибка сохранения: ' + e.message, 'error', 5000); }
  };
  shell.fileName = () => baseName(shell.title) + '.' + app.ext;
  shell.saveAs = async () => {
    const choice = await saveLocationDialog(shell);
    if (!choice) return;
    shell.loc = choice.loc; if (choice.name) shell.setTitle(baseName(choice.name));
    await shell.save();
  };
  shell.loadFile = async (file) => {
    try {
      await opts.load(file);
      shell.setTitle(baseName(file.name));
      const own = extOf(file.name) === app.ext;
      shell.loc = own ? file.loc : null;
      if (file.loc) recent.add({ app: app.id, name: file.name, ...(file.loc.source === 'native' ? { source: 'native', path: file.loc.path } : file.loc.source === 'library' ? { source: 'library', libId: file.loc.libId } : { source: 'cloud', account: file.loc.account, cloudPath: file.loc.path }) });
      shell.markClean();
      if (!own) toast(`Импортировано из ${extOf(file.name).toUpperCase()}. Сохранение создаст файл .${app.ext}`, 'info', 4000);
    } catch (e) { console.error(e); toast('Не удалось открыть файл: ' + e.message, 'error', 6000); }
  };
  shell.open = async () => {
    const f = await openLocalFile(app.open); if (!f) return;
    const target = appForFile(f.name);
    if (target && target.id !== app.id) { await handoff(target.id, f); return; }
    if (shell.dirty || !opts.isEmpty?.()) return openInNewWindowOrReplace(f);
    shell.loadFile(f);
  };
  async function openInNewWindowOrReplace(f) {
    if (native) { await handoff(app.id, f); return; }
    if (shell.dirty && !(await confirmDlg('Открыть файл', 'Несохранённые изменения будут потеряны. Продолжить?', 'Открыть'))) return;
    shell.loadFile(f);
  }
  shell.newDoc = async (tpl) => {
    if (native) { openApp(app.id, tpl ? { template: tpl } : {}); return; }
    if (shell.dirty && !(await confirmDlg('Новый документ', 'Несохранённые изменения будут потеряны. Продолжить?', 'Создать'))) return;
    location.href = app.page + (tpl ? '?template=' + tpl : '');
  };
  shell.export = async (name, data, mime) => { const p = await exportFile(name, data, mime); if (p) toast('Экспортировано: ' + (typeof p === 'string' ? p.split(/[\\/]/).pop() : name), 'ok'); };

  /* ---- backstage ---- */
  shell.openBackstage = (page = 'home') => openBackstage(shell, page);

  /* ---- keyboard ---- */
  document.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 's') { e.preventDefault(); e.shiftKey ? shell.saveAs() : shell.save(); }
    else if (e.key === 'F12') { e.preventDefault(); shell.saveAs(); }
    else if (mod && e.key.toLowerCase() === 'o') { e.preventDefault(); shell.openBackstage('open'); }
    else if (mod && e.key.toLowerCase() === 'n' && !e.shiftKey) { e.preventDefault(); shell.newDoc(); }
    else if (mod && e.key === 'F1') { e.preventDefault(); shell.ribbon.toggleCollapsed(); }
    else if (mod && e.key.toLowerCase() === 'p' && opts.print) { e.preventDefault(); opts.print(); }
  });
  window.addEventListener('beforeunload', (e) => { if (shell.dirty && !native) { e.preventDefault(); e.returnValue = ''; } });
  native?.onBeforeClose?.(async () => {
    if (!shell.dirty) return true;
    const r = await dialog({ title: 'Сохранить изменения?', body: `Сохранить изменения в «${esc(shell.title)}»?`, buttons: [{ label: 'Отмена', value: 'cancel' }, { label: 'Не сохранять', value: 'discard' }, { label: 'Сохранить', value: 'save', primary: true }] });
    if (r === 'save') { await shell.save(); return !shell.dirty; }
    if (r === 'discard') dropDraft();
    return r === 'discard';
  });

  /* ---- initial document from URL ---- */
  shell.boot = async () => {
    const q = new URLSearchParams(location.search);
    try {
      if (q.get('handoff')) { const f = await takeHandoff(q.get('handoff')); if (f) { await shell.loadFile(f); return; } }
      if (q.get('recent')) { const e = recent.list()[Number(q.get('recent'))]; if (e) { await shell.loadFile(await readFromLoc(locFromRecent(e), e.name)); return; } }
      if (q.get('path') && native) { await shell.loadFile(await readFromLoc({ source: 'native', path: q.get('path') })); return; }
      if (q.get('template')) { await opts.newDoc?.(q.get('template')); shell.markClean(); return; }
    } catch (e) { toast('Не удалось открыть: ' + e.message, 'error', 6000); }
    await opts.newDoc?.(null); shell.markClean();
    if (!q.get('board')) offerRecovery();
  };
  async function offerRecovery() {
    const list = (await Promise.all(drafts.orphans(app.id).map(async (k) => ({ k, d: await kv.get(k).catch(() => null) })))).filter((x) => x.d);
    if (!list.length) return;
    const body = h('div', { class: 'list', style: { width: 'min(460px,100%)' } }, h('p', { class: 'muted', style: { marginTop: 0 } }, 'Эти документы не были сохранены перед закрытием программы. Восстановить?'));
    let api;
    for (const { k, d } of list.sort((a, b) => b.d.time - a.d.time)) {
      const row = h('div', { class: 'list-item' }, h('span', { class: 'file-badge', style: { background: app.hex } }, app.letter), h('div', { class: 'grow' }, h('div', { class: 'name' }, d.title), h('div', { class: 'meta' }, fmtDate(d.time))));
      row.append(h('button', { class: 'btn sm primary', onclick: async () => { api.close(); await shell.loadFile({ name: baseName(d.title) + '.' + app.ext, bytes: new TextEncoder().encode(typeof d.data === 'string' ? d.data : ''), loc: d.loc }); shell.loc = d.loc; shell.markDirty(); kv.del(k); drafts.remove(k); toast('Документ восстановлен — не забудьте сохранить', 'ok'); } }, 'Восстановить'),
        h('button', { class: 'btn sm ghost', 'data-tip': 'Удалить черновик', html: icon('trash'), onclick: () => { kv.del(k); drafts.remove(k); row.remove(); if (!body.querySelector('.list-item')) api.close(); } }));
      body.append(row);
    }
    dialog({ title: 'Восстановление документов', body, buttons: [{ label: 'Позже', value: null }], onOpen: (a) => { api = a; } });
  }
  updateDocTitle();
  return shell;
}

/* draft index kept in localStorage; a draft is "orphaned" when its window stopped heartbeating */
const drafts = {
  list() { try { return JSON.parse(localStorage.getItem('rabbit.drafts') || '[]'); } catch { return []; } },
  add(k) { const l = drafts.list(); if (!l.includes(k)) { l.push(k); localStorage.setItem('rabbit.drafts', JSON.stringify(l)); } },
  remove(k) { localStorage.setItem('rabbit.drafts', JSON.stringify(drafts.list().filter((x) => x !== k))); localStorage.removeItem('rabbit.alive.' + k); },
  alive(k) { localStorage.setItem('rabbit.alive.' + k, String(Date.now())); },
  orphans(appId) { return drafts.list().filter((k) => k.startsWith(`draft:${appId}:`) && Date.now() - +(localStorage.getItem('rabbit.alive.' + k) || 0) > 15000); },
};

/* =================== Account menu =================== */
export function accountMenu(anchor, shell) {
  const p = profile(); const list = accounts.list(); const act = accounts.active();
  const items = [{ title: p.name }, { label: 'Изменить имя и цвет…', icon: 'user', action: () => profileDialog() }, '-', { title: 'Серверы Rabbit Office' }];
  for (const a of list) items.push({ label: `${a.displayName} · ${a.serverName}`, icon: 'server', checked: act?.id === a.id, action: () => accounts.setActive(a.id) });
  items.push({ label: 'Подключить сервер…', icon: 'plus', action: () => loginDialog() });
  if (act) items.push({ label: 'Сменить пароль…', icon: 'key', action: () => passwordDialog(act) });
  if (act) items.push({ label: 'Выйти с ' + act.serverName, icon: 'logout', action: async () => { await logout(act); toast('Вы вышли с сервера'); } });
  if (act?.role === 'admin') items.push({ label: 'Панель администратора', icon: 'settings', action: () => window.open(act.url + '/admin/', '_blank') });
  items.push('-', { label: 'Оформление…', icon: 'palette', action: () => (shell ? shell.openBackstage('settings') : themeMenu(anchor)) });
  menu(anchor, items);
}
export async function profileDialog() {
  const p = profile();
  const name = h('input', { class: 'input', value: p.name });
  let color = p.color;
  const sw = h('div', { class: 'row wrap' });
  const colors = ['#e5484d', '#f76b15', '#ffb224', '#30a46c', '#12a594', '#0090ff', '#3e63dd', '#8e4ec6', '#d6409f', '#6b7280'];
  const draw = () => { sw.innerHTML = ''; for (const c of colors) sw.append(h('button', { class: 'btn icon', style: { background: c, borderColor: c === color ? 'var(--text)' : 'transparent', borderWidth: '2px', borderRadius: '50%' }, onclick: () => { color = c; draw(); } })); };
  draw();
  const r = await dialog({ title: 'Профиль', body: h('div', {}, h('div', { class: 'field' }, h('label', {}, 'Имя (видно соавторам)'), name), h('div', { class: 'field' }, h('label', {}, 'Цвет курсора'), sw)), buttons: [{ label: 'Отмена', value: false }, { label: 'Сохранить', value: true, primary: true }] });
  if (r) { settings.set('userName', name.value.trim() || p.name); settings.set('userColor', color); }
}
export async function passwordDialog(acc) {
  const old = h('input', { class: 'input', type: 'password', autocomplete: 'current-password' });
  const nw = h('input', { class: 'input', type: 'password', autocomplete: 'new-password' });
  const nw2 = h('input', { class: 'input', type: 'password', autocomplete: 'new-password' });
  const err = h('div', { class: 'small', style: { color: 'var(--danger)', minHeight: '18px' } });
  return dialog({
    title: 'Смена пароля · ' + acc.serverName,
    body: h('div', { style: { width: 'min(380px,100%)' } }, h('div', { class: 'field' }, h('label', {}, 'Текущий пароль'), old), h('div', { class: 'field' }, h('label', {}, 'Новый пароль (не короче 6 символов)'), nw), h('div', { class: 'field' }, h('label', {}, 'Повторите новый пароль'), nw2), err),
    buttons: [{ label: 'Отмена', value: null }, { label: 'Сменить', primary: true, value: async () => {
      if (nw.value.length < 6) { err.textContent = 'Слишком короткий пароль'; return undefined; }
      if (nw.value !== nw2.value) { err.textContent = 'Пароли не совпадают'; return undefined; }
      try { await changePassword(acc, old.value, nw.value); toast('Пароль изменён', 'ok'); return true; } catch (e) { err.textContent = e.message; return undefined; }
    } }],
  });
}
export async function loginDialog(prefill = {}) {
  const url = h('input', { class: 'input', placeholder: 'office.example.com или 192.168.1.10:8080', value: prefill.url || '' });
  const user = h('input', { class: 'input', placeholder: 'Логин', value: prefill.username || '', autocomplete: 'username' });
  const pass = h('input', { class: 'input', type: 'password', placeholder: 'Пароль', autocomplete: 'current-password' });
  const err = h('div', { class: 'small', style: { color: 'var(--danger)', minHeight: '18px' } });
  const body = h('div', { style: { width: 'min(420px, 100%)' } },
    h('p', { class: 'muted', style: { marginTop: 0 } }, 'Подключитесь к своему серверу Rabbit Office: хранилище файлов и общие доски. Логин и пароль выдаёт администратор сервера.'),
    h('div', { class: 'field' }, h('label', {}, 'Адрес сервера'), url), h('div', { class: 'field' }, h('label', {}, 'Логин'), user), h('div', { class: 'field' }, h('label', {}, 'Пароль'), pass), err);
  return dialog({
    title: 'Подключение к серверу', body, buttons: [{ label: 'Отмена', value: null }, {
      label: 'Войти', primary: true, value: async () => {
        err.textContent = '';
        if (!url.value.trim() || !user.value.trim()) { err.textContent = 'Укажите адрес и логин'; return undefined; }
        try { const acc = await login(url.value, user.value.trim(), pass.value); toast(`Вы вошли на ${acc.serverName}`, 'ok'); return acc; }
        catch (e) { err.textContent = e.status === 401 ? 'Неверный логин или пароль' : e.message; return undefined; }
      },
    }],
  });
}

/* =================== Cloud file browser =================== */
export function cloudBrowser({ acc, mode = 'open', exts = [], defaultName = '' }) {
  let path = '/';
  const list = h('div', { class: 'list', style: { minHeight: '260px', maxHeight: '50vh', overflow: 'auto', border: '1px solid var(--border)', borderRadius: 'var(--radius)', padding: '4px' } });
  const crumbs = h('div', { class: 'row wrap', style: { marginBottom: '8px' } });
  const nameIn = h('input', { class: 'input', value: defaultName });
  let resolveSel;
  const load = async () => {
    list.innerHTML = '<div class="empty"><div class="spinner" style="margin:auto"></div></div>';
    crumbs.innerHTML = '';
    const parts = path.split('/').filter(Boolean);
    crumbs.append(h('button', { class: 'btn sm ghost', onclick: () => { path = '/'; load(); }, html: icon('home') + ' ' + esc(acc.serverName) }));
    parts.forEach((p, i) => { crumbs.append(h('span', { class: 'muted' }, '›'), h('button', { class: 'btn sm ghost', onclick: () => { path = '/' + parts.slice(0, i + 1).join('/'); load(); } }, p)); });
    try {
      const r = await cloudFiles.list(acc, path);
      list.innerHTML = '';
      const items = r.items.filter((it) => it.dir || !exts.length || exts.includes(extOf(it.name)));
      if (!items.length) list.append(h('div', { class: 'empty' }, 'Папка пуста'));
      for (const it of items) {
        const a = appForFile(it.name);
        const badge = it.dir ? h('span', { class: 'file-badge', style: { background: 'var(--surface-3)', color: 'var(--text-2)' }, html: icon('folder') }) : h('span', { class: 'file-badge', style: { background: a?.hex || '#888' } }, a?.letter || extOf(it.name).slice(0, 3).toUpperCase());
        const row = h('button', { class: 'list-item' }, badge, h('div', { class: 'grow' }, h('div', { class: 'name' }, it.name), h('div', { class: 'meta' }, it.dir ? 'Папка' : `${fmtSize(it.size)} · ${fmtDate(it.mtime)}`)));
        row.onclick = () => { if (it.dir) { path = (path === '/' ? '' : path) + '/' + it.name; load(); } else if (mode === 'open') resolveSel({ path: (path === '/' ? '' : path) + '/' + it.name, name: it.name }); else nameIn.value = it.name; };
        list.append(row);
      }
    } catch (e) { list.innerHTML = ''; list.append(h('div', { class: 'empty' }, 'Ошибка: ' + e.message)); }
  };
  const mk = h('button', { class: 'btn sm', html: icon('folder-plus') + ' Новая папка', onclick: async () => { const n = await promptDlg('Новая папка', 'Название', 'Новая папка'); if (n) { try { await cloudFiles.mkdir(acc, (path === '/' ? '' : path) + '/' + n); load(); } catch (e) { toast(e.message, 'error'); } } } });
  const body = h('div', { style: { width: 'min(620px, 100%)' } }, h('div', { class: 'row' }, crumbs, h('div', { class: 'grow' }), mk), list);
  if (mode === 'save') body.append(h('div', { class: 'field', style: { marginTop: '12px' } }, h('label', {}, 'Имя файла'), nameIn));
  load();
  return dialog({
    title: mode === 'open' ? 'Открыть с сервера' : 'Сохранить на сервер', body, wide: true,
    onOpen: (api) => { resolveSel = (v) => api.close(v); },
    buttons: mode === 'save' ? [{ label: 'Отмена', value: null }, { label: 'Сохранить', primary: true, value: () => { let n = nameIn.value.trim(); if (!n) return undefined; if (exts[0] && extOf(n) !== exts[0]) n += '.' + exts[0]; return { path: (path === '/' ? '' : path) + '/' + n, name: n }; } }] : [{ label: 'Отмена', value: null }],
  });
}

/* Save-location chooser */
async function saveLocationDialog(shell) {
  const app = shell.app; const name = shell.fileName();
  const opts = [];
  if (native) opts.push({ id: 'native', icon: 'folder', t: 'Этот компьютер', s: 'Выбрать папку и имя файла' });
  else opts.push({ id: 'library', icon: 'folder', t: 'Мои документы', s: isAndroid ? 'Хранится в приложении на устройстве' : 'Хранится в браузере на этом устройстве' });
  for (const a of accounts.list()) if (a.mode !== 'board') opts.push({ id: 'cloud:' + a.id, icon: 'cloud', t: a.serverName, s: `Сервер · ${a.displayName}` });
  if (!native) opts.push({ id: 'download', icon: 'download', t: 'Скачать файл', s: 'Сохранить копию .' + app.ext + ' на диск' });
  opts.push({ id: 'connect', icon: 'plus', t: 'Подключить сервер…', s: 'Свой сервер Rabbit Office' });
  const pick = await new Promise((res) => {
    const body = h('div', { class: 'list', style: { width: 'min(460px, 100%)' } });
    let api;
    for (const o of opts) body.append(h('button', { class: 'list-item', onclick: () => api.close(o.id), html: `<span class="file-badge" style="background:var(--accent-soft);color:var(--accent)">${icon(o.icon)}</span><div class="grow"><div class="name"><b>${esc(o.t)}</b></div><div class="meta">${esc(o.s)}</div></div>` }));
    dialog({ title: 'Сохранить как', body, buttons: [], onOpen: (a) => { api = a; } }).then(res);
  });
  if (!pick) return null;
  if (pick === 'native') { const path = await native.saveDialog({ defaultName: name, filters: [{ name: app.name, extensions: [app.ext] }] }); return path ? { loc: { source: 'native', path }, name: path.split(/[\\/]/).pop() } : null; }
  if (pick === 'library') { const n = await promptDlg('Сохранить в «Мои документы»', 'Имя документа', baseName(name)); if (!n) return null; return { loc: { source: 'library', libId: shell.loc?.source === 'library' ? shell.loc.libId : uid('d') }, name: n + '.' + app.ext }; }
  if (pick === 'download') { await shell.export(name, await shell.opts.serialize()); shell.markClean(); return null; }
  if (pick === 'connect') { const acc = await loginDialog(); if (!acc) return null; return saveLocationDialog(shell); }
  if (pick.startsWith('cloud:')) {
    const acc = accounts.get(pick.slice(6));
    const r = await cloudBrowser({ acc, mode: 'save', exts: [app.ext], defaultName: name });
    return r ? { loc: { source: 'cloud', account: acc.id, path: r.path }, name: r.name } : null;
  }
  return null;
}

export function greeting() {
  const hr = new Date().getHours();
  return hr < 5 ? 'Доброй ночи' : hr < 12 ? 'Доброе утро' : hr < 18 ? 'Добрый день' : 'Добрый вечер';
}

/* =================== Backstage =================== */
function openBackstage(shell, page) {
  const app = shell.app; const o = shell.opts;
  const bs = h('div', { class: 'backstage' });
  const nav = h('nav', { class: 'bs-nav' });
  const main = h('div', { class: 'bs-main' });
  const close = () => { bs.remove(); document.removeEventListener('keydown', esc_); o.focusEditor?.(); };
  const esc_ = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', esc_);
  const pages = [
    ['home', 'home', 'Главная'], ['new', 'file-plus', 'Создать'], ['open', 'folder', 'Открыть'], '-',
    ['info', 'info', 'Сведения'], ['save', 'save', 'Сохранить'], ['saveas', 'copy', 'Сохранить как'], ['export', 'share', 'Экспорт'],
    o.print ? ['print', 'print', 'Печать'] : null, '-', ['settings', 'settings', 'Параметры'], ['about', 'help', 'О программе'],
  ].filter(Boolean);
  nav.append(h('button', { class: 'back', onclick: close, html: icon('back') + '<span>Назад</span>' }));
  const btns = {};
  for (const p of pages) {
    if (p === '-') { nav.append(h('div', { class: 'sep' })); continue; }
    const [id, ic, label] = p;
    btns[id] = h('button', { onclick: () => show(id), html: icon(ic) + `<span>${label}</span>` });
    nav.append(btns[id]);
  }
  bs.append(nav, main);
  document.body.append(bs);

  const recentList = (filterApp) => {
    const l = recent.list().filter((e) => !filterApp || e.app === filterApp);
    const box = h('div', { class: 'list' });
    if (!l.length) box.append(h('div', { class: 'empty' }, 'Здесь появятся недавние документы'));
    l.slice(0, 15).forEach((e) => {
      const a = APPS[e.app];
      const where = e.source === 'native' ? e.path : e.source === 'library' ? 'Мои документы' : (accounts.get(e.account)?.serverName || 'Сервер') + ' · ' + e.cloudPath;
      box.append(h('button', { class: 'list-item', onclick: async () => { close(); const idx = recent.list().indexOf(recent.list().find((x) => x.time === e.time)); if (e.app !== app.id || shell.dirty || !o.isEmpty?.()) { if (native || e.app !== app.id) openApp(e.app, { recent: idx }); else shell.loadFile(await readFromLoc(locFromRecent(e), e.name)); } else shell.loadFile(await readFromLoc(locFromRecent(e), e.name)); },
        html: `<span class="file-badge" style="background:${a?.hex}">${a?.letter}</span><div class="grow"><div class="name">${esc(e.name)}</div><div class="meta">${esc(where)}</div></div><span class="meta">${fmtDate(e.time)}</span>` }));
    });
    return box;
  };
  const templateCards = () => {
    const c = h('div', { class: 'cards' });
    for (const t of o.templates || [{ id: '', name: 'Пустой документ' }]) {
      c.append(h('button', { class: 'card', onclick: () => { close(); shell.newDoc(t.id); }, html: `<div class="thumb">${t.thumb || ''}</div><div class="t">${esc(t.name)}</div>${t.desc ? `<div class="s">${esc(t.desc)}</div>` : ''}` }));
    }
    return c;
  };
  function show(id) {
    for (const [k, b] of Object.entries(btns)) b.classList.toggle('on', k === id);
    main.innerHTML = '';
    if (id === 'home') {
      main.append(h('h1', {}, greeting()), h('h2', {}, 'Создать'), templateCards(), h('h2', {}, 'Недавние'), recentList(app.id));
    } else if (id === 'new') {
      main.append(h('h1', {}, 'Создать'), templateCards());
    } else if (id === 'open') {
      main.append(h('h1', {}, 'Открыть'));
      const row = h('div', { class: 'cards', style: { gridTemplateColumns: 'repeat(auto-fill,minmax(200px,1fr))' } });
      const srcCard = (ic, t, s, fn) => h('button', { class: 'card', onclick: fn, html: `<div style="display:flex;gap:10px;align-items:center"><span class="file-badge" style="background:var(--accent-soft);color:var(--accent)">${icon(ic)}</span><div><div class="t">${esc(t)}</div><div class="s">${esc(s)}</div></div></div>` });
      row.append(srcCard('folder', native ? 'Этот компьютер' : 'Файл с устройства', app.open.map((e) => '.' + e).join(' '), async () => { close(); shell.open(); }));
      if (!native) row.append(srcCard('folder', 'Мои документы', 'Сохранённые в приложении', () => libraryPage()));
      for (const a of accounts.list()) if (a.mode !== 'board') row.append(srcCard('cloud', a.serverName, 'Сервер · ' + a.displayName, async () => { const r = await cloudBrowser({ acc: a, mode: 'open', exts: app.open }); if (r) { close(); const f = await readFromLoc({ source: 'cloud', account: a.id, path: r.path }, r.name); if (shell.dirty || !o.isEmpty?.()) { if (native) await handoff(app.id, f); else shell.loadFile(f); } else shell.loadFile(f); } }));
      row.append(srcCard('plus', 'Подключить сервер', 'Свой сервер Rabbit Office', async () => { if (await loginDialog()) show('open'); }));
      main.append(row, h('h2', {}, 'Недавние'), recentList());
    } else if (id === 'info') {
      const stats = o.stats?.() || [];
      main.append(h('h1', {}, 'Сведения'), h('div', { class: 'card', style: { maxWidth: '640px' }, html: `<div class="t" style="font-size:18px">${esc(shell.title)}</div><div class="s">${esc(locLabel(shell.loc))}</div>
        <table class="data" style="margin-top:10px">${stats.map(([k, v]) => `<tr><td class="muted">${esc(k)}</td><td>${esc(v)}</td></tr>`).join('')}
        <tr><td class="muted">Автор</td><td>${esc(profile().name)}</td></tr><tr><td class="muted">Состояние</td><td>${shell.dirty ? 'Есть несохранённые изменения' : 'Сохранено'}</td></tr></table>` }));
    } else if (id === 'save') { close(); shell.save(); }
    else if (id === 'saveas') { close(); shell.saveAs(); }
    else if (id === 'export') {
      main.append(h('h1', {}, 'Экспорт'));
      const c = h('div', { class: 'list', style: { maxWidth: '640px' } });
      for (const ex of o.exporters || []) c.append(h('button', { class: 'list-item', onclick: async () => { close(); try { await ex.run(); } catch (e) { console.error(e); toast('Ошибка экспорта: ' + e.message, 'error', 5000); } }, html: `<span class="file-badge" style="background:var(--accent-soft);color:var(--accent)">${icon(ex.icon || 'download')}</span><div class="grow"><div class="name"><b>${esc(ex.label)}</b></div><div class="meta">${esc(ex.desc || '')}</div></div>` }));
      main.append(c);
    } else if (id === 'print') { close(); o.print(); }
    else if (id === 'settings') {
      main.append(h('h1', {}, 'Параметры'), h('h2', {}, 'Оформление'), themeCards(), h('h2', {}, 'Профиль'),
        h('div', { class: 'row' }, h('span', { class: 'avatar', style: { background: profile().color, width: '32px', height: '32px', borderRadius: '50%', display: 'grid', placeItems: 'center', color: '#fff', fontWeight: 700 } }, profile().name[0]), h('b', {}, profile().name), h('button', { class: 'btn sm', onclick: () => profileDialog().then(() => show('settings')) }, 'Изменить')),
        h('h2', {}, 'Редактирование'),
        (() => { const c = h('input', { type: 'checkbox', checked: settings.get('autosave') }); c.onchange = () => settings.set('autosave', c.checked); return h('label', { class: 'check' }, c, 'Автосохранение (когда у документа есть место сохранения)'); })(), h('br'),
        (() => { const c = h('input', { type: 'checkbox', checked: settings.get('spellcheck') }); c.onchange = () => settings.set('spellcheck', c.checked); return h('label', { class: 'check', style: { marginTop: '8px' } }, c, 'Проверка орфографии'); })(),
        ...(o.settingsExtra ? [o.settingsExtra()] : []));
    } else if (id === 'about') {
      main.append(h('h1', {}, 'О программе'), h('div', { class: 'card', style: { maxWidth: '560px' }, html: `<div style="display:flex;gap:14px;align-items:center"><span class="logo logo-squares" style="width:44px;height:44px;gap:4px"><i></i><i></i><i></i><i></i></span><div><div class="t" style="font-size:18px">${esc(app.name)}</div><div class="s">Rabbit Office 1.0 · ${esc(platform)}</div></div></div>
        <p>Свободный офисный пакет без подписок: Documents, Xel, Point и Desk. Работает на Windows, macOS, Linux и Android. Файлы можно хранить на устройстве или на собственном сервере Rabbit Office.</p>
        <p class="muted small">Горячие клавиши: Ctrl+S — сохранить, Ctrl+Shift+S — сохранить как, Ctrl+O — открыть, Ctrl+N — новый, Ctrl+P — печать, Ctrl+F1 — свернуть ленту.</p>` }));
    }
  }
  async function libraryPage() {
    for (const b of Object.values(btns)) b.classList.remove('on');
    main.innerHTML = ''; main.append(h('h1', {}, 'Мои документы'));
    const l = (await library.list()).filter((d) => app.open.includes(extOf(d.name)));
    const box = h('div', { class: 'list', style: { maxWidth: '760px' } });
    if (!l.length) box.append(h('div', { class: 'empty' }, 'Нет сохранённых документов'));
    for (const d of l) {
      const row = h('div', { class: 'list-item' });
      row.innerHTML = `<span class="file-badge" style="background:${app.hex}">${app.letter}</span><div class="grow"><div class="name">${esc(d.name)}</div><div class="meta">${fmtDate(d.updated)} · ${fmtSize(d.size)}</div></div>`;
      const del = h('button', { class: 'btn ghost icon', html: icon('trash'), 'data-tip': 'Удалить', onclick: async (e) => { e.stopPropagation(); if (await confirmDlg('Удалить документ', `Удалить «${d.name}»?`, 'Удалить', true)) { await library.remove(d.id); libraryPage(); } } });
      row.append(del);
      row.onclick = async () => { close(); const f = await readFromLoc({ source: 'library', libId: d.id }); shell.loadFile(f); };
      box.append(row);
    }
    main.append(box);
  }
  show(page);
}

export { icon, h, $, $$, esc, toast, dialog, menu };
