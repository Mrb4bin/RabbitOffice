// Rabbit Office hub — launcher, recent files, server, settings
import { icon } from '../core/icons.js';
import { h, esc, toast, confirmDlg, promptDlg, fmtDate, fmtSize, menu } from '../core/ui.js';
import { recent, library, profile, settings } from '../core/store.js';
import { accounts, files as cloudFiles, boards as cloudBoards, sameOriginServer, logout } from '../core/cloud.js';
import { APPS, openApp, openLocalFile, appForFile, handoff, readFromLoc, locFromRecent, native, extOf, platform } from '../core/platform.js';
import { applyTheme, themeMenu, themeCards, accountMenu, loginDialog, profileDialog, greeting } from '../core/shell.js';
import { TEMPLATES } from '../core/templates.js';

applyTheme();
document.documentElement.dataset.app = 'hub';
const DESC = {
  docs: 'Текстовые документы: стили, таблицы, изображения, комментарии, экспорт в DOCX и PDF.',
  xel: 'Электронные таблицы: 120+ функций, диаграммы, фильтры, условное форматирование, XLSX и CSV.',
  point: 'Презентации: темы, фигуры, анимации и переходы, режим докладчика, экспорт в PPTX и PDF.',
  desk: 'Онлайн-доска: стикеры, фигуры, связи, рисование. Совместно — по IP, в локальной сети или через сервер.',
};

/* ---------- title bar ---------- */
const tb = h('header', { class: 'titlebar' });
tb.append(h('span', { class: 'logo logo-squares', style: { marginLeft: '4px' }, html: '<i></i><i></i><i></i><i></i>' }), h('span', { class: 'app-name', style: { marginLeft: '6px' } }, 'Rabbit Office'), h('div', { class: 'grow' }));
const openBtn = h('button', { class: 'tb-btn', 'data-tip': 'Открыть файл|Ctrl+O', html: icon('folder') + '<span class="hide-m">Открыть</span>', onclick: () => openAny() });
const themeBtn = h('button', { class: 'tb-btn', 'data-tip': 'Оформление', html: icon('palette'), onclick: (e) => themeMenu(e.currentTarget) });
const acctBtn = h('button', { class: 'tb-btn', 'data-tip': 'Учётная запись' });
acctBtn.onclick = () => accountMenu(acctBtn, null);
tb.append(openBtn, themeBtn, acctBtn);
const updAcct = () => { const p = profile(); const a = accounts.active(); acctBtn.innerHTML = `<span class="avatar" style="background:${p.color}">${esc((a?.displayName || p.name).slice(0, 1).toUpperCase())}</span>`; };

/* ---------- layout ---------- */
const nav = h('nav', { class: 'hub-nav' });
const main = h('div', { class: 'hub-main' });
document.body.append(tb, h('div', { class: 'hub' }, nav, main));
const pages = [['home', 'home', 'Главная'], ['files', 'folder', 'Мои документы'], ['server', 'server', 'Сервер'], ['settings', 'settings', 'Параметры']];
const navBtns = {};
for (const [id, ic, label] of pages) { navBtns[id] = h('button', { onclick: () => show(id), html: icon(ic) + `<span>${label}</span>` }); nav.append(navBtns[id]); }
const mini = h('div', { class: 'apps-mini' });
for (const a of Object.values(APPS)) mini.append(h('button', { onclick: () => openApp(a.id), html: `<span class="logo single" style="width:14px;height:14px;background:${a.hex};border-radius:4px"></span>${a.short}` }));
nav.append(mini);

let current = 'home';
function show(id) { current = id; for (const [k, b] of Object.entries(navBtns)) b.classList.toggle('on', k === id); main.innerHTML = ''; ({ home, files, server, settingsPage })[id === 'settings' ? 'settingsPage' : id](); }
window.addEventListener('rabbit-accounts', () => { updAcct(); show(current); });
window.addEventListener('rabbit-settings', updAcct);
updAcct();

async function openAny() {
  const all = Object.values(APPS).flatMap((a) => a.open);
  const f = await openLocalFile(all); if (!f) return;
  const a = appForFile(f.name); if (!a) return toast('Формат не поддерживается', 'error');
  await handoff(a.id, f);
}

/* ---------- pages ---------- */
function home() {
  main.append(h('h1', {}, `${greeting()}, ${profile().name.split(' ')[0]}`), h('div', { class: 'muted' }, 'Бесплатный офис без подписок. Выберите приложение или продолжите работу.'));
  sameOriginServer().then((srv) => {
    if (srv && !accounts.list().some((a) => a.url === srv.url)) {
      const b = h('div', { class: 'banner', html: `${icon('server')}<div class="grow"><b>${esc(srv.name)}</b><div class="small muted">Войдите, чтобы открыть файлы и доски этого сервера</div></div>` });
      b.append(h('button', { class: 'btn primary', onclick: () => loginDialog({ url: srv.url }) }, 'Войти'));
      main.insertBefore(b, main.children[2]);
    }
  });
  const tiles = h('div', { class: 'app-tiles' });
  for (const a of Object.values(APPS)) {
    const t = h('div', { class: 'app-tile', style: { '--tile': a.hex } });
    t.innerHTML = `<div class="sq">${a.letter}</div><div class="nm">${a.name}</div><div class="ds">${DESC[a.id]}</div>`;
    const acts = h('div', { class: 'acts' }, h('button', { class: 'btn primary', onclick: (e) => { e.stopPropagation(); openApp(a.id); } }, 'Создать'),
      h('button', { class: 'btn', onclick: (e) => { e.stopPropagation(); menu(e.currentTarget, TEMPLATES[a.id].map((tp) => ({ label: tp.name, hint: tp.desc, icon: 'template', action: () => openApp(a.id, tp.id ? { template: tp.id } : {}) }))); } }, 'Шаблоны'));
    t.append(acts); t.onclick = () => openApp(a.id);
    tiles.append(t);
  }
  main.append(tiles);

  main.append(h('h2', {}, 'Популярные шаблоны'));
  const tr = h('div', { class: 'tpl-row' });
  for (const [aid, list] of Object.entries(TEMPLATES)) for (const tp of list.filter((x) => x.id).slice(0, 2)) {
    tr.append(h('button', { class: 'tpl', onclick: () => openApp(aid, { template: tp.id }), html: `<div class="thumb">${tp.thumb}</div><div class="t">${esc(tp.name)}</div><div class="s">${APPS[aid].name}</div>` }));
  }
  main.append(tr);

  let filter = ''; let query = '';
  const search = h('input', { class: 'input', placeholder: 'Поиск по названию', style: { width: '220px', height: '30px' } });
  search.oninput = () => { query = search.value.trim().toLowerCase(); drawRecent(); };
  const head = h('h2', {}, 'Недавние', h('span', { class: 'grow' }), search);
  const seg = h('div', { class: 'seg' });
  for (const [id, label] of [['', 'Все'], ['docs', 'Documents'], ['xel', 'Xel'], ['point', 'Point'], ['desk', 'Desk']]) {
    const b = h('button', { class: id === filter ? 'on' : '', onclick: () => { filter = id; [...seg.children].forEach((x) => x.classList.toggle('on', x === b)); drawRecent(); } }, label);
    seg.append(b);
  }
  head.append(seg);
  const table = h('table', { class: 'recent-table' });
  main.append(head, table);
  function drawRecent() {
    table.innerHTML = '';
    const l = recent.list().map((e, i) => ({ ...e, idx: i })).filter((e) => (!filter || e.app === filter) && (!query || e.name.toLowerCase().includes(query)));
    if (!l.length) { table.innerHTML = `<tr><td class="empty">Пока пусто. Созданные и открытые файлы появятся здесь.</td></tr>`; return; }
    for (const e of l) {
      const a = APPS[e.app]; if (!a) continue;
      const where = e.source === 'native' ? e.path : e.source === 'library' ? 'Мои документы' : e.source === 'cloud' ? `${accounts.get(e.account)?.serverName || 'Сервер'} · ${e.cloudPath}` : e.source === 'board' ? 'Доска на сервере' : '';
      const tr = h('tr');
      tr.innerHTML = `<td><div class="nm"><span class="file-badge" style="background:${a.hex}">${a.letter}</span>${esc(e.name)}</div></td><td class="wh">${esc(where)}</td><td class="muted small hide-m">${fmtDate(e.time)}</td>`;
      const more = h('button', { class: 'btn ghost icon', html: icon('more-h'), onclick: (ev) => { ev.stopPropagation(); menu(ev.currentTarget, [{ label: 'Открыть', icon: 'folder', action: () => openApp(e.app, { recent: e.idx }) }, { label: 'Убрать из списка', icon: 'close', action: () => { recent.remove((x) => x.time === e.time); drawRecent(); } }]); } });
      const td = h('td', { style: { width: '40px' } }, more); tr.append(td);
      tr.onclick = () => openApp(e.app, { recent: e.idx });
      table.append(tr);
    }
  }
  drawRecent();
}

async function files() {
  main.append(h('h1', {}, 'Мои документы'));
  if (native) {
    main.append(h('p', { class: 'muted' }, 'На компьютере документы хранятся там, где вы их сохранили. Недавние файлы — на главной.'),
      h('button', { class: 'btn primary', onclick: openAny, html: icon('folder') + 'Открыть файл…' }));
  }
  const l = await library.list();
  if (!native || l.length) {
    main.append(h('h2', {}, 'Сохранённые в приложении', h('span', { class: 'grow' }), h('button', { class: 'btn sm', onclick: openAny, html: icon('upload') + 'Импорт файла' })));
    const box = h('table', { class: 'recent-table' });
    if (!l.length) box.innerHTML = '<tr><td class="empty">Нет документов. Сохраните документ через «Файл → Сохранить как → Мои документы».</td></tr>';
    for (const d of l) {
      const a = appForFile(d.name) || APPS[d.app]; if (!a) continue;
      const tr = h('tr', { html: `<td><div class="nm"><span class="file-badge" style="background:${a.hex}">${a.letter}</span>${esc(d.name)}</div></td><td class="muted small">${fmtDate(d.updated)}</td><td class="muted small hide-m">${fmtSize(d.size)}</td>` });
      const del = h('button', { class: 'btn ghost icon', html: icon('trash'), onclick: async (e) => { e.stopPropagation(); if (await confirmDlg('Удалить', `Удалить «${d.name}» без возможности восстановления?`, 'Удалить', true)) { await library.remove(d.id); show('files'); } } });
      tr.append(h('td', { style: { width: '40px' } }, del));
      tr.onclick = async () => handoff(a.id, await readFromLoc({ source: 'library', libId: d.id }));
      box.append(tr);
    }
    main.append(box);
  }
  for (const acc of accounts.list().filter((x) => x.mode !== 'board')) {
    main.append(h('h2', {}, `${acc.serverName}`, h('span', { class: 'badge' }, 'сервер')));
    main.append(serverFiles(acc));
  }
}

function serverFiles(acc) {
  let path = '/';
  const wrap = h('div');
  const load = async () => {
    wrap.innerHTML = '<div class="empty"><div class="spinner" style="margin:auto"></div></div>';
    try {
      const r = await cloudFiles.list(acc, path);
      wrap.innerHTML = '';
      const bar = h('div', { class: 'row', style: { marginBottom: '8px' } });
      const parts = path.split('/').filter(Boolean);
      bar.append(h('button', { class: 'btn sm ghost', onclick: () => { path = '/'; load(); }, html: icon('home') }));
      parts.forEach((p, i) => bar.append(h('span', { class: 'muted' }, '›'), h('button', { class: 'btn sm ghost', onclick: () => { path = '/' + parts.slice(0, i + 1).join('/'); load(); } }, p)));
      bar.append(h('span', { class: 'grow' }),
        h('button', { class: 'btn sm', html: icon('folder-plus') + 'Папка', onclick: async () => { const n = await promptDlg('Новая папка', 'Название', 'Новая папка'); if (n) { await cloudFiles.mkdir(acc, (path === '/' ? '' : path) + '/' + n).catch((e) => toast(e.message, 'error')); load(); } } }),
        h('button', { class: 'btn sm', html: icon('upload') + 'Загрузить', onclick: async () => { const f = await openLocalFile(Object.values(APPS).flatMap((a) => a.open).concat(['pdf', 'png', 'jpg', 'zip'])); if (f) { await cloudFiles.upload(acc, (path === '/' ? '' : path) + '/' + f.name, f.bytes).catch((e) => toast(e.message, 'error')); load(); } } }));
      wrap.append(bar);
      const t = h('table', { class: 'recent-table' });
      if (!r.items.length) t.innerHTML = '<tr><td class="empty">Папка пуста</td></tr>';
      if (r.quota) wrap.append(h('div', { class: 'small muted', style: { marginBottom: '6px' } }, `Занято ${fmtSize(r.used || 0)}${r.quota ? ' из ' + fmtSize(r.quota) : ''}`));
      for (const it of r.items) {
        const a = appForFile(it.name);
        const full = (path === '/' ? '' : path) + '/' + it.name;
        const tr = h('tr', { html: `<td><div class="nm"><span class="file-badge" style="background:${it.dir ? 'var(--surface-3)' : a?.hex || '#888'};color:${it.dir ? 'var(--text-2)' : '#fff'}">${it.dir ? icon('folder') : a?.letter || esc(extOf(it.name).slice(0, 3).toUpperCase())}</span>${esc(it.name)}</div></td><td class="muted small">${it.dir ? 'Папка' : fmtSize(it.size)}</td><td class="muted small hide-m">${fmtDate(it.mtime)}</td>` });
        const more = h('button', { class: 'btn ghost icon', html: icon('more-h'), onclick: (e) => { e.stopPropagation(); menu(e.currentTarget, [
          { label: 'Переименовать', icon: 'edit', action: async () => { const n = await promptDlg('Переименовать', 'Новое имя', it.name); if (n && n !== it.name) { await cloudFiles.move(acc, full, (path === '/' ? '' : path) + '/' + n).catch((er) => toast(er.message, 'error')); load(); } } },
          !it.dir ? { label: 'Скачать', icon: 'download', action: async () => { const { exportFile } = await import('../core/platform.js'); exportFile(it.name, await cloudFiles.download(acc, full)); } } : null,
          { label: 'Удалить', icon: 'trash', danger: true, action: async () => { if (await confirmDlg('Удалить', `Удалить «${it.name}»?`, 'Удалить', true)) { await cloudFiles.remove(acc, full).catch((er) => toast(er.message, 'error')); load(); } } },
        ]); } });
        tr.append(h('td', { style: { width: '40px' } }, more));
        tr.onclick = async () => {
          if (it.dir) { path = full; load(); return; }
          if (!a) return toast('Этот формат не открывается в Rabbit Office', 'error');
          handoff(a.id, await readFromLoc({ source: 'cloud', account: acc.id, path: full }, it.name));
        };
        t.append(tr);
      }
      wrap.append(t);
    } catch (e) { wrap.innerHTML = ''; wrap.append(h('div', { class: 'empty' }, 'Ошибка: ' + e.message)); }
  };
  load();
  return wrap;
}

async function server() {
  main.append(h('h1', {}, 'Сервер Rabbit Office'), h('p', { class: 'muted', style: { maxWidth: '680px' } }, 'Собственный сервер на Linux: арендуйте VPS, установите Rabbit Office Server одной командой — и получите хранилище файлов, общие доски и учётные записи для своей команды. Никаких подписок и чужих облаков.'));
  const list = accounts.list();
  for (const a of list) {
    const c = h('div', { class: 'srv-card', html: `<span class="file-badge" style="background:var(--accent);width:40px;height:40px">${icon('server')}</span><div class="grow"><div style="font-weight:600">${esc(a.serverName)}</div><div class="small muted">${esc(a.url)} · ${esc(a.displayName)} ${a.role === 'admin' ? '· администратор' : ''} · режим: ${({ storage: 'хранилище', board: 'доски', both: 'хранилище и доски' })[a.mode] || a.mode}</div></div>` });
    if (accounts.active()?.id === a.id) c.querySelector('.grow').append(h('span', { class: 'badge' }, 'активный'));
    else c.append(h('button', { class: 'btn sm', onclick: () => accounts.setActive(a.id) }, 'Сделать активным'));
    if (a.role === 'admin') c.append(h('button', { class: 'btn sm', onclick: () => window.open(a.url + '/admin/', '_blank'), html: icon('settings') + 'Админ-панель' }));
    c.append(h('button', { class: 'btn sm danger', onclick: async () => { await logout(a); show('server'); } }, 'Выйти'));
    main.append(c);
  }
  main.append(h('button', { class: 'btn primary', onclick: () => loginDialog(), html: icon('plus') + 'Подключить сервер' }));
  const act = accounts.active();
  if (act && act.mode !== 'storage') {
    main.append(h('h2', {}, 'Доски на сервере', h('span', { class: 'grow' }), h('button', { class: 'btn sm', html: icon('plus') + 'Новая доска', onclick: async () => { const t = await promptDlg('Новая доска', 'Название', 'Новая доска'); if (t) { const b = await cloudBoards.create(act, t).catch((e) => toast(e.message, 'error')); if (b) openApp('desk', { board: b.id, account: act.id }); } } })));
    const t = h('table', { class: 'recent-table' }); main.append(t);
    try {
      const bl = await cloudBoards.list(act);
      if (!bl.length) t.innerHTML = '<tr><td class="empty">Досок пока нет</td></tr>';
      for (const b of bl) {
        const tr = h('tr', { html: `<td><div class="nm"><span class="file-badge" style="background:${APPS.desk.hex}">B</span>${esc(b.title)}</div></td><td class="muted small">${b.online ? `${b.online} онлайн` : ''}</td><td class="muted small hide-m">${fmtDate(b.updated)}</td>` });
        tr.onclick = () => openApp('desk', { board: b.id, account: act.id });
        t.append(tr);
      }
    } catch (e) { t.innerHTML = `<tr><td class="empty">${esc(e.message)}</td></tr>`; }
  }
  main.append(h('h2', {}, 'Как установить свой сервер'), h('div', { class: 'card', style: { maxWidth: '720px' }, html: `<ol style="margin:0;padding-left:18px;line-height:1.8">
    <li>Арендуйте VPS с Ubuntu/Debian (подойдёт 1 ГБ ОЗУ) и направьте на него домен (A-запись).</li>
    <li>Скопируйте папку <span class="kbd">rabbit-office</span> на сервер и выполните <span class="kbd">sudo ./server/install.sh</span>.</li>
    <li>Мастер спросит домен, режим (хранилище / доски / оба) и создаст администратора. Сертификат HTTPS получится автоматически.</li>
    <li>Создайте пользователей в админ-панели <span class="kbd">https://ваш-домен/admin</span> и раздайте логины.</li></ol>` }));
}

function settingsPage() {
  main.append(h('h1', {}, 'Параметры'), h('h2', {}, 'Оформление'), themeCards());
  const p = profile();
  main.append(h('h2', {}, 'Профиль'), h('div', { class: 'row' }, h('span', { style: { background: p.color, width: '36px', height: '36px', borderRadius: '50%', display: 'grid', placeItems: 'center', color: '#fff', fontWeight: 700 } }, p.name[0]), h('div', {}, h('b', {}, p.name), h('div', { class: 'small muted' }, 'Имя и цвет видны соавторам на доске')), h('button', { class: 'btn sm', onclick: () => profileDialog().then(() => show('settings')) }, 'Изменить')));
  main.append(h('h2', {}, 'Rabbit Desk'));
  const port = h('input', { class: 'input', type: 'number', value: settings.get('deskPort'), style: { width: '120px' } });
  port.onchange = () => settings.set('deskPort', Number(port.value) || 47800);
  main.append(h('div', { class: 'field', style: { maxWidth: '360px' } }, h('label', {}, 'Порт для прямого подключения и локальной сети'), port));
  main.append(h('h2', {}, 'О программе'), h('div', { class: 'muted' }, `Rabbit Office 1.0 · платформа: ${platform}. Бесплатно и без подписок.`));
}

show('home');
document.addEventListener('keydown', (e) => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') { e.preventDefault(); openAny(); } });
// drag & drop files onto the hub
let dropHint;
document.addEventListener('dragover', (e) => { e.preventDefault(); if (!dropHint) { dropHint = h('div', { class: 'drop-hint' }, 'Отпустите, чтобы открыть'); document.body.append(dropHint); } });
document.addEventListener('dragleave', (e) => { if (!e.relatedTarget) { dropHint?.remove(); dropHint = null; } });
document.addEventListener('drop', async (e) => {
  e.preventDefault(); dropHint?.remove(); dropHint = null;
  const f = e.dataTransfer.files[0]; if (!f) return;
  const a = appForFile(f.name); if (!a) return toast('Формат не поддерживается', 'error');
  const path = native?.pathForFile?.(f);
  if (path) return openApp(a.id, { path });
  handoff(a.id, { name: f.name, bytes: new Uint8Array(await f.arrayBuffer()), loc: null });
});
