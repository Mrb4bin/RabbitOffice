import { h, esc, toast, dialog, confirmDlg, formDlg, fmtSize, fmtDate } from '/core/ui.js';
import { applyTheme } from '/core/shell.js';
applyTheme();
const app = document.getElementById('app');
let token = sessionStorage.getItem('rabbit.admin.token');
const api = async (method, path, body) => {
  const r = await fetch(path, { method, headers: { Authorization: 'Bearer ' + token, ...(body ? { 'Content-Type': 'application/json' } : {}) }, body: body ? JSON.stringify(body) : undefined });
  const d = await r.json().catch(() => ({}));
  if (r.status === 401) { token = null; sessionStorage.removeItem('rabbit.admin.token'); render(); }
  if (!r.ok) throw new Error(d.error || r.statusText);
  return d;
};
const genPass = () => Array.from(crypto.getRandomValues(new Uint8Array(10)), (b) => 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'[b % 55]).join('');
document.getElementById('logout').onclick = async () => { await api('POST', '/api/logout').catch(() => {}); token = null; sessionStorage.removeItem('rabbit.admin.token'); render(); };

async function render() {
  app.innerHTML = '';
  document.getElementById('logout').style.display = token ? '' : 'none';
  if (!token) return loginView();
  let st; try { st = await api('GET', '/api/admin/status'); } catch (e) { app.append(h('div', { class: 'box' }, e.message)); return; }
  const c = st.config;
  app.append(h('h1', {}, c.serverName), h('div', { class: 'muted' }, `Версия ${st.version} · Node ${st.node} · данные: ${st.dataDir}`));
  app.append(h('div', { class: 'stats', html: [['Пользователей', st.users], ['Досок', st.boards], ['Сейчас онлайн', st.online], ['Занято в хранилище', fmtSize(st.storageBytes)], ['Работает', Math.floor(st.uptime / 3600) + ' ч ' + Math.floor((st.uptime % 3600) / 60) + ' мин'], ['HTTPS', c.domain ? (st.cert ? 'до ' + new Date(st.cert.validTo).toLocaleDateString('ru-RU') : 'ожидание') : 'выключен']].map(([k, v]) => `<div class="stat"><span>${k}</span><b>${esc(v)}</b></div>`).join('') }));
  // users
  const users = await api('GET', '/api/admin/users');
  const ub = h('div', { class: 'box' }, h('div', { class: 'row', style: { marginBottom: '10px' } }, h('h2', { style: { margin: 0 } }, 'Пользователи'), h('span', { class: 'grow' }), h('button', { class: 'btn primary', onclick: addUser }, '+ Новый пользователь')));
  const t = h('table', { class: 'data' }); t.innerHTML = '<tr><th>Логин</th><th>Имя</th><th>Роль</th><th>Хранилище</th><th>Создан</th><th></th></tr>';
  for (const u of users) {
    const tr = h('tr', { html: `<td><b>${esc(u.username)}</b>${u.disabled ? ' <span class="badge" style="color:var(--danger)">отключён</span>' : ''}</td><td>${esc(u.displayName)}</td><td>${u.role === 'admin' ? 'Администратор' : 'Пользователь'}</td><td>${fmtSize(u.used)}${u.quotaMB ? ' / ' + u.quotaMB + ' МБ' : ''}</td><td>${fmtDate(u.created)}</td>` });
    const td = h('td', { style: { whiteSpace: 'nowrap', textAlign: 'right' } },
      h('button', { class: 'btn sm', onclick: () => editUser(u) }, 'Изменить'), ' ',
      h('button', { class: 'btn sm', onclick: async () => { const p = genPass(); if (!(await confirmDlg('Сброс пароля', `Новый пароль для ${u.username}: ${p}\nСохраните его и передайте пользователю.`, 'Установить'))) return; await api('PATCH', '/api/admin/users/' + encodeURIComponent(u.username), { password: p }).then(() => toast('Пароль изменён', 'ok')).catch((e) => toast(e.message, 'error')); } }, 'Сброс пароля'), ' ',
      h('button', { class: 'btn sm danger', onclick: async () => { if (await confirmDlg('Удалить пользователя', `Удалить ${u.username}? Его файлы останутся в папке данных.`, 'Удалить', true)) { await api('DELETE', '/api/admin/users/' + encodeURIComponent(u.username)).catch((e) => toast(e.message, 'error')); render(); } } }, 'Удалить'));
    tr.append(td); t.append(tr);
  }
  ub.append(t); app.append(ub);
  // settings
  const sb = h('div', { class: 'box' }, h('h2', {}, 'Настройки сервера'));
  const name = h('input', { class: 'input', value: c.serverName });
  const mode = h('div', { class: 'mode', html: [['storage', 'Хранилище', 'Файлы пользователей и общая папка'], ['board', 'Доски', 'Совместные доски Rabbit Desk'], ['both', 'Хранилище и доски', 'Все возможности']].map(([v, n, d]) => `<label><input type="radio" name="mode" value="${v}" ${c.mode === v ? 'checked' : ''}><span><b>${n}</b><br><span class="small muted">${d}</span></span></label>`).join('') });
  const domain = h('input', { class: 'input', value: c.domain || '', placeholder: 'office.example.com' });
  const email = h('input', { class: 'input', value: c.email || '', placeholder: 'admin@example.com' });
  const upl = h('input', { class: 'input', type: 'number', value: c.maxUploadMB }); const quota = h('input', { class: 'input', type: 'number', value: c.defaultQuotaMB });
  sb.append(h('div', { class: 'field' }, h('label', {}, 'Название сервера'), name), h('div', { class: 'field' }, h('label', {}, 'Режим работы'), mode),
    h('div', { class: 'row wrap', style: { alignItems: 'flex-start' } }, h('div', { class: 'field grow' }, h('label', {}, 'Домен (для автоматического HTTPS)'), domain, h('div', { class: 'small muted' }, 'A-запись домена должна указывать на этот сервер, порты 80 и 443 открыты.')), h('div', { class: 'field grow' }, h('label', {}, 'E-mail для Let\'s Encrypt'), email)),
    h('div', { class: 'row wrap' }, h('div', { class: 'field grow' }, h('label', {}, 'Макс. размер файла, МБ'), upl), h('div', { class: 'field grow' }, h('label', {}, 'Квота новых пользователей, МБ (0 — без лимита)'), quota)),
    st.cert ? h('p', { class: 'small muted' }, `Сертификат: ${st.cert.subject.replace(/\n/g, ' ')} · ${st.cert.issuer} · действует до ${new Date(st.cert.validTo).toLocaleString('ru-RU')}`) : '',
    h('div', { class: 'row' }, h('button', { class: 'btn primary', onclick: async () => { try { const r = await api('PUT', '/api/admin/settings', { serverName: name.value, mode: mode.querySelector('input:checked').value, domain: domain.value.trim().toLowerCase(), email: email.value.trim(), maxUploadMB: +upl.value || 200, defaultQuotaMB: +quota.value || 0 }); toast(r.restartNeeded ? 'Сохранено. Перезапустите службу: sudo systemctl restart rabbit-office' : 'Сохранено', 'ok', 6000); render(); } catch (e) { toast(e.message, 'error'); } } }, 'Сохранить'),
      c.domain ? h('button', { class: 'btn', onclick: async () => { await api('PUT', '/api/admin/settings', { renewCert: true }).then(() => toast('Запрос сертификата отправлен, проверьте через минуту', 'ok')).catch((e) => toast(e.message, 'error')); } }, 'Обновить сертификат') : ''));
  app.append(sb);
  // boards
  const boards = await api('GET', '/api/admin/boards').catch(() => []);
  const bb = h('div', { class: 'box' }, h('h2', {}, 'Доски'));
  if (!boards.length) bb.append(h('div', { class: 'muted' }, 'Досок пока нет'));
  for (const b of boards) bb.append(h('div', { class: 'list-item', html: `<div class="grow"><div class="name"><b>${esc(b.title)}</b></div><div class="meta">владелец ${esc(b.owner)} · изменена ${fmtDate(b.updated)}${b.online ? ' · онлайн: ' + b.online : ''}</div></div>` }, h('button', { class: 'btn sm danger', onclick: async () => { if (await confirmDlg('Удалить доску', `Удалить «${b.title}»?`, 'Удалить', true)) { await api('DELETE', '/api/boards/' + b.id); render(); } } }, 'Удалить')));
  app.append(bb);
  // maintenance
  const mb = h('div', { class: 'box' }, h('h2', {}, 'Обслуживание'),
    h('p', { class: 'muted', style: { marginTop: 0 } }, 'Резервная копия содержит пользователей, файлы, доски и настройки. Восстановление: rabbit-office restore <файл>.'),
    h('button', { class: 'btn primary', onclick: async (e) => {
      const btn = e.currentTarget; btn.disabled = true; btn.textContent = 'Готовлю архив…';
      try {
        const r = await fetch('/api/admin/backup', { headers: { Authorization: 'Bearer ' + token } }); if (!r.ok) throw new Error('Ошибка ' + r.status);
        const name = /filename="([^"]+)"/.exec(r.headers.get('content-disposition') || '')?.[1] || 'rabbit-office-backup.tar.gz';
        const url = URL.createObjectURL(await r.blob()); const a = h('a', { href: url, download: name }); document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 5000);
        toast('Резервная копия скачана', 'ok');
      } catch (er) { toast(er.message, 'error'); }
      btn.disabled = false; btn.textContent = 'Скачать резервную копию';
    } }, 'Скачать резервную копию'));
  app.append(mb);
}
async function addUser() {
  const p = genPass();
  const r = await formDlg('Новый пользователь', [{ key: 'username', label: 'Логин', placeholder: 'ivan.petrov' }, { key: 'displayName', label: 'Имя (видно соавторам)', placeholder: 'Иван Петров' }, { key: 'password', label: 'Пароль', value: p, hint: 'Сгенерирован автоматически — скопируйте и передайте пользователю' }, { key: 'role', label: 'Роль', type: 'select', options: [{ value: 'user', label: 'Пользователь' }, { value: 'admin', label: 'Администратор' }] }, { key: 'quotaMB', label: 'Квота хранилища, МБ (0 — без лимита)', type: 'number', value: 0 }], 'Создать');
  if (!r) return;
  try { await api('POST', '/api/admin/users', r); await dialog({ title: 'Пользователь создан', body: `<p>Передайте пользователю данные для входа:</p><div class="card mono" style="user-select:all">Сервер: ${esc(location.origin)}<br>Логин: ${esc(r.username)}<br>Пароль: ${esc(r.password)}</div>` }); render(); }
  catch (e) { toast(e.message, 'error'); }
}
async function editUser(u) {
  const r = await formDlg('Пользователь ' + u.username, [{ key: 'displayName', label: 'Имя', value: u.displayName }, { key: 'role', label: 'Роль', type: 'select', value: u.role, options: [{ value: 'user', label: 'Пользователь' }, { value: 'admin', label: 'Администратор' }] }, { key: 'quotaMB', label: 'Квота, МБ', type: 'number', value: u.quotaMB || 0 }, { key: 'disabled', label: 'Отключить вход', type: 'checkbox', value: u.disabled }], 'Сохранить');
  if (!r) return; await api('PATCH', '/api/admin/users/' + encodeURIComponent(u.username), r).catch((e) => toast(e.message, 'error')); render();
}
function loginView() {
  const user = h('input', { class: 'input', placeholder: 'admin', autocomplete: 'username' }); const pass = h('input', { class: 'input', type: 'password', autocomplete: 'current-password' }); const er = h('div', { class: 'small', style: { color: 'var(--danger)', minHeight: '18px' } });
  const go = async () => { er.textContent = ''; try { const r = await fetch('/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: user.value.trim(), password: pass.value }) }); const d = await r.json(); if (!r.ok) throw new Error(d.error); if (d.user.role !== 'admin') throw new Error('Нужны права администратора'); token = d.token; sessionStorage.setItem('rabbit.admin.token', token); render(); } catch (e) { er.textContent = e.message; } };
  pass.onkeydown = (e) => { if (e.key === 'Enter') go(); };
  app.append(h('div', { class: 'box login' }, h('h1', { style: { marginBottom: '16px' } }, 'Вход администратора'), h('div', { class: 'field' }, h('label', {}, 'Логин'), user), h('div', { class: 'field' }, h('label', {}, 'Пароль'), pass), er, h('button', { class: 'btn primary', style: { width: '100%' }, onclick: go }, 'Войти')));
}
render();
