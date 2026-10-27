// Client for a self-hosted Rabbit Office Server (storage + boards + accounts)
const AKEY = 'rabbit.accounts';
const ACTIVE = 'rabbit.activeAccount';

function normUrl(u) {
  u = String(u || '').trim().replace(/\/+$/, '');
  if (!u) return u;
  if (!/^https?:\/\//i.test(u)) u = (/^(localhost|\d+\.\d+\.\d+\.\d+)(:\d+)?$/.test(u) ? 'http://' : 'https://') + u;
  return u;
}

export const accounts = {
  list() { try { return JSON.parse(localStorage.getItem(AKEY) || '[]'); } catch { return []; } },
  save(l) { localStorage.setItem(AKEY, JSON.stringify(l)); window.dispatchEvent(new CustomEvent('rabbit-accounts')); },
  active() { const id = localStorage.getItem(ACTIVE); const l = accounts.list(); return l.find((a) => a.id === id) || l[0] || null; },
  setActive(id) { localStorage.setItem(ACTIVE, id); window.dispatchEvent(new CustomEvent('rabbit-accounts')); },
  get(id) { return accounts.list().find((a) => a.id === id) || null; },
  remove(id) { accounts.save(accounts.list().filter((a) => a.id !== id)); if (localStorage.getItem(ACTIVE) === id) localStorage.removeItem(ACTIVE); },
};

export class ApiError extends Error { constructor(msg, status) { super(msg); this.status = status; } }

export async function api(acc, method, path, body, { raw = false, binary = false, timeout = 30000 } = {}) {
  if (!acc) throw new ApiError('Нет подключения к серверу', 0);
  const headers = {};
  if (acc.token) headers.Authorization = 'Bearer ' + acc.token;
  let payload;
  if (body instanceof Uint8Array || body instanceof ArrayBuffer || body instanceof Blob) { payload = body; headers['Content-Type'] = 'application/octet-stream'; }
  else if (typeof body === 'string' && raw) { payload = body; headers['Content-Type'] = 'text/plain; charset=utf-8'; }
  else if (body !== undefined) { payload = JSON.stringify(body); headers['Content-Type'] = 'application/json'; }
  const ctl = new AbortController(); const t = setTimeout(() => ctl.abort(), timeout);
  let res;
  try { res = await fetch(acc.url + path, { method, headers, body: payload, signal: ctl.signal }); }
  catch (e) { throw new ApiError(e.name === 'AbortError' ? 'Сервер не отвечает' : 'Не удалось подключиться к серверу', 0); }
  finally { clearTimeout(t); }
  if (res.status === 401 && acc.token) { window.dispatchEvent(new CustomEvent('rabbit-auth-expired', { detail: acc })); }
  if (!res.ok) {
    let msg = res.statusText; try { msg = (await res.json()).error || msg; } catch { /* ignore */ }
    throw new ApiError(msg, res.status);
  }
  if (binary) return new Uint8Array(await res.arrayBuffer());
  const ct = res.headers.get('content-type') || '';
  return ct.includes('json') ? res.json() : res.text();
}

export async function serverInfo(url) {
  return api({ url: normUrl(url) }, 'GET', '/api/info', undefined, { timeout: 8000 });
}

export async function login(url, username, password) {
  url = normUrl(url);
  const r = await api({ url }, 'POST', '/api/login', { username, password });
  const info = await api({ url, token: r.token }, 'GET', '/api/info').catch(() => ({}));
  const acc = { id: `${url}|${r.user.username}`, url, token: r.token, username: r.user.username, displayName: r.user.displayName || r.user.username, role: r.user.role, serverName: info.name || url, mode: info.mode || 'both' };
  const l = accounts.list().filter((a) => a.id !== acc.id); l.unshift(acc); accounts.save(l); accounts.setActive(acc.id);
  return acc;
}
export const changePassword = (acc, oldPassword, newPassword) => api(acc, 'POST', '/api/me/password', { oldPassword, newPassword });
export async function logout(acc) { try { await api(acc, 'POST', '/api/logout'); } catch { /* ignore */ } accounts.remove(acc.id); }

export const files = {
  list: (acc, path = '/') => api(acc, 'GET', '/api/files?path=' + encodeURIComponent(path)),
  download: (acc, path) => api(acc, 'GET', '/api/files/download?path=' + encodeURIComponent(path), undefined, { binary: true }),
  upload: (acc, path, data) => api(acc, 'PUT', '/api/files/upload?path=' + encodeURIComponent(path), typeof data === 'string' ? new TextEncoder().encode(data) : data, { timeout: 120000 }),
  mkdir: (acc, path) => api(acc, 'POST', '/api/files/mkdir', { path }),
  remove: (acc, path) => api(acc, 'DELETE', '/api/files?path=' + encodeURIComponent(path)),
  move: (acc, from, to) => api(acc, 'POST', '/api/files/move', { from, to }),
};
export const boards = {
  list: (acc) => api(acc, 'GET', '/api/boards'),
  create: (acc, title) => api(acc, 'POST', '/api/boards', { title }),
  remove: (acc, id) => api(acc, 'DELETE', '/api/boards/' + encodeURIComponent(id)),
  rename: (acc, id, title) => api(acc, 'PATCH', '/api/boards/' + encodeURIComponent(id), { title }),
  wsUrl: (acc, id) => acc.url.replace(/^http/, 'ws') + '/ws/board/' + encodeURIComponent(id) + '?token=' + encodeURIComponent(acc.token),
};

// When the web office is served by a Rabbit server itself, offer that server automatically.
export async function sameOriginServer() {
  if (!/^https?:$/.test(location.protocol) || location.hostname === 'localhost' && location.port === '') return null;
  try { const info = await serverInfo(location.origin); return info && info.product === 'rabbit-office-server' ? { url: location.origin, ...info } : null; } catch { return null; }
}
export { normUrl };
