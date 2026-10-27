// Rabbit Office Server — self-hosted storage + collaborative boards + accounts + automatic HTTPS.
'use strict';
const http = require('http');
const https = require('https');
const fs = require('fs');
const fsp = fs.promises;
const path = require('path');
const crypto = require('crypto');
const { accept } = require('../../shared/ws.js');
const { Room } = require('../../shared/relay.js');
const store = require('./store.js');

const VERSION = '1.0.0';
const WEB_ROOT = path.resolve(__dirname, '../../web');
const ADMIN_ROOT = path.resolve(__dirname, '../admin');
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.json': 'application/json', '.ico': 'image/x-icon', '.woff2': 'font/woff2' };

let cfg = store.loadConfig();
const rooms = new Map();
const loginAttempts = new Map();
const acmeTokens = new Map();

/* ---------------- helpers ---------------- */
function send(res, code, body, headers = {}) {
  const isObj = body && typeof body === 'object' && !Buffer.isBuffer(body);
  res.writeHead(code, { 'Content-Type': isObj ? 'application/json; charset=utf-8' : 'text/plain; charset=utf-8', ...cors(), ...headers });
  res.end(isObj ? JSON.stringify(body) : body);
}
function cors() { return { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization, Content-Type', 'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS', 'Access-Control-Max-Age': '86400' }; }
const err = (res, code, message) => send(res, code, { error: message });
async function readBody(req, limit = cfg.maxUploadMB * 1024 * 1024) {
  const chunks = []; let size = 0;
  for await (const c of req) { size += c.length; if (size > limit) throw Object.assign(new Error('Файл слишком большой'), { status: 413 }); chunks.push(c); }
  return Buffer.concat(chunks);
}
async function json(req) { const b = await readBody(req, 2 * 1024 * 1024); try { return JSON.parse(b.toString('utf8') || '{}'); } catch { throw Object.assign(new Error('Некорректный JSON'), { status: 400 }); } }
function auth(req, url) {
  const h = req.headers.authorization || ''; const token = h.startsWith('Bearer ') ? h.slice(7) : url.searchParams.get('token');
  return token ? store.userByToken(token) : null;
}
const hasStorage = () => cfg.mode === 'storage' || cfg.mode === 'both';
const hasBoards = () => cfg.mode === 'board' || cfg.mode === 'both';

/* user file space: /home → users/<name>, /Общие → shared */
function resolveUserPath(user, p) {
  p = '/' + String(p || '/').replace(/\\/g, '/').split('/').filter((s) => s && s !== '.' && s !== '..').join('/');
  const sharedPrefix = '/Общие';
  let base, rel;
  if (p === sharedPrefix || p.startsWith(sharedPrefix + '/')) { base = path.join(store.dataDir(), 'files', '_shared'); rel = p.slice(sharedPrefix.length); }
  else { base = path.join(store.dataDir(), 'files', user.username); rel = p; }
  const full = path.join(base, rel);
  if (!full.startsWith(base)) throw Object.assign(new Error('Недопустимый путь'), { status: 400 });
  return { full, isRoot: p === '/', base };
}
async function dirSize(dir) { let total = 0; try { for (const e of await fsp.readdir(dir, { withFileTypes: true })) { const f = path.join(dir, e.name); total += e.isDirectory() ? await dirSize(f) : (await fsp.stat(f)).size; } } catch { /* empty */ } return total; }

/* ---------------- rooms ---------------- */
function getRoom(id) {
  let r = rooms.get(id); if (r) return r;
  const b = store.getBoard(id); if (!b) return null;
  r = new Room(id, { objects: b.objects || [], meta: { title: b.title }, onChange: () => scheduleSave(r) });
  rooms.set(id, r); return r;
}
const saveTimers = new Map();
function scheduleSave(r) { clearTimeout(saveTimers.get(r.id)); saveTimers.set(r.id, setTimeout(() => { store.saveBoardObjects(r.id, r.snapshot().objects); if (!r.online) rooms.delete(r.id); }, 1500)); }

/* ---------------- API ---------------- */
async function api(req, res, url) {
  const p = url.pathname; const m = req.method;
  if (p === '/api/info') return send(res, 200, { product: 'rabbit-office-server', version: VERSION, name: cfg.serverName, mode: cfg.mode, https: !!cfg.domain });
  if (p === '/api/health') return send(res, 200, { ok: true, uptime: process.uptime() });
  if (p === '/api/login' && m === 'POST') {
    const ip = req.socket.remoteAddress; const a = loginAttempts.get(ip) || { n: 0, t: Date.now() };
    if (Date.now() - a.t > 15 * 60e3) { a.n = 0; a.t = Date.now(); }
    if (a.n >= 10) return err(res, 429, 'Слишком много попыток. Подождите 15 минут.');
    const { username, password } = await json(req);
    const user = store.checkPassword(String(username || '').toLowerCase(), String(password || ''));
    if (!user) { a.n++; loginAttempts.set(ip, a); return err(res, 401, 'Неверный логин или пароль'); }
    if (user.disabled) return err(res, 403, 'Учётная запись отключена');
    loginAttempts.delete(ip);
    return send(res, 200, { token: store.createSession(user.username), user: store.publicUser(user) });
  }
  const user = auth(req, url);
  if (!user) return err(res, 401, 'Требуется вход');
  if (p === '/api/logout' && m === 'POST') { store.deleteSession((req.headers.authorization || '').slice(7)); return send(res, 200, { ok: true }); }
  if (p === '/api/me') return send(res, 200, store.publicUser(user));
  if (p === '/api/me/password' && m === 'POST') { const { oldPassword, newPassword } = await json(req); if (!store.checkPassword(user.username, oldPassword || '')) return err(res, 400, 'Старый пароль неверен'); if (String(newPassword || '').length < 6) return err(res, 400, 'Минимум 6 символов'); store.setPassword(user.username, newPassword); return send(res, 200, { ok: true }); }

  /* files */
  if (p.startsWith('/api/files')) {
    if (!hasStorage()) return err(res, 403, 'Хранилище на этом сервере отключено');
    if (p === '/api/files' && m === 'GET') {
      const { full, isRoot } = resolveUserPath(user, url.searchParams.get('path')); await fsp.mkdir(full, { recursive: true });
      const items = [];
      for (const e of await fsp.readdir(full, { withFileTypes: true })) { const st = await fsp.stat(path.join(full, e.name)); items.push({ name: e.name, dir: e.isDirectory(), size: st.size, mtime: st.mtimeMs }); }
      if (isRoot) items.unshift({ name: 'Общие', dir: true, size: 0, mtime: Date.now() });
      items.sort((a, b) => (b.dir - a.dir) || a.name.localeCompare(b.name, 'ru'));
      return send(res, 200, { items, used: await dirSize(path.join(store.dataDir(), 'files', user.username)), quota: user.quotaMB ? user.quotaMB * 1048576 : 0 });
    }
    if (p === '/api/files/download') { const { full } = resolveUserPath(user, url.searchParams.get('path')); try { const data = await fsp.readFile(full); res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Content-Length': data.length, ...cors() }); return res.end(data); } catch { return err(res, 404, 'Файл не найден'); } }
    if (p === '/api/files/upload' && m === 'PUT') {
      const { full, base } = resolveUserPath(user, url.searchParams.get('path'));
      const body = await readBody(req);
      if (user.quotaMB && !full.includes('_shared') && (await dirSize(base)) + body.length > user.quotaMB * 1048576) return err(res, 507, 'Превышена квота хранилища');
      await fsp.mkdir(path.dirname(full), { recursive: true }); await fsp.writeFile(full + '.tmp', body); await fsp.rename(full + '.tmp', full);
      return send(res, 200, { ok: true, size: body.length });
    }
    if (p === '/api/files/mkdir' && m === 'POST') { const { full } = resolveUserPath(user, (await json(req)).path); await fsp.mkdir(full, { recursive: true }); return send(res, 200, { ok: true }); }
    if (p === '/api/files/move' && m === 'POST') { const b = await json(req); const a = resolveUserPath(user, b.from).full, c = resolveUserPath(user, b.to).full; await fsp.mkdir(path.dirname(c), { recursive: true }); await fsp.rename(a, c); return send(res, 200, { ok: true }); }
    if (p === '/api/files' && m === 'DELETE') { const { full, isRoot } = resolveUserPath(user, url.searchParams.get('path')); if (isRoot) return err(res, 400, 'Нельзя удалить корень'); await fsp.rm(full, { recursive: true, force: true }); return send(res, 200, { ok: true }); }
  }

  /* boards */
  if (p.startsWith('/api/boards')) {
    if (!hasBoards()) return err(res, 403, 'Доски на этом сервере отключены');
    const id = decodeURIComponent(p.split('/')[3] || '');
    if (!id && m === 'GET') return send(res, 200, store.listBoards().map((b) => ({ ...b, online: rooms.get(b.id)?.online || 0 })));
    if (!id && m === 'POST') { const b = await json(req); return send(res, 200, store.createBoard(String(b.title || 'Доска').slice(0, 120), user.username)); }
    const board = store.getBoard(id); if (!board) return err(res, 404, 'Доска не найдена');
    if (m === 'PATCH') { const b = await json(req); store.renameBoard(id, String(b.title || board.title).slice(0, 120)); const r = rooms.get(id); if (r) { r.meta.title = b.title; r.broadcast({ t: 'meta', meta: r.meta }); } return send(res, 200, { ok: true }); }
    if (m === 'DELETE') { if (board.owner !== user.username && user.role !== 'admin') return err(res, 403, 'Удалить может владелец или администратор'); store.deleteBoard(id); rooms.get(id)?.broadcast({ t: 'error', message: 'Доска удалена' }); rooms.delete(id); return send(res, 200, { ok: true }); }
  }

  /* admin */
  if (p.startsWith('/api/admin')) {
    if (user.role !== 'admin') return err(res, 403, 'Нужны права администратора');
    if (p === '/api/admin/status') return send(res, 200, { version: VERSION, uptime: process.uptime(), users: store.listUsers().length, boards: store.listBoards().length, online: [...rooms.values()].reduce((a, r) => a + r.online, 0), storageBytes: await dirSize(path.join(store.dataDir(), 'files')), config: publicConfig(), cert: certInfo(), node: process.version, dataDir: store.dataDir() });
    if (p === '/api/admin/users' && m === 'GET') return send(res, 200, await Promise.all(store.listUsers().map(async (u) => ({ ...store.publicUser(u), used: await dirSize(path.join(store.dataDir(), 'files', u.username)) }))));
    if (p === '/api/admin/users' && m === 'POST') { const b = await json(req); try { const u = store.createUser(b); return send(res, 200, store.publicUser(u)); } catch (e) { return err(res, 400, e.message); } }
    const um = /^\/api\/admin\/users\/([^/]+)$/.exec(p);
    if (um && m === 'PATCH') { const b = await json(req); try { const u = store.updateUser(decodeURIComponent(um[1]), b); return send(res, 200, store.publicUser(u)); } catch (e) { return err(res, 400, e.message); } }
    if (um && m === 'DELETE') { const name = decodeURIComponent(um[1]); if (name === user.username) return err(res, 400, 'Нельзя удалить себя'); store.deleteUser(name); return send(res, 200, { ok: true }); }
    if (p === '/api/admin/settings' && m === 'PUT') {
      const b = await json(req); const before = { ...cfg };
      for (const k of ['serverName', 'mode', 'domain', 'email', 'maxUploadMB', 'defaultQuotaMB']) if (b[k] !== undefined) cfg[k] = b[k];
      if (!['storage', 'board', 'both'].includes(cfg.mode)) cfg.mode = 'both';
      store.saveConfig(cfg);
      if (cfg.domain && (cfg.domain !== before.domain || b.renewCert)) ensureCert(true).catch((e) => console.error('[acme]', e.message));
      return send(res, 200, { ok: true, restartNeeded: cfg.domain !== before.domain && !before.domain });
    }
    if (p === '/api/admin/boards') return send(res, 200, store.listBoards().map((b) => ({ ...b, online: rooms.get(b.id)?.online || 0 })));
    if (p === '/api/admin/backup') {
      // stream a tar.gz of the whole data directory (users, files, boards, config; certificates excluded)
      for (const r of rooms.values()) if (r.dirty) store.saveBoardObjects(r.id, r.snapshot().objects);
      const { spawn } = require('child_process');
      const name = `rabbit-office-backup-${new Date().toISOString().slice(0, 10)}.tar.gz`;
      const tar = spawn('tar', ['-czf', '-', '--exclude=./certs', '-C', store.dataDir(), '.']);
      res.writeHead(200, { 'Content-Type': 'application/gzip', 'Content-Disposition': `attachment; filename="${name}"`, ...cors() });
      tar.stdout.pipe(res); tar.on('error', () => res.end());
      return;
    }
  }
  return err(res, 404, 'Не найдено');
}
const publicConfig = () => ({ serverName: cfg.serverName, mode: cfg.mode, domain: cfg.domain, email: cfg.email, port: cfg.port, maxUploadMB: cfg.maxUploadMB, defaultQuotaMB: cfg.defaultQuotaMB });

/* ---------------- static ---------------- */
async function serveStatic(req, res, url) {
  let p = decodeURIComponent(url.pathname);
  let root = WEB_ROOT;
  if (p === '/admin') { res.writeHead(302, { Location: '/admin/' }); return res.end(); }
  if (p.startsWith('/admin/')) { root = ADMIN_ROOT; p = p.slice(6); }
  if (p.endsWith('/')) p += 'index.html';
  const file = path.join(root, path.normalize(p));
  if (!file.startsWith(root)) return err(res, 403, 'Forbidden');
  try {
    const data = await fsp.readFile(file);
    res.writeHead(200, { 'Content-Type': MIME[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-cache', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'same-origin' });
    res.end(data);
  } catch { err(res, 404, 'Не найдено'); }
}

async function handler(req, res) {
  const url = new URL(req.url, 'http://x');
  if (req.method === 'OPTIONS') { res.writeHead(204, cors()); return res.end(); }
  try {
    if (url.pathname.startsWith('/api/')) return await api(req, res, url);
    return await serveStatic(req, res, url);
  } catch (e) { console.error(e); return err(res, e.status || 500, e.status ? e.message : 'Внутренняя ошибка сервера'); }
}
function upgrade(req, sock) {
  const url = new URL(req.url, 'http://x'); const m = /^\/ws\/board\/([^/]+)$/.exec(url.pathname);
  if (!m || !hasBoards()) return sock.destroy();
  const user = auth(req, url); if (!user) { sock.end('HTTP/1.1 401 Unauthorized\r\n\r\n'); return; }
  const room = getRoom(decodeURIComponent(m[1])); if (!room) { sock.end('HTTP/1.1 404 Not Found\r\n\r\n'); return; }
  const ws = accept(req, sock); if (!ws) return;
  room.join(ws, user);
  ws.on('close', () => { if (!room.online) scheduleSave(room); });
}

/* ---------------- TLS / ACME (Let's Encrypt) ---------------- */
const certDir = () => path.join(store.dataDir(), 'certs');
function certInfo() {
  try { const pem = fs.readFileSync(path.join(certDir(), 'cert.pem'), 'utf8'); const x = new crypto.X509Certificate(pem); return { subject: x.subject, validTo: x.validTo, issuer: x.issuer.split('\n').find((l) => l.startsWith('O=')) || x.issuer }; } catch { return null; }
}
async function ensureCert(force = false) {
  if (!cfg.domain) return false;
  const info = certInfo();
  if (!force && info && new Date(info.validTo) - Date.now() > 30 * 86400e3 && info.subject.includes(cfg.domain)) return false;
  let acme; try { acme = require('acme-client'); } catch { console.error('[acme] пакет acme-client не установлен: выполните npm install в папке server'); return false; }
  console.log(`[acme] запрашиваю сертификат для ${cfg.domain}…`);
  await fsp.mkdir(certDir(), { recursive: true });
  const accKeyFile = path.join(certDir(), 'account.key');
  let accountKey; try { accountKey = await fsp.readFile(accKeyFile); } catch { accountKey = await acme.crypto.createPrivateKey(); await fsp.writeFile(accKeyFile, accountKey, { mode: 0o600 }); }
  const client = new acme.Client({ directoryUrl: cfg.acmeStaging ? acme.directory.letsencrypt.staging : acme.directory.letsencrypt.production, accountKey });
  const [key, csr] = await acme.crypto.createCsr({ commonName: cfg.domain });
  const cert = await client.auto({ csr, email: cfg.email || undefined, termsOfServiceAgreed: true, challengePriority: ['http-01'],
    challengeCreateFn: async (authz, ch, keyAuth) => { acmeTokens.set(ch.token, keyAuth); },
    challengeRemoveFn: async (authz, ch) => { acmeTokens.delete(ch.token); } });
  await fsp.writeFile(path.join(certDir(), 'key.pem'), key, { mode: 0o600 }); await fsp.writeFile(path.join(certDir(), 'cert.pem'), cert);
  console.log('[acme] сертификат получен');
  if (httpsServer) httpsServer.setSecureContext({ key, cert });
  return true;
}

/* ---------------- start ---------------- */
let httpsServer = null;
async function start() {
  store.init();
  if (!store.listUsers().length) {
    const pass = crypto.randomBytes(6).toString('base64url');
    store.createUser({ username: 'admin', password: pass, role: 'admin', displayName: 'Администратор' });
    console.log(`\n  Создан администратор: логин admin, пароль ${pass}\n  Смените пароль в админ-панели!\n`);
  }
  if (cfg.domain && !cfg.behindProxy) {
    // :80 serves ACME challenges and redirects to HTTPS
    http.createServer((req, res) => {
      const m = /^\/\.well-known\/acme-challenge\/(.+)$/.exec(req.url);
      if (m && acmeTokens.has(m[1])) { res.writeHead(200); return res.end(acmeTokens.get(m[1])); }
      res.writeHead(301, { Location: `https://${cfg.domain}${req.url}` }); res.end();
    }).listen(80, () => console.log('HTTP :80 → HTTPS'));
    try { await ensureCert(); } catch (e) { console.error('[acme] ошибка:', e.message); }
    let key, cert;
    try { key = fs.readFileSync(path.join(certDir(), 'key.pem')); cert = fs.readFileSync(path.join(certDir(), 'cert.pem')); }
    catch { const sc = selfSigned(); key = sc.key; cert = sc.cert; console.warn('[tls] временный самоподписанный сертификат'); }
    httpsServer = https.createServer({ key, cert }, handler);
    httpsServer.on('upgrade', upgrade);
    httpsServer.listen(443, () => console.log(`Rabbit Office Server ${VERSION}: https://${cfg.domain}  (админ-панель: https://${cfg.domain}/admin)`));
    setInterval(() => ensureCert().catch((e) => console.error('[acme]', e.message)), 12 * 3600e3);
  } else {
    const srv = http.createServer(handler); srv.on('upgrade', upgrade);
    srv.listen(cfg.port, cfg.host || '0.0.0.0', () => console.log(`Rabbit Office Server ${VERSION}: http://0.0.0.0:${cfg.port}  (админ-панель: /admin)`));
  }
  const flush = () => { for (const r of rooms.values()) if (r.dirty) store.saveBoardObjects(r.id, r.snapshot().objects); process.exit(0); };
  process.on('SIGINT', flush); process.on('SIGTERM', flush);
}
function selfSigned() {
  // Generate a throwaway self-signed certificate via openssl if available.
  const { execFileSync } = require('child_process'); const dir = certDir(); fs.mkdirSync(dir, { recursive: true });
  const k = path.join(dir, 'tmp-key.pem'), c = path.join(dir, 'tmp-cert.pem');
  try { execFileSync('openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', k, '-out', c, '-days', '7', '-subj', `/CN=${cfg.domain}`], { stdio: 'ignore' }); } catch { /* openssl missing */ }
  return { key: fs.readFileSync(k), cert: fs.readFileSync(c) };
}
if (require.main === module) start();
module.exports = { start, ensureCert };
