// Rabbit Office desktop shell (Windows, macOS, Linux)
'use strict';
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const http = require('http');
const dgram = require('dgram');

const res = (p) => { const local = path.join(__dirname, p); return fs.existsSync(local) ? local : path.join(__dirname, '..', p); };
const WEB = res('web');
const { accept } = require(path.join(res('shared'), 'ws.js'));
const { Room } = require(path.join(res('shared'), 'relay.js'));
const APP_BY_EXT = { rdoc: 'documents', docx: 'documents', md: 'documents', txt: 'documents', html: 'documents', rxel: 'xel', xlsx: 'xel', csv: 'xel', rpt: 'point', pptx: 'point', rdesk: 'desk' };
const isMac = process.platform === 'darwin';

function createWindow(page = 'index.html') {
  const [file, query] = page.split('?');
  const win = new BrowserWindow({
    width: 1360, height: 880, minWidth: 380, minHeight: 480, show: false, backgroundColor: nativeTheme.shouldUseDarkColors ? '#1b1b1c' : '#f3f3f3',
    titleBarStyle: 'hidden', titleBarOverlay: isMac ? undefined : { color: '#00000000', symbolColor: '#ffffff', height: 40 }, trafficLightPosition: { x: 14, y: 13 },
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, sandbox: false, spellcheck: true },
  });
  win.loadFile(path.join(WEB, file), query ? { search: query } : undefined);
  win.once('ready-to-show', () => win.show());
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('file:')) return { action: 'allow', overrideBrowserWindowOptions: { width: 1100, height: 700, titleBarStyle: 'default', autoHideMenuBar: true } };
    shell.openExternal(url); return { action: 'deny' };
  });
  // native context menu for editable fields: spelling suggestions + clipboard (apps with their own menus prevent this)
  win.webContents.on('context-menu', (e, p) => {
    if (!p.isEditable && !p.selectionText) return;
    const items = p.dictionarySuggestions.slice(0, 6).map((s) => ({ label: s, click: () => win.webContents.replaceMisspelling(s) }));
    if (p.misspelledWord) items.push({ label: 'Добавить в словарь', click: () => win.webContents.session.addWordToSpellCheckerDictionary(p.misspelledWord) }, { type: 'separator' });
    items.push({ role: 'cut', label: 'Вырезать', enabled: p.editFlags.canCut, visible: p.isEditable }, { role: 'copy', label: 'Копировать', enabled: p.editFlags.canCopy }, { role: 'paste', label: 'Вставить', enabled: p.editFlags.canPaste, visible: p.isEditable }, { type: 'separator' }, { role: 'selectAll', label: 'Выделить всё' });
    Menu.buildFromTemplate(items).popup({ window: win });
  });
  let forceClose = false;
  win.on('close', (e) => { if (forceClose || !win.dirty) return; e.preventDefault(); win.webContents.send('before-close'); });
  win.forceClose = () => { forceClose = true; win.close(); };
  return win;
}
function openPath(p) { const ext = path.extname(p).slice(1).toLowerCase(); const page = APP_BY_EXT[ext]; if (page) createWindow(`${page}.html?path=${encodeURIComponent(p)}`); }

const fromEvent = (e) => BrowserWindow.fromWebContents(e.sender);
ipcMain.handle('open-window', (e, url) => { createWindow(url); });
ipcMain.handle('open-file', async (e, { accept: exts }) => {
  const r = await dialog.showOpenDialog(fromEvent(e), { properties: ['openFile'], filters: [{ name: 'Документы', extensions: exts }, { name: 'Все файлы', extensions: ['*'] }] });
  if (r.canceled || !r.filePaths[0]) return null; const p = r.filePaths[0];
  app.addRecentDocument(p);
  return { path: p, name: path.basename(p), data: fs.readFileSync(p).toString('base64') };
});
ipcMain.handle('read-file', (e, p) => ({ path: p, name: path.basename(p), data: fs.readFileSync(p).toString('base64') }));
ipcMain.handle('write-file', (e, p, b64) => { fs.writeFileSync(p + '.tmp', Buffer.from(b64, 'base64')); fs.renameSync(p + '.tmp', p); app.addRecentDocument(p); return true; });
ipcMain.handle('save-dialog', async (e, { defaultName, filters }) => { const r = await dialog.showSaveDialog(fromEvent(e), { defaultPath: path.join(app.getPath('documents'), defaultName), filters }); return r.canceled ? null : r.filePath; });
ipcMain.handle('print-pdf', async (e, p, opts) => { const data = await e.sender.printToPDF({ printBackground: true, preferCSSPageSize: true, ...opts }); fs.writeFileSync(p, data); return true; });
ipcMain.on('set-titlebar', (e, o) => { try { fromEvent(e)?.setTitleBarOverlay?.(o); } catch { /* not supported */ } });
ipcMain.on('set-dirty', (e, d) => { const w = fromEvent(e); if (w) { w.dirty = d; w.setDocumentEdited?.(d); } });
ipcMain.on('close-ok', (e) => fromEvent(e)?.forceClose());

/* ---------- Rabbit Desk: direct (IP) / LAN host ---------- */
let host = null; let beacon = null; let listener = null;
const DISCOVERY_PORT = 47801;
// Real LAN addresses first (192.168/10/172.16-31 on physical adapters), virtual/VPN adapters last.
function addresses() {
  const rank = (name, ip) => {
    const virtual = /vethernet|wsl|docker|virtualbox|vmware|hyper-v|vpn|radmin|hamachi|zerotier|tailscale|utun|tun|tap|br-|veth/i.test(name);
    const priv = /^192\.168\./.test(ip) ? 0 : /^10\./.test(ip) ? 1 : /^172\.(1[6-9]|2\d|3[01])\./.test(ip) ? 2 : 3;
    return (virtual ? 10 : 0) + priv;
  };
  return Object.entries(os.networkInterfaces()).flatMap(([name, list]) => (list || []).filter((i) => i.family === 'IPv4' && !i.internal).map((i) => ({ ip: i.address, r: rank(name, i.address) })))
    .sort((a, b) => a.r - b.r).map((x) => x.ip);
}
const SESSION_ID = Math.random().toString(36).slice(2, 10);
ipcMain.handle('desk-host', async (e, { port, pin, lan, title, name }) => {
  if (host) await ipcStop();
  const room = new Room('local', { pin });
  const server = http.createServer((req, resp) => {
    if (req.url === '/rabbit/info') { resp.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' }); return resp.end(JSON.stringify({ product: 'rabbit-desk', title, host: name, needsPin: !!pin })); }
    resp.writeHead(404); resp.end();
  });
  server.on('upgrade', (req, sock) => { if (!req.url.startsWith('/desk')) return sock.destroy(); const ws = accept(req, sock); if (ws) room.join(ws); });
  let p = port || 47800;
  await new Promise((ok, fail) => { const tryListen = () => { server.once('error', (er) => { if (er.code === 'EADDRINUSE' && p < (port || 47800) + 20) { p++; tryListen(); } else fail(er); }); server.listen(p, '0.0.0.0', ok); }; tryListen(); });
  host = { server, port: p, pin, lan, room };
  if (lan) {
    beacon = dgram.createSocket({ type: 'udp4', reuseAddr: true });
    beacon.bind(() => { beacon.setBroadcast(true); const msg = () => Buffer.from(JSON.stringify({ product: 'rabbit-desk', sid: SESSION_ID, best: addresses()[0], title, host: name, port: p, needsPin: !!pin })); const tick = () => { for (const a of ['255.255.255.255', ...addresses().map((x) => x.replace(/\.\d+$/, '.255'))]) beacon.send(msg(), DISCOVERY_PORT, a, () => {}); }; tick(); beacon.timer = setInterval(tick, 2000); });
  }
  return { port: p, addresses: addresses() };
});
async function ipcStop() {
  if (beacon) { clearInterval(beacon.timer); beacon.close(); beacon = null; }
  if (host) { const h = host; host = null; h.room.broadcast({ t: 'error', message: 'Хост завершил сессию' }); for (const ws of h.room.clients.keys()) ws.close(); await new Promise((r) => h.server.close(r)); }
}
ipcMain.handle('desk-stop', ipcStop);
ipcMain.handle('desk-status', () => (host ? { hosting: true, port: host.port, pin: host.pin, lan: host.lan, addresses: addresses() } : { hosting: false }));
ipcMain.on('desk-discover', (e) => {
  const wc = e.sender;
  if (!listener) { listener = dgram.createSocket({ type: 'udp4', reuseAddr: true }); listener.bind(DISCOVERY_PORT); listener.subs = new Set(); listener.on('message', (m, rinfo) => { try { const d = JSON.parse(m); if (d.product === 'rabbit-desk') for (const s of listener.subs) if (!s.isDestroyed()) s.send('desk-found', { ...d, ip: rinfo.address }); } catch { /* ignore */ } }); }
  listener.subs.add(wc);
});

/* ---------- app lifecycle ---------- */
if (!app.requestSingleInstanceLock()) app.quit();
app.on('second-instance', (e, argv) => { const f = argv.find((a) => APP_BY_EXT[path.extname(a).slice(1).toLowerCase()] && fs.existsSync(a)); if (f) openPath(f); else createWindow(); });
app.on('open-file', (e, p) => { e.preventDefault(); if (app.isReady()) openPath(p); else app.once('ready', () => openPath(p)); });
app.whenReady().then(() => {
  if (!isMac) Menu.setApplicationMenu(null);
  try { require('electron').session.defaultSession.setSpellCheckerLanguages(['ru', 'en-US']); } catch { /* macOS uses the system checker */ }
  const f = process.argv.slice(1).find((a) => APP_BY_EXT[path.extname(a).slice(1).toLowerCase()] && fs.existsSync(a));
  if (f) openPath(f); else createWindow();
  app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
});
app.on('window-all-closed', () => { if (!isMac) app.quit(); });
