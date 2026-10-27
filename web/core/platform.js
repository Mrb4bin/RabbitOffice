// Platform abstraction: Electron (window.rabbitNative), Android (Capacitor) or plain web.
import { library, recent } from './store.js';
import { accounts, files as cloudFiles } from './cloud.js';
import { pickFile, readFileAs, uid, h } from './ui.js';
import { bytesToB64, b64ToBytes } from './zip.js';

export const native = window.rabbitNative || null;
const cap = window.Capacitor;
export const isAndroid = !!(cap && cap.isNativePlatform && cap.isNativePlatform() && cap.getPlatform() === 'android');
export const platform = native ? native.platform : isAndroid ? 'android' : 'web';
document.documentElement.classList.add('platform-' + (platform === 'darwin' ? 'mac' : platform));
if (native) document.documentElement.classList.add('is-desktop');

export const APPS = {
  docs: { id: 'docs', name: 'Rabbit Documents', short: 'Documents', page: 'documents.html', ext: 'rdoc', color: 'var(--c-docs)', hex: '#2f62c9', letter: 'D',
    open: ['rdoc', 'docx', 'html', 'htm', 'txt', 'md'] },
  xel: { id: 'xel', name: 'Rabbit Xel', short: 'Xel', page: 'xel.html', ext: 'rxel', color: 'var(--c-xel)', hex: '#1d8a52', letter: 'X',
    open: ['rxel', 'xlsx', 'csv', 'tsv'] },
  point: { id: 'point', name: 'Rabbit Point', short: 'Point', page: 'point.html', ext: 'rpt', color: 'var(--c-point)', hex: '#d2552d', letter: 'P',
    open: ['rpt', 'pptx'] },
  desk: { id: 'desk', name: 'Rabbit Desk', short: 'Desk', page: 'desk.html', ext: 'rdesk', color: 'var(--c-desk)', hex: '#7b4fd6', letter: 'B',
    open: ['rdesk'] },
};
export function extOf(name) { const m = /\.([a-z0-9]+)$/i.exec(name || ''); return m ? m[1].toLowerCase() : ''; }
export function appForFile(name) { const e = extOf(name); return Object.values(APPS).find((a) => a.open.includes(e)) || null; }
export function baseName(name) { return String(name || '').replace(/\.[a-z0-9]+$/i, ''); }

const toBytes = (data) => (typeof data === 'string' ? new TextEncoder().encode(data) : data instanceof Uint8Array ? data : new Uint8Array(data));

/* ---------- open / navigate between apps ---------- */
export function openApp(appId, params = {}) {
  const a = APPS[appId]; const qs = new URLSearchParams(params).toString();
  const url = a.page + (qs ? '?' + qs : '');
  if (native) native.openWindow(url); else location.href = url;
}
export function goHome() { if (native) native.openWindow('index.html'); else location.href = 'index.html'; }

// Hand a loaded file to an app page (works across windows via IndexedDB).
import { kv } from './store.js';
export async function handoff(appId, file) {
  const id = uid('h');
  await kv.set('handoff:' + id, file);
  openApp(appId, { handoff: id });
}
export async function takeHandoff(id) { const f = await kv.get('handoff:' + id); if (f) kv.del('handoff:' + id); return f; }

/* ---------- reading files ----------
   A "file" is { name, bytes:Uint8Array, loc } where loc describes where to save back:
   {source:'native', path} | {source:'library', libId} | {source:'cloud', account, path} | null */
export async function openLocalFile(accept) {
  if (native) {
    const r = await native.openFile({ accept });
    if (!r) return null;
    return { name: r.name, bytes: b64ToBytes(r.data), loc: { source: 'native', path: r.path } };
  }
  const f = await pickFile(accept.map((e) => '.' + e).join(','));
  if (!f) return null;
  return { name: f.name, bytes: new Uint8Array(await readFileAs(f, 'buffer')), loc: null };
}
export async function readFromLoc(loc, name) {
  if (loc.source === 'native') { const r = await native.readFile(loc.path); return { name: r.name, bytes: b64ToBytes(r.data), loc }; }
  if (loc.source === 'library') { const d = await library.get(loc.libId); if (!d) throw new Error('Документ не найден в библиотеке'); return { name: d.name, bytes: toBytes(d.data), loc }; }
  if (loc.source === 'cloud') { const acc = accounts.get(loc.account); if (!acc) throw new Error('Нет входа на сервер'); return { name: name || loc.path.split('/').pop(), bytes: await cloudFiles.download(acc, loc.path), loc }; }
  throw new Error('Неизвестный источник');
}

/* ---------- saving ---------- */
export async function writeToLoc(loc, name, data, appId) {
  const bytes = toBytes(data);
  if (loc.source === 'native') { await native.writeFile(loc.path, bytesToB64(bytes)); }
  else if (loc.source === 'library') { await library.put({ id: loc.libId, name, app: appId, data: bytes }); }
  else if (loc.source === 'cloud') { const acc = accounts.get(loc.account); if (!acc) throw new Error('Нет входа на сервер'); await cloudFiles.upload(acc, loc.path, bytes); }
  recent.add({ app: appId, name, ...locKey(loc) });
  return loc;
}
export function locKey(loc) {
  if (!loc) return {};
  if (loc.source === 'native') return { source: 'native', path: loc.path };
  if (loc.source === 'library') return { source: 'library', libId: loc.libId };
  if (loc.source === 'cloud') return { source: 'cloud', account: loc.account, cloudPath: loc.path };
  return {};
}
export function locFromRecent(e) {
  if (e.source === 'native') return { source: 'native', path: e.path };
  if (e.source === 'library') return { source: 'library', libId: e.libId };
  if (e.source === 'cloud') return { source: 'cloud', account: e.account, path: e.cloudPath };
  return null;
}
export function locLabel(loc) {
  if (!loc) return 'Не сохранено';
  if (loc.source === 'native') return loc.path;
  if (loc.source === 'library') return 'Мои документы (на устройстве)';
  if (loc.source === 'cloud') { const a = accounts.get(loc.account); return (a ? a.serverName + ' · ' : 'Сервер · ') + loc.path; }
  return '';
}

// Ask native dialog for a path (desktop)
export async function nativeSavePath(defaultName, filters) {
  if (!native) return null;
  return native.saveDialog({ defaultName, filters });
}

/* Export/download a file somewhere outside the app (always asks / downloads) */
export async function exportFile(name, data, mime = 'application/octet-stream') {
  const bytes = toBytes(data);
  if (native) {
    const ext = extOf(name);
    const path = await native.saveDialog({ defaultName: name, filters: [{ name: ext.toUpperCase(), extensions: [ext] }] });
    if (!path) return null;
    await native.writeFile(path, bytesToB64(bytes));
    return path;
  }
  if (isAndroid && cap.Plugins?.Filesystem) {
    const { Filesystem, Share } = cap.Plugins;
    const r = await Filesystem.writeFile({ path: 'Rabbit Office/' + name, data: bytesToB64(bytes), directory: 'DOCUMENTS', recursive: true });
    try { await Share?.share({ title: name, url: r.uri }); } catch { /* user closed share sheet */ }
    return r.uri;
  }
  if (window.showSaveFilePicker) {
    try {
      const hnd = await window.showSaveFilePicker({ suggestedName: name });
      const w = await hnd.createWritable(); await w.write(bytes); await w.close(); return name;
    } catch (e) { if (e.name === 'AbortError') return null; }
  }
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const a = h('a', { href: url, download: name }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  return name;
}

export async function printToPdf(defaultName, opts = {}) {
  if (native?.printToPDF) {
    const path = await native.saveDialog({ defaultName, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
    if (!path) return null;
    await native.printToPDF(path, opts);
    return path;
  }
  window.print();
  return null;
}

export function setTitleBarColors(bg, fg) { native?.setTitleBar?.({ color: bg, symbolColor: fg }); }
