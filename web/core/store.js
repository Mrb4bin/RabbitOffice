// Tiny IndexedDB key-value store + settings + recent files
const DB_NAME = 'rabbit-office';
let dbp;
function db() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB_NAME, 1);
    r.onupgradeneeded = () => { const d = r.result; d.createObjectStore('kv'); d.createObjectStore('library', { keyPath: 'id' }); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function tx(store, mode, fn) {
  const d = await db();
  return new Promise((res, rej) => {
    const t = d.transaction(store, mode); const s = t.objectStore(store);
    const req = fn(s);
    t.oncomplete = () => res(req?.result); t.onerror = () => rej(t.error);
  });
}
export const kv = {
  get: (k) => tx('kv', 'readonly', (s) => s.get(k)),
  set: (k, v) => tx('kv', 'readwrite', (s) => s.put(v, k)),
  del: (k) => tx('kv', 'readwrite', (s) => s.delete(k)),
};
// Local document library (used on web/Android where there is no real file system)
export const library = {
  list: async () => (await tx('library', 'readonly', (s) => s.getAll())).map(({ data, ...m }) => ({ ...m, size: data?.length || data?.byteLength || 0 })).sort((a, b) => b.updated - a.updated),
  get: (id) => tx('library', 'readonly', (s) => s.get(id)),
  put: (doc) => tx('library', 'readwrite', (s) => s.put({ ...doc, updated: Date.now() })),
  remove: (id) => tx('library', 'readwrite', (s) => s.delete(id)),
};

/* ---------- settings ---------- */
const SKEY = 'rabbit.settings';
const defaults = {
  theme: 'standard',
  userName: '',
  userColor: '',
  autosave: true,
  ribbonCollapsed: false,
  spellcheck: true,
  deskPort: 47800,
};
function readAll() { try { return { ...defaults, ...JSON.parse(localStorage.getItem(SKEY) || '{}') }; } catch { return { ...defaults }; } }
export const settings = {
  get(k) { return readAll()[k]; },
  all: readAll,
  set(k, v) { const a = readAll(); a[k] = v; localStorage.setItem(SKEY, JSON.stringify(a)); window.dispatchEvent(new CustomEvent('rabbit-settings', { detail: { k, v } })); },
};
const COLORS = ['#e5484d', '#f76b15', '#ffb224', '#30a46c', '#12a594', '#0090ff', '#3e63dd', '#8e4ec6', '#d6409f'];
export function profile() {
  let name = settings.get('userName');
  let color = settings.get('userColor');
  if (!name) { name = 'Гость ' + Math.floor(100 + Math.random() * 900); settings.set('userName', name); }
  if (!color) { color = COLORS[Math.floor(Math.random() * COLORS.length)]; settings.set('userColor', color); }
  return { name, color };
}
export const USER_COLORS = COLORS;

/* ---------- recent files ---------- */
const RKEY = 'rabbit.recent';
export const recent = {
  list() { try { return JSON.parse(localStorage.getItem(RKEY) || '[]'); } catch { return []; } },
  add(entry) {
    const key = (e) => `${e.source}|${e.path || e.libId || e.cloudPath || e.boardId}|${e.account || ''}`;
    const l = recent.list().filter((e) => key(e) !== key(entry));
    l.unshift({ ...entry, time: Date.now() });
    localStorage.setItem(RKEY, JSON.stringify(l.slice(0, 40)));
  },
  remove(pred) { localStorage.setItem(RKEY, JSON.stringify(recent.list().filter((e) => !pred(e)))); },
  clear() { localStorage.removeItem(RKEY); },
};
