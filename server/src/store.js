// Persistent storage for Rabbit Office Server: config, users, sessions, boards (JSON files on disk).
'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const DATA = process.env.RABBIT_DATA || (fs.existsSync('/var/lib/rabbit-office') ? '/var/lib/rabbit-office' : path.resolve(__dirname, '../data'));
const dataDir = () => DATA;
const file = (n) => path.join(DATA, n);
const DEFAULT_CFG = { serverName: 'Rabbit Office', mode: 'both', port: 8080, host: '0.0.0.0', domain: '', email: '', behindProxy: false, maxUploadMB: 200, defaultQuotaMB: 0, acmeStaging: false };

function readJson(n, def) { try { return JSON.parse(fs.readFileSync(file(n), 'utf8')); } catch { return def; } }
function writeJson(n, v) { fs.mkdirSync(DATA, { recursive: true }); const f = file(n); fs.writeFileSync(f + '.tmp', JSON.stringify(v, null, 2), { mode: 0o600 }); fs.renameSync(f + '.tmp', f); }
function init() { for (const d of ['', 'files', 'files/_shared', 'boards', 'certs']) fs.mkdirSync(path.join(DATA, d), { recursive: true }); }

function loadConfig() { return { ...DEFAULT_CFG, ...readJson('config.json', {}) }; }
function saveConfig(c) { writeJson('config.json', c); }

/* users */
let users = null;
const loadUsers = () => (users ??= readJson('users.json', []));
const saveUsers = () => writeJson('users.json', users);
function hash(password, salt = crypto.randomBytes(16).toString('hex')) { return { salt, hash: crypto.scryptSync(password, salt, 64).toString('hex') }; }
function checkPassword(username, password) {
  const u = loadUsers().find((x) => x.username === username); if (!u) { hash(password); return null; }
  const h = crypto.scryptSync(password, u.salt, 64);
  return crypto.timingSafeEqual(h, Buffer.from(u.hash, 'hex')) ? u : null;
}
function validName(n) { if (!/^[a-z0-9._-]{2,32}$/.test(n)) throw new Error('Логин: 2–32 символа, латиница, цифры, точка, дефис, подчёркивание'); }
function createUser({ username, password, role = 'user', displayName = '', quotaMB }) {
  username = String(username || '').toLowerCase().trim(); validName(username);
  if (String(password || '').length < 6) throw new Error('Пароль — минимум 6 символов');
  if (loadUsers().some((u) => u.username === username)) throw new Error('Такой пользователь уже существует');
  const u = { username, displayName: displayName || username, role: role === 'admin' ? 'admin' : 'user', quotaMB: quotaMB ?? loadConfig().defaultQuotaMB, disabled: false, created: Date.now(), ...hash(password) };
  users.push(u); saveUsers(); fs.mkdirSync(path.join(DATA, 'files', username), { recursive: true }); return u;
}
function updateUser(username, patch) {
  const u = loadUsers().find((x) => x.username === username); if (!u) throw new Error('Пользователь не найден');
  if (patch.displayName != null) u.displayName = String(patch.displayName).slice(0, 60);
  if (patch.role) { if (u.role === 'admin' && patch.role !== 'admin' && users.filter((x) => x.role === 'admin').length === 1) throw new Error('Должен остаться хотя бы один администратор'); u.role = patch.role === 'admin' ? 'admin' : 'user'; }
  if (patch.quotaMB != null) u.quotaMB = Math.max(0, +patch.quotaMB || 0);
  if (patch.disabled != null) { u.disabled = !!patch.disabled; if (u.disabled) dropSessions(username); }
  if (patch.password) { if (String(patch.password).length < 6) throw new Error('Пароль — минимум 6 символов'); Object.assign(u, hash(patch.password)); dropSessions(username); }
  saveUsers(); return u;
}
function setPassword(username, p) { return updateUser(username, { password: p }); }
function deleteUser(username) { loadUsers(); users = users.filter((u) => u.username !== username); saveUsers(); dropSessions(username); }
const listUsers = () => loadUsers();
const publicUser = (u) => ({ username: u.username, displayName: u.displayName, role: u.role, quotaMB: u.quotaMB, disabled: u.disabled, created: u.created });

/* sessions */
let sessions = null;
const loadSessions = () => (sessions ??= readJson('sessions.json', {}));
const tokenKey = (t) => crypto.createHash('sha256').update(t).digest('hex');
function createSession(username) { const t = crypto.randomBytes(32).toString('base64url'); loadSessions()[tokenKey(t)] = { username, exp: Date.now() + 30 * 86400e3 }; writeJson('sessions.json', sessions); return t; }
function userByToken(t) { const s = loadSessions()[tokenKey(String(t))]; if (!s || s.exp < Date.now()) return null; const u = loadUsers().find((x) => x.username === s.username); return u && !u.disabled ? u : null; }
function deleteSession(t) { delete loadSessions()[tokenKey(String(t))]; writeJson('sessions.json', sessions); }
function dropSessions(username) { const s = loadSessions(); for (const k of Object.keys(s)) if (s[k].username === username) delete s[k]; writeJson('sessions.json', s); }

/* boards */
const boardsIndex = () => readJson('boards.json', []);
const listBoards = () => boardsIndex();
function createBoard(title, owner) { const b = { id: crypto.randomBytes(8).toString('hex'), title, owner, created: Date.now(), updated: Date.now() }; const idx = boardsIndex(); idx.unshift(b); writeJson('boards.json', idx); writeJson(`boards/${b.id}.json`, { objects: [] }); return b; }
function getBoard(id) { if (!/^[a-f0-9]{16}$/.test(id)) return null; const meta = boardsIndex().find((b) => b.id === id); if (!meta) return null; return { ...meta, ...readJson(`boards/${id}.json`, { objects: [] }) }; }
function saveBoardObjects(id, objects) { writeJson(`boards/${id}.json`, { objects }); const idx = boardsIndex(); const b = idx.find((x) => x.id === id); if (b) { b.updated = Date.now(); writeJson('boards.json', idx); } }
function renameBoard(id, title) { const idx = boardsIndex(); const b = idx.find((x) => x.id === id); if (b) { b.title = title; writeJson('boards.json', idx); } }
function deleteBoard(id) { writeJson('boards.json', boardsIndex().filter((b) => b.id !== id)); try { fs.unlinkSync(file(`boards/${id}.json`)); } catch { /* gone */ } }

module.exports = { dataDir, init, loadConfig, saveConfig, createUser, updateUser, deleteUser, listUsers, publicUser, checkPassword, setPassword, createSession, userByToken, deleteSession, listBoards, createBoard, getBoard, saveBoardObjects, renameBoard, deleteBoard };
