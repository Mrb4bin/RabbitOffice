// Rabbit Xel formula engine: tokenizer, parser, evaluator, reference shifting.
export const MAX_ROWS = 10000;
export const MAX_COLS = 260;

export class FErr { constructor(e) { this.e = e; } toString() { return this.e; } }
export const ERR = { div0: new FErr('#DIV/0!'), value: new FErr('#VALUE!'), ref: new FErr('#REF!'), name: new FErr('#NAME?'), na: new FErr('#N/A'), num: new FErr('#NUM!'), circ: new FErr('#CIRC!'), null: new FErr('#NULL!') };
const isErr = (v) => v instanceof FErr;

export function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
export function colIndex(s) { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return c - 1; }
export function refName(r, c) { return colName(c) + (r + 1); }
export function parseRef(s) { const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim()); if (!m) return null; return { r: +m[2] - 1, c: colIndex(m[1]) }; }
export function parseRange(s) {
  const p = s.trim().split(':'); const a = parseRef(p[0]); if (!a) return null; const b = p[1] ? parseRef(p[1]) : a; if (!b) return null;
  return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
}
export function rangeName(g) { return g.r1 === g.r2 && g.c1 === g.c2 ? refName(g.r1, g.c1) : refName(g.r1, g.c1) + ':' + refName(g.r2, g.c2); }

/* ---------------- dates (Excel serials) ---------------- */
const EPOCH = Date.UTC(1899, 11, 30);
export const dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;
export const serialToDate = (s) => new Date(EPOCH + Math.round(s * 86400000));
export function parseDateInput(s) {
  let m = /^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(s);
  if (m) { let y = +m[3]; if (y < 100) y += 2000; const ser = dateToSerial(y, +m[2], +m[1]); return { v: ser + (m[4] ? (+m[4] * 3600 + +m[5] * 60 + +(m[6] || 0)) / 86400 : 0), fmt: m[4] ? 'datetime' : 'date' }; }
  m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s); if (m) return { v: dateToSerial(+m[1], +m[2], +m[3]), fmt: 'date' };
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(s); if (m) return { v: (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, fmt: 'time' };
  return null;
}

/* ---------------- tokenizer ---------------- */
const RE_REF = /^(?:(?:'((?:[^']|'')+)'|([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*))!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![\w(])/;
export function tokenize(src) {
  const t = []; let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    const rest = src.slice(i);
    let m;
    if ((m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(rest))) { t.push({ type: 'num', v: parseFloat(m[0]), pos: i, len: m[0].length }); i += m[0].length; continue; }
    if (ch === '"') { let j = i + 1, s = ''; while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } t.push({ type: 'str', v: s, pos: i, len: j + 1 - i }); i = j + 1; continue; }
    if ((m = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)/i.exec(rest))) { t.push({ type: 'err', v: new FErr(m[0].toUpperCase()), pos: i, len: m[0].length }); i += m[0].length; continue; }
    if ((m = RE_REF.exec(rest))) {
      const prev = t[t.length - 1];
      if (!(prev && prev.type === 'func')) { t.push({ type: 'ref', sheet: m[1] ? m[1].replace(/''/g, "'") : m[2] || null, ref: m[3], pos: i, len: m[0].length }); i += m[0].length; continue; }
    }
    if ((m = /^[A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*/.exec(rest))) {
      const name = m[0]; const after = src.slice(i + name.length).trimStart();
      if (after[0] === '(') t.push({ type: 'func', v: name.toUpperCase(), pos: i, len: name.length });
      else if (/^(TRUE|ИСТИНА)$/i.test(name)) t.push({ type: 'bool', v: true, pos: i, len: name.length });
      else if (/^(FALSE|ЛОЖЬ)$/i.test(name)) t.push({ type: 'bool', v: false, pos: i, len: name.length });
      else t.push({ type: 'name', v: name, pos: i, len: name.length });
      i += name.length; continue;
    }
    if ((m = /^(<>|<=|>=|[-+*/^&=<>%(),;:{}])/.exec(rest))) { t.push({ type: 'op', v: m[0] === ';' ? ',' : m[0], pos: i, len: m[0].length }); i += m[0].length; continue; }
    throw new Error('Неожиданный символ: ' + ch);
  }
  return t;
}

/* ---------------- parser (Pratt) ---------------- */
export function parse(src) {
  const toks = tokenize(src); let p = 0;
  const peek = () => toks[p]; const next = () => toks[p++];
  const expect = (v) => { const t = next(); if (!t || t.v !== v) throw new Error(`Ожидалось «${v}»`); };
  const BP = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  function expr(rbp = 0) {
    let left = prefix();
    for (;;) {
      const t = peek();
      if (!t || t.type !== 'op') break;
      if (t.v === '%') { next(); left = { k: 'pct', a: left }; continue; }
      const bp = BP[t.v]; if (!bp || bp <= rbp) break;
      next(); left = { k: 'bin', op: t.v, a: left, b: expr(t.v === '^' ? bp - 1 : bp) };
    }
    return left;
  }
  function prefix() {
    const t = next(); if (!t) throw new Error('Неполная формула');
    if (t.type === 'num') return { k: 'num', v: t.v };
    if (t.type === 'str') return { k: 'str', v: t.v };
    if (t.type === 'bool') return { k: 'bool', v: t.v };
    if (t.type === 'err') return { k: 'err', v: t.v };
    if (t.type === 'ref') return { k: 'ref', sheet: t.sheet, ref: t.ref };
    if (t.type === 'name') return { k: 'name', v: t.v };
    if (t.type === 'op' && (t.v === '-' || t.v === '+')) { const a = expr(5); return t.v === '-' ? { k: 'neg', a } : a; }
    if (t.type === 'op' && t.v === '(') { const e = expr(); expect(')'); return e; }
    if (t.type === 'op' && t.v === '{') { const rows = [[]]; for (;;) { const x = peek(); if (x?.v === '}') { next(); break; } if (x?.v === ',') { next(); continue; } rows[rows.length - 1].push(expr()); } return { k: 'arr', rows }; }
    if (t.type === 'func') {
      expect('('); const args = [];
      if (peek()?.v !== ')') {
        for (;;) {
          if (peek()?.v === ',' || peek()?.v === ')') args.push({ k: 'empty' }); else args.push(expr());
          if (peek()?.v === ',') { next(); continue; }
          break;
        }
      }
      expect(')');
      return { k: 'fn', name: ALIAS[t.v] || t.v, args };
    }
    throw new Error('Неожиданный элемент: ' + (t.v ?? t.ref));
  }
  const ast = expr();
  if (p < toks.length) throw new Error('Лишние символы в формуле');
  return ast;
}

/* ---------------- reference rewriting ---------------- */
function parseRefParts(ref) {
  const parts = ref.split(':').map((s) => {
    const m = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/.exec(s);
    return { ca: !!m[1] && !!m[2], c: m[2] ? colIndex(m[2]) : null, ra: m[3] === '$' || (!m[2] && m[1] === '$'), r: m[4] ? +m[4] - 1 : null };
  });
  return parts;
}
function partStr(p) { return (p.c != null ? (p.ca ? '$' : '') + colName(p.c) : '') + (p.r != null ? (p.ra ? '$' : '') + (p.r + 1) : ''); }
function rewrite(formula, fn) {
  let toks; try { toks = tokenize(formula.slice(1)); } catch { return formula; }
  let out = formula.slice(1); let shift = 0;
  for (const t of toks) {
    if (t.type !== 'ref') continue;
    const parts = parseRefParts(t.ref);
    const res = fn(parts, t.sheet);
    if (res === undefined) continue;
    const sheetPrefix = t.sheet ? (/^[\wА-Яа-яЁё.]+$/.test(t.sheet) ? t.sheet : `'${t.sheet.replace(/'/g, "''")}'`) + '!' : '';
    const str = res === null ? '#REF!' : sheetPrefix + res.map(partStr).join(':');
    out = out.slice(0, t.pos + shift) + str + out.slice(t.pos + shift + t.len);
    shift += str.length - t.len;
  }
  return '=' + out;
}
export function shiftFormula(formula, dr, dc) {
  return rewrite(formula, (parts) => {
    const res = parts.map((p) => ({ ...p, r: p.r != null && !p.ra ? p.r + dr : p.r, c: p.c != null && !p.ca ? p.c + dc : p.c }));
    if (res.some((p) => (p.r != null && p.r < 0) || (p.c != null && p.c < 0))) return null;
    return res;
  });
}
// Adjust refs when rows/cols are inserted (count>0) or deleted (count<0) at index `at`.
export function adjustFormula(formula, { sheet, isSameSheet, axis, at, count }) {
  return rewrite(formula, (parts, refSheet) => {
    const target = refSheet ? refSheet === sheet : isSameSheet; if (!target) return undefined;
    const key = axis === 'row' ? 'r' : 'c';
    if (parts.some((p) => p[key] == null)) return undefined;
    if (count > 0) return parts.map((p) => ({ ...p, [key]: p[key] >= at ? p[key] + count : p[key] }));
    const del = -count; const end = at + del - 1;
    if (parts.length === 1) { const v = parts[0][key]; if (v >= at && v <= end) return null; return [{ ...parts[0], [key]: v > end ? v - del : v }]; }
    let [a, b] = parts; let lo = Math.min(a[key], b[key]), hi = Math.max(a[key], b[key]);
    if (lo >= at && hi <= end) return null;
    const nl = lo > end ? lo - del : lo >= at ? at : lo; const nh = hi > end ? hi - del : hi >= at ? at - 1 : hi;
    return [{ ...a, [key]: nl }, { ...b, [key]: nh }];
  });
}
export function renameSheetInFormula(formula, oldName, newName) {
  return rewrite(formula, (parts, refSheet) => (refSheet === oldName ? parts : undefined)).replace(new RegExp(`(^|[^\\w'])(?:'${oldName.replace(/[.*+?^${}()|[\]\\']/g, (c) => (c === "'" ? "''" : '\\' + c))}'|${oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})!`, 'g'), (m, pre) => pre + (/^[\wА-Яа-яЁё.]+$/.test(newName) ? newName : `'${newName.replace(/'/g, "''")}'`) + '!');
}

/* ---------------- evaluator ---------------- */
class Range { constructor(sheet, r1, c1, r2, c2, get) { Object.assign(this, { sheet, r1, c1, r2, c2 }); this._get = get; }
  get rows() { if (!this._rows) { const out = []; for (let r = this.r1; r <= this.r2; r++) { const row = []; for (let c = this.c1; c <= this.c2; c++) row.push(this._get(r, c)); out.push(row); } this._rows = out; } return this._rows; }
  get h() { return this.r2 - this.r1 + 1; } get w() { return this.c2 - this.c1 + 1; }
  *values() { for (const row of this.rows) yield* row; }
}
const isRange = (v) => v instanceof Range || (v && v.arr);
const arrOf = (v) => (v instanceof Range ? v.rows : v && v.arr ? v.arr : [[v]]);

export class Engine {
  constructor(book) { this.book = book; this.cache = new Map(); this.astCache = new Map(); this.stack = new Set(); }
  invalidate() { this.cache.clear(); }
  sheetIndex(name) { if (name == null) return -1; const n = name.toLowerCase(); return this.book.sheets.findIndex((s) => s.name.toLowerCase() === n); }
  raw(si, r, c) { return this.book.sheets[si]?.cells.get(r * MAX_COLS + c)?.v; }
  value(si, r, c) {
    const key = si * 1e8 + r * MAX_COLS + c;
    if (this.cache.has(key)) return this.cache.get(key);
    const raw = this.raw(si, r, c);
    let v;
    if (raw == null || raw === '') v = null;
    else if (typeof raw === 'number' || typeof raw === 'boolean') v = raw;
    else if (raw[0] === '=' && raw.length > 1) {
      if (this.stack.has(key)) return ERR.circ;
      this.stack.add(key);
      try { v = this.evalFormula(raw, si, r, c); } finally { this.stack.delete(key); }
      if (isRange(v)) v = arrOf(v)[0]?.[0] ?? null;
      if (v == null) v = 0;
    } else v = raw;
    this.cache.set(key, v);
    return v;
  }
  ast(src) { let a = this.astCache.get(src); if (!a) { try { a = parse(src.slice(1)); } catch (e) { a = { k: 'perr', msg: e.message }; } this.astCache.set(src, a); if (this.astCache.size > 20000) this.astCache.clear(); } return a; }
  evalFormula(src, si, r, c) {
    const a = this.ast(src); if (a.k === 'perr') return ERR.name;
    const prev = this.ctx; this.ctx = { si, r, c };
    try { return this.ev(a, si); } catch (e) { if (e instanceof FErr) return e; console.warn(e); return ERR.value; } finally { this.ctx = prev; }
  }
  usedBounds(si) { const s = this.book.sheets[si]; let mr = 0, mc = 0; for (const k of s.cells.keys()) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (r > mr) mr = r; if (c > mc) mc = c; } return { mr, mc }; }
  refRange(node, si) {
    const tsi = node.sheet ? this.sheetIndex(node.sheet) : si;
    if (tsi < 0) throw ERR.ref;
    const parts = parseRefParts(node.ref);
    let [a, b] = parts; b = b || a;
    let r1 = a.r, c1 = a.c, r2 = b.r, c2 = b.c;
    if (r1 == null) { const ub = this.usedBounds(tsi); r1 = 0; r2 = ub.mr; }
    if (c1 == null) { const ub = this.usedBounds(tsi); c1 = 0; c2 = ub.mc; }
    if (r1 > r2) [r1, r2] = [r2, r1]; if (c1 > c2) [c1, c2] = [c2, c1];
    return new Range(tsi, r1, c1, r2, c2, (rr, cc) => this.value(tsi, rr, cc));
  }
  ev(n, si) {
    switch (n.k) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return n.v;
      case 'empty': return null;
      case 'arr': return { arr: n.rows.map((row) => row.map((x) => this.ev(x, si))) };
      case 'name': { // named range (Диспетчер имён)
        const nm = this.book.names?.[n.v.toUpperCase()];
        if (!nm) throw ERR.name;
        return this.ev({ k: 'ref', sheet: nm.sheet, ref: nm.ref }, si);
      }
      case 'ref': { const rg = this.refRange(n, si); return rg.r1 === rg.r2 && rg.c1 === rg.c2 && !n.ref.includes(':') ? this.value(rg.sheet, rg.r1, rg.c1) : rg; }
      case 'neg': { const v = num(this.scalar(this.ev(n.a, si))); return -v; }
      case 'pct': return num(this.scalar(this.ev(n.a, si))) / 100;
      case 'bin': return this.bin(n.op, this.scalar(this.ev(n.a, si)), this.scalar(this.ev(n.b, si)));
      case 'fn': return this.call(n, si);
    }
    throw ERR.value;
  }
  scalar(v) { if (isRange(v)) { const a = arrOf(v); if (a.length === 1 && a[0].length === 1) return a[0][0]; const { r, c } = this.ctx; if (v instanceof Range) { if (a[0].length === 1 && r >= v.r1 && r <= v.r2) return a[r - v.r1][0]; if (a.length === 1 && c >= v.c1 && c <= v.c2) return a[0][c - v.c1]; } return ERR.value; } return v; }
  bin(op, a, b) {
    if (isErr(a)) return a; if (isErr(b)) return b;
    if (op === '&') return txt(a) + txt(b);
    if (['=', '<>', '<', '>', '<=', '>='].includes(op)) {
      const cmp = compare(a, b);
      return { '=': cmp === 0, '<>': cmp !== 0, '<': cmp < 0, '>': cmp > 0, '<=': cmp <= 0, '>=': cmp >= 0 }[op];
    }
    const x = num(a), y = num(b);
    switch (op) { case '+': return x + y; case '-': return x - y; case '*': return x * y; case '/': if (y === 0) throw ERR.div0; return x / y; case '^': { const r = Math.pow(x, y); if (!isFinite(r)) throw ERR.num; return r; } }
    throw ERR.value;
  }
  call(n, si) {
    const name = n.name;
    if (name === 'IF') { const c = this.scalar(this.ev(n.args[0], si)); if (isErr(c)) return c; return bool(c) ? (n.args[1] ? this.ev(n.args[1], si) : true) : (n.args[2] ? this.ev(n.args[2], si) : false); }
    if (name === 'IFERROR' || name === 'IFNA') { let v; try { v = this.scalar(this.ev(n.args[0], si)); } catch (e) { v = e instanceof FErr ? e : ERR.value; } return isErr(v) && (name === 'IFERROR' || v.e === '#N/A') ? this.ev(n.args[1], si) : v; }
    if (name === 'IFS') { for (let i = 0; i + 1 < n.args.length; i += 2) if (bool(this.scalar(this.ev(n.args[i], si)))) return this.ev(n.args[i + 1], si); return ERR.na; }
    if (name === 'CHOOSE') { const i = Math.floor(num(this.scalar(this.ev(n.args[0], si)))); if (i < 1 || i >= n.args.length) throw ERR.value; return this.ev(n.args[i], si); }
    if (name === 'ROW' || name === 'COLUMN') { if (!n.args.length) return name === 'ROW' ? this.ctx.r + 1 : this.ctx.c + 1; const a = n.args[0]; if (a.k !== 'ref') throw ERR.value; const rg = this.refRange(a, si); return name === 'ROW' ? rg.r1 + 1 : rg.c1 + 1; }
    if (name === 'ISBLANK') { const v = this.ev(n.args[0], si); return (isRange(v) ? arrOf(v)[0][0] : v) == null; }
    const fn = FN[name]; if (!fn) return ERR.name;
    const args = n.args.map((a) => { try { return this.ev(a, si); } catch (e) { if (e instanceof FErr) return e; throw e; } });
    const ctx = this.ctx;
    const res = fn(args, this, ctx);
    this.ctx = ctx;
    return res;
  }
}

/* ---------------- coercion helpers ---------------- */
function num(v) {
  if (isErr(v)) throw v;
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const s = String(v).trim().replace(/\s/g, '').replace(',', '.');
  if (s !== '' && isFinite(+s)) return +s;
  const d = parseDateInput(String(v).trim()); if (d) return d.v;
  throw ERR.value;
}
function txt(v) { if (isErr(v)) throw v; if (v == null) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return fmtGeneral(v); return String(v); }
function bool(v) { if (isErr(v)) throw v; if (typeof v === 'string') { if (/^(true|истина)$/i.test(v)) return true; if (/^(false|ложь)$/i.test(v)) return false; throw ERR.value; } return !!v; }
function compare(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : 0; if (b == null) b = typeof a === 'string' ? '' : 0;
  const ta = typeof a, tb = typeof b;
  if (ta === tb) { if (ta === 'string') return a.localeCompare(b, 'ru', { sensitivity: 'base' }); return a < b ? -1 : a > b ? 1 : 0; }
  const rank = { number: 0, string: 1, boolean: 2 }; return rank[ta] - rank[tb];
}
export function fmtGeneral(v) {
  if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
  const a = Math.abs(v);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
  return String(+v.toPrecision(11));
}
function* flatAll(args) { for (const a of args) { if (isRange(a)) { for (const row of arrOf(a)) yield* row; } else yield a; } }
function nums(args, { strict = false } = {}) {
  const out = [];
  for (const a of args) {
    if (isRange(a)) { for (const row of arrOf(a)) for (const v of row) { if (isErr(v)) throw v; if (typeof v === 'number') out.push(v); } }
    else if (isErr(a)) throw a;
    else if (a == null) { if (strict) out.push(0); }
    else out.push(num(a));
  }
  return out;
}
const S = (args, i) => { const v = args[i]; return isRange(v) ? arrOf(v)[0][0] : v; };
const N = (args, i, d) => (args[i] === undefined || args[i] === null && d !== undefined ? d : num(S(args, i)));
const T = (args, i, d = '') => (args[i] === undefined ? d : txt(S(args, i)));
function criteria(c) {
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => compare(v, c) === 0 && v != null;
  const s = String(c ?? ''); const m = /^(<>|<=|>=|=|<|>)?(.*)$/s.exec(s); const op = m[1] || '='; let rhs = m[2];
  const rnum = rhs !== '' && isFinite(+rhs.replace(',', '.')) ? +rhs.replace(',', '.') : null;
  if ((op === '=' || op === '<>') && rnum == null) {
    const re = /[*?]/.test(rhs) ? new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i') : null;
    const f = (v) => (re ? re.test(txt(v ?? '')) : txt(v ?? '').toLowerCase() === rhs.toLowerCase());
    return op === '=' ? (v) => (rhs === '' ? v == null || v === '' : f(v)) : (v) => !f(v);
  }
  return (v) => { if (typeof v !== 'number') { if (op === '<>') return true; return false; } const cmp = v - rnum; return { '=': cmp === 0, '<>': cmp !== 0, '<': cmp < 0, '>': cmp > 0, '<=': cmp <= 0, '>=': cmp >= 0 }[op]; };
}
function lookupMatch(arr, key, mode) {
  // mode 0 exact; 1 largest <= key (sorted asc); -1 smallest >= key (sorted desc)
  if (mode === 0) { const f = criteria(typeof key === 'string' ? key : key); for (let i = 0; i < arr.length; i++) if (typeof key === 'string' ? f(arr[i]) : compare(arr[i], key) === 0) return i; return -1; }
  let best = -1; for (let i = 0; i < arr.length; i++) { const c = compare(arr[i], key); if (mode === 1 ? c <= 0 : c >= 0) best = i; else if (mode === 1) break; } return best;
}
const round = (x, d) => { const f = Math.pow(10, d); return Math.round((x + Number.EPSILON * Math.sign(x)) * f) / f; };
function stat(xs) { const n = xs.length; const m = xs.reduce((a, b) => a + b, 0) / n; const v = xs.reduce((a, b) => a + (b - m) ** 2, 0); return { n, m, v }; }
function fmtText(v, f) {
  if (typeof v !== 'number') return txt(v);
  f = String(f);
  if (/[дdгyмmчh]/i.test(f) && !/^[0#,.\s%]+$/.test(f)) {
    const d = serialToDate(v);
    const p = (x, n = 2) => String(x).padStart(n, '0');
    return f.replace(/гггг|yyyy/gi, d.getUTCFullYear()).replace(/гг|yy/gi, p(d.getUTCFullYear() % 100)).replace(/ММ|MM/g, p(d.getUTCMonth() + 1)).replace(/ДД|dd/gi, p(d.getUTCDate())).replace(/чч|hh/gi, p(d.getUTCHours())).replace(/мм|mm/g, p(d.getUTCMinutes())).replace(/сс|ss/gi, p(d.getUTCSeconds()));
  }
  const pct = f.includes('%'); if (pct) v *= 100;
  const dec = (f.split(/[.,]/)[1] || '').replace(/[^0#]/g, '').length;
  const group = /#[ ,]#|0[ ,]0|# ##/.test(f) || f.includes(' ');
  return v.toLocaleString('ru-RU', { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: group }) + (pct ? '%' : '');
}

/* ---------------- function library ---------------- */
export const FN = {
  SUM: (a) => nums(a).reduce((x, y) => x + y, 0),
  AVERAGE: (a) => { const x = nums(a); if (!x.length) throw ERR.div0; return x.reduce((p, q) => p + q, 0) / x.length; },
  AVERAGEA: (a) => { const x = [...flatAll(a)].filter((v) => v != null).map((v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0)); if (!x.length) throw ERR.div0; return x.reduce((p, q) => p + q, 0) / x.length; },
  MIN: (a) => { const x = nums(a); return x.length ? Math.min(...x) : 0; },
  MAX: (a) => { const x = nums(a); return x.length ? Math.max(...x) : 0; },
  COUNT: (a) => [...flatAll(a)].filter((v) => typeof v === 'number').length,
  COUNTA: (a) => [...flatAll(a)].filter((v) => v != null && v !== '').length,
  COUNTBLANK: (a) => [...flatAll(a)].filter((v) => v == null || v === '').length,
  PRODUCT: (a) => nums(a).reduce((x, y) => x * y, 1),
  MEDIAN: (a) => { const x = nums(a).sort((p, q) => p - q); if (!x.length) throw ERR.num; const m = x.length >> 1; return x.length % 2 ? x[m] : (x[m - 1] + x[m]) / 2; },
  MODE: (a) => { const x = nums(a); const cnt = new Map(); let best = null, bc = 1; for (const v of x) { const c = (cnt.get(v) || 0) + 1; cnt.set(v, c); if (c > bc) { bc = c; best = v; } } if (best == null) throw ERR.na; return best; },
  STDEV: (a) => { const s = stat(nums(a)); if (s.n < 2) throw ERR.div0; return Math.sqrt(s.v / (s.n - 1)); },
  STDEVP: (a) => { const s = stat(nums(a)); if (!s.n) throw ERR.div0; return Math.sqrt(s.v / s.n); },
  VAR: (a) => { const s = stat(nums(a)); if (s.n < 2) throw ERR.div0; return s.v / (s.n - 1); },
  VARP: (a) => { const s = stat(nums(a)); if (!s.n) throw ERR.div0; return s.v / s.n; },
  LARGE: (a) => { const x = nums([a[0]]).sort((p, q) => q - p); const k = num(S(a, 1)); if (k < 1 || k > x.length) throw ERR.num; return x[k - 1]; },
  SMALL: (a) => { const x = nums([a[0]]).sort((p, q) => p - q); const k = num(S(a, 1)); if (k < 1 || k > x.length) throw ERR.num; return x[k - 1]; },
  RANK: (a) => { const v = num(S(a, 0)); const x = nums([a[1]]); const asc = a[2] !== undefined && num(S(a, 2)) !== 0; return 1 + x.filter((y) => (asc ? y < v : y > v)).length; },
  PERCENTILE: (a) => { const x = nums([a[0]]).sort((p, q) => p - q); const k = num(S(a, 1)); if (!x.length || k < 0 || k > 1) throw ERR.num; const i = k * (x.length - 1); const lo = Math.floor(i); return x[lo] + (x[Math.min(lo + 1, x.length - 1)] - x[lo]) * (i - lo); },
  QUARTILE: (a, e, c) => FN.PERCENTILE([a[0], num(S(a, 1)) / 4], e, c),
  CORREL: (a) => { const x = arrOf(a[0]).flat().map(num), y = arrOf(a[1]).flat().map(num); const n = Math.min(x.length, y.length); const mx = x.reduce((p, q) => p + q, 0) / n, my = y.reduce((p, q) => p + q, 0) / n; let sxy = 0, sxx = 0, syy = 0; for (let i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; } return sxy / Math.sqrt(sxx * syy); },
  SUMSQ: (a) => nums(a).reduce((x, y) => x + y * y, 0),
  SUMPRODUCT: (a) => { const ms = a.map(arrOf); const h = ms[0].length, w = ms[0][0].length; let s = 0; for (let r = 0; r < h; r++) for (let c = 0; c < w; c++) { let p = 1; for (const m of ms) { const v = m[r]?.[c]; p *= typeof v === 'number' ? v : 0; } s += p; } return s; },
  SUMIF: (a) => { const rng = arrOf(a[0]); const f = criteria(S(a, 1)); const sr = a[2] !== undefined ? arrOf(a[2]) : rng; let s = 0; rng.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = sr[i]?.[j]; if (typeof x === 'number') s += x; } })); return s; },
  SUMIFS: (a) => { const sr = arrOf(a[0]); let s = 0; const conds = []; for (let i = 1; i + 1 < a.length; i += 2) conds.push([arrOf(a[i]), criteria(S(a, i + 1))]); sr.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'number' && conds.every(([r, f]) => f(r[i]?.[j]))) s += v; })); return s; },
  COUNTIF: (a) => { const f = criteria(S(a, 1)); let n = 0; for (const row of arrOf(a[0])) for (const v of row) if (f(v)) n++; return n; },
  COUNTIFS: (a) => { const conds = []; for (let i = 0; i + 1 < a.length; i += 2) conds.push([arrOf(a[i]), criteria(S(a, i + 1))]); let n = 0; conds[0][0].forEach((row, i) => row.forEach((_, j) => { if (conds.every(([r, f]) => f(r[i]?.[j]))) n++; })); return n; },
  AVERAGEIF: (a) => { const rng = arrOf(a[0]); const f = criteria(S(a, 1)); const sr = a[2] !== undefined ? arrOf(a[2]) : rng; let s = 0, n = 0; rng.forEach((row, i) => row.forEach((v, j) => { if (f(v)) { const x = sr[i]?.[j]; if (typeof x === 'number') { s += x; n++; } } })); if (!n) throw ERR.div0; return s / n; },
  AVERAGEIFS: (a) => { const sr = arrOf(a[0]); let s = 0, n = 0; const conds = []; for (let i = 1; i + 1 < a.length; i += 2) conds.push([arrOf(a[i]), criteria(S(a, i + 1))]); sr.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'number' && conds.every(([r, f]) => f(r[i]?.[j]))) { s += v; n++; } })); if (!n) throw ERR.div0; return s / n; },
  MAXIFS: (a) => { const sr = arrOf(a[0]); const conds = []; for (let i = 1; i + 1 < a.length; i += 2) conds.push([arrOf(a[i]), criteria(S(a, i + 1))]); let m = null; sr.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'number' && conds.every(([r, f]) => f(r[i]?.[j]))) m = m == null ? v : Math.max(m, v); })); return m ?? 0; },
  MINIFS: (a) => { const sr = arrOf(a[0]); const conds = []; for (let i = 1; i + 1 < a.length; i += 2) conds.push([arrOf(a[i]), criteria(S(a, i + 1))]); let m = null; sr.forEach((row, i) => row.forEach((v, j) => { if (typeof v === 'number' && conds.every(([r, f]) => f(r[i]?.[j]))) m = m == null ? v : Math.min(m, v); })); return m ?? 0; },
  ABS: (a) => Math.abs(N(a, 0)), SQRT: (a) => { const x = N(a, 0); if (x < 0) throw ERR.num; return Math.sqrt(x); },
  POWER: (a) => Math.pow(N(a, 0), N(a, 1)), EXP: (a) => Math.exp(N(a, 0)),
  LN: (a) => { const x = N(a, 0); if (x <= 0) throw ERR.num; return Math.log(x); }, LOG10: (a) => Math.log10(N(a, 0)),
  LOG: (a) => Math.log(N(a, 0)) / Math.log(a[1] !== undefined ? N(a, 1) : 10),
  MOD: (a) => { const d = N(a, 1); if (d === 0) throw ERR.div0; const x = N(a, 0); return x - d * Math.floor(x / d); },
  INT: (a) => Math.floor(N(a, 0)), TRUNC: (a) => { const f = Math.pow(10, a[1] !== undefined ? N(a, 1) : 0); return Math.trunc(N(a, 0) * f) / f; },
  ROUND: (a) => round(N(a, 0), a[1] !== undefined ? N(a, 1) : 0),
  ROUNDUP: (a) => { const f = Math.pow(10, a[1] !== undefined ? N(a, 1) : 0); const x = N(a, 0); return Math.sign(x) * Math.ceil(Math.abs(x) * f - 1e-9) / f; },
  ROUNDDOWN: (a) => { const f = Math.pow(10, a[1] !== undefined ? N(a, 1) : 0); const x = N(a, 0); return Math.sign(x) * Math.floor(Math.abs(x) * f + 1e-9) / f; },
  CEILING: (a) => { const s = a[1] !== undefined ? N(a, 1) : 1; return s ? Math.ceil(N(a, 0) / s) * s : 0; },
  FLOOR: (a) => { const s = a[1] !== undefined ? N(a, 1) : 1; return s ? Math.floor(N(a, 0) / s) * s : 0; },
  MROUND: (a) => { const m = N(a, 1); return m ? Math.round(N(a, 0) / m) * m : 0; },
  SIGN: (a) => Math.sign(N(a, 0)), PI: () => Math.PI, RAND: () => Math.random(),
  RANDBETWEEN: (a) => { const lo = Math.ceil(N(a, 0)), hi = Math.floor(N(a, 1)); return lo + Math.floor(Math.random() * (hi - lo + 1)); },
  SIN: (a) => Math.sin(N(a, 0)), COS: (a) => Math.cos(N(a, 0)), TAN: (a) => Math.tan(N(a, 0)), ASIN: (a) => Math.asin(N(a, 0)), ACOS: (a) => Math.acos(N(a, 0)), ATAN: (a) => Math.atan(N(a, 0)), ATAN2: (a) => Math.atan2(N(a, 1), N(a, 0)),
  DEGREES: (a) => (N(a, 0) * 180) / Math.PI, RADIANS: (a) => (N(a, 0) * Math.PI) / 180,
  FACT: (a) => { let n = Math.floor(N(a, 0)), r = 1; if (n < 0) throw ERR.num; while (n > 1) r *= n--; return r; },
  GCD: (a) => nums(a).map(Math.floor).reduce((x, y) => { while (y) [x, y] = [y, x % y]; return x; }, 0),
  LCM: (a) => nums(a).map(Math.floor).reduce((x, y) => { const g = FN.GCD([x, y]); return g ? (x * y) / g : 0; }, 1),
  EVEN: (a) => { const x = N(a, 0); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; }, ODD: (a) => { const x = N(a, 0); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r++; return x < 0 ? -r : r; },
  AND: (a) => [...flatAll(a)].filter((v) => v != null).every((v) => bool(v)), OR: (a) => [...flatAll(a)].filter((v) => v != null).some((v) => bool(v)),
  XOR: (a) => [...flatAll(a)].filter((v) => v != null).reduce((x, v) => x !== bool(v), false), NOT: (a) => !bool(S(a, 0)),
  TRUE: () => true, FALSE: () => false,
  ISNUMBER: (a) => typeof S(a, 0) === 'number', ISTEXT: (a) => typeof S(a, 0) === 'string', ISERROR: (a) => isErr(S(a, 0)), ISERR: (a) => isErr(S(a, 0)) && S(a, 0).e !== '#N/A', ISNA: (a) => isErr(S(a, 0)) && S(a, 0).e === '#N/A', ISLOGICAL: (a) => typeof S(a, 0) === 'boolean',
  ISEVEN: (a) => Math.floor(N(a, 0)) % 2 === 0, ISODD: (a) => Math.abs(Math.floor(N(a, 0)) % 2) === 1, NA: () => ERR.na,
  CONCAT: (a) => [...flatAll(a)].map(txt).join(''), CONCATENATE: (a) => [...flatAll(a)].map(txt).join(''),
  TEXTJOIN: (a) => { const d = T(a, 0); const skip = bool(S(a, 1)); return [...flatAll(a.slice(2))].filter((v) => !(skip && (v == null || v === ''))).map(txt).join(d); },
  LEN: (a) => T(a, 0).length, UPPER: (a) => T(a, 0).toLocaleUpperCase('ru'), LOWER: (a) => T(a, 0).toLocaleLowerCase('ru'),
  PROPER: (a) => T(a, 0).toLocaleLowerCase('ru').replace(/(^|[^\p{L}])(\p{L})/gu, (m, p, c) => p + c.toLocaleUpperCase('ru')),
  TRIM: (a) => T(a, 0).trim().replace(/\s+/g, ' '), CLEAN: (a) => T(a, 0).replace(/[\x00-\x1f]/g, ''),
  LEFT: (a) => T(a, 0).slice(0, a[1] !== undefined ? N(a, 1) : 1), RIGHT: (a) => { const s = T(a, 0), n = a[1] !== undefined ? N(a, 1) : 1; return n ? s.slice(-n) : ''; },
  MID: (a) => T(a, 0).substr(N(a, 1) - 1, N(a, 2)),
  FIND: (a) => { const i = T(a, 1).indexOf(T(a, 0), (a[2] !== undefined ? N(a, 2) : 1) - 1); if (i < 0) throw ERR.value; return i + 1; },
  SEARCH: (a) => { const i = T(a, 1).toLowerCase().indexOf(T(a, 0).toLowerCase(), (a[2] !== undefined ? N(a, 2) : 1) - 1); if (i < 0) throw ERR.value; return i + 1; },
  SUBSTITUTE: (a) => { const s = T(a, 0), o = T(a, 1), nw = T(a, 2); if (!o) return s; if (a[3] === undefined) return s.split(o).join(nw); let k = N(a, 3), i = -1; while (k-- > 0) { i = s.indexOf(o, i + 1); if (i < 0) return s; } return s.slice(0, i) + nw + s.slice(i + o.length); },
  REPLACE: (a) => { const s = T(a, 0), st = N(a, 1) - 1, n = N(a, 2); return s.slice(0, st) + T(a, 3) + s.slice(st + n); },
  REPT: (a) => T(a, 0).repeat(Math.max(0, N(a, 1))), EXACT: (a) => T(a, 0) === T(a, 1),
  TEXT: (a) => fmtText(S(a, 0), T(a, 1)), VALUE: (a) => num(S(a, 0)), T: (a) => (typeof S(a, 0) === 'string' ? S(a, 0) : ''), N: (a) => (typeof S(a, 0) === 'number' ? S(a, 0) : typeof S(a, 0) === 'boolean' ? +S(a, 0) : 0),
  CHAR: (a) => String.fromCharCode(N(a, 0)), CODE: (a) => T(a, 0).charCodeAt(0) || 0, UNICHAR: (a) => String.fromCodePoint(N(a, 0)), UNICODE: (a) => T(a, 0).codePointAt(0) || 0,
  FIXED: (a) => num(S(a, 0)).toLocaleString('ru-RU', { minimumFractionDigits: a[1] !== undefined ? N(a, 1) : 2, maximumFractionDigits: a[1] !== undefined ? N(a, 1) : 2, useGrouping: !(a[2] !== undefined && bool(S(a, 2))) }),
  TODAY: () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); },
  NOW: () => { const d = new Date(); return dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; },
  DATE: (a) => dateToSerial(N(a, 0), N(a, 1), N(a, 2)), TIME: (a) => (N(a, 0) * 3600 + N(a, 1) * 60 + N(a, 2)) / 86400,
  YEAR: (a) => serialToDate(N(a, 0)).getUTCFullYear(), MONTH: (a) => serialToDate(N(a, 0)).getUTCMonth() + 1, DAY: (a) => serialToDate(N(a, 0)).getUTCDate(),
  HOUR: (a) => serialToDate(N(a, 0)).getUTCHours(), MINUTE: (a) => serialToDate(N(a, 0)).getUTCMinutes(), SECOND: (a) => serialToDate(N(a, 0)).getUTCSeconds(),
  WEEKDAY: (a) => { const d = serialToDate(N(a, 0)).getUTCDay(); const t = a[1] !== undefined ? N(a, 1) : 1; return t === 2 ? ((d + 6) % 7) + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
  WEEKNUM: (a) => { const d = serialToDate(N(a, 0)); const y0 = Date.UTC(d.getUTCFullYear(), 0, 1); const wd = new Date(y0).getUTCDay(); return Math.floor(((d - y0) / 86400000 + wd) / 7) + 1; },
  EDATE: (a) => { const d = serialToDate(N(a, 0)); return dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + 1 + N(a, 1), d.getUTCDate()); },
  EOMONTH: (a) => { const d = serialToDate(N(a, 0)); return dateToSerial(d.getUTCFullYear(), d.getUTCMonth() + 2 + N(a, 1), 0); },
  DAYS: (a) => Math.floor(N(a, 0)) - Math.floor(N(a, 1)),
  DATEDIF: (a) => { const s = serialToDate(N(a, 0)), e = serialToDate(N(a, 1)); const u = T(a, 2).toUpperCase(); if (u === 'D') return Math.floor((e - s) / 86400000); let m = (e.getUTCFullYear() - s.getUTCFullYear()) * 12 + e.getUTCMonth() - s.getUTCMonth(); if (e.getUTCDate() < s.getUTCDate()) m--; return u === 'M' ? m : u === 'Y' ? Math.floor(m / 12) : ERR.value; },
  NETWORKDAYS: (a) => { let s = Math.floor(N(a, 0)), e = Math.floor(N(a, 1)), n = 0, sg = 1; if (s > e) { [s, e] = [e, s]; sg = -1; } for (let d = s; d <= e; d++) { const w = serialToDate(d).getUTCDay(); if (w !== 0 && w !== 6) n++; } return n * sg; },
  WORKDAY: (a) => { let d = Math.floor(N(a, 0)), k = N(a, 1); const st = k < 0 ? -1 : 1; while (k !== 0) { d += st; const w = serialToDate(d).getUTCDay(); if (w !== 0 && w !== 6) k -= st; } return d; },
  DATEVALUE: (a) => { const p = parseDateInput(T(a, 0)); if (!p) throw ERR.value; return Math.floor(p.v); },
  VLOOKUP: (a) => { const key = S(a, 0); const tbl = arrOf(a[1]); const col = N(a, 2); const approx = a[3] === undefined || bool(S(a, 3)); const i = lookupMatch(tbl.map((r) => r[0]), key, approx ? 1 : 0); if (i < 0) return ERR.na; if (col < 1 || col > tbl[0].length) throw ERR.ref; return tbl[i][col - 1] ?? 0; },
  HLOOKUP: (a) => { const key = S(a, 0); const tbl = arrOf(a[1]); const row = N(a, 2); const approx = a[3] === undefined || bool(S(a, 3)); const i = lookupMatch(tbl[0], key, approx ? 1 : 0); if (i < 0) return ERR.na; return tbl[row - 1]?.[i] ?? ERR.ref; },
  XLOOKUP: (a) => { const key = S(a, 0); const look = arrOf(a[1]).flat(); const ret = arrOf(a[2]); const i = lookupMatch(look, key, 0); if (i < 0) return a[3] !== undefined ? S(a, 3) : ERR.na; return ret.length === 1 ? ret[0][i] : ret[i][0]; },
  LOOKUP: (a) => { const key = S(a, 0); const look = arrOf(a[1]).flat(); const ret = a[2] !== undefined ? arrOf(a[2]).flat() : look; const i = lookupMatch(look, key, 1); return i < 0 ? ERR.na : ret[i]; },
  MATCH: (a) => { const key = S(a, 0); const arr = arrOf(a[1]).flat(); const t = a[2] !== undefined ? N(a, 2) : 1; const i = lookupMatch(arr, key, t === 0 ? 0 : t > 0 ? 1 : -1); return i < 0 ? ERR.na : i + 1; },
  INDEX: (a) => { const m = arrOf(a[0]); let r = a[1] !== undefined ? N(a, 1) : 1, c = a[2] !== undefined ? N(a, 2) : 1; if (m.length === 1 && a[2] === undefined) { c = r; r = 1; } const v = m[r - 1]?.[c - 1]; if (v === undefined) throw ERR.ref; return v; },
  ROWS: (a) => arrOf(a[0]).length, COLUMNS: (a) => arrOf(a[0])[0].length,
  TRANSPOSE: (a) => { const m = arrOf(a[0]); return { arr: m[0].map((_, c) => m.map((r) => r[c])) }; },
  UNIQUE: (a) => { const seen = new Set(); const out = []; for (const row of arrOf(a[0])) { const k = JSON.stringify(row); if (!seen.has(k)) { seen.add(k); out.push(row); } } return { arr: out }; },
  PMT: (a) => { const r = N(a, 0), n = N(a, 1), pv = N(a, 2), fv = a[3] !== undefined ? N(a, 3) : 0, t = a[4] !== undefined ? N(a, 4) : 0; if (!r) return -(pv + fv) / n; const x = Math.pow(1 + r, n); return -(r * (pv * x + fv)) / ((1 + r * t) * (x - 1)); },
  FV: (a) => { const r = N(a, 0), n = N(a, 1), p = N(a, 2), pv = a[3] !== undefined ? N(a, 3) : 0, t = a[4] !== undefined ? N(a, 4) : 0; if (!r) return -(pv + p * n); const x = Math.pow(1 + r, n); return -(pv * x + (p * (1 + r * t) * (x - 1)) / r); },
  PV: (a) => { const r = N(a, 0), n = N(a, 1), p = N(a, 2), fv = a[3] !== undefined ? N(a, 3) : 0, t = a[4] !== undefined ? N(a, 4) : 0; if (!r) return -(fv + p * n); const x = Math.pow(1 + r, n); return -(fv + (p * (1 + r * t) * (x - 1)) / r) / x; },
  NPV: (a) => { const r = N(a, 0); return nums(a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); },
  IRR: (a) => { const cf = nums([a[0]]); let r = a[1] !== undefined ? N(a, 1) : 0.1; for (let i = 0; i < 100; i++) { let f = 0, d = 0; cf.forEach((v, t) => { f += v / Math.pow(1 + r, t); d -= (t * v) / Math.pow(1 + r, t + 1); }); const nr = r - f / d; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw ERR.num; },
  NPER: (a) => { const r = N(a, 0), p = N(a, 1), pv = N(a, 2), fv = a[3] !== undefined ? N(a, 3) : 0; if (!r) return -(pv + fv) / p; return Math.log((p - fv * r) / (p + pv * r)) / Math.log(1 + r); },
  ADDRESS: (a) => { const abs = a[2] !== undefined ? N(a, 2) : 1; return (abs === 1 || abs === 3 ? '$' : '') + colName(N(a, 1) - 1) + (abs <= 2 ? '$' : '') + N(a, 0); },
  SUBTOTAL: (a, e, c) => { const code = N(a, 0) % 100; const f = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' }[code]; if (!f) throw ERR.value; return FN[f](a.slice(1), e, c); },
  SEQUENCE: (a) => { const r = N(a, 0), c = a[1] !== undefined ? N(a, 1) : 1, s = a[2] !== undefined ? N(a, 2) : 1, st = a[3] !== undefined ? N(a, 3) : 1; return { arr: Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => s + (i * c + j) * st)) }; },
  HYPERLINK: (a) => (a[1] !== undefined ? T(a, 1) : T(a, 0)),
  SWITCH: (a) => { const v = S(a, 0); for (let i = 1; i + 1 < a.length; i += 2) if (compare(v, S(a, i)) === 0) return S(a, i + 1); return a.length % 2 === 0 ? S(a, a.length - 1) : ERR.na; },
};
export const ALIAS = {
  СУММ: 'SUM', СРЗНАЧ: 'AVERAGE', СРЗНАЧА: 'AVERAGEA', МИН: 'MIN', МАКС: 'MAX', СЧЁТ: 'COUNT', СЧЕТ: 'COUNT', СЧЁТЗ: 'COUNTA', СЧЕТЗ: 'COUNTA', СЧИТАТЬПУСТОТЫ: 'COUNTBLANK',
  ПРОИЗВЕД: 'PRODUCT', МЕДИАНА: 'MEDIAN', МОДА: 'MODE', СТАНДОТКЛОН: 'STDEV', СТАНДОТКЛОНП: 'STDEVP', ДИСП: 'VAR', ДИСПР: 'VARP', НАИБОЛЬШИЙ: 'LARGE', НАИМЕНЬШИЙ: 'SMALL', РАНГ: 'RANK', ПЕРСЕНТИЛЬ: 'PERCENTILE', КВАРТИЛЬ: 'QUARTILE', КОРРЕЛ: 'CORREL',
  СУММКВ: 'SUMSQ', СУММПРОИЗВ: 'SUMPRODUCT', СУММЕСЛИ: 'SUMIF', СУММЕСЛИМН: 'SUMIFS', СЧЁТЕСЛИ: 'COUNTIF', СЧЕТЕСЛИ: 'COUNTIF', СЧЁТЕСЛИМН: 'COUNTIFS', СЧЕТЕСЛИМН: 'COUNTIFS', СРЗНАЧЕСЛИ: 'AVERAGEIF', СРЗНАЧЕСЛИМН: 'AVERAGEIFS', МАКСЕСЛИ: 'MAXIFS', МИНЕСЛИ: 'MINIFS',
  КОРЕНЬ: 'SQRT', СТЕПЕНЬ: 'POWER', ОСТАТ: 'MOD', ЦЕЛОЕ: 'INT', ОТБР: 'TRUNC', ОКРУГЛ: 'ROUND', ОКРУГЛВВЕРХ: 'ROUNDUP', ОКРУГЛВНИЗ: 'ROUNDDOWN', ОКРВВЕРХ: 'CEILING', ОКРВНИЗ: 'FLOOR', ОКРУГЛТ: 'MROUND', ЗНАК: 'SIGN', ПИ: 'PI', СЛЧИС: 'RAND', СЛУЧМЕЖДУ: 'RANDBETWEEN',
  ФАКТР: 'FACT', НОД: 'GCD', НОК: 'LCM', ГРАДУСЫ: 'DEGREES', РАДИАНЫ: 'RADIANS',
  ЕСЛИ: 'IF', ЕСЛИМН: 'IFS', ЕСЛИОШИБКА: 'IFERROR', ЕСНД: 'IFNA', И: 'AND', ИЛИ: 'OR', НЕ: 'NOT', ИСКЛИЛИ: 'XOR', ВЫБОР: 'CHOOSE', ПЕРЕКЛЮЧ: 'SWITCH',
  ЕПУСТО: 'ISBLANK', ЕЧИСЛО: 'ISNUMBER', ЕТЕКСТ: 'ISTEXT', ЕОШИБКА: 'ISERROR', ЕОШ: 'ISERR', ЕНД: 'ISNA', ЕЛОГИЧ: 'ISLOGICAL', ЕЧЁТН: 'ISEVEN', ЕНЕЧЁТ: 'ISODD', НД: 'NA',
  СЦЕП: 'CONCAT', СЦЕПИТЬ: 'CONCATENATE', ОБЪЕДИНИТЬ: 'TEXTJOIN', ДЛСТР: 'LEN', ПРОПИСН: 'UPPER', СТРОЧН: 'LOWER', ПРОПНАЧ: 'PROPER', СЖПРОБЕЛЫ: 'TRIM', ПЕЧСИМВ: 'CLEAN',
  ЛЕВСИМВ: 'LEFT', ПРАВСИМВ: 'RIGHT', ПСТР: 'MID', НАЙТИ: 'FIND', ПОИСК: 'SEARCH', ПОДСТАВИТЬ: 'SUBSTITUTE', ЗАМЕНИТЬ: 'REPLACE', ПОВТОР: 'REPT', СОВПАД: 'EXACT', ТЕКСТ: 'TEXT', ЗНАЧЕН: 'VALUE', СИМВОЛ: 'CHAR', КОДСИМВ: 'CODE', ФИКСИРОВАННЫЙ: 'FIXED',
  СЕГОДНЯ: 'TODAY', ТДАТА: 'NOW', ДАТА: 'DATE', ВРЕМЯ: 'TIME', ГОД: 'YEAR', МЕСЯЦ: 'MONTH', ДЕНЬ: 'DAY', ЧАС: 'HOUR', МИНУТЫ: 'MINUTE', СЕКУНДЫ: 'SECOND', ДЕНЬНЕД: 'WEEKDAY', НОМНЕДЕЛИ: 'WEEKNUM', ДАТАМЕС: 'EDATE', КОНМЕСЯЦА: 'EOMONTH', ДНИ: 'DAYS', РАЗНДАТ: 'DATEDIF', ЧИСТРАБДНИ: 'NETWORKDAYS', РАБДЕНЬ: 'WORKDAY', ДАТАЗНАЧ: 'DATEVALUE',
  ВПР: 'VLOOKUP', ГПР: 'HLOOKUP', ПРОСМОТРX: 'XLOOKUP', ПРОСМОТР: 'LOOKUP', ПОИСКПОЗ: 'MATCH', ИНДЕКС: 'INDEX', ЧСТРОК: 'ROWS', ЧИСЛСТОЛБ: 'COLUMNS', СТРОКА: 'ROW', СТОЛБЕЦ: 'COLUMN', ТРАНСП: 'TRANSPOSE', УНИК: 'UNIQUE', АДРЕС: 'ADDRESS', ПРОМЕЖУТОЧНЫЕ_ИТОГИ: 'SUBTOTAL', ПОСЛЕДОВ: 'SEQUENCE', ГИПЕРССЫЛКА: 'HYPERLINK',
  ПЛТ: 'PMT', БС: 'FV', ПС: 'PV', ЧПС: 'NPV', ВСД: 'IRR', КПЕР: 'NPER',
};
export const FN_HELP = {
  'Математические': [['SUM', 'СУММ(число1; [число2]; …)', 'Сумма чисел'], ['SUMIF', 'СУММЕСЛИ(диапазон; условие; [диапазон_суммирования])', 'Сумма по условию'], ['SUMIFS', 'СУММЕСЛИМН(диапазон_суммирования; диапазон1; условие1; …)', 'Сумма по нескольким условиям'], ['SUMPRODUCT', 'СУММПРОИЗВ(массив1; массив2; …)', 'Сумма произведений'], ['PRODUCT', 'ПРОИЗВЕД(число1; …)', 'Произведение'], ['ROUND', 'ОКРУГЛ(число; разряды)', 'Округление'], ['ROUNDUP', 'ОКРУГЛВВЕРХ(число; разряды)', 'Округление вверх'], ['ROUNDDOWN', 'ОКРУГЛВНИЗ(число; разряды)', 'Округление вниз'], ['ABS', 'ABS(число)', 'Модуль'], ['SQRT', 'КОРЕНЬ(число)', 'Квадратный корень'], ['POWER', 'СТЕПЕНЬ(число; степень)', 'Возведение в степень'], ['MOD', 'ОСТАТ(число; делитель)', 'Остаток от деления'], ['INT', 'ЦЕЛОЕ(число)', 'Целая часть'], ['RAND', 'СЛЧИС()', 'Случайное число 0…1'], ['RANDBETWEEN', 'СЛУЧМЕЖДУ(мин; макс)', 'Случайное целое'], ['PI', 'ПИ()', 'Число π']],
  'Статистические': [['AVERAGE', 'СРЗНАЧ(число1; …)', 'Среднее'], ['AVERAGEIF', 'СРЗНАЧЕСЛИ(диапазон; условие; [диапазон_усреднения])', 'Среднее по условию'], ['MIN', 'МИН(число1; …)', 'Минимум'], ['MAX', 'МАКС(число1; …)', 'Максимум'], ['COUNT', 'СЧЁТ(значение1; …)', 'Количество чисел'], ['COUNTA', 'СЧЁТЗ(значение1; …)', 'Количество непустых'], ['COUNTIF', 'СЧЁТЕСЛИ(диапазон; условие)', 'Количество по условию'], ['COUNTIFS', 'СЧЁТЕСЛИМН(диапазон1; условие1; …)', 'Количество по условиям'], ['MEDIAN', 'МЕДИАНА(число1; …)', 'Медиана'], ['STDEV', 'СТАНДОТКЛОН(число1; …)', 'Стандартное отклонение'], ['LARGE', 'НАИБОЛЬШИЙ(массив; k)', 'k-е наибольшее'], ['SMALL', 'НАИМЕНЬШИЙ(массив; k)', 'k-е наименьшее'], ['RANK', 'РАНГ(число; ссылка; [порядок])', 'Ранг числа']],
  'Логические': [['IF', 'ЕСЛИ(условие; если_истина; если_ложь)', 'Проверка условия'], ['IFS', 'ЕСЛИМН(условие1; значение1; …)', 'Несколько условий'], ['IFERROR', 'ЕСЛИОШИБКА(значение; значение_при_ошибке)', 'Перехват ошибок'], ['AND', 'И(условие1; …)', 'Все истинны'], ['OR', 'ИЛИ(условие1; …)', 'Хотя бы одно истинно'], ['NOT', 'НЕ(условие)', 'Отрицание'], ['SWITCH', 'ПЕРЕКЛЮЧ(выражение; значение1; результат1; …)', 'Выбор по значению']],
  'Текстовые': [['CONCAT', 'СЦЕП(текст1; …)', 'Объединение текста'], ['TEXTJOIN', 'ОБЪЕДИНИТЬ(разделитель; пропускать_пустые; текст1; …)', 'Объединение с разделителем'], ['LEN', 'ДЛСТР(текст)', 'Длина'], ['LEFT', 'ЛЕВСИМВ(текст; [число])', 'Символы слева'], ['RIGHT', 'ПРАВСИМВ(текст; [число])', 'Символы справа'], ['MID', 'ПСТР(текст; начало; число)', 'Часть строки'], ['UPPER', 'ПРОПИСН(текст)', 'ПРОПИСНЫЕ'], ['LOWER', 'СТРОЧН(текст)', 'строчные'], ['PROPER', 'ПРОПНАЧ(текст)', 'Каждое Слово С Прописной'], ['TRIM', 'СЖПРОБЕЛЫ(текст)', 'Убрать лишние пробелы'], ['SUBSTITUTE', 'ПОДСТАВИТЬ(текст; старый; новый; [номер])', 'Замена текста'], ['FIND', 'НАЙТИ(что; где; [начало])', 'Позиция подстроки'], ['TEXT', 'ТЕКСТ(значение; формат)', 'Число в текст по формату']],
  'Дата и время': [['TODAY', 'СЕГОДНЯ()', 'Текущая дата'], ['NOW', 'ТДАТА()', 'Текущие дата и время'], ['DATE', 'ДАТА(год; месяц; день)', 'Дата из частей'], ['YEAR', 'ГОД(дата)', 'Год'], ['MONTH', 'МЕСЯЦ(дата)', 'Месяц'], ['DAY', 'ДЕНЬ(дата)', 'День'], ['WEEKDAY', 'ДЕНЬНЕД(дата; [тип])', 'День недели'], ['EDATE', 'ДАТАМЕС(дата; месяцы)', 'Сдвиг на месяцы'], ['EOMONTH', 'КОНМЕСЯЦА(дата; месяцы)', 'Конец месяца'], ['DATEDIF', 'РАЗНДАТ(начало; конец; "D"|"M"|"Y")', 'Разница дат'], ['NETWORKDAYS', 'ЧИСТРАБДНИ(начало; конец)', 'Рабочие дни']],
  'Ссылки и поиск': [['VLOOKUP', 'ВПР(искомое; таблица; номер_столбца; [интервальный_просмотр])', 'Вертикальный поиск'], ['HLOOKUP', 'ГПР(искомое; таблица; номер_строки; [интервальный])', 'Горизонтальный поиск'], ['XLOOKUP', 'ПРОСМОТРX(искомое; где_искать; что_вернуть; [если_не_найдено])', 'Современный поиск'], ['INDEX', 'ИНДЕКС(массив; строка; [столбец])', 'Значение по позиции'], ['MATCH', 'ПОИСКПОЗ(искомое; массив; [тип])', 'Позиция значения'], ['ROW', 'СТРОКА([ссылка])', 'Номер строки'], ['COLUMN', 'СТОЛБЕЦ([ссылка])', 'Номер столбца']],
  'Финансовые': [['PMT', 'ПЛТ(ставка; кпер; пс; [бс]; [тип])', 'Платёж по кредиту'], ['FV', 'БС(ставка; кпер; плт; [пс]; [тип])', 'Будущая стоимость'], ['PV', 'ПС(ставка; кпер; плт; [бс]; [тип])', 'Приведённая стоимость'], ['NPV', 'ЧПС(ставка; значение1; …)', 'Чистая приведённая стоимость'], ['IRR', 'ВСД(значения; [предположение])', 'Внутренняя норма доходности']],
};
export const ALL_FN_NAMES = [...new Set([...Object.keys(FN), 'IF', 'IFERROR', 'IFNA', 'IFS', 'CHOOSE', 'ROW', 'COLUMN', 'ISBLANK', ...Object.keys(ALIAS)])].sort();
export { isErr, num as toNum, txt as toText };
