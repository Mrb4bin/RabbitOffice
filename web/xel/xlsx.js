// XLSX import/export for Rabbit Xel
import { zipCreate, zipRead, xmlEsc as x, parseXml } from '../core/zip.js';
import { MAX_COLS, colName, colIndex, refName, parseRange, rangeName, shiftFormula, isErr, ALIAS } from './formula.js';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NUMFMT = { number: [4, null], percent: [10, null], date: [14, null], time: [20, null], datetime: [22, null], sci: [11, null], text: [49, null] };
const hex = (c) => (c ? 'FF' + c.replace('#', '').toUpperCase().padEnd(6, '0').slice(0, 6) : null);
const RU2EN = Object.fromEntries(Object.entries(ALIAS));

function toExcelFormula(f) {
  // uppercase Russian aliases -> English names; ';' -> ','
  let out = ''; let inStr = false;
  for (let i = 1; i < f.length; i++) { const ch = f[i]; if (ch === '"') inStr = !inStr; out += !inStr && ch === ';' ? ',' : ch; }
  return out.replace(/([A-Za-zА-Яа-яЁё_][\wА-Яа-яЁё.]*)(?=\s*\()/g, (n) => RU2EN[n.toUpperCase()] || n.toUpperCase());
}

export async function exportXlsx(book, engine) {
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>']; const fontIdx = new Map([['', 0]]);
  const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']; const fillIdx = new Map([['', 0]]);
  const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']; const borderIdx = new Map([['', 0]]);
  const numFmts = []; const numFmtIdx = new Map();
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']; const xfIdx = new Map([['', 0]]);
  const strings = []; const strIdx = new Map();
  const sst = (s) => { if (!strIdx.has(s)) { strIdx.set(s, strings.length); strings.push(s); } return strIdx.get(s); };
  function styleId(s) {
    if (!s) return 0; const key = JSON.stringify(s); if (xfIdx.has(key)) return xfIdx.get(key);
    const fk = JSON.stringify([s.b, s.i, s.u, s.s, s.fc, s.ff, s.fs]);
    if (!fontIdx.has(fk)) { fontIdx.set(fk, fonts.length); fonts.push(`<font>${s.b ? '<b/>' : ''}${s.i ? '<i/>' : ''}${s.s ? '<strike/>' : ''}${s.u ? '<u/>' : ''}<sz val="${s.fs || 11}"/>${s.fc ? `<color rgb="${hex(s.fc)}"/>` : ''}<name val="${x(s.ff || 'Calibri')}"/></font>`); }
    if (s.bg && !fillIdx.has(s.bg)) { fillIdx.set(s.bg, fills.length); fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${hex(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`); }
    const bk = s.bd ? JSON.stringify(s.bd) : '';
    if (bk && !borderIdx.has(bk)) { const st = (w) => (w === 3 ? 'double' : w === 2 ? 'medium' : 'thin'); const side = (n, w) => (w ? `<${n} style="${st(w)}"><color auto="1"/></${n}>` : `<${n}/>`); borderIdx.set(bk, borders.length); borders.push(`<border>${side('left', s.bd.l)}${side('right', s.bd.r)}${side('top', s.bd.t)}${side('bottom', s.bd.b)}<diagonal/></border>`); }
    let nf = 0;
    if (s.fmt === 'currency' || s.dec != null && s.fmt !== 'percent' && s.fmt !== 'sci' || s.fmt === 'datelong') {
      const dec = s.dec ?? (s.fmt === 'general' || !s.fmt ? 0 : 2); const z = dec ? '.' + '0'.repeat(dec) : '';
      const code = s.fmt === 'currency' ? (s.cur === '₽' || !s.cur ? `#,##0${z}\\ "₽"` : `"${s.cur}"#,##0${z}`) : s.fmt === 'datelong' ? 'd mmmm yyyy' : s.fmt === 'number' ? `#,##0${z}` : `0${z}`;
      if (!numFmtIdx.has(code)) { numFmtIdx.set(code, 164 + numFmts.length); numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${x(code)}"/>`); }
      nf = numFmtIdx.get(code);
    } else if (s.fmt === 'percent') nf = s.dec ? (() => { const code = '0.' + '0'.repeat(s.dec) + '%'; if (!numFmtIdx.has(code)) { numFmtIdx.set(code, 164 + numFmts.length); numFmts.push(`<numFmt numFmtId="${164 + numFmts.length}" formatCode="${code}"/>`); } return numFmtIdx.get(code); })() : 9;
    else if (NUMFMT[s.fmt]) nf = NUMFMT[s.fmt][0];
    const al = s.al || s.va || s.wr ? `<alignment${s.al ? ` horizontal="${s.al}"` : ''}${s.va ? ` vertical="${s.va === 'middle' ? 'center' : s.va}"` : ''}${s.wr ? ' wrapText="1"' : ''}/>` : '';
    const xf = `<xf numFmtId="${nf}" fontId="${fontIdx.get(fk)}" fillId="${s.bg ? fillIdx.get(s.bg) : 0}" borderId="${bk ? borderIdx.get(bk) : 0}" xfId="0"${nf ? ' applyNumberFormat="1"' : ''} applyFont="1"${s.bg ? ' applyFill="1"' : ''}${bk ? ' applyBorder="1"' : ''}${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`;
    xfIdx.set(key, xfs.length); xfs.push(xf); return xfs.length - 1;
  }
  const sheetXml = book.sheets.map((s, si) => {
    const rows = new Map();
    for (const [k, cell] of s.cells) { const r = Math.floor(k / MAX_COLS), c = k % MAX_COLS; if (!rows.has(r)) rows.set(r, []); rows.get(r).push([c, cell]); }
    let data = '';
    for (const r of [...rows.keys()].sort((a, b) => a - b)) {
      const ht = s.rowH[r]; data += `<row r="${r + 1}"${ht != null ? ` ht="${(ht * 0.75).toFixed(2)}" customHeight="1"${ht === 0 ? ' hidden="1"' : ''}` : ''}>`;
      for (const [c, cell] of rows.get(r).sort((a, b) => a[0] - b[0])) {
        const ref = refName(r, c); const st = styleId(cell.s); const sa = st ? ` s="${st}"` : '';
        const v = cell.v;
        if (v == null || v === '') { data += `<c r="${ref}"${sa}/>`; continue; }
        if (typeof v === 'string' && v[0] === '=') {
          const res = engine.value(si, r, c);
          const f = `<f>${x(toExcelFormula(v))}</f>`;
          if (typeof res === 'number') data += `<c r="${ref}"${sa}>${f}<v>${res}</v></c>`;
          else if (typeof res === 'boolean') data += `<c r="${ref}"${sa} t="b">${f}<v>${res ? 1 : 0}</v></c>`;
          else if (isErr(res)) data += `<c r="${ref}"${sa} t="e">${f}<v>${x(res.e)}</v></c>`;
          else data += `<c r="${ref}"${sa} t="str">${f}<v>${x(res ?? '')}</v></c>`;
        } else if (typeof v === 'number') data += `<c r="${ref}"${sa}><v>${v}</v></c>`;
        else if (typeof v === 'boolean') data += `<c r="${ref}"${sa} t="b"><v>${v ? 1 : 0}</v></c>`;
        else data += `<c r="${ref}"${sa} t="s"><v>${sst(String(v))}</v></c>`;
      }
      data += '</row>';
    }
    const colEntries = Object.entries(s.colW).map(([c, w]) => [+c, w]).sort((a, b) => a[0] - b[0]);
    const cols = colEntries.length ? `<cols>${colEntries.map(([c, w]) => `<col min="${c + 1}" max="${c + 1}" width="${(w / 7).toFixed(2)}" customWidth="1"${w === 0 ? ' hidden="1"' : ''}/>`).join('')}</cols>` : '';
    const fr = s.freeze || { r: 0, c: 0 };
    const pane = fr.r || fr.c ? `<pane${fr.c ? ` xSplit="${fr.c}"` : ''}${fr.r ? ` ySplit="${fr.r}"` : ''} topLeftCell="${refName(fr.r, fr.c)}" activePane="${fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight'}" state="frozen"/>` : '';
    const merges = s.merges.length ? `<mergeCells count="${s.merges.length}">${s.merges.map((m) => `<mergeCell ref="${rangeName(m)}"/>`).join('')}</mergeCells>` : '';
    const af = s.filter ? `<autoFilter ref="${rangeName({ r1: s.filter.r, c1: s.filter.c1, r2: s.filter.r, c2: s.filter.c2 })}"/>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetViews><sheetView workbookViewId="0"${si === book.active ? ' tabSelected="1"' : ''}${s.gridlines === false ? ' showGridLines="0"' : ''}>${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="16.5"/>${cols}<sheetData>${data}</sheetData>${af}${merges}</worksheet>`;
  });
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const files = [
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${book.sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="${NS}" xmlns:r="${R}"><bookViews><workbookView activeTab="${book.active}"/></bookViews><sheets>${book.sheets.map((s, i) => `<sheet name="${x(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${book.sheets.some((s) => s.filter) ? `<definedNames>${book.sheets.map((s, i) => s.filter ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">'${x(s.name.replace(/'/g, "''"))}'!${rangeName({ r1: s.filter.r, c1: s.filter.c1, r2: s.filter.r, c2: s.filter.c2 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}</definedName>` : '').join('')}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${book.sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${R}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdSS" Type="${R}/sharedStrings" Target="sharedStrings.xml"/></Relationships>` },
    ...sheetXml.map((d, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: d })),
  ];
  files.push({ name: 'xl/sharedStrings.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${NS}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t xml:space="preserve">${x(s)}</t></si>`).join('')}</sst>` });
  files.push({ name: 'xl/styles.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` });
  return zipCreate(files);
}

/* ======================= IMPORT ======================= */
const kids = (el, name) => (el ? [...el.childNodes].filter((n) => n.nodeType === 1 && n.localName === name) : []);
const kid = (el, name) => kids(el, name)[0];
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
function color(el) {
  if (!el) return null; const rgb = el.getAttribute('rgb'); if (rgb) return '#' + rgb.slice(-6).toLowerCase();
  const th = el.getAttribute('theme'); if (th != null && THEME[+th]) return '#' + THEME[+th].toLowerCase();
  return null;
}
function fmtFromCode(id, code) {
  id = +id;
  if (id === 0) return null; if ([1, 3].includes(id)) return { fmt: 'number', dec: 0 }; if ([2, 4].includes(id)) return { fmt: 'number', dec: 2 };
  if (id === 9) return { fmt: 'percent', dec: 0 }; if (id === 10) return { fmt: 'percent', dec: 2 }; if (id === 11) return { fmt: 'sci' };
  if (id >= 14 && id <= 17) return { fmt: 'date' }; if (id >= 18 && id <= 21) return { fmt: 'time' }; if (id === 22) return { fmt: 'datetime' }; if (id === 49) return { fmt: 'text' };
  if (!code) return null;
  const c = code.replace(/"[^"]*"|\\./g, (m) => (/[₽$€£¥]|руб/.test(m) ? '¤' : ''));
  const dec = ((/0\.(0+)/.exec(c) || [])[1] || '').length;
  if (/[₽$€£¥¤]/.test(code) && !/[dmy]/i.test(c)) return { fmt: 'currency', dec, cur: (/[₽$€£¥]/.exec(code) || ['₽'])[0] };
  if (c.includes('%')) return { fmt: 'percent', dec };
  if (/[dy]/i.test(c) && /h/i.test(c)) return { fmt: 'datetime' };
  if (/[dy]/i.test(c) || /mmm/i.test(c)) return { fmt: 'date' };
  if (/h/i.test(c)) return { fmt: 'time' };
  if (/E\+/i.test(c)) return { fmt: 'sci' };
  if (/[#0]/.test(c)) return { fmt: 'number', dec };
  return null;
}
export async function importXlsx(bytes) {
  const z = await zipRead(bytes);
  const read = async (p) => (z.get(p) ? parseXml(await z.get(p).text()) : null);
  const wb = await read('xl/workbook.xml');
  const rels = {}; const rx = await read('xl/_rels/workbook.xml.rels');
  if (rx) for (const r of rx.getElementsByTagName('Relationship')) rels[r.getAttribute('Id')] = r.getAttribute('Target').replace(/^\/?xl\//, '');
  const strings = [];
  const ss = await read('xl/sharedStrings.xml');
  if (ss) for (const si of ss.getElementsByTagNameNS(NS, 'si')) strings.push([...si.getElementsByTagNameNS(NS, 't')].map((t) => t.textContent).join(''));
  const xfStyles = [];
  const st = await read('xl/styles.xml');
  if (st) {
    const nfCodes = {}; for (const nf of st.getElementsByTagNameNS(NS, 'numFmt')) nfCodes[nf.getAttribute('numFmtId')] = nf.getAttribute('formatCode');
    const fonts = kids(kid(st.documentElement, 'fonts'), 'font').map((f) => ({ b: !!kid(f, 'b') && kid(f, 'b').getAttribute('val') !== '0', i: !!kid(f, 'i') && kid(f, 'i').getAttribute('val') !== '0', u: !!kid(f, 'u'), s: !!kid(f, 'strike'), fs: +(kid(f, 'sz')?.getAttribute('val') || 11), ff: kid(f, 'name')?.getAttribute('val'), fc: color(kid(f, 'color')) }));
    const fills = kids(kid(st.documentElement, 'fills'), 'fill').map((f) => { const p = kid(f, 'patternFill'); return p && p.getAttribute('patternType') === 'solid' ? color(kid(p, 'fgColor')) : null; });
    const bw = (el) => { const s = el?.getAttribute('style'); return !s ? 0 : s === 'double' ? 3 : /medium|thick/.test(s) ? 2 : 1; };
    const borders = kids(kid(st.documentElement, 'borders'), 'border').map((b) => { const o = { l: bw(kid(b, 'left')), r: bw(kid(b, 'right')), t: bw(kid(b, 'top')), b: bw(kid(b, 'bottom')) }; return o.l || o.r || o.t || o.b ? o : null; });
    for (const xf of kids(kid(st.documentElement, 'cellXfs'), 'xf')) {
      const s = {}; const f = fonts[+xf.getAttribute('fontId') || 0] || {};
      if (f.b) s.b = true; if (f.i) s.i = true; if (f.u) s.u = true; if (f.s) s.s = true; if (f.fs && f.fs !== 11) s.fs = f.fs; if (f.ff && f.ff !== 'Calibri') s.ff = f.ff; if (f.fc && f.fc !== '#000000') s.fc = f.fc;
      const bg = fills[+xf.getAttribute('fillId') || 0]; if (bg && bg !== '#ffffff') s.bg = bg;
      const bd = borders[+xf.getAttribute('borderId') || 0]; if (bd) s.bd = bd;
      const nf = fmtFromCode(xf.getAttribute('numFmtId') || 0, nfCodes[xf.getAttribute('numFmtId')]); if (nf) Object.assign(s, nf);
      const al = kid(xf, 'alignment');
      if (al) { const hh = al.getAttribute('horizontal'); if (['left', 'center', 'right'].includes(hh)) s.al = hh; const v = al.getAttribute('vertical'); if (v === 'center') s.va = 'middle'; else if (v === 'top') s.va = 'top'; if (al.getAttribute('wrapText') === '1') s.wr = true; }
      xfStyles.push(Object.keys(s).length ? s : null);
    }
  }
  const sheets = [];
  for (const sEl of wb.getElementsByTagNameNS(NS, 'sheet')) {
    const rid = sEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || sEl.getAttribute('r:id');
    const doc = await read('xl/' + (rels[rid] || ''));
    const sheet = { name: sEl.getAttribute('name'), cells: [], colW: {}, rowH: {}, merges: [], freeze: { r: 0, c: 0 } };
    if (!doc) { sheets.push(sheet); continue; }
    const root = doc.documentElement;
    for (const col of kids(kid(root, 'cols'), 'col')) { const w = Math.round(+col.getAttribute('width') * 7); const hidden = col.getAttribute('hidden') === '1'; for (let c = +col.getAttribute('min') - 1; c <= Math.min(+col.getAttribute('max') - 1, MAX_COLS - 1); c++) sheet.colW[c] = hidden ? 0 : w; }
    const sv = kid(kid(root, 'sheetViews'), 'sheetView'); const pane = kid(sv, 'pane');
    if (pane && pane.getAttribute('state') === 'frozen') sheet.freeze = { r: +(pane.getAttribute('ySplit') || 0), c: +(pane.getAttribute('xSplit') || 0) };
    if (sv?.getAttribute('showGridLines') === '0') sheet.gridlines = false;
    const shared = {};
    for (const row of kids(kid(root, 'sheetData'), 'row')) {
      const r = +row.getAttribute('r') - 1;
      if (row.getAttribute('customHeight') === '1' && row.getAttribute('ht')) sheet.rowH[r] = Math.round(+row.getAttribute('ht') / 0.75);
      if (row.getAttribute('hidden') === '1') sheet.rowH[r] = 0;
      for (const c of kids(row, 'c')) {
        const ref = c.getAttribute('r'); const m = /^([A-Z]+)(\d+)$/.exec(ref); if (!m) continue;
        const cr = +m[2] - 1, cc = colIndex(m[1]); if (cc >= MAX_COLS) continue;
        const t = c.getAttribute('t'); const vEl = kid(c, 'v'); const fEl = kid(c, 'f'); const s = xfStyles[+c.getAttribute('s') || 0];
        let v = null;
        if (fEl) {
          let f = fEl.textContent;
          const si = fEl.getAttribute('si');
          if (fEl.getAttribute('t') === 'shared' && si != null) { if (f) shared[si] = { f, r: cr, c: cc }; else if (shared[si]) f = shiftFormula('=' + shared[si].f, cr - shared[si].r, cc - shared[si].c).slice(1); }
          if (f) v = '=' + f.replace(/_xlfn\./g, '');
        }
        if (v == null && vEl) {
          const raw = vEl.textContent;
          if (t === 's') v = strings[+raw] ?? ''; else if (t === 'b') v = raw === '1'; else if (t === 'str' || t === 'e') v = raw; else v = +raw;
        } else if (v == null && t === 'inlineStr') v = [...c.getElementsByTagNameNS(NS, 't')].map((x2) => x2.textContent).join('');
        if (v == null && !s) continue;
        sheet.cells.push([cr, cc, v, s || 0, 0]);
      }
    }
    for (const mc of kids(kid(root, 'mergeCells'), 'mergeCell')) { const g = parseRange(mc.getAttribute('ref')); if (g) sheet.merges.push(g); }
    const af = kid(root, 'autoFilter'); if (af) { const g = parseRange(af.getAttribute('ref')); if (g) sheet.filter = { r: g.r1, c1: g.c1, c2: g.c2, crit: {} }; }
    sheets.push(sheet);
  }
  return { sheets };
}
