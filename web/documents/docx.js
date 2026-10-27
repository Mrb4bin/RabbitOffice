// DOCX import/export for Rabbit Documents
import { zipCreate, zipRead, xmlEsc as x, parseXml, dataUrlToBytes, bytesToB64 } from '../core/zip.js';

const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MM_TW = 56.6929; // mm -> twips

function cssColorToHex(c) {
  if (!c) return null;
  if (c.startsWith('#')) return c.length === 4 ? c.slice(1).split('').map((ch) => ch + ch).join('') : c.slice(1, 7);
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(c); if (!m) return null;
  if (m[4] !== undefined && +m[4] === 0) return null;
  return [m[1], m[2], m[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('');
}
const HL = { ffff00: 'yellow', '00ff00': 'green', '00ffff': 'cyan', ff00ff: 'magenta', '0000ff': 'blue', ff0000: 'red', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', c0c0c0: 'lightGray', '000000': 'black' };

/* ======================= EXPORT ======================= */
export async function exportDocx(editor, doc) {
  const media = []; const links = [];
  let body = ''; let revId = 0; const fnIds = new Map();

  const runProps = (st) => {
    let p = '';
    if (st.font) p += `<w:rFonts w:ascii="${x(st.font)}" w:hAnsi="${x(st.font)}" w:cs="${x(st.font)}"/>`;
    if (st.b) p += '<w:b/>'; if (st.i) p += '<w:i/>'; if (st.s) p += '<w:strike/>';
    if (st.color) p += `<w:color w:val="${st.color}"/>`;
    if (st.size) p += `<w:sz w:val="${Math.round(st.size * 2)}"/><w:szCs w:val="${Math.round(st.size * 2)}"/>`;
    if (st.hl) p += HL[st.hl] ? `<w:highlight w:val="${HL[st.hl]}"/>` : `<w:shd w:val="clear" w:color="auto" w:fill="${st.hl}"/>`;
    if (st.u) p += '<w:u w:val="single"/>';
    if (st.va) p += `<w:vertAlign w:val="${st.va}"/>`;
    return p ? `<w:rPr>${p}</w:rPr>` : '';
  };
  const textRun = (t, st) => (st.del ? `<w:r>${runProps(st)}<w:delText xml:space="preserve">${x(t)}</w:delText></w:r>` : `<w:r>${runProps(st)}<w:t xml:space="preserve">${x(t)}</w:t></w:r>`);

  function imageRun(img) {
    if (!img.src.startsWith('data:')) return '';
    const { mime, bytes } = dataUrlToBytes(img.src);
    const ext = mime.includes('png') ? 'png' : mime.includes('gif') ? 'gif' : 'jpeg';
    const id = 'rIdImg' + (media.length + 1);
    const name = `image${media.length + 1}.${ext}`;
    media.push({ id, name, bytes, ext });
    const w = Math.round((img.width || img.naturalWidth || 300) * 9525), hgt = Math.round((img.height || img.naturalHeight || 200) * 9525);
    const n = media.length;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${hgt}"/><wp:docPr id="${n}" name="Picture ${n}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${n}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${hgt}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  function inline(node, st) {
    let out = '';
    for (const n of node.childNodes) {
      if (n.nodeType === 3) { if (n.data) out += textRun(n.data.replace(/\u00a0/g, ' '), st); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName.toLowerCase();
      if (tag === 'br') { out += '<w:r><w:br/></w:r>'; continue; }
      if (tag === 'img') { out += imageRun(n); continue; }
      if (n.classList.contains('cmt-anchor')) { out += inline(n, st); continue; }
      if ((tag === 'ins' || tag === 'del') && n.classList.contains('trk')) {
        const a = `w:id="${++revId}" w:author="${x(n.dataset.a || 'Автор')}" w:date="${new Date(+n.dataset.t || Date.now()).toISOString().replace(/\.\d+Z$/, 'Z')}"`;
        out += tag === 'ins' ? `<w:ins ${a}>${inline(n, st)}</w:ins>` : `<w:del ${a}>${inline(n, { ...st, del: true })}</w:del>`;
        continue;
      }
      if (tag === 'sup' && n.classList.contains('fn-ref')) {
        let fid = fnIds.get(n.dataset.fn); if (!fid) { fid = fnIds.size + 1; fnIds.set(n.dataset.fn, fid); }
        out += `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${fid}"/></w:r>`;
        continue;
      }
      if (n.classList.contains('fn-num')) continue;
      const s = { ...st };
      if (tag === 'b' || tag === 'strong') s.b = true;
      if (tag === 'i' || tag === 'em') s.i = true;
      if (tag === 'u') s.u = true;
      if (tag === 's' || tag === 'strike' || tag === 'del') s.s = true;
      if (tag === 'sub') s.va = 'subscript';
      if (tag === 'sup') s.va = 'superscript';
      if (tag === 'code') s.font = 'Consolas';
      if (tag === 'font') { if (n.color) s.color = cssColorToHex(n.color); if (n.face) s.font = n.face; }
      const cs = n.style;
      if (cs.fontWeight === 'bold' || +cs.fontWeight >= 600) s.b = true;
      if (cs.fontStyle === 'italic') s.i = true;
      if (cs.textDecoration.includes('underline') || cs.textDecorationLine?.includes('underline')) s.u = true;
      if (cs.textDecoration.includes('line-through') || cs.textDecorationLine?.includes('line-through')) s.s = true;
      if (cs.color) s.color = cssColorToHex(cs.color);
      if (cs.backgroundColor) s.hl = cssColorToHex(cs.backgroundColor);
      if (cs.fontFamily) s.font = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (cs.fontSize) { const v = parseFloat(cs.fontSize); s.size = cs.fontSize.endsWith('pt') ? v : cs.fontSize.endsWith('px') ? v * 0.75 : s.size; }
      if (cs.verticalAlign === 'sub') s.va = 'subscript'; if (cs.verticalAlign === 'super') s.va = 'superscript';
      if (tag === 'a' && n.href) {
        const id = 'rIdLink' + (links.length + 1); links.push({ id, href: n.getAttribute('href') });
        out += `<w:hyperlink r:id="${id}">${inline(n, { ...s, u: true, color: s.color || '0563c1' })}</w:hyperlink>`;
        continue;
      }
      out += inline(n, s);
    }
    return out;
  }

  function pPr(el, extra = '') {
    let p = extra;
    const al = el.style?.textAlign || el.getAttribute?.('align');
    if (al) p += `<w:jc w:val="${{ left: 'left', center: 'center', right: 'right', justify: 'both' }[al] || 'left'}"/>`;
    const lh = parseFloat(el.style?.lineHeight);
    if (lh && !el.style.lineHeight.endsWith('px')) p += `<w:spacing w:line="${Math.round(lh * 240)}" w:lineRule="auto"/>`;
    const ml = parseFloat(el.style?.marginLeft || el.style?.paddingLeft);
    if (ml) p += `<w:ind w:left="${Math.round(ml * 15)}"/>`;
    return p ? `<w:pPr>${reorderPPr(p)}</w:pPr>` : '';
  }
  // w:pPr children must follow schema order: pStyle, numPr, pBdr, spacing, ind, jc
  function reorderPPr(p) {
    const order = ['pStyle', 'numPr', 'pBdr', 'spacing', 'ind', 'jc'];
    const parts = p.match(/<w:(\w+)[^>]*?(\/>|>[\s\S]*?<\/w:\1>)/g) || [];
    return parts.sort((a, b) => order.indexOf(/<w:(\w+)/.exec(a)[1]) - order.indexOf(/<w:(\w+)/.exec(b)[1])).join('');
  }

  function block(el, ctx = {}) {
    if (el.nodeType === 3) { if (el.data.trim()) body += `<w:p>${textRun(el.data, {})}</w:p>`; return; }
    if (el.nodeType !== 1) return;
    const tag = el.tagName.toLowerCase();
    if (el.classList.contains('page-break')) { body += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'; return; }
    if (el.classList.contains('footnotes')) return; // exported as real Word footnotes
    if (tag === 'p' && el.classList.contains('dropcap')) { body += `<w:p>${pPr(el)}${inline(el, {})}</w:p>`; return; }
    if (el.classList.contains('toc')) { for (const a of el.querySelectorAll('a')) body += `<w:p><w:pPr><w:ind w:left="${(+a.dataset.level - 1) * 360}"/></w:pPr>${textRun(a.textContent, {})}</w:p>`; return; }
    if (/^h[1-6]$/.test(tag)) { const style = el.classList.contains('doc-title') ? 'Title' : 'Heading' + Math.min(3, +tag[1]); body += `<w:p>${pPr(el, `<w:pStyle w:val="${style}"/>`)}${inline(el, {})}</w:p>`; return; }
    if (tag === 'p' || tag === 'div') {
      if ([...el.children].some((c) => /^(P|DIV|TABLE|UL|OL|H\d|BLOCKQUOTE|PRE)$/.test(c.tagName))) { for (const c of el.childNodes) block(c, ctx); return; }
      body += `<w:p>${pPr(el, ctx.style ? `<w:pStyle w:val="${ctx.style}"/>` : '')}${inline(el, {})}</w:p>`; return;
    }
    if (tag === 'blockquote') { const kids = [...el.childNodes]; if (kids.some((c) => c.nodeType === 1 && /^(P|DIV)$/.test(c.tagName))) kids.forEach((c) => block(c, { style: 'Quote' })); else body += `<w:p>${pPr(el, '<w:pStyle w:val="Quote"/>')}${inline(el, {})}</w:p>`; return; }
    if (tag === 'pre') { for (const line of el.textContent.split('\n')) body += `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr>${textRun(line, {})}</w:p>`; return; }
    if (tag === 'hr') { body += '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>'; return; }
    if (tag === 'ul' || tag === 'ol') {
      const numId = tag === 'ol' ? 2 : 1; const lvl = ctx.lvl || 0;
      for (const li of el.children) {
        if (li.tagName !== 'LI') continue;
        const nested = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
        const clone = li.cloneNode(true); clone.querySelectorAll('ul,ol').forEach((n) => n.remove());
        const checked = el.classList.contains('checklist') ? (li.classList.contains('done') ? '☑ ' : '☐ ') : '';
        body += `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>${checked ? '' : `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${numId}"/></w:numPr>`}</w:pPr>${checked ? textRun(checked, {}) : ''}${inline(clone, {})}</w:p>`;
        nested.forEach((n) => block(n, { lvl: lvl + 1 }));
      }
      return;
    }
    if (tag === 'table') {
      const rows = [...el.querySelectorAll(':scope > tbody > tr, :scope > thead > tr, :scope > tr')];
      const cols = Math.max(1, ...rows.map((r) => [...r.cells].reduce((a, c) => a + (c.colSpan || 1), 0)));
      const tw = Math.round(9000 / cols);
      body += `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="999999"/>`).join('')}</w:tblBorders></w:tblPr><w:tblGrid>${'<w:gridCol w:w="' + tw + '"/>'.repeat(cols)}</w:tblGrid>`;
      for (const r of rows) {
        body += '<w:tr>';
        for (const c of r.cells) {
          const bg = cssColorToHex(c.style.backgroundColor);
          body += `<w:tc><w:tcPr><w:tcW w:w="${tw * (c.colSpan || 1)}" w:type="dxa"/>${c.colSpan > 1 ? `<w:gridSpan w:val="${c.colSpan}"/>` : ''}${bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ''}</w:tcPr>`;
          const save = body; body = '';
          const hasBlocks = [...c.children].some((k) => /^(P|DIV|UL|OL|H\d|TABLE)$/.test(k.tagName));
          if (hasBlocks) c.childNodes.forEach((k) => block(k)); else body += `<w:p>${pPr(c)}${inline(c, c.tagName === 'TH' ? { b: true } : {})}</w:p>`;
          if (!body.includes('<w:p')) body += '<w:p/>';
          body = save + body + '</w:tc>';
        }
        body += '</w:tr>';
      }
      body += '</w:tbl><w:p/>';
      return;
    }
    if (tag === 'img') { body += `<w:p>${imageRun(el)}</w:p>`; return; }
    if (tag === 'figure') { for (const c of el.childNodes) block(c); return; }
    body += `<w:p>${inline(el, {})}</w:p>`;
  }
  for (const n of editor.childNodes) block(n);

  const pg = doc.page; const W = Math.round(pg.w * MM_TW), Hh = Math.round(pg.h * MM_TW);
  const [mt, mr, mb, ml] = pg.margins.map((m) => Math.round(m * MM_TW));
  const hasHdr = !!(doc.header?.text), hasFtr = !!(doc.footer?.text || doc.footer?.pageNumbers);
  const sect = `<w:sectPr>${hasHdr ? '<w:headerReference w:type="default" r:id="rIdHdr"/>' : ''}${hasFtr ? '<w:footerReference w:type="default" r:id="rIdFtr"/>' : ''}<w:pgSz w:w="${pg.landscape ? Hh : W}" w:h="${pg.landscape ? W : Hh}"${pg.landscape ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${mt}" w:right="${mr}" w:bottom="${mb}" w:left="${ml}" w:header="567" w:footer="567" w:gutter="0"/>${pg.columns > 1 ? `<w:cols w:num="${pg.columns}" w:space="708"/>` : ''}</w:sectPr>`;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}${sect}</w:body></w:document>`;
  let footnotesXml = null;
  if (fnIds.size) {
    const notes = [...fnIds].map(([key, fid]) => {
      const p = editor.querySelector(`p.fn[data-fn="${key}"]`); const c = p ? p.cloneNode(true) : document.createElement('p'); c.querySelector('.fn-num')?.remove();
      return `<w:footnote w:id="${fid}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>${inline(c, {})}</w:p></w:footnote>`;
    }).join('');
    footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes ${NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>${notes}</w:footnotes>`;
  }
  const hf = (tag, part) => {
    const jc = { left: 'left', center: 'center', right: 'right' }[part.align || 'center'];
    const pn = part.pageNumbers ? `${part.text ? textRun('  ', {}) : ''}<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` : '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} ${NS}><w:p><w:pPr><w:jc w:val="${jc}"/></w:pPr>${part.text ? textRun(part.text, { size: 9, color: '666666' }) : ''}${pn}</w:p></w:${tag}>`;
  };
  const font = doc.style?.font || 'Calibri'; const size = (doc.style?.size || 11) * 2;
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${x(font)}" w:hAnsi="${x(font)}" w:cs="${x(font)}" w:eastAsia="${x(font)}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:sz w:val="56"/><w:color w:val="1F2A44"/></w:rPr></w:style>
${[1, 2, 3].map((n) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="${360 - n * 60}" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${[0, 36, 30, 26][n]}"/><w:color w:val="${['', '2F62C9', '2F62C9', '1F2A44'][n]}"/></w:rPr></w:style>`).join('')}
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/><w:shd w:val="clear" w:color="auto" w:fill="F3F4F6"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="40"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>`;
  const lvls = (fmt) => Array.from({ length: 9 }, (_, i) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt === 'bullet' ? 'bullet' : ['decimal', 'lowerLetter', 'lowerRoman'][i % 3]}"/><w:lvlText w:val="${fmt === 'bullet' ? ['•', '◦', '▪'][i % 3] : `%${i + 1}.`}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + i * 360}" w:hanging="360"/></w:pPr>${fmt === 'bullet' ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' : ''}</w:lvl>`).join('');
  const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS_W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls('bullet')}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls('decimal')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="${R}/numbering" Target="numbering.xml"/>
${hasHdr ? `<Relationship Id="rIdHdr" Type="${R}/header" Target="header1.xml"/>` : ''}${hasFtr ? `<Relationship Id="rIdFtr" Type="${R}/footer" Target="footer1.xml"/>` : ''}${footnotesXml ? `<Relationship Id="rIdFn" Type="${R}/footnotes" Target="footnotes.xml"/>` : ''}
${media.map((m) => `<Relationship Id="${m.id}" Type="${R}/image" Target="media/${m.name}"/>`).join('')}
${links.map((l) => `<Relationship Id="${l.id}" Type="${R}/hyperlink" Target="${x(l.href)}" TargetMode="External"/>`).join('')}</Relationships>`;
  const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
${footnotesXml ? '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>' : ''}${hasHdr ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${hasFtr ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${x(doc.title || '')}</dc:title><dc:creator>${x(doc.author || 'Rabbit Documents')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  const files = [
    { name: '[Content_Types].xml', data: types }, { name: '_rels/.rels', data: rootRels }, { name: 'docProps/core.xml', data: core },
    { name: 'word/document.xml', data: documentXml }, { name: 'word/styles.xml', data: stylesXml }, { name: 'word/numbering.xml', data: numberingXml },
    { name: 'word/_rels/document.xml.rels', data: docRels },
    ...media.map((m) => ({ name: 'word/media/' + m.name, data: m.bytes })),
  ];
  if (footnotesXml) files.push({ name: 'word/footnotes.xml', data: footnotesXml });
  if (hasHdr) files.push({ name: 'word/header1.xml', data: hf('hdr', doc.header) });
  if (hasFtr) files.push({ name: 'word/footer1.xml', data: hf('ftr', doc.footer) });
  return zipCreate(files);
}

/* ======================= IMPORT ======================= */
const qn = (el, name) => [...el.childNodes].filter((n) => n.nodeType === 1 && n.localName === name);
const q1 = (el, name) => qn(el, name)[0];
const attr = (el, name) => el?.getAttributeNS(NS_W, name) ?? el?.getAttribute('w:' + name);
const he = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export async function importDocx(bytes) {
  const z = await zipRead(bytes);
  const docXml = parseXml(await z.get('word/document.xml').text());
  const rels = {};
  if (z.get('word/_rels/document.xml.rels')) {
    const rx = parseXml(await z.get('word/_rels/document.xml.rels').text());
    for (const r of rx.getElementsByTagName('Relationship')) rels[r.getAttribute('Id')] = { target: r.getAttribute('Target'), type: r.getAttribute('Type'), mode: r.getAttribute('TargetMode') };
  }
  const styleMap = {};
  if (z.get('word/styles.xml')) {
    const sx = parseXml(await z.get('word/styles.xml').text());
    for (const s of sx.getElementsByTagNameNS(NS_W, 'style')) {
      const id = attr(s, 'styleId'); const name = (attr(s.getElementsByTagNameNS(NS_W, 'name')[0], 'val') || '').toLowerCase();
      const ol = s.getElementsByTagNameNS(NS_W, 'outlineLvl')[0];
      if (name === 'title') styleMap[id] = 'title';
      else if (/^heading (\d)/.test(name)) styleMap[id] = 'h' + Math.min(3, +/^heading (\d)/.exec(name)[1]);
      else if (ol) styleMap[id] = 'h' + Math.min(3, +attr(ol, 'val') + 1);
      else if (name.includes('quote')) styleMap[id] = 'quote';
    }
  }
  const numFmt = {};
  if (z.get('word/numbering.xml')) {
    const nx = parseXml(await z.get('word/numbering.xml').text());
    const abs = {};
    for (const a of nx.getElementsByTagNameNS(NS_W, 'abstractNum')) { const l0 = a.getElementsByTagNameNS(NS_W, 'lvl')[0]; abs[attr(a, 'abstractNumId')] = attr(l0?.getElementsByTagNameNS(NS_W, 'numFmt')[0], 'val'); }
    for (const n of nx.getElementsByTagNameNS(NS_W, 'num')) numFmt[attr(n, 'numId')] = abs[attr(n.getElementsByTagNameNS(NS_W, 'abstractNumId')[0], 'val')];
  }
  const imgCache = {}; const usedFootnotes = [];
  const imageFor = async (rid) => {
    if (imgCache[rid]) return imgCache[rid];
    const r = rels[rid]; if (!r) return '';
    const path = 'word/' + r.target.replace(/^\/?word\//, '').replace(/^\.\//, '');
    const f = z.get(path); if (!f) return '';
    const ext = path.split('.').pop().toLowerCase();
    const mime = ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/jpeg';
    return (imgCache[rid] = `data:${mime};base64,${bytesToB64(await f.bytes())}`);
  };

  async function runs(p) {
    let html = '';
    for (const n of p.childNodes) {
      if (n.nodeType !== 1) continue;
      if (n.localName === 'hyperlink') { const r = rels[n.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || n.getAttribute('r:id')]; html += `<a href="${he(r?.target || '#')}">${await runs(n)}</a>`; continue; }
      if (n.localName === 'ins' || n.localName === 'del') { const a = he(attr(n, 'author') || ''); const t = Date.parse(attr(n, 'date') || '') || ''; html += `<${n.localName} class="trk" data-a="${a}" data-t="${t}" title="${a}">${await runs(n)}</${n.localName}>`; continue; }
      if (n.localName === 'smartTag' || n.localName === 'sdt' || n.localName === 'sdtContent' || n.localName === 'fldSimple') { html += await runs(n.localName === 'sdt' ? (q1(n, 'sdtContent') || n) : n); continue; }
      if (n.localName !== 'r') continue;
      const rp = q1(n, 'rPr'); let style = ''; let open = '', close = '';
      if (rp) {
        const has = (t) => { const e = q1(rp, t); return e && attr(e, 'val') !== '0' && attr(e, 'val') !== 'false' && attr(e, 'val') !== 'none'; };
        if (has('b')) { open += '<b>'; close = '</b>' + close; }
        if (has('i')) { open += '<i>'; close = '</i>' + close; }
        if (has('u')) { open += '<u>'; close = '</u>' + close; }
        if (has('strike')) { open += '<s>'; close = '</s>' + close; }
        const va = attr(q1(rp, 'vertAlign'), 'val'); if (va === 'superscript') { open += '<sup>'; close = '</sup>' + close; } if (va === 'subscript') { open += '<sub>'; close = '</sub>' + close; }
        const col = attr(q1(rp, 'color'), 'val'); if (col && col !== 'auto') style += `color:#${col};`;
        const sz = attr(q1(rp, 'sz'), 'val'); if (sz) style += `font-size:${+sz / 2}pt;`;
        const fn = attr(q1(rp, 'rFonts'), 'ascii'); if (fn) style += `font-family:'${fn}';`;
        const hl = attr(q1(rp, 'highlight'), 'val'); if (hl && hl !== 'none') style += `background-color:${hl.replace(/^dark/, 'dark').toLowerCase()};`;
        const shd = attr(q1(rp, 'shd'), 'fill'); if (shd && shd !== 'auto') style += `background-color:#${shd};`;
      }
      let inner = '';
      for (const c of n.childNodes) {
        if (c.nodeType !== 1) continue;
        if (c.localName === 't' || c.localName === 'delText') inner += he(c.textContent);
        else if (c.localName === 'footnoteReference') { const fid = attr(c, 'id'); usedFootnotes.push(fid); inner += `<sup class="fn-ref" contenteditable="false" data-fn="f${fid}">*</sup>`; }
        else if (c.localName === 'tab') inner += '&emsp;';
        else if (c.localName === 'br') inner += attr(c, 'type') === 'page' ? '\u0000PB\u0000' : '<br>';
        else if (c.localName === 'drawing' || c.localName === 'pict') {
          const blip = c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'blip')[0] || c.getElementsByTagName('v:imagedata')[0];
          const rid = blip && (blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || blip.getAttribute('r:embed') || blip.getAttribute('r:id'));
          const ext = c.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', 'extent')[0];
          const w = ext ? Math.round(+ext.getAttribute('cx') / 9525) : null;
          if (rid) inner += `<img src="${await imageFor(rid)}" ${w ? `style="width:${w}px"` : ''}>`;
        }
      }
      if (!inner) continue; // e.g. the footnoteRef mark inside a footnote
      if (inner.includes('class="fn-ref"')) { html += inner; continue; } // reference mark carries its own superscript
      html += style ? `<span style="${style}">${open}${inner}${close}</span>` : open + inner + close;
    }
    return html;
  }

  async function para(p) {
    const pp = q1(p, 'pPr');
    const sid = attr(q1(pp || p, 'pStyle'), 'val');
    let kind = styleMap[sid] || (sid && /^heading(\d)/i.test(sid) ? 'h' + Math.min(3, +/(\d)/.exec(sid)[1]) : sid === 'Title' ? 'title' : null);
    const jc = attr(q1(pp || p, 'jc'), 'val');
    let style = '';
    if (jc) style += `text-align:${{ both: 'justify', center: 'center', right: 'right', end: 'right' }[jc] || 'left'};`;
    const sp = pp && q1(pp, 'spacing'); const line = attr(sp, 'line');
    if (line && (attr(sp, 'lineRule') || 'auto') === 'auto') style += `line-height:${(+line / 240).toFixed(2)};`;
    const ind = pp && q1(pp, 'ind'); const left = attr(ind, 'left') || attr(ind, 'start');
    const numPr = pp && q1(pp, 'numPr');
    if (left && !numPr && +left > 0) style += `margin-left:${Math.round(+left / 15)}px;`;
    let inner = await runs(p);
    const pageBreak = inner.includes('\u0000PB\u0000'); inner = inner.replace(/\u0000PB\u0000/g, '');
    const st = style ? ` style="${style}"` : '';
    let html;
    if (kind === 'title') html = `<h1 class="doc-title"${st}>${inner || '<br>'}</h1>`;
    else if (kind && kind[0] === 'h') html = `<${kind}${st}>${inner || '<br>'}</${kind}>`;
    else if (kind === 'quote') html = `<blockquote${st}>${inner || '<br>'}</blockquote>`;
    else html = `<p${st}>${inner || '<br>'}</p>`;
    return { html, pageBreak, list: numPr ? { lvl: +attr(q1(numPr, 'ilvl'), 'val') || 0, type: numFmt[attr(q1(numPr, 'numId'), 'val')] === 'bullet' ? 'ul' : 'ol', inner: inner || '<br>' } : null };
  }

  async function table(t) {
    let html = '<table class="doc-table"><tbody>';
    for (const tr of qn(t, 'tr')) {
      html += '<tr>';
      for (const tc of qn(tr, 'tc')) {
        const pr = q1(tc, 'tcPr'); const span = attr(pr && q1(pr, 'gridSpan'), 'val'); const fill = attr(pr && q1(pr, 'shd'), 'fill');
        html += `<td${span ? ` colspan="${span}"` : ''}${fill && fill !== 'auto' ? ` style="background-color:#${fill}"` : ''}>${await blocks(tc)}</td>`;
      }
      html += '</tr>';
    }
    return html + '</tbody></table>';
  }

  async function blocks(container) {
    let html = ''; let listStack = [];
    const closeLists = (to = 0) => { while (listStack.length > to) html += `</li></${listStack.pop()}>`; };
    for (const n of container.childNodes) {
      if (n.nodeType !== 1) continue;
      if (n.localName === 'p') {
        const r = await para(n);
        if (r.list) {
          const want = r.list.lvl + 1;
          if (listStack.length === want && listStack[want - 1] !== r.list.type) closeLists(want - 1);
          if (listStack.length < want) { while (listStack.length < want) { html += `<${r.list.type}><li>`; listStack.push(r.list.type); } html += r.list.inner; }
          else { closeLists(want); html += `</li><li>${r.list.inner}`; }
          continue;
        }
        closeLists();
        html += r.html;
        if (r.pageBreak) html += '<div class="page-break" contenteditable="false"></div>';
      } else if (n.localName === 'tbl') { closeLists(); html += await table(n); }
      else if (n.localName === 'sdt') { closeLists(); html += await blocks(q1(n, 'sdtContent') || n); }
    }
    closeLists();
    return html;
  }

  const body = docXml.getElementsByTagNameNS(NS_W, 'body')[0];
  let html = await blocks(body);
  if (usedFootnotes.length && z.get('word/footnotes.xml')) {
    const fx = parseXml(await z.get('word/footnotes.xml').text());
    const byId = new Map([...fx.getElementsByTagNameNS(NS_W, 'footnote')].map((f) => [attr(f, 'id'), f]));
    let sec = '<section class="footnotes">';
    for (const id of usedFootnotes) { const f = byId.get(id); let inner = ''; if (f) for (const p of qn(f, 'p')) inner += await runs(p); sec += `<p class="fn" data-fn="f${id}"><span class="fn-num" contenteditable="false"></span> ${inner}</p>`; }
    html += sec + '</section>';
  }
  const page = {};
  const sect = body.getElementsByTagNameNS(NS_W, 'sectPr'); const s = sect[sect.length - 1];
  if (s) {
    const sz = s.getElementsByTagNameNS(NS_W, 'pgSz')[0]; const mar = s.getElementsByTagNameNS(NS_W, 'pgMar')[0];
    if (sz) { const w = +attr(sz, 'w') / MM_TW, hh = +attr(sz, 'h') / MM_TW; page.landscape = w > hh; page.w = Math.min(w, hh); page.h = Math.max(w, hh); }
    if (mar) page.margins = ['top', 'right', 'bottom', 'left'].map((k) => Math.round(+attr(mar, k) / MM_TW));
  }
  return { html, page };
}
