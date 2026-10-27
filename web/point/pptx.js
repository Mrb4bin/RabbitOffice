// PPTX export/import for Rabbit Point (core features: text, shapes, images, tables, backgrounds)
import { zipCreate, zipRead, xmlEsc as x, parseXml, dataUrlToBytes, bytesToB64 } from '../core/zip.js';
import { SHAPES, THEMES, makeSlide, uidp } from './render.js';
import { renderChart, chartDataFromGrid } from '../core/charts.js';
import { parseGrid } from './render.js';

const E = 9525;
const PRST = { rect: 'rect', round: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtriangle: 'rtTriangle', diamond: 'diamond', pentagon: 'pentagon', hexagon: 'hexagon', octagon: 'octagon', star: 'star5', arrow: 'rightArrow', larrow: 'leftArrow', uarrow: 'upArrow', chevron: 'chevron', bubble: 'wedgeRoundRectCallout', plus: 'mathPlus', heart: 'heart', cloud: 'cloud' };
const hex6 = (c) => { if (!c) return null; const m = /#([0-9a-f]{6})/i.exec(c); if (m) return m[1].toUpperCase(); const r = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); return r ? [r[1], r[2], r[3]].map((v) => (+v).toString(16).padStart(2, '0')).join('').toUpperCase() : null; };
const fill = (c) => (c && c !== 'none' && hex6(c) ? `<a:solidFill><a:srgbClr val="${hex6(c)}"/></a:solidFill>` : '<a:noFill/>');

function paragraphs(html, t, th) {
  const d = document.createElement('div'); d.innerHTML = html || '';
  const out = [];
  const runs = (node, st) => { let r = ''; for (const n of node.childNodes) { if (n.nodeType === 3) { if (n.data) r += `<a:r><a:rPr lang="ru-RU" sz="${Math.round((st.size || 24) * 75)}"${st.b ? ' b="1"' : ''}${st.i ? ' i="1"' : ''}${st.u ? ' u="sng"' : ''} dirty="0">${fill(st.color)}<a:latin typeface="${x(st.font)}"/></a:rPr><a:t>${x(n.data)}</a:t></a:r>`; continue; } if (n.nodeType !== 1 || /^(UL|OL|LI|DIV|P)$/.test(n.tagName)) continue; const s = { ...st }; const tg = n.tagName; if (tg === 'B' || tg === 'STRONG') s.b = true; if (tg === 'I' || tg === 'EM') s.i = true; if (tg === 'U') s.u = true; if (n.style?.color) s.color = n.style.color; if (n.style?.fontSize) s.size = parseFloat(n.style.fontSize); if (tg === 'BR') { r += '<a:br/>'; continue; } r += runs(n, s); } return r; };
  const base = { size: t.size || 24, b: t.bold, i: t.italic, color: t.color || th.fg, font: t.font || th.body };
  const algn = { left: 'l', center: 'ctr', right: 'r', justify: 'just' }[t.align || 'left'];
  const walk = (node, bullet) => {
    let inline = document.createElement('span');
    const flush = () => { if (inline.childNodes.length) { out.push(`<a:p><a:pPr algn="${algn}"${bullet ? ` marL="342900" indent="-342900"><a:buChar char="•"/></a:pPr>` : '><a:buNone/></a:pPr>'}${runs(inline, base)}</a:p>`); inline = document.createElement('span'); } };
    for (const n of [...node.childNodes]) {
      if (n.nodeType === 1 && (n.tagName === 'UL' || n.tagName === 'OL')) { flush(); for (const li of n.children) { const sp = document.createElement('span'); sp.append(...[...li.childNodes].map((c) => c.cloneNode(true))); out.push(`<a:p><a:pPr algn="${algn}" marL="342900" indent="-342900">${n.tagName === 'OL' ? '<a:buAutoNum type="arabicPeriod"/>' : '<a:buChar char="•"/>'}</a:pPr>${runs(sp, base)}</a:p>`); } }
      else if (n.nodeType === 1 && (n.tagName === 'DIV' || n.tagName === 'P')) { flush(); walk(n, bullet); }
      else if (n.nodeType === 1 && n.tagName === 'BR') { flush(); }
      else inline.append(n.cloneNode(true));
    }
    flush();
  };
  walk(d, false);
  return out.length ? out.join('') : `<a:p><a:endParaRPr lang="ru-RU"/></a:p>`;
}

async function svgToPng(svg, w, h) {
  const img = new Image(); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg))); await img.decode();
  const c = document.createElement('canvas'); c.width = w * 2; c.height = h * 2; const g = c.getContext('2d'); g.scale(2, 2); g.drawImage(img, 0, 0, w, h);
  return dataUrlToBytes(c.toDataURL('image/png')).bytes;
}

export async function exportPptx(pres) {
  const th = pres.theme; const files = []; const media = [];
  const slidesXml = [];
  for (const [si, s] of pres.slides.entries()) {
    const rels = []; let id = 2; let shapes = '';
    const xfrm = (e) => `<a:xfrm${e.rot ? ` rot="${Math.round(e.rot * 60000)}"` : ''}${e.flipH ? ' flipH="1"' : ''}><a:off x="${Math.round(e.x * E)}" y="${Math.round(e.y * E)}"/><a:ext cx="${Math.round(Math.max(1, e.w) * E)}" cy="${Math.round(Math.max(1, e.h) * E)}"/></a:xfrm>`;
    const pic = async (e, bytes, ext) => { const n = `image${media.length + 1}.${ext}`; media.push({ n, bytes }); const rid = 'rId' + (rels.length + 2); rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${n}"/>`); shapes += `<p:pic><p:nvPicPr><p:cNvPr id="${id++}" name="Picture"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(e)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`; };
    for (const e of s.els) {
      if (e.type === 'image' && e.src?.startsWith('data:')) { const { mime, bytes } = dataUrlToBytes(e.src); await pic(e, bytes, mime.includes('png') ? 'png' : mime.includes('gif') ? 'gif' : 'jpeg'); continue; }
      if (e.type === 'chart') { const svg = renderChart({ type: e.chart.type, title: e.chart.title, ...chartDataFromGrid(parseGrid(e.chart.data)) }, e.w, e.h, { bg: '#ffffff' }); await pic(e, await svgToPng(svg, e.w, e.h), 'png'); continue; }
      if (e.type === 'table') {
        const rows = e.table.rows; const cw = Math.round((e.w * E) / rows[0].length);
        shapes += `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id++}" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round(e.x * E)}" y="${Math.round(e.y * E)}"/><a:ext cx="${Math.round(e.w * E)}" cy="${Math.round(e.h * E)}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="${e.table.header ? 1 : 0}" bandRow="${e.table.band ? 1 : 0}"/><a:tblGrid>${rows[0].map(() => `<a:gridCol w="${cw}"/>`).join('')}</a:tblGrid>${rows.map((r, ri) => `<a:tr h="${Math.round((e.h * E) / rows.length)}">${r.map((c) => `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs(c, { size: e.size || 18, bold: ri === 0 && e.table.header, color: ri === 0 && e.table.header ? '#ffffff' : e.color }, th)}</a:txBody><a:tcPr>${ri === 0 && e.table.header ? fill(th.accent) : ''}</a:tcPr></a:tc>`).join('')}</a:tr>`).join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
        continue;
      }
      const sh = SHAPES[e.shape];
      if (e.type === 'shape' && sh?.line) { shapes += `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id++}" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm({ ...e, y: e.y + e.h / 2, h: 0 })}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="${Math.round((e.strokeW || 3) * 12700)}">${fill(e.stroke || th.fg)}${sh.arrow ? '<a:tailEnd type="triangle"/>' : ''}</a:ln></p:spPr></p:cxnSp>`; continue; }
      const geom = e.type === 'shape' ? PRST[e.shape] || 'rect' : 'rect';
      const t = e.text || {};
      const sp = e.type === 'shape' ? `${fill(e.fill ?? th.accent)}${e.strokeW && e.stroke && e.stroke !== 'none' ? `<a:ln w="${e.strokeW * 12700}">${fill(e.stroke)}</a:ln>` : '<a:ln><a:noFill/></a:ln>'}` : '<a:noFill/>';
      const anchor = { top: 't', middle: 'ctr', bottom: 'b' }[t.valign || (e.type === 'shape' ? 'middle' : 'top')];
      const tt = { ...t, color: t.color || (e.type === 'shape' ? '#ffffff' : th.fg), font: t.font || (e.ph === 'title' ? th.title : th.body) };
      shapes += `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${e.type === 'text' ? 'TextBox' : 'Shape'} ${id}"/><p:cNvSpPr${e.type === 'text' ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(e)}<a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${sp}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="${12 * E}" rIns="${12 * E}" anchor="${anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>${paragraphs(t.html, tt, th)}</p:txBody></p:sp>`;
    }
    const bgc = s.bg?.color || (s.bg?.gradient || th.bg);
    const bgHex = hex6(bgc) || 'FFFFFF';
    let bg = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${bgHex}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    if (s.bg?.image) { const { mime, bytes } = dataUrlToBytes(s.bg.image); const n = `image${media.length + 1}.${mime.includes('png') ? 'png' : 'jpeg'}`; media.push({ n, bytes }); const rid = 'rId' + (rels.length + 2); rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${n}"/>`); bg = `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`; }
    const tr = s.transition?.type; const trx = tr && tr !== 'none' ? `<p:transition spd="${(s.transition.dur || 0.5) > 0.7 ? 'slow' : 'med'}">${{ fade: '<p:fade/>', push: '<p:push dir="l"/>', wipe: '<p:wipe dir="r"/>', zoom: '<p:zoom/>', cover: '<p:cover dir="u"/>', flip: '<p:fade/>' }[tr] || '<p:fade/>'}</p:transition>` : '';
    slidesXml.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${s.hidden ? ' show="0"' : ''}><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${trx}</p:sld>`);
    files.push({ name: `ppt/slides/_rels/slide${si + 1}.xml.rels`, data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${rels.join('')}</Relationships>` });
  }
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  const n = pres.slides.length;
  const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Rabbit"><a:themeElements><a:clrScheme name="Rabbit"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="${hex6(th.fg) || '262A33'}"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="${hex6(th.accent) || 'D2552D'}"/></a:accent1><a:accent2><a:srgbClr val="${hex6(th.accent2) || '2F62C9'}"/></a:accent2><a:accent3><a:srgbClr val="1D8A52"/></a:accent3><a:accent4><a:srgbClr val="7B4FD6"/></a:accent4><a:accent5><a:srgbClr val="C49A1A"/></a:accent5><a:accent6><a:srgbClr val="1AA3B8"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Rabbit"><a:majorFont><a:latin typeface="${x(th.title)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${x(th.body)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Rabbit"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
  const emptyTree = '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>';
  files.push(
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>${pres.slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('')}</Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>` },
    { name: 'ppt/presentation.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${pres.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 10}"/>`).join('')}</p:sldIdLst><p:sldSz cx="${pres.w * E}" cy="${pres.h * E}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>` },
    { name: 'ppt/_rels/presentation.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="theme/theme1.xml"/>${pres.slides.map((_, i) => `<Relationship Id="rId${i + 10}" Type="${R}/slide" Target="slides/slide${i + 1}.xml"/>`).join('')}</Relationships>` },
    { name: 'ppt/slideMasters/slideMaster1.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${NS}>${emptyTree}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>` },
    { name: 'ppt/slideMasters/_rels/slideMaster1.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${R}/theme" Target="../theme/theme1.xml"/></Relationships>` },
    { name: 'ppt/slideLayouts/slideLayout1.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${NS} type="blank">${emptyTree}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>` },
    { name: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>` },
    { name: 'ppt/theme/theme1.xml', data: theme },
    ...slidesXml.map((d, i) => ({ name: `ppt/slides/slide${i + 1}.xml`, data: d })),
    ...media.map((m) => ({ name: 'ppt/media/' + m.n, data: m.bytes })),
  );
  return zipCreate(files);
}

/* basic PPTX import: positions, text, shapes, pictures, solid backgrounds */
export async function importPptx(bytes) {
  const z = await zipRead(bytes); const A = 'http://schemas.openxmlformats.org/drawingml/2006/main', P = 'http://schemas.openxmlformats.org/presentationml/2006/main', RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const read = async (p) => (z.get(p) ? parseXml(await z.get(p).text()) : null);
  const pr = await read('ppt/presentation.xml'); const sz = pr.getElementsByTagNameNS(P, 'sldSz')[0];
  const cx = +sz.getAttribute('cx'), cy = +sz.getAttribute('cy'); const k = 1280 / (cx / E);
  const pres = { w: 1280, h: Math.round((cy / E) * k), theme: { ...THEMES[0] }, slides: [] };
  const prels = await read('ppt/_rels/presentation.xml.rels'); const relMap = {}; for (const r of prels.getElementsByTagName('Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');
  for (const sid of pr.getElementsByTagNameNS(P, 'sldId')) {
    const path = 'ppt/' + relMap[sid.getAttributeNS(RNS, 'id')]; const doc = await read(path); if (!doc) continue;
    const srels = await read(path.replace('slides/', 'slides/_rels/') + '.rels'); const sr = {}; if (srels) for (const r of srels.getElementsByTagName('Relationship')) sr[r.getAttribute('Id')] = r.getAttribute('Target');
    const slide = makeSlide(pres, 'blank');
    const bgc = doc.getElementsByTagNameNS(P, 'bg')[0]?.getElementsByTagNameNS(A, 'srgbClr')[0]; if (bgc) slide.bg = { color: '#' + bgc.getAttribute('val') };
    const geom = (el) => { const off = el.getElementsByTagNameNS(A, 'off')[0], ext = el.getElementsByTagNameNS(A, 'ext')[0]; if (!off || !ext) return null; return { x: Math.round((+off.getAttribute('x') / E) * k), y: Math.round((+off.getAttribute('y') / E) * k), w: Math.round((+ext.getAttribute('cx') / E) * k), h: Math.round((+ext.getAttribute('cy') / E) * k) }; };
    for (const sp of doc.getElementsByTagNameNS(P, 'sp')) {
      const ph = sp.getElementsByTagNameNS(P, 'ph')[0]?.getAttribute('type');
      const g = geom(sp) || (ph === 'title' || ph === 'ctrTitle' ? { x: 80, y: 50, w: 1120, h: 110 } : { x: 80, y: 180, w: 1120, h: 480 });
      let size = 24; const html = [...sp.getElementsByTagNameNS(A, 'p')].map((p) => { const t = [...p.getElementsByTagNameNS(A, 'r')].map((r) => { const rp = r.getElementsByTagNameNS(A, 'rPr')[0]; if (rp?.getAttribute('sz')) size = Math.round((+rp.getAttribute('sz') / 100) * (96 / 72) * k); let s = (r.getElementsByTagNameNS(A, 't')[0]?.textContent || '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); if (rp?.getAttribute('b') === '1') s = `<b>${s}</b>`; if (rp?.getAttribute('i') === '1') s = `<i>${s}</i>`; const c = rp?.getElementsByTagNameNS(A, 'srgbClr')[0]; if (c) s = `<span style="color:#${c.getAttribute('val')}">${s}</span>`; return s; }).join(''); return `<div>${t || '<br>'}</div>`; }).join('');
      const prst = sp.getElementsByTagNameNS(A, 'prstGeom')[0]?.getAttribute('prst');
      const fillC = sp.getElementsByTagNameNS(P, 'spPr')[0]?.getElementsByTagNameNS(A, 'solidFill')[0]?.getElementsByTagNameNS(A, 'srgbClr')[0];
      const shape = Object.entries(PRST).find(([, v]) => v === prst)?.[0];
      if (fillC && shape) slide.els.push({ id: uidp(), type: 'shape', shape, ...g, fill: '#' + fillC.getAttribute('val'), text: { html, size, align: 'center', valign: 'middle', color: '#ffffff' } });
      else slide.els.push({ id: uidp(), type: 'text', ...g, text: { html, size: ph && /title/i.test(ph) ? Math.max(size, 36) : size, align: 'left', bold: !!ph && /title/i.test(ph) } });
    }
    for (const pic of doc.getElementsByTagNameNS(P, 'pic')) {
      const g = geom(pic); const rid = pic.getElementsByTagNameNS(A, 'blip')[0]?.getAttributeNS(RNS, 'embed'); const target = sr[rid]; if (!g || !target) continue;
      const f = z.get('ppt/' + target.replace('../', '')); if (!f) continue;
      const ext = target.split('.').pop().toLowerCase(); slide.els.push({ id: uidp(), type: 'image', ...g, src: `data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,${bytesToB64(await f.bytes())}` });
    }
    pres.slides.push(slide);
  }
  return pres;
}
