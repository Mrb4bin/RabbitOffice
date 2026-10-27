// Rabbit Point — slide model helpers and renderer (shared by editor, thumbnails, show, export)
import { renderChart, chartDataFromGrid } from '../core/charts.js';

export const SHAPES = {
  rect: { n: 'Прямоугольник', d: 'M0 0H100V100H0Z' },
  round: { n: 'Скруглённый прямоугольник', d: 'M10 0H90Q100 0 100 10V90Q100 100 90 100H10Q0 100 0 90V10Q0 0 10 0Z' },
  ellipse: { n: 'Овал', d: 'M50 0A50 50 0 1 1 50 100A50 50 0 1 1 50 0Z' },
  triangle: { n: 'Треугольник', d: 'M50 0L100 100H0Z' },
  rtriangle: { n: 'Прямоугольный треугольник', d: 'M0 0L100 100H0Z' },
  diamond: { n: 'Ромб', d: 'M50 0L100 50L50 100L0 50Z' },
  pentagon: { n: 'Пятиугольник', d: 'M50 0L100 38L81 100H19L0 38Z' },
  hexagon: { n: 'Шестиугольник', d: 'M25 0H75L100 50L75 100H25L0 50Z' },
  octagon: { n: 'Восьмиугольник', d: 'M30 0H70L100 30V70L70 100H30L0 70V30Z' },
  star: { n: 'Звезда', d: 'M50 0L61 35H98L68 57L79 92L50 70L21 92L32 57L2 35H39Z' },
  arrow: { n: 'Стрелка вправо', d: 'M0 28H62V0L100 50L62 100V72H0Z' },
  larrow: { n: 'Стрелка влево', d: 'M100 28H38V0L0 50L38 100V72H100Z' },
  uarrow: { n: 'Стрелка вверх', d: 'M28 100V38H0L50 0L100 38H72V100Z' },
  chevron: { n: 'Шеврон', d: 'M0 0H72L100 50L72 100H0L28 50Z' },
  bubble: { n: 'Выноска', d: 'M8 0H92Q100 0 100 8V67Q100 75 92 75H42L22 100V75H8Q0 75 0 67V8Q0 0 8 0Z' },
  plus: { n: 'Плюс', d: 'M35 0H65V35H100V65H65V100H35V65H0V35H35Z' },
  heart: { n: 'Сердце', d: 'M50 96C22 74 0 58 0 32A25 25 0 0 1 50 20A25 25 0 0 1 100 32C100 58 78 74 50 96Z' },
  cloud: { n: 'Облако', d: 'M25 90A22 22 0 0 1 20 47A25 25 0 0 1 55 22A26 26 0 0 1 92 45A22 22 0 0 1 80 90Z' },
  line: { n: 'Линия', d: 'M0 50L100 50', line: true },
  arrowline: { n: 'Стрелка', d: 'M0 50L100 50', line: true, arrow: true },
};

export const THEMES = [
  { id: 'office', name: 'Стандартная', bg: '#ffffff', fg: '#262a33', accent: '#d2552d', accent2: '#2f62c9', title: 'Segoe UI', body: 'Segoe UI' },
  { id: 'midnight', name: 'Полночь', bg: 'linear-gradient(160deg,#151a2d,#0d1020)', fg: '#eef1fb', accent: '#6aa7ff', accent2: '#f7b955', title: 'Segoe UI', body: 'Segoe UI' },
  { id: 'pitch', name: 'Энергия', bg: 'linear-gradient(135deg,#ff8a5b,#d2552d)', fg: '#ffffff', accent: '#ffe0d1', accent2: '#2b1a12', title: 'Trebuchet MS', body: 'Segoe UI' },
  { id: 'minimal', name: 'Минимализм', bg: '#f5f1ea', fg: '#1f1f1f', accent: '#1d8a52', accent2: '#c2410c', title: 'Georgia', body: 'Segoe UI' },
  { id: 'ocean', name: 'Океан', bg: 'linear-gradient(160deg,#0f4c81,#1aa3b8)', fg: '#ffffff', accent: '#9ff0ff', accent2: '#ffd166', title: 'Segoe UI', body: 'Segoe UI' },
  { id: 'forest', name: 'Лес', bg: '#14281d', fg: '#e8f3e8', accent: '#7bd389', accent2: '#e9c46a', title: 'Georgia', body: 'Segoe UI' },
  { id: 'graphite', name: 'Графит', bg: '#26282c', fg: '#f2f2f2', accent: '#ffc83d', accent2: '#6bc5ff', title: 'Segoe UI', body: 'Segoe UI' },
  { id: 'lavender', name: 'Лаванда', bg: 'linear-gradient(160deg,#f3eefe,#e6ddfb)', fg: '#2b2240', accent: '#7b4fd6', accent2: '#e0679a', title: 'Segoe UI', body: 'Segoe UI' },
];

export const LAYOUTS = {
  title: { n: 'Титульный слайд', ph: [['title', 140, 220, 1000, 170, 60, 'center', 'Заголовок слайда'], ['subtitle', 240, 410, 800, 80, 26, 'center', 'Подзаголовок']] },
  content: { n: 'Заголовок и объект', ph: [['title', 80, 50, 1120, 110, 40, 'left', 'Заголовок слайда'], ['body', 80, 180, 1120, 480, 24, 'left', 'Текст слайда']] },
  two: { n: 'Два объекта', ph: [['title', 80, 50, 1120, 110, 40, 'left', 'Заголовок слайда'], ['body', 80, 180, 540, 480, 22, 'left', 'Текст'], ['body', 660, 180, 540, 480, 22, 'left', 'Текст']] },
  section: { n: 'Заголовок раздела', ph: [['title', 100, 260, 1080, 140, 54, 'left', 'Заголовок раздела'], ['subtitle', 100, 410, 1080, 70, 24, 'left', 'Описание раздела']] },
  titleonly: { n: 'Только заголовок', ph: [['title', 80, 50, 1120, 110, 40, 'left', 'Заголовок слайда']] },
  quote: { n: 'Цитата', ph: [['title', 160, 200, 960, 240, 40, 'center', '«Цитата»'], ['subtitle', 160, 470, 960, 60, 22, 'center', '— Автор']] },
  blank: { n: 'Пустой слайд', ph: [] },
};

export const uidp = () => Math.random().toString(36).slice(2, 9);

export function makeSlide(pres, layout = 'content') {
  const L = LAYOUTS[layout] || LAYOUTS.content;
  return {
    id: uidp(), layout, bg: null, transition: { type: 'fade', dur: 0.5 }, notes: '', hidden: false,
    els: L.ph.map(([ph, x, y, w, hh, size, align, hint]) => ({
      id: uidp(), type: 'text', ph, hint, x, y, w, h: hh,
      text: { html: '', size, align, valign: ph === 'body' ? 'top' : ph === 'title' && layout !== 'title' && layout !== 'section' ? 'bottom' : 'middle', bold: ph === 'title', list: ph === 'body' },
    })),
  };
}

export function slideBg(pres, slide) {
  const b = slide.bg; const th = pres.theme;
  if (!b) return th.bg;
  if (b.image) return `center/cover no-repeat url("${b.image}")`;
  if (b.gradient) return b.gradient;
  return b.color || th.bg;
}

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export function renderSlide(pres, slide, { edit = false, showAnim = false } = {}) {
  const th = pres.theme;
  const s = document.createElement('div');
  s.className = 'slide' + (edit ? ' edit-stage' : '');
  s.style.width = pres.w + 'px'; s.style.height = pres.h + 'px';
  s.style.background = slideBg(pres, slide);
  s.style.setProperty('--fg', th.fg); s.style.setProperty('--accent-c', th.accent);
  s.style.setProperty('--body-font', `"${th.body}", system-ui, sans-serif`);
  s.dataset.id = slide.id;
  const animated = slide.els.filter((e) => e.anim && e.anim.type !== 'none').sort((a, b) => (a.anim.order || 0) - (b.anim.order || 0));
  for (const el of slide.els) {
    const n = renderEl(pres, el, { edit });
    if (edit && showAnim && el.anim && el.anim.type !== 'none') { const b = document.createElement('div'); b.className = 'anim-badge'; b.textContent = animated.indexOf(el) + 1; n.append(b); }
    s.append(n);
  }
  return s;
}

export function renderEl(pres, el, { edit = false } = {}) {
  const th = pres.theme;
  const n = document.createElement('div');
  n.className = `el t-${el.type}`; n.dataset.id = el.id;
  Object.assign(n.style, { left: el.x + 'px', top: el.y + 'px', width: el.w + 'px', height: el.h + 'px', transform: `rotate(${el.rot || 0}deg)${el.flipH ? ' scaleX(-1)' : ''}${el.flipV ? ' scaleY(-1)' : ''}`, opacity: el.opacity ?? 1 });
  if (el.shadow) n.style.filter = 'drop-shadow(0 6px 14px rgba(0,0,0,.28))';
  if (el.type === 'shape') {
    const sh = SHAPES[el.shape] || SHAPES.rect;
    const gid = 'g' + el.id;
    const fill = sh.line ? 'none' : el.grad ? `url(#${gid})` : el.fill ?? th.accent; const stroke = el.stroke ?? (sh.line ? th.fg : 'none'); const sw = el.strokeW ?? (sh.line ? 3 : 0);
    const mid = 'm' + el.id;
    n.innerHTML = `<svg class="shape-svg" viewBox="0 0 100 100" preserveAspectRatio="none">${el.grad && !sh.line ? `<defs><linearGradient id="${gid}" x1="0" y1="0" x2="${el.grad[2] === 'h' ? 1 : 0}" y2="${el.grad[2] === 'h' ? 0 : 1}"><stop offset="0" stop-color="${el.grad[0]}"/><stop offset="1" stop-color="${el.grad[1]}"/></linearGradient></defs>` : ''}${sh.arrow ? `<defs><marker id="${mid}" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="${4}" markerHeight="${4}" orient="auto-start-reverse"><path d="M0 0L10 5L0 10z" fill="${stroke}"/></marker></defs>` : ''}<path d="${sh.d}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" vector-effect="non-scaling-stroke" stroke-linejoin="round" ${sh.arrow ? `marker-end="url(#${mid})"` : ''}/></svg>`;
  }
  if (el.type === 'image') { const img = document.createElement('img'); img.src = el.src; img.draggable = false; if (el.radius) img.style.borderRadius = el.radius + 'px'; if (el.stroke && el.strokeW) img.style.outline = `${el.strokeW}px solid ${el.stroke}`; if (el.filter) img.style.filter = IMG_FILTERS[el.filter] || ''; if (el.fit) img.style.objectFit = el.fit; n.append(img); }
  if (el.type === 'media') {
    if (el.kind === 'audio' && edit) { n.innerHTML = '<div style="position:absolute;inset:0;display:grid;place-items:center;font-size:' + Math.min(el.w, el.h) * 0.6 + 'px">🔊</div>'; }
    else { const m = document.createElement(el.kind === 'audio' ? 'audio' : 'video'); m.src = el.src; m.controls = !edit; m.preload = 'metadata'; m.loop = !!el.loop; m.style.cssText = `width:100%;height:100%;display:block;${el.kind === 'audio' ? '' : 'object-fit:contain;background:#000;'}${edit ? 'pointer-events:none;' : ''}`; if (!edit && el.autoplay) m.dataset.autoplay = '1'; n.append(m); }
  }
  if (el.type === 'table') {
    const t = el.table; const tb = document.createElement('table'); tb.className = (t.header ? 'hdr ' : '') + (t.band ? 'band' : '');
    tb.innerHTML = t.rows.map((r) => `<tr>${r.map((c) => `<td style="color:${el.color || th.fg};font-size:${el.size || 18}px">${c}</td>`).join('')}</tr>`).join('');
    if (edit) tb.querySelectorAll('td').forEach((td) => (td.dataset.edit = '1'));
    n.append(tb);
  }
  if (el.type === 'chart') {
    const d = chartDataFromGrid(parseGrid(el.chart.data));
    const h = document.createElement('div'); h.className = 'chart-holder';
    const dark = isDarkBg(th, el);
    h.innerHTML = renderChart({ type: el.chart.type, title: el.chart.title, ...d }, el.w, el.h, { text: dark ? '#e8e8e8' : '#444', grid: dark ? 'rgba(255,255,255,.18)' : 'rgba(0,0,0,.12)', font: th.body });
    n.append(h);
  }
  if (el.type === 'text' || (el.type === 'shape' && !SHAPES[el.shape]?.line)) {
    const t = el.text || {};
    const wrap = document.createElement('div'); wrap.className = 'el-text';
    wrap.style.justifyContent = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[t.valign || (el.type === 'shape' ? 'middle' : 'top')];
    const tx = document.createElement('div'); tx.className = 'tx';
    tx.innerHTML = t.html || '';
    tx.dataset.ph = edit && el.ph ? el.hint || '' : '';
    const font = el.ph === 'title' || t.heading ? th.title : th.body;
    Object.assign(tx.style, { fontSize: (t.size || 24) + 'px', color: t.color || (el.type === 'shape' ? '#ffffff' : th.fg), textAlign: t.align || (el.type === 'shape' ? 'center' : 'left'), fontFamily: `"${t.font || font}", system-ui, sans-serif`, fontWeight: t.bold ? '650' : '', fontStyle: t.italic ? 'italic' : '', lineHeight: t.lh || '' });
    if (el.ph === 'title' && !t.color) tx.style.color = th.fg;
    if (el.ph === 'subtitle' && !t.color) tx.style.color = `color-mix(in srgb, ${th.fg} 72%, transparent)`;
    wrap.append(tx); n.append(wrap);
    if (edit && el.ph && !t.html) n.classList.add('placeholder-frame');
  }
  return n;
}
export const IMG_FILTERS = { none: '', grayscale: 'grayscale(1)', sepia: 'sepia(.85)', bright: 'brightness(1.2) contrast(1.05)', dark: 'brightness(.75)', contrast: 'contrast(1.35)', blur: 'blur(4px)', saturate: 'saturate(1.6)' };
export function isDarkBg(th) { const m = /#([0-9a-f]{6})/i.exec(th.bg); if (!m) return false; const v = parseInt(m[1], 16); const r = v >> 16, g = (v >> 8) & 255, b = v & 255; return r * 0.299 + g * 0.587 + b * 0.114 < 128; }
export function parseGrid(text) { return String(text || '').trim().split('\n').map((l) => l.split(/[;\t]/).map((v) => { const t = v.trim(); return t !== '' && isFinite(+t.replace(',', '.')) ? +t.replace(',', '.') : t; })); }
