// Shared SVG chart renderer (Xel + Point)
export const CHART_COLORS = ['#3b6fe0', '#e8702a', '#2fa36b', '#d64545', '#8a5cd6', '#c49a1a', '#1aa3b8', '#d6589a', '#6b7280', '#5a8f29'];
export const CHART_TYPES = [
  { id: 'column', name: 'Гистограмма', icon: 'chart' },
  { id: 'bar', name: 'Линейчатая', icon: 'chart-bar' },
  { id: 'line', name: 'График', icon: 'chart-line' },
  { id: 'area', name: 'С областями', icon: 'chart-area' },
  { id: 'pie', name: 'Круговая', icon: 'chart-pie' },
  { id: 'doughnut', name: 'Кольцевая', icon: 'chart-pie' },
  { id: 'scatter', name: 'Точечная', icon: 'chart-scatter' },
];
const e = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function niceScale(min, max, ticks = 5) {
  if (min === max) { max = min + 1; min = Math.min(0, min); }
  const span = max - min; const step0 = span / ticks; const mag = Math.pow(10, Math.floor(Math.log10(step0)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= step0) || 10 * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}
const fmt = (v) => (Math.abs(v) >= 1e6 ? (v / 1e6).toFixed(1).replace(/\.0$/, '') + 'M' : Math.abs(v) >= 1e4 ? (v / 1e3).toFixed(0) + 'k' : +v.toFixed(2) + '');

export function renderChart(c, W = 480, H = 300, opt = {}) {
  const text = opt.text || '#444'; const grid = opt.grid || 'rgba(128,128,128,.25)'; const font = opt.font || 'Segoe UI, system-ui, sans-serif';
  const series = (c.series || []).map((s, i) => ({ ...s, color: s.color || CHART_COLORS[i % CHART_COLORS.length], values: s.values.map((v) => (typeof v === 'number' && isFinite(v) ? v : 0)) }));
  const cats = c.categories || series[0]?.values.map((_, i) => String(i + 1)) || [];
  let out = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="100%" height="100%" font-family="${e(font)}" font-size="11" preserveAspectRatio="none">`;
  if (opt.bg) out += `<rect width="${W}" height="${H}" fill="${opt.bg}"/>`;
  let top = 12;
  if (c.title) { out += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="14" font-weight="600" fill="${text}">${e(c.title)}</text>`; top = 34; }
  const showLegend = c.legend !== false && (series.length > 1 || c.type === 'pie' || c.type === 'doughnut');
  const legendH = showLegend ? 22 : 0;
  const plot = { x: 44, y: top, w: W - 58, h: H - top - 26 - legendH };
  if (!series.length || !cats.length) return out + `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" fill="${text}">Нет данных</text></svg>`;

  if (c.type === 'pie' || c.type === 'doughnut') {
    const vals = series[0].values.map((v) => Math.max(0, v)); const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2, cy = plot.y + (plot.h + 20) / 2, r = Math.min(plot.w, plot.h + 20) / 2 - 4, ir = c.type === 'doughnut' ? r * 0.55 : 0;
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2; const large = a1 - a0 > Math.PI ? 1 : 0; const col = CHART_COLORS[i % CHART_COLORS.length];
      const p = (a, rr) => `${cx + rr * Math.cos(a)},${cy + rr * Math.sin(a)}`;
      if (v / total >= 0.9999) out += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${col}"/>${ir ? `<circle cx="${cx}" cy="${cy}" r="${ir}" fill="${opt.bg || '#fff'}"/>` : ''}`;
      else out += `<path d="M${ir ? p(a0, ir) : `${cx},${cy}`} L${p(a0, r)} A${r},${r} 0 ${large} 1 ${p(a1, r)} ${ir ? `L${p(a1, ir)} A${ir},${ir} 0 ${large} 0 ${p(a0, ir)}` : ''}Z" fill="${col}" stroke="${opt.bg || '#fff'}" stroke-width="1.5"/>`;
      const am = (a0 + a1) / 2, lr = ir ? (r + ir) / 2 : r * 0.62;
      if (v / total > 0.05) out += `<text x="${cx + lr * Math.cos(am)}" y="${cy + lr * Math.sin(am) + 4}" text-anchor="middle" fill="#fff" font-weight="600">${Math.round((v / total) * 100)}%</text>`;
      a0 = a1;
    });
    if (showLegend) out += legend(cats.map((n, i) => ({ name: n, color: CHART_COLORS[i % CHART_COLORS.length] })), W, H, text);
    return out + '</svg>';
  }

  const all = series.flatMap((s) => s.values);
  const horiz = c.type === 'bar';
  if (horiz) { const lw = Math.min(110, 10 + Math.max(...cats.map((x) => Math.min(12, String(x).length))) * 6.2); plot.x = lw; plot.w = W - lw - 16; }
  if (c.type === 'scatter') {
    const xs = cats.map((x, i) => (isFinite(+x) ? +x : i + 1)); const sx = niceScale(Math.min(...xs), Math.max(...xs)); const sy = niceScale(Math.min(0, ...all), Math.max(...all));
    out += axes(plot, sy, grid, text, false);
    for (let v = sx.min; v <= sx.max + 1e-9; v += sx.step) { const x = plot.x + ((v - sx.min) / (sx.max - sx.min)) * plot.w; out += `<text x="${x}" y="${plot.y + plot.h + 14}" text-anchor="middle" fill="${text}">${fmt(v)}</text>`; }
    series.forEach((s) => s.values.forEach((v, i) => { const x = plot.x + ((xs[i] - sx.min) / (sx.max - sx.min)) * plot.w; const y = plot.y + plot.h - ((v - sy.min) / (sy.max - sy.min)) * plot.h; out += `<circle cx="${x}" cy="${y}" r="4" fill="${s.color}"/>`; }));
  } else {
    const sc = niceScale(Math.min(0, ...all), Math.max(0, ...all));
    out += axes(plot, sc, grid, text, horiz);
    const n = cats.length; const band = (horiz ? plot.h : plot.w) / n;
    const val2px = (v) => ((v - sc.min) / (sc.max - sc.min)) * (horiz ? plot.w : plot.h);
    const zero = val2px(0);
    cats.forEach((cat, i) => {
      const label = String(cat).length > 12 ? String(cat).slice(0, 11) + '…' : cat;
      if (horiz) out += `<text x="${plot.x - 4}" y="${plot.y + band * i + band / 2 + 4}" text-anchor="end" fill="${text}">${e(label)}</text>`;
      else if (n <= 24 || i % Math.ceil(n / 24) === 0) out += `<text x="${plot.x + band * i + band / 2}" y="${plot.y + plot.h + 14}" text-anchor="middle" fill="${text}">${e(label)}</text>`;
    });
    if (c.type === 'column' || c.type === 'bar') {
      const bw = (band * 0.72) / series.length;
      series.forEach((s, si) => s.values.forEach((v, i) => {
        const off = band * 0.14 + bw * si; const len = Math.abs(val2px(v) - zero); const start = Math.min(val2px(v), zero);
        if (horiz) out += `<rect x="${plot.x + start}" y="${plot.y + band * i + off}" width="${len}" height="${Math.max(1, bw - 1)}" fill="${s.color}" rx="1.5"/>`;
        else out += `<rect x="${plot.x + band * i + off}" y="${plot.y + plot.h - start - len}" width="${Math.max(1, bw - 1)}" height="${len}" fill="${s.color}" rx="1.5"/>`;
      }));
    } else {
      series.forEach((s) => {
        const pts = s.values.map((v, i) => [plot.x + band * i + band / 2, plot.y + plot.h - val2px(v)]);
        if (c.type === 'area') out += `<path d="M${pts[0][0]},${plot.y + plot.h - zero} ${pts.map((p) => `L${p[0]},${p[1]}`).join(' ')} L${pts[pts.length - 1][0]},${plot.y + plot.h - zero}Z" fill="${s.color}" fill-opacity=".28"/>`;
        out += `<polyline points="${pts.map((p) => p.join(',')).join(' ')}" fill="none" stroke="${s.color}" stroke-width="2.2" stroke-linejoin="round" stroke-linecap="round"/>`;
        if (pts.length <= 40) pts.forEach((p) => { out += `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${s.color}"/>`; });
      });
    }
  }
  if (showLegend) out += legend(series, W, H, text);
  return out + '</svg>';
}
function axes(plot, sc, grid, text, horiz) {
  let o = '';
  for (let v = sc.min; v <= sc.max + sc.step / 1e6; v += sc.step) {
    const t = (v - sc.min) / (sc.max - sc.min);
    if (horiz) { const x = plot.x + t * plot.w; o += `<line x1="${x}" x2="${x}" y1="${plot.y}" y2="${plot.y + plot.h}" stroke="${grid}"/><text x="${x}" y="${plot.y + plot.h + 14}" text-anchor="middle" fill="${text}">${fmt(v)}</text>`; }
    else { const y = plot.y + plot.h - t * plot.h; o += `<line x1="${plot.x}" x2="${plot.x + plot.w}" y1="${y}" y2="${y}" stroke="${grid}"/><text x="${plot.x - 6}" y="${y + 4}" text-anchor="end" fill="${text}">${fmt(v)}</text>`; }
  }
  return o;
}
function legend(items, W, H, text) {
  let x = 0; const parts = items.map((s) => { const w = 18 + Math.min(String(s.name).length, 18) * 6.2 + 12; const r = { s, x, w }; x += w; return r; });
  const start = Math.max(4, (W - x) / 2);
  return parts.map(({ s, x: px }) => `<rect x="${start + px}" y="${H - 15}" width="10" height="10" rx="2" fill="${s.color}"/><text x="${start + px + 14}" y="${H - 6}" fill="${text}">${e(String(s.name).slice(0, 18))}</text>`).join('');
}

// Build chart data from a 2D array of values (first row = headers?, first col = categories?)
export function chartDataFromGrid(grid) {
  if (!grid.length) return { categories: [], series: [] };
  const isNum = (v) => typeof v === 'number' || (v !== '' && v != null && isFinite(+v));
  const hasHeader = grid[0].some((v, i) => i > 0 && !isNum(v)) || (grid[0].length > 1 && !isNum(grid[0][1]) && grid[0][1] !== '');
  const hasCats = grid.slice(hasHeader ? 1 : 0).some((r) => !isNum(r[0])) && grid[0].length > 1;
  const rows = grid.slice(hasHeader ? 1 : 0);
  const c0 = hasCats ? 1 : 0;
  const series = [];
  for (let c = c0; c < grid[0].length; c++) series.push({ name: hasHeader ? String(grid[0][c] ?? `Ряд ${c}`) : `Ряд ${c - c0 + 1}`, values: rows.map((r) => +r[c] || 0) });
  return { categories: hasCats ? rows.map((r) => String(r[0] ?? '')) : rows.map((_, i) => String(i + 1)), series };
}
