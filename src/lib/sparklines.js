// Sparklines: tiny line, column and win/loss charts that live in a cell.
// Stored per sheet as [{ r, c, ref, type, color, negColor, markers }] where ref
// is the data range ("Sheet1!B2:G2"). Saved in Excel's x14 sparkline groups,
// so they show in Excel and WPS. Pure module (drawing takes a context or page).
import { parseRef, refValues, quoteSheet } from './xlsx-charts.js';
import { parseA1, rangeToA1, colName } from './xlsx-features.js';

export const SPARK_TYPES = [
  { id: 'line', label: 'Line' },
  { id: 'column', label: 'Column' },
  { id: 'stacked', label: 'Win/Loss' },
];
export const DEFAULT_SPARK = { type: 'line', color: '#4F46E5', negColor: '#DC2626', markers: false };

const toNum = (v) => (v === '' || v === null || v === undefined || typeof v === 'boolean' || !Number.isFinite(Number(v)) ? null : Number(v));

/** The numbers a sparkline plots (blanks become gaps: null) */
export function sparkValues(snapshot, sp, sheetName) {
  return refValues(snapshot, sp.ref, sheetName).map(toNum);
}

/**
 * Geometry in a w x h box: { kind, points: [[x, y] | null], bars: [{ x, y, w, h, neg }], axis }
 * y grows downwards (screen coordinates).
 */
export function sparkGeometry(values, type, w, h, pad = 2) {
  const nums = values.filter((v) => v !== null);
  const n = values.length;
  if (!n || !nums.length) return { kind: type, points: [], bars: [], axis: null };
  const iw = Math.max(1, w - pad * 2); const ih = Math.max(1, h - pad * 2);
  if (type === 'stacked') {
    const bw = iw / n;
    const mid = pad + ih / 2;
    const bars = values.map((v, i) => (v === null || v === 0 ? null : { x: pad + i * bw + bw * 0.12, y: v > 0 ? pad : mid, w: bw * 0.76, h: ih / 2, neg: v < 0 })).filter(Boolean);
    return { kind: type, points: [], bars, axis: mid };
  }
  let lo = Math.min(...nums); let hi = Math.max(...nums);
  if (type === 'column') { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (hi === lo) { hi += 1; lo -= type === 'column' ? 0 : 1; }
  const yOf = (v) => pad + ih - ((v - lo) / (hi - lo)) * ih;
  if (type === 'column') {
    const bw = iw / n;
    const zero = yOf(0);
    const bars = values.map((v, i) => (v === null ? null : { x: pad + i * bw + bw * 0.12, y: Math.min(zero, yOf(v)), w: bw * 0.76, h: Math.max(0.6, Math.abs(yOf(v) - zero)), neg: v < 0 })).filter(Boolean);
    return { kind: type, points: [], bars, axis: lo < 0 ? zero : null };
  }
  const step = n > 1 ? iw / (n - 1) : 0;
  const points = values.map((v, i) => (v === null ? null : [pad + (n > 1 ? i * step : iw / 2), yOf(v)]));
  return { kind: 'line', points, bars: [], axis: lo < 0 && hi > 0 ? yOf(0) : null };
}

/** Draws on a 2D canvas context, in the box (x, y, w, h) */
export function drawSparkCanvas(ctx, sp, values, x, y, w, h, { dpr = 1 } = {}) {
  const g = sparkGeometry(values, sp.type, w, h, Math.max(2, h * 0.12));
  ctx.save();
  ctx.translate(x, y);
  if (g.axis !== null && sp.type !== 'stacked') { ctx.strokeStyle = 'rgba(100,116,139,.45)'; ctx.lineWidth = 1 / dpr; ctx.beginPath(); ctx.moveTo(0, g.axis); ctx.lineTo(w, g.axis); ctx.stroke(); }
  if (g.kind === 'line') {
    ctx.strokeStyle = sp.color || DEFAULT_SPARK.color;
    ctx.lineWidth = Math.max(1.25, h / 18);
    ctx.lineJoin = 'round'; ctx.lineCap = 'round';
    ctx.beginPath();
    let pen = false;
    for (const p of g.points) { if (!p) { pen = false; continue; } if (pen) ctx.lineTo(p[0], p[1]); else ctx.moveTo(p[0], p[1]); pen = true; }
    ctx.stroke();
    const real = g.points.filter(Boolean);
    const mark = (p, color) => { ctx.fillStyle = color; ctx.beginPath(); ctx.arc(p[0], p[1], Math.max(1.6, h / 14), 0, Math.PI * 2); ctx.fill(); };
    if (sp.markers) for (const p of real) mark(p, sp.color || DEFAULT_SPARK.color);
    else if (real.length) {
      // High and low points, like Excel's "High point / Low point" option
      const ys = real.map((p) => p[1]);
      mark(real[ys.indexOf(Math.min(...ys))], sp.color || DEFAULT_SPARK.color);
      mark(real[ys.indexOf(Math.max(...ys))], sp.negColor || DEFAULT_SPARK.negColor);
    }
  } else {
    for (const b of g.bars) { ctx.fillStyle = b.neg ? sp.negColor || DEFAULT_SPARK.negColor : sp.color || DEFAULT_SPARK.color; ctx.fillRect(b.x, b.y, b.w, b.h); }
  }
  ctx.restore();
}

/** Draws on a pdf-lib page in the cell rectangle (PDF coordinates: y up) */
export function drawSparkline(page, sp, rect, { rgb, hexToRgb }) {
  const values = sp.values || [];
  const { x, y, width: w, height: h } = rect;
  const g = sparkGeometry(values, sp.type, w, h, Math.max(1.5, h * 0.14));
  const Y = (yy) => y + h - yy;
  const color = hexToRgb(sp.color || DEFAULT_SPARK.color);
  const neg = hexToRgb(sp.negColor || DEFAULT_SPARK.negColor);
  if (g.axis !== null && sp.type !== 'stacked') page.drawLine({ start: { x, y: Y(g.axis) }, end: { x: x + w, y: Y(g.axis) }, thickness: 0.3, color: rgb(0.6, 0.64, 0.7) });
  if (g.kind === 'line') {
    let prev = null;
    for (const p of g.points) {
      if (!p) { prev = null; continue; }
      if (prev) page.drawLine({ start: { x: x + prev[0], y: Y(prev[1]) }, end: { x: x + p[0], y: Y(p[1]) }, thickness: Math.max(0.6, h / 22), color });
      prev = p;
    }
    const real = g.points.filter(Boolean);
    if (real.length) {
      const ys = real.map((p) => p[1]);
      const dot = (p, c) => page.drawCircle({ x: x + p[0], y: Y(p[1]), size: Math.max(0.8, h / 18), color: c });
      if (sp.markers) real.forEach((p) => dot(p, color));
      else { dot(real[ys.indexOf(Math.min(...ys))], color); dot(real[ys.indexOf(Math.max(...ys))], neg); }
    }
  } else {
    for (const b of g.bars) page.drawRectangle({ x: x + b.x, y: Y(b.y + b.h), width: b.w, height: b.h, color: b.neg ? neg : color });
  }
}

// ------------------------------------------------------------------ Excel x14 XML
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
const SPARK_EXT = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const argb = (hex) => `FF${String(hex || '#000000').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)}`;
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** The <ext> element holding a sheet's sparkline groups (one group per style) */
export function sparklinesExtXml(list, sheetName) {
  if (!list?.length) return '';
  const groups = new Map();
  for (const sp of list) {
    const key = `${sp.type}|${sp.color}|${sp.negColor}|${sp.markers ? 1 : 0}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(sp);
  }
  const body = [...groups.values()].map((g) => {
    const s = g[0];
    const type = s.type === 'line' ? '' : ` type="${s.type}"`;
    const markers = s.type === 'line' ? (s.markers ? ' markers="1"' : ' high="1" low="1"') : ' negative="1"';
    const items = g.map((sp) => {
      const p = parseRef(sp.ref);
      const f = p ? `${quoteSheet(p.sheet || sheetName)}!${rangeToA1(p.range)}` : sp.ref;
      return `<x14:sparkline><xm:f>${esc(f)}</xm:f><xm:sqref>${colName(sp.c)}${sp.r + 1}</xm:sqref></x14:sparkline>`;
    }).join('');
    return `<x14:sparklineGroup${type} displayEmptyCellsAs="gap"${markers}><x14:colorSeries rgb="${argb(s.color)}"/><x14:colorNegative rgb="${argb(s.negColor)}"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers rgb="${argb(s.color)}"/><x14:colorFirst rgb="${argb(s.color)}"/><x14:colorLast rgb="${argb(s.color)}"/><x14:colorHigh rgb="${argb(s.color)}"/><x14:colorLow rgb="${argb(s.negColor)}"/><x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
  }).join('');
  return `<ext uri="${SPARK_EXT}" xmlns:x14="${X14}"><x14:sparklineGroups xmlns:xm="${XM}">${body}</x14:sparklineGroups></ext>`;
}

/** Adds the sparkline <ext> to a worksheet's XML (merging into an existing extLst) */
export function addSparklinesToSheetXml(xml, extXml) {
  if (!extXml) return xml;
  if (/<extLst>/.test(xml)) return xml.replace('</extLst>', `${extXml}</extLst>`);
  return xml.replace('</worksheet>', `<extLst>${extXml}</extLst></worksheet>`);
}

const hexOf = (a) => (a && /^[0-9A-F]{8}$/i.test(a) ? `#${a.slice(2).toUpperCase()}` : null);
/** Reads sparkline groups from a worksheet's XML */
export function parseSparklines(xml, sheetName) {
  const out = [];
  for (const g of String(xml).matchAll(/<(?:\w+:)?sparklineGroup\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?sparklineGroup>/g)) {
    const attrs = g[1];
    const type = attrs.match(/\btype="(\w+)"/)?.[1] || 'line';
    const color = hexOf(g[2].match(/<(?:\w+:)?colorSeries\b[^>]*\brgb="(\w+)"/)?.[1]) || DEFAULT_SPARK.color;
    const negColor = hexOf(g[2].match(/<(?:\w+:)?colorNegative\b[^>]*\brgb="(\w+)"/)?.[1]) || DEFAULT_SPARK.negColor;
    const markers = /\bmarkers="(1|true)"/.test(attrs);
    for (const s of g[2].matchAll(/<(?:\w+:)?sparkline>\s*<(?:\w+:)?f>([^<]*)<\/(?:\w+:)?f>\s*<(?:\w+:)?sqref>([^<]*)<\/(?:\w+:)?sqref>/g)) {
      const at = parseA1(s[2].trim());
      if (!at) continue;
      const ref = s[1].replace(/&amp;/g, '&').replace(/&apos;/g, "'").replace(/&quot;/g, '"');
      out.push({ r: at.startRow, c: at.startColumn, ref: parseRef(ref)?.sheet ? ref : `${quoteSheet(sheetName)}!${ref}`, type: ['line', 'column', 'stacked'].includes(type) ? type : 'line', color, negColor, markers });
    }
  }
  return out;
}

/**
 * One sparkline per row (or column) of data, placed in the location range.
 * data: { sheet, range } ; location: range of cells (a single row or column)
 */
export function makeSparklines(data, location, opts = {}) {
  const d = data.range; const L = location;
  const locCells = [];
  for (let r = L.startRow; r <= L.endRow; r++) for (let c = L.startColumn; c <= L.endColumn; c++) locCells.push({ r, c });
  const rows = d.endRow - d.startRow + 1; const cols = d.endColumn - d.startColumn + 1;
  const perRow = locCells.length === rows || (locCells.length !== cols && rows >= cols);
  const out = [];
  locCells.forEach((cell, i) => {
    const sub = perRow
      ? { startRow: d.startRow + i, endRow: d.startRow + i, startColumn: d.startColumn, endColumn: d.endColumn }
      : { startRow: d.startRow, endRow: d.endRow, startColumn: d.startColumn + i, endColumn: d.startColumn + i };
    if (sub.startRow > d.endRow || sub.startColumn > d.endColumn) return;
    out.push({ ...DEFAULT_SPARK, ...opts, ...cell, ref: `${quoteSheet(data.sheet)}!${rangeToA1(sub)}` });
  });
  return out;
}
