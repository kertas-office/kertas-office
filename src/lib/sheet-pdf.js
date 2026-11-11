// Renders a Univer workbook snapshot to PDF with pdf-lib.
// Handles values, number formats, fonts, fills, borders, alignment, wrap and merges,
// picks portrait or landscape, scales wide sheets and splits very wide ones.
// Pure module: no DOM, so it runs in Node tests too.
import { PDFDocument, StandardFonts, rgb } from '@cantoo/pdf-lib';
import { hexToRgb } from './pdf-ops.js';
import { createTextKit, loadFontBytes } from './pdf-text.js';
import { findFont, facePath } from './font-catalog.js';
import { readResources } from './xlsx-features.js';
import { cfEvaluator, iconLook } from './sheet-cf.js';
import { drawSparkline } from './sparklines.js';
import { MARGINS } from './page-margins.js';

export { MARGINS };

const A4 = [595.28, 841.89];
const MARGIN = 36;
const FOOTER = 18;
const PX_TO_PT = 0.75;
const MIN_SCALE = 0.55;
const MIN_FIT = 0.2; // the smallest "fit" scale a page setup can ask for
// Page setup (Page Layout tab), per sheet; a sheet without one prints as before
const PAPER = { a4: A4, letter: [612, 792], legal: [612, 1008], a3: [841.89, 1190.55], a5: [419.53, 595.28] };
const DEFAULT_MARGINS = { t: 36, b: 54, l: 36, r: 36, h: 18, f: 18 };
const DEFAULT_SETUP = { paper: 'a4', orientation: 'auto', margins: 'normal', fit: 'width', scale: 100, area: null, titleRows: null, breaks: [], header: ['', '', ''], footer: ['', 'Page {page} of {pages}', ''], center: false };
const validRange = (r) => (r && r.endRow >= r.startRow && r.endColumn >= r.startColumn && r.startRow >= 0 && r.startColumn >= 0 ? r : null);
const CELL = { STRING: 1, NUMBER: 2, BOOLEAN: 3 };
const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

// ------------------------------------------------------------- number formats
function stripLiterals(pattern) {
  return pattern.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, '');
}

function isDatePattern(p) {
  const s = stripLiterals(p).toLowerCase();
  return /[yd]/.test(s) || (/m/.test(s) && /[hs:]/.test(s));
}

function formatDate(serial, pattern) {
  const d = new Date(EXCEL_EPOCH + Math.round(serial * 86400000));
  const pad = (n) => String(n).padStart(2, '0');
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const p = pattern.replace(/\[[^\]]*\]/g, '');
  const tokens = p.match(/"[^"]*"|\\.|yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|am\/pm|./gi) || [];
  const ampm = /am\/pm/i.test(p);
  let afterHour = false;
  return tokens.map((tk, i) => {
    const t = tk.toLowerCase();
    if (tk.startsWith('"')) return tk.slice(1, -1);
    if (tk.startsWith('\\')) return tk.slice(1);
    const nextIsSec = /^s/.test((tokens.slice(i + 1).find((x) => /[a-z]/i.test(x)) || '').toLowerCase());
    switch (t) {
      case 'yyyy': return d.getUTCFullYear();
      case 'yy': return String(d.getUTCFullYear()).slice(2);
      case 'mmmm': return months[d.getUTCMonth()];
      case 'mmm': return months[d.getUTCMonth()].slice(0, 3);
      case 'mm': case 'm': {
        const minutes = afterHour || nextIsSec;
        afterHour = false;
        const v = minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1;
        return t === 'mm' ? pad(v) : v;
      }
      case 'dd': return pad(d.getUTCDate());
      case 'd': return d.getUTCDate();
      case 'hh': case 'h': {
        afterHour = true;
        let h = d.getUTCHours();
        if (ampm) h = h % 12 || 12;
        return t === 'hh' ? pad(h) : h;
      }
      case 'ss': return pad(d.getUTCSeconds());
      case 's': return d.getUTCSeconds();
      case 'am/pm': return d.getUTCHours() < 12 ? 'AM' : 'PM';
      default: return tk;
    }
  }).join('');
}

export function formatValue(v, t, pattern) {
  if (v === null || v === undefined) return '';
  if (t === CELL.BOOLEAN) return v === true || v === 1 || v === 'TRUE' ? 'TRUE' : 'FALSE';
  const n = typeof v === 'number' ? v : (t === CELL.NUMBER ? Number(v) : NaN);
  if (!Number.isFinite(n)) return String(v);
  if (!pattern || pattern === 'General' || pattern === '@') {
    if (Number.isInteger(n)) return String(n);
    return String(+n.toPrecision(10));
  }
  const section = pattern.split(';')[n < 0 && pattern.includes(';') ? 1 : 0] || pattern;
  if (isDatePattern(section)) return formatDate(n, section);
  const core = stripLiterals(section);
  const decimals = (core.split('.')[1] || '').replace(/[^0#]/g, '').length;
  const percent = core.includes('%');
  const group = /#,##|0,0/.test(core);
  let num = Math.abs(percent ? n * 100 : n);
  let out = num.toFixed(decimals);
  if (group) {
    const [i, f] = out.split('.');
    out = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f !== undefined ? `.${f}` : '');
  }
  // Literal prefix/suffix such as "RM" or $ around the number
  const firstNum = section.search(/[0#?]/);
  const lastNum = Math.max(section.lastIndexOf('0'), section.lastIndexOf('#'), section.lastIndexOf('?'));
  const clean = (s) => s.replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/[_*]./g, '').replace(/[%,]/g, '');
  const prefix = firstNum > 0 ? clean(section.slice(0, firstNum)) : '';
  const suffix = lastNum >= 0 ? clean(section.slice(lastNum + 1)) : '';
  const sign = n < 0 && !pattern.includes(';') ? '-' : '';
  const wrapNeg = n < 0 && pattern.includes(';') && /\(/.test(pattern.split(';')[1]);
  const body = `${sign}${prefix}${out}${percent ? '%' : ''}${suffix}`;
  return wrapNeg ? body.replace(/^\(|\)$/g, '') : body;
}

// ------------------------------------------------------------- helpers
const styleOf = (snap, s) => (s ? (typeof s === 'string' ? snap.styles?.[s] : s) || {} : {});

function cellText(c, style) {
  if (!c) return '';
  if ((c.v === undefined || c.v === null) && c.p?.body?.dataStream) return c.p.body.dataStream.replace(/[\r\n]+$/, '').replace(/\r/g, '\n');
  return formatValue(c.v, c.t ?? (typeof c.v === 'number' ? CELL.NUMBER : CELL.STRING), style.n?.pattern);
}

/** Pictures from the snapshot's drawing data: { [sheetId]: [{ col, colOff, row, rowOff, width, height, png|jpg }] } */
function sheetPictures(snapshot) {
  const out = {};
  let data;
  try { data = JSON.parse(snapshot.resources?.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')?.data || '{}'); } catch { return out; }
  for (const [sid, s] of Object.entries(data)) {
    for (const d of Object.values(s.data || {})) {
      const m = d.drawingType === 0 && typeof d.source === 'string' && d.source.match(/^data:image\/(png|jpe?g);base64,(.+)$/i);
      if (!m) continue;
      const bin = atob(m[2]);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const from = d.sheetTransform?.from || { column: d.column, columnOffset: d.columnOffset, row: d.row, rowOffset: d.rowOffset };
      (out[sid] ||= []).push({
        col: from.column || 0, colOff: from.columnOffset || 0, row: from.row || 0, rowOff: from.rowOffset || 0,
        width: d.transform?.width || d.width || 100, height: d.transform?.height || d.height || 100,
        ...(m[1].toLowerCase() === 'png' ? { png: bytes } : { jpg: bytes }), key: d.drawingId,
      });
    }
  }
  return out;
}

/** The last cell a floating object covers */
function cellAtPx(sh, im) {
  const w = (c) => sh.columnData?.[c]?.w ?? sh.defaultColumnWidth ?? 88;
  const h = (r) => sh.rowData?.[r]?.h ?? sh.defaultRowHeight ?? 24;
  let col = im.col; let x = (im.colOff || 0) + im.width;
  while (x > w(col) && col < 16000) { x -= w(col); col++; }
  let row = im.row; let y = (im.rowOff || 0) + im.height;
  while (y > h(row) && row < 1e6) { y -= h(row); row++; }
  return { col, row };
}

function usedRange(sh) {
  let maxR = -1; let maxC = -1;
  for (const [r, row] of Object.entries(sh.cellData || {})) {
    for (const [c, cd] of Object.entries(row || {})) {
      if (!cd) continue;
      const hasContent = (cd.v !== undefined && cd.v !== null && cd.v !== '') || cd.f || cd.p || cd.s;
      if (hasContent) { maxR = Math.max(maxR, +r); maxC = Math.max(maxC, +c); }
    }
  }
  for (const m of sh.mergeData || []) { maxR = Math.max(maxR, m.endRow); maxC = Math.max(maxC, m.endColumn); }
  return { maxR, maxC };
}

function wrapLines(text, measure, size, width) {
  const out = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      const test = line + word;
      if (measure(test, size) <= width || !line.trim()) line = test;
      else { out.push(line.trimEnd()); line = word.trimStart(); }
    }
    out.push(line);
  }
  return out;
}

function clip(text, measure, size, width) {
  if (measure(text, size) <= width) return text;
  const chars = [...text]; // never split a surrogate pair
  let lo = 0; let hi = chars.length;
  while (lo < hi) {
    const mid = Math.ceil((lo + hi) / 2);
    if (measure(chars.slice(0, mid).join(''), size) <= width) lo = mid; else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
}

const BORDER_WIDTH = { 1: 0.5, 2: 0.25, 3: 0.5, 4: 0.5, 5: 0.5, 6: 0.5, 7: 1.5, 8: 1, 9: 1, 10: 1, 11: 1, 12: 1, 13: 1.6 };
const BORDER_DASH = { 3: [1, 1.5], 4: [3, 2], 5: [3, 2, 1, 2], 6: [3, 2, 1, 2, 1, 2], 9: [4, 2], 10: [4, 2, 1, 2], 11: [4, 2, 1, 2, 1, 2], 12: [4, 2, 1, 2] };

// ------------------------------------------------------------- main
/**
 * @param snapshot IWorkbookData
 * @param opts { sheetIds?: string[], gridlines?: boolean, title?: string,
 *   images?: { [sheetId]: [{ col, colOff, row, rowOff, width, height, png: Uint8Array }] } }  (charts, drawn as pictures)
 *   print?: { [sheetId]: page setup } paper, orientation, margins, fit ('width'|'page'|'none'|'custom'), scale,
 *     area (range), titleRows [first, last], breaks [row before which a page starts], header/footer
 *     [left, centre, right] with {page} {pages} {date} {time} {file} {sheet}, gridlines, center
 * Pictures placed on the sheet (Univer drawings) are printed too.
 */
export async function snapshotToPdf(snapshot, { sheetIds, gridlines = true, title, images = {}, sparklines = {}, now, print = {} } = {}) {
  const pdf = await PDFDocument.create();
  pdf.setProducer('Kertas Office');
  pdf.setCreator('Kertas Office');
  if (title) pdf.setTitle(title);
  const fonts = {
    r: await pdf.embedFont(StandardFonts.Helvetica),
    b: await pdf.embedFont(StandardFonts.HelveticaBold),
    i: await pdf.embedFont(StandardFonts.HelveticaOblique),
    bi: await pdf.embedFont(StandardFonts.HelveticaBoldOblique),
  };
  const grid = rgb(0.84, 0.86, 0.9);
  const muted = rgb(0.45, 0.5, 0.58);

  // Non-Latin text (Arabic, Malayalam, Chinese...) is drawn with embedded Noto fonts
  const kit = createTextKit(pdf, { regular: fonts.r, bold: fonts.b });
  const allText = [];
  for (const sid of sheetIds || snapshot.sheetOrder) {
    const sh = snapshot.sheets[sid];
    if (!sh || sh.hidden) continue;
    allText.push(sh.name);
    for (const row of Object.values(sh.cellData || {})) for (const cd of Object.values(row || {})) if (cd) allText.push(cellText(cd, styleOf(snapshot, cd.s)));
  }
  await kit.prepare(allText);

  // Fonts chosen in the sheet (Roboto, Calibri...) are embedded so the PDF matches the screen
  const faceFonts = new Map(); // 'Family|b|i' -> PDFFont | null
  const ORDER = { regular: ['regular'], italic: ['italic', 'regular'], bold: ['bold', 'regular'], boldItalic: ['boldItalic', 'bold', 'italic', 'regular'] };
  const faceKey = (st) => `${st.ff}|${st.bl ? 1 : 0}|${st.it ? 1 : 0}`;
  for (const st of Object.values(snapshot.styles || {})) {
    const f = st?.ff && findFont(st.ff);
    if (!f || faceFonts.has(faceKey(st))) continue;
    faceFonts.set(faceKey(st), null);
    const want = st.bl && st.it ? 'boldItalic' : st.bl ? 'bold' : st.it ? 'italic' : 'regular';
    for (const face of ORDER[want]) {
      const bytes = await loadFontBytes(`face:${facePath(f, face)}`).catch(() => null);
      if (!bytes) continue;
      try { faceFonts.set(faceKey(st), await pdf.embedFont(bytes, { subset: true })); break; } catch { /* try the next face */ }
    }
  }
  const fontOf = (st) => (st.ff && faceFonts.get(faceKey(st))) || (st.bl && st.it ? fonts.bi : st.bl ? fonts.b : st.it ? fonts.i : fonts.r);

  const ids = (sheetIds || snapshot.sheetOrder).filter((id) => snapshot.sheets[id] && !snapshot.sheets[id].hidden);
  let pageCount = 0;

  const pictures = sheetPictures(snapshot);
  const cfRules = readResources(snapshot).cf || {};
  const embedded = new Map();
  const embed = async (key, bytes, jpeg) => {
    if (!embedded.has(key)) embedded.set(key, await (jpeg ? pdf.embedJpg(bytes) : pdf.embedPng(bytes)).catch(() => null));
    return embedded.get(key);
  };

  for (const sid of ids) {
    const sh = snapshot.sheets[sid];
    const custom = print[sid];
    const setup = { ...DEFAULT_SETUP, ...(custom || {}) };
    const floating = [...(pictures[sid] || []), ...(images[sid] || [])];
    const cf = cfEvaluator(sh, cfRules[sid] || [], { now: now || new Date() });
    const sparks = new Map((sparklines[sid] || []).map((sp) => [`${sp.r}:${sp.c}`, sp]));
    let { maxR, maxC } = usedRange(sh);
    // Make room for pictures and charts
    for (const im of floating) {
      const end = cellAtPx(sh, im);
      maxR = Math.max(maxR, end.row);
      maxC = Math.max(maxC, end.col);
    }
    const area = validRange(setup.area);
    if (maxR < 0 && !area) continue;
    // What is printed: the print area, or everything used
    const R0 = area ? area.startRow : 0; const R1 = area ? area.endRow : maxR;
    const C0 = area ? area.startColumn : 0; const C1 = area ? area.endColumn : maxC;
    const titles = Array.isArray(setup.titleRows) && setup.titleRows[0] >= 0 && setup.titleRows[1] >= setup.titleRows[0] ? setup.titleRows : null;
    const lastRow = Math.max(R1, titles ? titles[1] : 0);
    const lastCol = C1;

    const defW = sh.defaultColumnWidth || 88;
    const defH = sh.defaultRowHeight || 24;
    const colW = []; const rowH = [];
    for (let c = 0; c <= lastCol; c++) colW.push(sh.columnData?.[c]?.hd ? 0 : (sh.columnData?.[c]?.w ?? defW) * PX_TO_PT);
    for (let r = 0; r <= lastRow; r++) rowH.push(sh.rowData?.[r]?.hd ? 0 : (sh.rowData?.[r]?.h ?? sh.rowData?.[r]?.ah ?? defH) * PX_TO_PT);
    const totalW = sumRange(colW, C0, C1);
    const totalH = sumRange(rowH, R0, R1);

    // Paper, orientation, margins and scale
    const paper = PAPER[setup.paper] || A4;
    const mg = custom ? (MARGINS[setup.margins] || MARGINS.normal) : DEFAULT_MARGINS;
    const landscape = setup.orientation === 'landscape' ? true : setup.orientation === 'portrait' ? false : totalW > paper[0] - mg.l - mg.r;
    const [pageW, pageH] = landscape ? [paper[1], paper[0]] : paper;
    const availW = pageW - mg.l - mg.r;
    const availH = pageH - mg.t - mg.b;
    let scale;
    if (!custom || setup.fit === 'width') scale = Math.max(custom ? MIN_FIT : MIN_SCALE, Math.min(1, availW / totalW));
    else if (setup.fit === 'page') scale = Math.max(MIN_FIT, Math.min(1, availW / totalW, availH / Math.max(1, totalH)));
    else if (setup.fit === 'custom') scale = Math.min(4, Math.max(0.1, (Number(setup.scale) || 100) / 100));
    else scale = 1;
    const sheetGrid = setup.gridlines ?? gridlines;
    const breaks = new Set((setup.breaks || []).map(Number));

    // Column bands (only needed when the scaled sheet is still too wide)
    const bands = [];
    let start = C0; let acc = 0;
    for (let c = C0; c <= C1; c++) {
      if (acc + colW[c] * scale > availW && c > start) { bands.push([start, c - 1]); start = c; acc = 0; }
      acc += colW[c] * scale;
    }
    bands.push([start, C1]);

    // Merge lookup
    const mergeAt = new Map(); const covered = new Set();
    for (const m of sh.mergeData || []) {
      mergeAt.set(`${m.startRow}:${m.startColumn}`, m);
      for (let r = m.startRow; r <= m.endRow; r++) for (let c = m.startColumn; c <= m.endColumn; c++) {
        if (r !== m.startRow || c !== m.startColumn) covered.add(`${r}:${c}`);
      }
    }

    const sheetPages = [];
    for (const [c0, c1] of bands) {
      // Row pages: manual breaks start a new page, and title rows repeat at the top
      const titleRows = titles ? Array.from({ length: titles[1] - titles[0] + 1 }, (_, k) => titles[0] + k) : [];
      const titleH = titleRows.reduce((h, r) => h + rowH[r] * scale, 0);
      let r = R0;
      while (r <= R1) {
        const withTitles = titles && r > titles[1];
        let h = withTitles ? titleH : 0;
        const body = [];
        while (r <= R1 && (!body.length || (!breaks.has(r) && h + rowH[r] * scale <= availH))) {
          body.push(r); h += rowH[r] * scale; r++;
        }
        const pageRows = withTitles ? [...titleRows, ...body] : body;
        const onPage = new Set(body);
        const lastOnPage = pageRows[pageRows.length - 1];
        const page = pdf.addPage([pageW, pageH]);
        pageCount++;
        sheetPages.push(page);
        const top = pageH - mg.t;
        const bandW = sumRange(colW, c0, c1) * scale;
        const left = mg.l + (setup.center ? Math.max(0, (availW - bandW) / 2) : 0);
        const xOf = (c) => { let x = left; for (let k = c0; k < c; k++) x += colW[k] * scale; return x; };
        const yAt = new Map();
        { let y = top; for (const rr of pageRows) { yAt.set(rr, y); y -= rowH[rr] * scale; } }
        const yOf = (rr) => yAt.get(rr) ?? top;
        const span = (m, rr) => { let s = 0; for (let k = rr; k <= Math.min(m.endRow, lastOnPage); k++) s += rowH[k] || 0; return s; };

        // Pass 1: fills and gridlines
        for (const rr of pageRows) {
          for (let c = c0; c <= c1; c++) {
            if (covered.has(`${rr}:${c}`)) continue;
            const m = mergeAt.get(`${rr}:${c}`);
            const w = m ? sumRange(colW, c, Math.min(m.endColumn, c1)) * scale : colW[c] * scale;
            const ch = m ? span(m, rr) * scale : rowH[rr] * scale;
            if (!w || !ch) continue;
            const x = xOf(c); const y = yOf(rr) - ch;
            const st = styleOf(snapshot, sh.cellData?.[rr]?.[c]?.s);
            const fx = cf(rr, c);
            const bg = fx?.style?.bg?.rgb || st.bg?.rgb;
            if (bg) page.drawRectangle({ x, y, width: w, height: ch, color: hexToRgb(bg) });
            if (fx?.bar) drawDataBar(page, fx.bar, x, y, w, ch, scale);
            if (sheetGrid) page.drawRectangle({ x, y, width: w, height: ch, borderColor: grid, borderWidth: 0.35 });
            const sp = sparks.get(`${rr}:${c}`);
            if (sp) drawSparkline(page, sp, { x, y, width: w, height: ch }, { rgb, hexToRgb });
          }
        }

        // Pass 2: borders and text
        for (const rr of pageRows) {
          for (let c = c0; c <= c1; c++) {
            if (covered.has(`${rr}:${c}`)) continue;
            const cd = sh.cellData?.[rr]?.[c];
            const fx = cf(rr, c);
            const st = fx?.style ? { ...styleOf(snapshot, cd?.s), ...fx.style } : styleOf(snapshot, cd?.s);
            const m = mergeAt.get(`${rr}:${c}`);
            const w = m ? sumRange(colW, c, Math.min(m.endColumn, c1)) * scale : colW[c] * scale;
            const ch = m ? span(m, rr) * scale : rowH[rr] * scale;
            if (!w || !ch) continue;
            const x = xOf(c); const yTop = yOf(rr); const y = yTop - ch;

            if (st.bd) {
              const side = (key, sx, sy, ex, ey) => {
                const b = st.bd[key];
                if (!b?.s) return;
                page.drawLine({
                  start: { x: sx, y: sy }, end: { x: ex, y: ey },
                  thickness: BORDER_WIDTH[b.s] || 0.5, color: hexToRgb(b.cl?.rgb || '#000000'),
                  ...(BORDER_DASH[b.s] ? { dashArray: BORDER_DASH[b.s] } : {}),
                });
              };
              side('t', x, yTop, x + w, yTop);
              side('b', x, y, x + w, y);
              side('l', x, y, x, yTop);
              side('r', x + w, y, x + w, yTop);
            }

            let iconW = 0;
            if (fx?.icon) {
              const s2 = Math.min(ch * 0.62, 9 * scale);
              drawIcon(page, iconLook(fx.icon), x + 2.5 * scale + s2 / 2, y + ch / 2, s2);
              iconW = s2 + 2 * scale;
            }
            const raw = fx?.hideValue ? '' : cellText(cd, st);
            if (!raw) continue;
            const plain = kit.isPlain(raw);
            const text = raw;
            const font = fontOf(st);
            const measure = plain ? (t, sz) => font.widthOfTextAtSize(t, sz) : (t, sz) => kit.width(t, sz, !!st.bl);
            const size = (st.fs || 11) * scale;
            const color = st.cl?.rgb ? hexToRgb(st.cl.rgb) : rgb(0.06, 0.09, 0.16);
            const pad = 2.5 * scale;
            const isNum = typeof cd?.v === 'number' && cd?.t !== CELL.STRING;
            const align = st.ht === 2 ? 'center' : st.ht === 3 ? 'right' : st.ht === 1 ? 'left' : isNum ? 'right' : 'left';
            const lines = st.tb === 3 ? wrapLines(text, measure, size, w - pad * 2) : [text.split('\n')[0]];
            const lineH = size * 1.2;
            const blockH = lines.length * lineH;
            const vt = st.vt || 3;
            let baseline = vt === 1 ? yTop - pad - size * 0.85
              : vt === 2 ? y + ch / 2 + blockH / 2 - size * 0.85
                : y + pad + (lines.length - 1) * lineH + size * 0.2;

            // Text may spill into empty neighbours on the right, like Excel
            let maxW = w - pad * 2;
            if (st.tb !== 3 && align === 'left' && !isNum) {
              for (let k = c + 1; k <= c1 && !sh.cellData?.[rr]?.[k]?.v && !sh.cellData?.[rr]?.[k]?.f && !covered.has(`${rr}:${k}`); k++) maxW += colW[k] * scale;
            }
            for (const ln of lines) {
              const s = clip(ln, measure, size, maxW);
              const tw = measure(s, size);
              const tx = align === 'right' ? x + w - pad - tw : align === 'center' ? x + (w - tw) / 2 : x + pad + iconW;
              if (plain) page.drawText(s, { x: tx, y: baseline, size, font, color });
              else kit.draw(page, s, { x: tx, y: baseline, size, color, bold: !!st.bl });
              if (st.ul?.s) page.drawLine({ start: { x: tx, y: baseline - size * 0.12 }, end: { x: tx + tw, y: baseline - size * 0.12 }, thickness: size / 16, color });
              if (st.st?.s) page.drawLine({ start: { x: tx, y: baseline + size * 0.3 }, end: { x: tx + tw, y: baseline + size * 0.3 }, thickness: size / 16, color });
              baseline -= lineH;
            }
          }
        }

        // Pass 3: pictures and charts that start on this page
        for (const [k, im] of floating.entries()) {
          if (!onPage.has(im.row) || im.col < c0 || im.col > c1) continue;
          const img = await embed(`${sid}:${k}:${im.key || ''}`, im.png || im.jpg, !!im.jpg);
          if (!img) continue;
          const w = im.width * PX_TO_PT * scale;
          const h2 = im.height * PX_TO_PT * scale;
          const x = xOf(im.col) + (im.colOff || 0) * PX_TO_PT * scale;
          const yTop = yOf(im.row) - (im.rowOff || 0) * PX_TO_PT * scale;
          page.drawImage(img, { x, y: yTop - h2, width: w, height: h2 });
        }

        if (!custom) kit.draw(page, `${sh.name}  -  Page ${pageCount}`, { x: MARGIN, y: MARGIN / 2, size: 8, color: muted });
      }
    }

    // Headers and footers, once the sheet's page count is known
    if (custom) {
      const fill = (tpl, n) => String(tpl || '')
        .replace(/\{page\}/gi, String(n)).replace(/\{pages\}/gi, String(sheetPages.length))
        .replace(/\{date\}/gi, (now || new Date()).toLocaleDateString('en-GB')).replace(/\{time\}/gi, (now || new Date()).toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }))
        .replace(/\{file\}/gi, title || '').replace(/\{sheet\}/gi, sh.name);
      sheetPages.forEach((page, k) => {
        const w = page.getWidth(); const h = page.getHeight();
        for (const [cells, y] of [[setup.header, h - mg.h - 8], [setup.footer, mg.f]]) {
          (cells || []).slice(0, 3).forEach((tpl, i) => {
            const text = fill(tpl, k + 1);
            if (!text) return;
            const tw = kit.width(text, 9);
            const x = i === 0 ? mg.l : i === 1 ? (w - tw) / 2 : w - mg.r - tw;
            kit.draw(page, text, { x, y, size: 9, color: rgb(0.2, 0.24, 0.3) });
          });
        }
      });
    }
  }
  if (!pageCount) {
    const page = pdf.addPage(A4);
    page.drawText('(empty workbook)', { x: MARGIN, y: A4[1] - MARGIN - 12, size: 11, font: fonts.i, color: muted });
  }
  return pdf.save();
}

// Data bars: a solid or fading bar from the axis, like Excel's
function drawDataBar(page, bar, x, y, w, h, scale) {
  const pad = 1.5 * scale;
  const inner = w - pad * 2;
  const bh = h - pad * 2;
  const axisX = x + pad + inner * bar.axis;
  const len = Math.abs(bar.ratio) * inner * (bar.axis ? (bar.negative ? bar.axis : 1 - bar.axis) : 1);
  const bx = bar.negative ? axisX - len : axisX;
  if (len <= 0.2) return;
  const color = hexToRgb(bar.color);
  if (bar.gradient) {
    const steps = 12;
    for (let i = 0; i < steps; i++) {
      const t = i / steps;
      page.drawRectangle({ x: bar.negative ? bx + len * (1 - (i + 1) / steps) : bx + len * t, y: y + pad, width: len / steps + 0.05, height: bh, color, opacity: 0.85 - t * 0.65 });
    }
    page.drawRectangle({ x: bx, y: y + pad, width: len, height: bh, borderColor: color, borderWidth: 0.5, opacity: 0 });
  } else {
    page.drawRectangle({ x: bx, y: y + pad, width: len, height: bh, color, opacity: 0.85 });
  }
  if (bar.axis) page.drawLine({ start: { x: axisX, y: y + pad }, end: { x: axisX, y: y + h - pad }, thickness: 0.5, color: rgb(0.2, 0.2, 0.2), dashArray: [1, 1] });
}

// Icon sets: arrows, circles, flags and stars drawn as vector shapes
function drawIcon(page, look, cx, cy, size) {
  const color = hexToRgb(look.color);
  const r = size / 2;
  const path = (d) => page.drawSvgPath(d, { x: cx, y: cy, color, scale: size / 20 });
  switch (look.shape) {
    case 'up': path('M0 -9 L8 1 L3 1 L3 9 L-3 9 L-3 1 L-8 1 Z'); break;
    case 'down': path('M0 9 L8 -1 L3 -1 L3 -9 L-3 -9 L-3 -1 L-8 -1 Z'); break;
    case 'side': path('M9 0 L-1 8 L-1 3 L-9 3 L-9 -3 L-1 -3 L-1 -8 Z'); break;
    case 'flag': path('M-6 -9 L-4 -9 L-4 9 L-6 9 Z M-4 -9 L8 -5 L-4 -1 Z'); break;
    case 'star': path('M0 -9 L2.6 -3 L9 -2.8 L4 1.4 L5.6 8 L0 4.4 L-5.6 8 L-4 1.4 L-9 -2.8 L-2.6 -3 Z'); break;
    default: page.drawCircle({ x: cx, y: cy, size: r, color });
  }
}

function sumRange(arr, a, b) { let s = 0; for (let i = a; i <= b; i++) s += arr[i] || 0; return s; }
