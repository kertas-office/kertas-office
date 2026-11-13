// Converts between .xlsx (via ExcelJS) and Univer's IWorkbookData snapshot.
// Covers values, formulas, number formats, fonts, fills, alignment, wrap,
// borders, merges, column widths, row heights and frozen panes, plus (via
// xlsx-features.js) filters, conditional formatting, data validation,
// hyperlinks, notes and pictures.
// Pure module: no DOM, so it runs in Node tests too.
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  normaliseRels, stripDrawings, importFeatures, readResources, exportSheetFeatures, applyPatches,
  sheetParts, relsOf, xOf, yOf, cellAt, linkCell,
} from './xlsx-features.js';
import { chartXml, parseChartXml, injectCharts, refValues, chartId } from './xlsx-charts.js';
import { parseSparklines, sparklinesExtXml, addSparklinesToSheetXml } from './sparklines.js';
import { injectPivots, readPivots } from './pivot-xlsx.js';
import { readPageSetup, writePageSetup } from './xlsx-print.js';
import { writeThreadedComments } from './xlsx-comments.js';
import { NAMES_RESOURCE, readDefinedNames, writeDefinedNames, readSheetKeep, writeSheetKeep } from './xlsx-keep.js';
import { AGGREGATES, normSpec, aggLabel, fieldNames } from './pivot.js';

/** Values of a block on a sheet of the snapshot (2D, blanks as '') */
export function gridOf(snapshot, sheetName, range) {
  const sh = Object.values(snapshot.sheets || {}).find((s) => s.name === sheetName);
  if (!sh || !range) return [];
  const out = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    const row = [];
    for (let c = range.startColumn; c <= range.endColumn; c++) {
      const cd = sh.cellData?.[r]?.[c];
      const v = cd?.v;
      row.push(v === undefined || v === null ? '' : cd.t === 2 || (typeof v === 'number' && cd.t !== 1) ? Number(v) : v);
    }
    out.push(row);
  }
  return out;
}

const px = (emu) => Math.round((emu || 0) / 9525);

const CELL = { STRING: 1, NUMBER: 2, BOOLEAN: 3 };
const H_ALIGN = { left: 1, center: 2, right: 3, justify: 4, distributed: 4, fill: 1, centerContinuous: 2 };
const H_ALIGN_BACK = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify' };
const V_ALIGN = { top: 1, middle: 2, bottom: 3, justify: 2, distributed: 2 };
const V_ALIGN_BACK = { 1: 'top', 2: 'middle', 3: 'bottom' };
const WRAP = 3;
const BORDER = {
  thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8,
  mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13,
};
const BORDER_BACK = Object.fromEntries(Object.entries(BORDER).map(([k, v]) => [v, k]));
const BORDER_SIDES = { top: 't', bottom: 'b', left: 'l', right: 'r' };

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const DEFAULT_COL_PX = 88;
const DEFAULT_ROW_PX = 24;
const EXCEL_COL_PX = 64;
const EXCEL_ROW_PX = 20;
const MIN_ROWS = 1000;
const MIN_COLS = 26;

// Excel column width is in "characters"; ~7px per char + 5px padding at default font
const colWidthToPx = (w) => Math.round(w * 7 + 5);
const pxToColWidth = (px) => Math.max(0, +((px - 5) / 7).toFixed(2));
const ptToPx = (pt) => Math.round((pt * 4) / 3);
const pxToPt = (px) => +((px * 3) / 4).toFixed(2);

const dateToSerial = (d) => (d.getTime() - EXCEL_EPOCH) / 86400000;

function argbToRgb(color) {
  if (!color || !color.argb) return undefined;
  const hex = color.argb.length === 8 ? color.argb.slice(2) : color.argb;
  return `#${hex}`;
}
function rgbToArgb(rgb) {
  if (!rgb) return undefined;
  const hex = rgb.replace('#', '').toUpperCase();
  if (/^[0-9A-F]{6}$/.test(hex)) return { argb: `FF${hex}` };
  const m = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (m) return { argb: `FF${m.slice(1, 4).map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase()}` };
  return undefined;
}

// Functions newer than Excel 2007 are stored with a prefix in the file
// (_xlfn.MAXIFS); without it Excel and LibreOffice show #NAME?
const XLFN = ['AGGREGATE', 'BITAND', 'BITOR', 'BITXOR', 'CEILING.MATH', 'CONCAT', 'COVARIANCE.P', 'COVARIANCE.S', 'DAYS', 'FLOOR.MATH', 'FORECAST.LINEAR',
  'IFNA', 'IFS', 'ISOWEEKNUM', 'MAXIFS', 'MINIFS', 'MODE.SNGL', 'NORM.DIST', 'NORM.INV', 'NORM.S.DIST', 'PERCENTILE.EXC', 'PERCENTILE.INC',
  'QUARTILE.EXC', 'QUARTILE.INC', 'RANK.AVG', 'RANK.EQ', 'STDEV.P', 'STDEV.S', 'SWITCH', 'TEXTJOIN', 'VAR.P', 'VAR.S', 'XLOOKUP', 'XMATCH', 'LET',
  'LAMBDA', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'CHOOSECOLS', 'CHOOSEROWS', 'TAKE', 'DROP'];
const XLWS = ['FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY'];
const fnRe = (names) => new RegExp(`(^|[^A-Za-z0-9_.])(${names.map((n) => n.replace(/\./g, '\\.')).join('|')})\\(`, 'gi');
const XLFN_RE = fnRe(XLFN);
const XLWS_RE = fnRe(XLWS);
/** Adds the storage prefixes, leaving text in quotes alone */
export function prefixFormula(f) {
  return f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part
    .replace(XLFN_RE, (m, pre, name) => `${pre}_xlfn.${name.toUpperCase()}(`)
    .replace(XLWS_RE, (m, pre, name) => `${pre}_xlfn._xlws.${name.toUpperCase()}(`))).join('');
}
export const unprefixFormula = (f) => f.replace(/_xlfn\.(_xlws\.)?/gi, '');

const INDENT_PX = 9; // one Excel indent step at the default font
let uid = 0;
const newId = (p) => `${p}${Date.now().toString(36)}${(uid++).toString(36)}`;

// ---------------------------------------------------------------- import

function styleFromCell(cell) {
  const s = {};
  const f = cell.font;
  if (f) {
    if (f.name) s.ff = f.name;
    if (f.size) s.fs = f.size;
    if (f.bold) s.bl = 1;
    if (f.italic) s.it = 1;
    if (f.underline) s.ul = { s: 1 };
    if (f.strike) s.st = { s: 1 };
    if (f.vertAlign === 'superscript') s.va = 3;
    else if (f.vertAlign === 'subscript') s.va = 2;
    const c = argbToRgb(f.color);
    if (c) s.cl = { rgb: c };
  }
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const c = argbToRgb(fill.fgColor);
    if (c) s.bg = { rgb: c };
  }
  const a = cell.alignment;
  if (a) {
    if (a.horizontal && H_ALIGN[a.horizontal]) s.ht = H_ALIGN[a.horizontal];
    if (a.vertical && V_ALIGN[a.vertical]) s.vt = V_ALIGN[a.vertical];
    if (a.wrapText) s.tb = WRAP;
    // turned text: Excel's angle (or stacked letters); indent as a left padding of about 9px a step
    if (a.textRotation === 'vertical') s.tr = { a: 0, v: 1 };
    else if (Number(a.textRotation)) s.tr = { a: Number(a.textRotation), v: 0 };
    if (Number(a.indent) > 0) s.pd = { ...(s.pd || {}), l: Math.round(Number(a.indent) * INDENT_PX) };
  }
  const b = cell.border;
  if (b) {
    const bd = {};
    for (const [side, key] of Object.entries(BORDER_SIDES)) {
      const edge = b[side];
      if (edge && edge.style && BORDER[edge.style]) {
        bd[key] = { s: BORDER[edge.style], cl: { rgb: argbToRgb(edge.color) || '#000000' } };
      }
    }
    if (Object.keys(bd).length) s.bd = bd;
  }
  if (cell.numFmt && cell.numFmt !== 'General') s.n = { pattern: cell.numFmt };
  return Object.keys(s).length ? s : null;
}

function valueFromCell(cell) {
  const out = {};
  const v = cell.value;
  if (v === null || v === undefined) return out;

  let formula = null;
  try { formula = cell.formula || null; } catch { formula = null; }
  if (formula) {
    out.f = `=${unprefixFormula(formula)}`;
    const r = cell.result;
    if (r !== undefined && r !== null && typeof r !== 'object') {
      out.v = r;
      out.t = typeof r === 'number' ? CELL.NUMBER : typeof r === 'boolean' ? CELL.BOOLEAN : CELL.STRING;
    } else if (r instanceof Date) {
      out.v = dateToSerial(r); out.t = CELL.NUMBER;
    }
    return out;
  }
  if (typeof v === 'number') return { v, t: CELL.NUMBER };
  if (typeof v === 'boolean') return { v: v ? 1 : 0, t: CELL.BOOLEAN };
  if (typeof v === 'string') return { v, t: CELL.STRING };
  if (v instanceof Date) return { v: dateToSerial(v), t: CELL.NUMBER, _date: true };
  if (v.richText) return { v: v.richText.map((r) => r.text).join(''), t: CELL.STRING };
  if (v.text !== undefined && v.hyperlink) return { v: String(v.text), t: CELL.STRING };
  if (v.error) return { v: v.error, t: CELL.STRING };
  return { v: String(cell.text ?? ''), t: CELL.STRING };
}

/** xlsx bytes -> IWorkbookData */
export async function xlsxToSnapshot(bytes, name = 'Workbook') {
  let zip = null;
  try { zip = await JSZip.loadAsync(bytes); } catch { zip = null; }
  let forExcelJS = bytes;
  if (zip) {
    await normaliseRels(zip);
    // ExcelJS gets a copy without drawing parts; pictures and charts are read from the zip directly
    const copy = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
    await stripDrawings(copy);
    forExcelJS = await copy.generateAsync({ type: 'uint8array' });
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(forExcelJS);
  const unitId = newId('wb');
  const built = [];

  const styles = {};
  const styleKeyToId = new Map();
  const internStyle = (s) => {
    if (!s) return undefined;
    const key = JSON.stringify(s);
    if (!styleKeyToId.has(key)) {
      const id = `s${styleKeyToId.size}`;
      styleKeyToId.set(key, id);
      styles[id] = s;
    }
    return styleKeyToId.get(key);
  };

  const sheets = {};
  const sheetOrder = [];
  wb.eachSheet((ws) => {
    if (ws.state === 'veryHidden') return;
    const id = newId('sheet');
    const cellData = {};
    const arrays = []; // array formulas: the cells they spill into are filled by the formula
    ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        if (cell.isMerged && cell.master !== cell) return;
        const cv = cell.value;
        if (cv && typeof cv === 'object' && cv.shareType === 'array' && cv.ref) arrays.push({ ref: cv.ref, r: rowNum - 1, c: colNum - 1 });
        const val = valueFromCell(cell);
        let style = styleFromCell(cell);
        if (val._date) {
          style = { ...(style || {}), n: style?.n || { pattern: 'yyyy-mm-dd' } };
          delete val._date;
        }
        const sid = internStyle(style);
        if (val.v === undefined && !val.f && !sid) return;
        const r = rowNum - 1;
        const c = colNum - 1;
        (cellData[r] ||= {})[c] = { ...val, ...(sid ? { s: sid } : {}) };
      });
    });

    // The values Excel stored in a spill range are only a copy of the formula's answer; left in
    // place they would block the formula from spilling (#SPILL!), so only their formatting stays
    for (const a of arrays) {
      // without a stored answer, the formula is worked out (and spills) when the sheet opens
      const master = cellData[a.r]?.[a.c];
      if (master?.f) { delete master.v; delete master.t; }
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a.ref);
      if (!m) continue;
      const col = (l) => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
      for (let r = +m[2] - 1; r <= +m[4] - 1; r++) {
        for (let c = col(m[1]); c <= col(m[3]); c++) {
          if (r === a.r && c === a.c) continue;
          const cd = cellData[r]?.[c];
          if (!cd || cd.f) continue;
          if (cd.s) cellData[r][c] = { s: cd.s }; else delete cellData[r][c];
        }
      }
    }

    const mergeData = [];
    const merges = ws.model.merges || [];
    for (const ref of merges) {
      const [a, b] = ref.split(':');
      const p1 = ws.getCell(a);
      const p2 = ws.getCell(b);
      mergeData.push({
        startRow: +p1.row - 1, startColumn: +p1.col - 1,
        endRow: +p2.row - 1, endColumn: +p2.col - 1,
      });
    }

    const columnData = {};
    (ws.columns || []).forEach((col, i) => {
      if (col && col.width) columnData[i] = { w: colWidthToPx(col.width) };
      if (col && col.hidden) (columnData[i] ||= {}).hd = 1;
    });
    const rowData = {};
    ws.eachRow({ includeEmpty: true }, (row, rowNum) => {
      if (row.height) rowData[rowNum - 1] = { h: ptToPx(row.height) };
      if (row.hidden) (rowData[rowNum - 1] ||= {}).hd = 1;
    });

    const view = (ws.views || [])[0];
    let freeze;
    if (view && view.state === 'frozen' && (view.xSplit || view.ySplit)) {
      const xs = view.xSplit || 0;
      const ys = view.ySplit || 0;
      freeze = { xSplit: xs, ySplit: ys, startRow: ys, startColumn: xs };
    }

    sheets[id] = {
      id,
      name: ws.name,
      hidden: ws.state === 'hidden' ? 1 : 0,
      rowCount: Math.max(MIN_ROWS, ws.rowCount + 100),
      columnCount: Math.max(MIN_COLS, ws.columnCount + 10),
      // Without these, Excel draws its standard 8.43-character columns and 15 pt rows
      defaultColumnWidth: ws.properties?.defaultColWidth ? colWidthToPx(ws.properties.defaultColWidth) : EXCEL_COL_PX,
      defaultRowHeight: ws.properties?.defaultRowHeight ? ptToPx(ws.properties.defaultRowHeight) : EXCEL_ROW_PX,
      cellData,
      mergeData,
      columnData,
      rowData,
      ...(freeze ? { freeze } : {}),
      // a sheet shown without gridlines (common in forms and dashboards) stays that way
      ...(view && view.showGridLines === false ? { showGridlines: 0 } : {}),
    };
    sheetOrder.push(id);
    built.push({ ws, sh: sheets[id] });
  });

  if (!sheetOrder.length) return blankSnapshot(name);
  const { resources, charts: anchors } = await importFeatures({ wb, zip, sheets: built, unitId });

  // Grouped rows and columns, sparklines and PivotTables are kept beside the Univer data
  const extras = { pivots: {}, sparklines: {}, outline: {}, slicers: {}, keep: {}, print: {} };
  {
    const parts = zip ? await sheetParts(zip) : [];
    for (const { ws, sh } of built) {
      try {
        const part = parts.find((p) => p.name === ws.name);
        const xml = part ? await zip.file(part.path)?.async('string') : '';
        const setup = readPageSetup(ws, xml);
        if (setup) extras.print[sh.id] = setup;
        // links to a place in the workbook stored without a file target, which ExcelJS skips
        for (const m of (xml || '').matchAll(/<hyperlink\b([^>]*?)\/?>/g)) {
          const ref = /\bref="([A-Z]+)(\d+)/.exec(m[1]);
          const loc = /\blocation="([^"]*)"/.exec(m[1])?.[1];
          if (!ref || !loc || /\br:id=/.test(m[1])) continue;
          const at = /^#?(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)/i.exec(loc.replace(/&apos;/g, "'").replace(/&amp;/g, '&'));
          const target = at && built.find((b) => b.ws.name === (at[1] || at[2]).replace(/''/g, "'"));
          if (!target) continue;
          const r = +ref[2] - 1; const c = [...ref[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
          const existing = sh.cellData?.[r]?.[c] || {};
          if (existing.p) continue;
          const text = existing.v === undefined || existing.v === null ? loc : String(existing.v);
          (sh.cellData[r] ||= {})[c] = { ...existing, ...linkCell(text, `#gid=${target.sh.id}&range=${at[3].toUpperCase()}${at[4]}`) };
        }
      } catch (e) { console.warn('page setup', e); }
    }
  }
  // Named ranges go to Univer (formulas use them); protection and tables are carried to the save
  if (zip) {
    try {
      const names = await readDefinedNames(zip, built.map((b) => b.sh.id));
      if (Object.keys(names).length) resources.push({ name: NAMES_RESOURCE, data: JSON.stringify(names) });
    } catch (e) { console.warn('Could not read named ranges', e); }
    try {
      for (const [sheetName, keep] of Object.entries(await readSheetKeep(zip))) {
        const target = built.find((b) => b.ws.name === sheetName);
        if (target) extras.keep[target.sh.id] = keep;
      }
    } catch (e) { console.warn('Could not read protection and tables', e); }
  }
  for (const { ws, sh } of built) {
    const o = {};
    ws.eachRow({ includeEmpty: true }, (row, n) => { if (row.outlineLevel) (o.rows ||= {})[n - 1] = row.outlineLevel; });
    (ws.columns || []).forEach((col, i) => { if (col?.outlineLevel) (o.cols ||= {})[i] = col.outlineLevel; });
    if (o.rows || o.cols) extras.outline[sh.id] = o;
  }
  if (zip) {
    const parts = await sheetParts(zip);
    for (const { ws, sh } of built) {
      const part = parts.find((p) => p.name === ws.name);
      const xml = part && await zip.file(part.path)?.async('string');
      const list = xml && /sparklineGroup/.test(xml) ? parseSparklines(xml, ws.name) : [];
      if (list.length) extras.sparklines[sh.id] = list;
    }
    try {
      for (const [sheetName, list] of Object.entries(await readSlicers(zip))) {
        const target = built.find((b) => b.ws.name === sheetName);
        if (target && list.length) extras.slicers[target.sh.id] = list;
      }
    } catch (e) { console.warn('Could not read slicers', e); }
    try {
      for (const pv of await readPivots(zip)) {
        const part = parts.find((p) => p.path === pv.sheetPath);
        const target = part && built.find((b) => b.ws.name === part.name);
        if (!target) continue;
        const loc = pv.location;
        extras.pivots[target.sh.id] = { name: pv.name, source: pv.source, spec: pv.spec, fields: pv.fields, supported: pv.supported, size: loc ? { rows: loc.endRow + 1, cols: loc.endColumn + 1 } : null, imported: true };
      }
    } catch (e) { console.warn('Could not read PivotTables', e); }
  }

  // Charts: read each DrawingML chart into our model (kept beside the Univer data)
  const charts = {};
  for (const [sid, list] of Object.entries(anchors)) {
    const sh = sheets[sid];
    for (const a of list) {
      const xml = await zip.file(a.target)?.async('string');
      const model = xml && parseChartXml(xml);
      if (!model || !a.from) continue;
      const colOff = px(a.from.colOff);
      const rowOff = px(a.from.rowOff);
      let width = 480;
      let height = 288;
      if (a.to) {
        width = xOf(sh, a.to.col, px(a.to.colOff)) - xOf(sh, a.from.col, colOff);
        height = yOf(sh, a.to.row, px(a.to.rowOff)) - yOf(sh, a.from.row, rowOff);
      } else if (a.ext) { width = px(a.ext.cx); height = px(a.ext.cy); }
      (charts[sid] ||= []).push({ id: chartId(), ...model, anchor: { col: a.from.col, colOff, row: a.from.row, rowOff, width: Math.max(80, width), height: Math.max(60, height) } });
    }
  }
  // the workbook's own default font (LibreOffice and Google Sheets use Arial 10) and its author,
  // title and dates, so saving does not quietly change them
  const custom = {};
  try {
    const st = await zip?.file('xl/styles.xml')?.async('string');
    const first = st && /<fonts\b[^>]*>\s*<font>([\s\S]*?)<\/font>/.exec(st)?.[1];
    if (first) {
      const fname = /<name\s+val="([^"]+)"/.exec(first)?.[1];
      const fsize = Number(/<sz\s+val="([^"]+)"/.exec(first)?.[1]);
      if (fname && (fname !== 'Calibri' || (fsize && fsize !== 11))) custom.defaultFont = { name: fname, size: fsize || 11 };
    }
  } catch { /* keep the default */ }
  const props = {};
  for (const k of ['creator', 'lastModifiedBy', 'title', 'subject', 'keywords', 'description', 'category', 'company', 'manager']) if (wb[k]) props[k] = String(wb[k]);
  if (wb.created instanceof Date && !Number.isNaN(+wb.created)) props.created = wb.created.toISOString();
  if (Object.keys(props).length) custom.props = props;
  return { id: unitId, name, appVersion: '1.0.0', locale: 'enUS', styles, sheetOrder, sheets, resources, charts, extras, ...(Object.keys(custom).length ? { custom } : {}) };
}

export function blankSnapshot(name = 'Book1') {
  const id = newId('sheet');
  return {
    id: newId('wb'), name, appVersion: '1.0.0', locale: 'enUS', styles: {},
    sheetOrder: [id],
    sheets: {
      [id]: {
        id, name: 'Sheet1', rowCount: MIN_ROWS, columnCount: MIN_COLS,
        defaultColumnWidth: DEFAULT_COL_PX, defaultRowHeight: DEFAULT_ROW_PX,
        cellData: {}, mergeData: [], columnData: {}, rowData: {},
      },
    },
  };
}

// ---------------------------------------------------------------- export

function resolveStyle(snapshot, s) {
  if (!s) return null;
  return typeof s === 'string' ? snapshot.styles?.[s] || null : s;
}

function applyStyle(cell, s) {
  if (!s) return;
  const font = {};
  if (s.ff) font.name = s.ff;
  if (s.fs) font.size = s.fs;
  if (s.bl) font.bold = true;
  if (s.it) font.italic = true;
  if (s.ul && s.ul.s) font.underline = true;
  if (s.st && s.st.s) font.strike = true;
  if (s.va === 3) font.vertAlign = 'superscript';
  else if (s.va === 2) font.vertAlign = 'subscript';
  const fc = rgbToArgb(s.cl?.rgb);
  if (fc) font.color = fc;
  if (Object.keys(font).length) cell.font = font;

  const bg = rgbToArgb(s.bg?.rgb);
  if (bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: bg, bgColor: bg };

  const al = {};
  if (s.ht && H_ALIGN_BACK[s.ht]) al.horizontal = H_ALIGN_BACK[s.ht];
  if (s.vt && V_ALIGN_BACK[s.vt]) al.vertical = V_ALIGN_BACK[s.vt];
  if (s.tb === WRAP) al.wrapText = true;
  if (s.tr?.v === 1) al.textRotation = 'vertical';
  else if (Number(s.tr?.a)) al.textRotation = Math.max(-90, Math.min(90, Math.round(Number(s.tr.a))));
  if (Number(s.pd?.l) >= INDENT_PX / 2) al.indent = Math.min(15, Math.round(Number(s.pd.l) / INDENT_PX));
  if (Object.keys(al).length) cell.alignment = al;

  if (s.bd) {
    const border = {};
    for (const [side, key] of Object.entries(BORDER_SIDES)) {
      const e = s.bd[key];
      if (e && e.s && BORDER_BACK[e.s]) {
        border[side] = { style: BORDER_BACK[e.s], color: rgbToArgb(e.cl?.rgb) || { argb: 'FF000000' } };
      }
    }
    if (Object.keys(border).length) cell.border = border;
  }
  if (s.n?.pattern) cell.numFmt = s.n.pattern;
}

function cellValue(c) {
  if (c.f) {
    const result = c.v === undefined || c.v === null || c.v === '' ? undefined : c.v;
    return { formula: prefixFormula(c.f.replace(/^=/, '')), ...(result !== undefined ? { result } : {}) };
  }
  if (c.v === undefined || c.v === null) {
    // Rich text cells store their text in the paragraph model
    const body = c.p?.body?.dataStream;
    return body ? body.replace(/[\r\n]+$/, '').replace(/\r/g, '\n') : null;
  }
  if (c.t === CELL.BOOLEAN) return c.v === 1 || c.v === true || c.v === 'TRUE';
  if (c.t === CELL.NUMBER || (c.t === undefined && typeof c.v === 'number')) {
    const n = Number(c.v);
    return Number.isFinite(n) ? n : String(c.v);
  }
  return String(c.v);
}

/**
 * IWorkbookData -> xlsx bytes (Uint8Array).
 * charts: { [sheetId]: [chart model] } written as real Excel charts.
 */
/** True when the cell holds a note of its own (not a conversation) */
function hasOwnNote(res, sid, ref) {
  const m = String(ref || '').toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!m) return false;
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return !!res.note?.[sid]?.[+m[2] - 1]?.[col]?.note;
}

export async function snapshotToXlsx(snapshot, { charts = snapshot.charts || {}, extras = snapshot.extras || {} } = {}) {
  const wb = new ExcelJS.Workbook();
  const props = snapshot.custom?.props || {};
  for (const k of ['creator', 'lastModifiedBy', 'title', 'subject', 'keywords', 'description', 'category', 'company', 'manager']) if (props[k]) wb[k] = props[k];
  wb.created = props.created ? new Date(props.created) : new Date();
  wb.modified = new Date();
  // Excel and LibreOffice work the formulas out again when they open the file
  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };
  const res = readResources(snapshot);
  const sheetNameById = Object.fromEntries(Object.values(snapshot.sheets || {}).map((s) => [s.id, s.name]));
  const patches = [];
  const dynamicCells = {}; // sheet name -> cells holding dynamic array formulas

  for (const sid of snapshot.sheetOrder) {
    const sh = snapshot.sheets[sid];
    if (!sh) continue;
    const noGrid = sh.showGridlines === 0 || sh.showGridlines === false;
    const freeze = sh.freeze && (sh.freeze.xSplit || sh.freeze.ySplit)
      ? [{ state: 'frozen', xSplit: sh.freeze.xSplit || 0, ySplit: sh.freeze.ySplit || 0, ...(noGrid ? { showGridLines: false } : {}) }]
      : noGrid ? [{ showGridLines: false }] : undefined;
    const ws = wb.addWorksheet(sh.name || 'Sheet', {
      state: sh.hidden ? 'hidden' : 'visible',
      ...(freeze ? { views: freeze } : {}),
      // Keep the grid the same size in Excel as here, so pictures and charts line up
      properties: { defaultColWidth: pxToColWidth(sh.defaultColumnWidth || DEFAULT_COL_PX), defaultRowHeight: pxToPt(sh.defaultRowHeight || DEFAULT_ROW_PX) },
    });

    for (const [r, row] of Object.entries(sh.cellData || {})) {
      for (const [c, cd] of Object.entries(row || {})) {
        if (!cd) continue;
        const cell = ws.getCell(+r + 1, +c + 1);
        let v = cellValue(cd);
        // a formula that fills several cells is saved as Excel's dynamic array
        const spill = cd.f && extras.spills?.[sid]?.[`${r}:${c}`];
        if (spill && v && typeof v === 'object') {
          v = { ...v, shareType: 'array', ref: `${colLetters(spill.startColumn)}${spill.startRow + 1}:${colLetters(spill.endColumn)}${spill.endRow + 1}` };
          (dynamicCells[sh.name] ||= []).push(`${colLetters(+c)}${+r + 1}`);
        }
        if (v !== null && v !== undefined) cell.value = v;
        applyStyle(cell, resolveStyle(snapshot, cd.s));
      }
    }

    for (const m of sh.mergeData || []) {
      ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
    }

    const defW = sh.defaultColumnWidth || DEFAULT_COL_PX;
    for (const [c, cd] of Object.entries(sh.columnData || {})) {
      const col = ws.getColumn(+c + 1);
      if (cd.w && cd.w !== defW) col.width = pxToColWidth(cd.w);
      if (cd.hd) col.hidden = true;
    }
    for (const [r, rd] of Object.entries(sh.rowData || {})) {
      const row = ws.getRow(+r + 1);
      const h = rd.h ?? rd.ah;
      if (h && h !== (sh.defaultRowHeight || DEFAULT_ROW_PX)) row.height = pxToPt(h);
      if (rd.hd) row.hidden = true;
    }

    if (extras.print?.[sid]) writePageSetup(ws, extras.print[sid]);

    // Row and column groups (Excel's outline)
    const ol = extras.outline?.[sid];
    if (ol) {
      let maxR = 0; let maxC = 0;
      for (const [r, l] of Object.entries(ol.rows || {})) if (l > 0) { ws.getRow(+r + 1).outlineLevel = Math.min(7, l); maxR = Math.max(maxR, l); }
      for (const [c, l] of Object.entries(ol.cols || {})) if (l > 0) { ws.getColumn(+c + 1).outlineLevel = Math.min(7, l); maxC = Math.max(maxC, l); }
      ws.properties.outlineLevelRow = Math.min(7, maxR);
      ws.properties.outlineLevelCol = Math.min(7, maxC);
      ws.properties.outlineProperties = { summaryBelow: true, summaryRight: true };
    }

    const patch = exportSheetFeatures({ wb, ws, sh, res, sheetNameById });
    if (Object.keys(patch).length) patches.push({ name: ws.name, ...patch });
  }

  if (!wb.worksheets.length) wb.addWorksheet('Sheet1');
  const buf = await wb.xlsx.writeBuffer();
  const chartCount = Object.values(charts).reduce((n, l) => n + (l?.length || 0), 0);
  const sparkSheets = Object.entries(extras.sparklines || {}).filter(([sid, l]) => l?.length && snapshot.sheets[sid]);
  const pivots = Object.entries(extras.pivots || {}).filter(([sid, p]) => snapshot.sheets[sid] && p?.spec && p.supported !== false && p.source?.range);
  const slicerSheets = Object.entries(extras.slicers || {}).filter(([sid, l]) => l?.length && snapshot.sheets[sid]);
  let names = {};
  try { names = JSON.parse(snapshot.resources?.find((r) => r.name === NAMES_RESOURCE)?.data || '{}') || {}; } catch { names = {}; }
  const hasNames = Object.keys(names).length > 0;
  const keepSheets = Object.entries(extras.keep || {}).filter(([sid]) => snapshot.sheets[sid]);
  const hasDynamic = Object.keys(dynamicCells).length > 0;
  // conversations by sheet name; a cell with a note of its own keeps that note as it is
  const commentSheets = Object.entries(res.comment || {}).filter(([sid, l]) => l?.length && snapshot.sheets[sid])
    .map(([sid, l]) => [snapshot.sheets[sid].name, l.map((t) => ({ ...t, ownNote: hasOwnNote(res, sid, t.ref) }))]);
  const defaultFont = snapshot.custom?.defaultFont;
  const internalLinks = hasInternalLinks(snapshot);
  if (!hasDynamic && !patches.length && !chartCount && !sparkSheets.length && !pivots.length && !slicerSheets.length && !hasNames && !keepSheets.length && !commentSheets.length && !defaultFont && !internalLinks && props.creator && props.lastModifiedBy) return new Uint8Array(buf);
  const zip = await JSZip.loadAsync(buf);
  // ExcelJS writes "Unknown" as the author of a file that has none, and as the last editor
  {
    const core = await zip.file('docProps/core.xml')?.async('string');
    if (core) zip.file('docProps/core.xml', core.replace(props.lastModifiedBy ? /$^/ : /<cp:lastModifiedBy>Unknown<\/cp:lastModifiedBy>/, '<cp:lastModifiedBy></cp:lastModifiedBy>').replace(props.creator ? /$^/ : /<dc:creator>Unknown<\/dc:creator>/, '<dc:creator></dc:creator>'));
  }
  if (defaultFont) await setDefaultFont(zip, defaultFont);
  if (internalLinks) await fixInternalLinks(zip);
  await applyPatches(zip, patches);
  if (hasDynamic) await markDynamicArrays(zip, dynamicCells);
  if (hasNames) await writeDefinedNames(zip, names, snapshot.sheetOrder.filter((sid) => snapshot.sheets[sid]));
  if (keepSheets.length) {
    const byName = Object.fromEntries(keepSheets.map(([sid, k]) => [snapshot.sheets[sid].name, k]));
    const sheetByName = Object.fromEntries(Object.values(snapshot.sheets).map((x) => [x.name, x]));
    await writeSheetKeep(zip, byName, (sheetName, r, c) => {
      const cd = sheetByName[sheetName]?.cellData?.[r]?.[c];
      return cd?.v ?? '';
    });
  }
  if (commentSheets.length) {
    const pathByName = Object.fromEntries((await sheetParts(zip)).map((p) => [p.name, p.path]));
    await writeThreadedComments(zip, Object.fromEntries(commentSheets), pathByName);
  }
  if (slicerSheets.length) await writeSlicers(zip, Object.fromEntries(slicerSheets.map(([sid, l]) => [snapshot.sheets[sid].name, l])));
  if (sparkSheets.length || pivots.length) {
    const parts = await sheetParts(zip);
    for (const [sid, list] of sparkSheets) {
      const part = parts.find((p) => p.name === snapshot.sheets[sid].name);
      if (!part) continue;
      zip.file(part.path, addSparklinesToSheetXml(await zip.file(part.path).async('string'), sparklinesExtXml(list, snapshot.sheets[sid].name)));
    }
    const list = [];
    for (const [sid, p] of pivots) {
      const part = parts.find((x) => x.name === snapshot.sheets[sid].name);
      const grid = gridOf(snapshot, p.source.sheet, p.source.range);
      if (!part || grid.length < 2) continue;
      const fields = fieldNames(grid[0], p.source.range.startColumn);
      const nv = normSpec(p.spec).values;
      const label = nv.length === 1 ? `${aggLabel(nv[0].agg)} of ${fields[nv[0].fld]}` : 'Values';
      list.push({ sheetPath: part.path, name: p.name || `PivotTable${list.length + 1}`, source: p.source, grid, spec: p.spec, label });
    }
    await injectPivots(zip, list);
  }
  if (chartCount) {
    const parts = await sheetParts(zip);
    const list = [];
    for (const [sid, models] of Object.entries(charts)) {
      const sh = snapshot.sheets[sid];
      const part = sh && parts.find((p) => p.name === sh.name);
      if (!part) continue;
      for (const m of models || []) {
        const a = m.anchor || { col: 1, colOff: 0, row: 1, rowOff: 0, width: 480, height: 288 };
        const x1 = xOf(sh, a.col, a.colOff) + a.width;
        const y1 = yOf(sh, a.row, a.rowOff) + a.height;
        const ex = cellAt(sh, 'x', x1);
        const ey = cellAt(sh, 'y', y1);
        list.push({
          sheetPath: part.path,
          xml: chartXml(m, (ref) => refValues(snapshot, ref, sh.name)),
          anchorCells: { from: { col: a.col, colOff: a.colOff || 0, row: a.row, rowOff: a.rowOff || 0 }, to: { col: ex.index, colOff: ex.offset, row: ey.index, rowOff: ey.offset } },
        });
      }
    }
    await injectCharts(zip, list, relsOf);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------- dynamic arrays
// Excel tells a dynamic array formula (one that spills, like FILTER) from an old-style array
// formula by a cell metadata flag; without it Excel shows the formula as @FILTER(...) and only
// its first value.
const colLetters = (n) => { let s = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const DYNAMIC_METADATA = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';
/** Links to a place in the workbook (Sheet2!A1), which Excel stores without a file target */
function hasInternalLinks(snapshot) {
  for (const sh of Object.values(snapshot.sheets || {})) {
    for (const row of Object.values(sh.cellData || {})) {
      for (const cd of Object.values(row || {})) {
        if (cd?.p?.body?.customRanges?.some((r) => String(r.properties?.url || '').startsWith('#gid='))) return true;
      }
    }
  }
  return false;
}

/** ExcelJS also writes such a link as an outside file named "Sheet2!A1"; Excel would try to open that file */
async function fixInternalLinks(zip) {
  for (const part of await sheetParts(zip)) {
    let sx = await zip.file(part.path)?.async('string');
    if (!sx || !/<hyperlink\b[^>]*\blocation=/.test(sx)) continue;
    const relsPath = part.path.replace(/([^/]+)$/, '_rels/$1.rels');
    let rels = await zip.file(relsPath)?.async('string') || '';
    sx = sx.replace(/<hyperlink\b([^>]*?)\/>/g, (tag, attrs) => {
      const loc = /\blocation="([^"]*)"/.exec(attrs)?.[1];
      const id = /\br:id="([^"]*)"/.exec(attrs)?.[1];
      if (!loc || !id) return tag;
      const rel = new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*/>`).exec(rels)?.[0];
      const target = rel && /\bTarget="([^"]*)"/.exec(rel)?.[1];
      if (target !== loc && target !== `#${loc}`) return tag;
      rels = rels.replace(rel, '');
      return `<hyperlink${attrs.replace(/\s*r:id="[^"]*"/, '').replace(/location="#/, 'location="')}/>`;
    });
    zip.file(part.path, sx);
    if (rels) zip.file(relsPath, rels);
  }
}

/** The workbook's default font goes back to what it was when opened */
async function setDefaultFont(zip, { name, size }) {
  let st = await zip.file('xl/styles.xml')?.async('string');
  if (!st) return;
  const esc = (t) => String(t).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  st = st.replace(/(<fonts\b[^>]*>\s*<font>)([\s\S]*?)(<\/font>)/, (m, a, body, z) => {
    let b = body.replace(/<scheme\b[^>]*\/>/, ''); // a theme font would win over the name
    b = /<sz\b[^>]*\/>/.test(b) ? b.replace(/<sz\b[^>]*\/>/, `<sz val="${size}"/>`) : `<sz val="${size}"/>${b}`;
    b = /<name\b[^>]*\/>/.test(b) ? b.replace(/<name\b[^>]*\/>/, `<name val="${esc(name)}"/>`) : `${b}<name val="${esc(name)}"/>`;
    return a + b + z;
  });
  zip.file('xl/styles.xml', st);
}

async function markDynamicArrays(zip, bySheet) {
  const parts = await sheetParts(zip);
  for (const [name, cells] of Object.entries(bySheet)) {
    const part = parts.find((p) => p.name === name);
    if (!part) continue;
    let xml = await zip.file(part.path).async('string');
    for (const ref of cells) xml = xml.replace(new RegExp(`<c r="${ref}"(?![^>]*\\bcm=)`), `<c r="${ref}" cm="1"`);
    zip.file(part.path, xml);
  }
  if (zip.file('xl/metadata.xml')) return;
  zip.file('xl/metadata.xml', DYNAMIC_METADATA);
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>');
  zip.file('[Content_Types].xml', ct);
  let rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  let n = 1; while (rels.includes(`Id="rIdM${n}"`)) n++;
  rels = rels.replace('</Relationships>', `<Relationship Id="rIdM${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>`);
  zip.file('xl/_rels/workbook.xml.rels', rels);
}

// ---------------------------------------------------------------- slicers
// Kept in a custom XML part (by sheet name), which Excel carries along without showing.
// The filtering they do is saved as the sheet's ordinary filter, which Excel does show.
export const SLICER_NS = 'urn:kertas:slicers:1';
const xmlEsc = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const xmlUnesc = (t) => String(t).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

async function writeSlicers(zip, bySheetName) {
  const clean = Object.fromEntries(Object.entries(bySheetName).map(([n, list]) => [n, list.map((m) => ({
    title: m.title, range: m.range, col: m.col, selected: (m.selected || []).filter((k) => k !== '\u0000blank'), blank: (m.selected || []).includes('\u0000blank') || undefined, anchor: m.anchor,
  }))]));
  let n = 1;
  while (zip.file(`customXml/item${n}.xml`)) n++;
  const part = `customXml/item${n}.xml`;
  zip.file(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<kertasSlicers xmlns="${SLICER_NS}"><data>${xmlEsc(JSON.stringify(clean))}</data></kertasSlicers>`);
  const relsPath = 'xl/_rels/workbook.xml.rels';
  let rels = await zip.file(relsPath).async('string');
  let id = 1;
  while (rels.includes(`Id="rIdK${id}"`)) id++;
  rels = rels.replace('</Relationships>', `<Relationship Id="rIdK${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../${part}"/></Relationships>`);
  zip.file(relsPath, rels);
  let ct = await zip.file('[Content_Types].xml').async('string');
  if (!/Extension="xml"/i.test(ct)) {
    ct = ct.replace(/(<Types[^>]*>)/, '$1<Default Extension="xml" ContentType="application/xml"/>');
    zip.file('[Content_Types].xml', ct);
  }
}

async function readSlicers(zip) {
  for (const name of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(name)) continue;
    const text = await zip.file(name).async('string');
    if (!text.includes(SLICER_NS)) continue;
    const body = /<data>([\s\S]*?)<\/data>/.exec(text)?.[1];
    const data = body ? JSON.parse(xmlUnesc(body)) : {};
    for (const list of Object.values(data)) {
      for (const m of list) { m.selected = [...(m.selected || []), ...(m.blank ? ['\u0000blank'] : [])]; delete m.blank; }
    }
    return data;
  }
  return {};
}

// ---------------------------------------------------------------- CSV

export function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = '';
  let quoted = false;
  const delim = guessDelimiter(text);
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"' && text[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === delim) { row.push(field); field = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += ch;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

function guessDelimiter(text) {
  const head = text.slice(0, 2000).split(/\r?\n/)[0] || '';
  const counts = [',', ';', '\t'].map((d) => [d, head.split(d).length]);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 1 ? counts[0][0] : ',';
}

/**
 * A CSV value as a cell. Numbers become numbers so they can be added up, but a value is kept as
 * text when turning it into a number would change it: codes and phone numbers with leading zeros,
 * and numbers longer than a spreadsheet can hold exactly. A value starting with "=" stays text
 * (a CSV holds data, and running formulas from one is a known way to attack spreadsheet users).
 * The original spelling of a number ("1,5" style aside) is remembered, so saving without editing
 * gives back exactly what was opened.
 */
export function csvCell(raw) {
  const t = raw.trim();
  if (/^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(t)) {
    const digits = t.replace(/^[-+]/, '').replace(/[eE].*$/, '').replace('.', '').replace(/^0+/, '');
    const leadingZero = /^[-+]?0\d/.test(t);
    if (!leadingZero && digits.length <= 15) {
      const v = Number(t);
      return String(v) === raw ? { v, t: CELL.NUMBER } : { v, t: CELL.NUMBER, custom: { csv: raw } };
    }
  }
  return { v: raw, t: CELL.STRING };
}

export function csvToSnapshot(text, name = 'Book1') {
  const snap = blankSnapshot(name);
  const sh = snap.sheets[snap.sheetOrder[0]];
  sh.name = name.replace(/\.csv$/i, '').slice(0, 31) || 'Sheet1';
  const body = text.replace(/^\uFEFF/, '');
  const rows = parseCsv(body);
  snap.custom = { ...(snap.custom || {}), csvBom: text.startsWith('\uFEFF'), csvDelimiter: guessDelimiter(body), csvNewline: /\r\n/.test(body.slice(0, 20000)) || !/\n/.test(body) ? '\r\n' : '\n' };
  let maxCols = 0;
  rows.forEach((cols, r) => {
    maxCols = Math.max(maxCols, cols.length);
    cols.forEach((raw, c) => {
      if (raw === '') return;
      (sh.cellData[r] ||= {})[c] = csvCell(raw);
    });
  });
  sh.rowCount = Math.max(MIN_ROWS, rows.length + 100);
  sh.columnCount = Math.max(MIN_COLS, maxCols + 10);
  return snap;
}

/** Exports one sheet (the given id, or the first) to CSV text */
export function snapshotToCsv(snapshot, sheetId) {
  const sh = snapshot.sheets[sheetId || snapshot.sheetOrder[0]];
  const rows = Object.keys(sh.cellData || {}).map(Number);
  if (!rows.length) return '';
  const maxRow = Math.max(...rows);
  let maxCol = 0;
  for (const r of rows) for (const c of Object.keys(sh.cellData[r] || {})) maxCol = Math.max(maxCol, +c);
  // the file is written back with the separator it was opened with (, or ; or tab)
  const delim = [',', ';', '\t'].includes(snapshot.custom?.csvDelimiter) ? snapshot.custom.csvDelimiter : ',';
  const needsQuotes = new RegExp(`["\\n\\r${delim === '\t' ? '\\t' : delim}]`);
  const esc = (s) => (needsQuotes.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines = [];
  for (let r = 0; r <= maxRow; r++) {
    const line = [];
    for (let c = 0; c <= maxCol; c++) {
      const cd = sh.cellData[r]?.[c];
      const v = cd ? cellValue(cd) : null;
      let out = v === null || v === undefined ? '' : typeof v === 'object' ? String(v.result ?? '') : String(v);
      // a number not edited since the file was opened keeps its original spelling (1.50, 1E5, -0)
      const orig = cd?.custom?.csv;
      if (typeof orig === 'string' && !cd.f && typeof cd.v === 'number' && Number(orig.trim()) === cd.v) out = orig;
      line.push(esc(out));
    }
    lines.push(line.join(delim));
  }
  const nl = snapshot.custom?.csvNewline === '\n' ? '\n' : '\r\n';
  return `${lines.join(nl)}${nl}`;
}
