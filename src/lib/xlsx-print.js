// Page setup between Excel and Kertas: paper, orientation, margins, scaling, print area,
// rows repeated on every page, manual page breaks, header and footer. Read and written with
// ExcelJS's page setup model. Kertas's form is what sheet-pdf.js prints with.
import { MARGINS } from './page-margins.js';

const PAPER_CODES = { letter: 1, legal: 5, a3: 8, a4: 9, a5: 11 };
const PAPER_BY_CODE = Object.fromEntries(Object.entries(PAPER_CODES).map(([k, v]) => [v, k]));
const MARGIN_INCHES = Object.fromEntries(Object.entries(MARGINS).map(([k, m]) => [k, { left: m.l / 72, right: m.r / 72, top: m.t / 72, bottom: m.b / 72, header: m.h / 72, footer: m.f / 72 }]));
// Excel header/footer codes and Kertas's readable ones
const CODES = [['&P', '{page}'], ['&N', '{pages}'], ['&D', '{date}'], ['&T', '{time}'], ['&F', '{file}'], ['&A', '{sheet}']];

const colIndex = (l) => [...l.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const colLetters = (n) => { let s = ''; for (n += 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/** "A1:G40" (also with a sheet name and $) -> range */
export function parseArea(text) {
  const m = /^(?:.*!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(String(text || '').split(',')[0].trim());
  if (!m) return null;
  const a = { startColumn: colIndex(m[1]), startRow: +m[2] - 1, endColumn: colIndex(m[3] || m[1]), endRow: +(m[4] || m[2]) - 1 };
  return { startRow: Math.min(a.startRow, a.endRow), endRow: Math.max(a.startRow, a.endRow), startColumn: Math.min(a.startColumn, a.endColumn), endColumn: Math.max(a.startColumn, a.endColumn) };
}
export const areaText = (r) => (r ? `${colLetters(r.startColumn)}${r.startRow + 1}:${colLetters(r.endColumn)}${r.endRow + 1}` : '');
/** "1:2" or "$1:$2" -> [0, 1] */
export function parseTitleRows(text) {
  const m = /^\$?(\d+)(?::\$?(\d+))?$/.exec(String(text || '').replace(/^.*!/, '').trim());
  return m ? [+m[1] - 1, +(m[2] || m[1]) - 1].sort((a, b) => a - b) : null;
}

/** Excel's "&LLeft&CCentre&RRight" -> [left, centre, right] in Kertas's form */
export function parseHeaderFooter(text) {
  if (!text) return null;
  const out = ['', '', ''];
  let slot = 1;
  // &L &C &R pick the slot; other &codes (fonts, sizes) are dropped
  for (const part of String(text).split(/(&[LCR])/)) {
    if (part === '&L') slot = 0; else if (part === '&C') slot = 1; else if (part === '&R') slot = 2;
    else {
      let t = part;
      for (const [x, k] of CODES) t = t.split(x).join(k);
      t = t.replace(/&"[^"]*"/g, '').replace(/&\d+/g, '').replace(/&[BIUSEXYKOH]/g, '').replace(/&&/g, '&');
      out[slot] += t;
    }
  }
  return out.some(Boolean) ? out : null;
}
export function headerFooterText(cells) {
  if (!cells?.some(Boolean)) return undefined;
  const conv = (t) => { let s = String(t || '').replace(/&/g, '&&'); for (const [x, k] of CODES) s = s.split(k).join(x); return s; };
  return ['&L', '&C', '&R'].map((code, i) => (cells[i] ? code + conv(cells[i]) : '')).join('');
}

/** ExcelJS worksheet -> Kertas page setup, or null when the sheet has nothing special */
/** Manual page breaks from a sheet's XML (ExcelJS writes them but does not read them) */
export function rowBreaksFromXml(sheetXml) {
  const block = /<rowBreaks\b[\s\S]*?<\/rowBreaks>/.exec(sheetXml || '')?.[0] || '';
  return [...block.matchAll(/<brk\b([^>]*)\/?>/g)].map((m) => Number(/\bid="(\d+)"/.exec(m[1])?.[1])).filter((n) => n > 0);
}

export function readPageSetup(ws, sheetXml) {
  const ps = ws.pageSetup || {};
  const setup = {};
  if (PAPER_BY_CODE[ps.paperSize]) setup.paper = PAPER_BY_CODE[ps.paperSize];
  if (ps.orientation === 'landscape') setup.orientation = 'landscape';
  else if (ps.orientation === 'portrait' && (ps.printArea || ps.fitToPage)) setup.orientation = 'portrait';
  if (ps.fitToPage) setup.fit = (ps.fitToHeight || 0) >= 1 && (ps.fitToWidth || 0) >= 1 ? 'page' : 'width';
  else if (ps.scale && ps.scale !== 100) { setup.fit = 'custom'; setup.scale = ps.scale; }
  const m = ps.margins;
  if (m && typeof m.left === 'number') {
    const near = (a, b) => Math.abs(a - b) < 0.05;
    setup.margins = Object.entries(MARGIN_INCHES).find(([, v]) => near(v.left, m.left) && near(v.top, m.top))?.[0] || (m.left < 0.45 ? 'narrow' : m.left > 0.9 ? 'wide' : 'normal');
  }
  const area = parseArea(ps.printArea);
  if (area) setup.area = area;
  const titles = parseTitleRows(ps.printTitlesRow);
  if (titles) setup.titleRows = titles;
  if (ps.horizontalCentered) setup.center = true;
  if (ps.showGridLines) setup.gridlines = true;
  const fromModel = (ws.rowBreaks || []).map((b) => Number(b.id)).filter((n) => n > 0);
  const breaks = fromModel.length ? fromModel : rowBreaksFromXml(sheetXml);
  if (breaks.length) setup.breaks = breaks; // Excel's id is the row after which the page ends
  const hf = ws.headerFooter || {};
  const header = parseHeaderFooter(hf.oddHeader); const footer = parseHeaderFooter(hf.oddFooter);
  if (header) setup.header = header;
  if (footer) setup.footer = footer;
  // Only worth keeping when something was actually set up for printing
  const meaningful = setup.area || setup.titleRows || setup.breaks || setup.header || setup.footer || setup.fit || setup.orientation || setup.center;
  return meaningful ? setup : null;
}

/** Kertas page setup -> ExcelJS worksheet settings */
export function writePageSetup(ws, setup) {
  if (!setup) return;
  const ps = {
    paperSize: PAPER_CODES[setup.paper] || PAPER_CODES.a4,
    orientation: setup.orientation === 'landscape' ? 'landscape' : 'portrait',
    margins: MARGIN_INCHES[setup.margins] || MARGIN_INCHES.normal,
    horizontalCentered: !!setup.center,
    showGridLines: !!setup.gridlines,
  };
  if (setup.fit === 'width') Object.assign(ps, { fitToPage: true, fitToWidth: 1, fitToHeight: 0 });
  else if (setup.fit === 'page') Object.assign(ps, { fitToPage: true, fitToWidth: 1, fitToHeight: 1 });
  else Object.assign(ps, { fitToPage: false, scale: setup.fit === 'custom' ? Math.round(setup.scale || 100) : 100 });
  if (setup.area) ps.printArea = areaText(setup.area);
  if (setup.titleRows) ps.printTitlesRow = `${setup.titleRows[0] + 1}:${setup.titleRows[1] + 1}`;
  ws.pageSetup = { ...ws.pageSetup, ...ps };
  const oddHeader = headerFooterText(setup.header); const oddFooter = headerFooterText(setup.footer);
  if (oddHeader || oddFooter) ws.headerFooter = { ...(ws.headerFooter || {}), ...(oddHeader ? { oddHeader } : {}), ...(oddFooter ? { oddFooter } : {}) };
  for (const b of setup.breaks || []) if (b > 0) ws.getRow(b).addPageBreak(); // a page starts at row index b
}
