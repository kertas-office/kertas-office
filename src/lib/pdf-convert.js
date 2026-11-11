// PDF to Word (.docx) and PDF to Excel (.xlsx).
// Text, its position and style come from pdf.js; pdf-layout.js rebuilds
// paragraphs, headings, lists and tables; the docx library writes Word files.
// Works in the browser and in Node (tests pass pdf.js's legacy build).
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType,
  BorderStyle, ShadingType, ImageRun, PageBreak, LevelFormat,
} from 'docx';
import { analyse, pageRows, parseNumber } from './pdf-layout.js';
import { blankSnapshot } from './xlsx-bridge.js';

// --------------------------------------------------------------- reading

const STD = [
  [/times/i, 'Times New Roman'], [/arial|helvetica|liberationsans|arimo/i, 'Arial'], [/courier|mono/i, 'Courier New'],
  [/calibri|carlito/i, 'Calibri'], [/cambria|caladea/i, 'Cambria'], [/georgia/i, 'Georgia'], [/verdana/i, 'Verdana'],
  [/garamond/i, 'Garamond'], [/tahoma/i, 'Tahoma'], [/segoe/i, 'Segoe UI'], [/roboto/i, 'Roboto'], [/inter\b/i, 'Inter'],
  [/opensans|open sans/i, 'Open Sans'], [/lato/i, 'Lato'], [/noto ?sans/i, 'Noto Sans'], [/dejavu ?sans/i, 'DejaVu Sans'],
];
/** "ABCDEF+TimesNewRomanPS-BoldItalicMT" -> { family, bold, italic } */
export function fontInfo(raw = '') {
  const name = String(raw).replace(/^[A-Z]{6}\+/, '');
  const bold = /bold|black|heavy|semibold|demi/i.test(name);
  const italic = /italic|oblique/i.test(name);
  const family = STD.find(([re]) => re.test(name))?.[1] || null;
  return { family, bold, italic };
}

/**
 * Reads positioned text from an open pdf.js document.
 * Returns [{ width, height, items }] in the format pdf-layout.js expects.
 */
export async function extractPages(pdf, { Util, onProgress, pages: only } = {}) {
  const out = [];
  const total = pdf.numPages;
  for (let n = 1; n <= total; n++) {
    if (only && !only.includes(n - 1)) continue;
    const page = await pdf.getPage(n);
    const vp = page.getViewport({ scale: 1 });
    const tc = await page.getTextContent({ includeMarkedContent: false });
    // Loading the operator list resolves the fonts, which tells us bold and italic
    try { await page.getOperatorList(); } catch { /* fonts stay unknown */ }
    const fonts = new Map();
    const fontOf = (id) => {
      if (!fonts.has(id)) {
        let info = { family: null, bold: false, italic: false };
        try {
          const f = page.commonObjs.get(id);
          info = fontInfo(f?.name || f?.loadedName || '');
          if (f?.bold || f?.black) info.bold = true;
          if (f?.italic) info.italic = true;
        } catch { /* not loaded */ }
        fonts.set(id, info);
      }
      return fonts.get(id);
    };
    const mul = Util?.transform || ((a, b) => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]]);
    const items = [];
    for (const it of tc.items) {
      if (!it.str) continue;
      const t = mul(vp.transform, it.transform);
      const size = Math.hypot(t[2], t[3]);
      if (!size || Math.abs(t[1]) > Math.abs(t[0])) continue; // skip vertical text
      const f = fontOf(it.fontName);
      items.push({ str: it.str, x: t[4], y: t[5], w: it.width * vp.scale, size, bold: f.bold, italic: f.italic, font: f.family });
    }
    out.push({ width: vp.width, height: vp.height, items, index: n - 1 });
    page.cleanup?.();
    onProgress?.(n, total);
  }
  return out;
}

// --------------------------------------------------------------- Word

const HEADINGS = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
const ALIGN = { center: AlignmentType.CENTER, right: AlignmentType.RIGHT, left: AlignmentType.LEFT };
const twip = (pt) => Math.round(pt * 20);

function runs(list, { heading = false } = {}) {
  return list.filter((r) => r.text).map((r) => new TextRun({
    text: r.text,
    bold: heading ? undefined : r.bold || undefined,
    italics: r.italic || undefined,
    size: heading ? undefined : Math.round(r.size * 2),
    font: heading ? undefined : r.font || undefined,
  }));
}

/**
 * pages: from extractPages() (optionally with OCR words as items).
 * images: { [pageIndex]: [{ png: Uint8Array, width, height, y }] } placed by their top position.
 */
export async function pdfToDocx(pages, { title = 'Converted document', images = {} } = {}) {
  const doc = analyse(pages);
  const first = pages[0] || { width: 595, height: 842 };
  const allX = pages.flatMap((p) => p.items.map((it) => it.x));
  const allX1 = pages.flatMap((p) => p.items.map((it) => it.x + it.w));
  const margin = (v) => Math.max(36, Math.min(108, v));
  const left = margin(allX.length ? Math.min(...allX) : 72);
  const right = margin(allX1.length ? first.width - Math.max(...allX1) : 72);
  const contentWidth = twip(first.width - left - right);

  const children = [];
  doc.pages.forEach((pg, pi) => {
    if (pi > 0) children.push(new Paragraph({ children: [new PageBreak()] }));
    const pics = [...(images[pi] || [])].sort((a, b) => a.y - b.y);
    for (const b of pg.blocks) {
      if (b.type === 'heading') {
        children.push(new Paragraph({ heading: HEADINGS[b.level] || HeadingLevel.HEADING_3, alignment: ALIGN[b.align], children: runs(b.runs, { heading: true }) }));
      } else if (b.type === 'para') {
        children.push(new Paragraph({ alignment: ALIGN[b.align], indent: b.indent ? { firstLine: twip(b.indent) } : undefined, spacing: { after: 120 }, children: runs(b.runs) }));
      } else if (b.type === 'list') {
        for (const item of b.items) {
          children.push(new Paragraph({ ...(b.ordered ? { numbering: { reference: 'numbers', level: 0 } } : { bullet: { level: 0 } }), children: runs(item) }));
        }
      } else if (b.type === 'table') {
        const cols = Math.max(...b.rows.map((r) => r.length));
        const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFC7D5' };
        children.push(new Table({
          width: { size: contentWidth, type: WidthType.DXA },
          columnWidths: Array(cols).fill(Math.floor(contentWidth / cols)),
          rows: b.rows.map((row, ri) => new TableRow({
            tableHeader: ri === 0 && b.header,
            children: Array.from({ length: cols }, (_, ci) => new TableCell({
              borders: { top: border, bottom: border, left: border, right: border },
              shading: ri === 0 && b.header ? { type: ShadingType.CLEAR, fill: 'EEF2FE', color: 'auto' } : undefined,
              children: [new Paragraph({ alignment: parseNumber(row[ci] || '') ? AlignmentType.RIGHT : AlignmentType.LEFT, children: [new TextRun({ text: row[ci] || '', bold: ri === 0 && b.header ? true : undefined })] })],
            })),
          })),
        }));
        children.push(new Paragraph({ children: [] }));
      }
    }
    for (const im of pics) {
      const maxW = first.width - left - right;
      const scale = Math.min(1, maxW / im.width);
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: im.png, transformation: { width: Math.round(im.width * scale * 96 / 72), height: Math.round(im.height * scale * 96 / 72) } })] }));
    }
  });
  if (!children.length) children.push(new Paragraph({ children: [new TextRun('(No text was found in this PDF. If it is a scan, run Recognise text first.)')] }));

  const document = new Document({
    title,
    creator: 'Kertas Office',
    numbering: { config: [{ reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
    styles: { default: { document: { run: { font: 'Calibri', size: Math.round(doc.bodySize * 2) } } } },
    sections: [{
      properties: { page: { size: { width: twip(first.width), height: twip(first.height) }, margin: { left: twip(left), right: twip(right), top: twip(54), bottom: twip(54) } } },
      children,
    }],
  });
  const blob = await Packer.toBlob(document);
  return new Uint8Array(await blob.arrayBuffer());
}

// --------------------------------------------------------------- Excel

/**
 * A workbook snapshot from the pages: tables keep their columns, numbers become numbers.
 * mode: 'sheets' (one sheet per page) | 'single' (everything on one sheet)
 */
export function pdfToSnapshot(pages, { name = 'Converted', mode = pages.length > 30 ? 'single' : 'sheets' } = {}) {
  const snap = blankSnapshot(name);
  const firstId = snap.sheetOrder[0];
  const styles = {};
  const styleFor = (format) => {
    if (!format) return undefined;
    const id = `n${Object.keys(styles).length}`;
    const existing = Object.entries(styles).find(([, s]) => s.n.pattern === format);
    if (existing) return existing[0];
    styles[id] = { n: { pattern: format } };
    return id;
  };
  const fill = (sh, rows, startRow) => {
    rows.forEach((row, r) => row.forEach((text, c) => {
      if (!text) return;
      const num = parseNumber(text);
      const s = num ? styleFor(num.format) : undefined;
      (sh.cellData[startRow + r] ||= {})[c] = num ? { v: num.value, t: 2, ...(s ? { s } : {}) } : { v: text, t: 1 };
    }));
    const widths = {};
    rows.forEach((row) => row.forEach((t, c) => { widths[c] = Math.max(widths[c] || 0, Math.min(60, String(t || '').length)); }));
    for (const [c, w] of Object.entries(widths)) if (w > 10) (sh.columnData[c] ||= {}).w = Math.round(w * 7 + 12);
    return startRow + rows.length;
  };
  if (mode === 'single') {
    const sh = snap.sheets[firstId];
    sh.name = 'Converted';
    let r = 0;
    pages.forEach((p, i) => { r = fill(sh, pageRows(p), r); if (i < pages.length - 1) r += 1; });
    sh.rowCount = Math.max(1000, r + 100);
  } else {
    snap.sheetOrder = [];
    delete snap.sheets[firstId];
    pages.forEach((p, i) => {
      const id = `page${i + 1}`;
      const sh = { id, name: `Page ${(p.index ?? i) + 1}`, rowCount: 1000, columnCount: 26, defaultColumnWidth: 88, defaultRowHeight: 24, cellData: {}, mergeData: [], columnData: {}, rowData: {} };
      fill(sh, pageRows(p), 0);
      snap.sheets[id] = sh;
      snap.sheetOrder.push(id);
    });
  }
  snap.styles = styles;
  return snap;
}
