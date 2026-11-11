// Table designs for the Table tab. Each design is a real Word table style (header row,
// banded rows, first column, total row), added to the document's styles.xml the first time
// it is used, so the look travels with the file and Word, WPS and LibreOffice show it too.

const C = {
  blue: ['2563EB', 'DBEAFE', 'EFF6FF', '93C5FD'],
  green: ['059669', 'D1FAE5', 'ECFDF5', '6EE7B7'],
  orange: ['EA580C', 'FFEDD5', 'FFF7ED', 'FDBA74'],
  purple: ['7C3AED', 'EDE9FE', 'F5F3FF', 'C4B5FD'],
  red: ['E11D48', 'FFE4E6', 'FFF1F2', 'FDA4AF'],
  grey: ['475569', 'E2E8F0', 'F8FAFC', 'CBD5E1'],
};
const NAME = { blue: 'Blue', green: 'Green', orange: 'Orange', purple: 'Purple', red: 'Red', grey: 'Grey' };

/** kind: grid | header (solid header, banded) | light (accent line under header, banded) | minimal (lines between rows) */
export const TABLE_DESIGNS = [
  { id: 'KertasGrid', name: 'Grid', kind: 'grid', color: 'grey' },
  { id: 'KertasMinimal', name: 'Minimal', kind: 'minimal', color: 'grey' },
  ...['blue', 'green', 'orange', 'purple', 'red', 'grey'].map((c) => ({ id: `KertasHeader${NAME[c]}`, name: `Header, ${NAME[c].toLowerCase()}`, kind: 'header', color: c })),
  ...['blue', 'green', 'orange', 'purple', 'red', 'grey'].map((c) => ({ id: `KertasLight${NAME[c]}`, name: `Light, ${NAME[c].toLowerCase()}`, kind: 'light', color: c })),
];

// ---------------------------------------------------------------- Word XML
const el = (name, attributes, elements) => ({ type: 'element', name, ...(attributes ? { attributes } : {}), ...(elements ? { elements } : {}) });
const border = (side, val, sz, color) => el(`w:${side}`, { 'w:val': val, 'w:sz': String(sz), 'w:space': '0', 'w:color': color });
const nil = (side) => el(`w:${side}`, { 'w:val': 'nil' });

function spec(d) {
  const [accent, tint, band, line] = C[d.color];
  // returns plain description used by both the XML and the editor's parsed copy
  const s = { borders: {}, cond: {} };
  if (d.kind === 'grid') {
    for (const k of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) s.borders[k] = ['single', 4, '000000'];
    s.cond.firstRow = { bold: true };
  } else if (d.kind === 'minimal') {
    s.borders.insideH = ['single', 4, 'D1D5DB'];
    s.borders.bottom = ['single', 4, 'D1D5DB'];
    s.cond.firstRow = { bold: true, cellBorders: { bottom: ['single', 12, '334155'] } };
    s.cond.lastRow = { bold: true, cellBorders: { top: ['single', 8, '334155'] } };
  } else if (d.kind === 'header') {
    for (const k of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) s.borders[k] = ['single', 4, line];
    s.cond.firstRow = { bold: true, color: 'FFFFFF', fill: accent, cellBorders: { top: ['single', 4, accent], left: ['single', 4, accent], right: ['single', 4, accent], bottom: ['single', 4, accent] } };
    s.cond.lastRow = { bold: true, cellBorders: { top: ['double', 6, accent] } };
    s.cond.firstCol = { bold: true };
    s.cond.band1Horz = { fill: band };
  } else {
    s.borders.top = ['single', 4, line];
    s.borders.bottom = ['single', 4, line];
    s.borders.insideH = ['single', 4, tint];
    s.cond.firstRow = { bold: true, color: accent, cellBorders: { bottom: ['single', 12, accent] } };
    s.cond.lastRow = { bold: true, cellBorders: { top: ['single', 8, accent] } };
    s.cond.firstCol = { bold: true };
    s.cond.band1Horz = { fill: band };
  }
  return s;
}

function styleXml(d) {
  const s = spec(d);
  const tblBorders = Object.entries(s.borders).map(([k, [v, sz, c]]) => border(k, v, sz, c));
  const cond = Object.entries(s.cond).map(([type, c]) => {
    const kids = [];
    kids.push(el('w:pPr', null, [el('w:spacing', { 'w:before': '0', 'w:after': '0' })]));
    const rPr = [];
    if (c.bold) rPr.push(el('w:b'), el('w:bCs'));
    if (c.color) rPr.push(el('w:color', { 'w:val': c.color }));
    if (rPr.length) kids.push(el('w:rPr', null, rPr));
    kids.push(el('w:tblPr'));
    const tcPr = [];
    if (c.cellBorders) tcPr.push(el('w:tcBorders', null, Object.entries(c.cellBorders).map(([k, [v, sz, col]]) => border(k, v, sz, col))));
    if (c.fill) tcPr.push(el('w:shd', { 'w:val': 'clear', 'w:color': 'auto', 'w:fill': c.fill }));
    if (tcPr.length) kids.push(el('w:tcPr', null, tcPr));
    return el('w:tblStylePr', { 'w:type': type }, kids);
  });
  return el('w:style', { 'w:type': 'table', 'w:customStyle': '1', 'w:styleId': d.id }, [
    el('w:name', { 'w:val': `Kertas ${d.name}` }),
    el('w:basedOn', { 'w:val': 'TableNormal' }),
    el('w:uiPriority', { 'w:val': '40' }),
    el('w:pPr', null, [el('w:spacing', { 'w:after': '0', 'w:line': '240', 'w:lineRule': 'auto' })]),
    el('w:tblPr', null, [
      el('w:tblStyleRowBandSize', { 'w:val': '1' }),
      el('w:tblStyleColBandSize', { 'w:val': '1' }),
      el('w:tblBorders', null, tblBorders.length ? tblBorders : [nil('top')]),
      el('w:tblCellMar', null, [el('w:top', { 'w:w': '57', 'w:type': 'dxa' }), el('w:left', { 'w:w': '108', 'w:type': 'dxa' }), el('w:bottom', { 'w:w': '57', 'w:type': 'dxa' }), el('w:right', { 'w:w': '108', 'w:type': 'dxa' })]),
    ]),
    ...cond,
  ]);
}

// the editor's parsed form of the same style (what the page layout reads)
function styleParsed(d) {
  const s = spec(d);
  const b = ([val, size, color]) => ({ val, size, space: 0, color: `#${color}` });
  const tableStyleProperties = {};
  for (const [type, c] of Object.entries(s.cond)) {
    const p = { type, paragraphProperties: { spacing: { before: 0, after: 0 } } };
    if (c.bold || c.color) p.runProperties = { ...(c.bold ? { bold: true, boldCs: true } : {}), ...(c.color ? { color: { val: c.color } } : {}) };
    const tc = {};
    if (c.cellBorders) tc.borders = Object.fromEntries(Object.entries(c.cellBorders).map(([k, v]) => [k, b(v)]));
    if (c.fill) tc.shading = { val: 'clear', color: 'auto', fill: c.fill };
    if (Object.keys(tc).length) p.tableCellProperties = tc;
    tableStyleProperties[type] = p;
  }
  return {
    type: 'table', styleId: d.id, name: `Kertas ${d.name}`, basedOn: 'TableNormal', uiPriority: 40, customStyle: true,
    paragraphProperties: { spacing: { after: 0, line: 240, lineRule: 'auto' } },
    tableProperties: {
      tableStyleRowBandSize: 1, tableStyleColBandSize: 1,
      borders: Object.fromEntries(Object.entries(s.borders).map(([k, v]) => [k, b(v)])),
      cellMargins: { marginTop: { value: 57, type: 'dxa' }, marginLeft: { value: 108, type: 'dxa' }, marginBottom: { value: 57, type: 'dxa' }, marginRight: { value: 108, type: 'dxa' } },
    },
    tableStyleProperties,
  };
}

/** Adds the design's style to the document if it is not there yet */
export function ensureTableStyle(ed, d) {
  const c = ed?.converter;
  const styles = c?.translatedLinkedStyles?.styles;
  if (!styles) throw new Error('The document is still opening');
  if (styles[d.id]) return d.id;
  const root = c.convertedXml?.['word/styles.xml']?.elements?.find((e) => e.name === 'w:styles');
  if (!root) throw new Error('This document has no style sheet');
  root.elements = root.elements || [];
  root.elements.push(styleXml(d));
  styles[d.id] = styleParsed(d);
  return d.id;
}

/** A small picture of the design for the gallery */
export function designPreview(d) {
  const s = spec(d);
  const W = 64; const H = 44; const rows = 5; const cols = 3;
  const rh = H / rows; const cw = W / cols;
  const col = (v) => `#${v}`;
  let out = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg"><rect width="${W}" height="${H}" fill="#fff"/>`;
  for (let r = 0; r < rows; r++) {
    const cond = r === 0 ? s.cond.firstRow : (r % 2 === 1 ? s.cond.band1Horz : null);
    if (cond?.fill) out += `<rect x="0" y="${r * rh}" width="${W}" height="${rh}" fill="${col(cond.fill)}"/>`;
    for (let k = 0; k < cols; k++) {
      const tc = r === 0 ? (s.cond.firstRow?.color ? col(s.cond.firstRow.color) : '#334155') : '#94A3B8';
      out += `<rect x="${k * cw + 4}" y="${r * rh + rh / 2 - 1}" width="${cw - 9}" height="2" rx="1" fill="${tc}" opacity="${r === 0 ? 1 : 0.8}"/>`;
    }
  }
  const line = (x1, y1, x2, y2, [v, sz, c]) => { out += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${col(c)}" stroke-width="${Math.max(0.8, sz / 6)}"${v === 'double' ? ' stroke-dasharray="0"' : ''}/>`; };
  const bd = s.borders;
  if (bd.top) line(0, 0.5, W, 0.5, bd.top);
  if (bd.bottom) line(0, H - 0.5, W, H - 0.5, bd.bottom);
  if (bd.left) line(0.5, 0, 0.5, H, bd.left);
  if (bd.right) line(W - 0.5, 0, W - 0.5, H, bd.right);
  for (let r = 1; r < rows; r++) if (bd.insideH) line(0, r * rh, W, r * rh, bd.insideH);
  for (let k = 1; k < cols; k++) if (bd.insideV) line(k * cw, 0, k * cw, H, bd.insideV);
  if (s.cond.firstRow?.cellBorders?.bottom) line(0, rh, W, rh, s.cond.firstRow.cellBorders.bottom);
  return `${out}</svg>`;
}
