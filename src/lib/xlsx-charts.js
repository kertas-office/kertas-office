// Charts: our chart model, its Excel (DrawingML) form, and helpers to read
// cell ranges. Pure module, shared by the Sheets UI and the xlsx bridge.
//
// Chart model:
// {
//   id, type: 'column'|'bar'|'line'|'area'|'pie'|'doughnut'|'scatter',
//   title, legend: 'bottom'|'right'|'top'|'left'|'none', stacked: bool,
//   range?: "'Sheet 1'!A1:C5",            // the block it was made from (for editing)
//   byRows?: bool, headers?: bool, labels?: bool,
//   categories?: "Sheet1!$A$2:$A$5",       // x values for scatter
//   series: [{ name?: "Sheet1!$B$1" | 'text', values: "Sheet1!$B$2:$B$5", color? }],
//   anchor: { col, colOff, row, rowOff, width, height }   // pixels
// }
import { parseA1, rangeToA1, colName } from './xlsx-features.js';

export const CHART_TYPES = [
  { id: 'column', label: 'Column' },
  { id: 'bar', label: 'Bar' },
  { id: 'line', label: 'Line' },
  { id: 'area', label: 'Area' },
  { id: 'pie', label: 'Pie' },
  { id: 'doughnut', label: 'Doughnut' },
  { id: 'scatter', label: 'Scatter' },
];

// Bright, distinct colours; the same ones are written into the file so Excel matches
export const PALETTE = ['#4F7DF3', '#22B07D', '#F59E0B', '#EF4444', '#8B5CF6', '#06B6D4', '#EC4899', '#84CC16', '#F97316', '#64748B'];

const EMU = 9525;
const xmlEsc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const xmlUnesc = (s) => String(s).replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi, (m, e) => {
  const map = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  if (map[e.toLowerCase()]) return map[e.toLowerCase()];
  return String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
});

let n = 0;
export const chartId = () => `chart${Date.now().toString(36)}${(n++).toString(36)}`;

// ---------------------------------------------------------------- references

export const quoteSheet = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${String(name).replace(/'/g, "''")}'`);

/** "'My Sheet'!$A$1:$B$4" -> { sheet: 'My Sheet', range } */
export function parseRef(ref) {
  const m = String(ref || '').trim().replace(/^=/, '').match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
  if (!m) { const r = parseA1(ref); return r ? { sheet: null, range: r } : null; }
  const range = parseA1(m[3]);
  return range ? { sheet: (m[1] || m[2]).replace(/''/g, "'"), range } : null;
}

export const makeRef = (sheet, range) => `${quoteSheet(sheet)}!${rangeToA1(range, true)}`;

const sheetByName = (snapshot, name) => Object.values(snapshot.sheets || {}).find((s) => s.name === name);

function cellText(snapshot, cd) {
  if (!cd) return '';
  if (cd.v !== undefined && cd.v !== null) return cd.v;
  const ds = cd.p?.body?.dataStream;
  return ds ? ds.replace(/[\r\n]+$/, '') : '';
}

/** Cell values of a reference, read from a Univer snapshot (row by row) */
export function refValues(snapshot, ref, fallbackSheet) {
  const p = parseRef(ref);
  if (!p) return [];
  const sh = sheetByName(snapshot, p.sheet || fallbackSheet) || (!p.sheet && Object.values(snapshot.sheets)[0]);
  if (!sh) return [];
  const out = [];
  for (let r = p.range.startRow; r <= p.range.endRow; r++) {
    for (let c = p.range.startColumn; c <= p.range.endColumn; c++) out.push(cellText(snapshot, sh.cellData?.[r]?.[c]));
  }
  return out;
}

const isNum = (v) => v !== '' && v !== null && v !== undefined && typeof v !== 'boolean' && Number.isFinite(Number(v));

/**
 * Splits a block of cells into series, the way Excel's Insert Chart does.
 * grid: 2D array of values for range. Returns { categories, series, headers, labels, byRows }
 */
export function seriesFromGrid(sheet, range, grid, opts = {}) {
  const rows = grid.length;
  const cols = rows ? grid[0].length : 0;
  // A header row: the first row has text where the data columns are
  const headers = opts.headers ?? (rows > 1 && grid[0].slice(cols > 1 ? 1 : 0).some((v) => v !== '' && !isNum(v)));
  // Category labels: the first column is text below the header
  const labels = opts.labels ?? (cols > 1 && grid.slice(headers ? 1 : 0).some((r) => r[0] !== '' && !isNum(r[0])));
  const dataRows = rows - (headers ? 1 : 0);
  const dataCols = cols - (labels ? 1 : 0);
  // Excel's rule: more rows than columns means one series per column
  const byRows = opts.byRows ?? (dataCols > dataRows);
  const r0 = range.startRow + (headers ? 1 : 0);
  const c0 = range.startColumn + (labels ? 1 : 0);
  const series = [];
  let categories;
  if (!byRows) {
    if (labels) categories = makeRef(sheet, { startRow: r0, startColumn: range.startColumn, endRow: range.endRow, endColumn: range.startColumn });
    for (let c = c0; c <= range.endColumn; c++) {
      series.push({
        ...(headers ? { name: makeRef(sheet, { startRow: range.startRow, startColumn: c, endRow: range.startRow, endColumn: c }) } : { name: `Series ${series.length + 1}` }),
        values: makeRef(sheet, { startRow: r0, startColumn: c, endRow: range.endRow, endColumn: c }),
      });
    }
  } else {
    if (headers) categories = makeRef(sheet, { startRow: range.startRow, startColumn: c0, endRow: range.startRow, endColumn: range.endColumn });
    for (let r = r0; r <= range.endRow; r++) {
      series.push({
        ...(labels ? { name: makeRef(sheet, { startRow: r, startColumn: range.startColumn, endRow: r, endColumn: range.startColumn }) } : { name: `Series ${series.length + 1}` }),
        values: makeRef(sheet, { startRow: r, startColumn: c0, endRow: r, endColumn: range.endColumn }),
      });
    }
  }
  return { categories, series, headers, labels, byRows };
}

/** Values to draw: { labels: [], datasets: [{ name, values, color }] } using a value reader */
export function resolveChart(model, read) {
  const isRef = (s) => typeof s === 'string' && /![$A-Z]/i.test(s);
  const labels = model.categories ? read(model.categories).map((v) => (v === '' ? '' : String(v))) : null;
  const datasets = (model.series || []).map((s, i) => {
    const name = s.name ? (isRef(s.name) ? String(read(s.name)[0] ?? '') : s.name) : `Series ${i + 1}`;
    return { name, values: read(s.values).map((v) => (isNum(v) ? Number(v) : null)), color: s.color || PALETTE[i % PALETTE.length] };
  });
  const len = Math.max(0, ...datasets.map((d) => d.values.length));
  const all = labels || Array.from({ length: len }, (_, i) => String(i + 1));
  // Rows with no label and no numbers (blank rows inside the range) are left out
  const keep = all.map((l, i) => (labels ? l !== '' : true) || datasets.some((d) => d.values[i] !== null && d.values[i] !== undefined));
  return {
    labels: all.filter((_, i) => keep[i]),
    datasets: datasets.map((d) => ({ ...d, values: d.values.filter((_, i) => keep[i]) })),
  };
}

// ---------------------------------------------------------------- Excel chart XML

const hex6 = (c) => String(c || '#4F7DF3').replace('#', '').toUpperCase();

function strCache(vals) {
  return `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${xmlEsc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
}
function numCache(vals) {
  return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (isNum(v) ? `<c:pt idx="${i}"><c:v>${Number(v)}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
}
const strRef = (f, vals) => `<c:strRef><c:f>${xmlEsc(f)}</c:f>${strCache(vals)}</c:strRef>`;
const numRef = (f, vals) => `<c:numRef><c:f>${xmlEsc(f)}</c:f>${numCache(vals)}</c:numRef>`;

function richText(text, size = 1400) {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${size}" b="0"/></a:pPr><a:r><a:rPr lang="en-US" sz="${size}" b="0"/><a:t>${xmlEsc(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

const LEGEND_POS = { bottom: 'b', right: 'r', top: 't', left: 'l' };
const LEGEND_BACK = { b: 'bottom', r: 'right', t: 'top', l: 'left', tr: 'right' };

/** Our model -> c:chartSpace XML. read(ref) returns cell values for caches. */
export function chartXml(model, read) {
  const type = model.type || 'column';
  const isRef = (s) => typeof s === 'string' && /![$A-Z]/i.test(s);
  const catVals = model.categories ? read(model.categories).map((v) => String(v ?? '')) : null;
  const sers = (model.series || []).map((s, i) => {
    const color = hex6(s.color || PALETTE[i % PALETTE.length]);
    const tx = s.name ? (isRef(s.name) ? `<c:tx>${strRef(s.name, read(s.name).map(String))}</c:tx>` : `<c:tx><c:v>${xmlEsc(s.name)}</c:v></c:tx>`) : '';
    const vals = read(s.values);
    const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
    if (type === 'pie' || type === 'doughnut') {
      const pts = vals.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex6(PALETTE[k % PALETTE.length])}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
      return `<c:ser>${head}${pts}${catVals ? `<c:cat>${strRef(model.categories, catVals)}</c:cat>` : ''}<c:val>${numRef(s.values, vals)}</c:val></c:ser>`;
    }
    if (type === 'scatter') {
      const ln = `<c:spPr><a:ln w="25400" cap="rnd"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>`;
      const x = model.categories ? `<c:xVal>${numRef(model.categories, read(model.categories))}</c:xVal>` : '';
      return `<c:ser>${head}${ln}${x}<c:yVal>${numRef(s.values, vals)}</c:yVal><c:smooth val="0"/></c:ser>`;
    }
    if (type === 'line') {
      const sp = `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>`;
      return `<c:ser>${head}${sp}${catVals ? `<c:cat>${strRef(model.categories, catVals)}</c:cat>` : ''}<c:val>${numRef(s.values, vals)}</c:val><c:smooth val="0"/></c:ser>`;
    }
    const fill = `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
    const inv = type === 'column' || type === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
    return `<c:ser>${head}${fill}${inv}${catVals ? `<c:cat>${strRef(model.categories, catVals)}</c:cat>` : ''}<c:val>${numRef(s.values, vals)}</c:val></c:ser>`;
  }).join('');

  const AX1 = 50010;
  const AX2 = 50020;
  const axes = (catPos, valPos, scatter = false) => {
    const first = scatter
      ? `<c:valAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
      : `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
    const grid = '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E5E7EB"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>';
    return `${first}<c:valAx><c:axId val="${AX2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${grid}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${AX1}"/><c:crosses val="autoZero"/><c:crossBetween val="${scatter ? 'midCat' : 'between'}"/></c:valAx>`;
  };
  const axIds = `<c:axId val="${AX1}"/><c:axId val="${AX2}"/>`;
  const grouping = model.stacked ? 'stacked' : null;
  let plot;
  switch (type) {
    case 'bar':
    case 'column':
      plot = `<c:barChart><c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping || 'clustered'}"/><c:varyColors val="0"/>${sers}<c:gapWidth val="${model.stacked ? 60 : 120}"/>${model.stacked ? '<c:overlap val="100"/>' : '<c:overlap val="-10"/>'}${axIds}</c:barChart>${type === 'bar' ? axes('l', 'b') : axes('b', 'l')}`;
      break;
    case 'line':
      plot = `<c:lineChart><c:grouping val="${grouping || 'standard'}"/><c:varyColors val="0"/>${sers}<c:marker val="1"/>${axIds}</c:lineChart>${axes('b', 'l')}`;
      break;
    case 'area':
      plot = `<c:areaChart><c:grouping val="${grouping || 'standard'}"/><c:varyColors val="0"/>${sers}${axIds}</c:areaChart>${axes('b', 'l')}`;
      break;
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
      break;
    case 'scatter':
      plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}${axIds}</c:scatterChart>${axes('b', 'l', true)}`;
      break;
    default:
      throw new Error(`Unknown chart type ${type}`);
  }
  const title = model.title ? `<c:title>${richText(model.title)}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = model.legend === 'none' ? '' : `<c:legend><c:legendPos val="${LEGEND_POS[model.legend] || 'b'}"/><c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>`;
}

const tagAttr = (xml, tag, attr = 'val') => {
  const m = xml.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*\\b${attr}="([^"]*)"`));
  return m ? xmlUnesc(m[1]) : undefined;
};
const refIn = (xml, tag) => {
  const m = xml.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${tag}>`));
  if (!m) return undefined;
  const f = m[1].match(/<(?:\w+:)?f\b[^>]*>([\s\S]*?)<\/(?:\w+:)?f>/);
  if (f) return xmlUnesc(f[1]).replace(/^\(|\)$/g, '');
  const v = m[1].match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
  return v ? xmlUnesc(v[1]) : undefined;
};
const allText = (xml) => [...xml.matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((m) => xmlUnesc(m[1])).join('');

/** c:chartSpace XML -> our model (without id/anchor), or null if the chart kind is not supported */
export function parseChartXml(xml) {
  const plot = xml.match(/<(?:\w+:)?plotArea\b[^>]*>([\s\S]*?)<\/(?:\w+:)?plotArea>/)?.[1] || '';
  const kinds = { barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter' };
  let type = null;
  let body = '';
  for (const [tag, t] of Object.entries(kinds)) {
    const m = plot.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${tag}>`));
    if (m) { type = t; body = m[1]; break; }
  }
  if (!type) return null;
  if (type === 'column' && tagAttr(body, 'barDir') === 'bar') type = 'bar';
  const grouping = tagAttr(body, 'grouping');
  const series = [];
  let categories;
  for (const s of body.matchAll(/<(?:\w+:)?ser\b[^>]*>([\s\S]*?)<\/(?:\w+:)?ser>/g)) {
    const x = s[1];
    const name = refIn(x, 'tx');
    const values = refIn(x, type === 'scatter' ? 'yVal' : 'val');
    if (!values) continue;
    const cat = refIn(x, type === 'scatter' ? 'xVal' : 'cat');
    if (cat && !categories) categories = cat;
    const color = x.match(/<(?:\w+:)?spPr\b[^>]*>(?:(?!<\/(?:\w+:)?spPr>)[\s\S])*?<(?:\w+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/)?.[1];
    series.push({ ...(name ? { name } : {}), values, ...(color && type !== 'pie' && type !== 'doughnut' ? { color: `#${color.toUpperCase()}` } : {}) });
  }
  if (!series.length) return null;
  const titleXml = xml.match(/<(?:\w+:)?chart\b[^>]*>[\s\S]*?<(?:\w+:)?title\b[^>]*>([\s\S]*?)<\/(?:\w+:)?title>/)?.[1];
  const titleOutsidePlot = titleXml && xml.indexOf(titleXml) < xml.indexOf('plotArea');
  let title = titleOutsidePlot ? allText(titleXml) : '';
  if (!title && titleOutsidePlot && tagAttr(xml, 'autoTitleDeleted') !== '1' && series.length === 1 && series[0].name && !/!/.test(series[0].name)) title = series[0].name;
  const legendXml = xml.match(/<(?:\w+:)?legend\b[^>]*>([\s\S]*?)<\/(?:\w+:)?legend>/)?.[1];
  const legend = legendXml ? LEGEND_BACK[tagAttr(legendXml, 'legendPos')] || 'right' : 'none';
  return { type, title, legend, stacked: grouping === 'stacked' || grouping === 'percentStacked', categories, series };
}

// ---------------------------------------------------------------- packaging

function drawingAnchor(c, idx, rId) {
  const a = c.anchorCells;
  return `<xdr:twoCellAnchor><xdr:from><xdr:col>${a.from.col}</xdr:col><xdr:colOff>${a.from.colOff * EMU}</xdr:colOff><xdr:row>${a.from.row}</xdr:row><xdr:rowOff>${a.from.rowOff * EMU}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${a.to.col}</xdr:col><xdr:colOff>${a.to.colOff * EMU}</xdr:colOff><xdr:row>${a.to.row}</xdr:row><xdr:rowOff>${a.to.rowOff * EMU}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${idx}" name="Chart ${idx}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const relsDoc = (inner) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;
function nextRid(relsXml) {
  let max = 0;
  for (const m of relsXml.matchAll(/Id="rId(\d+)"/g)) max = Math.max(max, +m[1]);
  return `rId${max + 1}`;
}
function addOverride(ct, part, type) {
  if (ct.includes(`PartName="/${part}"`)) return ct;
  return ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}
// Elements that must come after <drawing> inside <worksheet>
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/**
 * Adds charts to a written xlsx.
 * list: [{ sheetPath, xml, anchorCells: { from: {col,colOff,row,rowOff}, to } }]  (offsets in px)
 */
export async function injectCharts(zip, list, relsOf) {
  if (!list.length) return;
  let ct = await zip.file('[Content_Types].xml').async('string');
  const used = (re) => Object.keys(zip.files).map((f) => f.match(re)).filter(Boolean).map((m) => +m[1]);
  let chartNo = Math.max(0, ...used(/^xl\/charts\/chart(\d+)\.xml$/));
  let drawingNo = Math.max(0, ...used(/^xl\/drawings\/drawing(\d+)\.xml$/));
  const bySheet = new Map();
  for (const c of list) { if (!bySheet.has(c.sheetPath)) bySheet.set(c.sheetPath, []); bySheet.get(c.sheetPath).push(c); }

  for (const [sheetPath, charts] of bySheet) {
    const sheetDir = sheetPath.replace(/[^/]+$/, '');
    const sheetRelsPath = `${sheetDir}_rels/${sheetPath.slice(sheetDir.length)}.rels`;
    const rels = await relsOf(zip, sheetPath);
    let drawingPath = rels.find((r) => r.type === 'drawing' && !r.external)?.target;
    if (!drawingPath) {
      drawingNo += 1;
      drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
      zip.file(drawingPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>');
      ct = addOverride(ct, drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
      let sr = (await zip.file(sheetRelsPath)?.async('string')) || relsDoc('');
      const rid = nextRid(sr);
      sr = sr.replace('</Relationships>', `<Relationship Id="${rid}" Type="${REL}/drawing" Target="../drawings/drawing${drawingNo}.xml"/></Relationships>`);
      zip.file(sheetRelsPath, sr);
      let sx = await zip.file(sheetPath).async('string');
      if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace(/<worksheet\b/, `<worksheet xmlns:r="${REL}"`);
      const el = `<drawing r:id="${rid}"/>`;
      const at = AFTER_DRAWING.map((t) => sx.search(new RegExp(`<${t}\\b`))).filter((i) => i >= 0).sort((a, b) => a - b)[0];
      sx = at !== undefined ? sx.slice(0, at) + el + sx.slice(at) : sx.replace('</worksheet>', `${el}</worksheet>`);
      zip.file(sheetPath, sx);
    }
    const drawingDir = drawingPath.replace(/[^/]+$/, '');
    const drawingRelsPath = `${drawingDir}_rels/${drawingPath.slice(drawingDir.length)}.rels`;
    let dr = (await zip.file(drawingRelsPath)?.async('string')) || relsDoc('');
    let dx = await zip.file(drawingPath).async('string');
    let shapeId = 1000 + Math.max(0, ...[...dx.matchAll(/cNvPr id="(\d+)"/g)].map((m) => +m[1] % 1000));
    const prefix = dx.match(/<(\w+):wsDr\b/)?.[1];
    for (const c of charts) {
      chartNo += 1;
      const chartPath = `xl/charts/chart${chartNo}.xml`;
      zip.file(chartPath, c.xml);
      ct = addOverride(ct, chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const rid = nextRid(dr);
      dr = dr.replace('</Relationships>', `<Relationship Id="${rid}" Type="${REL}/chart" Target="../charts/chart${chartNo}.xml"/></Relationships>`);
      shapeId += 1;
      let anchor = drawingAnchor(c, shapeId, rid);
      if (prefix !== 'xdr') {
        // A drawing written with another prefix (or none): declare ours on the anchor
        anchor = anchor.replace('<xdr:twoCellAnchor>', '<xdr:twoCellAnchor xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">');
      }
      if (!/xmlns:a=/.test(dx) && prefix === 'xdr') dx = dx.replace(/<xdr:wsDr\b/, '<xdr:wsDr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"');
      dx = dx.replace(new RegExp(`</${prefix ? `${prefix}:` : ''}wsDr>`), `${anchor}$&`);
    }
    zip.file(drawingPath, dx);
    zip.file(drawingRelsPath, dr);
  }
  zip.file('[Content_Types].xml', ct);
}

export { colName };
