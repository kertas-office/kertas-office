// What is inside an Office file, so Kertas can say before saving what it cannot keep.
//
// detectFeatures(bytes, kind) looks through a .docx / .xlsx package for kinds of content
// (macros, SmartArt, embedded objects...) and returns the ones found. The app compares that
// with what each editor keeps (KEEPS below, worked out by opening and re-saving real files,
// see test/fidelity.test.js) and warns about the rest when the file is saved.
import JSZip from 'jszip';

// part: a file inside the package; xml: text in one of the main parts
const WORD = [
  { id: 'macros', label: 'Macros (VBA code)', part: /^word\/vbaProject\.bin$/i },
  { id: 'smartart', label: 'SmartArt diagrams', part: /^word\/diagrams\/data\d*\.xml$/i },
  { id: 'chart', label: 'Charts', part: /^word\/charts\/chart\d*\.xml$/i },
  { id: 'ole', label: 'Embedded objects (such as a spreadsheet inside the document)', part: /^word\/embeddings\//i, xml: /<o:OLEObject\b/ },
  { id: 'textbox', label: 'Text boxes', xml: /<w:txbxContent\b/ },
  { id: 'shape', label: 'Drawn shapes', xml: /<wps:wsp\b|<v:(oval|rect|roundrect|line|polyline|shape)\b(?![^>]*o:ole)/ },
  { id: 'formfield', label: 'Form fields (check boxes, text fields)', xml: /<w:ffData\b|<w14:checkbox\b/ },
  { id: 'activex', label: 'ActiveX controls', part: /^word\/activeX\//i },
  { id: 'signature', label: 'Digital signature', part: /^_xmlsignatures\//i },
  { id: 'fonts', label: 'Embedded fonts', part: /^word\/fonts\/.+\.(odttf|ttf)$/i },
  { id: 'comment', label: 'Comments', part: /^word\/comments\.xml$/i },
  { id: 'footnote', label: 'Footnotes', xml: /<w:footnoteReference\b/ },
  { id: 'toc', label: 'Table of contents', xml: /TOC \\\\?o|docPartGallery w:val="Table of Contents"/ },
  { id: 'link', label: 'Links', xml: /<w:hyperlink\b/ },
];
const SHEET = [
  { id: 'macros', label: 'Macros (VBA code)', part: /^xl\/vbaProject\.bin$/i },
  { id: 'pivot', label: 'PivotTables', part: /^xl\/pivotTables\/pivotTable\d*\.xml$/i },
  { id: 'chart', label: 'Charts', part: /^xl\/charts\/chart\d*\.xml$/i },
  { id: 'chart-other', label: 'Radar, stock, bubble and 3-D charts', chartTypes: /<c:(radarChart|stockChart|bubbleChart|surfaceChart|surface3DChart|ofPieChart|bar3DChart|line3DChart|pie3DChart|area3DChart)\b/ },
  { id: 'image', label: 'Pictures', xml: /<xdr:pic\b/, where: /^xl\/drawings\/drawing\d*\.xml$/i },
  { id: 'shape', label: 'Shapes and text boxes', xml: /<xdr:sp\b/, where: /^xl\/drawings\/drawing\d*\.xml$/i },
  { id: 'control', label: 'Buttons and form controls', part: /^xl\/(ctrlProps|activeX)\//i, xml: /<(control|formControlPr)\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'table', label: 'Excel tables (Format as Table)', part: /^xl\/tables\/table\d*\.xml$/i },
  { id: 'slicer', label: 'Excel slicers and timelines', part: /^xl\/(slicers|slicerCaches|timelines|timelineCaches)\//i },
  { id: 'extlink', label: 'Links to other workbooks', part: /^xl\/externalLinks\/externalLink\d*\.xml$/i },
  { id: 'connection', label: 'Data connections and queries', part: /^xl\/(connections\.xml|queryTables\/)/i },
  { id: 'threaded', label: 'Threaded comments (replies)', part: /^xl\/threadedComments\//i },
  { id: 'comment', label: 'Notes', part: /^xl\/comments\d*\.xml$/i },
  { id: 'names', label: 'Named ranges', xml: /<definedName\b(?![^>]*_xlnm\.)/, where: /^xl\/workbook\.xml$/i },
  { id: 'filter', label: 'Filters', xml: /<autoFilter\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'validation', label: 'Drop-down lists and input rules', xml: /<dataValidation\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'cf', label: 'Conditional formatting', xml: /<conditionalFormatting\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'protect', label: 'Sheet protection', xml: /<sheetProtection\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'hidden', label: 'Hidden sheets', xml: /<sheet\b[^>]*state="(hidden|veryHidden)"/, where: /^xl\/workbook\.xml$/i },
  { id: 'link', label: 'Links', xml: /<hyperlink\b/, where: /^xl\/worksheets\/sheet\d*\.xml$/i },
  { id: 'signature', label: 'Digital signature', part: /^_xmlsignatures\//i },
];
const WORD_XML = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

/** Kinds of content each editor keeps when it saves, found by opening and re-saving real
 *  files (test/fidelity.test.js); anything else that is found is mentioned before saving */
export const KEEPS = {
  writer: new Set(['smartart', 'chart', 'ole', 'textbox', 'shape', 'formfield', 'fonts', 'comment', 'footnote', 'toc', 'link']),
  sheets: new Set(['chart', 'pivot', 'image', 'comment', 'names', 'filter', 'validation', 'cf', 'protect', 'hidden', 'link', 'table']),
};

async function texts(zip, test) {
  const out = [];
  for (const name of Object.keys(zip.files)) if (test.test(name) && !zip.files[name].dir) out.push(await zip.file(name).async('string'));
  return out;
}

/** The kinds of content in a .docx/.xlsx package: [{ id, label, count }] */
export async function detectFeatures(bytes, kind) {
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { return []; }
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const list = kind === 'writer' ? WORD : kind === 'sheets' ? SHEET : [];
  const found = [];
  const cache = new Map();
  const xmlOf = async (re) => { const k = String(re); if (!cache.has(k)) cache.set(k, await texts(zip, re)); return cache.get(k); };
  for (const f of list) {
    let count = 0;
    if (f.part) count += names.filter((n) => f.part.test(n)).length;
    if (f.xml) {
      const parts = await xmlOf(f.where || WORD_XML);
      const g = new RegExp(f.xml.source, 'g');
      for (const t of parts) count += (t.match(g) || []).length;
    }
    if (f.chartTypes) {
      const charts = await xmlOf(/^xl\/charts\/chart\d*\.xml$/i);
      count += charts.filter((t) => f.chartTypes.test(t)).length;
    }
    if (count) found.push({ id: f.id, label: f.label, count });
  }
  return found;
}

/** What would be lost by saving a file with these features in this editor */
export function atRisk(features, kind) {
  const keeps = KEEPS[kind] || new Set();
  return features.filter((f) => !keeps.has(f.id));
}
