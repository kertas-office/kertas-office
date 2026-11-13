// Sheets: spreadsheet editing powered by Univer (Apache-2.0),
// with our own ExcelJS bridge for .xlsx and .csv files, an Excel-style
// ribbon, live charts and PivotTable summaries.
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { FormulaDataModel } from '@univerjs/engine-formula';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import filterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import '@univerjs/preset-sheets-filter/lib/index.css';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import '@univerjs/preset-sheets-sort/lib/index.css';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import cfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import dvEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import findEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import linkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import '@univerjs/preset-sheets-hyper-link/lib/index.css';
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note';
import noteEnUS from '@univerjs/preset-sheets-note/locales/en-US';
import '@univerjs/preset-sheets-note/lib/index.css';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import drawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import '@univerjs/preset-sheets-drawing/lib/index.css';
import {
  xlsxToSnapshot, snapshotToXlsx, csvToSnapshot, snapshotToCsv, blankSnapshot,
} from '../lib/xlsx-bridge.js';
import { snapshotToPdf } from '../lib/sheet-pdf.js';
import { univerFonts } from '../lib/font-assets.js';
import { FONT_CATALOG, findFont } from '../lib/font-catalog.js';
import { createRibbon } from '../lib/ribbon.js';
import { seriesFromGrid, parseRef, makeRef, refValues, CHART_TYPES } from '../lib/xlsx-charts.js';
import { rangeToA1 } from '../lib/xlsx-features.js';
import { parseTitleRows } from '../lib/xlsx-print.js';
import { buildPivot, AGGREGATES, normSpec, fieldNames } from '../lib/pivot.js';
import { makeSparklines, SPARK_TYPES, sparkValues, DEFAULT_SPARK } from '../lib/sparklines.js';
import { createSparkLayer } from './sheet-sparklines.js';
import { createSlicerLayer } from './sheet-slicers.js';
import { openDialog } from '../lib/dialog.js';

// Kertas's wording for the name box and the named ranges panel
// (locales merge one level deep only, so the rest of sheets-ui's wording is carried over)
const NAME_WORDS = {
  'sheets-ui': { ...sheetsEnUS['sheets-ui'], definedName: {
    managerTitle: 'Named ranges',
    managerDescription: 'To name cells: select them, type a name in this box and press Enter. Names can be used in formulas instead of cell addresses.',
    addButton: 'New name', featureTitle: 'Named ranges', ratioRange: 'Range', ratioFormula: 'Formula', confirm: 'Save', cancel: 'Cancel',
    scopeWorkbook: 'Whole workbook', inputNamePlaceholder: 'Name (letters, numbers and underscores, no spaces)', inputCommentPlaceholder: 'Comment (optional)',
    formulaOrRefStringInvalid: 'That is not a cell range or formula', defaultName: 'MyRange', updateButton: 'Save', deleteButton: 'Delete',
    deleteConfirmText: 'Delete this name? Formulas that use it will show #NAME?.',
  } },
};
import { parseA1 } from '../lib/xlsx-features.js';
import { createChartLayer, renderChartPng } from './sheet-charts.js';
import {
  chartDialog, pivotDialog, functionDialog, sizeDialog, cfRuleDialog, removeDuplicatesDialog, FUNCTIONS,
} from './sheets-dialogs.js';

const MUTATION = 2;
const CHART_ICON = { column: 'chart-column', bar: 'chart-bar', line: 'chart-line', area: 'chart-area', pie: 'chart-pie', doughnut: 'chart-donut', scatter: 'chart-scatter' };

// Cell styles (Home > Cell styles), the same set Excel and WPS offer
const CELL_STYLES = [
  { id: 'normal', label: 'Normal', apply: (r) => { r.clearFormat(); } },
  { id: 'good', label: 'Good', bg: '#C6EFCE', fg: '#006100' },
  { id: 'bad', label: 'Bad', bg: '#FFC7CE', fg: '#9C0006' },
  { id: 'neutral', label: 'Neutral', bg: '#FFEB9C', fg: '#9C5700' },
  { id: 'title', label: 'Title', fg: '#1F2937', bold: true, size: 18 },
  { id: 'h1', label: 'Heading 1', fg: '#1E3A8A', bold: true, size: 15, bottom: ['#4F7DF3', 'THICK'] },
  { id: 'h2', label: 'Heading 2', fg: '#1E3A8A', bold: true, size: 13, bottom: ['#A5B8F8', 'MEDIUM'] },
  { id: 'total', label: 'Total', bold: true, top: ['#4F7DF3', 'THIN'], bottom: ['#4F7DF3', 'DOUBLE'] },
  { id: 'input', label: 'Input', bg: '#FFCC99', fg: '#3F3F76' },
  { id: 'calc', label: 'Calculation', bg: '#F2F2F2', fg: '#FA7D00', bold: true },
  { id: 'note', label: 'Note', bg: '#FFFFCC', fg: '#1F2937' },
  { id: 'currency', label: 'Currency', format: '#,##0.00' },
  { id: 'percent', label: 'Percent', format: '0%' },
  { id: 'comma', label: 'Comma', format: '#,##0' },
];

// Format as table: header colour and the light band colour
const TABLE_STYLES = [
  { id: 'blue', label: 'Blue', head: '#4F7DF3', band: '#EEF2FE' },
  { id: 'green', label: 'Green', head: '#16A34A', band: '#ECFDF3' },
  { id: 'orange', label: 'Orange', head: '#EA7317', band: '#FFF4EA' },
  { id: 'purple', label: 'Purple', head: '#7C3AED', band: '#F3EEFE' },
  { id: 'teal', label: 'Teal', head: '#0D9488', band: '#E8F7F5' },
  { id: 'gray', label: 'Gray', head: '#475569', band: '#F1F5F9' },
];

const CF_STYLES = {
  red: { bg: '#FFC7CE', fg: '#9C0006' }, yellow: { bg: '#FFEB9C', fg: '#9C5700' }, green: { bg: '#C6EFCE', fg: '#006100' },
  redText: { fg: '#C00000' }, bold: { bold: true },
};

export async function mount(el, { data, name, dark, setDirty, setStatus, toast, fileMenu }) {
  el.innerHTML = '<div data-ribbon></div><div class="sheets-host"></div>';
  const host = el.querySelector('.sheets-host');
  const ext = name.toLowerCase().split('.').pop();

  let snapshot;
  if (!data) snapshot = blankSnapshot(name.replace(/\.[^.]+$/, ''));
  else if (ext === 'csv') {
    // UTF-8 when the bytes are valid UTF-8, else the Windows code page Excel uses for "CSV" files,
    // so accented letters are not turned into question marks (and then saved that way)
    let text; let legacy = false;
    try { text = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(data); } catch { text = new TextDecoder('windows-1252').decode(data); legacy = true; }
    snapshot = csvToSnapshot(text, name);
    // such a file is saved as UTF-8 with a byte-order mark, which Excel reads correctly
    if (legacy) snapshot.custom = { ...(snapshot.custom || {}), csvBom: true };
  }
  else snapshot = await xlsxToSnapshot(data, name.replace(/\.[^.]+$/, ''));
  const initialCharts = snapshot.charts || {};
  delete snapshot.charts;
  // PivotTables, sparklines and row/column groups live beside the Univer data
  const extras = snapshot.extras || {};
  delete snapshot.extras;
  const pivots = new Map(Object.entries(extras.pivots || {})); // pivot sheet id -> definition
  const outline = extras.outline || {}; // sheet id -> { rows: { r: level }, cols: { c: level } }
  const initialSparks = extras.sparklines || {};
  const keepParts = extras.keep || {}; // sheet protection and Excel tables, written back on save
  const printSetup = extras.print || {}; // page setup per sheet id (Page Layout tab)

  // The grid is drawn on a canvas, so fonts the workbook uses must be loaded before it draws
  const usedFonts = [...new Set(Object.values(snapshot.styles || {}).map((st) => st?.ff).filter((ff) => ff && findFont(ff)))];
  await Promise.race([
    Promise.all(usedFonts.flatMap((ff) => ['', 'bold ', 'italic '].map((w) => document.fonts?.load(`${w}16px "${ff}"`).catch(() => {})))),
    new Promise((r) => setTimeout(r, 2500)),
  ]);
  // The rest load quietly, so picking one from the list draws right away
  setTimeout(() => FONT_CATALOG.forEach((f) => document.fonts?.load(`16px "${f.family}"`).catch(() => {})), 3000);

  const { univer, univerAPI } = createUniver({
    locale: LocaleType.EN_US,
    darkMode: !!dark,
    locales: { [LocaleType.EN_US]: mergeLocales(sheetsEnUS, filterEnUS, sortEnUS, cfEnUS, dvEnUS, findEnUS, linkEnUS, noteEnUS, drawingEnUS, NAME_WORDS) },
    presets: [
      UniverSheetsCorePreset({
        container: host,
        header: true,
        toolbar: false, // the Home tab has the formatting tools
        customFontFamily: { override: true, list: univerFonts() },
        footer: { sheetBar: true, statisticBar: true, zoomSlider: true },
        formulaBar: true,
        workerURL: new Worker(new URL('./univer.worker.js', import.meta.url), { type: 'module' }),
      }),
      UniverSheetsFilterPreset(),
      UniverSheetsSortPreset(),
      UniverSheetsConditionalFormattingPreset(),
      UniverSheetsDataValidationPreset(),
      UniverSheetsFindReplacePreset(),
      UniverSheetsHyperLinkPreset(),
      UniverSheetsNotePreset(),
      UniverSheetsDrawingPreset(),
    ],
  });

  const workbook = univerAPI.createWorkbook(snapshot);

  // Formulas are worked out in the background. Saving (or printing) waits for that to finish, so
  // the file never holds answers from before the last edit: Kertas and other apps show the
  // answers stored in a file when they open it.
  let calcRunning = false;
  let calcWaiters = [];
  const calcDone = () => { calcRunning = false; const w = calcWaiters; calcWaiters = []; w.forEach((f) => f()); };
  try {
    const engine = univerAPI.getFormula();
    engine.calculationStart(() => { calcRunning = true; });
    engine.calculationEnd(() => calcDone());
  } catch (e) { console.warn('formula events', e); }
  async function settleFormulas() {
    await new Promise((r) => setTimeout(r, 80)); // an edit starts its calculation on the next tick
    if (!calcRunning) return;
    setStatus('Finishing calculations...');
    await new Promise((resolve) => { const t = setTimeout(resolve, 120000); calcWaiters.push(() => { clearTimeout(t); resolve(); }); });
    try { await univerAPI.getFormula().onCalculationResultApplied(5000); } catch { /* applied already */ }
  }
  const markDirty = () => setDirty(true);

  // ---------------------------------------------------------------- helpers
  const ws = () => workbook.getActiveSheet();
  const selection = () => ws().getActiveRange() || workbook.getActiveRange() || ws().getRange(0, 0);
  const run = (id, params) => univerAPI.executeCommand(id, params);
  // ---- Home tab formatting (Univer's own commands; the controls follow the active cell)
  const NUMBER_FORMATS = [
    { label: 'General', pattern: 'General' }, { label: 'Number', pattern: '0.00' }, { label: 'Comma', pattern: '#,##0.00' }, { label: 'Currency', pattern: '"$"#,##0.00' },
    { label: 'Accounting', pattern: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)' }, { label: 'Short date', pattern: 'dd/mm/yyyy' },
    { label: 'Long date', pattern: 'dddd, d mmmm yyyy' }, { label: 'Time', pattern: 'h:mm:ss AM/PM' }, { label: 'Percentage', pattern: '0.00%' },
    { label: 'Fraction', pattern: '# ?/?' }, { label: 'Scientific', pattern: '0.00E+00' }, { label: 'Text', pattern: '@' },
  ];
  const homeState = { wrap: false };
  async function cmd(id, params) {
    try { await run(id, params); } catch (e) { toast(String(e?.message || e), 'error'); }
    syncHomeSoon();
  }
  function setNumFmt(v) {
    if (v === '__more') { run('sheet.operation.open.numfmt.panel'); return; }
    const f = NUMBER_FORMATS.find((x) => x.label.toLowerCase() === String(v).toLowerCase());
    const pattern = f ? f.pattern : String(v);
    try { selection().setNumberFormat(pattern === 'General' ? '' : pattern); } catch (e) { toast(String(e?.message || e), 'error'); }
    syncHomeSoon();
  }
  function borderMenu() {
    const b = (type, style = 1) => () => cmd('sheet.command.set-border-basic', { value: { type, color: '#000000', style } });
    return [
      { label: 'Bottom border', icon: 'border', action: b('bottom') },
      { label: 'Top border', icon: 'border', action: b('top') },
      { label: 'Left border', icon: 'border', action: b('left') },
      { label: 'Right border', icon: 'border', action: b('right') },
      'sep',
      { label: 'No border', icon: 'close', action: b('none', 0) },
      { label: 'All borders', icon: 'borders-all', action: b('all') },
      { label: 'Outside borders', icon: 'border', action: b('outside') },
      { label: 'Thick outside borders', icon: 'border', action: b('outside', 13) },
      { label: 'Inside borders', icon: 'borders-all', action: b('inside') },
      { label: 'Bottom double border', icon: 'border', action: b('bottom', 7) },
    ];
  }
  function syncHome() {
    let st = {}; let nf = '';
    try { const r = selection(); st = r.getCellStyleData?.() || {}; nf = r.getNumberFormat?.() || ''; homeState.wrap = st.tb === 3; } catch { return; }
    ribbon.setOn('b', st.bl === 1); ribbon.setOn('i', st.it === 1); ribbon.setOn('u', st.ul?.s === 1); ribbon.setOn('s', st.st?.s === 1);
    ribbon.setOn('al', st.ht === 1); ribbon.setOn('ac', st.ht === 2); ribbon.setOn('ar', st.ht === 3);
    ribbon.setOn('vt', st.vt === 1); ribbon.setOn('vm', st.vt === 2); ribbon.setOn('vb', st.vt === 3);
    ribbon.setOn('wrap', st.tb === 3);
    ribbon.setValue('fontname', st.ff || 'Arial');
    ribbon.setValue('fontsize', String(st.fs || 11));
    const f = NUMBER_FORMATS.find((x) => x.pattern === nf) || (!nf || nf === 'General' ? NUMBER_FORMATS[0] : null);
    ribbon.setValue('numfmt', f ? f.label : nf);
  }
  let homeTimer = null;
  const syncHomeSoon = () => { if (homeTimer) return; homeTimer = requestAnimationFrame(() => { homeTimer = null; syncHome(); }); };
  const bounds = (r) => { const g = r.getRange(); return { startRow: g.startRow, startColumn: g.startColumn, endRow: g.endRow, endColumn: g.endColumn }; };
  const single = (r) => { const b = bounds(r); return b.startRow === b.endRow && b.startColumn === b.endColumn; };
  /** The selection, or the block of data around the cursor when one cell is selected */
  const region = () => {
    const r = selection();
    if (!single(r)) return r;
    const g = r.getDataRegion();
    return g && !single(g) ? g : r;
  };
  // raw cell values: getValues() returns what a formatted cell shows ("38,000" rather than 38000)
  const values = (r) => (r.getRawValues?.() || r.getValues()).map((row) => row.map((v) => (v === null || v === undefined ? '' : v)));
  const looksLikeHeader = (grid) => grid.length > 1 && grid[0].every((v) => v !== '' && typeof v === 'string') && grid[1].some((v) => typeof v === 'number');
  const busy = (msg, fn) => { try { return fn(); } catch (e) { console.error(e); toast(`${msg}: ${e.message || e}`, 'error'); return null; } };

  // ---------------------------------------------------------------- charts
  const charts = createChartLayer({ univerAPI, workbook, dark, hooks: { onEdit: editChart, onChange: markDirty } });
  // Charts need the grid to be drawn first
  const loadCharts = (tries = 0) => {
    try { charts.load(initialCharts); } catch (e) {
      if (tries < 10) setTimeout(() => loadCharts(tries + 1), 200); else console.error(e);
    }
  };
  if (Object.keys(initialCharts).length) setTimeout(() => loadCharts(), 50);

  const sparks = createSparkLayer({ univerAPI, workbook, hooks: { onChange: markDirty } });
  if (Object.keys(initialSparks).length) setTimeout(() => { try { sparks.load(initialSparks); } catch (e) { console.warn('sparklines', e); } }, 120);
  const slicers = createSlicerLayer({ univerAPI, workbook, toast: (m, kind) => toast(m, kind), hooks: { onChange: markDirty } });
  if (Object.keys(extras.slicers || {}).length) setTimeout(() => { try { slicers.load(extras.slicers); } catch (e) { console.warn('slicers', e); } }, 160);
  const sheetById = (sid) => workbook.getSheets().find((x) => x.getSheetId() === sid);
  for (const def of pivots.values()) def.sourceId ||= workbook.getSheetByName(def.source?.sheet || '')?.getSheetId();

  function chartFrom(opts, sheetForRange) {
    const p = parseRef(opts.range);
    if (!p) throw new Error('That range is not valid');
    const sheetName = p.sheet || sheetForRange;
    const src = workbook.getSheetByName(sheetName);
    if (!src) throw new Error(`There is no sheet called ${sheetName}`);
    const grid = values(src.getRange(rangeToA1(p.range)));
    const s = seriesFromGrid(sheetName, p.range, grid, { headers: opts.headers, labels: opts.labels, byRows: opts.byRows });
    if (!s.series.length) throw new Error('No numbers to plot in that range');
    return {
      type: opts.type, title: opts.title, legend: opts.legend, stacked: !!opts.stacked,
      range: makeRef(sheetName, p.range).replace(/\$/g, ''), byRows: s.byRows, headers: s.headers, labels: s.labels,
      categories: s.categories, series: s.series,
    };
  }

  async function insertChart(type = 'column') {
    const sheet = ws();
    const r = region();
    const b = bounds(r);
    const grid = values(r);
    const guess = seriesFromGrid(sheet.getSheetName(), b, grid);
    const opts = await chartDialog({
      type, range: makeRef(sheet.getSheetName(), b).replace(/\$/g, ''),
      headers: guess.headers, labels: guess.labels, byRows: guess.byRows, legend: type === 'pie' || type === 'doughnut' ? 'right' : 'bottom',
      title: guess.headers && guess.series.length === 1 ? String(grid[0][b.endColumn - b.startColumn] ?? '') : '',
    });
    if (!opts) return;
    busy('Could not make that chart', () => {
      const model = chartFrom(opts, sheet.getSheetName());
      model.anchor = { col: b.endColumn + 2, colOff: 0, row: b.startRow, rowOff: 0, width: 480, height: 300 };
      charts.add(sheet.getSheetId(), model);
      toast('Chart added. Drag it to move, drag a corner to resize.');
    });
  }

  async function editChart(model) {
    if (!model) return;
    const opts = await chartDialog(model, { editing: true });
    if (!opts) return;
    busy('Could not update the chart', () => charts.update(model.id, chartFrom(opts, ws().getSheetName())));
  }

  // ---------------------------------------------------------------- pivot
  /** Writes a pivot's cells and styling onto its sheet; returns the area used */
  /** Number format of each source column (from its first data cell), so dates stay dates */
  function sourceFormats(sheet, range) {
    const out = [];
    for (let c = range.startColumn; c <= range.endColumn; c++) {
      let f = '';
      try { f = sheet.getRange(Math.min(range.startRow + 1, range.endRow), c).getNumberFormat() || ''; } catch { /* none */ }
      out.push(f === 'General' ? '' : f);
    }
    return out;
  }
  function writePivot(out, pv, spec, formats = []) {
    const width = Math.max(...pv.cells.map((row) => row.length));
    const rows = pv.cells.map((row) => Array.from({ length: width }, (_, j) => {
      const c = row[j];
      return c ? (c.f ?? c.v) : null;
    }));
    out.getRange(0, 0, rows.length, width).setValues(rows);
    out.getRange(0, 0).setFontWeight('bold').setFontSize(13);
    const heads = (pv.headerRows || 2) - 1; // header rows under the title
    const lead = pv.rowFields || 1; // row-field columns
    out.getRange(1, 0, heads, width).setFontWeight('bold').setBackground('#4F7DF3').setFontColor('#FFFFFF');
    out.getRange(rows.length - 1, 0, 1, width).setFontWeight('bold').setBackground('#EEF2FE');
    const body = rows.length - 1 - heads; // data rows and the total row
    // labels keep their source format (a date stays a date), values keep theirs unless counted
    (pv.rowCols || []).forEach((src, j) => { if (formats[src] && rows.length - 2 - heads > 0) out.getRange(1 + heads, j, rows.length - 2 - heads, 1).setNumberFormat?.(formats[src]); });
    (pv.colRows || []).forEach((src, k) => { if (formats[src] && width > lead) out.getRange(1 + k, lead, 1, width - lead).setNumberFormat?.(formats[src]); });
    (pv.valueCols || []).forEach((v, j) => {
      const f = v.agg === 'count' ? '' : v.agg === 'average' ? (formats[v.fld] && !/[dmyhs]/i.test(formats[v.fld]) ? formats[v.fld] : '#,##0.00') : formats[v.fld];
      if (f && body > 0) out.getRange(1 + heads, lead + j, body, 1).setNumberFormat?.(f);
    });
    for (let j = 0; j < width; j++) out.setColumnWidth(j, j < lead ? 150 : 120);
    out.setFrozenRows?.(1 + heads);
    return { rows: rows.length, cols: width };
  }
  const pivotName = () => { const used = new Set([...pivots.values()].map((p) => p.name)); let n = 1; while (used.has(`PivotTable${n}`)) n++; return `PivotTable${n}`; };

  async function insertPivot() {
    const sheet = ws();
    const r = region();
    const b = bounds(r);
    const grid = values(r);
    if (grid.length < 2 || grid[0].length < 2) { toast('Select a table with a header row and at least two columns first', 'error'); return; }
    const fields = fieldNames(grid[0], b.startColumn);
    const spec = await pivotDialog(fields, rangeToA1(b));
    if (!spec) return;
    busy('Could not make the PivotTable', () => {
      const pv = buildPivot(grid, { sheet: sheet.getSheetName(), startRow: b.startRow, startColumn: b.startColumn }, spec);
      const names = new Set(workbook.getSheets().map((x) => x.getSheetName()));
      let nm = 'Pivot'; for (let i = 2; names.has(nm); i++) nm = `Pivot ${i}`;
      const out = workbook.insertSheet(nm);
      const size = writePivot(out, pv, spec, sourceFormats(sheet, b));
      pivots.set(out.getSheetId(), { name: pivotName(), sourceId: sheet.getSheetId(), source: { sheet: sheet.getSheetName(), range: b }, spec, fields, size });
      workbook.setActiveSheet(out);
      markDirty();
      toast(`${pv.label} by ${normSpec(spec).rows.map((i) => fields[i]).join(' and ')} is on the "${nm}" sheet. It saves as a real PivotTable for Excel.`);
    });
  }

  /** Rebuilds a pivot from its source, picking up new rows and new categories */
  function refreshPivot(pid) {
    const def = pivots.get(pid);
    const out = sheetById(pid);
    if (!def || !out) return false;
    if (!def.spec) throw new Error(`"${out.getSheetName()}" uses pivot settings Kertas cannot rebuild (it summarises a named range, or uses a summary other than sum, count, average, max or min)`);
    const src = (def.sourceId && sheetById(def.sourceId)) || workbook.getSheetByName(def.source.sheet);
    if (!src) throw new Error(`The source sheet "${def.source.sheet}" is no longer in the workbook`);
    def.sourceId = src.getSheetId();
    def.source.sheet = src.getSheetName();
    const range = { ...def.source.range };
    const width = range.endColumn - range.startColumn + 1;
    // Rows typed straight under the data are part of it
    for (let k = 0; k < 200000; k++) {
      const next = values(src.getRange(range.endRow + 1, range.startColumn, 1, width))[0];
      if (!next.some((v) => v !== '')) break;
      range.endRow++;
    }
    def.source.range = range;
    const grid = values(src.getRange(rangeToA1(range)));
    const pv = buildPivot(grid, { sheet: def.source.sheet, startRow: range.startRow, startColumn: range.startColumn }, def.spec);
    const old = def.size || { rows: 0, cols: 0 };
    const clearRows = Math.max(old.rows, pv.cells.length) + 2;
    const clearCols = Math.max(old.cols, ...pv.cells.map((row) => row.length));
    const area = out.getRange(0, 0, clearRows, clearCols);
    area.clear ? area.clear() : area.setValues(Array.from({ length: clearRows }, () => Array(clearCols).fill(null)));
    def.size = writePivot(out, pv, def.spec, sourceFormats(src, range));
    def.fields = fieldNames(grid[0], range.startColumn);
    def.imported = false;
    return true;
  }
  function refreshPivots() {
    const active = ws().getSheetId();
    const ids = pivots.has(active) ? [active] : [...pivots.keys()].filter((id) => sheetById(id));
    if (!ids.length) { toast('There are no PivotTables in this workbook yet'); return; }
    busy('Could not refresh', () => {
      let n = 0;
      for (const id of ids) if (refreshPivot(id)) n++;
      markDirty();
      toast(n === 1 ? 'PivotTable refreshed from its data' : `${n} PivotTables refreshed`);
    });
  }

  // ---------------------------------------------------------------- sparklines
  async function insertSparklines(type = 'line') {
    const sheet = ws();
    const r = region();
    const b = bounds(r);
    const dataRef = rangeToA1(b);
    const loc = b.endColumn > b.startColumn
      ? { startRow: b.startRow, endRow: b.endRow, startColumn: b.endColumn + 1, endColumn: b.endColumn + 1 }
      : { startRow: b.endRow + 1, endRow: b.endRow + 1, startColumn: b.startColumn, endColumn: b.endColumn };
    const res = await openDialog({
      title: 'Sparklines',
      description: 'Small charts inside cells, one for each row (or column) of the data.',
      width: 460,
      fields: [
        { id: 'type', label: 'Type', type: 'segmented', value: type, options: SPARK_TYPES.map((t) => ({ value: t.id, label: t.label })) },
        { id: 'data', label: 'Data range', value: dataRef, placeholder: 'B2:G10' },
        { id: 'loc', label: 'Put them in', value: rangeToA1(loc), placeholder: 'H2:H10' },
        { id: 'color', label: 'Colour', type: 'color', value: DEFAULT_SPARK.color },
        { id: 'neg', label: 'Negative / low point', type: 'color', value: DEFAULT_SPARK.negColor },
        { id: 'markers', label: 'Show every point', type: 'checkbox', value: false, show: (v) => v.type === 'line' },
      ],
      validate: (v) => {
        const d = parseA1(v.data.replace(/^.*!/, '')); const l = parseA1(v.loc.replace(/^.*!/, ''));
        if (!d) return 'The data range is not valid';
        if (!l) return 'The location range is not valid';
        if (l.startRow !== l.endRow && l.startColumn !== l.endColumn) return 'Put sparklines in one row or one column';
        return null;
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    busy('Could not add sparklines', () => {
      const list = makeSparklines({ sheet: sheet.getSheetName(), range: parseA1(v.data.replace(/^.*!/, '')) }, parseA1(v.loc.replace(/^.*!/, '')), { type: v.type, color: v.color, negColor: v.neg, markers: !!v.markers });
      sparks.add(sheet.getSheetId(), list);
      toast(`${list.length} sparkline${list.length === 1 ? '' : 's'} added`);
    });
  }
  function clearSparklines() {
    const n = sparks.removeIn(ws().getSheetId(), bounds(selection()));
    toast(n ? `Removed ${n} sparkline${n === 1 ? '' : 's'}` : 'There are no sparklines in the selection');
  }
  const sparkMenu = () => [
    ...SPARK_TYPES.map((t) => ({ label: t.label, icon: 'sparkline', hint: t.id === 'stacked' ? 'Up or down for each value' : null, action: () => insertSparklines(t.id) })),
    'sep',
    { label: 'Colour of selected sparklines...', icon: 'palette', action: recolourSparklines },
    { label: 'Clear sparklines in the selection', icon: 'clear', action: clearSparklines },
  ];
  async function recolourSparklines() {
    const g = bounds(selection());
    const res = await openDialog({ title: 'Sparkline colour', fields: [{ id: 'color', label: 'Colour', type: 'color', value: DEFAULT_SPARK.color }, { id: 'type', label: 'Type', type: 'segmented', value: 'keep', options: [{ value: 'keep', label: 'Keep' }, ...SPARK_TYPES.map((t) => ({ value: t.id, label: t.label }))] }], actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }] });
    if (!res) return;
    const n = sparks.styleIn(ws().getSheetId(), g, { color: res.values.color, ...(res.values.type !== 'keep' ? { type: res.values.type } : {}) });
    toast(n ? `Updated ${n} sparkline${n === 1 ? '' : 's'}` : 'Select the cells with the sparklines first');
  }

  // ---------------------------------------------------------------- row and column groups (outline)
  const outlineOf = (sid) => (outline[sid] ||= { rows: {}, cols: {} });
  function group(axis, delta) {
    const sheet = ws();
    const b = bounds(selection());
    const o = outlineOf(sheet.getSheetId());
    const map = axis === 'rows' ? (o.rows ||= {}) : (o.cols ||= {});
    const [from, to] = axis === 'rows' ? [b.startRow, b.endRow] : [b.startColumn, b.endColumn];
    if (to - from > 5000) { toast('Select fewer rows or columns to group', 'error'); return; }
    let changed = 0;
    for (let i = from; i <= to; i++) {
      const lv = Math.max(0, Math.min(7, (map[i] || 0) + delta));
      if (lv === (map[i] || 0)) continue;
      changed++;
      if (lv) map[i] = lv; else delete map[i];
      if (delta < 0 && !lv) (axis === 'rows' ? sheet.showRows(i, 1) : sheet.showColumns(i, 1));
    }
    if (!changed) { toast(delta > 0 ? 'Those are grouped as deeply as Excel allows (7 levels)' : 'Nothing to ungroup there'); return; }
    markDirty();
    const what = axis === 'rows' ? `row${to > from ? 's' : ''} ${from + 1}${to > from ? `-${to + 1}` : ''}` : `column${to > from ? 's' : ''} ${colLabel(from)}${to > from ? `-${colLabel(to)}` : ''}`;
    toast(delta > 0 ? `Grouped ${what}. Use Collapse to hide them.` : `Ungrouped ${what}`);
  }
  const colLabel = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  /** Runs of grouped rows/columns at level >= 1: [{ axis, start, end, level }] */
  function groupsOf(sid) {
    const out = [];
    const o = outline[sid] || {};
    for (const axis of ['rows', 'cols']) {
      const keys = Object.keys(o[axis] || {}).map(Number).sort((a, b) => a - b);
      let start = null; let prev = null;
      for (const k of keys) {
        if (start === null) { start = k; prev = k; continue; }
        if (k === prev + 1) { prev = k; continue; }
        out.push({ axis, start, end: prev }); start = k; prev = k;
      }
      if (start !== null) out.push({ axis, start, end: prev });
    }
    return out;
  }
  const isHidden = (sheet, axis, i) => {
    try { const w = sheet.getSheet(); return axis === 'rows' ? !w.getRowRawVisible(i) : !w.getColVisible(i); } catch { return false; }
  };
  function setGroup(g, collapse) {
    const sheet = ws();
    const n = g.end - g.start + 1;
    if (g.axis === 'rows') (collapse ? sheet.hideRows(g.start, n) : sheet.showRows(g.start, n));
    else (collapse ? sheet.hideColumns(g.start, n) : sheet.showColumns(g.start, n));
    markDirty();
  }
  function collapseAt(collapse) {
    const sheet = ws();
    const b = bounds(selection());
    const groups = groupsOf(sheet.getSheetId());
    // The group holding the cursor, or the one just above / left of it (where Excel puts the button)
    const hit = groups.find((g) => (g.axis === 'rows' ? b.startRow >= g.start - 1 && b.startRow <= g.end + 1 : b.startColumn >= g.start - 1 && b.startColumn <= g.end + 1));
    if (!hit) { toast('Put the cursor in or next to a group first (Data > Group)'); return; }
    setGroup(hit, collapse);
  }
  function allGroups(collapse) {
    const groups = groupsOf(ws().getSheetId());
    if (!groups.length) { toast('This sheet has no groups yet'); return; }
    for (const g of groups) setGroup(g, collapse);
  }
  const groupsMenu = () => {
    const sheet = ws();
    const groups = groupsOf(sheet.getSheetId());
    const label = (g) => (g.axis === 'rows' ? `Rows ${g.start + 1}-${g.end + 1}` : `Columns ${colLabel(g.start)}-${colLabel(g.end)}`);
    return [
      { label: 'Group rows', icon: 'group', hint: 'The selected rows', action: () => group('rows', 1) },
      { label: 'Group columns', icon: 'group', hint: 'The selected columns', action: () => group('cols', 1) },
      { label: 'Ungroup rows', icon: 'ungroup', action: () => group('rows', -1) },
      { label: 'Ungroup columns', icon: 'ungroup', action: () => group('cols', -1) },
      ...(groups.length ? ['sep', { header: 'Groups on this sheet (tick = showing)' }, ...groups.map((g) => {
        const shown = !isHidden(sheet, g.axis, g.start);
        return { label: label(g), checked: shown, action: () => setGroup(g, shown) };
      })] : []),
    ];
  };

  // Rows and columns inserted or removed move groups and sparklines with them
  const shiftSub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => {
    const m = /^sheet\.mutation\.(insert|remove)-(row|rows|col)$/.exec(e.id || '');
    if (!m) return;
    const p = e.params || {};
    const range = p.range || p.ranges?.[0];
    const sid = p.subUnitId;
    if (!range || !sid) return;
    const axis = m[2].startsWith('row') ? 'rows' : 'cols';
    const at = axis === 'rows' ? range.startRow : range.startColumn;
    const count = (axis === 'rows' ? range.endRow - range.startRow : range.endColumn - range.startColumn) + 1;
    const sign = m[1] === 'insert' ? 1 : -1;
    const o = outline[sid];
    if (o?.[axis]) {
      const next = {};
      for (const [k, lv] of Object.entries(o[axis])) {
        const i = Number(k);
        if (i < at) next[i] = lv;
        else if (sign > 0) next[i + count] = lv;
        else if (i >= at + count) next[i - count] = lv;
      }
      o[axis] = next;
    }
    sparks.shift?.(sid, axis, at, count * sign);
  });

  // ---------------------------------------------------------------- cells
  function autoSum(fn = 'SUM') {
    busy('AutoSum', () => {
      const sheet = ws();
      const r = selection();
      const b = bounds(r);
      if (!single(r)) {
        // Totals under each selected column, like Excel
        const target = b.endRow + 1;
        for (let c = b.startColumn; c <= b.endColumn; c++) {
          sheet.getRange(target, c).setValue(`=${fn}(${rangeToA1({ startRow: b.startRow, startColumn: c, endRow: b.endRow, endColumn: c })})`);
        }
        return;
      }
      // One cell: add up the numbers right above it (or to its left)
      const col = b.startColumn;
      const row = b.startRow;
      const num = (v) => typeof v === 'number';
      let top = row;
      if (row > 0) {
        const above = values(sheet.getRange(Math.max(0, row - 500), col, Math.min(500, row), 1)).map((x) => x[0]).reverse();
        let k = 0; while (k < above.length && num(above[k])) k++;
        top = row - k;
      }
      let left = col;
      if (top === row && col > 0) {
        const before = values(sheet.getRange(row, Math.max(0, col - 200), 1, Math.min(200, col)))[0].reverse();
        let k = 0; while (k < before.length && num(before[k])) k++;
        left = col - k;
      }
      const ref = top < row ? rangeToA1({ startRow: top, startColumn: col, endRow: row - 1, endColumn: col })
        : left < col ? rangeToA1({ startRow: row, startColumn: left, endRow: row, endColumn: col - 1 }) : '';
      r.setValue(`=${fn}(${ref})`);
    });
  }

  // ---------------------------------------------------------------- named ranges
  const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}$/;
  const looksLikeCell = (n) => /^[A-Za-z]{1,3}\d{1,7}$/.test(n) || /^R\d*C\d*$/i.test(n);
  async function defineName() {
    const r = selection();
    const sheetName = ws().getSheetName();
    const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
    const ref = `${quoted}!${rangeToA1(bounds(r), true)}`;
    // a label in the cell above or to the left is a good suggestion, as in Excel
    const b = bounds(r);
    const textAt = (row, col) => { const x = row >= 0 && col >= 0 ? ws().getRange(row, col).getValue() : ''; return typeof x === 'string' ? x : ''; };
    const near = textAt(b.startRow - 1, b.startColumn) || textAt(b.startRow, b.startColumn - 1) || (b.endRow > b.startRow ? textAt(b.startRow, b.startColumn) : '');
    const suggest = String(near || '').trim().replace(/[^A-Za-z0-9_.]+/g, '_').replace(/^(\d)/, '_$1').slice(0, 40);
    const taken = new Set(workbook.getDefinedNames().map((d) => d.getName().toLowerCase()));
    const res = await openDialog({
      title: 'Define name',
      description: 'A name for cells, to use in formulas, for example =SUM(Sales) instead of =SUM(B2:B20).',
      width: 460,
      fields: [
        { id: 'name', label: 'Name', value: suggest && !looksLikeCell(suggest) ? suggest : '' },
        { id: 'scope', label: 'Works in', type: 'select', value: 'workbook', options: [{ value: 'workbook', label: 'The whole workbook' }, { value: 'sheet', label: `Only the sheet "${sheetName}"` }] },
        { id: 'ref', label: 'Refers to', value: `=${ref}` },
        { id: 'comment', label: 'Comment (optional)', value: '' },
      ],
      validate: (v) => {
        const n = v.name.trim();
        if (!NAME_RE.test(n)) return 'Start with a letter or _, then letters, numbers, _ or . (no spaces)';
        if (looksLikeCell(n) || /^(true|false)$/i.test(n)) return 'That looks like a cell address or a value; choose another name';
        if (taken.has(n.toLowerCase())) return `"${n}" is already a name. Use Name manager to change it.`;
        if (!v.ref.trim().replace(/^=/, '')) return 'Say which cells (or formula) the name stands for';
        return null;
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Define', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    busy('Define name', () => {
      let bld = workbook.newDefinedNameBuilder().setName(v.name.trim()).setRef(v.ref.trim().replace(/^=/, ''));
      if (v.comment.trim()) bld = bld.setComment(v.comment.trim());
      bld = v.scope === 'sheet' ? bld.setScopeToWorksheet(ws()) : bld.setScopeToWorkbook();
      workbook.insertDefinedNameBuilder(bld.build());
      markDirty();
      toast(`"${v.name.trim()}" now stands for ${v.ref.trim().replace(/^=/, '')}`);
    });
  }
  /** Names to put into the selected cell's formula */
  function useNameMenu() {
    const names = workbook.getDefinedNames().filter((d) => !d.isHidden?.()).map((d) => d.getName()).sort((a, b) => a.localeCompare(b));
    if (!names.length) return [{ label: 'No named ranges yet', icon: 'bookmark', disabled: true }, 'sep', { label: 'Define name...', icon: 'plus', action: () => defineName() }];
    return [
      ...names.map((n) => ({ label: n, icon: 'bookmark', action: () => busy('Use name', () => {
        const r = selection(); const cell = ws().getRange(r.getRow(), r.getColumn());
        const cur = String(cell.getFormula?.() || '');
        // into the formula already there (at its end), or a new formula
        cell.setValue(cur.startsWith('=') ? `${cur}${/[(,+\-*/^&=<>]$/.test(cur) ? '' : '+'}${n}` : `=${n}`);
        markDirty();
      }) })),
      'sep',
      { label: 'Define name...', icon: 'plus', action: () => defineName() },
    ];
  }

  async function insertFunction(preset) {
    const r = selection();
    const a1 = rangeToA1({ ...bounds(r), endRow: r.getRow(), endColumn: r.getColumn() });
    const f = await functionDialog(a1, '', preset);
    if (f) busy('Insert function', () => { ws().getRange(r.getRow(), r.getColumn()).setValue(f); markDirty(); });
  }

  function applyCellStyle(st) {
    busy('Cell style', () => {
      const r = selection();
      if (st.apply) { st.apply(r); return; }
      if (st.bg) r.setBackground(st.bg);
      if (st.fg) r.setFontColor(st.fg);
      if (st.bold) r.setFontWeight('bold');
      if (st.size) r.setFontSize(st.size);
      if (st.format) r.setNumberFormat?.(st.format);
      const B = univerAPI.Enum.BorderType;
      const S = univerAPI.Enum.BorderStyleTypes;
      if (st.top) r.setBorder(B.TOP, S[st.top[1]], st.top[0]);
      if (st.bottom) r.setBorder(B.BOTTOM, S[st.bottom[1]], st.bottom[0]);
    });
  }

  function formatAsTable(t) {
    busy('Format as table', () => {
      const sheet = ws();
      const r = region();
      const b = bounds(r);
      if (b.endRow === b.startRow) { toast('Select a table with a header row first', 'error'); return; }
      const cols = b.endColumn - b.startColumn + 1;
      sheet.getRange(b.startRow, b.startColumn, 1, cols).setBackground(t.head).setFontColor('#FFFFFF').setFontWeight('bold');
      for (let i = b.startRow + 1; i <= b.endRow; i++) {
        sheet.getRange(i, b.startColumn, 1, cols).setBackground((i - b.startRow) % 2 === 0 ? t.band : '#FFFFFF');
      }
      const B = univerAPI.Enum.BorderType;
      const S = univerAPI.Enum.BorderStyleTypes;
      r.setBorder(B.ALL, S.THIN, '#D8DEE9');
      if (!sheet.getFilter?.()) { try { r.createFilter(); } catch { /* a filter elsewhere on the sheet */ } }
      markDirty();
    });
  }

  function cfAdd(build) {
    busy('Conditional formatting', () => {
      const sheet = ws();
      const rule = build(sheet.newConditionalFormattingRule()).setRanges([selection().getRange()]).build();
      sheet.addConditionalFormattingRule(rule);
      markDirty();
    });
  }
  async function cfHighlight(kind) {
    const v = await cfRuleDialog(kind);
    if (!v) return;
    const st = CF_STYLES[v.style];
    cfAdd((b) => {
      let x = kind === 'greaterThan' ? b.whenNumberGreaterThan(v.a)
        : kind === 'lessThan' ? b.whenNumberLessThan(v.a)
          : kind === 'between' ? b.whenNumberBetween(Math.min(v.a, v.b), Math.max(v.a, v.b))
            : kind === 'equal' ? b.whenNumberEqualTo(v.a) : b.whenTextContains(v.a);
      if (st.bg) x = x.setBackground(st.bg);
      if (st.fg) x = x.setFontColor(st.fg);
      if (st.bold) x = x.setBold(true);
      return x;
    });
  }
  const scale = (colors) => cfAdd((b) => b.setColorScale([
    { index: 0, color: colors[0], value: { type: 'min' } },
    { index: 1, color: colors[1], value: { type: 'percentile', value: 50 } },
    { index: 2, color: colors[2], value: { type: 'max' } },
  ]));
  const bar = (color) => cfAdd((b) => b.setDataBar({ min: { type: 'min' }, max: { type: 'max' }, positiveColor: color, nativeColor: '#EF4444', isGradient: true, isShowValue: true }));
  const icons = (iconType) => cfAdd((b) => b.setIconSet({
    iconConfigs: [
      { iconType, iconId: '0', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 } },
      { iconType, iconId: '1', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 } },
      { iconType, iconId: '2', operator: 'lessThanOrEqual', value: { type: 'percent', value: 33 } },
    ],
    isShowValue: true,
  }));
  function cfClear(scope) {
    busy('Clear rules', () => {
      const sheet = ws();
      if (scope === 'sheet') { sheet.clearConditionalFormatRules(); return; }
      const s = bounds(selection());
      for (const rule of sheet.getConditionalFormattingRules()) {
        const hit = (rule.ranges || []).some((r) => r.startRow <= s.endRow && r.endRow >= s.startRow && r.startColumn <= s.endColumn && r.endColumn >= s.startColumn);
        if (hit) sheet.deleteConditionalFormattingRule(rule.cfId);
      }
    });
  }

  function sort(ascending) {
    busy('Sort', () => {
      const sel = selection();
      const r = region();
      const b = bounds(r);
      const grid = values(r);
      const header = looksLikeHeader(grid);
      const body = header ? ws().getRange(b.startRow + 1, b.startColumn, b.endRow - b.startRow, b.endColumn - b.startColumn + 1) : r;
      const col = Math.min(Math.max(sel.getColumn() - b.startColumn, 0), b.endColumn - b.startColumn);
      body.sort({ column: col, ascending });
      markDirty();
    });
  }

  async function removeDuplicates() {
    const r = region();
    const b = bounds(r);
    const grid = values(r);
    if (r.getFormulas?.().some((row) => row.some(Boolean))) { toast('Remove duplicates works on plain values. This range has formulas.', 'error'); return; }
    const head = looksLikeHeader(grid);
    const names = grid[0].map((v, i) => (head && v !== '' ? String(v) : `Column ${String.fromCharCode(65 + ((b.startColumn + i) % 26))}`));
    const opt = await removeDuplicatesDialog(names);
    if (!opt) return;
    busy('Remove duplicates', () => {
      const rows = opt.headers ? grid.slice(1) : grid;
      const seen = new Set();
      const keep = rows.filter((row) => {
        const key = JSON.stringify(opt.cols.map((i) => (typeof row[i] === 'string' ? row[i].trim().toLowerCase() : row[i])));
        if (seen.has(key)) return false;
        seen.add(key); return true;
      });
      const removed = rows.length - keep.length;
      const start = b.startRow + (opt.headers ? 1 : 0);
      const width = b.endColumn - b.startColumn + 1;
      if (keep.length) ws().getRange(start, b.startColumn, keep.length, width).setValues(keep);
      if (removed) ws().getRange(start + keep.length, b.startColumn, removed, width).clearContent();
      markDirty();
      toast(removed ? `${removed} duplicate row${removed === 1 ? '' : 's'} removed, ${keep.length} kept` : 'No duplicates found');
    });
  }

  /** Excel's quickest dropdown: the distinct values already in the column become the list */
  function dropdownFromSelection() {
    busy('Dropdown list', () => {
      const r = selection();
      const vals = [...new Set(values(r).flat().filter((v) => v !== '').map(String))].slice(0, 200);
      if (!vals.length) { toast('Type a few values in the cells first, then make the dropdown', 'error'); return; }
      r.setDataValidation(univerAPI.newDataValidation().requireValueInList(vals, true).build());
      markDirty();
      toast(`Dropdown with ${vals.length} choice${vals.length === 1 ? '' : 's'}`);
    });
  }

  function split(kind) {
    busy('Text to columns', () => {
      const D = univerAPI.Enum.SplitDelimiterType;
      selection().splitTextToColumns(true, D[kind]);
      markDirty();
    });
  }

  // ---------------------------------------------------------------- slicers
  async function insertSlicer() {
    const sheet = ws();
    const r = region();
    const b = bounds(r);
    const grid = values(sheet.getRange(b.startRow, b.startColumn, 1, b.endColumn - b.startColumn + 1));
    if (b.endRow <= b.startRow) { toast('Select a table with a header row first, or click inside one', 'error'); return; }
    const heads = fieldNames(grid[0], b.startColumn).slice(0, 40);
    const res = await openDialog({
      title: 'Insert slicers',
      description: `Pick the columns to filter ${rangeToA1(b)} by. Each one gets a panel of buttons; click a button to show only those rows.`,
      fields: heads.map((h, i) => ({ id: `c${i}`, label: h, type: 'checkbox', value: false })),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
      validate: (v, a) => (a === 'ok' && !heads.some((_, i) => v[`c${i}`]) ? 'Pick at least one column' : null),
    });
    if (!res || res.action !== 'ok') return;
    busy('Could not insert the slicer', () => {
      let n = 0;
      heads.forEach((h, i) => {
        if (!res.values[`c${i}`]) return;
        // tall enough for its buttons, up to about eight of them
        const count = new Set(values(sheet.getRange(b.startRow + 1, b.startColumn + i, b.endRow - b.startRow, 1)).map((x) => String(x[0]))).size;
        slicers.add(sheet.getSheetId(), {
          title: h, range: b, col: i, selected: [],
          anchor: { col: b.endColumn + 2, colOff: n * 196, row: b.startRow, rowOff: 0, width: 184, height: Math.min(300, 54 + count * 31) },
        });
        n++;
      });
      toast(n === 1 ? 'Click a button in the slicer to filter the table. Ctrl+click picks several.' : 'Click the buttons in the slicers to filter the table. Ctrl+click picks several.');
    });
  }

  // ---------------------------------------------------------------- page layout
  const setupOf = () => printSetup[ws().getSheetId()] || {};
  function setSetup(patch, msg) {
    const sid = ws().getSheetId();
    printSetup[sid] = { ...(printSetup[sid] || {}), ...patch };
    markDirty();
    syncLayout();
    if (msg) toast(msg);
  }
  const PAPERS = [['a4', 'A4', '210 x 297 mm'], ['letter', 'Letter', '8.5 x 11 in'], ['legal', 'Legal', '8.5 x 14 in'], ['a3', 'A3', '297 x 420 mm'], ['a5', 'A5', '148 x 210 mm']];
  const marginsMenu = () => [['normal', 'Normal', 'Top and bottom 1.9 cm, sides 1.8 cm'], ['narrow', 'Narrow', 'Sides 0.6 cm'], ['wide', 'Wide', '2.5 cm all round']]
    .map(([v, label, hint]) => ({ label, hint, icon: 'margins', checked: (setupOf().margins || 'normal') === v, action: () => setSetup({ margins: v }) }));
  const orientationMenu = () => [['auto', 'Automatic', 'Landscape when the sheet is wide'], ['portrait', 'Portrait'], ['landscape', 'Landscape']]
    .map(([v, label, hint]) => ({ label, hint, icon: v === 'landscape' ? 'landscape' : 'portrait', checked: (setupOf().orientation || 'auto') === v, action: () => setSetup({ orientation: v }) }));
  const paperMenu = () => PAPERS.map(([v, label, hint]) => ({ label, hint, icon: 'paper', checked: (setupOf().paper || 'a4') === v, action: () => setSetup({ paper: v }) }));
  const fitMenu = () => {
    const cur = setupOf().fit || 'width';
    return [
      { label: 'Fit all columns on one page', hint: 'Rows continue on the next pages', icon: 'fit', checked: cur === 'width', action: () => setSetup({ fit: 'width' }) },
      { label: 'Fit sheet on one page', icon: 'fit', checked: cur === 'page', action: () => setSetup({ fit: 'page' }) },
      { label: 'Actual size (100%)', icon: 'fit', checked: cur === 'none', action: () => setSetup({ fit: 'none' }) },
      { label: `Custom${cur === 'custom' ? ` (${setupOf().scale}%)` : ''}...`, icon: 'edit', checked: cur === 'custom', action: () => customScale() },
    ];
  };
  async function customScale() {
    const res = await openDialog({
      title: 'Print scale', width: 360,
      fields: [{ id: 'scale', label: 'Scale (%)', type: 'number', value: setupOf().scale || 100, min: 10, max: 400 }],
      validate: (v) => (!(+v.scale >= 10 && +v.scale <= 400) ? 'Enter a number from 10 to 400' : null),
    });
    if (res?.action === 'ok') setSetup({ fit: 'custom', scale: +res.values.scale });
  }
  const printAreaMenu = () => [
    { label: 'Set print area', hint: 'Print only the selected cells', icon: 'select-all', action: () => { const b = bounds(selection()); setSetup({ area: b }, `Only ${rangeToA1(b)} will be printed`); } },
    { label: 'Clear print area', icon: 'clear', disabled: !setupOf().area, action: () => setSetup({ area: null }, 'The whole sheet will be printed') },
  ];
  const breaksMenu = () => {
    const row = selection().getRow();
    const breaks = setupOf().breaks || [];
    return [
      { label: 'Insert page break', hint: `A new page starts at row ${row + 1}`, icon: 'page-break', disabled: row === 0 || breaks.includes(row), action: () => setSetup({ breaks: [...breaks, row].sort((a, b) => a - b) }, `A new page starts at row ${row + 1}`) },
      { label: 'Remove page break', icon: 'clear', disabled: !breaks.includes(row), action: () => setSetup({ breaks: breaks.filter((b) => b !== row) }) },
      { label: 'Reset all page breaks', icon: 'refresh', disabled: !breaks.length, action: () => setSetup({ breaks: [] }, 'Page breaks removed') },
    ];
  };
  async function printTitles() {
    const cur = setupOf().titleRows;
    const sel = bounds(selection());
    const res = await openDialog({
      title: 'Print titles',
      description: 'Rows printed at the top of every page, such as a header row.',
      width: 420,
      fields: [{ id: 'rows', label: 'Rows to repeat at the top', value: cur ? `${cur[0] + 1}:${cur[1] + 1}` : `${sel.startRow + 1}:${sel.endRow + 1}`, placeholder: '1:1', hint: 'For example 1:1 for the first row, or 1:2 for the first two. Leave empty for none.' }],
      validate: (v) => (v.rows.trim() && !parseTitleRows(v.rows) ? 'Write the rows like 1:1 or 1:2' : null),
    });
    if (res?.action !== 'ok') return;
    const t = res.values.rows.trim() ? parseTitleRows(res.values.rows) : null;
    setSetup({ titleRows: t }, t ? `Row${t[0] === t[1] ? ` ${t[0] + 1}` : `s ${t[0] + 1} to ${t[1] + 1}`} will print at the top of every page` : 'No rows repeat');
  }
  const HF_PRESETS = [
    ['', '(none)'], ['Page {page}', 'Page 1'], ['Page {page} of {pages}', 'Page 1 of 3'], ['{sheet}', 'Sheet name'], ['{file}', 'File name'], ['{date}', 'Date'],
  ];
  async function headerFooter() {
    const cur = setupOf();
    const h = cur.header || ['', '', '']; const f = cur.footer || ['', 'Page {page} of {pages}', ''];
    const res = await openDialog({
      title: 'Header and footer',
      description: 'You can type text and these: {page} {pages} {date} {time} {file} {sheet}',
      width: 560,
      fields: [
        { id: 'hl', label: 'Header: left', value: h[0] }, { id: 'hc', label: 'Header: centre', value: h[1] }, { id: 'hr', label: 'Header: right', value: h[2] },
        { id: 'fl', label: 'Footer: left', value: f[0] }, { id: 'fc', label: 'Footer: centre', value: f[1] }, { id: 'fr', label: 'Footer: right', value: f[2] },
        { id: 'preset', label: 'Quick footer', type: 'select', value: '', options: [{ value: '', label: 'Choose to fill the footer centre...' }, ...HF_PRESETS.slice(1).map(([v, l]) => ({ value: v, label: l }))] },
      ],
      onChange: (v, form) => {
        if (!v.preset) return;
        const fc = form.querySelector('[data-f="fc"]'); fc.value = v.preset;
        form.querySelector('[data-f="preset"]').value = '';
      },
    });
    if (res?.action !== 'ok') return;
    const v = res.values;
    setSetup({ header: [v.hl, v.hc, v.hr], footer: [v.fl, v.fc, v.fr] }, 'Header and footer set');
  }
  function syncLayout() {
    const st = setupOf();
    ribbon.setOn('pl-center', !!st.center);
    ribbon.setOn('pl-grid', st.gridlines ?? printGrid);
    ribbon.setOn('pl-titles', !!st.titleRows);
    ribbon.setOn('pl-area', !!st.area);
  }
  univerAPI.addEvent(univerAPI.Event.ActiveSheetChanged, () => setTimeout(syncLayout, 0));
  const printPreview = () => fileMenu.find((m) => m.id === 'print')?.action?.();

  // ---------------------------------------------------------------- formulas across sheets
  // While a formula is being typed, clicking another sheet's tab shows that sheet so its cells
  // can be picked. Univer keeps the cell being edited floating over the other sheet, which
  // looks like that sheet has the value; as in Excel, it is hidden there and the formula
  // stays in the formula bar until you come back or press Enter.
  let editingSheet = null;
  function editorBits() {
    const out = [];
    for (const fe of el.querySelectorAll('[data-u-comp="formula-editor"]')) {
      const box = fe.closest('.univer-absolute');
      if (!box || box.closest('.univer-formula-bar, [class*="formula-bar"]')) continue;
      out.push(box);
      // the small "'Sheet2'!B3" tag above it
      box.parentElement?.querySelectorAll(':scope > .univer-pointer-events-none.univer-absolute').forEach((t) => out.push(t));
    }
    return out;
  }
  function syncEditorVisibility() {
    const away = !!editingSheet && ws()?.getSheetId() !== editingSheet;
    for (const n of editorBits()) n.style.visibility = away ? 'hidden' : '';
  }
  univerAPI.addEvent(univerAPI.Event.SheetEditStarted, (p) => { editingSheet = p?.worksheet?.getSheetId?.() || ws()?.getSheetId(); syncEditorVisibility(); });
  univerAPI.addEvent(univerAPI.Event.SheetEditEnded, () => { editingSheet = null; syncEditorVisibility(); });
  univerAPI.addEvent(univerAPI.Event.ActiveSheetChanged, () => requestAnimationFrame(syncEditorVisibility));

  // ---------------------------------------------------------------- split panes
  // A split shows two parts of the sheet at once. It uses the frozen-pane drawing, with the top
  // (and left) pane free to scroll on its own with the mouse wheel.
  const splits = new Map(); // sheet id -> { rows, cols }
  function toggleSplit() {
    busy('Split', () => {
      const sheet = ws();
      const sid = sheet.getSheetId();
      if (splits.has(sid)) { splits.delete(sid); sheet.cancelFreeze(); ribbon.setOn('splitpanes', false); return; }
      const sc = sheet.getScrollState?.() || { sheetViewStartRow: 0, sheetViewStartColumn: 0 };
      const f = sheet.getFreeze?.() || {};
      // the first row and column in view (a frozen part counts from where it starts)
      const top = f.ySplit ? Math.max(0, f.startRow - f.ySplit) : (sc.sheetViewStartRow || 0);
      const left = f.xSplit ? Math.max(0, f.startColumn - f.xSplit) : (sc.sheetViewStartColumn || 0);
      const r = selection();
      let row = r.getRow(); let col = r.getColumn();
      // split in the middle of the window when the cursor is at the top-left corner of the view
      if (row <= top && col <= left) row = top + 12;
      const rows = Math.max(0, row - top); const cols = Math.max(0, col - left);
      if (!rows && !cols) { toast('Click a cell lower down or further right to split there', 'error'); return; }
      sheet.cancelFreeze();
      sheet.setFreeze({ xSplit: cols, ySplit: rows, startRow: rows ? row : -1, startColumn: cols ? col : -1 });
      splits.set(sid, { rows, cols });
      ribbon.setOn('splitpanes', true);
      toast('The top and left panes scroll on their own: use the mouse wheel over them. Click Split again to close it.');
    });
  }
  /** Scrolls the top (rows) or left (cols) pane of a split by n cells */
  function scrollPane(sheet, dir, n) {
    const sp = splits.get(sheet.getSheetId());
    const f = sheet.getFreeze();
    const sc = sheet.getScrollState();
    if (!sp || !f) return;
    if (dir === 'rows' && f.ySplit) {
      const mainFirst = sc.sheetViewStartRow + f.ySplit;
      const start = Math.max(f.ySplit, Math.min(mainFirst, f.startRow + n));
      if (start === f.startRow) return;
      sheet.setFreeze({ ...f, startRow: start });
      run('sheet.command.scroll-view', { sheetViewStartRow: mainFirst - f.ySplit, sheetViewStartColumn: sheet.getScrollState().sheetViewStartColumn, offsetX: 0, offsetY: 0 });
    } else if (dir === 'cols' && f.xSplit) {
      const mainFirst = sc.sheetViewStartColumn + f.xSplit;
      const start = Math.max(f.xSplit, Math.min(mainFirst, f.startColumn + n));
      if (start === f.startColumn) return;
      sheet.setFreeze({ ...f, startColumn: start });
      run('sheet.command.scroll-view', { sheetViewStartColumn: mainFirst - f.xSplit, sheetViewStartRow: sheet.getScrollState().sheetViewStartRow, offsetX: 0, offsetY: 0 });
    }
  }
  function onSplitWheel(e) {
    const sheet = ws();
    const sp = sheet && splits.get(sheet.getSheetId());
    if (!sp || e.ctrlKey) return;
    const canvas = e.target.closest?.('canvas');
    if (!canvas) return;
    const box = canvas.getBoundingClientRect();
    const f = sheet.getFreeze();
    const zoom = sheet.getZoom?.() || 1;
    const px = (from, count, size) => { let t = 0; for (let i = from; i < from + count; i++) t += size(i); return t * zoom; };
    const topH = 20 * zoom + px(f.startRow - f.ySplit, f.ySplit, (i) => sheet.getRowHeight(i));
    const leftW = 46 * zoom + px(f.startColumn - f.xSplit, f.xSplit, (i) => sheet.getColumnWidth(i));
    const x = e.clientX - box.left; const y = e.clientY - box.top;
    const vertical = Math.abs(e.deltaY) >= Math.abs(e.deltaX) && !e.shiftKey;
    const step = Math.sign(vertical ? e.deltaY : (e.deltaX || e.deltaY)) * 3;
    if (vertical && f.ySplit && y < topH) { e.preventDefault(); e.stopPropagation(); scrollPane(sheet, 'rows', step); }
    else if (!vertical && f.xSplit && x < leftW) { e.preventDefault(); e.stopPropagation(); scrollPane(sheet, 'cols', step); }
  }
  el.addEventListener('wheel', onSplitWheel, { capture: true, passive: false });
  const splitWheelOff = () => el.removeEventListener('wheel', onSplitWheel, { capture: true });
  // the button follows the sheet that is showing
  univerAPI.addEvent(univerAPI.Event.ActiveSheetChanged || 'ActiveSheetChanged', () => ribbon.setOn('splitpanes', splits.has(ws()?.getSheetId())));

  function freeze(kind) {
    busy('Freeze panes', () => {
      const sheet = ws();
      if (splits.delete(sheet.getSheetId())) ribbon.setOn('splitpanes', false);
      if (kind === 'none') { sheet.cancelFreeze(); return; }
      if (kind === 'row') { sheet.cancelFreeze(); sheet.setFrozenRows(1); return; }
      if (kind === 'col') { sheet.cancelFreeze(); sheet.setFrozenColumns(1); return; }
      const r = selection();
      sheet.cancelFreeze();
      if (r.getRow() > 0) sheet.setFrozenRows(r.getRow());
      if (r.getColumn() > 0) sheet.setFrozenColumns(r.getColumn());
    });
  }

  const zoomBy = (d) => busy('Zoom', () => { const s = ws(); const z = s.getZoom?.() || 1; s.zoom(Math.min(4, Math.max(0.1, +(d === 0 ? 1 : z + d).toFixed(2)))); });
  const rowsCols = (fn) => busy('Rows and columns', () => {
    const r = selection();
    const b = bounds(r);
    fn(ws(), r, b.startRow, b.endRow - b.startRow + 1, b.startColumn, b.endColumn - b.startColumn + 1);
    markDirty();
  });

  async function setSize(kind) {
    const sheet = ws();
    const r = selection();
    const cur = kind === 'row' ? sheet.getRowHeight(r.getRow()) : sheet.getColumnWidth(r.getColumn());
    const px = await sizeDialog(kind, Math.round(cur));
    if (!px) return;
    rowsCols((s, _r, row, nr, col, nc) => (kind === 'row' ? s.setRowHeights(row, nr, px) : s.setColumnWidths(col, nc, px)));
  }

  // ---------------------------------------------------------------- ribbon
  let printGrid = true;
  const cfMenu = () => [
    { header: 'Highlight cells' },
    { label: 'Greater than...', icon: 'up', action: () => cfHighlight('greaterThan') },
    { label: 'Less than...', icon: 'down', action: () => cfHighlight('lessThan') },
    { label: 'Between...', icon: 'sort', action: () => cfHighlight('between') },
    { label: 'Equal to...', icon: 'hash', action: () => cfHighlight('equal') },
    { label: 'Text that contains...', icon: 'text', action: () => cfHighlight('textContains') },
    { label: 'Duplicate values', icon: 'dedupe', action: () => cfAdd((b) => b.setDuplicateValues().setBackground('#FFC7CE').setFontColor('#9C0006')) },
    { header: 'Top and bottom' },
    { label: 'Top 10 items', icon: 'up', action: () => cfAdd((b) => b.setRank({ isBottom: false, isPercent: false, value: 10 }).setBackground('#C6EFCE').setFontColor('#006100')) },
    { label: 'Bottom 10 items', icon: 'down', action: () => cfAdd((b) => b.setRank({ isBottom: true, isPercent: false, value: 10 }).setBackground('#FFC7CE').setFontColor('#9C0006')) },
    { label: 'Above average', icon: 'up', action: () => cfAdd((b) => b.setAverage('greaterThan').setBackground('#C6EFCE').setFontColor('#006100')) },
    { label: 'Below average', icon: 'down', action: () => cfAdd((b) => b.setAverage('lessThan').setBackground('#FFC7CE').setFontColor('#9C0006')) },
    { header: 'Bars, scales and icons' },
    { label: 'Blue data bars', icon: 'chart-bar', action: () => bar('#4F7DF3') },
    { label: 'Green data bars', icon: 'chart-bar', action: () => bar('#22B07D') },
    { label: 'Orange data bars', icon: 'chart-bar', action: () => bar('#F59E0B') },
    { label: 'Green - yellow - red scale', icon: 'palette', action: () => scale(['#63BE7B', '#FFEB84', '#F8696B']) },
    { label: 'Red - yellow - green scale', icon: 'palette', action: () => scale(['#F8696B', '#FFEB84', '#63BE7B']) },
    { label: 'Blue - white - red scale', icon: 'palette', action: () => scale(['#5A8AC6', '#FCFCFF', '#F8696B']) },
    { label: 'Arrows', icon: 'up', action: () => icons('3Arrows') },
    { label: 'Traffic lights', icon: 'ellipse', action: () => icons('3TrafficLights1') },
    { label: 'Flags', icon: 'bookmark', action: () => icons('3Flags') },
    { label: 'Stars', icon: 'sparkles', action: () => icons('3Stars') },
    'sep',
    { label: 'Clear rules from selected cells', icon: 'clear', action: () => cfClear('selection') },
    { label: 'Clear rules from this sheet', icon: 'trash', action: () => cfClear('sheet') },
    { label: 'New rule...', icon: 'plus', action: () => run('sheet.operation.open.conditional.formatting.panel', { value: 1 }) },
    { label: 'Manage rules...', icon: 'list', action: () => run('sheet.operation.open.conditional.formatting.panel', { value: 2 }) },
  ];
  const sortMenu = () => [
    { label: 'Sort A to Z', icon: 'sort-az', action: () => sort(true) },
    { label: 'Sort Z to A', icon: 'sort-za', action: () => sort(false) },
    { label: 'Custom sort...', icon: 'sort', action: () => run('sheet.command.sort-range-custom') },
    'sep',
    { label: 'Filter', icon: 'filter', shortcut: 'Ctrl+Shift+L', action: () => run('sheet.command.smart-toggle-filter') },
    { label: 'Clear filter', icon: 'filter-x', action: () => run('sheet.command.clear-filter-criteria') },
    { label: 'Reapply', icon: 'refresh', action: () => run('sheet.command.re-calc-filter') },
  ];
  const fnMenu = (cat) => () => (FUNCTIONS.find((g) => g.cat === cat)?.items || []).map(([fname, , desc]) => ({ label: fname, hint: desc, icon: 'function', action: () => insertFunction(fname) }));
  const sumMenu = [
    { label: 'Sum', icon: 'sigma', shortcut: 'Alt+=', action: () => autoSum('SUM') },
    { label: 'Average', icon: 'function', action: () => autoSum('AVERAGE') },
    { label: 'Count numbers', icon: 'hash', action: () => autoSum('COUNT') },
    { label: 'Max', icon: 'up', action: () => autoSum('MAX') },
    { label: 'Min', icon: 'down', action: () => autoSum('MIN') },
    'sep',
    { label: 'More functions...', icon: 'function', action: () => insertFunction() },
  ];
  const chartMenu = () => CHART_TYPES.map((t) => ({ label: t.label, icon: CHART_ICON[t.id], action: () => insertChart(t.id) }));

  const ribbon = createRibbon(el.querySelector('[data-ribbon]'), {
    app: 'sheets',
    fileMenu,
    inline: [{ id: 'stats', kind: 'label', label: '' }],
    tabs: [
      {
        id: 'home', label: 'Home', groups: [
          { label: 'Undo', items: [
            { id: 'undo', icon: 'undo', title: 'Undo (Ctrl+Z)', action: () => cmd('univer.command.undo') },
            { id: 'redo', icon: 'redo', title: 'Redo (Ctrl+Y)', action: () => cmd('univer.command.redo') },
          ] },
          { label: 'Clipboard', items: [
            { id: 'paste', kind: 'big', icon: 'paste', label: 'Paste', title: 'Paste (Ctrl+V)', action: () => cmd('univer.command.paste') },
            { id: 'cut', icon: 'cut', title: 'Cut (Ctrl+X)', action: () => cmd('univer.command.cut') },
            { id: 'copy', icon: 'copy', title: 'Copy (Ctrl+C)', action: () => cmd('univer.command.copy') },
            { id: 'paste-values', icon: 'paste-plain', title: 'Paste values only', action: () => cmd('sheet.command.paste-value') },
            { id: 'painter', icon: 'painter', title: 'Format painter: copy the formatting of the selection, then click where it goes', action: () => cmd('sheet.command.set-once-format-painter') },
          ] },
          { label: 'Font', rows: [
            [
              { id: 'fontname', kind: 'combo', width: 150, title: 'Font', items: () => univerFonts().map((f) => ({ label: f.label, value: f.value, style: `font-family:"${f.value}";font-size:14px` })), action: (v) => v && cmd('sheet.command.set-range-font-family', { value: String(v).trim() }) },
              { id: 'fontsize', kind: 'combo', width: 54, title: 'Font size', items: () => [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72].map((n) => ({ label: String(n), value: n })), action: (v) => { const n = parseFloat(String(v).replace(',', '.')); if (n > 0 && n < 410) cmd('sheet.command.set-range-fontsize', { value: n }); else { toast('Enter a size between 1 and 409', 'error'); syncHome(); } } },
              { id: 'grow', icon: 'grow', title: 'Increase font size', action: () => cmd('sheet.command.set-range-font-increase') },
              { id: 'shrink', icon: 'shrink', title: 'Decrease font size', action: () => cmd('sheet.command.set-range-font-decrease') },
            ],
            [
              { id: 'b', icon: 'bold', title: 'Bold (Ctrl+B)', action: () => cmd('sheet.command.set-range-bold') },
              { id: 'i', icon: 'italic', title: 'Italic (Ctrl+I)', action: () => cmd('sheet.command.set-range-italic') },
              { id: 'u', icon: 'underline', title: 'Underline (Ctrl+U)', action: () => cmd('sheet.command.set-range-underline') },
              { id: 's', icon: 'strike', title: 'Strikethrough', action: () => cmd('sheet.command.set-range-stroke') },
              { id: 'borders', kind: 'menu', icon: 'borders-all', title: 'Borders', items: () => borderMenu() },
              { id: 'fill', kind: 'swatch', icon: 'fill', palette: 'fill', value: '#FFFF00', title: 'Fill colour', action: (c) => (c === 'none' ? cmd('sheet.command.reset-background-color') : cmd('sheet.command.set-background-color', { value: c })) },
              { id: 'fc', kind: 'swatch', icon: 'font-color', palette: 'text', value: '#C00000', title: 'Font colour', action: (c) => (c === 'none' ? cmd('sheet.command.reset-range-text-color') : cmd('sheet.command.set-range-text-color', { value: c })) },
            ],
          ] },
          { label: 'Alignment', rows: [
            [
              { id: 'vt', icon: 'v-top', title: 'Top align', action: () => cmd('sheet.command.set-vertical-text-align', { value: 1 }) },
              { id: 'vm', icon: 'v-mid', title: 'Middle align', action: () => cmd('sheet.command.set-vertical-text-align', { value: 2 }) },
              { id: 'vb', icon: 'v-bottom', title: 'Bottom align', action: () => cmd('sheet.command.set-vertical-text-align', { value: 3 }) },
              { id: 'rotate', kind: 'menu', icon: 'rotate', title: 'Orientation', items: () => [
                { label: 'Angle counterclockwise', icon: 'rotate', action: () => cmd('sheet.command.set-text-rotation', { value: 45 }) },
                { label: 'Angle clockwise', icon: 'rotate', action: () => cmd('sheet.command.set-text-rotation', { value: -45 }) },
                { label: 'Vertical text', icon: 'rotate', action: () => cmd('sheet.command.set-text-rotation', { value: 'v' }) },
                { label: 'Rotate text up', icon: 'rotate', action: () => cmd('sheet.command.set-text-rotation', { value: 90 }) },
                { label: 'Rotate text down', icon: 'rotate', action: () => cmd('sheet.command.set-text-rotation', { value: -90 }) },
                'sep',
                { label: 'No rotation', icon: 'text', action: () => cmd('sheet.command.set-text-rotation', { value: 0 }) },
              ] },
              { id: 'wrap', icon: 'wrap', title: 'Wrap text: show long text on several lines', action: () => cmd('sheet.command.set-text-wrap', { value: homeState.wrap ? 1 : 3 }) },
            ],
            [
              { id: 'al', icon: 'align-left', title: 'Align left', action: () => cmd('sheet.command.set-horizontal-text-align', { value: 1 }) },
              { id: 'ac', icon: 'align-center', title: 'Centre', action: () => cmd('sheet.command.set-horizontal-text-align', { value: 2 }) },
              { id: 'ar', icon: 'align-right', title: 'Align right', action: () => cmd('sheet.command.set-horizontal-text-align', { value: 3 }) },
              { id: 'merge', kind: 'menu', icon: 'merge', title: 'Merge cells', items: () => [
                { label: 'Merge & centre', icon: 'merge', action: async () => { await cmd('sheet.command.add-worksheet-merge-all'); cmd('sheet.command.set-horizontal-text-align', { value: 2 }); } },
                { label: 'Merge across', hint: 'Each row on its own', icon: 'merge', action: () => cmd('sheet.command.add-worksheet-merge-horizontal') },
                { label: 'Merge cells', icon: 'merge', action: () => cmd('sheet.command.add-worksheet-merge-all') },
                { label: 'Unmerge cells', icon: 'split', action: () => cmd('sheet.command.remove-worksheet-merge') },
              ] },
              { id: 'clip', icon: 'text', title: 'Cut off long text at the cell edge', action: () => cmd('sheet.command.set-text-wrap', { value: 2 }) },
            ],
          ] },
          { label: 'Number', rows: [
            [
              { id: 'numfmt', kind: 'combo', width: 132, title: 'Number format', items: () => [...NUMBER_FORMATS.map((f) => ({ label: f.label, value: f.pattern })), { label: 'More number formats...', value: '__more' }], action: (v) => setNumFmt(v) },
            ],
            [
              { id: 'currency', icon: 'currency', title: 'Currency', action: () => cmd('sheet.command.numfmt.set.currency') },
              { id: 'percent', icon: 'percent', title: 'Percent style', action: () => cmd('sheet.command.numfmt.set.percent') },
              { id: 'comma', icon: 'num-format', title: 'Thousands separator (1,234.00)', action: () => setNumFmt('#,##0.00') },
              { id: 'dec-more', icon: 'dec-more', title: 'Increase decimals', action: () => cmd('sheet.command.numfmt.add.decimal.command') },
              { id: 'dec-less', icon: 'dec-less', title: 'Decrease decimals', action: () => cmd('sheet.command.numfmt.subtract.decimal.command') },
            ],
          ] },
          { label: 'Cells', items: [
            { id: 'ins', kind: 'menu', icon: 'insert-row', label: 'Insert', items: [
              { label: 'Insert rows above', icon: 'insert-row', action: () => rowsCols((s, r, row, nr) => s.insertRowsBefore(row, nr)) },
              { label: 'Insert rows below', icon: 'insert-row', action: () => rowsCols((s, r, row, nr) => s.insertRowsAfter(row + nr - 1, nr)) },
              { label: 'Insert columns left', icon: 'insert-col', action: () => rowsCols((s, r, row, nr, col, nc) => s.insertColumnsBefore(col, nc)) },
              { label: 'Insert columns right', icon: 'insert-col', action: () => rowsCols((s, r, row, nr, col, nc) => s.insertColumnsAfter(col + nc - 1, nc)) },
            ] },
            { id: 'del', kind: 'menu', icon: 'trash', label: 'Delete', items: [
              { label: 'Delete rows', icon: 'rows', action: () => rowsCols((s, r, row, nr) => s.deleteRows(row, nr)) },
              { label: 'Delete columns', icon: 'columns', action: () => rowsCols((s, r, row, nr, col, nc) => s.deleteColumns(col, nc)) },
            ] },
            { id: 'fmt', kind: 'menu', icon: 'table-style', label: 'Format', items: [
              { header: 'Cell size' },
              { label: 'Row height...', icon: 'rows', action: () => setSize('row') },
              { label: 'Column width...', icon: 'columns', action: () => setSize('col') },
              { label: 'AutoFit column width', icon: 'fit', action: () => run('sheet.command.set-col-auto-width') },
              { header: 'Visibility' },
              { label: 'Hide rows', icon: 'eye-off', action: () => rowsCols((s, r, row, nr) => s.hideRows(row, nr)) },
              { label: 'Hide columns', icon: 'eye-off', action: () => rowsCols((s, r, row, nr, col, nc) => s.hideColumns(col, nc)) },
              { label: 'Unhide rows', icon: 'view', action: () => rowsCols((s, r, row, nr) => s.showRows(Math.max(0, row - 1), nr + 2)) },
              { label: 'Unhide columns', icon: 'view', action: () => rowsCols((s, r, row, nr, col, nc) => s.showColumns(Math.max(0, col - 1), nc + 2)) },
            ] },
          ] },
          { label: 'Editing', items: [
            { id: 'autosum', kind: 'menu', icon: 'sigma', label: 'AutoSum', items: sumMenu },
            { id: 'sortfilter', kind: 'menu', icon: 'list-filter', label: 'Sort & Filter', items: sortMenu },
            { id: 'clear', kind: 'menu', icon: 'clear', label: 'Clear', items: [
              { label: 'Clear all', icon: 'clear', action: () => busy('Clear', () => selection().clear()) },
              { label: 'Clear contents', icon: 'text', shortcut: 'Del', action: () => busy('Clear', () => selection().clearContent()) },
              { label: 'Clear formats', icon: 'clear-format', action: () => busy('Clear', () => selection().clearFormat()) },
            ] },
            { id: 'find', kind: 'menu', icon: 'search', label: 'Find', items: [
              { label: 'Find...', icon: 'search', shortcut: 'Ctrl+F', action: () => run('ui.operation.open-find-dialog') },
              { label: 'Replace...', icon: 'replace', shortcut: 'Ctrl+H', action: () => run('ui.operation.open-replace-dialog') },
            ] },
          ] },
          { label: 'Styles', items: [
            { id: 'cf', kind: 'menu', icon: 'palette', label: 'Conditional', title: 'Conditional formatting', items: cfMenu },
            { id: 'tablefmt', kind: 'menu', icon: 'table-style', label: 'As table', title: 'Format as table', items: TABLE_STYLES.map((t) => ({ label: t.label, hint: 'Banded rows, header and filter', icon: 'table-style', action: () => formatAsTable(t) })) },
            { id: 'cellstyles', kind: 'menu', icon: 'brush', label: 'Cell styles', items: CELL_STYLES.map((s) => ({ label: s.label, icon: 'brush', action: () => applyCellStyle(s) })) },
          ] },
        ],
      },
      {
        id: 'insert', label: 'Insert', groups: [
          { label: 'Tables', items: [
            { id: 'pivot', kind: 'big', icon: 'pivot', label: 'PivotTable', title: 'Summarise data with a PivotTable', action: insertPivot },
            { id: 'table', kind: 'menu', icon: 'table-style', label: 'Table', items: TABLE_STYLES.map((t) => ({ label: `${t.label} table`, icon: 'table-style', action: () => formatAsTable(t) })) },
          ] },
          { label: 'Filters', items: [
            { id: 'slicer', kind: 'big', icon: 'slicer', label: 'Slicer', title: 'Buttons that filter a table by the values in a column', action: insertSlicer },
          ] },
          { label: 'Sparklines', items: [
            { id: 'sparklines', kind: 'menu', icon: 'sparkline', label: 'Sparklines', title: 'Tiny charts inside cells', items: sparkMenu },
          ] },
          { label: 'Charts', items: [
            { id: 'ch-column', kind: 'big', icon: 'chart-column', label: 'Column', action: () => insertChart('column') },
            { id: 'ch-line', kind: 'big', icon: 'chart-line', label: 'Line', action: () => insertChart('line') },
            { id: 'ch-pie', kind: 'big', icon: 'chart-pie', label: 'Pie', action: () => insertChart('pie') },
            { id: 'ch-bar', kind: 'big', icon: 'chart-bar', label: 'Bar', action: () => insertChart('bar') },
            { id: 'ch-more', kind: 'menu', icon: 'chart-area', label: 'More', items: chartMenu },
          ] },
          { label: 'Illustrations', items: [
            { id: 'picture', kind: 'big', icon: 'image', label: 'Picture', action: () => run('sheet.command.insert-float-image') },
          ] },
          { label: 'Links and notes', items: [
            { id: 'link', kind: 'big', icon: 'link', label: 'Link', title: 'Insert link (Ctrl+K)', action: () => run('sheet.operation.insert-hyper-link-toolbar') },
            { id: 'note', kind: 'big', icon: 'note-add', label: 'Note', action: () => run('sheet.operation.add-note-popup') },
          ] },
          { label: 'Functions', items: [
            { id: 'fx', kind: 'big', icon: 'function', label: 'Function', action: () => insertFunction() },
          ] },
        ],
      },
      {
        id: 'layout', label: 'Page Layout', groups: [
          { label: 'Page setup', items: [
            { id: 'pl-margins', kind: 'menu', icon: 'margins', label: 'Margins', items: marginsMenu },
            { id: 'pl-orient', kind: 'menu', icon: 'orientation', label: 'Orientation', items: orientationMenu },
            { id: 'pl-size', kind: 'menu', icon: 'paper', label: 'Size', title: 'Paper size', items: paperMenu },
            { id: 'pl-area', kind: 'menu', icon: 'select-all', label: 'Print area', items: printAreaMenu },
            { id: 'pl-breaks', kind: 'menu', icon: 'page-break', label: 'Breaks', title: 'Page breaks', items: breaksMenu },
            { id: 'pl-titles', kind: 'big', icon: 'rows', label: 'Print titles', title: 'Rows printed at the top of every page', action: printTitles },
          ] },
          { label: 'Scale to fit', items: [
            { id: 'pl-fit', kind: 'menu', icon: 'fit', label: 'Scale', title: 'How the sheet is fitted to the page', items: fitMenu },
          ] },
          { label: 'Sheet options', items: [
            { id: 'pl-grid', kind: 'big', icon: 'gridlines', label: 'Print gridlines', title: 'Print the cell grid', action: () => setSetup({ gridlines: !(setupOf().gridlines ?? printGrid) }) },
            { id: 'pl-center', kind: 'big', icon: 'align-center', label: 'Centre on page', title: 'Centre the printout across the page', action: () => setSetup({ center: !setupOf().center }) },
            { id: 'pl-hf', kind: 'big', icon: 'header', label: 'Header & footer', action: headerFooter },
          ] },
          { label: 'Print', items: [
            { id: 'pl-preview', kind: 'big', icon: 'print', label: 'Print preview', title: 'See the pages as they will print (Ctrl+P)', action: printPreview },
          ] },
        ],
      },
      {
        id: 'formulas', label: 'Formulas', groups: [
          { label: 'Named ranges', items: [
            { id: 'name-define', kind: 'big', icon: 'bookmark', label: 'Define name', title: 'Name the selected cells, to use the name in formulas', action: () => defineName() },
            { id: 'name-use', kind: 'menu', icon: 'function', label: 'Use in formula', items: () => useNameMenu() },
            { id: 'name-manager', kind: 'big', icon: 'list', label: 'Name manager', title: 'All named ranges: change, rename or delete them', action: () => run('sidebar.operation.defined-name', { value: 'open' }) },
          ] },
          { label: 'Function library', items: [
            { id: 'fx2', kind: 'big', icon: 'function', label: 'Insert function', action: () => insertFunction() },
            { id: 'autosum2', kind: 'menu', icon: 'sigma', label: 'AutoSum', items: sumMenu },
            ...['Math', 'Logical', 'Lookup', 'Text', 'Date', 'Statistical', 'Financial'].map((cat) => ({ id: `fn-${cat}`, kind: 'menu', icon: { Math: 'calculator', Logical: 'magic', Lookup: 'search', Text: 'text', Date: 'calendar', Statistical: 'chart-column', Financial: 'percent' }[cat], label: cat, items: fnMenu(cat) })),
          ] },
        ],
      },
      {
        id: 'data', label: 'Data', groups: [
          { label: 'Sort and filter', items: [
            { id: 'sort-az', kind: 'big', icon: 'sort-az', label: 'A to Z', action: () => sort(true) },
            { id: 'sort-za', kind: 'big', icon: 'sort-za', label: 'Z to A', action: () => sort(false) },
            { id: 'sort-custom', kind: 'big', icon: 'sort', label: 'Sort', title: 'Custom sort', action: () => run('sheet.command.sort-range-custom') },
            { id: 'filter', kind: 'big', icon: 'filter', label: 'Filter', title: 'Filter (Ctrl+Shift+L)', action: () => run('sheet.command.smart-toggle-filter') },
            { id: 'filter-clear', icon: 'filter-x', label: 'Clear', action: () => run('sheet.command.clear-filter-criteria') },
            { id: 'filter-re', icon: 'refresh', label: 'Reapply', action: () => run('sheet.command.re-calc-filter') },
            { id: 'slicer2', icon: 'slicer', label: 'Slicer', title: 'Buttons that filter a table by the values in a column', action: insertSlicer },
          ] },
          { label: 'Data tools', items: [
            { id: 'dv', kind: 'menu', icon: 'validation', label: 'Validation', title: 'Data validation: dropdown lists and input rules', items: [
              { label: 'Add a rule to the selection...', icon: 'plus', action: () => run('data-validation.command.addRuleAndOpen') },
              { label: 'Dropdown list from the selection', icon: 'list', hint: 'Turns the selected cells into a dropdown of their values', action: dropdownFromSelection },
              { label: 'Manage rules...', icon: 'list', action: () => run('data-validation.operation.open-validation-panel', {}) },
              'sep',
              { label: 'Clear validation from the selection', icon: 'clear', action: () => busy('Clear validation', () => { for (const d of selection().getDataValidations?.() || []) d.delete(); markDirty(); }) },
            ] },
            { id: 'dedupe', kind: 'big', icon: 'dedupe', label: 'Remove duplicates', action: removeDuplicates },
            { id: 'split', kind: 'menu', icon: 'columns', label: 'Text to columns', items: [
              { label: 'Split at commas', icon: 'columns', action: () => split('Comma') },
              { label: 'Split at semicolons', icon: 'columns', action: () => split('Semicolon') },
              { label: 'Split at spaces', icon: 'columns', action: () => split('Space') },
              { label: 'Split at tabs', icon: 'columns', action: () => split('Tab') },
            ] },
          ] },
          { label: 'Outline', items: [
            { id: 'group', kind: 'menu', icon: 'group', label: 'Group', title: 'Group rows or columns so they can be hidden together', items: groupsMenu },
            { id: 'collapse', icon: 'fold', label: 'Collapse', title: 'Hide the group at the cursor', action: () => collapseAt(true) },
            { id: 'expand', icon: 'unfold', label: 'Expand', title: 'Show the group at the cursor', action: () => collapseAt(false) },
            { id: 'collapse-all', icon: 'outline', label: 'All', title: 'Collapse or expand every group on the sheet', action: () => allGroups(!groupsOf(ws().getSheetId()).some((g) => isHidden(ws(), g.axis, g.start))) },
          ] },
          { label: 'Analyse', items: [
            { id: 'pivot2', kind: 'big', icon: 'pivot', label: 'PivotTable', action: insertPivot },
            { id: 'pivot-refresh', kind: 'big', icon: 'refresh', label: 'Refresh', title: 'Rebuild PivotTables from their data (new rows and categories included)', action: refreshPivots },
          ] },
        ],
      },
      {
        id: 'review', label: 'Review', groups: [
          { label: 'Notes', items: [
            { id: 'note2', kind: 'big', icon: 'note-add', label: 'New note', action: () => run('sheet.operation.add-note-popup') },
          ] },
          { label: 'Protect', items: [
            { id: 'protect', kind: 'big', icon: 'lock', label: 'Protect range', action: () => run('sheet.command.add-range-protection-from-toolbar') },
          ] },
          { label: 'Find', items: [
            { id: 'find2', kind: 'big', icon: 'search', label: 'Find', action: () => run('ui.operation.open-find-dialog') },
            { id: 'replace2', kind: 'big', icon: 'replace', label: 'Replace', action: () => run('ui.operation.open-replace-dialog') },
          ] },
        ],
      },
      {
        id: 'view', label: 'View', groups: [
          { label: 'Window', items: [
            { id: 'freeze', kind: 'menu', icon: 'freeze', label: 'Freeze panes', items: [
              { label: 'Freeze panes at the selection', icon: 'freeze', action: () => freeze('selection') },
              { label: 'Freeze top row', icon: 'header', action: () => freeze('row') },
              { label: 'Freeze first column', icon: 'sidebar', action: () => freeze('col') },
              'sep',
              { label: 'Unfreeze panes', icon: 'unlock', action: () => freeze('none') },
            ] },
            { id: 'splitpanes', kind: 'big', icon: 'split-panes', label: 'Split', title: 'Split the window at the selection into panes that scroll on their own', action: toggleSplit },
          ] },
          { label: 'Show', items: [
            { id: 'gridlines', kind: 'big', icon: 'gridlines', label: 'Gridlines', action: () => busy('Gridlines', () => { const s = ws(); s.setHiddenGridlines(!s.hasHiddenGridLines()); ribbon.setOn('gridlines', !s.hasHiddenGridLines()); }) },
          ] },
          { label: 'Zoom', items: [
            { id: 'zin', kind: 'big', icon: 'zoom-in', label: 'Zoom in', action: () => zoomBy(0.1) },
            { id: 'zout', kind: 'big', icon: 'zoom-out', label: 'Zoom out', action: () => zoomBy(-0.1) },
            { id: 'z100', kind: 'big', icon: 'fit', label: '100%', action: () => zoomBy(0) },
          ] },
        ],
      },
    ],
  });
  syncLayout();
  ribbon.setOn('gridlines', true);

  // ---------------------------------------------------------------- dirty tracking
  // Univer emits mutations while it loads and recalculates, so only count
  // changes once the user has actually interacted with the sheet
  let touched = false;
  const touch = () => { touched = true; };
  host.addEventListener('pointerdown', touch, true);
  host.addEventListener('keydown', touch, true);
  host.addEventListener('paste', touch, true);
  el.querySelector('[data-ribbon]').addEventListener('click', touch, true);
  const sub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (event) => {
    if (touched && event.type === MUTATION) setDirty(true);
    if (event.type === MUTATION || /selection/i.test(event.id)) syncHomeSoon();
  });
  setTimeout(syncHome, 300);

  // Keyboard shortcuts Univer does not have
  const keys = (e) => {
    if (e.altKey && (e.key === '=' || e.code === 'Equal')) { e.preventDefault(); e.stopPropagation(); autoSum('SUM'); }
    else if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 'l') { e.preventDefault(); e.stopPropagation(); run('sheet.command.smart-toggle-filter'); }
  };
  host.addEventListener('keydown', keys, true);

  let lastCount = -1;
  const updateStats = () => {
    const n = workbook.getSheets().length;
    if (n === lastCount) return;
    lastCount = n;
    setStatus(`${n} sheet${n === 1 ? '' : 's'}`);
    ribbon.setLabel('stats', `${n} sheet${n === 1 ? '' : 's'}`);
  };
  updateStats();
  const statSub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => { if (/sheet\.(command|mutation)\.(insert|remove|copy)-sheet/.test(e.id)) setTimeout(updateStats, 0); });
  if (ext === 'csv') toast('CSV files keep values only. Use Save As .xlsx to keep formatting.');

  /**
   * Formulas whose answer fills several cells (FILTER, SORT, UNIQUE, SEQUENCE...): where each
   * one spills to, for saving them as Excel's dynamic arrays. The spilled values are also put in
   * the snapshot, so they are cached in the file and printed.
   */
  function spillRanges(snap) {
    const out = {};
    try {
      const fdm = univer.__getInjector().get(FormulaDataModel);
      const unit = workbook.getId();
      const ranges = fdm.getArrayFormulaRange()?.[unit] || {};
      const values = fdm.getArrayFormulaCellData()?.[unit] || {};
      for (const [sid, rows] of Object.entries(ranges)) {
        const sh = snap.sheets[sid];
        if (!sh) continue;
        for (const [r, cols] of Object.entries(rows || {})) {
          for (const [c, rg] of Object.entries(cols || {})) {
            if (!rg || rg.endRow == null) continue;
            (out[sid] ||= {})[`${r}:${c}`] = { startRow: rg.startRow, endRow: rg.endRow, startColumn: rg.startColumn, endColumn: rg.endColumn };
          }
        }
        for (const [r, cols] of Object.entries(values[sid] || {})) {
          for (const [c, cell] of Object.entries(cols || {})) {
            if (!cell || cell.v === undefined || cell.v === null) continue;
            const row = (sh.cellData[r] ||= {});
            if (row[c]?.f || (row[c]?.v !== undefined && row[c]?.v !== null && row[c]?.v !== '')) continue;
            row[c] = { ...(row[c] || {}), v: cell.v, ...(cell.t ? { t: cell.t } : {}) };
          }
        }
      }
    } catch (e) { console.warn('spill ranges', e); }
    return out;
  }

  /** The workbook as data plus its charts (with their current positions) */
  const current = () => {
    const snap = workbook.save();
    const models = charts.models(snap);
    charts.stripFromSnapshot(snap);
    sparks.stripFromSnapshot(snap);
    const slicerModels = slicers.all(snap);
    const spills = spillRanges(snap);
    slicers.stripFromSnapshot(snap);
    // a split is a way of looking at the sheet, not part of it
    for (const sid of splits.keys()) if (snap.sheets[sid]) delete snap.sheets[sid].freeze;
    const live = new Set(snap.sheetOrder);
    const extrasNow = {
      pivots: Object.fromEntries([...pivots].filter(([id]) => live.has(id)).map(([id, p]) => {
        const src = p.sourceId && snap.sheets[p.sourceId];
        return [id, { ...p, source: { ...p.source, sheet: src?.name || p.source.sheet } }];
      })),
      sparklines: sparks.all(),
      slicers: slicerModels,
      spills,
      print: Object.fromEntries(Object.entries(printSetup).filter(([id]) => live.has(id))),
      keep: Object.fromEntries(Object.entries(keepParts).filter(([id]) => live.has(id))),
      outline: Object.fromEntries(Object.entries(outline).filter(([id]) => live.has(id))),
    };
    return { snap, models, extras: extrasNow };
  };

  return {
    async save(format, { background = false } = {}) {
      // a cell still being typed into is entered first, as in Excel (not for the background
      // recovery copy, which must not interrupt typing)
      if (!background) { try { await workbook.endEditingAsync?.(true); } catch { /* not editing */ } }
      await settleFormulas();
      const { snap, models, extras: ex } = current();
      if (format === 'csv') {
        const activeId = workbook.getActiveSheet()?.getSheetId();
        if (snap.sheetOrder.length > 1) toast('CSV saves the current sheet only');
        // a CSV that was opened without a byte-order mark is saved without one; new ones get it, for Excel
        return new TextEncoder().encode(`${snap.custom?.csvBom === false ? '' : '\uFEFF'}${snapshotToCsv(snap, activeId)}`);
      }
      return snapshotToXlsx(snap, { charts: models, extras: ex });
    },
    async exportPdf() {
      try { await workbook.endEditingAsync?.(true); } catch { /* not editing */ }
      await settleFormulas();
      const { snap, models, extras: ex } = current();
      const images = {};
      for (const [sid, list] of Object.entries(models)) {
        const sheetName = snap.sheets[sid]?.name;
        for (const m of list) {
          const png = await renderChartPng(m, (ref) => refValues(snap, ref, sheetName));
          if (png) (images[sid] ||= []).push({ ...m.anchor, png });
        }
      }
      const sparklines = {};
      for (const [sid, list] of Object.entries(ex.sparklines || {})) sparklines[sid] = list.map((sp) => ({ ...sp, values: sparkValues(snap, sp, snap.sheets[sid]?.name) }));
      return snapshotToPdf(snap, { gridlines: printGrid, title: name.replace(/\.[^.]+$/, ''), images, sparklines, print: ex.print || {} });
    },
    setTheme(isDark) { try { univerAPI.toggleDarkMode(isDark); } catch { /* older Univer */ } charts.setTheme(isDark); },
    markSaved() { setDirty(false); },
    dispose() {
      try { sub?.dispose?.(); statSub?.dispose?.(); shiftSub?.dispose?.(); } catch { /* noop */ }
      charts.dispose();
      sparks.dispose();
      slicers.dispose();
      splitWheelOff?.();
      ribbon.dispose();
      univer.dispose();
    },
    univerAPI,
    workbook,
    charts,
    insertChart,
    insertPivot,
    refreshPivots,
    insertSparklines,
    sparks,
    slicers,
    splits,
    pivots,
    outline,
  };
}
