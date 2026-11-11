import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { xlsxToSnapshot, snapshotToXlsx, blankSnapshot } from '../src/lib/xlsx-bridge.js';
import { buildPivot } from '../src/lib/pivot.js';
import { pivotXml, readPivots } from '../src/lib/pivot-xlsx.js';
import { sparkGeometry, sparklinesExtXml, parseSparklines, makeSparklines, addSparklinesToSheetXml } from '../src/lib/sparklines.js';
import { cfEvaluator, iconLook } from '../src/lib/sheet-cf.js';
import { snapshotToPdf } from '../src/lib/sheet-pdf.js';
import { setFontLoader } from '../src/lib/pdf-text.js';
import { nodeFontLoader } from './font-loader.mjs';

setFontLoader(nodeFontLoader);

const SALES = [
  ['Region', 'Quarter', 'Revenue'],
  ['North', 'Q1', 100], ['South', 'Q1', 80], ['North', 'Q2', 120], ['East', 'Q2', 60], ['South', 'Q2', 90],
];
function book(grid = SALES) {
  const snap = blankSnapshot('t');
  const sid = snap.sheetOrder[0];
  const sh = snap.sheets[sid];
  sh.name = 'Data';
  grid.forEach((r, i) => r.forEach((v, j) => { (sh.cellData[i] ||= {})[j] = typeof v === 'number' ? { v, t: 2 } : { v, t: 1 }; }));
  return { snap, sid, sh };
}
function addPivotSheet(snap, grid, spec) {
  const pv = buildPivot(grid, { sheet: 'Data', startRow: 0, startColumn: 0 }, spec);
  const id = 'pivotsheet';
  const cellData = {};
  pv.cells.forEach((row, r) => row.forEach((c, k) => { if (c) (cellData[r] ||= {})[k] = c.f ? { f: c.f } : typeof c.v === 'number' ? { v: c.v, t: 2 } : { v: c.v, t: 1 }; }));
  snap.sheets[id] = { id, name: 'Pivot', rowCount: 100, columnCount: 20, cellData, mergeData: [], columnData: {}, rowData: {} };
  snap.sheetOrder.push(id);
  return { id, pv };
}

test('a PivotTable is saved as a real Excel pivot and read back', async () => {
  const { snap } = book();
  const spec = { rows: 0, columns: 1, values: 2, agg: 'sum' };
  const { id } = addPivotSheet(snap, SALES, spec);
  const extras = { pivots: { [id]: { name: 'PivotTable1', source: { sheet: 'Data', range: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 } }, spec } } };
  const bytes = await snapshotToXlsx(snap, { extras });
  const zip = await JSZip.loadAsync(bytes);
  const table = await zip.file('xl/pivotTables/pivotTable1.xml').async('string');
  assert.match(table, /<location ref="A1:D6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"\/>/, "label column, two quarters and a total; two header rows, three regions and a total");
  assert.match(table, /<rowFields count="1"><field x="0"\/><\/rowFields>/);
  assert.match(table, /<colFields count="1"><field x="1"\/><\/colFields>/);
  assert.match(table, /<dataField name="Sum of Revenue" fld="2"/);
  const def = await zip.file('xl/pivotCache/pivotCacheDefinition1.xml').async('string');
  assert.match(def, /refreshOnLoad="1"/);
  assert.match(def, /<worksheetSource ref="A1:C6" sheet="Data"\/>/);
  assert.match(def, /<cacheField name="Region" numFmtId="0"><sharedItems count="3"><s v="East"\/><s v="North"\/><s v="South"\/><\/sharedItems>/);
  assert.match(await zip.file('xl/workbook.xml').async('string'), /<pivotCaches><pivotCache cacheId="101" r:id="rId\d+"\/><\/pivotCaches>/);
  assert.match(await zip.file('[Content_Types].xml').async('string'), /pivotTable\+xml/);
  const sheetRels = await zip.file('xl/worksheets/_rels/sheet2.xml.rels').async('string');
  assert.match(sheetRels, /pivotTable1\.xml/);

  const back = await xlsxToSnapshot(bytes, 'back');
  const pid = back.sheetOrder.find((s) => back.sheets[s].name === 'Pivot');
  assert.deepEqual(back.extras.pivots[pid].spec, spec);
  assert.equal(back.extras.pivots[pid].source.sheet, 'Data');
  assert.deepEqual(back.extras.pivots[pid].fields, ['Region', 'Quarter', 'Revenue']);
});

test('pivots without a column field start below the caption, and counts use Excel subtotals', () => {
  const { table, ref } = pivotXml({ name: 'P', source: { sheet: 'Data', range: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 2 } }, grid: SALES, spec: { rows: 0, columns: null, values: 2, agg: 'count' }, label: 'Count of Revenue' }, 7);
  assert.equal(ref, 'A2:B6');
  assert.match(table, /<colItems count="1"><i\/><\/colItems>/);
  assert.match(table, /subtotal="count"/);
  assert.match(table, /rowHeaderCaption="Region"/);
});

test('sparklines are drawn to scale and survive a save in Excel form', async () => {
  const g = sparkGeometry([1, 3, 2, null, 5], 'line', 100, 20, 2);
  assert.equal(g.points.length, 5);
  assert.equal(g.points[3], null, 'blanks are gaps');
  assert.equal(g.points[4][1], 2, 'the highest value touches the top');
  const cols = sparkGeometry([4, -2, 6], 'column', 60, 30, 0);
  assert.equal(cols.bars.length, 3);
  assert.ok(cols.bars[1].neg && cols.axis > 0);
  const wl = sparkGeometry([1, -1, 0, 2], 'stacked', 40, 20, 0);
  assert.equal(wl.bars.length, 3, 'zero draws nothing in win/loss');

  const made = makeSparklines({ sheet: 'Data', range: { startRow: 1, startColumn: 1, endRow: 3, endColumn: 4 } }, { startRow: 1, startColumn: 5, endRow: 3, endColumn: 5 }, { type: 'column', color: '#0EA5E9' });
  assert.deepEqual(made.map((s) => [s.r, s.c, s.ref]), [[1, 5, 'Data!B2:E2'], [2, 5, 'Data!B3:E3'], [3, 5, 'Data!B4:E4']]);
  const ext = sparklinesExtXml(made, 'Data');
  assert.match(ext, /<x14:sparklineGroup type="column"/);
  const sheetXml = addSparklinesToSheetXml('<worksheet><sheetData/></worksheet>', ext);
  assert.deepEqual(parseSparklines(sheetXml, 'Data').map((s) => [s.r, s.c, s.ref, s.type, s.color]), made.map((s) => [s.r, s.c, s.ref, 'column', '#0EA5E9']));

  const { snap, sid } = book([['A', 'B', 'C', 'D', 'Trend'], [1, 4, 2, 6], [3, 2, 5, 1]]);
  const sparks = makeSparklines({ sheet: 'Data', range: { startRow: 1, startColumn: 0, endRow: 2, endColumn: 3 } }, { startRow: 1, startColumn: 4, endRow: 2, endColumn: 4 });
  const bytes = await snapshotToXlsx(snap, { extras: { sparklines: { [sid]: sparks } } });
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(xml, /<extLst><ext uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"/);
  assert.match(xml, /<xm:f>Data!A2:D2<\/xm:f><xm:sqref>E2<\/xm:sqref>/);
  const back = await xlsxToSnapshot(bytes, 'back');
  assert.equal(back.extras.sparklines[back.sheetOrder[0]].length, 2);
});

test('grouped rows keep their outline level and collapsed state', async () => {
  const { snap, sid, sh } = book();
  sh.rowData = { 2: { hd: 1 }, 3: { hd: 1 } };
  const extras = { outline: { [sid]: { rows: { 1: 1, 2: 2, 3: 2, 4: 1 }, cols: { 1: 1 } } } };
  const bytes = await snapshotToXlsx(snap, { extras });
  const xml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(xml, /<row r="3"[^>]*hidden="1"[^>]*outlineLevel="2"|<row r="3"[^>]*outlineLevel="2"[^>]*hidden="1"/);
  assert.match(xml, /outlineLevelRow="2"/);
  const back = await xlsxToSnapshot(bytes, 'back');
  assert.deepEqual(back.extras.outline[back.sheetOrder[0]], { rows: { 1: 1, 2: 2, 3: 2, 4: 1 }, cols: { 1: 1 } });
});

test('conditional formatting is worked out for export', () => {
  const { sh } = book([['n'], [5], [15], [25], [15], ['x']]);
  const R = (r0, r1) => [{ startRow: r0, endRow: r1, startColumn: 0, endColumn: 0 }];
  const ev = cfEvaluator(sh, [
    { ranges: R(1, 4), rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 20, style: { bg: { rgb: '#FFC7CE' }, cl: { rgb: '#9C0006' } } } },
    { ranges: R(1, 4), rule: { type: 'highlightCell', subType: 'duplicateValues', style: { bl: 1, bg: { rgb: '#FFEB9C' } } } },
    { ranges: R(1, 4), rule: { type: 'colorScale', config: [{ index: 0, color: '#FFFFFF', value: { type: 'min' } }, { index: 1, color: '#00FF00', value: { type: 'max' } }] } },
    { ranges: R(1, 4), rule: { type: 'dataBar', config: { min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#638EC6', isGradient: false } } },
    { ranges: R(1, 4), rule: { type: 'iconSet', config: [{ iconType: '3Arrows', iconId: '0', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 } }, { iconType: '3Arrows', iconId: '1', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 } }, { iconType: '3Arrows', iconId: '2', operator: 'lessThanOrEqual', value: { type: 'percent', value: 0 } }] } },
  ]);
  assert.equal(ev(3, 0).style.bg.rgb, '#FFC7CE', 'first matching rule wins the fill');
  assert.equal(ev(2, 0).style.bl, 1, 'duplicates are bold');
  assert.equal(ev(2, 0).style.bg.rgb, '#FFEB9C');
  assert.equal(ev(1, 0).style.bg.rgb.toLowerCase(), "#ffffff", "colour scale at the minimum");
  assert.equal(ev(1, 0).bar.ratio, 0);
  assert.equal(ev(3, 0).bar.ratio, 1);
  assert.deepEqual(iconLook(ev(3, 0).icon), { shape: 'up', color: '#16A34A' });
  assert.deepEqual(iconLook(ev(1, 0).icon), { shape: 'down', color: '#DC2626' });
  assert.equal(ev(5, 0), null, 'outside the ranges');
});

test('the PDF shows conditional colours, data bars and sparklines', async () => {
  const { snap, sid } = book([['Score', 'Trend'], [10, ''], [90, ''], [50, '']]);
  snap.resources = [{ name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: JSON.stringify({ [sid]: [
    { cfId: 'a', ranges: [{ startRow: 1, endRow: 3, startColumn: 0, endColumn: 0 }], rule: { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 60, style: { bg: { rgb: '#FF0000' } } } },
  ] }) }];
  const plain = await snapshotToPdf(snap, {});
  const withSparks = await snapshotToPdf(snap, { sparklines: { [sid]: [{ r: 1, c: 1, type: 'line', color: '#4F46E5', values: [1, 3, 2, 5] }] } });
  assert.ok(withSparks.length > plain.length, 'sparkline adds drawing');
  const doc = await pdfjs.getDocument({ data: plain, standardFontDataUrl: new URL('../node_modules/pdfjs-dist/standard_fonts/', import.meta.url).pathname }).promise;
  const ops = await (await doc.getPage(1)).getOperatorList();
  const fills = [];
  for (let i = 0; i < ops.fnArray.length; i++) if (ops.fnArray[i] === pdfjs.OPS.setFillRGBColor) fills.push(ops.argsArray[i].join(','));
  assert.ok(fills.some((f) => /^255,0,0$|^#ff0000$/i.test(f) || f === '1,0,0'), `a red fill is drawn (${[...new Set(fills)].join(' | ')})`);
});

test('PivotTables with several row, column and value fields are built, saved and read back', async () => {
  const { buildPivot } = await import('../src/lib/pivot.js');
  const { pivotXml, readPivots } = await import('../src/lib/pivot-xlsx.js');
  const grid = [
    ['Region', 'Rep', 'Year', 'Sales', 'Units'],
    ['North', 'Ann', 2023, 100, 1], ['North', 'Ann', 2024, 150, 2], ['North', 'Bob', 2023, 80, 1],
    ['South', 'Cy', 2024, 200, 4], ['South', 'Cy', 2023, 50, 1],
  ];
  const spec = { rows: [0, 1], columns: [2], values: [{ fld: 3, agg: 'sum' }, { fld: 4, agg: 'count' }] };
  const pv = buildPivot(grid, { sheet: 'Data', startRow: 0, startColumn: 0 }, spec);
  // title, a Year row, a values row, 3 rep rows, grand total
  assert.equal(pv.cells.length, 1 + 2 + 3 + 1);
  assert.deepEqual(pv.cells[1].slice(0, 4).map((c) => c?.v ?? null), ['Year', null, 2023, 2023]);
  assert.deepEqual(pv.cells[2].slice(0, 4).map((c) => c?.v), ['Region', 'Rep', 'Sum of Sales', 'Count of Units']);
  assert.deepEqual(pv.cells[3].slice(0, 2).map((c) => c.v), ['North', 'Ann']);
  // North / Ann / 2023: sum of sales filtered by region, rep and the year header
  assert.equal(pv.cells[3][2].f, "=SUMIFS(Data!$D$2:$D$6,Data!$A$2:$A$6,$A4,Data!$B$2:$B$6,$B4,Data!$C$2:$C$6,C$2)");
  assert.equal(pv.cells[3][3].f, "=COUNTIFS(Data!$A$2:$A$6,$A4,Data!$B$2:$B$6,$B4,Data!$C$2:$C$6,D$2)");
  assert.equal(pv.cells.at(-1)[0].v, 'Grand Total');

  // saved as a real PivotTable, then read back with every field
  const parts = pivotXml({ name: 'P1', source: { sheet: 'Data', range: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 4 } }, grid, spec, label: 'Values' }, 1);
  assert.match(parts.table, /<rowFields count="2"><field x="0"\/><field x="1"\/><\/rowFields>/);
  assert.match(parts.table, /<colFields count="2"><field x="2"\/><field x="-2"\/><\/colFields>/);
  assert.match(parts.table, /<dataFields count="2">.*subtotal="count"/);
  const JSZip = (await import('jszip')).default;
  const zip = new JSZip();
  zip.file('xl/pivotTables/pivotTable1.xml', parts.table);
  zip.file('xl/pivotTables/_rels/pivotTable1.xml.rels', '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>');
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml', parts.cacheDef);
  const [back] = await readPivots(zip);
  assert.equal(back.supported, true);
  assert.deepEqual(back.spec, { rows: [0, 1], columns: [2], values: [{ fld: 3, agg: 'sum' }, { fld: 4, agg: 'count' }] });
});
