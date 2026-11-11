import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { xlsxToSnapshot, snapshotToXlsx, blankSnapshot, prefixFormula, unprefixFormula } from '../src/lib/xlsx-bridge.js';
import { seriesFromGrid, chartXml, parseChartXml, refValues, parseRef, resolveChart, CHART_TYPES } from '../src/lib/xlsx-charts.js';
import { buildPivot, distinct } from '../src/lib/pivot.js';
import { snapshotToPdf } from '../src/lib/sheet-pdf.js';

const GRID = [['Item', 'Qty', 'Price'], ['Pen', 3, 1.5], ['Book', 8, 12], ['Bag', 1, 40], ['Pen', 5, 1.5]];

function sheetWith(grid, name = 'Data') {
  const snap = blankSnapshot('t');
  const sh = snap.sheets[snap.sheetOrder[0]];
  sh.name = name;
  grid.forEach((r, i) => r.forEach((v, j) => { (sh.cellData[i] ||= {})[j] = typeof v === 'number' ? { v, t: 2 } : { v, t: 1 }; }));
  return { snap, sh };
}

test('chart series are found the way Excel finds them', () => {
  const s = seriesFromGrid('Data', { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 }, GRID);
  assert.equal(s.headers, true);
  assert.equal(s.labels, true);
  assert.equal(s.byRows, false);
  assert.equal(s.categories, 'Data!$A$2:$A$5');
  assert.deepEqual(s.series, [{ name: 'Data!$B$1', values: 'Data!$B$2:$B$5' }, { name: 'Data!$C$1', values: 'Data!$C$2:$C$5' }]);
  const wide = seriesFromGrid('My Data', { startRow: 0, startColumn: 0, endRow: 1, endColumn: 4 }, [['', 'Q1', 'Q2', 'Q3', 'Q4'], ['Sales', 10, 20, 30, 40]]);
  assert.equal(wide.byRows, true, 'more columns than rows gives one series per row');
  assert.deepEqual(wide.series, [{ name: "'My Data'!$A$2", values: "'My Data'!$B$2:$E$2" }]);
  assert.equal(wide.categories, "'My Data'!$B$1:$E$1");
  assert.deepEqual(parseRef("'It''s'!$B$2:$C$3"), { sheet: "It's", range: { startRow: 1, startColumn: 1, endRow: 2, endColumn: 2, rangeType: 0 } });
});

test('charts are saved as real Excel charts and read back', async () => {
  const { snap, sh } = sheetWith(GRID);
  const s = seriesFromGrid('Data', { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 }, GRID);
  snap.charts = {
    [sh.id]: CHART_TYPES.map((t, i) => ({
      id: `c${i}`, type: t.id, title: `${t.label} chart`, legend: ['bottom', 'right', 'top', 'none'][i % 4], stacked: t.id === 'column',
      categories: s.categories, series: s.series, anchor: { col: 5, colOff: 10, row: i * 15, rowOff: 4, width: 400 + i, height: 250 },
    })),
  };
  const bytes = await snapshotToXlsx(snap);
  const zip = await JSZip.loadAsync(bytes);
  const types = await zip.file('[Content_Types].xml').async('string');
  assert.equal((types.match(/drawingml\.chart\+xml/g) || []).length, 7);
  const chart1 = await zip.file('xl/charts/chart1.xml').async('string');
  assert.match(chart1, /<c:barDir val="col"\/><c:grouping val="stacked"\/>/);
  assert.match(chart1, /<c:pt idx="1"><c:v>8<\/c:v><\/c:pt>/, 'values are cached for viewers that do not recalculate');
  assert.match(await zip.file('xl/worksheets/sheet1.xml').async('string'), /<drawing r:id="rId\d+"\/>/);

  const back = await xlsxToSnapshot(bytes, 'back');
  const bs = Object.values(back.sheets)[0];
  const charts = back.charts[bs.id];
  assert.deepEqual(charts.map((c) => c.type), CHART_TYPES.map((t) => t.id));
  assert.deepEqual(charts.map((c) => c.title), CHART_TYPES.map((t) => `${t.label} chart`));
  assert.deepEqual(charts.map((c) => c.legend), ['bottom', 'right', 'top', 'none', 'bottom', 'right', 'top']);
  assert.equal(charts[0].stacked, true);
  assert.deepEqual(charts[2].series.map((x) => [x.name, x.values]), [['Data!$B$1', 'Data!$B$2:$B$5'], ['Data!$C$1', 'Data!$C$2:$C$5']]);
  assert.equal(charts[2].categories, 'Data!$A$2:$A$5');
  charts.forEach((c, i) => assert.deepEqual(c.anchor, { col: 5, colOff: 10, row: i * 15, rowOff: 4, width: 400 + i, height: 250 }));
});

test('chart values resolve from cells, including other sheets', () => {
  const { snap } = sheetWith(GRID, 'Sales 2026');
  const model = { type: 'column', categories: "'Sales 2026'!$A$2:$A$5", series: [{ name: "'Sales 2026'!$B$1", values: "'Sales 2026'!$B$2:$B$5" }] };
  const r = resolveChart(model, (ref) => refValues(snap, ref));
  assert.deepEqual(r.labels, ['Pen', 'Book', 'Bag', 'Pen']);
  assert.deepEqual(r.datasets[0], { name: 'Qty', values: [3, 8, 1, 5], color: '#4F7DF3' });
  const xml = chartXml({ ...model, type: 'pie', title: 'Share' }, (ref) => refValues(snap, ref));
  const parsed = parseChartXml(xml);
  assert.equal(parsed.type, 'pie');
  assert.equal(parsed.title, 'Share');
});

test('pivot summaries use live formulas', () => {
  assert.deepEqual(distinct(['b', 2, 'A', '', 10, 'a', null, 2]), [2, 10, 'A', 'b']);
  const grid = [['Region', 'Item', 'Qty'], ['North', 'Pen', 3], ['South', 'Book', 8], ['North', 'Bag', 1]];
  const p = buildPivot(grid, { sheet: 'My Data', startRow: 4, startColumn: 1 }, { rows: 0, columns: 1, values: 2, agg: 'sum' });
  assert.equal(p.cells[0][0].v, 'Sum of Qty');
  assert.deepEqual(p.cells[1].map((c) => c.v), ['Region', 'Bag', 'Book', 'Pen', 'Grand Total']);
  assert.equal(p.cells[2][0].v, 'North');
  assert.equal(p.cells[2][1].f, "=SUMIFS('My Data'!$D$6:$D$8,'My Data'!$B$6:$B$8,$A3,'My Data'!$C$6:$C$8,B$2)");
  assert.equal(p.cells.at(-1).at(-1).f, "=SUMIFS('My Data'!$D$6:$D$8,'My Data'!$B$6:$B$8,\"<>\")");
  const avg = buildPivot(grid, { sheet: 'S', startRow: 0, startColumn: 0 }, { rows: 1, columns: null, values: 2, agg: 'average' });
  assert.equal(avg.cells[2][1].f, '=IFERROR(AVERAGEIFS(S!$C$2:$C$4,S!$B$2:$B$4,$A3),"")');
});

test('newer Excel functions get their file prefix', () => {
  assert.equal(prefixFormula('MAXIFS(A:A,B:B,"x")+XLOOKUP(1,A1:A3,B1:B3)+SUM(UNIQUE(A1:A3))+"MINIFS(1)"'),
    '_xlfn.MAXIFS(A:A,B:B,"x")+_xlfn.XLOOKUP(1,A1:A3,B1:B3)+SUM(_xlfn._xlws.UNIQUE(A1:A3))+"MINIFS(1)"');
  assert.equal(unprefixFormula('_xlfn.MINIFS(A1:A3,B1:B3,1)+_xlfn._xlws.SORT(C1:C3)'), 'MINIFS(A1:A3,B1:B3,1)+SORT(C1:C3)');
});

test('sheet PDF prints pictures and charts', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/features.xlsx', import.meta.url)), 'f');
  const png = new Uint8Array(await readFile(new URL('./fixtures/logo.png', import.meta.url)));
  const data = Object.values(snap.sheets).find((s) => s.name === 'Data');
  const pdf = await snapshotToPdf(snap, { images: { [data.id]: [{ col: 2, colOff: 0, row: 8, rowOff: 0, width: 300, height: 200, png }] } });
  const text = new TextDecoder('latin1').decode(pdf);
  // Each transparent PNG is an image plus its alpha mask
  const images = (text.match(/\/Subtype \/Image/g) || []).length - (text.match(/\/SMask \d+ 0 R/g) || []).length;
  assert.equal(images, 2, 'the sheet picture and the chart');
});

test('blank rows inside a chart range are skipped', () => {
  const { snap } = sheetWith([...GRID, ['', '', ''], ['', '', '']]);
  const r = resolveChart({ type: 'pie', categories: 'Data!$A$2:$A$7', series: [{ values: 'Data!$B$2:$B$7' }] }, (ref) => refValues(snap, ref));
  assert.deepEqual(r.labels, ['Pen', 'Book', 'Bag', 'Pen']);
  assert.deepEqual(r.datasets[0].values, [3, 8, 1, 5]);
});
