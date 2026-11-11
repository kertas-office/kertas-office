import JSZip from 'jszip';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { xlsxToSnapshot, snapshotToXlsx, csvToSnapshot, snapshotToCsv } from '../src/lib/xlsx-bridge.js';

const first = (snap, name) => Object.values(snap.sheets).find((s) => s.name === name);

test('xlsx -> snapshot keeps values, formulas, styles, merges, layout', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/sample.xlsx', import.meta.url)), 'sample');
  assert.equal(snap.sheetOrder.length, 2);
  const s = first(snap, 'Sales');
  assert.equal(s.cellData[0][0].v, 'Item');
  assert.equal(s.cellData[1][1].v, 10);
  assert.equal(s.cellData[1][3].f, '=B2*C2');
  assert.equal(s.cellData[4][3].f, '=SUM(D2:D4)');
  const hdr = snap.styles[s.cellData[0][0].s];
  assert.equal(hdr.bl, 1); assert.equal(hdr.cl.rgb, '#FFFFFF'); assert.equal(hdr.bg.rgb, '#1F4E79'); assert.equal(hdr.ff, 'Arial');
  assert.equal(snap.styles[s.cellData[1][2].s].n.pattern, '#,##0.00');
  assert.equal(snap.styles[s.cellData[1][4].s].n.pattern !== undefined, true, 'date gets a format');
  assert.deepEqual(s.mergeData[0], { startRow: 6, startColumn: 0, endRow: 6, endColumn: 2 });
  assert.equal(s.columnData[0].w, 145);
  assert.equal(s.rowData[0].h, 40);
  assert.equal(s.freeze.ySplit, 1);
  assert.equal(snap.styles[s.cellData[4][3].s].bd.b.s, 7);
});

test('snapshot -> xlsx round-trips through Excel format', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/sample.xlsx', import.meta.url)), 'sample');
  const bytes = await snapshotToXlsx(snap);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet('Sales');
  assert.equal(ws.getCell('A1').value, 'Item');
  assert.equal(ws.getCell('A1').font.bold, true);
  assert.equal(ws.getCell('A1').fill.fgColor.argb, 'FF1F4E79');
  assert.equal(ws.getCell('D2').formula, 'B2*C2');
  assert.equal(ws.getCell('C2').numFmt, '#,##0.00');
  assert.equal(ws.getCell('B1').alignment.horizontal, 'center');
  assert.equal(ws.getCell('D5').border.bottom.style, 'double');
  assert.ok(ws.getCell('A7').isMerged);
  assert.equal(ws.getColumn(1).width, 20);
  assert.equal(ws.getRow(1).height, 30);
  assert.equal(ws.views[0].ySplit, 1);
  assert.equal(wb.getWorksheet('Notes').getCell('B2').value, true);
});

test('csv parse and export', () => {
  const snap = csvToSnapshot('name,amount,note\r\n"Ali, M",12.5,"says ""hi"""\r\nSiti,7,\r\n', 'people.csv');
  const s = snap.sheets[snap.sheetOrder[0]];
  assert.equal(s.cellData[1][0].v, 'Ali, M');
  assert.equal(s.cellData[1][1].v, 12.5);
  assert.equal(s.cellData[1][2].v, 'says "hi"');
  assert.equal(snapshotToCsv(snap), 'name,amount,note\r\n"Ali, M",12.5,"says ""hi"""\r\nSiti,7,\r\n');
});

// ---------------------------------------------------------------- Round 4 features

const res = (snap, name) => JSON.parse(snap.resources.find((r) => r.name === name)?.data || '{}');
const only = (obj) => Object.values(obj)[0];

test('xlsx from another writer: filters, conditional formats, validation, links, notes, pictures', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/features.xlsx', import.meta.url)), 'features');
  const data = first(snap, 'Data');
  const lists = first(snap, 'Lists');

  const f = res(snap, 'SHEET_FILTER_PLUGIN')[data.id];
  assert.deepEqual([f.ref.startRow, f.ref.startColumn, f.ref.endRow, f.ref.endColumn], [0, 0, 4, 3]);
  assert.deepEqual(f.filterColumns[0], { colId: 3, filters: { filters: ['North'] } });
  assert.deepEqual(f.filterColumns[1].customFilters.customFilters, [{ operator: 'greaterThan', val: '2' }]);
  assert.deepEqual(f.cachedFilteredOut, [2], 'the filtered-out row is handed to the filter');
  assert.ok(!data.rowData[2]?.hd, 'and is not hidden twice');

  const cf = res(snap, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[data.id].map((x) => x.rule);
  const byType = (t, sub) => cf.find((r) => r.type === t && (!sub || r.subType === sub));
  assert.deepEqual(byType('highlightCell', 'number'), { type: 'highlightCell', subType: 'number', operator: 'greaterThan', value: 4, style: { bl: 1, cl: { rgb: '#9C0006' }, bg: { rgb: '#FFC7CE' } } });
  assert.equal(byType('highlightCell', 'formula').value, '=$B2>4');
  assert.deepEqual(byType('highlightCell', 'text'), { type: 'highlightCell', subType: 'text', operator: 'containsText', value: 'Pe', style: { it: 1 } });
  assert.equal(byType('highlightCell', 'rank').value, 2);
  assert.equal(byType('dataBar').config.positiveColor, '#638EC6');
  assert.deepEqual(byType('colorScale').config.map((c) => c.value.type), ['min', 'percentile', 'max']);
  const icons = byType('iconSet').config;
  assert.deepEqual(icons.map((c) => [c.operator, c.value.value]), [['greaterThanOrEqual', 67], ['greaterThanOrEqual', 33], ['lessThanOrEqual', 33]]);

  const dv = res(snap, 'SHEET_DATA_VALIDATION_PLUGIN')[data.id];
  const list = dv.find((d) => d.type === 'list' && d.formula1.startsWith('['));
  assert.deepEqual(JSON.parse(list.formula1), ['North', 'South', 'East', 'West']);
  assert.deepEqual(list.ranges[0], { startRow: 1, startColumn: 3, endRow: 19, endColumn: 3, rangeType: 0 }, 'cells are merged back into one range');
  assert.equal(dv.find((d) => d.type === 'whole').error, '1 to 100 only');
  assert.match(dv.find((d) => d.type === 'list' && d.formula1.startsWith('=')).formula1, /^=\[.+\]Lists!\$A\$1:\$A\$3$/);
  assert.equal(dv.find((d) => d.type === 'date').formula1, '2026-01-01');

  assert.equal(res(snap, 'SHEET_NOTE_PLUGIN')[data.id][1][0].note, 'Best seller');
  assert.equal(data.cellData[1][4].p.body.customRanges[0].properties.url, 'https://example.com/kertas');
  assert.equal(data.cellData[2][4].p.body.customRanges[0].properties.url, `#gid=${lists.id}&range=A1`);

  const pic = only(res(snap, 'SHEET_DRAWING_PLUGIN')[data.id].data);
  assert.deepEqual([pic.column, pic.row, pic.width, pic.height], [7, 1, 64, 64]);
  assert.match(pic.source, /^data:image\/png;base64,iVBOR/);
});

test('Univer features survive a save to .xlsx and back', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/sample.xlsx', import.meta.url)), 'sample');
  const s = first(snap, 'Sales');
  const R = (r1, c1, r2 = r1, c2 = c1) => ({ startRow: r1, startColumn: c1, endRow: r2, endColumn: c2, rangeType: 0 });
  const rule = (r, ranges) => ({ cfId: Math.random().toString(36), ranges, rule: r, stopIfTrue: false });
  const cf = [
    rule({ type: 'highlightCell', subType: 'number', operator: 'between', value: [1, 5], style: { cl: { rgb: 'rgb(0,0,255)' } } }, [R(1, 1, 3, 1)]),
    rule({ type: 'highlightCell', subType: 'text', operator: 'beginsWith', value: 'Pe', style: { bl: 1 } }, [R(1, 0, 3, 0)]),
    rule({ type: 'highlightCell', subType: 'duplicateValues', operator: 'containsText', value: '', style: { bg: { rgb: '#FFEB9C' } } }, [R(1, 0, 3, 0)]),
    rule({ type: 'highlightCell', subType: 'average', operator: 'lessThan', value: '', style: { ul: { s: 1 } } }, [R(1, 2, 3, 2)]),
    rule({ type: 'highlightCell', subType: 'rank', operator: '', value: 1, isBottom: true, isPercent: false, style: { it: 1 } }, [R(1, 3, 3, 3)]),
    rule({ type: 'highlightCell', subType: 'timePeriod', operator: 'last7Days', value: '', style: { bg: { rgb: '#C6EFCE' } } }, [R(1, 4, 3, 4)]),
    rule({ type: 'dataBar', isShowValue: true, config: { min: { type: 'min' }, max: { type: 'max' }, positiveColor: '#5B9BD5', nativeColor: '#FF0000', isGradient: false } }, [R(1, 2, 3, 2)]),
    rule({ type: 'iconSet', isShowValue: true, config: [
      { iconType: '3Arrows', iconId: '0', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 67 } },
      { iconType: '3Arrows', iconId: '1', operator: 'greaterThanOrEqual', value: { type: 'percent', value: 33 } },
      { iconType: '3Arrows', iconId: '2', operator: 'lessThanOrEqual', value: { type: 'percent', value: 33 } }] }, [R(1, 1, 3, 1)]),
  ];
  const dv = [
    { uid: 'a', ranges: [R(1, 5, 9, 5)], type: 'list', formula1: '["Pen","Book","Bag"]', showDropDown: true },
    { uid: 'b', ranges: [R(1, 6, 3, 6)], type: 'checkbox' },
    { uid: 'c', ranges: [R(1, 7, 3, 7)], type: 'decimal', operator: 'greaterThan', formula1: '0', showErrorMessage: true, error: 'Must be positive' },
    { uid: 'd', ranges: [R(1, 8, 3, 8)], type: 'textLength', operator: 'lessThanOrEqual', formula1: '10' },
    { uid: 'e', ranges: [R(1, 9, 3, 9)], type: 'custom', formula1: '=J2<>""' },
    { uid: 'f', ranges: [R(1, 10, 3, 10)], type: 'list', formula1: `=[${snap.id}]Sales!A2:A4`, showDropDown: true },
  ];
  const filter = { ref: R(0, 0, 3, 4), filterColumns: [{ colId: 0, filters: { filters: ['Pen', 'Bag'] } }, { colId: 1, customFilters: { and: 1, customFilters: [{ operator: 'greaterThan', val: '2' }, { operator: 'lessThan', val: '50' }] } }], cachedFilteredOut: [2] };
  const png = (await readFile(new URL('./fixtures/logo.png', import.meta.url))).toString('base64');
  const drawing = { data: { pic1: { drawingId: 'pic1', drawingType: 0, imageSourceType: 'BASE64', source: `data:image/png;base64,${png}`, column: 6, columnOffset: 5, row: 1, rowOffset: 0, width: 80, height: 60, sheetTransform: { from: { column: 6, columnOffset: 5, row: 1, rowOffset: 0 }, to: { column: 6, columnOffset: 85, row: 4, rowOffset: 0 } }, transform: { left: 636, top: 60, width: 80, height: 60 } } }, order: ['pic1'] };
  const notes = { 3: { 1: { id: 'n', row: 3, col: 1, note: 'Check this', width: 200, height: 80, show: false } } };
  const notesSheet = first(snap, 'Notes');
  s.cellData[9] = { 0: { v: 'Docs', t: 1, p: { id: 'd', documentStyle: {}, body: { dataStream: 'Docs\r\n', paragraphs: [{ startIndex: 4 }], customRanges: [{ rangeType: 0, rangeId: 'l1', properties: { url: 'https://kertas.app/docs' }, startIndex: 0, endIndex: 3 }] } } },
    1: { v: 'Notes tab', t: 1, p: { id: 'e', documentStyle: {}, body: { dataStream: 'Notes tab\r\n', paragraphs: [{ startIndex: 9 }], customRanges: [{ rangeType: 0, rangeId: 'l2', properties: { url: `#gid=${notesSheet.id}&range=B2` }, startIndex: 0, endIndex: 8 }] } } } };
  snap.resources = [
    { name: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN', data: JSON.stringify({ [s.id]: cf }) },
    { name: 'SHEET_DATA_VALIDATION_PLUGIN', data: JSON.stringify({ [s.id]: dv }) },
    { name: 'SHEET_FILTER_PLUGIN', data: JSON.stringify({ [s.id]: filter }) },
    { name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify({ [s.id]: drawing }) },
    { name: 'SHEET_NOTE_PLUGIN', data: JSON.stringify({ [s.id]: notes }) },
  ];

  const bytes = await snapshotToXlsx(snap);

  // What Excel will see
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const ws = wb.getWorksheet('Sales');
  assert.equal(ws.getCell('B4').note?.texts?.map((t) => t.text).join('') ?? ws.getCell('B4').note, 'Check this');
  assert.deepEqual(ws.getCell('A10').value, { text: 'Docs', hyperlink: 'https://kertas.app/docs' });
  // a link to a place in the workbook is stored as a location only, as Excel does
  const sheetXml = await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheetXml, /<hyperlink ref="B10" location="Notes!B2"\/>/);
  assert.equal(ws.getRow(3).hidden, true, 'filtered rows are hidden in the file');
  assert.equal(ws.getImages().length, 1);
  assert.equal(ws.getCell('F2').dataValidation.formulae[0], '"Pen,Book,Bag"');
  assert.equal(ws.getCell('K3').dataValidation.formulae[0], 'Sales!A2:A4');

  // And back again
  const back = await xlsxToSnapshot(bytes, 'back');
  const bs = first(back, 'Sales');
  assert.ok(bs.cellData[9][1].p?.body?.customRanges?.[0]?.properties?.url?.startsWith('#gid='), 'the workbook link is read back');
  const cf2 = res(back, 'SHEET_CONDITIONAL_FORMATTING_PLUGIN')[bs.id].map((x) => x.rule);
  assert.deepEqual(cf2.map((r) => r.subType || r.type), ['number', 'text', 'duplicateValues', 'average', 'rank', 'timePeriod', 'dataBar', 'iconSet']);
  assert.deepEqual(cf2[0].value, [1, 5]);
  assert.deepEqual(cf2[0].style, { cl: { rgb: '#0000FF' } });
  assert.deepEqual([cf2[1].operator, cf2[1].value], ['beginsWith', 'Pe']);
  assert.equal(cf2[3].operator, 'lessThan');
  assert.deepEqual([cf2[4].value, cf2[4].isBottom], [1, true]);
  assert.equal(cf2[5].operator, 'last7Days');
  assert.equal(cf2[6].config.positiveColor, '#5B9BD5');
  assert.deepEqual(cf2[7].config.map((c) => c.value.value), [67, 33, 33]);

  const dv2 = res(back, 'SHEET_DATA_VALIDATION_PLUGIN')[bs.id];
  const t = (type) => dv2.find((d) => d.type === type);
  assert.deepEqual(JSON.parse(dv2.find((d) => d.type === 'list' && d.formula1.startsWith('[')).formula1), ['Pen', 'Book', 'Bag']);
  assert.deepEqual([t('checkbox').formula1, t('checkbox').formula2], ['1', '0']);
  assert.deepEqual([t('decimal').operator, t('decimal').formula1, t('decimal').error], ['greaterThan', '0', 'Must be positive']);
  assert.equal(t('textLength').operator, 'lessThanOrEqual');
  assert.equal(t('custom').formula1, '=J2<>""');
  assert.equal(dv2.find((d) => d.type === 'list' && d.formula1.startsWith('=')).formula1, `=[${back.id}]Sales!A2:A4`);

  const f2 = res(back, 'SHEET_FILTER_PLUGIN')[bs.id];
  assert.deepEqual(f2.filterColumns, filter.filterColumns);
  assert.deepEqual(f2.cachedFilteredOut, [2]);

  const pic = only(res(back, 'SHEET_DRAWING_PLUGIN')[bs.id].data);
  assert.deepEqual([pic.column, pic.columnOffset, pic.row, pic.width, pic.height], [6, 5, 1, 80, 60]);
  assert.equal(res(back, 'SHEET_NOTE_PLUGIN')[bs.id][3][1].note, 'Check this');
  assert.equal(bs.cellData[9][1].p.body.customRanges[0].properties.url, `#gid=${first(back, 'Notes').id}&range=B2`);
});
