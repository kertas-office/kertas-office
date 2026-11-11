import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { PDFDocument } from '@cantoo/pdf-lib';
import { xlsxToSnapshot, blankSnapshot } from '../src/lib/xlsx-bridge.js';
import { snapshotToPdf, formatValue } from '../src/lib/sheet-pdf.js';

test('number formats', () => {
  assert.equal(formatValue(1234.5, 2, '#,##0.00'), '1,234.50');
  assert.equal(formatValue(0.256, 2, '0.0%'), '25.6%');
  assert.equal(formatValue(46328, 2, 'yyyy-mm-dd'), '2026-11-02');
  assert.equal(formatValue(46328.5, 2, 'dd/mm/yyyy hh:mm'), '02/11/2026 12:00');
  assert.equal(formatValue(46328, 2, 'd mmm yyyy'), '2 Nov 2026');
  assert.equal(formatValue(1500, 2, '"RM"#,##0.00'), 'RM1,500.00');
  assert.equal(formatValue(-42, 2, '#,##0'), '-42');
  assert.equal(formatValue(0.1 + 0.2, 2), '0.3');
  assert.equal(formatValue(1, 3), 'TRUE');
});

test('workbook renders to PDF with one page per sheet', async () => {
  const snap = await xlsxToSnapshot(await readFile(new URL('./fixtures/sample.xlsx', import.meta.url)), 'sample');
  const bytes = await snapshotToPdf(snap, { title: 'sample' });
  const doc = await PDFDocument.load(bytes);
  assert.equal(doc.getPageCount(), 2);
});

test('wide sheets go landscape and long sheets paginate', async () => {
  const snap = blankSnapshot('big');
  const sh = snap.sheets[snap.sheetOrder[0]];
  for (let r = 0; r < 120; r++) { sh.cellData[r] = {}; for (let c = 0; c < 9; c++) sh.cellData[r][c] = { v: r * c, t: 2 }; }
  const doc = await PDFDocument.load(await snapshotToPdf(snap));
  const { width, height } = doc.getPage(0).getSize();
  assert.ok(width > height, 'landscape');
  assert.ok(doc.getPageCount() >= 3);
});
