import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { spreadsheetToXlsx, docTextToHtml, openLegacy, isLegacy } from '../src/lib/legacy.js';
import { xlsxToSnapshot } from '../src/lib/xlsx-bridge.js';

const require = createRequire(import.meta.url);
const { docText, convertLegacy, findSoffice } = require('../electron/legacy.cjs');
const fixture = (n) => readFile(new URL(`./fixtures/${n}`, import.meta.url));

test('.xls opens through SheetJS with values, formulas and layout', async () => {
  assert.equal(isLegacy('Old Book.XLS'), true);
  const snap = await xlsxToSnapshot(spreadsheetToXlsx(new Uint8Array(await fixture('legacy.xls'))), 'old');
  const sales = Object.values(snap.sheets).find((s) => s.name === 'Sales');
  assert.equal(sales.cellData[0][0].v, 'Item');
  assert.equal(sales.cellData[1][1].v, 10);
  assert.equal(sales.cellData[1][3].f, '=B2*C2');
  assert.deepEqual(sales.mergeData[0], { startRow: 6, startColumn: 0, endRow: 6, endColumn: 2 });
  assert.equal(snap.styles[sales.cellData[1][2].s].n.pattern, '#,##0.00');
});

test('.doc text comes through with its table, without LibreOffice', async () => {
  const parts = await docText(new Uint8Array(await fixture('legacy.doc')));
  assert.match(parts.body, /Quarterly report/);
  const html = docTextToHtml(parts);
  assert.match(html, /<p>Revenue grew in every region this quarter\.<\/p>/);
  assert.match(html, /<table><tr><td>Region<\/td><td>Q1<\/td><td>Q2<\/td><\/tr><tr><td>North<\/td><td>100<\/td><td>120<\/td><\/tr>/);
  // The renderer side: a desktop that only has the text path
  const res = await openLegacy({ name: 'r.doc', data: new Uint8Array(await fixture('legacy.doc')) }, async (f) => ({ text: await docText(f.data), via: 'text' }));
  assert.ok(res.html.includes('Prepared by Kertas.'));
  await assert.rejects(openLegacy({ name: 'r.odt', data: new Uint8Array(1) }, null), /desktop app/);
});

test('LibreOffice converts .doc to .docx when it is installed', { skip: !findSoffice() && 'LibreOffice not installed' }, async () => {
  const res = await convertLegacy({ name: 'legacy.doc', data: new Uint8Array(await fixture('legacy.doc')) });
  assert.equal(res.via, 'libreoffice');
  assert.equal(res.ext, 'docx');
  assert.equal(String.fromCharCode(res.data[0], res.data[1]), 'PK');
});
