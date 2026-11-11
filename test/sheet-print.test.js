// Page setup: print area, repeated title rows, page breaks, scaling, header and footer
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PDFDocument } from '@cantoo/pdf-lib';
import { snapshotToPdf } from '../src/lib/sheet-pdf.js';
import { blankSnapshot, snapshotToXlsx, xlsxToSnapshot } from '../src/lib/xlsx-bridge.js';
import { parseHeaderFooter, headerFooterText } from '../src/lib/xlsx-print.js';
import { setFontLoader } from '../src/lib/pdf-text.js';
import { nodeFontLoader } from './font-loader.mjs';

setFontLoader(nodeFontLoader);

function book(rows = 120) {
  const snap = blankSnapshot('t');
  const sid = snap.sheetOrder[0];
  const sh = snap.sheets[sid];
  sh.name = 'Sales';
  sh.cellData[0] = { 0: { v: 'Item', t: 1 }, 1: { v: 'Amount', t: 1 } };
  for (let r = 1; r < rows; r++) sh.cellData[r] = { 0: { v: `Item ${r}`, t: 1 }, 1: { v: r * 10, t: 2 } };
  return { snap, sid };
}
const pdfText = (bytes) => {
  const dir = mkdtempSync(join(tmpdir(), 'print-'));
  writeFileSync(join(dir, 'a.pdf'), bytes);
  return execFileSync('pdftotext', ['-layout', join(dir, 'a.pdf'), '-']).toString().split('\f').filter((p) => p.trim());
};

test('title rows repeat, breaks start pages, footer counts pages', async () => {
  const { snap, sid } = book();
  const print = { [sid]: { titleRows: [0, 0], breaks: [10], footer: ['{sheet}', 'Page {page} of {pages}', ''], header: ['', 'Quarterly', ''] } };
  const pages = pdfText(await snapshotToPdf(snap, { print }));
  assert.ok(pages.length >= 3, `${pages.length} pages`);
  assert.match(pages[0], /Item 9\b/);
  assert.doesNotMatch(pages[0], /Item 10\b/, 'the break starts a new page at row 11');
  for (const [i, p] of pages.entries()) {
    assert.match(p, /Item\s+Amount/, `titles on page ${i + 1}`);
    assert.match(p, new RegExp(`Page ${i + 1} of ${pages.length}`));
    assert.match(p, /Quarterly/);
  }
});

test('print area and fit to one page', async () => {
  const { snap, sid } = book();
  const area = { startRow: 0, endRow: 59, startColumn: 0, endColumn: 1 };
  const one = pdfText(await snapshotToPdf(snap, { print: { [sid]: { area, fit: 'page' } } }));
  assert.equal(one.length, 1);
  assert.match(one[0], /Item 59\b/);
  assert.doesNotMatch(one[0], /Item 60\b/, 'rows outside the area are left out');
  const doc = await PDFDocument.load(await snapshotToPdf(snap, { print: { [sid]: { paper: 'letter', orientation: 'landscape' } } }));
  assert.deepEqual(doc.getPage(0).getSize(), { width: 792, height: 612 });
});

test('page setup is saved in the workbook and read back', async () => {
  const { snap, sid } = book(40);
  const setup = { paper: 'letter', orientation: 'landscape', margins: 'narrow', fit: 'width', area: { startRow: 0, endRow: 30, startColumn: 0, endColumn: 1 }, titleRows: [0, 0], breaks: [15], header: ['', 'Report', ''], footer: ['{file}', '', 'Page {page} of {pages}'], center: true };
  const bytes = await snapshotToXlsx(snap, { extras: { print: { [sid]: setup } } });
  const back = await xlsxToSnapshot(bytes, 'b.xlsx');
  const got = back.extras.print[back.sheetOrder[0]];
  for (const k of ['paper', 'orientation', 'margins', 'fit', 'center']) assert.equal(got[k], setup[k], k);
  assert.deepEqual(got.area, setup.area);
  assert.deepEqual(got.titleRows, [0, 0]);
  assert.deepEqual(got.breaks, [15]);
  assert.deepEqual(got.header, setup.header);
  assert.deepEqual(got.footer, setup.footer);
});

test('header and footer codes convert both ways', () => {
  assert.deepEqual(parseHeaderFooter('&L&"Arial,Bold"&10Report&CPage &P of &N&R&D'), ['Report', 'Page {page} of {pages}', '{date}']);
  assert.equal(headerFooterText(['Q&A', '', '{sheet}']), '&LQ&&A&R&A');
});
