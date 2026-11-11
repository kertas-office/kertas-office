// Fixes from the November audit
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { xlsxToSnapshot, snapshotToXlsx } from '../src/lib/xlsx-bridge.js';
import { fieldNames } from '../src/lib/pivot.js';
import JSZip from 'jszip';

test('a time rule (Data validation, Time) survives opening and saving', async () => {
  const snap = await xlsxToSnapshot(new Uint8Array(fs.readFileSync('src/templates/timesheet.xlsx')), 'Timesheet');
  const out = await snapshotToXlsx(snap);
  const zip = await JSZip.loadAsync(out);
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
  const rules = [...sheet.matchAll(/<dataValidation\b[^>]*sqref="([^"]+)"/g)].map((m) => m[1]);
  assert.ok(rules.some((r) => r.includes('D10')), `time rule kept (rules on ${rules.join(', ')})`);
});

test('blank table headers are named the same way everywhere ("Column F")', () => {
  assert.deepEqual(fieldNames(['Name', '', 'Total'], 4), ['Name', 'Column F', 'Total']);
});

// ---------------------------------------------------------------- data-integrity round
import { csvToSnapshot, csvCell, snapshotToCsv } from '../src/lib/xlsx-bridge.js';
import { anchorPointComments } from '../src/lib/docx-prepare.js';
import ExcelJS from 'exceljs';

test('CSV values that are not plain numbers stay exactly as written', () => {
  assert.deepEqual(csvCell('007'), { v: '007', t: 1 }, 'leading zero');
  assert.deepEqual(csvCell('1234567890123456789'), { v: '1234567890123456789', t: 1 }, 'too long for a number');
  assert.deepEqual(csvCell('=1+1'), { v: '=1+1', t: 1 }, 'formulas from a CSV are not run');
  assert.equal(csvCell('12.5').v, 12.5);
  assert.equal(csvCell('1.50').custom.csv, '1.50', 'original spelling remembered');
});

test('a CSV saved without edits comes back byte for byte', () => {
  for (const text of ['a;b\r\n007;1.50\r\n"x;y";-0\r\n', 'name,n\n"Ali, M",1E5\n"line1\nline2",  42 \n', 'id\tcode\n1\t00123\n']) {
    assert.equal(snapshotToCsv(csvToSnapshot(text, 't.csv')), text);
  }
});

test('the workbook default font, author and turned or indented text survive a save', async () => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Siti'; wb.title = 'Budget';
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'turned'; ws.getCell('A1').alignment = { textRotation: 45, indent: 2 };
  ws.getCell('A2').value = 'x2'; ws.getCell('A2').font = { vertAlign: 'superscript' };
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  // a default font of Arial 10, as LibreOffice and Google Sheets write
  zip.file('xl/styles.xml', (await zip.file('xl/styles.xml').async('string')).replace(/(<fonts\b[^>]*>\s*<font>)[\s\S]*?(<\/font>)/, '$1<sz val="10"/><name val="Arial"/>$2'));
  const snap = await xlsxToSnapshot(await zip.generateAsync({ type: 'uint8array' }), 'x');
  const back = new ExcelJS.Workbook();
  const out = await snapshotToXlsx(snap);
  await back.xlsx.load(out);
  const styles = await (await JSZip.loadAsync(out)).file('xl/styles.xml').async('string');
  assert.match(/<fonts\b[^>]*>\s*<font>([\s\S]*?)<\/font>/.exec(styles)[1], /<name val="Arial"\/>/);
  assert.equal(back.creator, 'Siti');
  assert.equal(back.title, 'Budget');
  const c = back.getWorksheet('S').getCell('A1');
  assert.equal(c.alignment.textRotation, 45);
  assert.equal(c.alignment.indent, 2);
  assert.equal(back.getWorksheet('S').getCell('A2').font.vertAlign, 'superscript');
});

test('a comment attached to a single spot gets a range so it is kept', () => {
  const xml = '<w:body><w:p><w:r><w:t>Words</w:t></w:r><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body>';
  assert.equal(anchorPointComments(xml), '<w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Words</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p></w:body>');
  const ranged = '<w:p><w:commentRangeStart w:id="1"/><w:r><w:t>A</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>';
  assert.equal(anchorPointComments(ranged), ranged, 'comments with a range are left alone');
});
