import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { execFileSync } from 'node:child_process';
import { writeFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { xlsxToSnapshot, snapshotToXlsx, blankSnapshot } from '../src/lib/xlsx-bridge.js';
import { distinctItems, filterKey } from '../src/modules/sheet-slicers.js';

function book() {
  const snap = blankSnapshot('t');
  const sid = snap.sheetOrder[0];
  const sh = snap.sheets[sid];
  sh.name = 'Sales data';
  [['Region', 'Revenue'], ['North', 100], ['South', 80], ['North', 120]].forEach((r, i) => r.forEach((v, j) => { (sh.cellData[i] ||= {})[j] = typeof v === 'number' ? { v, t: 2 } : { v, t: 1 }; }));
  return { snap, sid };
}

test('slicer values match what the sheet filter compares', () => {
  assert.equal(filterKey(120), '120');
  assert.equal(filterKey(true), 'TRUE');
  assert.equal(filterKey('a\nb'), 'ab');
  const items = distinctItems(['North', 12, '', 3, 'North', 'east'], ['North', '12.00', '', '3.00', 'North', 'east']);
  assert.deepEqual(items.map((x) => x.label), ['3.00', '12.00', 'east', 'North', '(Blank)']);
  assert.equal(items[0].key, '3');
});

test('slicers are saved with the workbook and read back by sheet', async () => {
  const { snap, sid } = book();
  const model = { id: 'x', title: 'Region', range: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 1 }, col: 0, selected: ['North', '\u0000blank'], anchor: { col: 3, colOff: 0, row: 0, rowOff: 0, width: 184, height: 232 } };
  const bytes = await snapshotToXlsx(snap, { extras: { slicers: { [sid]: [model] } } });
  const zip = await JSZip.loadAsync(bytes);
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  assert.match(rels, /relationships\/customXml" Target="\.\.\/customXml\/item\d+\.xml"/);
  const back = await xlsxToSnapshot(bytes, 'b.xlsx');
  const list = Object.values(back.extras.slicers)[0];
  assert.equal(list.length, 1);
  assert.equal(list[0].title, 'Region');
  assert.deepEqual(list[0].selected, ['North', '\u0000blank']);
  assert.deepEqual(list[0].anchor, model.anchor);
  // LibreOffice still opens the workbook
  const dir = mkdtempSync(join(tmpdir(), 'slicer-'));
  writeFileSync(join(dir, 'a.xlsx'), bytes);
  try {
    execFileSync('soffice', ['--headless', '--convert-to', 'csv', '--outdir', dir, join(dir, 'a.xlsx')], { timeout: 90000, stdio: 'ignore' });
    const csv = execFileSync('cat', [join(dir, 'a.csv')]).toString();
    assert.match(csv, /North,120/);
  } catch (e) { if (e.code !== 'ENOENT') throw e; }
});
