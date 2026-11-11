// What Kertas keeps when it saves files made in other apps.
// The fixtures were made in LibreOffice (and by hand for macros and SmartArt); the
// "-saved-by-editor" files are what the Word editor wrote for them before the repair pass.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { detectFeatures, atRisk } from '../src/lib/fidelity.js';
import { repairDocx } from '../src/lib/docx-repair.js';
import { xlsxToSnapshot, snapshotToXlsx } from '../src/lib/xlsx-bridge.js';

const fx = (n) => readFile(new URL(`./fixtures/fidelity/${n}`, import.meta.url));
const ids = (list) => list.map((f) => f.id).sort();

test('finds the kinds of content in Word and Excel files', async () => {
  assert.deepEqual(ids(await detectFeatures(await fx('smartart.docx'), 'writer')), ['link', 'smartart']);
  assert.deepEqual(ids(await detectFeatures(await fx('ole.docx'), 'writer')), ['ole']);
  assert.ok(ids(await detectFeatures(await fx('macro.docm'), 'writer')).includes('macros'));
  assert.deepEqual(ids(await detectFeatures(await fx('chart-radar.xlsx'), 'sheets')), ['chart', 'chart-other']);
  assert.ok(ids(await detectFeatures(await fx('button.xlsx'), 'sheets')).includes('control'));
});

test('warns only about what is lost', async () => {
  assert.deepEqual(ids(atRisk(await detectFeatures(await fx('smartart.docx'), 'writer'), 'writer')), []);
  assert.deepEqual(ids(atRisk(await detectFeatures(await fx('macro.docm'), 'writer'), 'writer')), ['macros']);
  assert.deepEqual(ids(atRisk(await detectFeatures(await fx('chart-radar.xlsx'), 'sheets'), 'sheets')), ['chart-other']);
  assert.deepEqual(ids(atRisk(await detectFeatures(await fx('names.xlsx'), 'sheets'), 'sheets')), []);
});

async function dangling(bytes) {
  const z = await JSZip.loadAsync(bytes); const out = [];
  for (const p of Object.keys(z.files).filter((n) => /^word\/[^/]+\.xml$/.test(n))) {
    const used = new Set([...(await z.file(p).async('string')).matchAll(/\br:[A-Za-z]+="([^"]+)"/g)].map((m) => m[1]));
    const rels = await z.file(p.replace(/([^/]+)$/, '_rels/$1.rels'))?.async('string') || '';
    for (const id of used) if (!rels.includes(`Id="${id}"`)) out.push(id);
  }
  return out;
}

test('an embedded object keeps its link to the embedded file', async () => {
  const saved = await fx('ole-saved-by-editor.docx');
  assert.deepEqual(await dangling(saved), ['rId2'], 'the editor drops it');
  const fixed = await repairDocx(new Uint8Array(saved), new Uint8Array(await fx('ole.docx')));
  assert.deepEqual(await dangling(fixed), []);
  const z = await JSZip.loadAsync(fixed);
  assert.match(await z.file('word/_rels/document.xml.rels').async('string'), /Id="rId2"[^>]*Target="embeddings\/oleObject1\.xlsx"/);
  assert.ok(z.file('word/embeddings/oleObject1.xlsx'));
});

test('macro code nothing refers to is taken out of a .docx', async () => {
  const fixed = await repairDocx(new Uint8Array(await fx('macro-saved-by-editor.docx')), new Uint8Array(await fx('macro.docm')));
  const z = await JSZip.loadAsync(fixed);
  assert.equal(z.file('word/vbaProject.bin'), null);
  assert.doesNotMatch(await z.file('[Content_Types].xml').async('string'), /vbaProject/);
});

test('named ranges, sheet protection and Excel tables survive a save', async () => {
  for (const [file, id] of [['names.xlsx', 'names'], ['protect.xlsx', 'protect'], ['table.xlsx', 'table']]) {
    const out = await snapshotToXlsx(await xlsxToSnapshot(await fx(file), file));
    assert.ok(ids(await detectFeatures(out, 'sheets')).includes(id), `${file} keeps ${id}`);
  }
  // the named range is usable by Univer's formulas
  const snap = await xlsxToSnapshot(await fx('names.xlsx'), 'n');
  const names = JSON.parse(snap.resources.find((r) => r.name === 'SHEET_DEFINED_NAME_PLUGIN').data);
  const sales = Object.values(names).find((n) => n.name === 'Sales');
  assert.equal(sales.formulaOrRefString, 'Sheet1!$B$2:$D$5');
  // a renamed table header goes into the saved table, as Excel requires
  const t = await xlsxToSnapshot(await fx('table.xlsx'), 't');
  const sid = t.sheetOrder[0];
  t.sheets[sid].cellData[0][1] = { v: 'First quarter', t: 1 };
  const z = await JSZip.loadAsync(await snapshotToXlsx(t));
  const tableXml = await z.file(Object.keys(z.files).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n))).async('string');
  assert.match(tableXml, /<tableColumn\b[^>]*name="First quarter"/);
});

test('a formula that spills is saved as an Excel dynamic array and spills again on opening', async () => {
  const { blankSnapshot } = await import('../src/lib/xlsx-bridge.js');
  const snap = blankSnapshot('d');
  const sid = snap.sheetOrder[0];
  const sh = snap.sheets[sid];
  [['N', 'S'], ['a', 12], ['b', 7], ['c', 15]].forEach((r, i) => r.forEach((v, j) => { (sh.cellData[i] ||= {})[j] = typeof v === 'number' ? { v, t: 2 } : { v, t: 1 }; }));
  sh.cellData[0][3] = { f: '=FILTER(A2:B4,B2:B4>10)', v: 'a', t: 1 };
  sh.cellData[0][4] = { v: 12, t: 2 }; sh.cellData[1] ||= {}; sh.cellData[1][3] = { v: 'c', t: 1 }; sh.cellData[1][4] = { v: 15, t: 2 };
  const bytes = await snapshotToXlsx(snap, { extras: { spills: { [sid]: { '0:3': { startRow: 0, endRow: 1, startColumn: 3, endColumn: 4 } } } } });
  const z = await JSZip.loadAsync(bytes);
  const sheet = await z.file('xl/worksheets/sheet1.xml').async('string');
  assert.match(sheet, /<c r="D1" cm="1"[^>]*><f t="array" ref="D1:E2">_xlfn\._xlws\.FILTER\(/);
  assert.ok(z.file('xl/metadata.xml'));
  assert.match(await z.file('xl/_rels/workbook.xml.rels').async('string'), /sheetMetadata/);
  const back = await xlsxToSnapshot(bytes, 'd');
  const cd = back.sheets[back.sheetOrder[0]].cellData;
  assert.equal(cd[0][3].f, '=FILTER(A2:B4,B2:B4>10)');
  assert.equal(cd[0][3].v, undefined, 'worked out again when the sheet opens');
  assert.equal(cd[1]?.[3], undefined, 'the stored copy of the answer does not block the spill');
});
