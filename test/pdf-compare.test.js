import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractPages } from '../src/lib/pdf-convert.js';
import { compareText, diffSequences, pairPages, buildReport, joinWords } from '../src/lib/pdf-compare.js';
import { setFontLoader } from '../src/lib/pdf-text.js';
import { nodeFontLoader } from './font-loader.mjs';

setFontLoader(nodeFontLoader);
const fonts = new URL('../node_modules/pdfjs-dist/standard_fonts/', import.meta.url).pathname;
const open = async (bytes) => pdfjs.getDocument({ data: new Uint8Array(bytes), fontExtraProperties: true, standardFontDataUrl: fonts }).promise;
const read = async (f) => new Uint8Array(await readFile(new URL(`./fixtures/${f}`, import.meta.url)));
const text = async (bytes) => extractPages(await open(bytes), { Util: pdfjs.Util });

test('the diff keeps what matches and marks the rest', () => {
  const ops = diffSequences(['a', 'b', 'c', 'd'], ['a', 'x', 'c', 'd', 'e']);
  assert.deepEqual(ops.map((o) => o[0]).join(''), '=-+==+');
  // Long inputs split on unique anchors and stay exact
  const a = Array.from({ length: 3000 }, (_, i) => `w${i}`);
  const b = [...a.slice(0, 1500), 'new', ...a.slice(1500, 2999)];
  const big = diffSequences(a, b);
  assert.equal(big.filter((o) => o[0] === '+').length, 1);
  assert.equal(big.filter((o) => o[0] === '-').length, 1);
});

test('pages pair by shared text, and a new page stands alone', () => {
  const t = new Map([['0:0', 40], ['1:2', 12]]);
  assert.deepEqual(pairPages(2, 3, t), [[0, 0], [null, 1], [1, 2]]);
});

test('two versions of a contract: replaced figure, added and removed text, a new page', async () => {
  const a = await text(await read('compare-original.pdf'));
  const b = await text(await read('compare-revised.pdf'));
  const r = compareText(a, b);
  const said = r.changes.map((c) => [c.kind, joinWords(c.removed), joinWords(c.added)]);
  assert.deepEqual(said, [
    ['replaced', '4,500', '5,200'],
    ['added', '', 'Urgent issues are answered within four hours.'],
    ['added', '', 'Data protection Client data stays in Malaysia and is deleted when the agreement ends.'],
    ['removed', 'Signed for the provider by the head of operations.', ''],
  ]);
  assert.deepEqual(r.pairs, [[0, 0], [null, 1], [1, 2]]);
  // The removed line sits on the second page of the original, the new page is page 2 of the revision
  assert.equal(r.changes[3].removed[0].page, 1);
  assert.equal(r.changes[2].added[0].page, 1);
});

test('the report has a summary, then the pages side by side with the changes boxed', async () => {
  const A = await read('compare-original.pdf');
  const B = await read('compare-revised.pdf');
  const pa = await text(A);
  const pb = await text(B);
  const r = compareText(pa, pb);
  const bytes = await buildReport({ bytes: A, pages: pa, name: 'Agreement v1.pdf' }, { bytes: B, pages: pb, name: 'Agreement v2.pdf' }, r, { title: 'Agreement v1 compared with v2' });
  const doc = await open(bytes);
  assert.equal(doc.numPages, 1 + 3);
  const summary = (await (await doc.getPage(1)).getTextContent()).items.map((i) => i.str).join(' ');
  assert.match(summary, /4 changes: 1 replaced, 2 added, 1 removed/);
  assert.match(summary, /Was\s+4,500/);
  assert.match(summary, /Now\s+5,200/);
  // Each summary entry links to its page
  const links = (await (await doc.getPage(1)).getAnnotations()).filter((x) => x.subtype === 'Link');
  assert.equal(links.length, 4);
  // A side-by-side page is two pages wide and carries both versions' text
  const p2 = await doc.getPage(2);
  const [, , w] = p2.view;
  assert.ok(w > 595 * 2);
  const t2 = (await p2.getTextContent()).items.map((i) => i.str).join(' ');
  assert.match(t2, /Original, page 1/);
  assert.match(t2, /Revised, page 1/);
  assert.match(t2, /RM 4,500/);
  assert.match(t2, /RM 5,200/);
  const t3 = (await (await doc.getPage(3)).getTextContent()).items.map((i) => i.str).join(' ');
  assert.match(t3, /Original: no matching page/);
});

test('identical files have no changes', async () => {
  const a = await text(await read('compare-original.pdf'));
  const r = compareText(a, a);
  assert.equal(r.changes.length, 0);
  assert.deepEqual(r.pairs, [[0, 0], [1, 1]]);
});
