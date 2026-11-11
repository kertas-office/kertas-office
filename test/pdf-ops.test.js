import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { PDFDocument } from '@cantoo/pdf-lib';
import { buildPdf, toWinAnsi } from '../src/lib/pdf-ops.js';

const fx = (n) => readFile(new URL(`./fixtures/${n}`, import.meta.url));
// 1x1 transparent PNG
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64'));

test('reorder, rotate, delete, merge and annotate', async () => {
  const sources = { s0: new Uint8Array(await fx('sample.pdf')), s1: new Uint8Array(await fx('extra.pdf')) };
  const pages = [
    { src: 's0', index: 2, rotate: 90, annots: [{ type: 'text', x: 72, y: 600, size: 14, color: '#e5483d', text: 'Approved \u2014 Hashir\nline two', rot: 90 }] },
    { src: 's1', index: 0, rotate: 0, annots: [{ type: 'highlight', x1: 70, y1: 495, x2: 260, y2: 515, color: '#ffd400' }] },
    { src: 's0', index: 0, rotate: 0, annots: [
      { type: 'rect', x1: 60, y1: 720, x2: 300, y2: 780, color: '#2f6fed', width: 2 },
      { type: 'ink', points: [[100, 100], [150, 140], [200, 110]], color: '#000000', width: 2 },
      { type: 'image', x: 300, y: 100, w: 120, h: 40, rot: 0, png: PNG },
    ] },
  ];
  const out = await buildPdf(sources, pages);
  const doc = await PDFDocument.load(out, { updateMetadata: false });
  assert.equal(doc.getPageCount(), 3);
  assert.equal(doc.getPage(0).getRotation().angle, 90);
  assert.deepEqual([doc.getPage(1).getWidth(), doc.getPage(1).getHeight()], [842, 595]);
  assert.equal(doc.getProducer(), 'Kertas Office');
});

test('non-WinAnsi characters are replaced, not crashing', () => {
  assert.equal(toWinAnsi('café \u201cok\u201d ക'), 'café "ok" ?');
});
