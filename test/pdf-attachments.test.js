import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildPdf } from '../src/lib/pdf-ops.js';

const fx = (n) => readFile(new URL(`./fixtures/${n}`, import.meta.url));
const csv = new TextEncoder().encode('name,qty\npen,3\n');

async function attachedFiles(bytes) {
  const doc = await pdfjs.getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
  const found = await doc.getAttachments();
  const out = {};
  for (const [id, a] of found instanceof Map ? found : Object.entries(found || {})) {
    out[a.filename] = { data: a.content || await doc.getAttachmentContent(id), description: a.description };
  }
  await doc.loadingTask?.destroy?.();
  return out;
}

test('files attached to a PDF are written, read back and can be removed', async () => {
  const sources = { s0: new Uint8Array(await fx('sample.pdf')) };
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [] }];
  const one = await buildPdf(sources, pages, { attachments: [{ name: 'data.csv', data: csv, description: 'Stock' }] });
  const got = await attachedFiles(one);
  assert.deepEqual(Object.keys(got), ['data.csv']);
  assert.equal(new TextDecoder().decode(got['data.csv'].data), 'name,qty\npen,3\n');
  assert.equal(got['data.csv'].description, 'Stock');

  // saving again with the list edited replaces the old set instead of adding to it
  const two = await buildPdf({ s0: one }, pages, { attachments: [] });
  assert.deepEqual(Object.keys(await attachedFiles(two)), []);
});
