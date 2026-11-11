// Layers (optional content groups): reading, creating, moving, deleting, renaming, visibility.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { PDFDocument, PDFName, PDFDict, PDFArray } from '@cantoo/pdf-lib';
import { buildPdf } from '../src/lib/pdf-ops.js';
import { listLayers } from '../src/lib/pdf-layers.js';
import { tokenize } from '../src/lib/pdf-textedit.js';

const layered = new Uint8Array(readFileSync(new URL('./fixtures/layered.pdf', import.meta.url)));
const png = new Uint8Array(readFileSync(new URL('./fixtures/logo.png', import.meta.url)));

async function content(doc, i = 0) {
  const page = doc.getPage(i);
  const c = page.node.get(PDFName.of('Contents'));
  const refs = c instanceof PDFArray ? c.asArray() : [c];
  const { decodePDFRawStream, PDFRawStream } = await import('@cantoo/pdf-lib');
  return refs.map((r) => { const s = doc.context.lookup(r); const b = s instanceof PDFRawStream ? decodePDFRawStream(s).decode() : s.getContents(); return Buffer.from(b).toString('latin1'); }).join('\n');
}
const dParts = (doc) => {
  const d = doc.catalog.lookup(PDFName.of('OCProperties'), PDFDict).lookup(PDFName.of('D'), PDFDict);
  const names = (k) => (d.lookupMaybe(PDFName.of(k), PDFArray)?.asArray() || []).map((r) => doc.context.lookup(r).get(PDFName.of('Name')).decodeText());
  return { on: names('ON'), off: names('OFF'), order: names('Order').filter(Boolean) };
};

test('reads the layers another app wrote', async () => {
  const doc = await PDFDocument.load(layered);
  assert.deepEqual(listLayers(doc).map((l) => l.name), ['Background', 'Text', 'Artwork', 'Draft notes']);
});

test('new layer: items are wrapped in marked content and listed in the catalogue', async () => {
  const src = { s0: layered };
  const doc0 = await PDFDocument.load(layered);
  const existing = listLayers(doc0).map((l) => ({ key: `s0|${l.id}`, src: 's0', id: l.id, name: l.name, origName: l.name, visible: l.name !== 'Draft notes', deleted: false }));
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [
    { type: 'image', x: 300, y: 100, w: 80, h: 40, rot: 0, img: png, fmt: 'png', layer: 'new1' },
    { type: 'rect', x1: 50, y1: 50, x2: 90, y2: 90, color: '#ff0000', width: 2, layer: 'new1' },
    { type: 'text', x: 60, y: 300, size: 12, color: '#000000', text: 'on Text', rot: 0, layer: existing[1].key },
  ] }];
  const out = await buildPdf(src, pages, { layers: { order: ['new1', ...existing.map((l) => l.key)], existing, created: [{ key: 'new1', name: 'Stamps & photos', visible: true }] } });
  const doc = await PDFDocument.load(out);
  const names = listLayers(doc).map((l) => l.name);
  assert.ok(names.includes('Stamps & photos'), names.join());
  assert.equal(names.length, 5);
  const parts = dParts(doc);
  assert.equal(parts.order[0], 'Stamps & photos');
  assert.deepEqual(parts.off, ['Draft notes']);
  const props = doc.getPage(0).node.Resources().lookup(PDFName.of('Properties'), PDFDict);
  const ocNames = props.entries().filter(([, v]) => doc.context.lookup(v).get(PDFName.of('Name'))?.decodeText?.() === 'Stamps & photos').map(([k]) => k.asString().slice(1));
  assert.equal(ocNames.length, 1);
  const c = await content(doc);
  assert.match(c, new RegExp(`/OC /${ocNames[0]} BDC[\\s\\S]*Do[\\s\\S]*EMC`));
  // the text went onto the existing Text layer (MC1)
  const toks = tokenize(c);
  assert.ok(toks.some((t, i) => t.v === 'BDC' && toks[i - 1].v === 'MC1') , 'existing layer reused');
});

test('existing layer moved: its content gets a transform; deleted layer: its content is gone', async () => {
  const doc0 = await PDFDocument.load(layered);
  const ls = listLayers(doc0);
  const key = (n) => `s0|${ls.find((l) => l.name === n).id}`;
  const existing = ls.map((l) => ({ key: `s0|${l.id}`, src: 's0', id: l.id, name: l.name === 'Artwork' ? 'Pictures' : l.name, origName: l.name, visible: true, deleted: l.name === 'Draft notes' }));
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [], layerMoves: { [key('Artwork')]: [0.5, 0, 0, 0.5, 100, 50] } }];
  const out = await buildPdf({ s0: layered }, pages, { layers: { order: existing.filter((l) => !l.deleted).map((l) => l.key), existing, created: [] } });
  const doc = await PDFDocument.load(out);
  const c = await content(doc);
  const hex = (t) => Buffer.from(t, 'latin1').toString('hex');
  assert.ok(!c.includes(hex('DRAFT')) && !c.includes('DRAFT'), 'deleted layer content removed');
  assert.ok(c.includes(hex('Unlayered footer text')) || c.includes('Unlayered footer text'), 'content outside layers kept');
  assert.equal((c.match(/0\.5 0 0 0\.5 100 50 cm/g) || []).length, 2, 'circle and box');
  assert.match(c, /q 0\.5 0 0 0\.5 [-\d.]+ [-\d.]+ cm\s*\/fzImg0 Do\s*Q/, 'the image (an /OC image XObject) moves too');
  const names = listLayers(doc).map((l) => l.name);
  assert.deepEqual(names, ['Background', 'Text', 'Pictures']);
});

test('assembled documents keep their layers', async () => {
  const doc0 = await PDFDocument.load(layered);
  const ls = listLayers(doc0);
  const existing = ls.map((l) => ({ key: `s0|${l.id}`, src: 's0', id: l.id, name: l.name, origName: l.name, visible: l.name !== 'Text', deleted: false }));
  // the page twice: a new document is assembled
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [] }, { src: 's0', index: 0, rotate: 90, annots: [] }];
  const out = await buildPdf({ s0: layered }, pages, { layers: { order: existing.map((l) => l.key), existing, created: [] } });
  const doc = await PDFDocument.load(out);
  assert.deepEqual(listLayers(doc).map((l) => l.name).sort(), ['Artwork', 'Background', 'Draft notes', 'Text']);
  assert.deepEqual(dParts(doc).off, ['Text']);
});

// Illustrator and InDesign put layer marks inside text objects: BT ... /OC /L BDC ... EMC ... ET
async function textLayerPdf() {
  const { PDFDocument: D, PDFName: N, PDFHexString, StandardFonts } = await import('@cantoo/pdf-lib');
  const doc = await D.create();
  const page = doc.addPage([400, 400]);
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const ocg = doc.context.register(doc.context.obj({ Type: 'OCG', Name: PDFHexString.fromText('Words') }));
  const res = page.node.Resources();
  res.set(N.of('Properties'), doc.context.obj({ L1: ocg }));
  const fd = doc.context.obj({}); fd.set(N.of('F1'), font.ref); res.set(N.of('Font'), fd);
  doc.catalog.set(N.of('OCProperties'), doc.context.obj({ OCGs: [ocg], D: { ON: [ocg], Order: [ocg] } }));
  const cs = [
    'BT /F1 20 Tf 40 300 Td',
    '/OC /L1 BDC 1 0 0 rg (Moved) Tj EMC', // a whole line on the layer
    '0 -40 Td (Stays) Tj', // relies on the red set inside the block
    '0 -40 Td (Mid) Tj /OC /L1 BDC ( line) Tj EMC', // starts in the middle of a line
    'ET',
  ].join('\n');
  page.node.set(N.of('Contents'), doc.context.register(doc.context.flateStream(new TextEncoder().encode(cs))));
  return { bytes: await doc.save(), ocgId: `${ocg.objectNumber}R` };
}

test('layer content inside a text object moves, and the text after it keeps its colour and place', async () => {
  const { bytes, ocgId } = await textLayerPdf();
  const key = `s0|${ocgId}`;
  const existing = [{ key, src: 's0', id: ocgId, name: 'Words', origName: 'Words', visible: true, deleted: false }];
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [], layerMoves: { [key]: [1, 0, 0, 1, 100, -50] } }];
  const out = await buildPdf({ s0: bytes }, pages, { layers: { order: [key], existing, created: [] } });
  const doc = await PDFDocument.load(out);
  const c = await content(doc);
  // the first block moved; the one in the middle of a line was left alone
  assert.equal((c.match(/ET\s*q 1 0 0 1 100 -50 cm\s*BT/g) || []).length, 1, c);
  assert.match(c, /ET Q\s*1 0 0 rg\s*BT/, 'red is set again after the moved block');
  // where things land, read back by poppler
  const { execFileSync } = await import('node:child_process');
  const { writeFileSync, mkdtempSync } = await import('node:fs');
  const dir = mkdtempSync('/tmp/kly-');
  writeFileSync(`${dir}/o.pdf`, out);
  let html;
  try { html = execFileSync('pdftotext', ['-bbox', `${dir}/o.pdf`, '-'], { encoding: 'utf8' }); } catch { return; } // poppler not installed
  const word = (w) => { const m = new RegExp(`xMin="([\\d.]+)" yMin="([\\d.]+)"[^>]*>${w}<`).exec(html); return m ? [Number(m[1]), Number(m[2])] : null; };
  const moved = word('Moved'); const stays = word('Stays'); const mid = word('Mid');
  assert.ok(moved && Math.abs(moved[0] - 140) < 1, `Moved at ${moved}`);
  assert.ok(Math.abs(moved[1] - (stays[1] + 10)) < 2, `Moved went 50pt down (now 10pt below Stays): ${moved} vs ${stays}`);
  assert.ok(Math.abs(stays[0] - 40) < 1, `Stays did not move: ${stays}`);
  assert.ok(Math.abs(mid[0] - 40) < 1, `Mid did not move: ${mid}`);
});

// Some apps draw each layer's content inside a reusable form XObject
test('layer content inside a form XObject moves and is deleted', async () => {
  const { PDFDocument: D, PDFName: N, PDFHexString } = await import('@cantoo/pdf-lib');
  const doc = await D.create();
  const page = doc.addPage([400, 400]);
  const mk = (name) => doc.context.register(doc.context.obj({ Type: 'OCG', Name: PDFHexString.fromText(name) }));
  const shapes = mk('Shapes'); const notes = mk('Notes');
  const formContent = '/OC /S BDC 0 0 1 rg 10 10 50 50 re f EMC\n/OC /N BDC 1 0 0 rg 100 10 50 50 re f EMC';
  const form = doc.context.flateStream(new TextEncoder().encode(formContent), {
    Type: 'XObject', Subtype: 'Form', BBox: [0, 0, 400, 400], Matrix: [1, 0, 0, 1, 20, 30],
    Resources: { Properties: { S: shapes, N: notes } },
  });
  const formRef = doc.context.register(form);
  const res = page.node.Resources();
  const xo = doc.context.obj({}); xo.set(N.of('Fm1'), formRef); res.set(N.of('XObject'), xo);
  doc.catalog.set(N.of('OCProperties'), doc.context.obj({ OCGs: [shapes, notes], D: { ON: [shapes, notes], Order: [shapes, notes] } }));
  page.node.set(N.of('Contents'), doc.context.register(doc.context.flateStream(new TextEncoder().encode('q /Fm1 Do Q'))));
  const bytes = await doc.save();
  const id = (r) => `${r.objectNumber}R`;
  const existing = [
    { key: `s0|${id(shapes)}`, src: 's0', id: id(shapes), name: 'Shapes', origName: 'Shapes', visible: true, deleted: false },
    { key: `s0|${id(notes)}`, src: 's0', id: id(notes), name: 'Notes', origName: 'Notes', visible: true, deleted: true },
  ];
  const pages = [{ src: 's0', index: 0, rotate: 0, annots: [], layerMoves: { [existing[0].key]: [1, 0, 0, 1, 100, 200] } }];
  const out = await buildPdf({ s0: bytes }, pages, { layers: { order: [existing[0].key], existing, created: [] } });
  const back = await PDFDocument.load(out);
  const xobj = back.getPage(0).node.Resources().lookup(N.of('XObject'), PDFDict).lookup(N.of('Fm1'));
  const { decodePDFRawStream } = await import('@cantoo/pdf-lib');
  const text = Buffer.from(decodePDFRawStream(xobj).decode()).toString('latin1');
  // the move is in page space: inside the form it reads the same because the form only translates
  assert.match(text, /q 1 0 0 1 100 200 cm\s*0 0 1 rg 10 10 50 50 re f\s*Q/, text);
  assert.doesNotMatch(text, /1 0 0 rg/, 'the deleted layer is gone');
  assert.ok(xobj.dict.get(N.of('Matrix')), 'the form keeps its matrix');
});
