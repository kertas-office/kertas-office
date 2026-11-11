import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { PDFDocument, PDFName, StandardFonts } from '@cantoo/pdf-lib';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import {
  buildPdf, parseRanges, splitEvery, fillTokens, listFormFields, imagesToPdf, readOutline,
} from '../src/lib/pdf-ops.js';

const fx = async (n) => new Uint8Array(await readFile(new URL(`./fixtures/${n}`, import.meta.url)));
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64'));
const FONTS = new URL('../node_modules/pdfjs-dist/standard_fonts/', import.meta.url).pathname;
const model = (n, src = 's0') => Array.from({ length: n }, (_, i) => ({ src, index: i, rotate: 0, annots: [] }));

async function pageTexts(bytes, password) {
  const task = pdfjs.getDocument({ data: bytes.slice(), password, isEvalSupported: false, standardFontDataUrl: FONTS });
  const doc = await task.promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const tc = await (await doc.getPage(i)).getTextContent();
    out.push(tc.items.map((t) => t.str).join(' '));
  }
  await task.destroy();
  return out;
}

async function formFixture() {
  const doc = await PDFDocument.create();
  const page = doc.addPage([595, 842]);
  const form = doc.getForm();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  page.drawText('Application form', { x: 50, y: 790, size: 18, font });
  const name = form.createTextField('applicant.name');
  name.addToPage(page, { x: 50, y: 700, width: 300, height: 24 });
  const agree = form.createCheckBox('agree');
  agree.addToPage(page, { x: 50, y: 650, width: 16, height: 16 });
  const plan = form.createRadioGroup('plan');
  plan.addOptionToPage('basic', page, { x: 50, y: 600, width: 14, height: 14 });
  plan.addOptionToPage('pro', page, { x: 120, y: 600, width: 14, height: 14 });
  const country = form.createDropdown('country');
  country.addOptions(['Malaysia', 'Qatar', 'India']);
  country.addToPage(page, { x: 50, y: 540, width: 200, height: 24 });
  doc.addPage([595, 842]);
  return doc.save();
}

test('page range parsing and splitting', () => {
  assert.deepEqual(parseRanges('1-3, 5', 6), [0, 1, 2, 4]);
  assert.deepEqual(parseRanges('4-', 6), [3, 4, 5]);
  assert.deepEqual(parseRanges('3-1', 6), [2, 1, 0]);
  assert.deepEqual(parseRanges('2,2,2', 6), [1]);
  assert.throws(() => parseRanges('9', 6), /outside/);
  assert.throws(() => parseRanges('a', 6), /not a page/);
  assert.deepEqual(splitEvery(5, 2), [[0, 1], [2, 3], [4]]);
  assert.equal(fillTokens('Page {n} of {total} - {file}', { n: 2, total: 9, date: 'x', file: 'a.pdf' }), 'Page 2 of 9 - a.pdf');
});

test('password protect, then reopen with the password', async () => {
  const src = await fx('sample.pdf');
  const out = await buildPdf({ s0: src }, model(3), { encrypt: { userPassword: 'kertas', permissions: { copying: false } } });
  await assert.rejects(PDFDocument.load(out), /encrypted/i);
  assert.equal((await PDFDocument.load(out, { password: 'kertas' })).getPageCount(), 3);
  // An open password alone must not switch off printing or copying
  const openOnly = await buildPdf({ s0: src }, model(3), { encrypt: { userPassword: 'kertas' } });
  const p = (await pdfjs.getDocument({ data: openOnly.slice(), password: 'kertas', isEvalSupported: false }).promise);
  const perms = await p.getPermissions();
  for (const flag of [4, 8, 16, 32, 256, 2048]) assert.ok(perms.has(flag), `permission ${flag} granted`);
  await p.destroy?.();
  // Saving a protected source with its password removes protection unless asked to keep it
  const plain = await buildPdf({ s0: out }, model(3), { passwords: { s0: 'kertas' } });
  assert.equal((await PDFDocument.load(plain)).isEncrypted, false);
  await assert.rejects(buildPdf({ s0: out }, model(3), { passwords: { s0: 'nope' } }), /incorrect/);
});

test('watermark and header/footer text lands on every page', async () => {
  const out = await buildPdf({ s0: await fx('sample.pdf') }, model(3), {
    watermark: { text: 'CONFIDENTIAL', opacity: 0.2 },
    headerFooter: { header: ['', 'Kertas Ltd', ''], footer: ['{file}', '', 'Page {n} of {total}'], file: 'sample.pdf' },
  });
  const texts = await pageTexts(out);
  texts.forEach((t, i) => {
    assert.match(t, /CONFIDENTIAL/);
    assert.match(t, /Kertas Ltd/);
    assert.match(t, new RegExp(`Page ${i + 1} of 3`));
  });
});

test('header/footer also works on rotated pages', async () => {
  const pages = model(3); pages[1].rotate = 90;
  const out = await buildPdf({ s0: await fx('sample.pdf') }, pages, { headerFooter: { footer: ['', 'Page {n}', ''], start: 5 } });
  const texts = await pageTexts(out);
  assert.match(texts[1], /Page 6/);
});

test('forms: list, fill in place (still editable), flatten when pages are merged', async () => {
  const bytes = await formFixture();
  const fields = await listFormFields(bytes);
  const byName = Object.fromEntries(fields.map((f) => [f.name, f]));
  assert.equal(byName['applicant.name'].type, 'text');
  assert.equal(byName.plan.widgets.length, 2);
  assert.deepEqual(byName.plan.widgets.map((w) => w.onValue), ['basic', 'pro']);
  assert.deepEqual(byName.country.options, ['Malaysia', 'Qatar', 'India']);
  assert.equal(byName.agree.widgets[0].page, 0);

  const values = { s0: { 'applicant.name': 'Hashir', agree: true, plan: 'pro', country: 'Qatar' } };
  const filled = await buildPdf({ s0: bytes }, model(2), { formValues: values });
  const again = Object.fromEntries((await listFormFields(filled)).map((f) => [f.name, f.value]));
  assert.equal(again['applicant.name'], 'Hashir');
  assert.equal(again.agree, true);
  assert.equal(again.plan, 'pro');
  assert.deepEqual(again.country, ['Qatar']);

  // Reordering forces a new document: values must still be visible as page content
  const reordered = [{ src: 's0', index: 1, rotate: 0, annots: [] }, { src: 's0', index: 0, rotate: 0, annots: [] }];
  const flat = await buildPdf({ s0: bytes }, reordered, { formValues: values });
  assert.equal((await listFormFields(flat)).length, 0);
  assert.match((await pageTexts(flat))[1], /Hashir/);
});

test('sticky notes become real PDF comments', async () => {
  const pages = model(1, 's0').slice(0, 1);
  pages[0].annots.push({ type: 'note', x: 100, y: 700, text: 'Please check the total', color: '#facc15' });
  const out = await buildPdf({ s0: await fx('sample.pdf') }, [...pages, ...model(3).slice(1)]);
  const doc = await PDFDocument.load(out);
  const annots = doc.getPage(0).node.Annots();
  assert.equal(annots.size(), 1);
  const note = doc.context.lookup(annots.get(0));
  assert.equal(note.get(PDFName.of('Subtype')).toString(), '/Text');
  assert.equal(note.lookup(PDFName.of('Contents')).decodeText(), 'Please check the total');
});

test('every annotation type draws without error', async () => {
  const pages = model(3);
  pages[0].annots.push(
    { type: 'ellipse', x1: 50, y1: 50, x2: 150, y2: 100, color: '#2563eb', width: 2 },
    { type: 'line', x1: 50, y1: 200, x2: 250, y2: 260, color: '#e11d48', width: 2, arrow: true },
    { type: 'markup', style: 'underline', rects: [[72, 725, 230, 742]], color: '#2563eb' },
    { type: 'markup', style: 'strike', rects: [[72, 725, 230, 742]], color: '#e11d48' },
    { type: 'markup', style: 'highlight', rects: [[72, 725, 230, 742]], color: '#facc15' },
    { type: 'stamp', x: 300, y: 500, w: 140, h: 44, rot: 0, label: 'APPROVED', sub: '09/11/2026', color: '#16a34a' },
    { type: 'stamp', x: 300, y: 400, w: 140, h: 30, rot: 90, label: 'DRAFT', color: '#2563eb' },
    { type: 'image', x: 300, y: 100, w: 50, h: 50, img: PNG, fmt: 'png' },
  );
  const texts = await pageTexts(await buildPdf({ s0: await fx('sample.pdf') }, pages));
  assert.match(texts[0], /APPROVED/);
  assert.match(texts[0], /DRAFT/);
});

test('redaction replaces the page with an image so hidden text is gone', async () => {
  const pages = model(3);
  pages[0].annots.push({ type: 'redact', x1: 60, y1: 720, x2: 400, y2: 745 });
  const out = await buildPdf({ s0: await fx('sample.pdf') }, pages, { rasters: { 0: { bytes: PNG, fmt: 'png' } } });
  const texts = await pageTexts(out);
  assert.doesNotMatch(texts[0], /Invoice/);
  assert.match(texts[1], /Invoice total/);
  const doc = await PDFDocument.load(out);
  assert.equal(doc.getPageCount(), 3);
  assert.equal(Math.round(doc.getPage(0).getWidth()), 595);
});

test('blank pages, images to PDF and bookmarks', async () => {
  const pages = [...model(3)];
  pages.splice(1, 0, { kind: 'blank', width: 595.28, height: 841.89, rotate: 0, annots: [] });
  const outline = [{ title: 'Intro', page: 0, children: [{ title: 'Details', page: 2, children: [] }] }, { title: 'End', page: 3, children: [] }];
  const out = await buildPdf({ s0: await fx('sample.pdf') }, pages, { outline });
  assert.equal((await PDFDocument.load(out)).getPageCount(), 4);
  assert.deepEqual(await readOutline(out), outline);

  const imgPdf = await imagesToPdf([{ bytes: PNG, fmt: 'png' }, { bytes: PNG, fmt: 'png' }], { pageSize: 'a4' });
  assert.equal((await PDFDocument.load(imgPdf)).getPageCount(), 2);
});

test('in-place editing keeps the original bookmarks', async () => {
  const withOutline = await buildPdf({ s0: await fx('sample.pdf') }, model(3), { outline: [{ title: 'Start', page: 1, children: [] }] });
  const edited = await buildPdf({ s0: withOutline }, model(3).map((p) => ({ ...p, annots: [{ type: 'rect', x1: 1, y1: 1, x2: 9, y2: 9, color: '#000' }] })));
  assert.deepEqual(await readOutline(edited), [{ title: 'Start', page: 1, children: [] }]);
});
