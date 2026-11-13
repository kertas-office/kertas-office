// PDF, end to end: page tools, typed text, form filling and attached files through the
// interface, then the saved PDF read back with pdf-lib and pdf.js (not with Kertas).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument } from '@cantoo/pdf-lib';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { openApp, openFile, openFixture, saveBytes, fixture, stopServer, sleep } from './harness.mjs';

let page;
before(async () => { page = await openApp(); });
after(async () => { await page?.close(); stopServer(); });

async function ribbon(tab, id) {
  await page.locator(`.rb-tab[data-tab="${tab}"]:visible`).click();
  await page.locator(`[data-id="${id}"]:visible`).first().click();
  await sleep(400);
}

async function textOf(bytes) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map((t) => t.str).join(' '));
  return out;
}

test('pages: rotate, delete, and type text on a page', async () => {
  await openFixture(page, 'pdf', 'sample.pdf');
  await ribbon('organize', 'p-rotr');
  await page.locator('.pdf-thumbs .thumb').nth(2).click();
  await ribbon('organize', 'p-del');

  await ribbon('home', 'tool-text');
  const first = page.locator('.pdf-page:visible').first();
  await first.scrollIntoViewIfNeeded();
  await first.click({ position: { x: 120, y: 120 } });
  await page.keyboard.type('Checked by Kertas', { delay: 10 });
  await page.locator('.pdf-thumbs').click({ position: { x: 5, y: 5 } }); // click away commits the text
  await sleep(400);

  const bytes = await saveBytes(page, 'pdf');
  const doc = await PDFDocument.load(bytes);
  assert.equal(doc.getPageCount(), 2, 'one of three pages deleted');
  assert.equal(doc.getPage(0).getRotation().angle % 360, 90, `first page rotated: ${doc.getPages().map((p) => p.getRotation().angle)}`);
  assert.match((await textOf(bytes))[0], /Checked by Kertas/);
  assert.deepEqual(page.errors, []);
});

test('forms: a filled form keeps its values', async () => {
  await openFixture(page, 'pdf', 'form.pdf');
  const layer = page.locator('.form-layer:visible');
  await layer.locator('input[type=text], input:not([type])').first().fill('Aisha Rahman');
  await layer.locator('input[type=checkbox]').first().check();
  await sleep(300);

  const form = (await PDFDocument.load(await saveBytes(page, 'pdf'))).getForm();
  assert.equal(form.getTextField('name').getText(), 'Aisha Rahman');
  assert.equal(form.getCheckBox('terms').isChecked(), true);
  assert.deepEqual(page.errors, []);
});

test('attachments: attach, save, reopen, remove', async () => {
  await openFixture(page, 'pdf', 'sample.pdf');
  await page.locator('[data-side="attach"]:visible').click();
  const chooser = page.waitForEvent('filechooser');
  await page.locator('[data-att-add]:visible').click();
  await (await chooser).setFiles([fixture('sample.xlsx'), fixture('logo.png')]);
  await sleep(600);
  assert.match(await page.locator('.pdf-attach:visible').innerText(), /sample\.xlsx[\s\S]*logo\.png|logo\.png[\s\S]*sample\.xlsx/);

  const saved = await saveBytes(page, 'pdf');
  await openFile(page, 'pdf', 'with-files.pdf', saved);
  await page.locator('[data-side="attach"]:visible').click();
  const panel = page.locator('.pdf-attach:visible');
  assert.match(await panel.innerText(), /sample\.xlsx/);
  const row = panel.locator('.att-row', { hasText: 'logo.png' });
  await row.hover();
  await row.locator('[data-att="del"]').click();
  await sleep(300);

  const doc = await pdfjs.getDocument({ data: new Uint8Array(await saveBytes(page, 'pdf')), isEvalSupported: false }).promise;
  const names = [...((await doc.getAttachments()) || new Map()).values()].map((a) => a.filename);
  assert.deepEqual(names, ['sample.xlsx']);
  assert.deepEqual(page.errors, []);
});

test('merge: another PDF is appended from the Organize tab', async () => {
  await openFixture(page, 'pdf', 'sample.pdf');
  const chooser = page.waitForEvent('filechooser');
  await ribbon('organize', 'o-merge');
  await (await chooser).setFiles([fixture('extra.pdf')]);
  await page.waitForFunction(() => [...document.querySelectorAll('.pdf-thumbs .thumb')].filter((t) => t.offsetParent).length > 3, null, { timeout: 15000 });
  const extra = (await PDFDocument.load(await (await import('node:fs/promises')).readFile(fixture('extra.pdf')))).getPageCount();
  const doc = await PDFDocument.load(await saveBytes(page, 'pdf'));
  assert.equal(doc.getPageCount(), 3 + extra);
  assert.deepEqual(page.errors, []);
});

