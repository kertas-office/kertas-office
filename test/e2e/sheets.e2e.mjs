// Sheets, end to end: typing into cells, formulas, named ranges and comments through the
// interface, then the saved .xlsx read back with ExcelJS and JSZip (not with Kertas).
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { openApp, openFile, openFixture, saveBytes, fixture, stopServer, sleep } from './harness.mjs';

let page;
before(async () => { page = await openApp(); });
after(async () => { await page?.close(); stopServer(); });

const readBook = async (bytes) => { const wb = new ExcelJS.Workbook(); await wb.xlsx.load(bytes); return wb; };

/** Selects cells by typing their address in the name box, as a person would. */
async function goTo(ref) {
  const box = page.locator('[data-u-comp=defined-name] input').first();
  await box.click();
  await box.fill(ref);
  await box.press('Enter');
  await sleep(250);
}

/** Types into the selected cell and moves down. */
async function enter(text) {
  await page.keyboard.type(text, { delay: 10 });
  await page.keyboard.press('Enter');
  await sleep(150);
}

async function ribbon(tab, id) {
  await page.locator(`.rb-tab[data-tab="${tab}"]`).click();
  await page.locator(`[data-id="${id}"]`).first().click();
  await sleep(400);
}

test('a new workbook: values and a formula are saved and calculated', async () => {
  await openFile(page, 'sheets', 'Budget.xlsx');
  await goTo('A1');
  await enter('Price');
  await enter('5');
  await enter('7');
  await enter('=SUM(A2:A3)');

  const ws = (await readBook(await saveBytes(page, 'xlsx'))).worksheets[0];
  assert.equal(ws.getCell('A1').value, 'Price');
  assert.equal(ws.getCell('A2').value, 5);
  assert.equal(ws.getCell('A4').formula, 'SUM(A2:A3)');
  assert.equal(ws.getCell('A4').result, 12);
  assert.deepEqual(page.errors, []);
});

test('a named range is defined from the Formulas tab and saved', async () => {
  await goTo('A2:A3');
  await ribbon('formulas', 'name-define');
  const dlg = page.locator('.modal.dialog');
  await dlg.locator('[data-f="name"]').fill('Prices');
  await dlg.locator('button.primary').click();
  await sleep(400);
  await goTo('B1');
  await enter('=SUM(Prices)');

  const wb = await readBook(await saveBytes(page, 'xlsx'));
  const names = wb.definedNames.model.map((n) => n.name);
  assert.ok(names.includes('Prices'), `defined names: ${names}`);
  assert.equal(wb.worksheets[0].getCell('B1').result, 12);
});

test('a comment is added from the Review tab and saved as a threaded comment', async () => {
  await goTo('A1');
  await ribbon('review', 'comment-new');
  await page.getByText('Write a comment or a reply').first().waitFor({ state: 'visible' });
  await sleep(300);
  await page.keyboard.type('Check this price', { delay: 10 });
  await page.getByRole('button', { name: /^Post$/ }).first().click();
  await sleep(500);

  const zip = await JSZip.loadAsync(await saveBytes(page, 'xlsx'));
  const parts = Object.keys(zip.files);
  const threaded = parts.find((p) => /^xl\/threadedComments\/threadedComment\d+\.xml$/.test(p));
  assert.ok(threaded, `threaded comment part in ${parts.join(', ')}`);
  assert.match(await zip.file(threaded).async('string'), /Check this price/);
  assert.ok(parts.includes('xl/persons/person.xml'), 'comment authors');
  assert.deepEqual(page.errors, []);
});

test('saved workbook reopens with values, name and comment', async () => {
  const bytes = await saveBytes(page, 'xlsx');
  await openFile(page, 'sheets', 'Reopened.xlsx', bytes);
  const again = await readBook(await saveBytes(page, 'xlsx'));
  const ws = again.worksheets[0];
  assert.equal(ws.getCell('A4').result, 12);
  assert.equal(ws.getCell('B1').formula, 'SUM(Prices)');
  const zip = await JSZip.loadAsync(await saveBytes(page, 'xlsx'));
  assert.ok(Object.keys(zip.files).some((p) => p.startsWith('xl/threadedComments/')), 'comment survives a second save');
  assert.deepEqual(page.errors, []);
});

test('CSV: the current sheet saves as CSV', async () => {
  const csv = (await saveBytes(page, 'csv')).toString('utf8').replace(/^﻿/, '');
  assert.match(csv.split(/\r?\n/)[0], /^Price/);
  assert.match(csv, /\n5,?/);
});

// Cells with a value or formula in each sheet part, read from the XML itself (ExcelJS cannot
// load every real-world file, for example ones with some drawings)
async function filledCells(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const out = [];
  for (const p of Object.keys(zip.files).filter((f) => /^xl\/worksheets\/sheet\d+\.xml$/.test(f)).sort()) {
    const xml = await zip.file(p).async('string');
    out.push((xml.match(/<c [^>]*>(?:(?!<\/c>).)*?<(?:v|f|is)\b/gs) || []).length);
  }
  return out;
}

test('existing workbooks open and round-trip their cells', async () => {
  for (const name of ['sample.xlsx', 'features.xlsx']) {
    const want = await filledCells(await readFile(fixture(name)));
    await openFixture(page, 'sheets', name);
    const got = await filledCells(await saveBytes(page, 'xlsx'));
    assert.equal(got.length, want.length, `${name}: sheets`);
    want.forEach((n, i) => assert.ok(got[i] >= n, `${name} sheet ${i + 1}: ${got[i]} of ${n} cells`));
  }
  assert.deepEqual(page.errors, []);
});
