// The new Sheets and PDF editors (trial), end to end: the desktop app with spreadsheets and
// PDFs opening in them, driven the way a person would and checked in the files that get saved.
// Needs the editors built (node scripts/build-writer2.cjs); skipped when they are not there.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync, readdirSync } from 'node:fs';
import { mkdtemp, copyFile, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import ExcelJS from 'exceljs';
import { PDFDocument } from '@cantoo/pdf-lib';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const W2 = path.join(ROOT, 'electron', 'writer2');
const hasSheets = existsSync(path.join(W2, 'sheets', 'renderer', 'index.html'))
  && existsSync(path.join(W2, 'native')) && readdirSync(path.join(W2, 'native')).some((f) => /^xlsx-sidecar(\.exe)?$/.test(f));
const hasPdf = existsSync(path.join(W2, 'pdf', 'renderer', 'index.html'));
const sheetsOpts = { skip: hasSheets ? false : 'the new Sheets is not built (node scripts/build-writer2.cjs, with cargo installed)' };
const pdfOpts = { skip: hasPdf ? false : 'the new PDF editor is not built (node scripts/build-writer2.cjs)' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let app;
let k; // Kertas's own page: tab strip, Home, palette
let w; // the editor of the file that is open
let dir;
let file;
let n = 0;

before(async () => {
  if (!hasSheets && !hasPdf) return;
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-ne-'));
  app = await _electron.launch({
    executablePath: require('electron'),
    args: [ROOT, '--no-sandbox', '--window-size=1400,1000'],
    env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production', KERTAS_NEW_SHEETS: '1', KERTAS_NEW_PDF: '1' },
  });
  k = await app.firstWindow();
  k.setDefaultTimeout(15000);
  await k.waitForSelector('#tabs');
});

after(async () => {
  if (!app) return;
  const child = app.process();
  try { await Promise.race([app.close(), sleep(5000)]); } catch { /* closing */ }
  try { child?.kill('SIGKILL'); } catch { /* gone */ }
  await rm(dir, { recursive: true, force: true }).catch(() => {});
  setTimeout(() => process.exit(), 3000).unref();
});

const editors = (kind) => app.context().pages().filter((p) => p.url().startsWith(`kertas-writer://${kind}/`));

/** Native dialogs answered by the test: Open gives these files, Save a path in the temp folder, questions "don't save". */
async function answerDialogs(openPaths = [], saveName = null) {
  await app.evaluate(({ dialog }, [openPaths, dir, saveName]) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: openPaths });
    dialog.showSaveDialog = async (_w, o) => ({ canceled: false, filePath: `${dir}/${saveName || String((o || _w)?.defaultPath || 'file').split(/[\\/]/).pop()}` });
    dialog.showMessageBox = async (_w, o) => {
      const buttons = (o || _w)?.buttons || [];
      const dont = buttons.findIndex((b) => /don.t save|discard/i.test(b));
      return { response: dont >= 0 ? dont : 0, checkboxChecked: false };
    };
  }, [openPaths, dir, saveName]);
}

async function closeAll() {
  while (await k.locator('#tabs .tab .close').count()) {
    await answerDialogs();
    await k.locator('#tabs .tab .close').first().click();
    await sleep(900);
  }
}

/** A fresh copy of a fixture, opened through Kertas's own Open command; `w` becomes its editor. */
async function open(name, kind) {
  await closeAll();
  file = path.join(dir, `${++n}-${name}`);
  await copyFile(fixture(name), file);
  await answerDialogs([file]);
  const had = new Set(editors(kind));
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+o');
  w = null;
  for (let i = 0; i < 80 && !w; i++) { w = editors(kind).find((p) => !had.has(p)); if (!w) await sleep(250); }
  assert.ok(w, `${name} opened in the new ${kind} editor`);
  w.setDefaultTimeout(15000);
  w.errors = [];
  w.on('pageerror', (e) => w.errors.push(String(e.message || e)));
  await w.waitForSelector(kind === 'sheets' ? '#univer-container canvas' : '.ribbon-tabs');
  await sleep(kind === 'sheets' ? 3500 : 2500);
}

/** A key press that goes through the app the way the keyboard does (Kertas's shortcuts are caught there). */
async function press(kind, keyCode, modifiers = ['control']) {
  await app.evaluate(({ webContents }, [prefix, keyCode, modifiers]) => {
    const wc = webContents.getAllWebContents().filter((c) => c.getURL().startsWith(prefix)).at(-1);
    wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
    wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
  }, [`kertas-writer://${kind}/`, keyCode, modifiers]);
}
const dirtyTabs = () => k.locator('#tabs .tab.is-dirty').count();
const tabTitles = () => k.locator('#tabs .tab .title').allTextContents();
/** a cell of the first screen, by where the grid draws it at this window size (columns 69 px from 46, rows 20 px from 205) */
const CELLS = { G3: [588, 235], H2: [657, 215] };
async function goTo(cell) {
  await w.mouse.click(...CELLS[cell]);
  await sleep(300);
}
/** a ribbon tab of the PDF editor, without collapsing the ribbon by clicking the open one */
async function pdfTab(name) {
  const t = w.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) });
  if (!/\bactive\b/.test(await t.getAttribute('class')) || !(await w.locator('.ribbon-body').isVisible())) await t.click();
  await sleep(400);
}

test('neither editor\'s build can reach an AI service: no service addresses, no assistant channels, no sign-in', pdfOpts, async () => {
  const files = [];
  for (const kind of ['sheets', 'pdf']) {
    if (kind === 'sheets' && !hasSheets) continue;
    files.push(path.join(W2, kind, 'preload', 'index.js'));
    for (const name of await readdir(path.join(W2, kind, 'renderer', 'assets'))) if (/\.(c?js|mjs)$/.test(name)) files.push(path.join(W2, kind, 'renderer', 'assets', name));
  }
  // the services the assistant called, its channels to the main process, its sign-in
  const marks = /api\.openai\.com|api\.anthropic\.com|generativelanguage\.googleapis|api\.deepseek\.com|genspark\.ai|chat\/completions|["'`](?:ai|kb|project):[a-zA-Z-]+["'`]|gsk-login|x-api-key/;
  const found = [];
  for (const f of files) {
    const text = await readFile(f, 'utf8');
    const m = marks.exec(text);
    if (m) found.push(`${path.relative(W2, f)}: ${text.slice(Math.max(0, m.index - 40), m.index + 60).replace(/\s+/g, ' ')}`);
  }
  assert.deepEqual(found, []);
});

// ------------------------------------------------------------------ Sheets

test('a workbook opens in the new Sheets as a Kertas tab, with nothing of an assistant', sheetsOpts, async () => {
  await open('sample.xlsx', 'sheets');
  assert.deepEqual((await tabTitles()).slice(1), [path.basename(file)]);
  assert.equal(await w.title(), 'Sheets');
  const text = await w.locator('body').innerText();
  assert.doesNotMatch(text, /genspark|genoffice|\bAI\b|copilot|translate/i);
  assert.equal(await w.locator('.ai-entry, .qa-name-with-ai, .copilot').count(), 0);
  assert.deepEqual(w.errors, []);
});

test('typing a value and a formula, then Ctrl+S, saves them into the file', sheetsOpts, async () => {
  await open('sample.xlsx', 'sheets');
  await goTo('G3');
  await w.keyboard.type('Hello Kertas');
  await w.keyboard.press('Enter');
  await w.keyboard.type('=SUM(B2:B4)');
  await w.keyboard.press('Enter');
  await sleep(2200);
  assert.equal(await dirtyTabs(), 1, 'the tab shows unsaved changes');
  await press('sheets', 'S');
  await sleep(3500);
  assert.equal(await dirtyTabs(), 0, 'saved');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet('Sales');
  assert.equal(ws.getCell('G3').value, 'Hello Kertas');
  assert.equal(ws.getCell('G4').value.formula, 'SUM(B2:B4)');
  // what was there before is still there
  assert.equal(ws.getCell('A2').value, 'Pen');
  assert.equal(ws.getCell('D5').value.formula, 'SUM(D2:D4)');
  assert.ok(wb.getWorksheet('Notes'), 'the second sheet survived');
});

test('Ctrl+Shift+N opens a blank workbook in the new Sheets', sheetsOpts, async () => {
  await closeAll();
  const had = editors('sheets').length;
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+Shift+N');
  for (let i = 0; i < 60 && editors('sheets').length === had; i++) await sleep(250);
  assert.equal(editors('sheets').length, had + 1);
  assert.match((await tabTitles()).at(-1), /^Book\d+\.xlsx$/);
});

test('closing a workbook with unsaved changes asks first, and "Don\'t save" leaves the file alone', sheetsOpts, async () => {
  await open('sample.xlsx', 'sheets');
  const before = await readFile(file);
  await goTo('H2');
  await w.keyboard.type('not kept');
  await w.keyboard.press('Enter');
  await sleep(2200);
  assert.equal(await dirtyTabs(), 1);
  let asked = false;
  await app.evaluate(({ dialog }) => {
    globalThis.__asked = 0;
    dialog.showMessageBox = async (_w, o) => {
      globalThis.__asked++;
      const buttons = (o || _w)?.buttons || [];
      const dont = buttons.findIndex((b) => /don.t save|discard/i.test(b));
      return { response: dont >= 0 ? dont : 0, checkboxChecked: false };
    };
  });
  await k.locator('#tabs .tab .close').first().click();
  await sleep(1500);
  asked = (await app.evaluate(() => globalThis.__asked)) > 0;
  assert.ok(asked, 'a question was shown');
  assert.equal((await tabTitles()).length, 1, 'only Home is left');
  assert.ok(before.equals(await readFile(file)), 'the file is unchanged');
});

// --------------------------------------------------------------------- PDF

test('a PDF opens in the new PDF editor as a Kertas tab, with nothing of an assistant', pdfOpts, async () => {
  await open('sample.pdf', 'pdf');
  assert.deepEqual((await tabTitles()).slice(1), [path.basename(file)]);
  assert.equal(await w.title(), 'PDF');
  for (const name of ['Home', 'Annotate', 'Edit', 'Pages', 'View', 'Fill Form']) {
    await pdfTab(name);
    const text = await w.locator('.ribbon-tabs, .ribbon-body').allInnerTexts();
    assert.doesNotMatch(text.join(' '), /genspark|genoffice|\bAI\b/i, `${name} tab`);
  }
  assert.equal(await w.locator('.ai-entry, .ai-dock, .ai-rail, .pdf-sel-ask').count(), 0);
  assert.deepEqual(w.errors, []);
});

test('the Convert menu offers what Kertas can open: Word, Excel and pictures', pdfOpts, async () => {
  await open('sample.pdf', 'pdf');
  await pdfTab('Home');
  await w.getByText('PDF Converter', { exact: true }).click();
  await sleep(400);
  const items = (await w.locator('.rb-menu button').allInnerTexts()).map((s) => s.trim());
  assert.deepEqual(items, ['PDF to Word', 'PDF to Excel', 'Export as PNG', 'Export as JPG']);
  await w.keyboard.press('Escape');
});

test('rotating a page and pressing Ctrl+S saves the rotation into the file', pdfOpts, async () => {
  await open('sample.pdf', 'pdf');
  const was = (await PDFDocument.load(await readFile(file))).getPage(0).getRotation().angle;
  await pdfTab('Pages');
  await w.getByText('Rotate right', { exact: true }).click();
  await sleep(2500);
  if (await dirtyTabs()) { await press('pdf', 'S'); await sleep(3500); }
  assert.equal(await dirtyTabs(), 0);
  const doc = await PDFDocument.load(await readFile(file));
  assert.equal(doc.getPageCount(), 3);
  assert.equal((doc.getPage(0).getRotation().angle - was + 360) % 360, 90);
});

test('PDF to Word from the Convert menu opens the result as a document, and PDF to Excel as a workbook', pdfOpts, async () => {
  await open('report.pdf', 'pdf');
  await pdfTab('Home');
  const docsBefore = editors('docs').length;
  await w.getByText('PDF Converter', { exact: true }).click();
  await w.getByText('PDF to Word', { exact: true }).click();
  for (let i = 0; i < 160 && editors('docs').length === docsBefore; i++) await sleep(250);
  assert.equal(editors('docs').length, docsBefore + 1, 'a Writer tab opened with the converted document');
  const doc = editors('docs').at(-1);
  await doc.waitForSelector('.ProseMirror');
  assert.ok((await doc.locator('.ProseMirror').innerText()).trim().length > 20, 'it has the text of the PDF');

  if (!hasSheets) return;
  await answerDialogs([], 'report-converted.xlsx');
  const sheetsBefore = editors('sheets').length;
  await k.locator('#tabs .tab', { hasText: path.basename(file) }).click();
  await sleep(600);
  await w.getByText('PDF Converter', { exact: true }).click();
  await w.getByText('PDF to Excel', { exact: true }).click();
  for (let i = 0; i < 160 && editors('sheets').length === sheetsBefore; i++) await sleep(250);
  assert.equal(editors('sheets').length, sheetsBefore + 1, 'the workbook opened in the new Sheets');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(dir, 'report-converted.xlsx'));
  assert.ok(wb.worksheets.length >= 1);
});

test('Kertas\'s dark mode reaches both editors', pdfOpts, async () => {
  await open('sample.pdf', 'pdf');
  const theme = () => w.evaluate(() => document.documentElement.getAttribute('data-theme'));
  const was = await theme();
  await k.locator('#theme-toggle, [data-action="theme"], button[title*="theme" i], button[title*="dark" i], button[title*="light" i]').first().click();
  await sleep(900);
  assert.notEqual(await theme(), was);
  await k.locator('#theme-toggle, [data-action="theme"], button[title*="theme" i], button[title*="dark" i], button[title*="light" i]').first().click();
  await sleep(500);
});
