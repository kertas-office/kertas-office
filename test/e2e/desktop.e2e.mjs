// The desktop app, end to end: Electron with its real main process, file access guard and
// native dialogs (answered by the test), on a private settings folder.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, copyFile, readFile, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let app;
let win;
let dir;

before(async () => {
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-e2e-'));
  app = await _electron.launch({
    executablePath: require('electron'),
    args: [ROOT, '--no-sandbox'],
    env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production' },
  });
  win = await app.firstWindow();
  win.setDefaultTimeout(15000);
  win.errors = [];
  win.on('pageerror', (e) => win.errors.push(String(e.message || e)));
  win.on('pageerror', (e) => { if (process.env.E2E_STACKS) console.error('PAGE ERROR', e.stack); });
  await win.waitForSelector('#tabs');
});

after(async () => {
  try { await Promise.race([app?.close(), sleep(5000)]); } catch { /* closing */ }
  try { app?.process().kill('SIGKILL'); } catch { /* gone */ }
  await rm(dir, { recursive: true, force: true });
});

/** Answers the next native Open dialogs with these files, and Save dialogs with a path in the temp folder. */
async function answerDialogs(openPaths, saveDir = dir) {
  await app.evaluate(({ dialog }, [openPaths, saveDir]) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: openPaths });
    dialog.showSaveDialog = async (_w, o) => ({ canceled: false, filePath: `${saveDir}/${String(o.defaultPath || 'file').split(/[\\/]/).pop()}` });
  }, [openPaths, saveDir]);
}

/** Opens files through the app's own Open command (Ctrl+O), as a person would. */
async function openWithDialog(paths) {
  await answerDialogs(paths);
  const before = await win.locator('#tabs .tab').count();
  await win.keyboard.press('Control+o');
  await win.waitForFunction((n) => document.querySelectorAll('#tabs .tab').length > n, before);
  await sleep(2500);
}

const toHome = async () => { await win.locator('#tabs .tab').first().click(); await sleep(600); };

test('the app starts on Home', async () => {
  assert.ok(await win.locator('#tabs .tab').count() >= 1);
  assert.deepEqual(win.errors, []);
});

test('a document opened from disk is listed in Recent and can be removed from it', async () => {
  const file = path.join(dir, 'Letter.docx');
  await copyFile(fixture('sample.docx'), file);
  await openWithDialog([file]);
  await toHome();
  const row = win.locator('.recent-item', { hasText: 'Letter.docx' });
  await row.waitFor({ state: 'visible' });
  await row.hover();
  await row.locator('.rm-btn').click();
  await sleep(500);
  assert.equal(await win.locator('.recent-item', { hasText: 'Letter.docx' }).count(), 0);
});

test('Home lists at most 10 recent files', async () => {
  const files = [];
  for (let i = 1; i <= 12; i++) { const f = path.join(dir, `Note ${i}.docx`); await copyFile(fixture('sample.docx'), f); files.push(f); }
  await openWithDialog(files);
  await toHome();
  const n = await win.locator('.recent-item').count();
  assert.ok(n > 0 && n <= 10, `${n} recent files shown`);
});

test('Writer: a picture picked in Insert > Picture goes in (1.0.1 regression)', async () => {
  const doc = path.join(dir, 'Pictures.docx');
  await copyFile(fixture('sample.docx'), doc);
  await openWithDialog([doc]);
  const imgs = () => win.locator('.superdoc-page:visible img').count();
  const before = await imgs();
  await answerDialogs([fixture('logo.png')]);
  await win.locator('.rb-tab[data-tab="insert"]:visible').click();
  await win.locator('[data-id="picture"]:visible').first().click();
  await sleep(1500);
  assert.equal(await imgs(), before + 1, 'picture inserted');
  assert.deepEqual(win.errors, []);
});

test('Writer: saving writes the file in place, and Export PDF makes a PDF', async () => {
  const doc = path.join(dir, 'Pictures.docx');
  await win.keyboard.press('Control+s');
  await sleep(2500);
  const zip = await JSZip.loadAsync(await readFile(doc));
  assert.ok(Object.keys(zip.files).some((f) => f.startsWith('word/media/')), 'saved with the picture');

  const pdf = await win.evaluate(async () => {
    const t = [...window.__kertas.tabs.values()].find((x) => x.kind === 'writer' && x.name === 'Pictures.docx');
    const b = new Uint8Array(await t.instance.exportPdf());
    return String.fromCharCode(...b.subarray(0, 5));
  });
  assert.equal(pdf, '%PDF-');
});

test('PDF: an attached file is saved where the person chooses', async () => {
  const pdf = path.join(dir, 'Report.pdf');
  await copyFile(fixture('sample.pdf'), pdf);
  await openWithDialog([pdf]);
  await win.locator('[data-side="attach"]:visible').click();
  await answerDialogs([fixture('sample.xlsx')]);
  await win.locator('[data-att-add]:visible').click();
  await sleep(800);
  const row = win.locator('.pdf-attach:visible .att-row', { hasText: 'sample.xlsx' });
  await row.hover();
  const out = path.join(dir, 'saved');
  await (await import('node:fs/promises')).mkdir(out);
  await answerDialogs([], out);
  await row.locator('[data-att="save"]').click();
  await sleep(800);
  assert.deepEqual(await readFile(path.join(out, 'sample.xlsx')), await readFile(fixture('sample.xlsx')));
  assert.deepEqual(win.errors, []);
});

test('files the person never opened stay out of reach of the page', async () => {
  const secret = path.join(dir, 'secret.txt');
  await writeFile(secret, 'private');
  const res = await win.evaluate(async (p) => {
    try { await window.kertasNative.readPath(p); return 'read'; } catch (e) { return `blocked: ${e.message}`; }
  }, secret);
  assert.match(res, /^blocked/);
});

