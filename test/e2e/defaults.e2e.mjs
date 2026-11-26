// The new editors are what Kertas opens by default, and the commands that used to need the
// previous PDF editor (Merge PDFs, Images to PDF, Sign a PDF) work with the new one.
// Needs the editors built (node scripts/build-writer2.cjs); skipped when they are not there.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, copyFile, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { PDFDocument } from '@cantoo/pdf-lib';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const built = existsSync(path.join(ROOT, 'electron', 'writer2', 'pdf', 'renderer', 'index.html'));
const opts = { skip: built ? false : 'the new editors are not built (node scripts/build-writer2.cjs)' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let app; let k; let dir; let n = 0;

before(async () => {
  if (!built) return;
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-def-'));
  const env = { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production' };
  for (const v of ['KERTAS_NEW_WRITER', 'KERTAS_NEW_SHEETS', 'KERTAS_NEW_PDF']) delete env[v];
  app = await _electron.launch({ executablePath: require('electron'), args: [ROOT, '--no-sandbox', '--window-size=1400,1000'], env });
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
const copy = async (name) => { const f = path.join(dir, `${++n}-${name}`); await copyFile(fixture(name), f); return f; };
async function dialogs(openPaths, saveName) {
  await app.evaluate(({ dialog }, [openPaths, dir, saveName]) => {
    dialog.showOpenDialog = async () => ({ canceled: false, filePaths: openPaths });
    dialog.showSaveDialog = async () => ({ canceled: false, filePath: `${dir}/${saveName}` });
    dialog.showMessageBox = async (_w, o) => {
      const buttons = (o || _w)?.buttons || [];
      const dont = buttons.findIndex((b) => /don.t save|discard/i.test(b));
      return { response: dont >= 0 ? dont : 0, checkboxChecked: false };
    };
  }, [openPaths, dir, saveName]);
}
async function command(label) {
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+k');
  await k.keyboard.type(label);
  await k.keyboard.press('Enter');
}
async function waitFor(fn, ms = 20000) { for (let t = 0; t < ms; t += 250) { if (await fn()) return true; await sleep(250); } return false; }

test('without any setting, Word documents, workbooks and PDFs open in the new editors', opts, async () => {
  for (const [name, kind] of [['sample.docx', 'docs'], ['sample.xlsx', 'sheets'], ['sample.pdf', 'pdf']]) {
    const had = editors(kind).length;
    await dialogs([await copy(name)], 'unused');
    await command('Open...');
    assert.ok(await waitFor(async () => editors(kind).length > had), `${name} opened in the new editor`);
  }
});

test('Merge PDFs joins the files, saves the result and opens it in the new PDF editor', opts, async () => {
  const a = await copy('sample.pdf'); const b = await copy('report.pdf');
  const pages = async (f) => (await PDFDocument.load(await readFile(f))).getPageCount();
  const want = (await pages(a)) + (await pages(b));
  await dialogs([a, b], 'merged.pdf');
  const had = editors('pdf').length;
  await command('Merge PDFs');
  const out = path.join(dir, 'merged.pdf');
  assert.ok(await waitFor(async () => existsSync(out) && editors('pdf').length > had), 'saved and opened');
  assert.equal(await pages(out), want);
});

test('Images to PDF makes a PDF of the pictures, saves it and opens it in the new PDF editor', opts, async () => {
  await dialogs([await copy('photo.jpg'), await copy('logo.png')], 'pictures.pdf');
  const had = editors('pdf').length;
  await command('Images to PDF');
  // the page-size question of Kertas's own dialog
  await k.waitForSelector('.modal .chip.primary');
  await k.locator('.modal .chip.primary').click();
  const out = path.join(dir, 'pictures.pdf');
  assert.ok(await waitFor(async () => existsSync(out) && editors('pdf').length > had), 'saved and opened');
  assert.equal((await PDFDocument.load(await readFile(out))).getPageCount(), 2);
});

test('Sign a PDF opens the file in the new PDF editor with the signature step started', opts, async () => {
  await dialogs([await copy('sample.pdf')], 'unused');
  const had = editors('pdf').length;
  await command('Sign a PDF');
  assert.ok(await waitFor(async () => editors('pdf').length > had));
  const w = editors('pdf').at(-1);
  assert.ok(await waitFor(async () => (await w.locator('.pdf-modal, [role="dialog"]').count()) > 0), 'the signature dialog is open');
});
