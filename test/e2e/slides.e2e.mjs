// The Slides editor, end to end: presentations open in it as Kertas tabs, edits are saved into
// the .pptx, and there is nothing of an assistant. Checked in the saved files with JSZip and
// pdf-lib. Needs the editors built (node scripts/build-writer2.cjs); skipped when they are not.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, copyFile, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { PDFDocument } from '@cantoo/pdf-lib';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const W2 = path.join(ROOT, 'electron', 'writer2');
const hasSlides = existsSync(path.join(W2, 'slides', 'renderer', 'index.html'));
const opts = { skip: hasSlides ? false : 'the Slides editor is not built (node scripts/build-writer2.cjs)' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let app;
let k; // Kertas's own page
let w; // the Slides editor of the open deck
let dir;
let file;
let n = 0;

before(async () => {
  if (!hasSlides) return;
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-slides-'));
  app = await _electron.launch({
    executablePath: require('electron'),
    args: [ROOT, '--no-sandbox', '--window-size=1400,1000'],
    env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production' },
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

const editors = () => app.context().pages().filter((p) => p.url().startsWith('kertas-writer://slides/'));
const dirtyTabs = () => k.locator('#tabs .tab.is-dirty').count();
const tabTitles = () => k.locator('#tabs .tab .title').allTextContents();

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

async function waitEditor(had) {
  w = null;
  for (let i = 0; i < 80 && !w; i++) { w = editors().find((p) => !had.has(p)); if (!w) await sleep(250); }
  assert.ok(w, 'the deck opened in Slides');
  w.setDefaultTimeout(15000);
  w.errors = [];
  w.on('pageerror', (e) => w.errors.push(String(e.message || e)));
  await w.waitForSelector('.slide-list .thumb');
  await sleep(2500);
}

/** A fresh copy of a fixture, opened with Kertas's Open command */
async function open(name) {
  await closeAll();
  file = path.join(dir, `${++n}-${name}`);
  await copyFile(fixture(name), file);
  await answerDialogs([file]);
  const had = new Set(editors());
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+o');
  await waitEditor(had);
}

async function press(keyCode, modifiers = ['control']) {
  await app.evaluate(({ webContents }, [keyCode, modifiers]) => {
    const wc = webContents.getAllWebContents().filter((c) => c.getURL().startsWith('kertas-writer://slides/')).at(-1);
    wc.sendInputEvent({ type: 'keyDown', keyCode, modifiers });
    wc.sendInputEvent({ type: 'keyUp', keyCode, modifiers });
  }, [keyCode, modifiers]);
}

const partsText = async (bytes, re) => {
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter((f) => re.test(f)).sort();
  return (await Promise.all(names.map((f) => zip.file(f).async('string')))).join('\n');
};

test('the Slides build cannot reach an AI service: no service addresses, no assistant channels, no sign-in', opts, async () => {
  const files = [path.join(W2, 'slides', 'preload', 'index.js')];
  for (const name of await readdir(path.join(W2, 'slides', 'renderer', 'assets'))) if (/\.(c?js|mjs)$/.test(name)) files.push(path.join(W2, 'slides', 'renderer', 'assets', name));
  const marks = /api\.openai\.com|api\.anthropic\.com|generativelanguage\.googleapis|api\.deepseek\.com|genspark\.ai|chat\/completions|["'`](?:ai|kb|project):[a-zA-Z-]+["'`]|gsk-login|x-api-key/;
  const found = [];
  for (const f of files) {
    const text = await readFile(f, 'utf8');
    const m = marks.exec(text);
    if (m) found.push(`${path.relative(W2, f)}: ${text.slice(Math.max(0, m.index - 40), m.index + 60).replace(/\s+/g, ' ')}`);
  }
  assert.deepEqual(found, []);
});

test('a presentation opens in Slides as a Kertas tab, with Kertas\'s File button and no assistant', opts, async () => {
  await open('sample.pptx');
  assert.ok((await tabTitles()).includes(`${n}-sample.pptx`));
  assert.equal(await w.locator('.slide-list .thumb').count(), 3);
  assert.equal(await w.locator('.ai-dock, .stage-ai-bar').count(), 0);
  assert.doesNotMatch(await w.locator('body').innerText(), /Genspark|GenOffice|\bAI\b/);
  await w.locator('.kertas-file-tab').click();
  const items = (await w.locator('.kertas-file-menu button span').allInnerTexts()).map((s) => s.trim());
  for (const want of ['New presentation', 'Save', 'Export as PDF...', 'Export slides as PNG pictures...']) assert.ok(items.includes(want), want);
  await w.keyboard.press('Escape');
  assert.deepEqual(w.errors, []);
});

test('speaker notes typed and saved with Ctrl+S are in the .pptx', opts, async () => {
  await open('sample.pptx');
  await w.locator('textarea[placeholder*="speaker notes" i], textarea[placeholder*="notes" i]').first().click();
  await sleep(300);
  await w.keyboard.type('Thank the team');
  await w.locator('.slide-list .thumb').first().click(); // notes are kept when the box loses focus
  await sleep(1500);
  assert.equal(await dirtyTabs(), 1, 'the tab shows unsaved changes');
  await press('S');
  for (let i = 0; i < 40 && await dirtyTabs(); i++) await sleep(250);
  assert.equal(await dirtyTabs(), 0, 'saved');
  assert.match(await partsText(await readFile(file), /^ppt\/notesSlides\/.*\.xml$/), /Thank the team/);
  assert.match(await partsText(await readFile(file), /^ppt\/slides\/slide\d+\.xml$/), /Quarterly review/);
});

test('a new slide added with Ctrl+M is saved as a fourth slide', opts, async () => {
  await open('sample.pptx');
  await w.locator('.slide-list .thumb').nth(2).click();
  await press('M');
  await sleep(1200);
  assert.equal(await w.locator('.slide-list .thumb').count(), 4);
  await press('S');
  for (let i = 0; i < 40 && await dirtyTabs(); i++) await sleep(250);
  const zip = await JSZip.loadAsync(await readFile(file));
  assert.equal(Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f)).length, 4);
});

test('Export as PDF from the File button writes one page per slide', opts, async () => {
  await open('sample.pptx');
  await answerDialogs([], 'deck.pdf');
  await w.locator('.kertas-file-tab').click();
  await w.locator('.kertas-file-menu button').filter({ has: w.locator('span', { hasText: /^Export as PDF/ }) }).click();
  const out = path.join(dir, 'deck.pdf');
  for (let i = 0; i < 80 && !existsSync(out); i++) await sleep(250);
  await sleep(800);
  const pdf = await PDFDocument.load(await readFile(out));
  assert.equal(pdf.getPageCount(), 3);
});

test('New presentation on Home opens a blank deck, and its first Save asks where to keep it', opts, async () => {
  await closeAll();
  await k.locator('#tabs .tab').first().click();
  const had = new Set(editors());
  await k.locator('.app-card.slides [data-cmd="new-slides"]').click();
  await waitEditor(had);
  assert.match((await tabTitles()).join(','), /Presentation\d+\.pptx/);
  await w.locator('textarea[placeholder*="speaker notes" i], textarea[placeholder*="notes" i]').first().click();
  await w.keyboard.type('A new deck');
  await w.locator('.slide-list .thumb').first().click();
  await sleep(1200);
  await answerDialogs([], 'new-deck.pptx');
  await press('S');
  const out = path.join(dir, 'new-deck.pptx');
  for (let i = 0; i < 60 && !existsSync(out); i++) await sleep(250);
  await sleep(800);
  assert.match(await partsText(await readFile(out), /^ppt\/notesSlides\/.*\.xml$/), /A new deck/);
  for (let i = 0; i < 20 && !(await tabTitles()).includes('new-deck.pptx'); i++) await sleep(250);
  assert.ok((await tabTitles()).includes('new-deck.pptx'), 'the tab takes the saved name');
});

test('the slide show covers the whole window and Escape brings the editor back', opts, async () => {
  await open('sample.pptx');
  const before = await w.evaluate(() => [innerWidth, innerHeight]);
  await press('F5', []);
  let during = before;
  for (let i = 0; i < 20; i++) { await sleep(250); during = await w.evaluate(() => [innerWidth, innerHeight]); if (during[1] > before[1]) break; }
  assert.ok(during[1] > before[1], `the show is taller than the editor (${before} -> ${during})`);
  await press('Escape', []);
  let after = during;
  for (let i = 0; i < 20; i++) { await sleep(250); after = await w.evaluate(() => [innerWidth, innerHeight]); if (after[1] === before[1]) break; }
  assert.deepEqual(after, before);
});

test('Kertas\'s dark mode reaches Slides', opts, async () => {
  await open('sample.pptx');
  const theme = () => w.evaluate(() => document.documentElement.getAttribute('data-theme'));
  const toggle = '#theme-toggle, [data-action="theme"], button[title*="theme" i], button[title*="dark" i], button[title*="light" i]';
  const was = await theme();
  await k.locator(toggle).first().click();
  await sleep(900);
  assert.notEqual(await theme(), was);
  await k.locator(toggle).first().click();
  await sleep(500);
});

test('a slide copied and pasted with Ctrl+C and Ctrl+V becomes a new slide (the clipboard works)', opts, async () => {
  await open('sample.pptx');
  await w.locator('.slide-list .thumb').first().click();
  await sleep(300);
  await press('C');
  await sleep(600);
  await press('V');
  for (let i = 0; i < 20 && await w.locator('.slide-list .thumb').count() < 4; i++) await sleep(250);
  assert.equal(await w.locator('.slide-list .thumb').count(), 4);
  assert.deepEqual(w.errors, []);
});

test('a template from the Slides card opens as a new deck, and its first Save asks where to keep it', opts, async () => {
  await closeAll();
  await k.locator('#tabs .tab').first().click();
  await k.locator('.app-card.slides [data-cmd="templates-slides"]').click();
  assert.equal(await k.locator('.tpl-gallery.slides [data-tpl]').count(), 6);
  const had = new Set(editors());
  await k.locator('.tpl-gallery.slides [data-tpl="pitch-deck"]').click();
  await waitEditor(had);
  assert.equal(await w.locator('.slide-list .thumb').count(), 7);
  assert.ok((await tabTitles()).some((t) => /^Pitch deck \d+\.pptx$/.test(t)), (await tabTitles()).join(','));
  await answerDialogs([], 'my-pitch.pptx');
  await w.locator('textarea[placeholder*="notes" i]').first().click();
  await w.keyboard.type('Open with the shop story');
  await w.locator('.slide-list .thumb').nth(1).click();
  await sleep(1000);
  await press('S');
  const out = path.join(dir, 'my-pitch.pptx');
  for (let i = 0; i < 60 && !existsSync(out); i++) await sleep(250);
  await sleep(800);
  const slides = await partsText(await readFile(out), /^ppt\/slides\/slide\d+\.xml$/);
  assert.match(slides, /The simplest way for small/);
  assert.match(await partsText(await readFile(out), /^ppt\/notesSlides\/.*\.xml$/), /Open with the shop story/);
  // the template itself is untouched
  const zip = await JSZip.loadAsync(await readFile(path.join(ROOT, 'src', 'templates', 'pitch-deck.pptx')));
  assert.equal(Object.keys(zip.files).some((f) => f.startsWith('ppt/notesSlides/')), false);
});
