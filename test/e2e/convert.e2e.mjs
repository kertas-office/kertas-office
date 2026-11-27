// The Convert app, end to end: the desktop app with files added the way a drop adds them,
// converted to each format and checked with independent readers. Word to PDF needs the
// new Writer built (node scripts/build-writer2.cjs); audio and video need an FFmpeg on PATH.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, copyFile, readFile, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import { PDFDocument } from '@cantoo/pdf-lib';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const hasWriter = existsSync(path.join(ROOT, 'electron', 'writer2', 'renderer', 'index.html'));
let ffmpeg = null;
try { ffmpeg = execFileSync(process.platform === 'win32' ? 'where' : 'which', ['ffmpeg']).toString().split(/\r?\n/)[0].trim() || null; } catch { /* none */ }

let app;
let k;
let dir;

async function launch(env = {}) {
  if (app) { try { await Promise.race([app.close(), sleep(4000)]); } catch { /* closing */ } }
  app = await _electron.launch({
    executablePath: require('electron'),
    args: [ROOT, '--no-sandbox', '--window-size=1300,900'],
    env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, `config-${Date.now()}`), NODE_ENV: 'production', ...env },
  });
  k = await app.firstWindow();
  k.setDefaultTimeout(20000);
  await k.waitForSelector('#tabs');
}

before(async () => {
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-convert-'));
  await launch(ffmpeg ? { KERTAS_FFMPEG: ffmpeg } : {});
});
after(async () => {
  const child = app?.process();
  try { await Promise.race([app?.close(), sleep(5000)]); } catch { /* closing */ }
  try { child?.kill('SIGKILL'); } catch { /* gone */ }
  await rm(dir, { recursive: true, force: true }).catch(() => {});
  setTimeout(() => process.exit(), 3000).unref();
});

/** A fresh folder with copies of these fixtures, added to the Convert list */
async function addFiles(names) {
  const sub = await mkdtemp(path.join(dir, 'job-'));
  const paths = [];
  for (const n of names) { const p = path.join(sub, path.basename(n)); await copyFile(fixture(n), p); paths.push(p); }
  await k.evaluate(() => window.__kertas.openConvert());
  await k.waitForSelector('.cv');
  await k.evaluate(() => { const t = [...window.__kertas.tabs.values()].find((x) => x.kind === 'convert'); t.instance && document.querySelector('[data-act="clear"]')?.click(); });
  await k.evaluate(async (list) => {
    const t = [...window.__kertas.tabs.values()].find((x) => x.kind === 'convert');
    await t.instance.add(await window.kertasNative.convert('add', { paths: list }));
  }, paths);
  return sub;
}
const rowOf = (name) => k.locator('.cv-row', { has: k.locator('strong', { hasText: new RegExp(`^${name.replace(/\./g, '\\.')}$`) }) });
async function convert(name, to) {
  const row = rowOf(name);
  await row.locator('select').selectOption(to);
  const id = await row.getAttribute('data-id');
  const ok = await k.evaluate((i) => [...window.__kertas.tabs.values()].find((x) => x.kind === 'convert').instance.convertOne(i), id);
  const status = (await rowOf(name).locator('.cv-status').textContent()).trim();
  assert.ok(ok, `${name} to ${to}: ${status}`);
}
const magic = async (file) => (await readFile(file)).subarray(0, 12);

test('Home has a Convert card that opens the Convert tab', async () => {
  await k.evaluate(() => window.__kertas.activate('home'));
  await k.locator('.app-card.convert [data-convert="open"]').click();
  await k.waitForSelector('.cv-drop');
  assert.equal(await k.locator('#tabs .tab', { hasText: 'Convert' }).count(), 1);
  // a second time brings the same tab forward
  await k.evaluate(() => window.__kertas.openConvert());
  assert.equal(await k.locator('#tabs .tab', { hasText: 'Convert' }).count(), 1);
});

test('pictures convert between formats, including HEIC, TIFF and SVG in, and AVIF, ICO and PDF out', async () => {
  const sub = await addFiles(['photo.jpg', 'convert/sample.heic', 'convert/sample.tiff', 'convert/sample.svg']);
  await convert('photo.jpg', 'webp');
  await convert('photo.jpg', 'avif');
  await convert('sample.heic', 'jpg');
  await convert('sample.tiff', 'png');
  await convert('sample.svg', 'ico');
  await convert('photo.jpg', 'pdf');
  const m = (f) => magic(path.join(sub, f));
  assert.equal((await m('photo.webp')).subarray(8, 12).toString(), 'WEBP');
  assert.equal((await m('photo.avif')).subarray(4, 12).toString(), 'ftypavif');
  assert.deepEqual([...(await m('sample.jpg')).subarray(0, 3)], [0xff, 0xd8, 0xff]);
  assert.equal((await m('sample.png')).subarray(1, 4).toString(), 'PNG');
  assert.deepEqual([...(await m('sample.ico')).subarray(0, 4)], [0, 0, 1, 0]);
  assert.equal((await PDFDocument.load(await readFile(path.join(sub, 'photo.pdf')))).getPageCount(), 1);
});

test('a PDF becomes Word, Excel, pictures per page, text, a smaller copy and one file per page', async () => {
  const sub = await addFiles(['report.pdf', 'sample.pdf']);
  await convert('report.pdf', 'docx');
  await convert('report.pdf', 'xlsx');
  await convert('report.pdf', 'txt');
  await convert('sample.pdf', 'png');
  await convert('sample.pdf', 'pages');
  await convert('report.pdf', 'compressed');
  const files = await readdir(sub);
  const docx = await JSZip.loadAsync(await readFile(path.join(sub, 'report.docx')));
  assert.ok((await docx.file('word/document.xml').async('string')).length > 500);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(sub, 'report.xlsx'));
  assert.ok(wb.worksheets.length >= 1);
  assert.ok((await readFile(path.join(sub, 'report.txt'), 'utf8')).trim().length > 50);
  const pages = (await PDFDocument.load(await readFile(fixture('sample.pdf')))).getPageCount();
  assert.equal(files.filter((f) => /^sample - page \d+\.png$/.test(f)).length, pages);
  assert.equal(files.filter((f) => /^sample - page \d+\.pdf$/.test(f)).length, pages);
  assert.ok(files.includes('report (compressed).pdf'));
});

test('a scan with recognised text converts to Word as paragraphs in one size', async () => {
  const sub = await addFiles(['convert/scan-recognised.pdf']);
  await convert('scan-recognised.pdf', 'docx');
  const xml = await (await JSZip.loadAsync(await readFile(path.join(sub, 'scan-recognised.docx')))).file('word/document.xml').async('string');
  assert.match(xml, /Lorem ipsum/);
  const sizes = new Set([...xml.matchAll(/<w:sz w:val="(\d+)"/g)].map((m) => m[1]));
  assert.ok(sizes.size <= 2, `one body size (and perhaps a heading): ${[...sizes].join(', ')}`);
  assert.ok([...sizes].some((v) => Number(v) >= 20), 'the size comes from the line spacing, not the 8 pt it was stored at');
});

test('documents: Markdown, HTML and text to Word and PDF, and Word to text, HTML and Markdown', async () => {
  const sub = await addFiles(['convert/notes.md', 'convert/page.html', 'convert/plain.txt', 'sample.docx']);
  await convert('notes.md', 'docx');
  await convert('page.html', 'pdf');
  await convert('plain.txt', 'docx');
  await convert('sample.docx', 'txt');
  await convert('sample.docx', 'md');
  await convert('sample.docx', 'html');
  const xml = await (await JSZip.loadAsync(await readFile(path.join(sub, 'notes.docx')))).file('word/document.xml').async('string');
  assert.match(xml, /Trip notes/);
  assert.match(xml, /w:pStyle w:val="Heading1"/, 'the heading is a Word heading');
  assert.match(xml, /<w:tbl>/, 'the table is a Word table');
  assert.match(xml, /w:numPr/, 'the lists are Word lists');
  assert.match(xml, /Passport/);
  assert.equal((await PDFDocument.load(await readFile(path.join(sub, 'page.pdf')))).getPageCount(), 1);
  assert.match(await (await JSZip.loadAsync(await readFile(path.join(sub, 'plain.docx')))).file('word/document.xml').async('string'), /café/);
  assert.ok((await readFile(path.join(sub, 'sample.txt'), 'utf8')).trim().length > 10);
  assert.ok((await readFile(path.join(sub, 'sample.md'), 'utf8')).trim().length > 10);
  assert.match(await readFile(path.join(sub, 'sample.html'), 'utf8'), /<html/);
});

test('Word to PDF goes through the new Writer, out of sight', { skip: hasWriter ? false : 'the new Writer is not built' }, async () => {
  const sub = await addFiles(['long.docx']);
  await convert('long.docx', 'pdf');
  const doc = await PDFDocument.load(await readFile(path.join(sub, 'long.pdf')));
  assert.ok(doc.getPageCount() > 1, 'the document keeps its pages');
  // no Writer tab or window was left open
  assert.equal(app.windows().length, 1);
});

test('spreadsheets: CSV to Excel, Excel to CSV per sheet, JSON and PDF', async () => {
  const sub = await addFiles(['convert/people.csv', 'sample.xlsx', 'convert/people.json']);
  await convert('people.csv', 'xlsx');
  await convert('sample.xlsx', 'csv');
  await convert('sample.xlsx', 'json');
  await convert('sample.xlsx', 'pdf');
  await convert('people.json', 'csv');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(sub, 'people.xlsx'));
  assert.equal(wb.worksheets[0].getCell('C4').value, 'Paris');
  const files = await readdir(sub);
  const csvs = files.filter((f) => /^sample - .+\.csv$/.test(f));
  assert.ok(csvs.length >= 2, `one CSV per sheet: ${files.join(', ')}`);
  assert.ok(typeof JSON.parse(await readFile(path.join(sub, 'sample.json'), 'utf8')) === 'object');
  assert.ok((await PDFDocument.load(await readFile(path.join(sub, 'sample.pdf')))).getPageCount() >= 1);
  assert.match(await readFile(path.join(sub, 'people (1).csv'), 'utf8'), /Aisha/);
});

test('several PDFs merge into one, and several pictures go into one PDF', async () => {
  const sub = await addFiles(['report.pdf', 'sample.pdf', 'photo.jpg', 'convert/sample.webp']);
  await k.locator('[data-act="merge"]').click();
  await k.locator('[data-act="combine"]').click();
  for (let i = 0; i < 40 && (await readdir(sub)).filter((f) => /\((merged|pictures)\)\.pdf$/.test(f)).length < 2; i++) await sleep(250);
  const count = async (f) => (await PDFDocument.load(await readFile(path.join(sub, f)))).getPageCount();
  const pages = async (f) => (await PDFDocument.load(await readFile(fixture(f)))).getPageCount();
  assert.equal(await count('report (merged).pdf'), (await pages('report.pdf')) + (await pages('sample.pdf')));
  assert.equal(await count('photo (pictures).pdf'), 2);
});

test('converting never overwrites: a second run gets a new name', async () => {
  const sub = await addFiles(['convert/people.csv']);
  await convert('people.csv', 'tsv');
  await k.evaluate(() => { const r = document.querySelector('.cv-row select'); r.value = 'json'; r.dispatchEvent(new Event('change', { bubbles: true })); });
  await convert('people.csv', 'tsv');
  const files = await readdir(sub);
  assert.ok(files.includes('people.tsv') && files.includes('people (1).tsv'), files.join(', '));
});

test('without LibreOffice, .odt and .rtf are not offered, and older Word files still convert', async () => {
  await launch({ KERTAS_NO_LIBREOFFICE: '1' });
  const sub = await addFiles(['sample.docx', 'legacy.doc']);
  const offered = await rowOf('sample.docx').locator('option').evaluateAll((os) => os.map((o) => o.value));
  assert.ok(!offered.includes('odt') && !offered.includes('rtf'), offered.join(','));
  await convert('legacy.doc', 'docx');
  const xml = await (await JSZip.loadAsync(await readFile(path.join(sub, 'legacy.docx')))).file('word/document.xml').async('string');
  assert.ok(xml.length > 200);
});

test('audio and video convert with FFmpeg', { skip: ffmpeg ? false : 'no ffmpeg on PATH' }, async () => {
  await launch({ KERTAS_FFMPEG: ffmpeg });
  const sub = await addFiles(['convert/clip.mp4', 'convert/tone.wav']);
  assert.ok(await k.locator('.cv-engine').isHidden(), 'the engine is ready, so nothing asks to download it');
  await convert('clip.mp4', 'mp3');
  await convert('clip.mp4', 'webm');
  await convert('clip.mp4', 'gif');
  await convert('tone.wav', 'flac');
  const m = (f) => magic(path.join(sub, f));
  assert.match((await m('clip.mp3')).subarray(0, 3).toString('latin1'), /^(ID3|\xff\xfb|\xff\xf3)/);
  assert.deepEqual([...(await m('clip.webm')).subarray(0, 4)], [0x1a, 0x45, 0xdf, 0xa3]);
  assert.equal((await m('clip.gif')).subarray(0, 3).toString(), 'GIF');
  assert.equal((await m('tone.flac')).subarray(0, 4).toString(), 'fLaC');
});
