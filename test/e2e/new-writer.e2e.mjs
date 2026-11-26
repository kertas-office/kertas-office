// The new Writer (trial), end to end: the desktop app with Word documents opening in the
// new editor, driven the way a person would and checked in the file that gets saved.
// Needs the editor built (node scripts/build-writer2.cjs); skipped when it is not there.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdtemp, copyFile, readFile, readdir, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { _electron } from 'playwright';
import { ROOT, fixture } from './harness.mjs';

const require = createRequire(import.meta.url);
const built = existsSync(path.join(ROOT, 'electron', 'writer2', 'main', 'index.cjs'));
const opts = { skip: built ? false : 'the new Writer is not built (node scripts/build-writer2.cjs)' };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let app;
let k; // Kertas's own page: tab strip, Home, palette
let w; // the editor of the document that is open
let dir;
let file; // the copy of the fixture being edited
let n = 0;

before(async () => {
  if (!built) return;
  dir = await mkdtemp(path.join(tmpdir(), 'kertas-nw-'));
  app = await _electron.launch({
    executablePath: require('electron'),
    args: [ROOT, '--no-sandbox', '--window-size=1400,1000'],
    env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production', KERTAS_NEW_WRITER: '1' },
  });
  k = await app.firstWindow();
  k.setDefaultTimeout(15000);
  await k.waitForSelector('#tabs');
});

after(async () => {
  if (!built) return;
  // a failed test can leave a native prompt waiting, so closing is not waited for long
  const child = app?.process();
  try { await Promise.race([app?.close(), sleep(5000)]); } catch { /* closing */ }
  try { child?.kill('SIGKILL'); } catch { /* gone */ }
  await rm(dir, { recursive: true, force: true }).catch(() => {});
  // The connections to the editor views that went away with their tabs can keep this process
  // alive after everything is reported; end it, with the exit code the results gave it.
  setTimeout(() => process.exit(), 3000).unref();
});

const editors = () => app.context().pages().filter((p) => p.url().startsWith('kertas-writer://'));

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

/** A fresh copy of a fixture, opened through Kertas's own Open command; `w` becomes its editor. */
async function open(name) {
  // one document at a time: close what the last test left open
  while (await k.locator('#tabs .tab .close').count()) {
    await answerDialogs();
    await k.locator('#tabs .tab .close').first().click();
    await sleep(700);
  }
  file = path.join(dir, `${++n}-${name}`);
  await copyFile(fixture(name), file);
  await answerDialogs([file]);
  const had = new Set(editors());
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+o');
  for (let i = 0; i < 60; i++) {
    w = editors().find((p) => !had.has(p));
    if (w) break;
    await sleep(250);
  }
  assert.ok(w, `${name} opened in the new Writer`);
  w.setDefaultTimeout(15000);
  w.errors = [];
  w.on('pageerror', (e) => w.errors.push(String(e.message || e)));
  await w.waitForSelector('.ProseMirror');
  await sleep(1800);
  await inDoc();
}

const zipOf = async (p = file) => JSZip.loadAsync(await readFile(p));
const xml = async (part = 'word/document.xml') => { const f = (await zipOf()).file(part); return f ? f.async('string') : null; };
const partNames = async () => Object.keys((await zipOf()).files);
const save = async () => { await w.keyboard.press('Control+s'); await sleep(2200); };
/** the cursor in the body text: the first paragraph, clear of tables */
const inDoc = async () => { const b = await w.locator('.ProseMirror > p').first().boundingBox(); await w.mouse.click(b.x + Math.min(b.width / 2, 200), b.y + b.height / 2); await sleep(200); };
/** a new empty paragraph at the end of the document */
const toEnd = async () => { await inDoc(); await w.keyboard.press('Control+End'); await sleep(250); await w.keyboard.press('Enter'); await sleep(100); };
/** a command from the slash menu's own search box */
const slash = async (query) => { await w.keyboard.press('Control+/'); await sleep(350); await w.keyboard.type(query); await sleep(300); await w.keyboard.press('Enter'); await sleep(800); };
/** a ribbon tab, without collapsing the ribbon by clicking the open one */
async function tab(name) {
  const t = w.locator('.ribbon-tab', { hasText: new RegExp(`^${name}$`) });
  if (!/\bactive\b/.test(await t.getAttribute('class')) || !(await w.locator('.ribbon-group:visible').count())) await t.click();
  await sleep(450);
}
/** a ribbon control by the start of its tooltip */
const tip = (start) => w.locator(`.ribbon [data-tip^="${start}"]:visible`).first();
const menuItem = (text) => w.locator('button:visible, [role="menuitem"]:visible', { hasText: new RegExp(`^${text}$`) }).last();
const dialogButton = (text) => w.locator('.modal button:visible, [role="dialog"] button:visible', { hasText: new RegExp(`^${text}$`) }).last();
const dialogOk = async () => { await w.locator('.modal button.btn-primary:visible, [role="dialog"] button.btn-primary:visible').last().click(); await sleep(900); };
/** the paragraphs of a document's XML that hold this text */
const paragraphsWith = (x, text) => (x.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).filter((para) => para.replace(/<[^>]+>/g, '').includes(text));
const kertasTabs = () => k.evaluate(() => [...document.querySelectorAll('#tabs .tab')].slice(1).map((t) => ({ title: t.querySelector('.title')?.textContent, dirty: t.classList.contains('is-dirty'), active: t.classList.contains('active') })));

test('the build has no AI assistant in it: no provider names, no assistant channels, no prompts', opts, async () => {
  const root = path.join(ROOT, 'electron', 'writer2');
  const files = [];
  for (const sub of ['main', 'preload', path.join('renderer', 'assets')]) {
    for (const name of await readdir(path.join(root, sub))) if (/\.(c?js|mjs)$/.test(name)) files.push(path.join(root, sub, name));
  }
  // what only the assistant has: the services it called, its channels to the main process, its sign-in
  const marks = /api\.openai\.com|api\.anthropic\.com|generativelanguage\.googleapis|api\.deepseek\.com|chat\/completions|["'`](?:ai|kb|project):[a-z-]+["'`]|aiStream(?!s:)|aiChat|gsk-login|x-api-key/;
  const found = [];
  for (const f of files) {
    const text = await readFile(f, 'utf8');
    const m = marks.exec(text);
    if (m) found.push(`${path.relative(root, f)}: ${text.slice(Math.max(0, m.index - 40), m.index + 60).replace(/\s+/g, ' ')}`);
  }
  assert.deepEqual(found, []);
  const bridge = await readFile(path.join(root, 'preload', 'index.js'), 'utf8');
  assert.doesNotMatch(bridge, /aiNamingApi|getAiSettings|kbSearch|onMcpCommand/);
});

test('a Word document opens as a Kertas tab; typing marks it unsaved and saving clears the mark', opts, async () => {
  await open('sample.docx');
  let tabs = await kertasTabs();
  assert.equal(tabs.length, 1);
  assert.match(tabs[0].title, /sample\.docx$/);
  assert.equal(tabs[0].dirty, false);
  await toEnd();
  await w.keyboard.type('Typed in the new Writer.');
  await sleep(2300);
  assert.equal((await kertasTabs())[0].dirty, true, 'the tab shows unsaved changes');
  await save();
  await sleep(1800);
  assert.equal((await kertasTabs())[0].dirty, false, 'saved');
  assert.match(await xml(), /Typed in the new Writer\./);
  assert.deepEqual(w.errors, []);
});

test('Save from Kertas (the command palette) saves the open document', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await w.keyboard.type('Saved by Kertas.');
  await k.locator('#btn-palette').click();
  await sleep(500);
  await k.keyboard.type('Save');
  await sleep(300);
  await k.keyboard.press('Enter');
  await sleep(2500);
  assert.match(await xml(), /Saved by Kertas\./);
});

test('rulers: shown by default; the top margin and the header distance are dragged on the side ruler', opts, async () => {
  await open('sample.docx');
  assert.ok(await w.locator('.ruler').count(), 'top ruler');
  assert.ok(await w.locator('.vruler').count(), 'side ruler');
  const pgMar = async () => Object.fromEntries([...(/<w:pgMar([^>]*)>/.exec(await xml())?.[1] ?? '').matchAll(/w:(\w+)="(-?\d+)"/g)].map((m) => [m[1], Number(m[2])]));
  const was = await pgMar();
  const drag = async (sel, dy) => {
    const b = await w.locator(sel).boundingBox();
    const x = b.x + b.width / 2;
    const y = b.y + b.height / 2;
    await w.mouse.move(x, y);
    await w.mouse.down();
    await w.mouse.move(x, y + dy / 2, { steps: 4 });
    await w.mouse.move(x, y + dy, { steps: 4 });
    await w.mouse.up();
    await sleep(500);
  };
  await drag('.vruler-edge-handle[data-side="top"]', -20);
  await drag('.vruler-margin-handle[data-side="top"]', -30);
  await save();
  const now = await pgMar();
  assert.ok(now.header < (was.header ?? 720), `header distance ${was.header} -> ${now.header}`);
  assert.ok(now.top < was.top, `top margin ${was.top} -> ${now.top}`);
});

test('form fields: a text field, a drop-down, a date and a tick box are filled in and saved as Word controls', opts, async () => {
  await open('form-fields.docx');
  const fields = () => w.evaluate(() => [...document.querySelectorAll('[data-ctrl-field], [data-ctrl-checkbox]')].map((e) => `${e.dataset.ctrlField ?? 'checkbox'}:${e.textContent}`));
  assert.deepEqual(await fields(), ['text:Click here to type', 'list:Choose an item', 'date:Pick a date', 'text:Kuala Lumpur', 'checkbox:☐']);
  await w.locator('[data-ctrl-field="text"]').first().click();
  await w.keyboard.type('Aminah Yusof');
  await w.locator('[data-ctrl-field="list"]').first().click();
  await w.locator('.kertas-field-menu button', { hasText: 'Engineering' }).click();
  await w.locator('[data-ctrl-field="date"]').first().click();
  await w.locator('.kertas-field-date input').fill('2026-12-01');
  await w.locator('.kertas-field-date button.is-primary').click();
  await w.locator('[data-ctrl-checkbox]').first().click();
  await sleep(300);
  assert.deepEqual(await fields(), ['text:Aminah Yusof', 'list:Engineering', 'date:1 December 2026', 'text:Kuala Lumpur', 'checkbox:☒']);
  await save();
  const x = await xml();
  assert.match(x, /<w:id w:val="101"\/><w:text\/><\/w:sdtPr><w:sdtContent>.*?Aminah Yusof/s, 'text field keeps its control, without the placeholder flag');
  assert.match(x, /<w:dropDownList>.*?<\/w:dropDownList><\/w:sdtPr><w:sdtContent>.*?Engineering/s);
  assert.match(x, /<w:date w:fullDate="2026-12-01T00:00:00Z">/);
  assert.match(x, /<w14:checked w14:val="1"\/>/);
  assert.equal((x.match(/<w:sdt>/g) || []).length, 5, 'all five controls are still controls');
});

test('form fields: all four kinds are inserted from the Insert tab', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await tab('Insert');
  await tip('Text field').click();
  await sleep(300);
  await w.keyboard.type(' then ');
  await tip('Tick box').click();
  await sleep(300);
  await w.keyboard.type(' then ');
  await tip('Dropdown list').click();
  await sleep(400);
  await w.locator('.kertas-field-dialog textarea').fill('Red\nGreen\nBlue');
  await w.locator('.kertas-field-dialog button.is-primary').click();
  await sleep(400);
  await w.keyboard.type(' then ');
  await tip('Date picker').click();
  await sleep(400);
  await w.locator('.kertas-field-dialog button.is-primary').click();
  await sleep(400);
  await save();
  const x = await xml();
  assert.match(x, /<w:showingPlcHdr\/><w:text\/>/);
  assert.match(x, /<w14:checkbox>/);
  assert.match(x, /<w:listItem w:displayText="Green" w:value="Green"\/>/);
  assert.match(x, /<w:dateFormat w:val="d MMMM yyyy"\/>/);
  assert.match(x, /xmlns:w14=/, 'the namespace the tick box needs is declared');
});

test('slash menu: / at the start of a line opens it, typing filters, Enter applies; not inside a date', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await w.keyboard.type('/');
  await sleep(500);
  assert.ok(await w.locator('.kertas-slash').count(), 'opened');
  await w.keyboard.type('bul');
  await sleep(300);
  assert.equal((await w.locator('.kertas-slash-item strong').allTextContents())[0], 'Bulleted list', 'best match first');
  await w.keyboard.press('Enter');
  await sleep(500);
  await w.keyboard.type('first item');
  await w.keyboard.press('Enter');
  await w.keyboard.press('Enter');
  await w.keyboard.type('Due 12/01');
  await sleep(400);
  assert.equal(await w.locator('.kertas-slash').count(), 0, 'a slash inside text is just a slash');
  await save();
  const x = await xml();
  assert.ok(paragraphsWith(x, 'first item').some((para) => /<w:numPr>|w:val="List/.test(para)), 'a list item');
  assert.match(x, /Due 12\/01/);
  assert.doesNotMatch(x, /\/bul/);
});

test('headings and a table of contents built from them', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await slash('heading 1');
  await w.keyboard.type('First chapter');
  await w.keyboard.press('Enter');
  await w.keyboard.type('Body under it.');
  await w.keyboard.press('Enter');
  await slash('heading 2');
  await w.keyboard.type('A section');
  await w.keyboard.press('Control+Home');
  await sleep(250);
  await w.keyboard.press('Enter');
  await w.keyboard.press('ArrowUp');
  await sleep(150);
  await slash('text');
  await slash('table of contents');
  await sleep(900);
  await save();
  const x = await xml();
  assert.ok(paragraphsWith(x, 'First chapter').some((para) => /<w:pStyle w:val="Heading1"\/>/.test(para)), 'Heading 1');
  assert.ok(paragraphsWith(x, 'A section').some((para) => /<w:pStyle w:val="Heading2"\/>/.test(para)), 'Heading 2');
  assert.match(x, /<w:instrText[^>]*>\s*TOC /, 'a real TOC field');
  assert.match(x, /TOC[\s\S]*First chapter[\s\S]*A section/, 'listing the headings');
});

test('footnote, endnote and a table', opts, async () => {
  await open('sample.docx');
  await inDoc();
  await w.keyboard.press('End');
  await slash('footnote');
  await w.locator('.modal textarea:visible').fill('A footnote typed here.');
  await dialogOk();
  await inDoc();
  await w.keyboard.press('End');
  await slash('endnote');
  await w.locator('.modal textarea:visible').fill('An endnote typed here.');
  await dialogOk();
  await toEnd();
  await slash('table');
  await dialogOk();
  await w.keyboard.type('cell one');
  await w.keyboard.press('Tab');
  await w.keyboard.type('cell two');
  await save();
  const x = await xml();
  assert.match(x, /<w:footnoteReference\b/);
  assert.match(await xml('word/footnotes.xml'), /A footnote typed here\./);
  assert.match(x, /<w:endnoteReference\b/);
  assert.match(await xml('word/endnotes.xml'), /An endnote typed here\./);
  assert.match(x, /<w:tbl>[\s\S]*cell one[\s\S]*cell two[\s\S]*<\/w:tbl>/);
});

test('bookmark, a cross-reference to it, a caption and a comment', opts, async () => {
  await open('sample.docx');
  const b = await w.locator('.ProseMirror p').first().boundingBox();
  await w.mouse.dblclick(b.x + 40, b.y + b.height / 2);
  await tab('Insert');
  await tip('Bookmark').click();
  await sleep(600);
  await w.locator('[placeholder="Bookmark name"]').fill('KeyPoint');
  await dialogButton('Add').click();
  await sleep(500);
  await dialogButton('Close').click();
  await toEnd();
  await w.keyboard.type('See ');
  await tip('Cross-reference').click();
  await sleep(700);
  await w.locator('.modal select:visible').first().selectOption('bookmark');
  await sleep(400);
  await w.locator('.modal:visible').getByText('KeyPoint').first().click();
  await dialogButton('Insert').click();
  await sleep(700);
  await toEnd();
  await slash('caption');
  await dialogButton('Insert').click();
  await sleep(600);
  await w.mouse.dblclick(b.x + 40, b.y + b.height / 2);
  await slash('comment');
  await sleep(500);
  await w.keyboard.type('Please check this sentence.');
  await w.keyboard.press('Control+Enter');
  await sleep(800);
  await save();
  const x = await xml();
  assert.match(x, /<w:bookmarkStart[^>]*w:name="KeyPoint"/);
  assert.match(x, /REF KeyPoint/);
  assert.match(x, /SEQ Figure/);
  assert.match(x, /<w:commentRangeStart\b/);
  assert.match(await xml('word/comments.xml'), /Please check this sentence\./);
});

test('track changes: what is typed while it is on is saved as an insertion', opts, async () => {
  await open('sample.docx');
  await tab('Review');
  await w.locator('.ribbon button:visible', { hasText: /^Track Changes/ }).first().click();
  await sleep(400);
  assert.match(await w.locator('.status-bar').textContent(), /Track Changes: On/);
  await toEnd();
  await w.keyboard.type('Tracked insertion.');
  await save();
  assert.match(await xml(), /<w:ins [^>]*>[\s\S]{0,400}?Tracked insertion\./);
});

test('header with a different first page, page numbers and a watermark', opts, async () => {
  await open('sample.docx');
  await tab('Insert');
  await tip('Edit header').click();
  await sleep(400);
  await menuItem('Edit Header').click();
  await sleep(800);
  await w.keyboard.type('Header typed in Kertas');
  await tip('Use a separate header').click();
  await sleep(400);
  const body = await w.locator('.ProseMirror').first().boundingBox();
  await w.mouse.dblclick(body.x + 200, Math.max(body.y + 200, 520));
  await sleep(700);
  await tab('Insert');
  await tip('Page Number').click();
  await sleep(400);
  await menuItem('Bottom Center').click();
  await sleep(700);
  await tab('Design');
  await tip('Add ghosted text').click();
  await sleep(400);
  await menuItem('DRAFT').click();
  await sleep(700);
  await save();
  const names = await partNames();
  const texts = await Promise.all(names.filter((p) => /^word\/(header|footer)\d*\.xml$/.test(p)).map(async (p) => [p, await xml(p)]));
  assert.ok(texts.some(([p, t]) => /header/.test(p) && /Header typed in Kertas/.test(t)), 'header text');
  assert.ok(texts.some(([p, t]) => /footer/.test(p) && /PAGE/.test(t)), 'page number field in a footer');
  assert.ok(texts.some(([p, t]) => /header/.test(p) && /DRAFT/.test(t)), 'watermark in a header');
  assert.match(await xml(), /<w:titlePg\b/);
});

test('a picture from the computer, and an equation typed as LaTeX then edited by double-click', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await answerDialogs([fixture('logo.png')]);
  await slash('picture');
  await sleep(1500);
  await toEnd();
  await slash('equation');
  await w.locator('[placeholder^="e.g. x ="]').fill('E = mc^2');
  await sleep(400);
  await w.locator('button.btn-primary:visible', { hasText: 'Insert' }).click();
  await sleep(1000);
  const eq = await w.locator('.ProseMirror math').last().boundingBox();
  await w.mouse.dblclick(eq.x + eq.width / 2, eq.y + eq.height / 2);
  await sleep(800);
  assert.equal(await w.locator('[placeholder^="e.g. x ="]').inputValue(), 'E = mc^2', 'the dialog reopens with the equation');
  await w.locator('[placeholder^="e.g. x ="]').fill('E = mc^2 + \\frac{a}{b}');
  await sleep(400);
  await w.locator('button.btn-primary:visible').last().click();
  await sleep(1000);
  await save();
  const x = await xml();
  assert.match(x, /<w:drawing>[\s\S]*?<a:blip\b/, 'the picture is in the body');
  assert.ok((await partNames()).some((p) => /^word\/media\/.+\.png$/i.test(p)), 'its file is in the package');
  assert.match(x, /<m:oMath>[\s\S]*?<m:f>/, 'the edited equation has its fraction');
});

test('page layout: two columns and landscape', opts, async () => {
  await open('sample.docx');
  await tab('Layout');
  await tip('Columns').click();
  await sleep(400);
  await menuItem('Two').click();
  await sleep(600);
  await tip('Orientation').click();
  await sleep(400);
  await menuItem('Landscape').click();
  await sleep(700);
  await save();
  const x = await xml();
  assert.match(x, /<w:cols\b[^>]*w:num="2"/);
  assert.match(x, /<w:pgSz\b[^>]*w:orient="landscape"/);
});

test('Export as PDF from Kertas writes a PDF', opts, async () => {
  await open('sample.docx');
  await answerDialogs([], 'exported.pdf');
  await k.locator('#btn-palette').click();
  await sleep(500);
  await k.keyboard.type('Export as PDF');
  await sleep(300);
  await k.keyboard.press('Enter');
  const out = path.join(dir, 'exported.pdf');
  for (let i = 0; i < 40 && !existsSync(out); i++) await sleep(500);
  assert.ok(existsSync(out), 'the PDF was written');
  assert.ok((await stat(out)).size > 2000);
  assert.equal((await readFile(out)).subarray(0, 5).toString(), '%PDF-');
});

test('a template and a PDF converted to Word both open in the new Writer', opts, async () => {
  await open('sample.docx');
  const before = editors().length;
  await k.locator('#tabs .tab').first().click();
  await k.getByText('Templates', { exact: true }).first().click();
  await sleep(700);
  await k.locator('[data-tpl]').first().click();
  for (let i = 0; i < 40 && editors().length === before; i++) await sleep(300);
  assert.equal(editors().length, before + 1, 'the template opened as a new-Writer tab');
  await answerDialogs([fixture('report.pdf')]);
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+o');
  await k.waitForSelector('[title="Turn this PDF into an editable Word document"]', { state: 'attached', timeout: 20000 });
  await sleep(1500);
  const button = k.locator('[title="Turn this PDF into an editable Word document"]').first();
  if (!(await button.isVisible())) { await k.getByText('Convert', { exact: true }).first().click(); await sleep(400); }
  await button.click();
  for (let i = 0; i < 60 && editors().length === before + 1; i++) await sleep(500);
  assert.equal(editors().length, before + 2, 'the converted document opened as a new-Writer tab');
  const converted = editors().at(-1);
  await converted.waitForSelector('.ProseMirror');
  await sleep(1500);
  assert.ok((await converted.locator('.ProseMirror').first().textContent()).length > 200, 'with the text of the PDF');
});

test('find and replace all', opts, async () => {
  await open('sample.docx');
  await w.keyboard.press('Control+h');
  await sleep(600);
  await w.locator('input[placeholder="Find"]:visible').fill('Bullet');
  await w.locator('input[placeholder="Replace with"]:visible').fill('Point');
  await sleep(400);
  await w.locator('button:visible', { hasText: /^Replace All$/ }).click();
  await sleep(600);
  await save();
  const text = (await xml()).replace(/<[^>]+>/g, ' ');
  assert.match(text, /Point one/);
  assert.doesNotMatch(text, /Bullet one/);
});

test('tables: Tab in the last cell adds a row, Sort orders the rows, Merge Cells joins two', opts, async () => {
  await open('sample.docx');
  const last = await w.locator('.ProseMirror td').last().boundingBox();
  await w.mouse.click(last.x + last.width - 12, last.y + last.height / 2);
  await sleep(250);
  await w.keyboard.press('End');
  await w.keyboard.press('Tab');
  await w.keyboard.type('Apple');
  await w.keyboard.press('Tab');
  await w.keyboard.type('3');
  await sleep(300);
  await tab('Table Layout');
  await tip('Sort').click();
  await sleep(600);
  await w.locator('.modal:visible label', { hasText: /^Header row$/ }).last().click();
  await dialogButton('OK').click();
  await sleep(700);
  await save();
  let cells = [...(await xml()).matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map((m) => m[0].replace(/<[^>]+>/g, '').trim());
  assert.deepEqual(cells, ['Name', 'Qty', 'Apple', '3', 'Pen', '10'], 'the header row stays, the rest are in order');
  // merge the two cells of the last row
  const cellsOnPage = w.locator('.ProseMirror td');
  const count = await cellsOnPage.count();
  const a = await cellsOnPage.nth(count - 2).boundingBox();
  const b = await cellsOnPage.nth(count - 1).boundingBox();
  // click one cell, Shift+click the other: the two cells are selected as cells
  await w.mouse.click(a.x + a.width / 2, a.y + a.height / 2);
  await sleep(300);
  await w.keyboard.down('Shift');
  await w.mouse.click(b.x + b.width / 2, b.y + b.height / 2);
  await w.keyboard.up('Shift');
  await sleep(400);
  await w.mouse.click(b.x + b.width / 2, b.y + b.height / 2, { button: 'right' });
  await sleep(500);
  await menuItem('Merge Cells').click();
  await sleep(600);
  await save();
  assert.match(await xml(), /<w:gridSpan w:val="2"\/>/);
});

test('word count, narrow margins and a section break', opts, async () => {
  await open('sample.docx');
  await tab('Review');
  await tip('Word Count').click();
  await sleep(500);
  assert.match((await w.locator('.modal:visible').last().textContent()).replace(/\s+/g, ' '), /Words\s*16/);
  await dialogButton('Close').click();
  await tab('Layout');
  await tip('Margins').click();
  await sleep(400);
  await w.locator('button:visible', { hasText: /^Narrow/ }).last().click();
  await sleep(600);
  await toEnd();
  await tab('Layout');
  await tip('Add a break').click();
  await sleep(400);
  await w.locator('button:visible', { hasText: /^Next Page/ }).last().click();
  await sleep(600);
  await w.keyboard.type('In a new section.');
  await save();
  const x = await xml();
  assert.match(x, /<w:pgMar\b[^>]*w:top="720"[^>]*w:left="720"|<w:pgMar\b[^>]*w:left="720"[^>]*w:top="720"/);
  assert.equal((x.match(/<w:sectPr\b/g) || []).length, 2, 'two sections');
  assert.match(x, /In a new section\./);
});

test('a text box holds its own text', opts, async () => {
  await open('sample.docx');
  await toEnd();
  await slash('text box');
  await sleep(800);
  await w.keyboard.type('Inside the box');
  await sleep(400);
  await inDoc();
  await save();
  const x = await xml();
  assert.match(x, /<w:txbxContent>[\s\S]*?Inside the box[\s\S]*?<\/w:txbxContent>/);
});
