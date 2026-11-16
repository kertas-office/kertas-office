// Writer, end to end: what a person types and clicks, then the saved .docx read back
// with JSZip (not with Kertas), then reopened in Kertas.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { openApp, openFile, openFixture, saveBytes, fixture, stopServer, sleep } from './harness.mjs';

let page;
before(async () => { page = await openApp(); });
after(async () => { await page?.close(); stopServer(); });

const docXml = async (bytes) => (await JSZip.loadAsync(bytes)).file('word/document.xml').async('string');
const pageText = () => page.locator('.superdoc-page:visible').allInnerTexts().then((t) => t.join('\n'));

/** Puts the caret at the end of the document body. */
async function caretToEnd() {
  // a click could land on a table or picture, so the editor is focused at the end directly
  await page.evaluate(() => {
    const ed = window.__kertasOoxml.editor();
    let end = 1;
    ed.state.doc.descendants((n, pos) => { if (n.isTextblock) end = pos + n.nodeSize - 1; });
    ed.commands.setTextSelection({ from: end, to: end });
    ed.view?.focus?.();
  });
  await sleep(150);
}

/** Runs a slash command by typing its name, as a person would. */
async function slash(query) {
  await page.keyboard.press('Enter');
  await page.keyboard.type('/');
  await page.locator('.slash-menu').waitFor({ state: 'visible' });
  await page.keyboard.type(query, { delay: 20 });
  await sleep(150);
  await page.keyboard.press('Enter');
  await sleep(400);
}

/** Confirms a Kertas dialog if one opened. */
async function confirmDialog() {
  const dlg = page.locator('.modal.dialog');
  if (await dlg.count()) { await dlg.locator('button.primary, button[type=submit]').last().click(); await sleep(400); }
}

test('a new document: type, bold, save and reopen', async () => {
  await openFile(page, 'writer', 'New document.docx');
  await caretToEnd();
  await page.keyboard.type('Kertas end to end ');
  await page.keyboard.press('Control+b');
  await page.keyboard.type('bold words');
  await page.keyboard.press('Control+b');
  await sleep(300);

  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.match(xml, /Kertas end to end/);
  assert.match(xml, /<w:r>(?:(?!<\/w:r>).)*<w:b\s*\/>(?:(?!<\/w:r>).)*bold words/s, 'bold run saved');
  assert.deepEqual(page.errors, []);
});

test('line spacing menu offers 1.08 as the default', async () => {
  await page.locator('.rb-tab[data-tab="home"]').click();
  await page.locator('[data-id="spacing"]').click();
  await sleep(250);
  const items = await page.locator('.menu').last().innerText();
  assert.match(items, /1\.08 \(default\)/);
  assert.doesNotMatch(items, /Word default/);
  await page.keyboard.press('Escape');
});

test('slash menu: table, bulleted list and tick box field are saved', async () => {
  await caretToEnd();
  await slash('table');
  await confirmDialog();
  await caretToEnd();
  await slash('bulleted');
  await page.keyboard.type('First point');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Second point');
  await caretToEnd();
  await slash('tick box field');
  await confirmDialog();

  const bytes = await saveBytes(page, 'docx');
  const xml = await docXml(bytes);
  assert.match(xml, /<w:tbl>/, 'table');
  assert.match(xml, /<w:numPr>/, 'list numbering');
  assert.match(xml, /First point/);
  assert.match(xml, /<w:sdt>(?:(?!<\/w:sdt>).)*w14:checkbox/s, 'tick box form field');
  assert.deepEqual(page.errors, []);
});

test('right-click on a table shows the table menu, and scrolling closes it', async () => {
  const table = page.locator('.superdoc-page:visible .superdoc-table-fragment').first();
  await table.click({ button: 'right', position: { x: 20, y: 8 } });
  const menu = page.locator('.context-menu').first();
  await menu.waitFor({ state: 'visible' });
  const text = await menu.innerText();
  for (const want of [/Borders/i, /Row height|Row/i, /Table design|Design/i]) assert.match(text, want);
  await page.mouse.wheel(0, 400);
  await sleep(400);
  assert.equal(await page.locator('.context-menu:visible').count(), 0, 'menu closes on scroll');
});

test('a picture dropped on the page is inserted and saved', async () => {
  const png = (await readFile(fixture('logo.png'))).toString('base64');
  const first = page.locator('.superdoc-page:visible').first();
  await first.evaluate((el) => el.scrollIntoView({ block: 'start' }));
  await sleep(300);
  const box = await first.boundingBox();
  await page.evaluate(async ([png, x, y]) => {
    const bytes = Uint8Array.from(atob(png), (c) => c.charCodeAt(0));
    const dt = new DataTransfer();
    dt.items.add(new File([bytes], 'logo.png', { type: 'image/png' }));
    const target = document.elementFromPoint(x, y);
    for (const t of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(t, { bubbles: true, cancelable: true, clientX: x, clientY: y, dataTransfer: dt }));
  }, [png, box.x + 200, box.y + 160]);
  await sleep(1500);
  assert.equal(await page.evaluate(() => document.body.classList.contains('dragging')), false, 'drop cover cleared');

  const zip = await JSZip.loadAsync(await saveBytes(page, 'docx'));
  assert.ok(Object.keys(zip.files).some((f) => f.startsWith('word/media/')), 'picture stored');
  assert.match(await zip.file('word/document.xml').async('string'), /<w:drawing>/);
});

test('saved file reopens with its content', async () => {
  const bytes = await saveBytes(page, 'docx');
  await openFile(page, 'writer', 'Reopened.docx', bytes);
  const text = await pageText();
  assert.match(text, /Kertas end to end/);
  assert.match(text, /Second point/);
  assert.ok(await page.locator('.superdoc-page:visible img').count() >= 1, 'picture shown');
  assert.deepEqual(page.errors, []);
});

test('an existing Word file opens and round-trips', async () => {
  const original = await readFile(fixture('sample.docx'));
  await openFixture(page, 'writer', 'sample.docx');
  const before = (await docXml(original)).match(/<w:t[^>]*>[^<]+/g)?.length || 0;
  const xml = await docXml(await saveBytes(page, 'docx'));
  const after = xml.match(/<w:t[^>]*>[^<]+/g)?.length || 0;
  assert.ok(after >= before * 0.9, `text runs kept (${after} of ${before})`);
  assert.deepEqual(page.errors, []);
});

test('two documents open: each gets what is typed into it (1.0.1 bug)', async () => {
  await openFile(page, 'writer', 'First.docx');
  await caretToEnd();
  await page.keyboard.type('typed in first', { delay: 10 });
  await openFile(page, 'writer', 'Second.docx');
  await caretToEnd();
  await page.keyboard.type('typed in second', { delay: 10 });
  await sleep(400);
  assert.match(await pageText(), /typed in second/);
  await page.locator('#tabs .tab', { hasText: 'First.docx' }).click();
  await sleep(800);
  await caretToEnd();
  await page.keyboard.type(' again', { delay: 10 });
  await sleep(400);
  const first = await pageText();
  assert.match(first, /typed in first again/);
  assert.doesNotMatch(first, /typed in second/);
  assert.deepEqual(page.errors, []);
});

// Compare documents: Review > Compare, then pick the other version
async function compareWith(role, otherFixture) {
  await page.locator('.rb-tab[data-tab="review"]:visible').click();
  await page.locator('[data-id="compare"]:visible').click();
  if (role === 'revised') await page.locator('.modal.dialog [data-f="role"] [data-v="revised"]').click();
  const chooser = page.waitForEvent('filechooser');
  await page.locator('.modal.dialog button.primary').click();
  await (await chooser).setFiles(fixture(otherFixture));
  await page.waitForFunction(() => /differences? found/.test(document.body.innerText), null, { timeout: 30000 });
  await sleep(500);
  // the comparison is the tab the helpers save from now on
  await page.evaluate(() => { window.__t = [...window.__kertas.tabs.values()].find((t) => / compared with /.test(t.name)); });
}
const textBlocks = () => page.evaluate(() => { const out = []; window.__kertasOoxml.editor().state.doc.descendants((n) => { if (n.isTextblock) { out.push(n.textContent); return false; } return true; }); return out; });
const trackedTexts = () => page.evaluate(() => { const out = []; window.__kertasOoxml.editor().state.doc.descendants((n) => { if (n.isText) for (const m of n.marks) if (m.type.name === 'trackInsert' || m.type.name === 'trackDelete') out.push(`${m.type.name === 'trackInsert' ? '+' : '-'}${n.text}`); }); return out; });

test('compare documents: differences open as tracked changes; accepting them all gives the revised text', async () => {
  await openFixture(page, 'writer', 'compare-revised.docx');
  const revised = await textBlocks();
  await openFixture(page, 'writer', 'compare-original.docx');
  const original = await textBlocks();
  await compareWith('original', 'compare-revised.docx');
  assert.match(await page.locator('#tabs .tab.active, #tabs .tab[aria-selected="true"]').first().innerText(), /compare-original compared with compare-revised/);
  assert.deepEqual(await trackedTexts(), [
    '+Draft for the board.', '-two', '+three', '-fourteen', '+eighteen',
    '-This paragraph will be removed in the revised version.', '+A new paragraph added in the middle.',
    '-stayed level', '+grew', '-at', '+to', '-two', '+five', '-361,500', '+412,300',
  ]);
  // credited to the revised version, styled like it
  const info = await page.evaluate(() => {
    const ed = window.__kertasOoxml.editor(); const authors = new Set(); let draftStyle;
    ed.state.doc.descendants((n) => { if (n.isTextblock && n.textContent === 'Draft for the board.') draftStyle = n.attrs.paragraphProperties?.styleId ?? null; for (const m of n.marks) if (m.attrs.author) authors.add(m.attrs.author); return true; });
    return { authors: [...authors], draftStyle };
  });
  assert.deepEqual(info, { authors: ['compare-revised'], draftStyle: null });
  // saved as Word tracked changes
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.match(xml, /<w:ins [^>]*w:author="compare-revised"/);
  assert.match(xml, /<w:del [^>]*w:author="compare-revised"/);
  // accepting one whole-paragraph deletion takes its paragraph with it
  await page.locator('.writer-side:visible .tc', { hasText: 'This paragraph will be removed' }).locator('[data-a="accept"]').click();
  await sleep(500);
  assert.ok(!(await textBlocks()).includes(''), 'no empty paragraph left behind');
  // accept the rest: exactly the revised document
  await page.locator('.rb-tab[data-tab="review"]:visible').click();
  await page.locator('[data-id="acc-all"]:visible').click();
  await sleep(500);
  assert.deepEqual(await textBlocks(), revised);
  assert.equal(original.length, 10);
  assert.deepEqual(page.errors, []);
});

test('compare documents: from the revised version, rejecting every change gives the original text', async () => {
  await openFixture(page, 'writer', 'compare-original.docx');
  const original = await textBlocks();
  await openFixture(page, 'writer', 'compare-revised.docx');
  await compareWith('revised', 'compare-original.docx');
  const tracked = await trackedTexts();
  assert.ok(tracked.includes('+eighteen'), tracked.join(' '));
  await page.locator('.rb-tab[data-tab="review"]:visible').click();
  await page.locator('[data-id="rej-all"]:visible').click();
  await sleep(500);
  assert.deepEqual(await textBlocks(), original);
  assert.deepEqual(page.errors, []);
});
