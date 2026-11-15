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

// Bookmarks and cross-references
const rb = async (tab, id) => { await page.locator(`.rb-tab[data-tab="${tab}"]:visible`).click(); await page.locator(`[data-id="${id}"]:visible`).click(); };
const caretAtDocEnd = () => page.evaluate(() => { const ed = window.__kertasOoxml.editor(); let last = null; ed.state.doc.descendants((n, p) => { if (n.isTextblock) last = p + n.nodeSize - 1; return !n.isTextblock; }); ed.commands.setTextSelection({ from: last, to: last }); ed.view.focus(); });
const crossRefs = () => page.evaluate(() => { const out = []; window.__kertasOoxml.editor().state.doc.descendants((n) => { if (n.type.name === 'crossReference') out.push([n.attrs.instruction.replace(/_Ref\d+/, '_Ref'), n.attrs.resolvedText]); return true; }); return out; });
async function insertCrossRef(kind, item, show) {
  await caretAtDocEnd();
  await rb('insert', 'xref');
  await page.locator(`.modal.dialog [data-f="kind"] [data-v="${kind}"]`).click();
  await page.locator(`.modal.dialog select[data-f="item-${kind}"]`).selectOption({ label: item });
  await page.locator('.modal.dialog select[data-f="show"]').selectOption(show);
  await page.locator('.modal.dialog button.primary').click();
  await sleep(300);
  await page.keyboard.type(' / ');
}

test('cross-references show a heading, its page and where it is, follow edits, and save as Word fields', async () => {
  await openFixture(page, 'writer', 'references.docx');
  await insertCrossRef('heading', 'Results', 'text');
  await insertCrossRef('heading', 'Results', 'page');
  await insertCrossRef('heading', '   Funding sources', 'aboveBelow');
  await insertCrossRef('caption', 'Table 1: Spending by department', 'text');
  assert.deepEqual(await crossRefs(), [
    ['REF _Ref \\h', 'Results'], ['PAGEREF _Ref \\h', '2'], ['REF _Ref \\p \\h', 'above'], ['REF _Ref \\h', 'Table 1: Spending by department'],
  ]);
  // renaming the heading updates the reference by itself
  await page.evaluate(() => { const ed = window.__kertasOoxml.editor(); let at = null; ed.state.doc.descendants((n, p) => { if (n.isText && n.text === 'Results') at = p; return true; }); ed.view.dispatch(ed.state.tr.insertText('Final results', at, at + 7)); });
  await page.waitForFunction(() => { let t = ''; window.__kertasOoxml.editor().state.doc.descendants((n) => { if (n.type.name === 'crossReference' && !t) t = n.attrs.resolvedText; return true; }); return t === 'Final results'; }, null, { timeout: 5000 });
  assert.match(await page.locator('.view.active .superdoc-page').nth(1).innerText(), /See also: Final results \/ 2 \/ above \/ Table 1: Spending by department/);
  const xml = await docXml(await saveBytes(page, 'docx'));
  const fields = [...xml.matchAll(/<w:instrText[^>]*>([^<]*)</g)].map((m) => m[1].trim().replace(/_Ref\d+/, '_Ref'));
  assert.deepEqual(fields, ['REF _Ref \\h', 'PAGEREF _Ref \\h', 'REF _Ref \\p \\h', 'REF _Ref \\h']);
  assert.equal((xml.match(/<w:bookmarkStart [^>]*w:name="_Ref\d+"/g) || []).length, 3, 'one hidden bookmark per heading or caption referred to');
  assert.deepEqual(page.errors, []);
});

test('bookmarks: add one around selected text, refer to it, go to it and delete it', async () => {
  await openFixture(page, 'writer', 'references.docx');
  // select "city council"
  await page.evaluate(() => { const ed = window.__kertasOoxml.editor(); let at = null; ed.state.doc.descendants((n, p) => { if (n.isText && n.text.includes('city council')) at = p + n.text.indexOf('city council'); return true; }); ed.commands.setTextSelection({ from: at, to: at + 12 }); ed.view.focus(); });
  await rb('insert', 'bookmark');
  assert.equal(await page.locator('.modal.dialog [data-f="name"]').inputValue(), 'city_council', 'name suggested from the selection');
  await page.locator('.modal.dialog [data-f="name"]').fill('Council 1');
  await page.locator('.modal.dialog button.primary').click();
  assert.match(await page.locator('.modal.dialog .dlg-error').innerText(), /letters, numbers and _/, 'spaces are refused');
  await page.locator('.modal.dialog [data-f="name"]').fill('Council');
  await page.locator('.modal.dialog button.primary').click();
  await insertCrossRef('bookmark', 'Council  (city council)', 'text');
  assert.deepEqual(await crossRefs(), [['REF Council \\h', 'city council']]);
  // Go to selects the bookmarked text
  await caretAtDocEnd();
  await rb('insert', 'bookmark');
  await page.locator('.modal.dialog button[data-a="goto"]').click();
  assert.equal(await page.evaluate(() => { const ed = window.__kertasOoxml.editor(); const { from, to } = ed.state.selection; return ed.state.doc.textBetween(from, to); }), 'city council');
  // deleting the bookmark keeps the text and the reference's last value
  await rb('insert', 'bookmark');
  await page.locator('.modal.dialog button[data-a="delete"]').click();
  await sleep(300);
  assert.equal(await page.evaluate(() => { let n = 0; window.__kertasOoxml.editor().state.doc.descendants((x) => { if (x.type.name === 'bookmarkStart' && !x.attrs.name.startsWith('_')) n++; return true; }); return n; }), 0);
  assert.match(await page.locator('.view.active .superdoc-page').first().innerText(), /Most money comes from the city council\./);
  assert.deepEqual(page.errors, []);
});
