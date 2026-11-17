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

test('cross-references saved by Word are left as they are on open; heading numbers are never replaced', async () => {
  // REF \r (the heading's number, 1.1) and a REF whose saved text is out of date. (The PAGEREF in
  // the fixture opens as the editor's own page reference, which this code does not touch.)
  await openFixture(page, 'writer', 'references-word.docx');
  await sleep(1500);
  assert.deepEqual(await crossRefs(), [['REF _Ref \\r \\h', '1.1'], ['REF _Ref \\h \\* MERGEFORMAT', 'Old title']], 'opening the document changes nothing');
  // once the user edits, the ones Kertas can work out follow the document; the heading number stays
  await caretAtDocEnd();
  await page.keyboard.type(' x');
  await page.waitForFunction(() => { const t = []; window.__kertasOoxml.editor().state.doc.descendants((n) => { if (n.type.name === 'crossReference') t.push(n.attrs.resolvedText); return true; }); return t[1] === 'Results'; }, null, { timeout: 5000 });
  assert.deepEqual(await crossRefs(), [['REF _Ref \\r \\h', '1.1'], ['REF _Ref \\h \\* MERGEFORMAT', 'Results']]);
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.match(xml, /REF _Ref100 \\r \\h/);
  assert.match(xml, /<w:t[^>]*>1\.1<\/w:t>/, 'the heading number is saved as Word had it');
  assert.deepEqual(page.errors, []);
});

test('a merged cell taller than a page breaks across pages; a short one moves whole', async () => {
  const pagesText = () => page.evaluate(() => [...document.querySelectorAll('.view.active .superdoc-page')].map((p) => p.innerText.replace(/\s+/g, ' ')));
  // 60 rows joined by one merged cell: the table starts on page 1 and continues on page 2
  await openFixture(page, 'writer', 'merged-tall.docx');
  let pages = await pagesText();
  assert.equal(pages.length, 2, 'two pages, not three');
  assert.match(pages[0], /One merged cell Item 01/, 'the table starts on page 1');
  assert.match(pages[1], /^Group Item Value Item \d\d/, 'page 2 repeats the header and goes on');
  assert.match(pages[1], /Item 60 20/);
  // the merged cell on page 1 ends with the table, not at the bottom of the page
  const fit = await page.evaluate(() => {
    const p1 = document.querySelector('.view.active .superdoc-page');
    const table = p1.querySelector('.superdoc-table-fragment').getBoundingClientRect();
    const cells = [...p1.querySelectorAll('.superdoc-table-fragment div')].filter((d) => d.innerText.trim() === 'One merged cell' && d.style.position === 'absolute');
    return cells.length ? cells[0].getBoundingClientRect().bottom - table.bottom : null;
  });
  assert.ok(fit !== null && fit <= 1, `merged cell stays inside the table on page 1 (${fit})`);
  // six rows joined by a merged cell near the end of a page move to the next page together
  await openFixture(page, 'writer', 'merged-short.docx');
  pages = await pagesText();
  assert.match(pages[1], /Group Item Long merged text that must stay whole: [a-z ]+twelve Row 1 Row 2 Row 3 Row 4 Row 5 Row 6/);
  assert.doesNotMatch(pages[0], /Row 1/);
  assert.deepEqual(page.errors, []);
});

test('diagonal cell borders are drawn in their colours, and saved unchanged', async () => {
  await openFixture(page, 'writer', 'diagonal.docx');
  const lines = await page.evaluate(() => [...document.querySelectorAll('.view.active svg[data-kertas-diagonal] line')]
    .map((l) => ({ stroke: l.getAttribute('stroke'), down: Number(l.getAttribute('y2')) > Number(l.getAttribute('y1')), right: Number(l.getAttribute('x2')) > Number(l.getAttribute('x1')) })));
  assert.deepEqual(lines, [
    { stroke: '#C00000', down: true, right: true },  // top left to bottom right
    { stroke: '#0070C0', down: true, right: false }, // top right to bottom left
  ]);
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.match(xml, /<w:tl2br [^>]*w:color="C00000"/);
  assert.match(xml, /<w:tr2bl [^>]*w:color="0070C0"/);
  assert.deepEqual(page.errors, []);
});

test('symbol fonts: Wingdings and Symbol characters are drawn, and saved unchanged', async () => {
  await openFixture(page, 'writer', 'symbols.docx');
  const shown = await page.locator('.view.active .superdoc-page').first().innerText();
  for (const ch of ['✔', '✘', 'α', '☑', '☑', 'π', 'ρ']) assert.ok(shown.includes(ch), `${ch} is drawn`);
  assert.ok(!/[-]/.test(shown), 'no private-use codes on screen');
  // the tick typed as the letter u-umlaut in Wingdings shows as a tick, not as the letter
  assert.match(shown, /✔ tick typed as a letter/);
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.equal((xml.match(/<w:sym /g) || []).length, 3, 'w:sym elements kept');
  assert.match(xml, /w:font="Wingdings" w:char="F0FC"/);
  assert.ok(xml.includes('') && xml.includes('ü'), 'symbol text saved with its original codes');
  assert.match(xml, /p = 2r/);
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

test('equations: insert, double-click to edit, right-click to edit, save', async () => {
  const t = await page.evaluate(async () => { const t = await window.__kertas.openTab('writer', { name: 'Equations.docx' }); while (!t.instance) await new Promise((r) => setTimeout(r, 100)); window.__t = t; return t.id; });
  assert.ok(t);
  await sleep(2500);
  const rb = async (tab, id) => { await page.locator(`.rb-tab[data-tab="${tab}"]:visible`).click(); await page.locator(`[data-id="${id}"]:visible`).first().click(); };
  await rb('insert', 'equation');
  const tex = page.locator('.modal [data-f="tex"]');
  assert.match(await page.locator('.modal').innerText(), /Writer and Word can both edit it/);
  await tex.fill('a^2+b^2=c^2');
  await page.locator('.modal').getByRole('button', { name: 'Insert', exact: true }).click();
  await page.locator('.view.active .sd-math').first().waitFor();
  await sleep(500);
  // double-click opens it as LaTeX
  await page.locator('.view.active .sd-math').first().dblclick();
  await page.locator('.modal').waitFor();
  assert.match(await page.locator('.modal').innerText(), /Edit equation/);
  assert.equal(await tex.inputValue(), 'a^2+b^2=c^2');
  await tex.fill('x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}');
  await page.locator('.modal').getByRole('button', { name: 'Update', exact: true }).click();
  await sleep(800);
  // right-click > Edit equation
  await page.locator('.view.active .sd-math').first().click({ button: 'right' });
  await page.getByText('Edit equation...').click();
  await page.locator('.modal').waitFor();
  assert.equal(await tex.inputValue(), 'x=\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}');
  await tex.fill('E=mc^2');
  await page.locator('.modal').getByRole('button', { name: 'Update', exact: true }).click();
  await sleep(800);
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.equal(xml.match(/<m:oMathPara/g)?.length, 1);
  const text = [...xml.matchAll(/<m:t[^>]*>([^<]*)<\/m:t>/g)].map((m) => m[1]).join('');
  assert.equal(text, 'E=mc2');
  assert.deepEqual(page.errors, []);
});

test('equations from Word: one on its own line and one in a sentence can both be edited', async () => {
  await openFixture(page, 'writer', 'equations.docx');
  await sleep(1000);
  const tex = page.locator('.modal [data-f="tex"]');
  const eqs = page.locator('.view.active .sd-math');
  assert.equal(await eqs.count(), 2);
  await eqs.nth(0).dblclick();
  await page.locator('.modal').waitFor();
  assert.equal(await tex.inputValue(), 'A=\\pi r^2');
  await tex.fill('A=\\pi r^2+1');
  await page.locator('.modal').getByRole('button', { name: 'Update', exact: true }).click();
  await sleep(800);
  await page.locator('.view.active .sd-math').nth(1).dblclick();
  await page.locator('.modal').waitFor();
  assert.equal(await tex.inputValue(), '\\sqrt{2}');
  await tex.fill('\\sqrt{3}');
  await page.locator('.modal').getByRole('button', { name: 'Update', exact: true }).click();
  await sleep(800);
  const xml = await docXml(await saveBytes(page, 'docx'));
  assert.equal(xml.match(/<m:oMathPara/g)?.length, 1);
  assert.match(xml, /Each side has length/);
  assert.match(xml, /metres\./);
  const text = [...xml.matchAll(/<m:t[^>]*>([^<]*)<\/m:t>/g)].map((m) => m[1]).join('|');
  assert.match(text, /A\|=\|π\|r\|2\|\+\|1/);
  assert.match(text, /3/);
  assert.doesNotMatch(text, /\|2\|.*\|2$/);
  assert.deepEqual(page.errors, []);
});
