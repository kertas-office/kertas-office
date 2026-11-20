import { test } from 'node:test';
import assert from 'node:assert/strict';
import { textLooksGarbled } from '../src/lib/pdf-convert.js';
import { settleSizes, uniformRuns, analyse } from '../src/lib/pdf-layout.js';

const page = (text, fontName) => text.split('\n').map((str) => ({ str, fontName }));

test('ordinary text in several languages is not garbled', () => {
  const samples = [
    'The library opened two new branches this year and visits rose by fourteen percent.\nFunding came from the city council (RM 4,500 a month) and a private trust.',
    'Perjanjian ini bermula pada 1 Mac 2026 dan berjalan selama dua belas bulan.\nYuran bulanan ialah RM 4,500, dibayar sebelum hari kelima setiap bulan.',
    'Le présent contrat prend effet le 1er mars 2026 pour une durée de douze mois.\nLa société s\'engage à répondre aux demandes urgentes en moins de quatre heures, à l\'égard du client.',
    'കേരള പബ്ലിക് സർവീസ് കമ്മീഷൻ നടത്തുന്ന പരീക്ഷകൾക്കുള്ള പാഠ്യപദ്ധതി ഇവിടെ നൽകിയിരിക്കുന്നു. ഓരോ വിഷയത്തിനും പ്രത്യേക ഭാഗങ്ങളുണ്ട്.',
    'يبدأ هذا الاتفاق في الأول من مارس ويستمر لمدة اثني عشر شهرا. الرسوم الشهرية تدفع قبل اليوم الخامس من كل شهر.',
    'Total 1,234.50 (12%) | Q1 361,500 | Q2 402,100 | Q3 398,400 | Q4 415,900 | email: info@example.org | https://example.org/a_b',
  ];
  for (const s of samples) assert.deepEqual(textLooksGarbled(page(s)), { garbled: false, reason: null }, s.slice(0, 30));
});

test('characters with no meaning (missing character table) are garbled', () => {
  const s = 'Annual\u0001\u0002\u0003\u0004 report of the\u0011\u0012\u0013 library network\u0005\u0006';
  assert.equal(textLooksGarbled(page(s)).reason, 'unmapped');
});

test('a pre-Unicode Indian font is garbled, by its name or by its text', () => {
  const legacy = 'tIcf ]»nIv kÀhokv I½oj³ \\S¯p¶ ]co£IÄ¡pÅ ]mTy]²Xn ChnsS \\ÂInbncn¡p¶p. Hmtcm hnjb¯n\\pw {]tXyI `mK§fp­v.';
  assert.equal(textLooksGarbled(page(legacy, 'ABCDEF+ML-TTKarthika')).reason, 'legacy-font');
  assert.equal(textLooksGarbled(page(legacy, 'ABCDEF+SomeEmbeddedFont')).reason, 'symbols');
  assert.equal(textLooksGarbled(page('Hkkjr ds lafo/kku esa ekSfyd vf/kdkjksa dk o.kZu fd;k x;k gS vkSj ;g lHkh ukxfjdksa ij ykxw gksrk gSA', 'Kruti Dev 010')).reason, 'legacy-font');
  // a Unicode font with a similar family name is fine
  assert.equal(textLooksGarbled(page('The library opened two new branches this year and visits rose.', 'ABCDEF+Kartika')).garbled, false);
});

test('a few words are never judged', () => {
  assert.equal(textLooksGarbled(page('\u0001\u0002 ')).garbled, false);
});

// ---- font sizes
const line = (y, words, sizes, extra = {}) => words.map((str, i) => ({ str, x: 60 + i * 60, y, w: 50, size: Array.isArray(sizes) ? sizes[i] : sizes, bold: false, italic: false, font: null, ...extra }));

test('recognised text: words on a line share one size, and lines of the same type one size', () => {
  const p = { width: 595, height: 842, items: [
    ...line(100, ['They', 'are', 'on', 'one', 'line'], [11.2, 8.1, 8.0, 8.3, 10.9], { ocr: true }).map((it, i) => ({ ...it, size: [11.2, 10.4, 10.2, 10.3, 10.9][i] })),
    ...line(116, ['and', 'this', 'is', 'the', 'next'], [10.1, 10.6, 9.9, 10.4, 10.2], { ocr: true }),
    ...line(132, ['a', 'third', 'line', 'of', 'body'], 9.9, { ocr: true }),
    ...line(200, ['A', 'big', 'heading'], 20.3, { ocr: true }),
  ] };
  const out = settleSizes([p])[0].items;
  const body = new Set(out.filter((it) => it.y < 190).map((it) => it.size));
  assert.equal(body.size, 1, [...body].join(','));
  assert.equal(new Set(out.filter((it) => it.y === 200).map((it) => it.size)).size, 1);
  assert.ok(out.find((it) => it.y === 200).size > [...body][0] * 1.5);
  // the input is not changed
  assert.equal(p.items[0].size, 11.2);
});

test('typeset text keeps its real sizes: 10 and 11 point stay apart, a superscript stays small', () => {
  const p = { width: 595, height: 842, items: [
    ...line(100, ['Body', 'text', 'at', 'eleven', 'points'], 11),
    ...line(116, ['with', 'a', 'note', 'mark'], [11, 11, 11, 11]), { str: '2', x: 320, y: 112, w: 5, size: 7, bold: false, italic: false, font: null },
    ...line(160, ['A', 'caption', 'at', 'ten', 'points'], 10),
    ...line(176, ['measured', 'with', 'rounding', 'noise'], [10.02, 9.98, 10.01, 10]),
  ] };
  const out = settleSizes([p])[0].items;
  assert.deepEqual([...new Set(out.map((it) => it.size))].sort((a, b) => a - b), [7, 10, 11]);
});

test('a paragraph takes one size, keeping bold and a real small run', () => {
  const runs = uniformRuns([
    { text: 'Most of it at ', size: 11, bold: false, italic: false, font: null },
    { text: 'ten and a half', size: 10.5, bold: false, italic: false, font: null },
    { text: ' and bold', size: 11.5, bold: true, italic: false, font: null },
    { text: '3', size: 7, bold: false, italic: false, font: null },
  ]);
  assert.deepEqual(runs.map((r) => [r.text, r.size, r.bold]), [['Most of it at ten and a half', 11, false], [' and bold', 11, true], ['3', 7, false]]);
});

test('analyse settles sizes before building paragraphs', () => {
  const p = { width: 595, height: 842, items: [
    ...line(100, ['One', 'paragraph', 'that', 'runs', 'over'], [10.4, 10.9, 10.1, 10.6, 10.2], { ocr: true }),
    ...line(114, ['two', 'lines', 'of', 'the', 'page.'], [10.8, 10.0, 10.3, 10.5, 10.7], { ocr: true }),
  ] };
  const blocks = analyse([p]).pages[0].blocks;
  const sizes = new Set(blocks.flatMap((b) => (b.runs || []).map((r) => r.size)));
  assert.equal(sizes.size, 1);
});

test('lines with wide spacing still join into one paragraph when the line runs to the edge', () => {
  const row = (y, text) => [{ str: text, x: 60, y, w: text.length * 5.2, size: 10, bold: false, italic: false, font: null }];
  const full = 'Our contact centre handles around nine thousand calls a month, and six in ten are customers asking';
  const p = { width: 595, height: 842, items: [
    ...row(100, full),
    ...row(120, 'where their order is or wanting to change a delivery date, which a portal would let them do without'),
    ...row(140, 'calling us.'),
    ...row(175, 'We propose building the portal over six months.'),
  ] };
  const blocks = analyse([p]).pages[0].blocks.filter((b) => b.type === 'para');
  assert.equal(blocks.length, 2);
  assert.match(blocks[0].runs.map((r) => r.text).join(''), /asking where their order.*do without calling us\.$/);
});

// ---- scans with a scanner app's text layer
import { textLayerRough, textEdges } from '../src/lib/pdf-layout.js';
import { isScan } from '../src/lib/pdf-convert.js';

test('a scanner app\'s text layer, sized word by word, is rough; text sized by the line is not', () => {
  const words = 'The Owner must ensure that the Property is fully compliant with the Brand Standards and fully adherent to the system sixty days prior to the opening of the hotel to customers'.split(' ');
  const rows = (sizeOf) => ({ width: 595, height: 842, scan: true, items: Array.from({ length: 6 }, (_, r) => words.slice(r * 5, r * 5 + 5).map((str, i) => ({ str, x: 60 + i * 80, y: 100 + r * 16, w: 70, size: sizeOf(str), bold: false, italic: false, font: null }))).flat() });
  // short words without tall letters come out much smaller, as scanner apps size them
  const rough = rows((w) => (/^(to|as|a|is|of|the|and|must|that|with)$/i.test(w) ? 6.2 : 10.4));
  assert.equal(textLayerRough(rough), true);
  assert.equal(textLayerRough(rows(() => 10.4)), false);
  // and converting the rough one still gives one size per line
  const out = settleSizes([rough])[0].items;
  assert.equal(new Set(out.map((it) => it.size)).size, 1);
});

test('a page is a scan when one picture covers most of it', () => {
  const OPS = { save: 1, restore: 2, transform: 3, paintImageXObject: 4, paintFormXObjectBegin: 5, paintFormXObjectEnd: 6 };
  const ops = (m) => ({ fnArray: [OPS.save, OPS.transform, OPS.paintImageXObject, OPS.restore], argsArray: [null, m, ['img'], null] });
  assert.equal(isScan(ops([595, 0, 0, 842, 0, 0]), OPS, 595 * 842), true);
  assert.equal(isScan(ops([120, 0, 0, 60, 40, 700]), OPS, 595 * 842), false); // a logo
  assert.equal(isScan(null, OPS, 595 * 842), false);
});

test('body lines are not centred because of a stray mark in the margin', () => {
  const row = (y, x, text) => ({ str: text, x, y, w: text.length * 4.6, size: 10, bold: false, italic: false, font: null });
  const body = 'itself or its assigned designate in order to maintain quality, brand standards and overall conformity with';
  const p = { width: 595, height: 842, items: [
    row(60, 277, 'Agreement'),                    // a real centred title
    row(100, 60, body), row(114, 60, body), row(128, 60, body), row(142, 60, 'the system.'),
    row(180, 60, body), row(194, 60, body),
    row(300, 8, '|'),                              // a scribble at the far left
    row(320, 560, 'ii'),                           // and one at the far right
  ] };
  const blocks = analyse([p]).pages[0].blocks;
  const paras = blocks.filter((b) => b.type === 'para' && b.runs.map((r) => r.text).join('').length > 60);
  assert.ok(paras.length >= 2);
  assert.ok(paras.every((b) => b.align === 'left'), paras.map((b) => b.align).join(','));
  assert.equal(blocks.find((b) => /Agreement/.test((b.runs || []).map((r) => r.text).join('')))?.align, 'center');
  assert.ok(textEdges([{ x0: 8, x1: 12, text: '|' }, { x0: 60, x1: 540, text: body }, { x0: 60, x1: 540, text: body }]).left === 60);
});
