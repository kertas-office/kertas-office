import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { extractPages, pdfToDocx, pdfToSnapshot, fontInfo } from '../src/lib/pdf-convert.js';
import { analyse, parseNumber } from '../src/lib/pdf-layout.js';
import { buildPdf } from '../src/lib/pdf-ops.js';
import { setFontLoader } from '../src/lib/pdf-text.js';
import { nodeFontLoader } from './font-loader.mjs';

setFontLoader(nodeFontLoader);
const open = async (bytes) => pdfjs.getDocument({ data: new Uint8Array(bytes), fontExtraProperties: true, standardFontDataUrl: new URL('../node_modules/pdfjs-dist/standard_fonts/', import.meta.url).pathname }).promise;
const report = async () => extractPages(await open(await readFile(new URL('./fixtures/report.pdf', import.meta.url))), { Util: pdfjs.Util });

test('numbers and fonts are read the way people write them', () => {
  assert.deepEqual(parseNumber('1,250,000'), { value: 1250000, format: '#,##0' });
  assert.deepEqual(parseNumber('(1,234.50)'), { value: -1234.5, format: '#,##0.00' });
  assert.deepEqual(parseNumber('34.4%'), { value: 0.344, format: '0.0%' });
  assert.equal(parseNumber('RM 12').value, 12);
  assert.equal(parseNumber('Q3 2026'), null);
  assert.deepEqual(fontInfo('ABCDEF+TimesNewRomanPS-BoldItalicMT'), { family: 'Times New Roman', bold: true, italic: true });
});

test('a report PDF becomes headings, paragraphs, lists and a table', async () => {
  const pages = await report();
  const { pages: [p] } = analyse(pages);
  const kinds = p.blocks.map((b) => (b.type === 'heading' ? `h${b.level}` : b.type === 'list' ? (b.ordered ? 'ol' : 'ul') : b.type));
  assert.deepEqual(kinds, ['h1', 'h2', 'para', 'para', 'h3', 'ul', 'h3', 'ol', 'h2', 'table', 'para']);
  const table = p.blocks.find((b) => b.type === 'table');
  assert.deepEqual(table.rows[1], ['North', '1,250,000', '820,000', '34.4%']);
  assert.match(p.blocks[2].runs.map((r) => r.text).join(''), /^Sales grew in every region this year\. .* a quarter of all revenue\.$/);

  const docx = await JSZip.loadAsync(await pdfToDocx(pages));
  const xml = await docx.file('word/document.xml').async('string');
  assert.match(xml, /w:pStyle w:val="Heading1"/);
  assert.equal((xml.match(/<w:tbl>/g) || []).length, 1);
  assert.equal((xml.match(/<w:numPr>/g) || []).length, 6, 'three bullets and three numbered items');
  assert.match(xml, /Prepared by the finance team\./);
});

test('PDF tables become numbers in a workbook', async () => {
  const snap = pdfToSnapshot(await report(), { name: 'r' });
  const sh = snap.sheets[snap.sheetOrder[0]];
  const rows = Object.values(sh.cellData).map((r) => Object.values(r));
  const north = rows.find((r) => r[0].v === 'North');
  assert.deepEqual(north.map((c) => c.v), ['North', 1250000, 820000, 0.344]);
  assert.equal(snap.styles[north[3].s].n.pattern, '0.0%');
});

test('recognised words are saved as an invisible, searchable text layer', async () => {
  const scan = await readFile(new URL('./fixtures/scan.pdf', import.meta.url));
  const words = [{ str: 'Annual', x: 60, y: 760, w: 70, size: 20 }, { str: 'Sales', x: 136, y: 760, w: 55, size: 20 }, { str: 'മലയാളം', x: 60, y: 720, w: 80, size: 14 }];
  const out = await buildPdf({ a: new Uint8Array(scan) }, [{ src: 'a', index: 0, annots: [{ type: 'ocrText', words }] }]);
  const page = await (await open(out)).getPage(1);
  const text = (await page.getTextContent()).items.map((i) => i.str).join(' ');
  assert.match(text, /Annual/);
  assert.match(text, /Sales/);
  const ops = await page.getOperatorList();
  assert.ok(ops.fnArray.includes(pdfjs.OPS.setTextRenderingMode), 'drawn with an invisible text mode');
});

test('existing PDF text can be replaced in place', async () => {
  const bytes = await readFile(new URL('./fixtures/report.pdf', import.meta.url));
  const doc = await open(bytes);
  const page = await doc.getPage(1);
  const { height } = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent();
  const item = tc.items.find((it) => it.str.includes('Sales grew in every region'));
  assert.ok(item, 'fixture has the paragraph');
  const [, , , size, x, y] = item.transform;
  // Box around the whole run in PDF space
  const box = { x0: x - 1, y0: y - size * 0.3, x1: x + item.width + 1, y1: y + size * 0.9 };
  const out = await buildPdf({ a: new Uint8Array(bytes) }, [{ src: 'a', index: 0, annots: [{ type: 'editText', box, text: 'Revenue rose across all markets' }] }]);
  const edited = await open(out);
  const text = (await (await edited.getPage(1)).getTextContent()).items.map((i) => i.str).join(' ');
  assert.doesNotMatch(text, /Sales grew in every region/);
  assert.match(text, /Revenue rose across all markets/);
  assert.match(text, /Prepared by the finance team/, 'other text is untouched');
  assert.ok(height > 0);
});
