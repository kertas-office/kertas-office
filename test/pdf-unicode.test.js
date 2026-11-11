import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { PDFDocument, PDFName } from '@cantoo/pdf-lib';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { setFontLoader, splitRuns, scriptOf } from '../src/lib/pdf-text.js';
import { buildPdf } from '../src/lib/pdf-ops.js';
import { snapshotToPdf } from '../src/lib/sheet-pdf.js';
import { blankSnapshot } from '../src/lib/xlsx-bridge.js';
import { nodeFontLoader } from './font-loader.mjs';

setFontLoader(nodeFontLoader);
async function fontNames(bytes) {
  const doc = await PDFDocument.load(bytes);
  const names = new Set();
  for (const [, obj] of doc.context.enumerateIndirectObjects()) {
    const base = obj?.get?.(PDFName.of('BaseFont'));
    if (base) names.add(base.asString().replace(/^\//, '').replace(/^[A-Z]{6}\+/, ''));
  }
  return [...names];
}

async function pageText(bytes) {
  const doc = await pdfjs.getDocument({ data: new Uint8Array(bytes), standardFontDataUrl: new URL('../node_modules/pdfjs-dist/standard_fonts/', import.meta.url).pathname }).promise;
  const out = [];
  for (let i = 1; i <= doc.numPages; i++) out.push((await (await doc.getPage(i)).getTextContent()).items.map((t) => t.str).join(' '));
  return out.join('\n');
}

test('scripts are split into runs and right-to-left text is ordered', () => {
  assert.equal(scriptOf('മ'.codePointAt(0)), 'malayalam');
  assert.equal(scriptOf('合'.codePointAt(0)), 'sc');
  assert.equal(scriptOf('é'.codePointAt(0)), 'std');
  const runs = splitRuns('المجموع: 1,250 ريال (QAR)');
  assert.deepEqual(runs.map((r) => r.text), ['(', 'QAR', 'ريال )', ' ', '1,250', 'المجموع: '], 'right to left, the number kept whole, brackets mirrored');
  assert.deepEqual(splitRuns('Total 合计 120').map((r) => [r.script, r.text]), [['std', 'Total '], ['sc', '合计 120']]);
});

test('sheet PDFs embed fonts for Malayalam, Arabic and Chinese cells', async () => {
  const snap = blankSnapshot('u');
  const sh = snap.sheets[snap.sheetOrder[0]];
  sh.name = 'ഫലം';
  sh.cellData = { 0: { 0: { v: 'പേര്', t: 1 }, 1: { v: 'المدينة', t: 1 }, 2: { v: '城市', t: 1 } }, 1: { 0: { v: 'Ali', t: 1 }, 1: { v: 'الدوحة', t: 1 }, 2: { v: 'Plain', t: 1 } } };
  const pdf = await snapshotToPdf(snap);
  const names = await fontNames(pdf);
  for (const f of ['NotoSansMalayalam', 'NotoSansArabic', 'NotoSansSC', 'Helvetica']) assert.ok(names.some((n) => n.startsWith(f)), `${f} in ${names}`);
  const text = await pageText(pdf);
  assert.match(text, /Ali/);
  assert.match(text, /城市/);
});

test('stamps, watermarks, page numbers and form fields take any script', async () => {
  const src = await readFile(new URL('./fixtures/form.pdf', import.meta.url));
  const doc = await PDFDocument.load(src);
  const field = doc.getForm().getFields().find((f) => f.constructor.name === 'PDFTextField');
  const out = await buildPdf({ a: new Uint8Array(src) }, [{ src: 'a', index: 0, annots: [
    { type: 'stamp', x: 60, y: 600, w: 160, h: 50, label: 'അംഗീകരിച്ചു', sub: '10/11/2026', color: '#16a34a' },
    { type: 'text', x: 60, y: 520, size: 14, text: 'ملاحظة: تمت المراجعة\nसमीक्षा पूर्ण', color: '#111827' },
  ] }], {
    watermark: { text: '机密 CONFIDENTIAL', opacity: 0.2 },
    headerFooter: { footer: ['', 'പേജ് {n} / {total}', ''] },
    formValues: { a: { [field.getName()]: 'ഹാഷിർ' } },
  });
  const names = await fontNames(out);
  for (const f of ['NotoSansMalayalam', 'NotoSansArabic', 'NotoSansDevanagari', 'NotoSansSC']) assert.ok(names.some((n) => n.startsWith(f)), `${f} in ${names}`);
  const back = await PDFDocument.load(out);
  assert.equal(back.getForm().getTextField(field.getName()).getText(), 'ഹാഷിർ');
});

test('fonts chosen in a sheet are embedded in its PDF', async () => {
  const snap = blankSnapshot('f');
  const sh = snap.sheets[snap.sheetOrder[0]];
  snap.styles = { a: { ff: 'Roboto', bl: 1 }, b: { ff: 'Playfair Display', it: 1 }, c: { ff: 'Calibri' }, d: { ff: 'Some Font Nobody Has' } };
  sh.cellData = { 0: { 0: { v: 'Bold Roboto', t: 1, s: 'a' }, 1: { v: 'Italic Playfair', t: 1, s: 'b' }, 2: { v: 'Calibri text', t: 1, s: 'c' }, 3: { v: 'Fallback', t: 1, s: 'd' } } };
  const names = await fontNames(await snapshotToPdf(snap));
  for (const f of ['Roboto-Bold', 'PlayfairDisplay-Italic', 'Carlito', 'Helvetica']) assert.ok(names.some((n) => n.startsWith(f)), `${f} in ${names}`);
});

test('every catalog font package is bundled by the app', async () => {
  const { FONT_CATALOG } = await import('../src/lib/font-catalog.js');
  const src = await readFile(new URL('../src/lib/font-assets.js', import.meta.url), 'utf8');
  const glob = src.match(/@expo-google-fonts\/\{([^}]+)\}/)[1].split(',');
  assert.deepEqual(FONT_CATALOG.map((f) => f.pkg).filter((p) => !glob.includes(p)), []);
});
