import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { PDFDocument, PDFTextField, PDFCheckBox, PDFRadioGroup, PDFDropdown, PDFOptionList, PDFName, PDFRawStream } from '@cantoo/pdf-lib';
import { buildPdf, listFormFields } from '../src/lib/pdf-ops.js';

const fx = (n) => readFile(new URL(`./fixtures/${n}`, import.meta.url));

test('new form fields are saved as real fields', async () => {
  const src = { s0: new Uint8Array(await fx('sample.pdf')) };
  const f = (o) => ({ type: 'field', rot: 0, ...o });
  const annots = [
    f({ kind: 'text', name: 'Full name', x1: 72, y1: 700, x2: 300, y2: 720, required: true }),
    f({ kind: 'multiline', name: 'Comments', x1: 72, y1: 560, x2: 300, y2: 640 }),
    f({ kind: 'checkbox', name: 'Agree', x1: 72, y1: 520, x2: 86, y2: 534 }),
    f({ kind: 'radio', name: 'Size', value: 'Small', x1: 72, y1: 480, x2: 86, y2: 494 }),
    f({ kind: 'radio', name: 'Size', value: 'Large', x1: 120, y1: 480, x2: 134, y2: 494 }),
    f({ kind: 'dropdown', name: 'Country', options: ['Malaysia', 'India', 'Japan'], value: 'India', x1: 72, y1: 440, x2: 220, y2: 460 }),
    f({ kind: 'list', name: 'Colours', options: ['Red', 'Green'], x1: 240, y1: 400, x2: 340, y2: 460 }),
    f({ kind: 'text', name: 'Full name', x1: 72, y1: 380, x2: 300, y2: 400 }),
  ];
  const out = await buildPdf(src, [{ src: 's0', index: 0, rotate: 0, annots }, { src: 's0', index: 1, rotate: 0, annots: [] }]);
  const doc = await PDFDocument.load(out);
  const form = doc.getForm();
  const byName = Object.fromEntries(form.getFields().map((x) => [x.getName(), x]));
  assert.ok(byName['Full name'] instanceof PDFTextField);
  assert.ok(byName['Full name'].isRequired());
  assert.ok(byName['Full name 2'] instanceof PDFTextField, 'a repeated name gets a number');
  assert.ok(byName.Comments.isMultiline());
  assert.ok(byName.Agree instanceof PDFCheckBox);
  assert.ok(byName.Size instanceof PDFRadioGroup);
  assert.deepEqual(byName.Size.getOptions(), ['Small', 'Large']);
  assert.ok(byName.Country instanceof PDFDropdown);
  assert.deepEqual(byName.Country.getOptions(), ['Malaysia', 'India', 'Japan']);
  assert.deepEqual(byName.Country.getSelected(), ['India']);
  assert.ok(byName.Colours instanceof PDFOptionList);
  const w = byName.Comments.acroField.getWidgets()[0].getRectangle();
  [[w.x, 72], [w.y, 560], [w.width, 228], [w.height, 80]].forEach(([got, want]) => assert.ok(Math.abs(got - want) <= 1.5, `${got} vs ${want}`));
  // and the app's own reader lists them, so they can be filled in after reopening
  const listed = await listFormFields(out);
  assert.ok(listed.length >= 7);
});

// A JPEG maker for tests, using Pillow
function pillowJpeg({ data, kind, width, height, channels, maxSide, quality }) {
  const s = Math.min(1, maxSide / Math.max(width, height));
  const w = Math.max(1, Math.round(width * s)); const h = Math.max(1, Math.round(height * s));
  const py = `
import sys, io
from PIL import Image
raw = sys.stdin.buffer.read()
kind, w0, h0, ch, w, h, q = sys.argv[1], int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4]), int(sys.argv[5]), int(sys.argv[6]), int(sys.argv[7])
im = Image.open(io.BytesIO(raw)) if kind == 'jpeg' else Image.frombytes('RGB' if ch == 3 else 'L', (w0, h0), raw[: w0 * h0 * ch])
im = im.convert('RGB').resize((w, h))
b = io.BytesIO(); im.save(b, 'JPEG', quality=q); sys.stdout.buffer.write(b.getvalue())`;
  const bytes = execFileSync('python3', ['-c', py, kind, width, height, channels, w, h, Math.round(quality * 100)].map(String), { input: Buffer.from(data), maxBuffer: 1 << 28 });
  return { bytes: new Uint8Array(bytes), width: w, height: h };
}

test('compressing shrinks pictures and drops unused objects', async () => {
  // a noisy photo-like PNG with transparency, 1600 x 1200
  const png = execFileSync('python3', ['-c', `
import sys, io, random
from PIL import Image
random.seed(3)
im = Image.effect_noise((1600, 1200), 60).convert('RGB')
im.putalpha(200)
b = io.BytesIO(); im.save(b, 'PNG'); sys.stdout.buffer.write(b.getvalue())`], { maxBuffer: 1 << 28 });
  const d = await PDFDocument.create();
  const img = await d.embedPng(new Uint8Array(png));
  const pg = d.addPage([600, 800]);
  pg.drawImage(img, { x: 50, y: 200, width: 500, height: 375 });
  pg.drawText('Hello', { x: 50, y: 700 });
  // an orphan left over from an earlier edit
  d.context.register(d.context.flateStream(new Uint8Array(50000).fill(7)));
  const bytes = await d.save();

  const plain = await buildPdf({ s0: bytes }, [{ src: 's0', index: 0, rotate: 0, annots: [] }]);
  const small = await buildPdf({ s0: bytes }, [{ src: 's0', index: 0, rotate: 0, annots: [] }], { compress: { images: 'strong', recodeImage: pillowJpeg } });
  assert.ok(small.length < plain.length * 0.6, `${small.length} vs ${plain.length}`);
  const back = await PDFDocument.load(small);
  const xo = back.getPage(0).node.Resources().lookup(PDFName.of('XObject'));
  const im = xo.lookup(xo.keys()[0]);
  assert.equal(im.dict.lookup(PDFName.of('Filter')).decodeText(), 'DCTDecode');
  assert.equal(im.dict.lookup(PDFName.of('Width')).asNumber(), 1300);
  assert.ok(im.dict.get(PDFName.of('SMask')), 'transparency is kept');
  const orphans = back.context.enumerateIndirectObjects().filter(([, o]) => o instanceof PDFRawStream && o.contents.length > 40000 && !o.dict.get(PDFName.of('Subtype')));
  assert.equal(orphans.length, 0, 'the unused object is gone');
  // pdftoppm still draws the page
  const { writeFileSync } = await import('node:fs');
  writeFileSync('/tmp/kertas-compressed.pdf', small);
  execFileSync('pdftoppm', ['-r', '20', '-png', '/tmp/kertas-compressed.pdf', '/tmp/kertas-compressed']);
});

test('compressing without pictures still works on a plain document', async () => {
  const src = { s0: new Uint8Array(await fx('sample.pdf')) };
  const out = await buildPdf(src, [{ src: 's0', index: 0, rotate: 0, annots: [] }], { compress: { images: null } });
  const doc = await PDFDocument.load(out);
  assert.equal(doc.getPageCount(), 1);
});
