// Generates the PDF fixtures used by the browser checks
import { writeFile, readFile } from 'node:fs/promises';
import { PDFDocument, StandardFonts } from '@cantoo/pdf-lib';
import { buildPdf } from '../src/lib/pdf-ops.js';

const dir = new URL('./fixtures/', import.meta.url);
const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
const bold = await doc.embedFont(StandardFonts.HelveticaBold);
const page = doc.addPage([595, 842]);
const form = doc.getForm();
page.drawText('Membership application', { x: 50, y: 780, size: 22, font: bold });
const label = (t, y) => page.drawText(t, { x: 50, y, size: 11, font });
label('Full name', 718); form.createTextField('name').addToPage(page, { x: 160, y: 710, width: 300, height: 22 });
label('Email', 678); form.createTextField('email').addToPage(page, { x: 160, y: 670, width: 300, height: 22 });
label('Country', 638);
const c = form.createDropdown('country'); c.addOptions(['Malaysia', 'Qatar', 'India', 'Singapore']); c.addToPage(page, { x: 160, y: 630, width: 200, height: 22 });
label('Plan', 598);
const plan = form.createRadioGroup('plan');
plan.addOptionToPage('Basic', page, { x: 160, y: 596, width: 14, height: 14 }); page.drawText('Basic', { x: 180, y: 598, size: 11, font });
plan.addOptionToPage('Pro', page, { x: 240, y: 596, width: 14, height: 14 }); page.drawText('Pro', { x: 260, y: 598, size: 11, font });
form.createCheckBox('terms').addToPage(page, { x: 160, y: 556, width: 14, height: 14 }); page.drawText('I agree to the terms', { x: 180, y: 558, size: 11, font });
label('Notes', 518);
const notes = form.createTextField('notes'); notes.enableMultiline(); notes.addToPage(page, { x: 160, y: 430, width: 300, height: 90 });
await writeFile(new URL('form.pdf', dir), await doc.save());

const sample = new Uint8Array(await readFile(new URL('sample.pdf', dir)));
const m = [0, 1, 2].map((i) => ({ src: 's0', index: i, rotate: 0, annots: [] }));
await writeFile(new URL('outlined.pdf', dir), await buildPdf({ s0: sample }, m, { outline: [
  { title: 'Introduction', page: 0, children: [{ title: 'Invoice details', page: 1, children: [] }] },
  { title: 'Appendix', page: 2, children: [] },
] }));
await writeFile(new URL('locked.pdf', dir), await buildPdf({ s0: sample }, m, { encrypt: { userPassword: 'open123' } }));
console.log('fixtures written');
