import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { Document, Packer, Paragraph } from 'docx';
import { readRecipients, fillFields, fieldsInDocx, mergeSingle, mergeSeparate, mergeFieldName, fieldName, recordLabel } from '../src/lib/mailmerge.js';
import { latexToOmml, mathmlToOmml, EQUATION_PRESETS } from '../src/lib/omml.js';

const complex = (name, shown) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> MERGEFIELD ${name} \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>«${shown}»</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
async function letter() {
  const zip = await JSZip.loadAsync(await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('BODY')] }] })));
  let xml = await zip.file('word/document.xml').async('string');
  const body = `<w:p><w:r><w:t xml:space="preserve">Dear </w:t></w:r>${complex('FirstName', 'FirstName')}<w:r><w:t>,</w:t></w:r></w:p>`
    + `<w:p><w:r><w:t xml:space="preserve">Your balance is </w:t></w:r><w:fldSimple w:instr=" MERGEFIELD &quot;Amount Due&quot; "><w:r><w:rPr><w:i/></w:rPr><w:t>«Amount Due»</w:t></w:r></w:fldSimple></w:p>`
    + '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';
  xml = xml.replace(/<w:p>(?:(?!<w:p>).)*?BODY.*?<\/w:p>/s, body);
  zip.file('word/document.xml', xml);
  return zip.generateAsync({ type: 'uint8array' });
}
const text = (xml) => [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');

test('recipients are read from CSV with tidy, unique headings', () => {
  const csv = new TextEncoder().encode('First Name,Amount Due,First Name\nAminah,"RM 1,200.00",x\n,,\nRavi,RM 80,y\n');
  const r = readRecipients(csv);
  assert.deepEqual(r.columns, ['First Name', 'Amount Due', 'First Name 2']);
  assert.equal(r.rows.length, 2);
  assert.equal(r.rows[0]['Amount Due'], 'RM 1,200.00');
  assert.equal(fieldName('Amount Due'), 'Amount_Due');
  assert.equal(mergeFieldName(' MERGEFIELD "Amount Due" \\* MERGEFORMAT'), 'Amount Due');
  assert.equal(recordLabel({ 'First Name': 'Aminah', 'Last Name': 'Yusof' }, ['First Name', 'Last Name'], 0), 'Aminah Yusof');
});

test('merge fields are filled in place and other fields are left alone', async () => {
  const docx = await letter();
  assert.deepEqual(await fieldsInDocx(docx), ['FirstName', 'Amount Due']);
  const xml = fillFields(await (await JSZip.loadAsync(docx)).file('word/document.xml').async('string'), (n) => ({ FirstName: 'Aminah', 'Amount Due': 'RM 1,200' }[n]));
  assert.match(text(xml), /Dear Aminah,Your balance is RM 1,200/);
  assert.match(xml, /<w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve">Aminah<\/w:t><\/w:r>/, 'keeps the field result formatting');
  assert.match(xml, /<w:instrText> PAGE <\/w:instrText>/, 'PAGE field untouched');
  assert.doesNotMatch(xml, /MERGEFIELD/);
  // Words that slipped into a field's result after the placeholder are kept
  const slipped = `<w:p>${complex('Amount', 'Amount').replace('<w:t>«Amount»</w:t>', '<w:t xml:space="preserve">«Amount» is due.</w:t>')}\u200B</w:p>`;
  assert.equal(text(fillFields(slipped, () => 'RM 80')), 'RM 80 is due.');
});

test('a single merged document has one copy per record, on its own page', async () => {
  const rows = [{ FirstName: 'Aminah', 'Amount Due': 'RM 1,200' }, { FirstName: 'Ravi', 'Amount Due': 'RM 80' }, { FirstName: 'Mei <Ling>', 'Amount Due': 'RM 5' }];
  const merged = await JSZip.loadAsync(await mergeSingle(await letter(), rows));
  const xml = await merged.file('word/document.xml').async('string');
  const t = text(xml);
  assert.match(t, /Dear Aminah,.*Dear Ravi,.*Dear Mei &lt;Ling&gt;,/s);
  assert.equal((xml.match(/w:type="page"/g) || []).length, 2);
  assert.equal((xml.match(/<w:sectPr/g) || []).length, 1);
  const separate = await mergeSeparate(await letter(), rows);
  assert.equal(separate.length, 3);
  assert.match(text(await (await JSZip.loadAsync(separate[1].data)).file('word/document.xml').async('string')), /Dear Ravi,Your balance is RM 80/);
});

test('LaTeX becomes Office Math', () => {
  const { omath, paragraph } = latexToOmml('x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}');
  assert.match(paragraph, /^<w:p><m:oMathPara><m:oMath>/);
  assert.match(omath, /<m:f><m:num>.*<m:rad><m:radPr><m:degHide m:val="1"\/><\/m:radPr><m:deg\/><m:e><m:sSup>/);
  const sum = latexToOmml('\\sum_{i=1}^{n} i').omath;
  assert.match(sum, /<m:nary><m:naryPr><m:chr m:val="∑"\/><m:limLoc m:val="undOvr"\/><\/m:naryPr><m:sub>.*<\/m:sub><m:sup>.*<\/m:sup><m:e>.*>i<\/m:t>/);
  assert.match(latexToOmml('\\left(1+x\\right)').omath, /<m:d><m:dPr><m:begChr m:val="\("\/><m:endChr m:val="\)"\/><\/m:dPr>/);
  assert.match(latexToOmml('\\hat{x}').omath, /<m:acc><m:accPr><m:chr m:val="̂"\/>/);
  assert.match(latexToOmml('\\begin{pmatrix}1&2\\\\3&4\\end{pmatrix}').omath, /<m:m>.*<m:count m:val="2"\/>.*<m:mr>.*<m:mr>/);
  assert.throws(() => latexToOmml('\\frac{1'), /expected/);
  for (const p of EQUATION_PRESETS) assert.ok(latexToOmml(p.tex).omath.length > 50, p.label);
  assert.match(mathmlToOmml('<math><mi>x</mi></math>'), /<m:oMath><m:r>.*x<\/m:t><\/m:r><\/m:oMath>/);
});

test('citations and references follow each style', async () => {
  const { inText, entryText, parseAuthors, bibliographyXml, numbering } = await import('../src/lib/citations.js');
  const book = { sourceId: 'a', type: 'book', fields: { title: 'Deep work: rules for focused success', authors: parseAuthors('Cal Newport'), year: '2016', publisher: 'Grand Central', city: 'New York' } };
  const art = { sourceId: 'b', type: 'journalArticle', fields: { title: 'The restorative benefits of nature', authors: parseAuthors('Kaplan, Stephen; Kaplan, Rachel'), year: '1995', journalName: 'Journal of Environmental Psychology', volume: '15', issue: '3', pages: '169–182' } };
  assert.equal(inText([book, art], 'apa'), '(Newport, 2016; Kaplan & Kaplan, 1995)');
  assert.equal(inText([book], 'mla', { page: '12' }), '(Newport 12)');
  assert.equal(inText([art], 'chicago'), '(Kaplan and Kaplan 1995)');
  assert.equal(inText([book, art], 'ieee', { numbers: numbering([book, art], ['b', 'a']) }), '[2], [1]');
  assert.equal(entryText(art, 'apa'), 'Kaplan, S., & Kaplan, R. (1995). The restorative benefits of nature. Journal of Environmental Psychology, 15(3), 169-182.');
  assert.equal(entryText(book, 'mla'), 'Newport, Cal. Deep Work: Rules for Focused Success. Grand Central, 2016.');
  assert.equal(entryText(book, 'harvard'), 'Newport, C. (2016) Deep work: Rules for focused success. New York: Grand Central.');
  assert.match(entryText(art, 'ieee', 1), /^\[1\]\tS\. Kaplan and R\. Kaplan, "The restorative benefits of nature," Journal of Environmental Psychology, vol\. 15, no\. 3, pp\. 169-182, 1995\.$/);
  const xml = bibliographyXml([book, art], 'mla');
  assert.match(xml, /<w:t>Works Cited<\/w:t>/);
  assert.ok(xml.indexOf('Kaplan') < xml.indexOf('Newport'), 'alphabetical');
  assert.match(xml, /<w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">Journal of Environmental Psychology<\/w:t>/);
});
