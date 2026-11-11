// Citations: plain text in the editor, Word CITATION fields in the saved file.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { textToFields, fieldsToText, citationInstr, parseInstr } from '../src/lib/docx-citations.js';

const doc = (body) => `<w:document xmlns:w="w"><w:body>${body}</w:body></w:document>`;

test('a registered citation inside a run becomes a field, with the text around it kept', () => {
  const src = doc('<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">As shown (Newport, 2016, p. 4) and later (Kaplan, 1995).</w:t></w:r></w:p>');
  const entries = [
    { text: '(Newport, 2016, p. 4)', sourceIds: ['New16'], page: '4' },
    { text: '(Kaplan, 1995)', sourceIds: ['Kap95'] },
  ];
  const { xml, count } = textToFields(src, entries);
  assert.equal(count, 2);
  assert.match(xml, /<w:instrText xml:space="preserve"> CITATION New16 \\p 4 \\l 1033 <\/w:instrText>/);
  assert.match(xml, /CITATION Kap95 \\l 1033/);
  assert.equal((xml.match(/<w:rPr><w:i\/><\/w:rPr>/g) || []).length, 3 + 5 * 2, 'every new run keeps the formatting');
  // and back again
  const back = fieldsToText(xml);
  assert.equal(back.found.length, 2);
  assert.deepEqual(back.found.map((f) => [f.text, f.sourceIds, f.page]), [['(Newport, 2016, p. 4)', ['New16'], '4'], ['(Kaplan, 1995)', ['Kap95'], '']]);
  const text = [...back.xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
  assert.equal(text, 'As shown (Newport, 2016, p. 4) and later (Kaplan, 1995).');
  assert.doesNotMatch(back.xml, /fldChar|instrText/);
});

test('Word citation content controls and several sources', () => {
  const src = doc(`<w:p><w:sdt><w:sdtPr><w:id w:val="1"/><w:citation/></w:sdtPr><w:sdtContent>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> CITATION Smi20 \\m Doe19 \\l 1033 </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>(Smith, 2020; Doe, 2019)</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
  </w:sdtContent></w:sdt></w:p>`);
  const r = fieldsToText(src);
  assert.deepEqual(r.found[0].sourceIds, ['Smi20', 'Doe19']);
  assert.doesNotMatch(r.xml, /w:sdt|fldChar/);
  assert.match(r.xml, /\(Smith, 2020; Doe, 2019\)/);
});

test('other fields are left alone', () => {
  const src = doc('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>');
  assert.equal(fieldsToText(src).xml, src);
  assert.deepEqual(parseInstr(citationInstr(['A', 'B'], '12')), { sourceIds: ['A', 'B'], page: '12' });
});
