import { test } from 'node:test';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { xlsxToSnapshot, snapshotToXlsx, blankSnapshot } from '../src/lib/xlsx-bridge.js';
import { COMMENT_RES, toBody, fromBody, commentsOf } from '../src/lib/xlsx-comments.js';

function withThreads() {
  const snap = blankSnapshot('t');
  const sid = snap.sheetOrder[0];
  const sh = snap.sheets[sid];
  sh.name = 'Budget';
  sh.cellData[1] = { 1: { v: 120, t: 2 }, 2: { v: 'Rent', t: 1 } };
  const threads = [
    {
      id: 'a1', threadId: 'a1', ref: 'B2', dT: '2026-11-13T09:00:00.00', personId: 'p1', authorName: 'Hashir', text: toBody('Is this the new rent?\nPlease check.'), resolved: false,
      children: [{ id: 'a2', threadId: 'a1', ref: 'B2', dT: '2026-11-13T10:30:00.00', personId: 'p2', authorName: 'Aisha & Co', text: toBody('Yes, from December <confirmed>') }],
    },
    { id: 'b1', threadId: 'b1', ref: 'C2', dT: '2026-11-13T11:00:00.00', personId: 'p1', authorName: 'Hashir', text: toBody('Done'), resolved: true, children: [] },
  ];
  snap.resources = [...(snap.resources || []), { name: COMMENT_RES, data: JSON.stringify({ [sid]: threads }) }];
  // a cell with a note of its own and a conversation keeps the note
  snap.resources.push({ name: 'SHEET_NOTE_PLUGIN', data: JSON.stringify({ [sid]: { 1: { 2: { id: 'n', row: 1, col: 2, note: 'My own note', width: 220, height: 92, show: false } } } }) });
  return snap;
}

test('conversations are saved as Excel threaded comments, with notes for older readers', async () => {
  const bytes = await snapshotToXlsx(withThreads());
  const zip = await JSZip.loadAsync(bytes);
  const tc = await zip.file('xl/threadedComments/threadedComment1.xml')?.async('string');
  assert.ok(tc, 'threaded comments part written');
  const comments = [...tc.matchAll(/<threadedComment\b([^>]*)>/g)].map((m) => m[1]);
  assert.equal(comments.length, 3);
  const root = comments.find((a) => /ref="B2"/.test(a) && !/parentId/.test(a));
  const reply = comments.find((a) => /parentId=/.test(a));
  assert.match(root, /id="\{[0-9A-F-]{36}\}"/, 'Excel ids are GUIDs');
  assert.match(reply, new RegExp(`parentId="${root.match(/id="(\{[^"]+\})"/)[1].replace(/[{}]/g, '\\$&')}"`));
  assert.match(comments.find((a) => /ref="C2"/.test(a)), /done="1"/, 'a resolved conversation is marked done');
  assert.match(tc, /Aisha &amp; Co|&lt;confirmed&gt;/, 'text is escaped');
  const persons = await zip.file('xl/persons/person.xml').async('string');
  assert.match(persons, /displayName="Hashir"/);
  assert.match(persons, /displayName="Aisha &amp; Co"/);
  const ct = await zip.file('[Content_Types].xml').async('string');
  assert.match(ct, /application\/vnd\.ms-excel\.threadedcomments\+xml/);
  assert.match(ct, /application\/vnd\.ms-excel\.person\+xml/);
  assert.match(await zip.file('xl/_rels/workbook.xml.rels').async('string'), /relationships\/person/);
  const sheetRels = await zip.file('xl/worksheets/_rels/sheet1.xml.rels').async('string');
  assert.match(sheetRels, /relationships\/threadedComment/);
  // the notes: B2 holds the conversation with author tc={id}; C2 keeps the user's own note
  const commentsPath = Object.keys(zip.files).find((n) => /^xl\/comments\d*\.xml$/.test(n));
  const legacy = await zip.file(commentsPath).async('string');
  const authors = [...legacy.matchAll(/<author>([^<]*)<\/author>/g)].map((m) => m[1]);
  const b2 = legacy.match(/<comment\b[^>]*ref="B2"[^>]*>/)[0];
  assert.match(authors[+b2.match(/authorId="(\d+)"/)[1]], /^tc=\{/);
  assert.match(legacy, /Hashir: Is this the new rent\?/);
  const c2 = legacy.match(/<comment\b[^>]*ref="C2"[^>]*>([\s\S]*?)<\/comment>/);
  assert.doesNotMatch(authors[+c2[0].match(/authorId="(\d+)"/)[1]], /^tc=/, 'own note is not turned into the conversation');
  assert.match(c2[1], /My own note/);
});

test('conversations come back when the file is opened again', async () => {
  const bytes = await snapshotToXlsx(withThreads());
  const snap = await xlsxToSnapshot(bytes, 'b');
  const sid = snap.sheetOrder[0];
  const threads = commentsOf(snap)[sid];
  assert.equal(threads.length, 2);
  const t = threads.find((x) => x.ref === 'B2');
  assert.equal(t.authorName, 'Hashir');
  assert.equal(fromBody(t.text), 'Is this the new rent?\nPlease check.');
  assert.equal(t.children.length, 1);
  assert.equal(t.children[0].authorName, 'Aisha & Co');
  assert.equal(fromBody(t.children[0].text), 'Yes, from December <confirmed>');
  assert.equal(t.children[0].threadId, t.threadId);
  assert.equal(threads.find((x) => x.ref === 'C2').resolved, true);
  // B2's stand-in note is not shown as a second note; C2's own note stays
  const notes = JSON.parse(snap.resources.find((r) => r.name === 'SHEET_NOTE_PLUGIN')?.data || '{}')[sid] || {};
  assert.equal(notes[1]?.[1], undefined);
  assert.equal(notes[1]?.[2]?.note, 'My own note');
});
