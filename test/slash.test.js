import { test } from 'node:test';
import assert from 'node:assert/strict';
import { matchScore, filterCommands } from '../src/modules/writer-slash.js';

const CMDS = [
  { id: 'h1', label: 'Heading 1', keys: ['h1'] },
  { id: 'h2', label: 'Heading 2', keys: ['h2', 'subheading'] },
  { id: 'table', label: 'Table', keys: ['grid'] },
  { id: 'toc', label: 'Table of contents', keys: ['toc', 'contents'] },
  { id: 'bullets', label: 'Bulleted list', keys: ['ul', 'bullet'] },
  { id: 'date', label: 'Date', keys: ['today'] },
];
const ids = (q, recent) => filterCommands(CMDS, q, recent).map((c) => c.id);

test('slash commands are found by name, keyword or initials', () => {
  assert.deepEqual(ids('h2'), ['h2']);
  assert.deepEqual(ids('tab'), ['table', 'toc'], 'shorter label first on a prefix tie');
  assert.equal(ids('toc')[0], 'toc');
  assert.equal(ids('bul')[0], 'bullets');
  assert.equal(ids('today')[0], 'date');
  assert.equal(ids('contents')[0], 'toc');
  assert.ok(ids('tbl').includes('table'), 'letters in order');
  assert.deepEqual(ids('xyz'), []);
  assert.equal(matchScore(CMDS[0], ''), 1);
});

test('with no query, recently used commands come first', () => {
  const list = filterCommands(CMDS, '', ['date', 'toc']);
  assert.deepEqual(list.slice(0, 2).map((c) => [c.id, c.recent]), [['date', true], ['toc', true]]);
  assert.equal(list.length, CMDS.length, 'nothing listed twice');
});
