import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import nspell from 'nspell';
import { wordsIn, dictionaryFor } from '../src/lib/spell/words.js';


test('words are found, skipping links, emails, acronyms and code', () => {
  const text = "Teh report's figures (see https://kertas.app/docs or me@mail.com) for UPSC and iPhone users, v2.3 file_name.txt colour.";
  assert.deepEqual(wordsIn(text).map((w) => w.word), ['Teh', "report's", 'figures', 'see', 'or', 'for', 'and', 'users', 'colour']);
  const w = wordsIn('Hello wrold')[1];
  assert.deepEqual([w.start, w.end], [6, 11]);
  assert.equal(dictionaryFor('en-MY'), 'en-GB');
  assert.equal(dictionaryFor('en'), 'en-US');
  assert.equal(dictionaryFor('ml-IN'), null);
});

test('Hunspell dictionaries flag misspellings and suggest fixes', async () => {
  const dir = (pkg) => new URL(`../node_modules/${pkg}/`, import.meta.url);
  const load = async (pkg) => nspell(await readFile(new URL('index.aff', dir(pkg)), 'utf8'), await readFile(new URL('index.dic', dir(pkg)), 'utf8'));
  const us = await load('dictionary-en');
  assert.equal(us.correct('receive'), true);
  assert.equal(us.correct('recieve'), false);
  assert.ok(us.suggest('recieve').includes('receive'));
  assert.equal(us.correct('colour'), false);
  const gb = await load('dictionary-en-gb');
  assert.equal(gb.correct('colour'), true);
});
