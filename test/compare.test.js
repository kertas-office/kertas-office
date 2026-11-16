// Compare documents: how paragraphs are lined up and how words inside them are compared
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { planChanges, wordOps } from '../src/modules/writer-compare.js';

const paras = (...texts) => texts.map((text) => ({ text }));
const apply = (a, ops) => { let out = a; for (const op of [...ops].reverse()) out = out.slice(0, op.at) + op.ins + out.slice(op.at + op.del); return out; };

test('word changes inside a paragraph turn it into the revised text', () => {
  const a = 'The library opened two new branches and visits rose by fourteen percent.';
  const b = 'The library opened three new branches and visits rose by eighteen percent.';
  const ops = wordOps(a, b);
  assert.equal(apply(a, ops), b);
  assert.deepEqual(ops.map((o) => [a.slice(o.at, o.at + o.del), o.ins]), [['two', 'three'], ['fourteen', 'eighteen']]);
});

test('numbers with separators change as one piece', () => {
  const ops = wordOps('Visits 361,500 in Q1', 'Visits 412,300 in Q1');
  assert.deepEqual(ops.map((o) => o.ins), ['412,300']);
});

test('unchanged paragraphs are left alone; removed, added and edited ones are found', () => {
  const plan = planChanges(
    paras('Title', 'Intro has two words changed here.', 'This one goes.', 'Same.', 'End.'),
    paras('New first line.', 'Title', 'Intro has three words changed here.', 'Same.', 'Added in the middle.', 'End.'),
  );
  assert.deepEqual(plan.map((c) => c.kind), ['insert', 'words', 'delete', 'insert']);
  assert.equal(plan[0].text, 'New first line.');
  assert.equal(plan[0].after, -1);
  assert.equal(plan[2].para, 2);
  assert.deepEqual([plan[3].text, plan[3].after], ['Added in the middle.', 3]);
});

test('a paragraph replaced by an unlike one in the same place is shown as a replacement', () => {
  const plan = planChanges(paras('Q1', '361,500', 'End.'), paras('Q1', '412,300', 'End.'));
  assert.deepEqual(plan.map((c) => c.kind), ['words']);
});

test('identical documents have no changes', () => {
  assert.deepEqual(planChanges(paras('A b c.', 'D e.'), paras('A b c.', 'D e.')), []);
});
