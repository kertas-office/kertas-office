// Cross-references: which REF and PAGEREF fields Kertas works out itself, and which keep Word's text
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { parseInstruction, canResolve } from '../src/modules/writer-references.js';

test('instructions are read with switch arguments skipped and quotes removed', () => {
  assert.deepEqual(parseInstruction('REF _Ref123 \\h \\* MERGEFORMAT'), { field: 'REF', name: '_Ref123', switches: ['\\h', '\\*'] });
  assert.deepEqual(parseInstruction(' PAGEREF "Budget" \\h '), { field: 'PAGEREF', name: 'Budget', switches: ['\\h'] });
  assert.deepEqual(parseInstruction('ref _Ref1 \\R \\H'), { field: 'REF', name: '_Ref1', switches: ['\\r', '\\h'] });
});

test('references Kertas inserts can be worked out', () => {
  for (const instruction of ['REF _Ref1 \\h', 'PAGEREF _Ref1 \\h', 'REF _Ref1 \\p \\h', 'REF Council', 'REF _Ref1 \\h \\* MERGEFORMAT']) {
    assert.equal(canResolve({ instruction }), true, instruction);
  }
});

test('heading numbers and other switches keep the text Word saved', () => {
  for (const instruction of ['REF _Ref1 \\r \\h', 'REF _Ref1 \\n', 'REF _Ref1 \\w \\h', 'REF _Ref1 \\f', 'REF _Ref1 \\d "-"', 'PAGEREF _Ref1 \\p \\h', 'SEQ Table', 'NOTEREF _Ref1']) {
    assert.equal(canResolve({ instruction }), false, instruction);
  }
});
