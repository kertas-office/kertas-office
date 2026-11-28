// Text recognised on a scan is often stored smaller than the type on the picture; its size
// is corrected from the line spacing, after the layout is read
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { analyse, scanScales } from '../src/lib/pdf-layout.js';

function page({ scan, size, pitch, lines = 12 }) {
  const items = [];
  for (let l = 0; l < lines; l++) {
    const y = 100 + l * pitch + (l >= 6 ? pitch : 0); // a blank line after six lines: two paragraphs
    for (let w = 0; w < 8; w++) items.push({ str: 'word', x: 72 + w * 40, y, w: 30, size, bold: false, italic: false, font: null });
  }
  return { width: 595, height: 842, items, scan };
}

test('a scanned page with text stored at 8 pt on 14 pt spacing comes out near 11.5 pt', () => {
  const [k] = scanScales([page({ scan: true, size: 8, pitch: 14 })]);
  assert.ok(Math.abs(8 * k - 14 / 1.2) < 0.01);
  const doc = analyse([page({ scan: true, size: 8, pitch: 14 })]);
  assert.equal(doc.bodySize, 11.5);
  const paras = doc.pages[0].blocks.filter((b) => b.type === 'para');
  assert.equal(paras.length, 2, 'the paragraph break is still found');
  assert.ok(paras.every((p) => p.runs.every((r) => r.size === 11.5)));
});

test('typeset pages, and scans whose sizes already fit their spacing, keep their sizes', () => {
  assert.deepEqual(scanScales([page({ scan: false, size: 8, pitch: 14 }), page({ scan: true, size: 11, pitch: 13.2 })]), [1, 1]);
  assert.equal(analyse([page({ scan: false, size: 11, pitch: 13.2 })]).bodySize, 11);
});
