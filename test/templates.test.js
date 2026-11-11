// The ready-made templates offered on the Home screen: every one is present, opens, and holds
// nothing Kertas would warn about losing when it is saved.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { xlsxToSnapshot } from '../src/lib/xlsx-bridge.js';
import { detectFeatures, atRisk } from '../src/lib/fidelity.js';

const src = fs.readFileSync('src/lib/templates.js', 'utf8');
const list = [...src.matchAll(/\{ id: '([^']+)', kind: '(writer|sheets)', name: '([^']+)'/g)].map((m) => ({ id: m[1], kind: m[2], name: m[3] }));
const file = (t) => `src/templates/${t.id}.${t.kind === 'writer' ? 'docx' : 'xlsx'}`;

test('eight templates each for Writer and Sheets', () => {
  assert.equal(list.filter((t) => t.kind === 'writer').length, 8);
  assert.equal(list.filter((t) => t.kind === 'sheets').length, 8);
});

for (const t of list) {
  test(`template ${t.id}: file and preview are there`, () => {
    assert.ok(fs.existsSync(file(t)), file(t));
    assert.ok(fs.existsSync(`src/templates/${t.id}.jpg`), `${t.id}.jpg`);
  });
  test(`template ${t.id}: nothing would be lost on save`, async () => {
    const bytes = new Uint8Array(fs.readFileSync(file(t)));
    assert.deepEqual(atRisk(await detectFeatures(bytes, t.kind), t.kind), []);
  });
  if (t.kind === 'sheets') {
    test(`template ${t.id}: opens with its formulas`, async () => {
      const snap = await xlsxToSnapshot(new Uint8Array(fs.readFileSync(file(t))), t.name);
      const formulas = Object.values(snap.sheets).reduce((n, sh) => n + Object.values(sh.cellData || {}).reduce((k, row) => k + Object.values(row).filter((c) => c?.f).length, 0), 0);
      assert.ok(formulas > 0, 'has formulas');
    });
  }
}
