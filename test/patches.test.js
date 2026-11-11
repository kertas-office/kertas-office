// Every fix in scripts/patch-deps.cjs must be present in the installed packages.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve('node_modules');
const script = fs.readFileSync('scripts/patch-deps.cjs', 'utf8');
const markers = [...script.matchAll(/marker: '([^']+)'/g)].map((m) => m[1]);

function filesUnder(dir) {
  const d = path.join(root, dir);
  return fs.existsSync(d) ? fs.readdirSync(d).filter((f) => /\.m?js$/.test(f)).map((f) => path.join(d, f)) : [];
}
const listed = [...script.matchAll(/files: \[([^\]]*)\]/g)].flatMap((m) => [...m[1].matchAll(/'([^']+)'/g)].map((x) => path.join(root, x[1]))).filter((f) => fs.existsSync(f));
const sources = [...new Set([...filesUnder('superdoc/dist/chunks'), ...filesUnder('fontkit/dist'), ...filesUnder('@univerjs/engine-formula/lib/es'), ...listed])].map((f) => fs.readFileSync(f, 'utf8')).join('\n');

test('patch script lists its fixes', () => {
  assert.ok(markers.length >= 10, `found ${markers.length} fixes`);
});
for (const m of markers) {
  test(`fix applied: ${m}`, () => assert.ok(sources.includes(m), `"${m}" is missing: run npm install, or the package changed`));
}
test('no closed-source SuperDoc engine is installed', () => {
  assert.equal(fs.existsSync(path.join(root, '@superdoc/docx-engine')), false);
  assert.equal(JSON.parse(fs.readFileSync(path.join(root, 'superdoc/package.json'), 'utf8')).version, '1.47.2');
});
