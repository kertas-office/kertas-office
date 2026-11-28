#!/usr/bin/env node
// Builds the new editors into electron/writer2/: the new Writer, and the Sheets, PDF and
// Slides editors beside it (electron/writer2/sheets, /pdf and /slides).
//
// The editors come from GenOffice (Apache-2.0, Copyright Mainfunc, Inc.), taken at
// a fixed commit and changed by scripts/writer2/genoffice.patch: Kertas hosts them
// as tabs of its own window, the AI assistant is left out, product names are
// Kertas's, and they keep their files in a folder of their own. Only the Apache-2.0
// parts are used; the separately licensed ee/ directory is deleted before
// anything is built.
//
//   node scripts/build-writer2.cjs            fetch, patch, build, copy
//   KERTAS_GENOFFICE_DIR=<checkout>           use a checkout that is already patched
//   KERTAS_SIDECAR=<file>                     a ready-built spreadsheet helper to ship
//                                             (when packaging for another system)
//
// The new Sheets reads and writes workbooks through a small helper program written in
// Rust (apps/sheets/native/xlsx-engine). It is built here with cargo when cargo is
// installed; without it the new Sheets is simply not offered.
//
// Without this step Kertas builds and runs as before, with none of the new editors.
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const REPO = 'https://github.com/genspark-ai/genoffice.git';
const COMMIT = '99865c27f685f1e7f40d6df273545db8ea86aa17';
const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'electron', 'writer2');
const PATCH = path.join(__dirname, 'writer2', 'genoffice.patch');
const given = process.env.KERTAS_GENOFFICE_DIR;
const SRC = given || path.join(ROOT, '.cache', 'genoffice');

const run = (cmd, args, cwd) => execFileSync(cmd, args, { cwd, stdio: 'inherit', shell: process.platform === 'win32' });

if (!given) {
  const stamp = path.join(SRC, '.kertas-patched');
  const want = `${COMMIT}\n${fs.readFileSync(PATCH, 'utf8').length}`;
  if (!fs.existsSync(stamp) || fs.readFileSync(stamp, 'utf8') !== want) {
    fs.rmSync(SRC, { recursive: true, force: true });
    fs.mkdirSync(SRC, { recursive: true });
    run('git', ['init', '-q'], SRC);
    // keep the source's own line endings: on Windows git would turn them into CRLF
    // and the patch (LF) would no longer apply
    run('git', ['config', 'core.autocrlf', 'false'], SRC);
    run('git', ['config', 'core.eol', 'lf'], SRC);
    run('git', ['remote', 'add', 'origin', REPO], SRC);
    run('git', ['fetch', '-q', '--depth', '1', 'origin', COMMIT], SRC);
    run('git', ['checkout', '-q', 'FETCH_HEAD'], SRC);
    fs.rmSync(path.join(SRC, 'ee'), { recursive: true, force: true });
    run('git', ['apply', '--whitespace=nowarn', PATCH], SRC);
    run('npm', ['ci', '--ignore-scripts', '--no-audit', '--no-fund'], SRC);
    fs.writeFileSync(stamp, want);
  }
}

const app = path.join(SRC, 'apps', 'docs');
// Kertas's own interface font, for the editor's Kertas look (apps/docs/src/renderer/kertas/skin.css)
const inter = path.join(ROOT, 'node_modules', '@fontsource-variable', 'inter', 'files');
fs.mkdirSync(path.join(app, 'src', 'renderer', 'kertas'), { recursive: true });
for (const f of ['inter-latin-wght-normal.woff2', 'inter-latin-ext-wght-normal.woff2']) fs.copyFileSync(path.join(inter, f), path.join(app, 'src', 'renderer', 'kertas', f));
run('npx', ['electron-vite', 'build', '--config', 'kertas.vite.config.ts'], app);
try { run('node', ['tools/gen-third-party-notices.mjs'], SRC); } catch { console.warn('Third-party notices were not regenerated'); }

// The spreadsheet, PDF and presentation editors: their pages and bridges (their main-process code is in the Writer's bundle)
for (const name of ['sheets', 'pdf', 'slides']) run('npx', ['electron-vite', 'build', '--config', 'kertas.vite.config.ts'], path.join(SRC, 'apps', name));

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'wasm'), { recursive: true });
for (const part of ['main', 'preload', 'renderer']) fs.cpSync(path.join(app, 'out-kertas', part), path.join(OUT, part), { recursive: true });
for (const name of ['sheets', 'pdf', 'slides']) {
  for (const part of ['preload', 'renderer']) fs.cpSync(path.join(SRC, 'apps', name, 'out-kertas', part), path.join(OUT, name, part), { recursive: true });
}
// The spreadsheet helper, for the system this build is for
try {
  const engine = path.join(SRC, 'apps', 'sheets', 'native', 'xlsx-engine');
  let helper = process.env.KERTAS_SIDECAR;
  if (!helper) {
    run('cargo', ['build', '--release', '--locked'], engine);
    helper = path.join(engine, 'target', 'release', process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar');
  }
  fs.mkdirSync(path.join(OUT, 'native'), { recursive: true });
  fs.copyFileSync(helper, path.join(OUT, 'native', path.basename(helper)));
  if (process.platform !== 'win32') fs.chmodSync(path.join(OUT, 'native', path.basename(helper)), 0o755);
  fs.copyFileSync(path.join(engine, 'Cargo.lock'), path.join(OUT, 'native', 'Cargo.lock')); // the exact libraries inside it
} catch (e) {
  console.warn('The spreadsheet helper was not built, so the new Sheets is left out:', e.message);
}
fs.copyFileSync(path.join(SRC, 'node_modules', '@embedpdf', 'pdfium', 'dist', 'pdfium.wasm'), path.join(OUT, 'wasm', 'pdfium.wasm'));
fs.copyFileSync(path.join(SRC, 'node_modules', 'harfbuzzjs', 'dist', 'harfbuzz-subset.wasm'), path.join(OUT, 'wasm', 'hb-subset.wasm'));
// Spelling dictionaries, so the checker never has to fetch one: Electron's own set for the version
// Kertas ships, of which the languages Kertas offers are kept (electron/writer2.cjs serves them
// to the checker from this computer). Skipped quietly when offline; spelling then uses whatever the
// system provides, and on Linux stays off.
try {
  const version = require(path.join(ROOT, 'node_modules', 'electron', 'package.json')).version;
  const cache = path.join(ROOT, '.cache', `hunspell-${version}`);
  const zip = path.join(cache, 'hunspell_dictionaries.zip');
  if (!fs.existsSync(zip)) {
    fs.mkdirSync(cache, { recursive: true });
    run('curl', ['-fsSL', '-o', zip, `https://github.com/electron/electron/releases/download/v${version}/hunspell_dictionaries.zip`]);
  }
  const dict = path.join(OUT, 'dict');
  fs.mkdirSync(dict, { recursive: true });
  run('unzip', ['-o', '-q', zip, 'en-US-*.bdic', 'en-GB-[0-9]*.bdic', 'de-DE-*.bdic', 'es-ES-*.bdic', 'fr-FR-*.bdic', 'id-ID-*.bdic', 'LICENSE*', '-d', dict]);
} catch (e) {
  console.warn('Spelling dictionaries were not added:', e.message);
}
// The licences travel with the code they cover
for (const f of ['LICENSE', 'NOTICE', 'LICENSE-UNICODE.txt']) fs.copyFileSync(path.join(SRC, f), path.join(OUT, f));
const notices = path.join(SRC, 'apps', 'shell', 'build', 'THIRD-PARTY-NOTICES.txt');
if (fs.existsSync(notices)) fs.copyFileSync(notices, path.join(OUT, 'THIRD-PARTY-NOTICES.txt'));
fs.copyFileSync(path.join(app, 'src', 'renderer', 'fonts', 'README.md'), path.join(OUT, 'FONTS.md'));
fs.writeFileSync(path.join(OUT, 'README.txt'), `The new Writer, Sheets and PDF editors in Kertas Office are built on GenOffice (https://github.com/genspark-ai/genoffice),
commit ${COMMIT}, used under the Apache License 2.0 (see LICENSE and NOTICE here).
Kertas changed it: see scripts/writer2/genoffice.patch in the Kertas Office source.
Kertas Office is not affiliated with or endorsed by the authors of GenOffice.
`);
console.log(`New editors built into ${path.relative(ROOT, OUT)}`);
