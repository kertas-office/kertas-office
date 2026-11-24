#!/usr/bin/env node
// Builds the new Writer (trial) into electron/writer2/.
//
// The editor comes from GenOffice (Apache-2.0, Copyright Mainfunc, Inc.), taken at
// a fixed commit and changed by scripts/writer2/genoffice.patch: Kertas hosts it
// in its own window, the AI assistant is switched off, product names are
// Kertas's, and it keeps its files in a folder of its own. Only the Apache-2.0
// parts are used; the separately licensed ee/ directory is deleted before
// anything is built.
//
//   node scripts/build-writer2.cjs            fetch, patch, build, copy
//   KERTAS_GENOFFICE_DIR=<checkout>           use a checkout that is already patched
//
// Without this step Kertas builds and runs as before, with no new Writer.
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
run('npx', ['electron-vite', 'build', '--config', 'kertas.vite.config.ts'], app);
try { run('node', ['tools/gen-third-party-notices.mjs'], SRC); } catch { console.warn('Third-party notices were not regenerated'); }

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'wasm'), { recursive: true });
for (const part of ['main', 'preload', 'renderer']) fs.cpSync(path.join(app, 'out-kertas', part), path.join(OUT, part), { recursive: true });
fs.copyFileSync(path.join(SRC, 'node_modules', '@embedpdf', 'pdfium', 'dist', 'pdfium.wasm'), path.join(OUT, 'wasm', 'pdfium.wasm'));
// The licences travel with the code they cover
for (const f of ['LICENSE', 'NOTICE', 'LICENSE-UNICODE.txt']) fs.copyFileSync(path.join(SRC, f), path.join(OUT, f));
const notices = path.join(SRC, 'apps', 'shell', 'build', 'THIRD-PARTY-NOTICES.txt');
if (fs.existsSync(notices)) fs.copyFileSync(notices, path.join(OUT, 'THIRD-PARTY-NOTICES.txt'));
fs.copyFileSync(path.join(app, 'src', 'renderer', 'fonts', 'README.md'), path.join(OUT, 'FONTS.md'));
fs.writeFileSync(path.join(OUT, 'README.txt'), `The new Writer in Kertas Office is built on GenOffice (https://github.com/genspark-ai/genoffice),
commit ${COMMIT}, used under the Apache License 2.0 (see LICENSE and NOTICE here).
Kertas changed it: see scripts/writer2/genoffice.patch in the Kertas Office source.
Kertas Office is not affiliated with or endorsed by the authors of GenOffice.
`);
console.log(`New Writer built into ${path.relative(ROOT, OUT)}`);
