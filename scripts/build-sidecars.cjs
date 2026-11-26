#!/usr/bin/env node
// Builds the new Sheets' helper program (a small Rust program that reads and writes
// workbooks) for every processor this system's installers are made for, into
// build/sidecar/<arch>/, where electron-builder picks up the one for each installer.
//
//   node scripts/build-sidecars.cjs            the architectures of this system's installers
//   node scripts/build-sidecars.cjs x64        just one
//
// Needs the editors' source fetched by scripts/build-writer2.cjs (or KERTAS_GENOFFICE_DIR)
// and Rust (rustup); missing targets are added with rustup.
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC = process.env.KERTAS_GENOFFICE_DIR || path.join(ROOT, '.cache', 'genoffice');
const ENGINE = path.join(SRC, 'apps', 'sheets', 'native', 'xlsx-engine');
const TARGETS = {
  win32: { x64: 'x86_64-pc-windows-msvc', arm64: 'aarch64-pc-windows-msvc' },
  darwin: { x64: 'x86_64-apple-darwin', arm64: 'aarch64-apple-darwin' },
  linux: { x64: 'x86_64-unknown-linux-gnu', arm64: 'aarch64-unknown-linux-gnu' },
};
// what electron-builder.config.cjs makes on each system
const ARCHES = { win32: ['x64', 'arm64'], darwin: ['x64', 'arm64'], linux: ['x64'] };

const run = (cmd, args, cwd) => execFileSync(cmd, args, { cwd, stdio: 'inherit', shell: process.platform === 'win32' });
const exe = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar';

if (!fs.existsSync(path.join(ENGINE, 'Cargo.toml'))) {
  console.error('The editors\' source is not there: run node scripts/build-writer2.cjs first.');
  process.exit(1);
}
const arches = process.argv.slice(2).length ? process.argv.slice(2) : ARCHES[process.platform];
for (const arch of arches) {
  const target = TARGETS[process.platform]?.[arch];
  if (!target) throw new Error(`No helper target for ${process.platform} ${arch}`);
  try { run('rustup', ['target', 'add', target]); } catch { console.warn(`rustup could not add ${target}; trying with what is installed`); }
  run('cargo', ['build', '--release', '--locked', '--target', target], ENGINE);
  const out = path.join(ROOT, 'build', 'sidecar', arch);
  fs.mkdirSync(out, { recursive: true });
  fs.copyFileSync(path.join(ENGINE, 'target', target, 'release', exe), path.join(out, exe));
  if (process.platform !== 'win32') fs.chmodSync(path.join(out, exe), 0o755);
  fs.copyFileSync(path.join(ENGINE, 'Cargo.lock'), path.join(out, 'Cargo.lock'));
  console.log(`Helper for ${arch} built into ${path.relative(ROOT, out)}`);
}
