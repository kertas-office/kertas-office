// End-to-end harness: serves the production build, drives it in Chromium the way a
// person would, and hands saved files back to Node so they can be checked with
// independent readers (JSZip, ExcelJS, pdf-lib, pdf.js) rather than by Kertas itself.
import { spawn } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { chromium } from 'playwright';

export const ROOT = fileURLToPath(new URL('../..', import.meta.url));
export const fixture = (name) => path.join(ROOT, 'test', 'fixtures', name);

let server = null;
let url = null;

/** Starts `vite preview` on the built app (run `vite build` first; `npm run e2e` does). */
export async function startServer() {
  if (url) return url;
  const port = 4300 + Math.floor(Math.random() * 500);
  server = spawn(process.execPath, [path.join(ROOT, 'node_modules/vite/bin/vite.js'), 'preview', '--port', String(port), '--strictPort'], { cwd: ROOT, stdio: 'pipe' });
  url = `http://localhost:${port}/`;
  for (let i = 0; i < 100; i++) {
    try { if ((await fetch(url)).ok) return url; } catch { /* not up yet */ }
    await new Promise((r) => setTimeout(r, 150));
  }
  throw new Error('preview server did not start');
}

export function stopServer() { server?.kill(); server = null; url = null; }

/**
 * Opens Kertas in a fresh browser page. Uncaught page errors are collected on
 * `page.errors` so each test can assert that none happened.
 */
export async function openApp({ dark = false } = {}) {
  const base = await startServer();
  const browser = await chromium.launch();
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 }, acceptDownloads: true, colorScheme: dark ? 'dark' : 'light' });
  const page = await context.newPage();
  page.setDefaultTimeout(10000);
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(String(e.message || e)));
  await page.goto(base);
  await page.waitForSelector('#tabs');
  page.close = async () => { await browser.close(); };
  return page;
}

/** Opens a file from test/fixtures in a new tab. */
export async function openFixture(page, kind, name) {
  return openFile(page, kind, name, await readFile(fixture(name)));
}

/** Opens bytes in a new tab of the given kind (no bytes: a new blank file) and waits for the editor. */
export async function openFile(page, kind, name, bytes) {
  const data = bytes ? Buffer.from(bytes) : null;
  await page.evaluate(async ([kind, name, b64]) => {
    const data = b64 ? Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)) : undefined;
    const t = await window.__kertas.openTab(kind, { name: path(name), data });
    window.__t = t;
    for (let i = 0; i < 300 && !t.instance; i++) await new Promise((r) => setTimeout(r, 100));
    if (!t.instance) throw new Error(`${kind} did not open ${name}`);
    function path(n) { return n ? n.split('/').pop() : undefined; }
  }, [kind, name || null, data ? data.toString('base64') : null]);
  await settle(page, kind);
}

/** Waits until the editor has drawn its first page. */
export async function settle(page, kind) {
  const sel = { writer: '.superdoc-page:visible', sheets: 'canvas:visible', pdf: '.pdf-page:visible' }[kind];
  await page.locator(sel).first().waitFor({ state: 'visible', timeout: 30000 });
  await page.waitForTimeout(kind === 'writer' ? 1200 : 800);
}

/** Saves the active tab through its editor and returns the bytes as a Buffer. */
export async function saveBytes(page, ext) {
  const b64 = await page.evaluate(async (ext) => {
    const bytes = await window.__t.instance.save(ext);
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s);
  }, ext);
  return Buffer.from(b64, 'base64');
}

/** Exports the active tab to PDF and returns the bytes (Writer and Sheets print through the browser). */
export async function exportPdfBytes(page) {
  const b64 = await page.evaluate(async () => {
    const bytes = new Uint8Array(await window.__t.instance.exportPdf());
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return btoa(s);
  });
  return Buffer.from(b64, 'base64');
}

/** Clicks a ribbon tab by its label, then a button in it by data-id. */
export async function ribbon(page, tabLabel, id) {
  if (tabLabel) await page.locator('.rb-tabs button, .ribbon-tabs button, [role=tab]').filter({ hasText: new RegExp(`^${tabLabel}$`) }).first().click();
  if (id) await page.locator(`[data-id="${id}"]`).first().click();
  await page.waitForTimeout(250);
}

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
