// Website screenshots of the editors, light and dark: node scripts/site-shots/shots.mjs <samples-dir> <out-dir>
// Opens each sample in the desktop app (built, with the editors built), and writes
// writer/sheets/pdf .webp at 2000x1250 plus -sm (1000x625), and -dark versions.
// The editors are views laid over the Kertas window, so each picture is the window with
// the editor's own capture placed where the editor sits.
import { _electron } from 'playwright';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('../..', import.meta.url));
const require = createRequire(path.join(ROOT, 'package.json'));
const [src, out] = process.argv.slice(2);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const SCALE = 1.6; // 1250 x 781 window -> 2000 x 1250

const dir = await mkdtemp(path.join(tmpdir(), 'kertas-shots-'));
const app = await _electron.launch({
  executablePath: require('electron'),
  args: [ROOT, '--no-sandbox', `--force-device-scale-factor=${SCALE}`],
  env: { ...process.env, XDG_CONFIG_HOME: path.join(dir, 'config'), NODE_ENV: 'production' },
});
const k = await app.firstWindow();
await k.waitForSelector('#tabs');
await app.evaluate(({ BrowserWindow }) => { const w = BrowserWindow.getAllWindows()[0]; w.setContentSize(1250, 781); });
await sleep(800);

async function open(file) {
  await app.evaluate(({ dialog }, p) => { dialog.showOpenDialog = async () => ({ canceled: false, filePaths: [p] }); }, file);
  await k.keyboard.press('Control+o');
}
/** The window, with every visible editor view's own capture laid on top, as PNG bytes */
async function shot() {
  const base = await k.screenshot({ type: 'png' });
  const layers = await app.evaluate(async ({ BrowserWindow }) => {
    const w = BrowserWindow.getAllWindows()[0];
    const out = [];
    for (const v of w.contentView.children) {
      if (!v.webContents || !v.getVisible?.()) continue;
      const b = v.getBounds();
      if (b.width < 10) continue;
      const img = await v.webContents.capturePage();
      out.push({ b, png: img.toPNG().toString('base64') });
    }
    return out;
  });
  return { base, layers };
}

const results = {};
async function capture(name) {
  const { base, layers } = await shot();
  results[name] = { base: base.toString('base64'), layers };
}

// Writer
await open(path.join(src, 'Annual report 2025.docx'));
await sleep(6000);
await capture('writer');
// Sheets
await open(path.join(src, 'Household budget 2026.xlsx'));
await sleep(7000);
await capture('sheets');
// PDF, on its Convert tab
await open(path.join(src, 'Annual report 2025.pdf'));
await sleep(5000);
await app.evaluate(async ({ webContents }) => {
  const wc = webContents.getAllWebContents().find((w) => /^kertas-writer:\/\/pdf\//.test(w.getURL()));
  await wc?.executeJavaScript(`[...document.querySelectorAll('.ribbon-tab')].find((b) => b.textContent.trim() === 'Convert')?.click()`);
});
await sleep(1500);
await capture('pdf');

// dark: each tab again
await k.evaluate(() => window.__kertas.applyTheme(true));
await sleep(1500);
const tabIds = await k.evaluate(() => [...window.__kertas.tabs.values()].map((t) => t.id));
for (const [i, name] of ['writer', 'sheets', 'pdf'].entries()) {
  await k.evaluate((id) => window.__kertas.activate(id), tabIds[i]);
  await sleep(2500);
  await capture(`${name}-dark`);
}

// compose.py lays the editors over the window and writes the .webp files
const { writeFile } = await import('node:fs/promises');
await writeFile(path.join(dir, 'layers.json'), JSON.stringify({ scale: SCALE, results }));
console.log(path.join(dir, 'layers.json'));
await Promise.race([app.close(), sleep(3000)]);
process.exit(0);
