// Renders every slide of the Slides templates with the Slides editor itself (File > Export slides
// as PNG pictures) into <out>/<name>/, for the preview pictures and for checking the decks.
// Run after npm run build and node scripts/build-writer2.cjs:
//   xvfb-run -a node scripts/templates/slide-previews.mjs <out-dir>
import { _electron } from 'playwright';
import { createRequire } from 'node:module';
import { mkdirSync, mkdtempSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const OUT = path.resolve(process.argv[2] || path.join(tmpdir(), 'slide-previews'));
const NAMES = (process.argv[3] || 'simple,pitch-deck,business-review,project-status,training,product-launch').split(',');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const app = await _electron.launch({
  executablePath: require('electron'),
  args: [ROOT, '--no-sandbox', '--window-size=1400,1000'],
  env: { ...process.env, XDG_CONFIG_HOME: mkdtempSync(path.join(tmpdir(), 'kertas-previews-')), NODE_ENV: 'production' },
});
const k = await app.firstWindow();
await k.waitForSelector('#tabs');
const editors = () => app.context().pages().filter((p) => p.url().startsWith('kertas-writer://slides/'));

for (const name of NAMES) {
  const dir = path.join(OUT, name);
  mkdirSync(dir, { recursive: true });
  const file = path.join(ROOT, 'src', 'templates', `${name}.pptx`);
  await app.evaluate(({ dialog }, [file, dir]) => {
    dialog.showOpenDialog = async (_w, o) => {
      const props = (o || _w)?.properties || [];
      return { canceled: false, filePaths: [props.includes('openDirectory') ? dir : file] };
    };
    dialog.showSaveDialog = async (_w, o) => ({ canceled: false, filePath: `${dir}/${String((o || _w)?.defaultPath || 'slides.png').split(/[\\/]/).pop()}` });
  }, [file, dir]);
  const had = new Set(editors());
  await k.locator('#tabs .tab').first().click();
  await k.keyboard.press('Control+o');
  let w = null;
  for (let i = 0; i < 80 && !w; i++) { w = editors().find((p) => !had.has(p)); if (!w) await sleep(250); }
  if (!w) throw new Error(`${name} did not open`);
  await w.waitForSelector('.slide-list .thumb');
  await sleep(3000);
  await w.locator('.kertas-file-tab').click();
  await w.locator('.kertas-file-menu button').filter({ has: w.locator('span', { hasText: /^Export slides as PNG/ }) }).click();
  for (let i = 0; i < 120 && !readdirSync(dir).some((f) => f.endsWith('.png')); i++) await sleep(250);
  await sleep(1500);
  console.log(name, readdirSync(dir).filter((f) => f.endsWith('.png')).length, 'slides');
  await k.locator('#tabs .tab .close').last().click();
  await sleep(800);
}
app.process().kill('SIGKILL');
process.exit(0);
