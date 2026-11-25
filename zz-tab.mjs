// Launch Kertas with the new Writer forced on, open a docx as a tab, capture the real screen.
import { chromium } from 'playwright';
import { spawn, execFileSync } from 'node:child_process';
const S='/tmp/claude-0/-home-claude/bb666f01-839e-591e-8c77-11557a8dcc1d/scratchpad';
const [file, out, script] = process.argv.slice(2);
const args = ['--no-sandbox', '--remote-debugging-port=9333', `--user-data-dir=${S}/w2/ud4`, '--window-size=1400,1000', '--window-position=0,0', '.'];
if (file && file !== '-') args.push(file);
const p = spawn('node_modules/.bin/electron', args, { env: { ...process.env, KERTAS_NEW_WRITER: '1' }, stdio: ['ignore', 'pipe', 'pipe'] });
p.stderr.on('data', d => { const t = String(d); if (!/dbus|gpu|GPU|vaapi|DevTools listening|openssl|APPIMAGE/.test(t)) process.stdout.write('[err] ' + t); });
await new Promise(r => setTimeout(r, 10000));
const b = await chromium.connectOverCDP('http://127.0.0.1:9333');
const ctx = b.contexts()[0];
const pages = () => ctx.pages();
const k = pages().find(p => p.url().startsWith('app://'));
const w = pages().find(p => p.url().startsWith('kertas-writer://'));
console.log('pages', pages().map(p => p.url()));
for (const pg of pages()) { pg.on('pageerror', e => console.log('[pageerror]', String(e).slice(0, 300))); pg.on('console', m => { if (m.type() === 'error') console.log('[console]', m.text().slice(0, 300)); }); }
const shot = (name) => execFileSync('import', ['-window', 'root', `${S}/w2/${name}`]);
shot(out);
if (script) { const mod = await import(script); await mod.default({ k, w, pages, shot, ctx }); }
p.kill(); process.exit(0);
