// Builds the website into website/dist: wraps each page in src/pages with src/layout.html
// (shared header and footer), copies assets and images, and writes the sitemap.
// Run with `node website/build.mjs` (no dependencies). Cloudflare Pages runs the same command.
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const out = path.join(root, 'dist');
const SITE = 'https://kertasoffice.com';

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });

const copyDir = (from, to) => {
  fs.mkdirSync(to, { recursive: true });
  for (const e of fs.readdirSync(from, { withFileTypes: true })) {
    const a = path.join(from, e.name);
    const b = path.join(to, e.name);
    if (e.isDirectory()) copyDir(a, b); else fs.copyFileSync(a, b);
  }
};
copyDir(path.join(root, 'assets'), path.join(out, 'assets'));
copyDir(path.join(root, 'img'), path.join(out, 'img'));
for (const f of fs.readdirSync(path.join(root, 'static'))) fs.copyFileSync(path.join(root, 'static', f), path.join(out, f));

// Cache-busting version from the stylesheet and script contents
const version = crypto.createHash('sha1')
  .update(fs.readFileSync(path.join(root, 'assets/site.css')))
  .update(fs.readFileSync(path.join(root, 'assets/site.js')))
  .digest('hex').slice(0, 8);

const layout = fs.readFileSync(path.join(root, 'src/layout.html'), 'utf8');
const pages = [];
for (const file of fs.readdirSync(path.join(root, 'src/pages')).filter((f) => f.endsWith('.html'))) {
  const src = fs.readFileSync(path.join(root, 'src/pages', file), 'utf8');
  const meta = {};
  const body = src.replace(/^\s*<!--\s*(\w+):\s*([\s\S]*?)\s*-->\s*$/gm, (_, k, v) => { meta[k] = v; return ''; }).trim();
  for (const k of ['title', 'description', 'path', 'nav']) if (!meta[k]) throw new Error(`${file}: missing <!-- ${k}: ... -->`);
  const html = layout
    .replace(/\{\{cur:(\w+)\}\}/g, (_, n) => (n === meta.nav ? ' aria-current="page"' : ''))
    .replace(/\{\{(\w+)\}\}/g, (_, k) => (k === 'content' ? body : k === 'version' ? version : (meta[k] ?? '').replace(/"/g, '&quot;')));
  if (/\u2014/.test(html)) throw new Error(`${file}: contains an em dash`);
  fs.writeFileSync(path.join(out, file), html);
  if (meta.path !== '/404') pages.push(meta.path);
}

fs.writeFileSync(path.join(out, 'sitemap.xml'), `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${pages.sort().map((p) => `  <url><loc>${SITE}${p}</loc></url>`).join('\n')}
</urlset>
`);
console.log(`Built ${pages.length + 1} pages into ${path.relative(process.cwd(), out) || out}`);
