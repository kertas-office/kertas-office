// Renders the Kertas app icon at every size the app, installers and website need.
// The mark is a sheet of paper folded into a K (kertas is Malay for paper): a white top sheet
// and a lavender flap folded along one straight line, on the purple tile. Smaller sizes draw
// the mark larger on the tile so it stays clear in the taskbar, title bar and browser tab:
//   tiny   (16-24 px)  the mark almost edge to edge, no background circles
//   small  (32-48 px)  the mark a little larger than in the full icon
//   full   (64 px up)  Windows, Linux and the website
//   mac    (.icns)     the full icon inside Apple's standard margin, with a drop shadow
// Run after changing the design: node scripts/make-icons.cjs   (needs Playwright's Chromium)
const fs = require('fs');
const path = require('path');
const png2icons = require('png2icons');

let playwright;
try { playwright = require('playwright'); } catch { playwright = require(process.env.PLAYWRIGHT_PATH || '/opt/npm-tools/node_modules/playwright'); }

const ROOT = path.join(__dirname, '..');
const out = (...p) => { const f = path.join(ROOT, ...p); fs.mkdirSync(path.dirname(f), { recursive: true }); return f; };

// ---------------------------------------------------------------- the design
// Everything is drawn on a 1024 grid. The mark is drawn on a 256 grid and placed with mark().
const TILE = `<linearGradient id="tile" x1="0" y1="0" x2="1024" y2="1024" gradientUnits="userSpaceOnUse">
    <stop offset="0" stop-color="#5b5ff5"/><stop offset=".52" stop-color="#8b5cf6"/><stop offset="1" stop-color="#d946ef"/></linearGradient>`;
// shadow along the fold, falling onto the flap
const CREASE = `<linearGradient id="crease" gradientUnits="userSpaceOnUse" x1="104" y1="166" x2="124" y2="192">
    <stop offset="0" stop-color="#2e1065" stop-opacity=".28"/><stop offset="1" stop-color="#2e1065" stop-opacity="0"/></linearGradient>`;

// The fold runs in one straight line from the top right corner to the left edge
const SHEET = 'M64 72Q64 40 96 40H238L64 208Z';
const FLAP = 'M142 132L246 232H102C78 232 64 222 64 211C64 206 66 202.5 70 200.3Z';

/** The folded K, size px wide, centred on the 1024 tile */
function mark(size) {
  const k = size / 256;
  const at = 512 - size / 2;
  // the mark spans x 64-246 and y 40-232 on its grid; (155, 136) is its centre
  return `<g transform="translate(${at} ${at}) scale(${k}) translate(-27 -8)">
      <path d="${SHEET}" fill="#fff"/>
      <path d="${FLAP}" fill="#e3d4ff"/>
      <path d="${FLAP}" fill="url(#crease)"/>
    </g>`;
}

function tile(t, r, inner, { shadow = false, circles = true } = {}) {
  const s = 1024 - 2 * t;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1024 1024">
  <defs>${TILE}${CREASE}
    <filter id="drop" x="-20%" y="-20%" width="140%" height="150%"><feDropShadow dx="0" dy="22" stdDeviation="26" flood-color="#1e1b4b" flood-opacity=".38"/></filter>
    <filter id="lift" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="10" stdDeviation="14" flood-color="#312e81" flood-opacity=".28"/></filter>
    <clipPath id="body"><rect x="${t}" y="${t}" width="${s}" height="${s}" rx="${r}"/></clipPath>
  </defs>
  <g${shadow ? ' filter="url(#drop)"' : ''}><rect x="${t}" y="${t}" width="${s}" height="${s}" rx="${r}" fill="url(#tile)"/></g>
  <g clip-path="url(#body)">
    ${circles ? '<circle cx="860" cy="170" r="300" fill="#fff" opacity=".10"/>\n    <circle cx="140" cy="930" r="340" fill="#312e81" opacity=".16"/>' : ''}
    <g filter="url(#lift)">${inner}</g>
  </g>
  <rect x="${t + 2}" y="${t + 2}" width="${s - 4}" height="${s - 4}" rx="${r - 2}" fill="none" stroke="#fff" stroke-opacity=".3" stroke-width="4"/>
</svg>`;
}

// Full icon: the tile fills the canvas (Windows and Linux add no frame of their own)
const full = () => tile(24, 230, mark(724));
// macOS: Apple's grid puts the 824 px tile inside a 1024 canvas, with a shadow
const mac = () => full().replace('viewBox="0 0 1024 1024">', 'viewBox="0 0 1024 1024"><g transform="translate(100 100) scale(.8047)">').replace(/<\/svg>$/, '</g></svg>')
  .replace('<g><rect', '<g filter="url(#drop)"><rect');
// Small: the mark a little larger
const small = () => tile(16, 220, mark(800));
// Tiny: the mark almost edge to edge, plain tile
const tiny = () => tile(8, 210, mark(880), { circles: false });

const forSize = (s) => (s <= 24 ? tiny() : s <= 48 ? small() : full());

// ---------------------------------------------------------------- .ico with hand-drawn sizes
function ico(pngs) { // pngs: [{ size, data }]
  const head = Buffer.alloc(6 + 16 * pngs.length);
  head.writeUInt16LE(0, 0); head.writeUInt16LE(1, 2); head.writeUInt16LE(pngs.length, 4);
  let offset = head.length;
  pngs.forEach(({ size, data }, i) => {
    const e = 6 + 16 * i;
    head.writeUInt8(size >= 256 ? 0 : size, e); head.writeUInt8(size >= 256 ? 0 : size, e + 1);
    head.writeUInt8(0, e + 2); head.writeUInt8(0, e + 3);
    head.writeUInt16LE(1, e + 4); head.writeUInt16LE(32, e + 6);
    head.writeUInt32LE(data.length, e + 8); head.writeUInt32LE(offset, e + 12);
    offset += data.length;
  });
  return Buffer.concat([head, ...pngs.map((p) => p.data)]);
}

(async () => {
  const browser = await playwright.chromium.launch();
  const page = await browser.newPage();
  const render = async (svg, size) => {
    await page.setViewportSize({ width: size, height: size });
    await page.setContent(`<html><body style="margin:0;background:transparent">${svg.replace('<svg ', `<svg width="${size}" height="${size}" `)}</body></html>`);
    return page.screenshot({ omitBackground: true, clip: { x: 0, y: 0, width: size, height: size } });
  };

  // Sources, kept for reference and for the website and app
  fs.writeFileSync(out('build/icon-src/app.svg'), full());
  fs.writeFileSync(out('build/icon-src/app-mac.svg'), mac());
  fs.writeFileSync(out('build/icon-src/app-small.svg'), small());
  fs.writeFileSync(out('build/icon-src/app-tiny.svg'), tiny());

  // electron-builder: icon.png (Linux and fallback), icon.icns (macOS), icon.ico (Windows)
  fs.writeFileSync(out('build/icon.png'), await render(full(), 1024));
  fs.writeFileSync(out('build/icon.icns'), png2icons.createICNS(await render(mac(), 1024), png2icons.BICUBIC2, 0));
  const icoSizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256];
  const icoPngs = [];
  for (const s of icoSizes) icoPngs.push({ size: s, data: await render(forSize(s), s) });
  fs.writeFileSync(out('build/icon.ico'), ico(icoPngs));

  // Linux icon theme sizes, and the window icon Electron uses on Linux
  for (const s of [16, 24, 32, 48, 64, 128, 256, 512]) fs.writeFileSync(out(`build/icons/${s}x${s}.png`), await render(forSize(s), s));
  fs.writeFileSync(out('electron/icon.png'), await render(full(), 512));

  // In-app logo and favicon
  fs.writeFileSync(out('src/public/logo.svg'), small()); // shown at 26 px in the title bar
  fs.writeFileSync(out('src/public/favicon.svg'), small());
  fs.writeFileSync(out('src/public/icon-192.png'), await render(full(), 192));

  // Website
  fs.writeFileSync(out('website/img/logo.svg'), small()); // header and footer, 30-34 px
  fs.writeFileSync(out('website/img/logo-full.svg'), full());
  fs.writeFileSync(out('website/img/favicon.svg'), small());
  fs.writeFileSync(out('website/img/favicon-32.png'), await render(small(), 32));
  fs.writeFileSync(out('website/img/favicon.ico'), ico([16, 32, 48].map((s, i) => ({ size: s, data: icoPngs[[0, 3, 5][i]].data }))));
  // Apple touch icon: square and opaque (iOS rounds the corners itself)
  const touch = full().replace('rx="230"', 'rx="0"').replace('<rect x="24" y="24" width="976" height="976" rx="0"', '<rect x="0" y="0" width="1024" height="1024" rx="0"')
    .replace('<rect x="26" y="26" width="972" height="972" rx="228" fill="none" stroke="#fff" stroke-opacity=".3" stroke-width="4"/>', '');
  fs.writeFileSync(out('website/img/apple-touch-icon.png'), await render(touch.replace(/<clipPath id="body"><rect[^>]*\/><\/clipPath>/, '<clipPath id="body"><rect x="0" y="0" width="1024" height="1024"/></clipPath>'), 180));
  fs.writeFileSync(out('website/img/icon-512.png'), await render(full(), 512));

  await browser.close();
  console.log('Icons written: build/, electron/icon.png, src/public/, website/img/');
})().catch((e) => { console.error(e); process.exit(1); });
