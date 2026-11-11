// Font files for non-Latin text in generated PDFs. Vite copies them into the
// build; they are fetched only when a document actually needs that script.
import base from '@expo-google-fonts/noto-sans/400Regular/NotoSans_400Regular.ttf?url';
import baseBold from '@expo-google-fonts/noto-sans/700Bold/NotoSans_700Bold.ttf?url';
import arabic from '@expo-google-fonts/noto-sans-arabic/400Regular/NotoSansArabic_400Regular.ttf?url';
import malayalam from '@expo-google-fonts/noto-sans-malayalam/400Regular/NotoSansMalayalam_400Regular.ttf?url';
import devanagari from '@expo-google-fonts/noto-sans-devanagari/400Regular/NotoSansDevanagari_400Regular.ttf?url';
import bengali from '@expo-google-fonts/noto-sans-bengali/400Regular/NotoSansBengali_400Regular.ttf?url';
import tamil from '@expo-google-fonts/noto-sans-tamil/400Regular/NotoSansTamil_400Regular.ttf?url';
import thai from '@expo-google-fonts/noto-sans-thai/400Regular/NotoSansThai_400Regular.ttf?url';
import sc from '@expo-google-fonts/noto-sans-sc/400Regular/NotoSansSC_400Regular.ttf?url';
import kr from '@expo-google-fonts/noto-sans-kr/400Regular/NotoSansKR_400Regular.ttf?url';
import { setFontLoader } from './pdf-text.js';
import { FONT_CATALOG, FACE_DIRS, facePath, cssStack } from './font-catalog.js';

// Every face of every catalog font (regular, italic, bold, bold italic), as URLs.
// Keep the package list in step with FONT_CATALOG (a test checks it).
const FACE_URLS = import.meta.glob('../../node_modules/@expo-google-fonts/{carlito,caladea,arimo,tinos,cousine,inter,roboto,open-sans,lato,montserrat,poppins,source-sans-3,nunito,work-sans,raleway,noto-sans,merriweather,lora,playfair-display,eb-garamond,libre-baskerville,source-serif-4,jetbrains-mono,roboto-mono,caveat,dancing-script,manjari,noto-sans-malayalam,cairo,amiri,noto-sans-arabic}/{400Regular,400Regular_Italic,700Bold,700Bold_Italic}/*.ttf', { query: '?url', import: 'default', eager: true });
const urlOf = (path) => FACE_URLS[`../../node_modules/${path}`] || null;
export const hasFace = (path) => !!urlOf(path);

const URLS = { base, baseBold, arabic, malayalam, devanagari, bengali, tamil, thai, sc, kr };
const cache = new Map();

export function installFontLoader() {
  setFontLoader((key) => {
    // 'face:<package path>' loads one face of a catalog font (Sheets PDF export)
    if (key.startsWith('face:')) {
      const url = urlOf(key.slice(5));
      if (!url) return Promise.resolve(null);
      if (!cache.has(key)) cache.set(key, fetch(url).then((r) => r.arrayBuffer()).then((b) => new Uint8Array(b)));
      return cache.get(key);
    }
    if (!URLS[key]) return Promise.resolve(null);
    if (!cache.has(key)) {
      cache.set(key, fetch(URLS[key]).then((r) => {
        if (!r.ok) throw new Error(`Font ${key}: HTTP ${r.status}`);
        return r.arrayBuffer();
      }).then((b) => new Uint8Array(b)).catch((e) => { cache.delete(key); throw e; }));
    }
    return cache.get(key);
  });
}

/** Declares @font-face rules for the catalog, so the fonts draw on screen (Sheets canvas, previews). */
export function installFontFaces() {
  const rules = [];
  for (const f of FONT_CATALOG) {
    for (const [face, dir] of Object.entries(FACE_DIRS)) {
      const url = urlOf(facePath(f, face));
      if (!url) continue;
      const weight = face.startsWith('bold') ? 700 : 400;
      const style = face.toLowerCase().includes('italic') ? 'italic' : 'normal';
      // An installed copy of the real font is used first (Calibri on Windows, for example)
      rules.push(`@font-face{font-family:'${f.family}';src:local('${f.family}${weight === 700 ? ' Bold' : ''}${style === 'italic' ? ' Italic' : ''}'),url('${url}') format('truetype');font-weight:${weight};font-style:${style};font-display:swap}`);
    }
  }
  const el = document.createElement('style');
  el.id = 'kertas-fonts';
  el.textContent = rules.join('\n');
  document.head.appendChild(el);
}

/** SuperDoc's font registration and dropdown rows */
export function superdocFonts() {
  const families = FONT_CATALOG.map((f) => ({
    family: f.family,
    faces: Object.keys(FACE_DIRS).map((face) => { const src = urlOf(facePath(f, face)); return { source: src, url: src, weight: face.startsWith('bold') ? 700 : 400, style: face.toLowerCase().includes('italic') ? 'italic' : 'normal' }; }).filter((x) => x.source),
  }));
  const options = FONT_CATALOG.map((f) => ({ label: f.family, key: f.family, props: { style: { fontFamily: cssStack(f) } } }));
  return { families, options };
}

/** Univer's font list */
export const univerFonts = () => FONT_CATALOG.map((f) => ({ value: f.family, label: f.family, category: { mono: 'monospace', serif: 'serif', hand: 'handwriting' }[f.category] || 'sans-serif' }));
