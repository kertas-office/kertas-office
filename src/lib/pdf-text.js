// Unicode text for PDFs we generate (sheet export, stamps, watermarks, page
// numbers, filled forms). The 14 standard PDF fonts only cover Western
// European text, so any other script is drawn with an embedded Noto font,
// subset to the glyphs used, and shaped with fontkit (Arabic joining,
// Indic reordering, Hangul) with mark positions applied glyph by glyph.
//
//   const kit = createTextKit(pdf, { regular, bold });   // standard PDFFonts
//   await kit.prepare([...all strings]);                 // embeds what is needed
//   kit.width(text, size, bold); kit.draw(page, text, { x, y, size, color, bold });
//
// Pure module: fonts are loaded through setFontLoader (browser and Node differ).
// Upstream fontkit 2: complete Arabic, Indic and Hangul shapers
import * as fontkit from 'fontkit';
import {
  pushGraphicsState, popGraphicsState, beginText, endText, setFontAndSize, setTextMatrix, showText, setFillingColor,
  setGraphicsState, PDFHexString, degrees,
} from '@cantoo/pdf-lib';

/** Fonts by key; the loader turns a key into font bytes */
export const FONT_KEYS = ['base', 'baseBold', 'arabic', 'malayalam', 'devanagari', 'bengali', 'tamil', 'thai', 'sc', 'kr'];

let loader = null;
/** fn(key) -> Promise<Uint8Array | ArrayBuffer>. Keys: FONT_KEYS, or 'face:<package path>' for catalog fonts */
export function setFontLoader(fn) { loader = fn; }
export const loadFontBytes = async (key) => (loader ? loader(key) : null);

// Characters the standard (WinAnsi) fonts can draw
const WIN_ANSI_EXTRA = new Set([0x20ac, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021, 0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x017d,
  0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x017e, 0x0178]);
export const isWinAnsi = (cp) => (cp >= 0x20 && cp <= 0x7e) || (cp >= 0xa0 && cp <= 0xff) || WIN_ANSI_EXTRA.has(cp) || cp === 0x09 || cp === 0x0a;

/** Which font a character needs: 'std' or a FONT_KEYS entry */
export function scriptOf(cp) {
  if (isWinAnsi(cp)) return 'std';
  if ((cp >= 0x0600 && cp <= 0x06ff) || (cp >= 0x0750 && cp <= 0x077f) || (cp >= 0x08a0 && cp <= 0x08ff) || (cp >= 0xfb50 && cp <= 0xfdff) || (cp >= 0xfe70 && cp <= 0xfeff)) return 'arabic';
  if (cp >= 0x0d00 && cp <= 0x0d7f) return 'malayalam';
  if ((cp >= 0x0900 && cp <= 0x097f) || (cp >= 0xa8e0 && cp <= 0xa8ff)) return 'devanagari';
  if (cp >= 0x0980 && cp <= 0x09ff) return 'bengali';
  if (cp >= 0x0b80 && cp <= 0x0bff) return 'tamil';
  if (cp >= 0x0e00 && cp <= 0x0e7f) return 'thai';
  if ((cp >= 0x1100 && cp <= 0x11ff) || (cp >= 0x3130 && cp <= 0x318f) || (cp >= 0xac00 && cp <= 0xd7af)) return 'kr';
  if ((cp >= 0x2e80 && cp <= 0x9fff) || (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xff00 && cp <= 0xffef) || (cp >= 0x20000 && cp <= 0x2fa1f)) return 'sc';
  return 'base';
}

const isMark = (cp) => (cp >= 0x0300 && cp <= 0x036f) || cp === 0x200c || cp === 0x200d || (cp >= 0xfe00 && cp <= 0xfe0f);
const isDigit = (cp) => cp >= 0x30 && cp <= 0x39;
const isLetterStd = (cp) => (cp >= 0x41 && cp <= 0x5a) || (cp >= 0x61 && cp <= 0x7a) || (cp >= 0xc0 && cp <= 0xff && cp !== 0xd7 && cp !== 0xf7);
const RTL = new Set(['arabic']);

/**
 * Splits text into runs that each use one font, in visual (left to right) order.
 * Spaces stay inside the word runs around them so words shape as a whole.
 * A small bidi pass orders right-to-left (Arabic) text: runs of Arabic, with
 * the numbers and spaces between them, read right to left as a block while
 * each number keeps its own left-to-right digits.
 */
export function splitRuns(text) {
  const runs = [];
  let cur = null;
  const push = (script, kind, ch) => { cur = { script, kind, text: ch }; runs.push(cur); };
  const hasRtl = /[\u0600-\u06ff\u0750-\u077f\u08a0-\u08ff\ufb50-\ufdff\ufe70-\ufeff]/.test(String(text));
  for (const ch of String(text)) {
    const cp = ch.codePointAt(0);
    const s = scriptOf(cp);
    if (cur && isMark(cp)) { cur.text += ch; continue; }
    if (s === 'arabic') {
      if (cur && cur.kind === 'rtl') cur.text += ch; else push('arabic', 'rtl', ch);
    } else if (cur && cur.kind === 'weak' && (ch === ',' || ch === '.' || ch === ':' || ch === '/')) {
      cur.text += ch; // 1,250.00 and 10:30 stay one number
    } else if (isDigit(cp)) {
      // Digits join a non-Latin word run (e.g. 120 after Chinese), but never an Arabic one
      if (cur && (cur.kind === 'weak' || (cur.kind === 'ltr' && cur.script !== 'std'))) cur.text += ch;
      else push('std', 'weak', ch);
    } else if (s === 'std' && !isLetterStd(cp)) {
      // Spaces and punctuation
      if (cur && (cur.kind === 'neutral' || cur.kind === 'rtl' || (cur.kind === 'ltr' && cur.script !== 'std' && !hasRtl))) cur.text += ch;
      else if (cur && cur.kind === 'ltr' && !hasRtl) cur.text += ch; // with Arabic around, punctuation takes its side from the context
      else push('std', 'neutral', ch);
    } else if (cur && cur.kind === 'ltr' && cur.script === s) {
      cur.text += ch;
    } else {
      push(s, 'ltr', ch);
    }
  }
  return reorder(runs);
}

const MIRROR = { '(': ')', ')': '(', '[': ']', ']': '[', '{': '}', '}': '{', '<': '>', '>': '<', '«': '»', '»': '«' };

function reorder(runs) {
  if (!runs.some((r) => r.kind === 'rtl')) return runs;
  const strong = runs.map((r) => (r.kind === 'rtl' ? 'R' : r.kind === 'ltr' ? 'L' : null));
  const para = strong.find(Boolean) || 'L';
  // Resolve numbers and neutrals from their surroundings
  const level = runs.map((r, i) => {
    if (strong[i]) return strong[i];
    let prev = para;
    for (let k = i - 1; k >= 0; k--) if (strong[k]) { prev = strong[k]; break; }
    if (r.kind === 'weak') return prev === 'R' ? 'R' : 'L';
    let next = para;
    for (let k = i + 1; k < runs.length; k++) if (strong[k] || runs[k].kind === 'weak') { next = strong[k] || (prev === 'R' ? 'R' : 'L'); break; }
    return prev === next ? prev : para;
  });
  // Brackets in right-to-left text are drawn mirrored, as every text engine does
  runs.forEach((r, i) => { if (level[i] === 'R' && r.kind !== 'weak') r.text = r.text.replace(/[()[\]{}<>«»]/g, (c) => MIRROR[c]); });
  const out = [];
  if (para === 'L') {
    for (let i = 0; i < runs.length;) {
      if (level[i] !== 'R') { out.push(runs[i++]); continue; }
      let j = i; while (j < runs.length && level[j] === 'R') j++;
      out.push(...runs.slice(i, j).reverse());
      i = j;
    }
    return out;
  }
  // Right-to-left paragraph: everything reversed, left-to-right stretches kept in order
  const rev = runs.map((r, i) => [r, level[i]]).reverse();
  for (let i = 0; i < rev.length;) {
    if (rev[i][1] !== 'L') { out.push(rev[i++][0]); continue; }
    let j = i; while (j < rev.length && rev[j][1] === 'L') j++;
    out.push(...rev.slice(i, j).reverse().map((x) => x[0]));
    i = j;
  }
  return out;
}

/** The fonts (beyond the standard ones) a set of strings needs */
export function fontsNeeded(texts, { bold = false } = {}) {
  const need = new Set();
  for (const t of texts) {
    if (t === null || t === undefined) continue;
    for (const ch of String(t)) {
      const s = scriptOf(ch.codePointAt(0));
      if (s !== 'std') need.add(s === 'base' && bold ? 'baseBold' : s);
    }
  }
  return need;
}

export function createTextKit(pdf, { regular, bold }) {
  pdf.registerFontkit(fontkit);
  const embedded = new Map(); // key -> PDFFont

  const fontFor = (script, isBold) => {
    if (script === 'std') return isBold ? bold : regular;
    if (script === 'base' && isBold && embedded.has('baseBold')) return embedded.get('baseBold');
    return embedded.get(script) || embedded.get('base') || null;
  };

  /** Runs with their fonts; characters no font covers become '?' in the standard font */
  function runs(text, isBold) {
    return splitRuns(String(text ?? '')).map((r) => {
      const font = fontFor(r.script, isBold);
      if (font) return { ...r, font };
      return { ...r, script: 'std', font: isBold ? bold : regular, text: r.text.replace(/[^\x20-\x7e\xa0-\xff]/g, '?') };
    });
  }

  function runWidth(r, size) {
    try { return r.font.widthOfTextAtSize(r.text, size); } catch { return r.text.length * size * 0.5; }
  }

  return {
    /** Embeds the fonts these strings need. Safe to call more than once. */
    async prepare(texts, { boldToo = true } = {}) {
      const need = fontsNeeded(texts);
      if (boldToo && need.has('base')) need.add('baseBold');
      for (const key of need) {
        if (embedded.has(key)) continue;
        if (!loader) throw new Error('No font loader set for non-Latin text');
        let bytes;
        try { bytes = await loader(key); } catch (e) { console.warn(`Font ${key} not available`, e); continue; }
        if (!bytes) continue;
        embedded.set(key, await pdf.embedFont(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes), { subset: true }));
      }
    },
    /** True when the text can be drawn with the standard fonts alone */
    isPlain(text) { for (const ch of String(text ?? '')) if (!isWinAnsi(ch.codePointAt(0))) return false; return true; },
    width(text, size, isBold = false) { return runs(text, isBold).reduce((w, r) => w + runWidth(r, size), 0); },
    /** Draws one line at (x, y), optionally rotated (degrees) about that point. Returns its width. */
    draw(page, text, { x, y, size, color, bold: isBold = false, opacity, rotate = 0 }) {
      const t = (rotate * Math.PI) / 180;
      const cos = Math.cos(t);
      const sin = Math.sin(t);
      let pen = 0;
      for (const r of runs(text, isBold)) {
        if (!r.text) continue;
        const ox = x + pen * cos;
        const oy = y + pen * sin;
        if (r.script === 'std') {
          page.drawText(r.text, { x: ox, y: oy, size, font: r.font, color, rotate: degrees(rotate), ...(opacity !== undefined ? { opacity } : {}) });
        } else {
          drawShaped(page, r.font, r.text, { x: ox, y: oy, size, color, opacity, cos, sin });
        }
        pen += runWidth(r, size);
      }
      return pen;
    },
    /** The embedded font for a string's main non-Latin script (for form fields), or null */
    fontForText(text) {
      const need = [...fontsNeeded([text])].filter((k) => k !== 'baseBold');
      const key = need.find((k) => k !== 'base') || need[0];
      return key ? embedded.get(key) || null : null;
    },
    /** Fonts that were embedded (for tests) */
    embeddedKeys: () => [...embedded.keys()],
  };
}

/**
 * Draws a shaped run glyph by glyph, applying the font's own mark and
 * kerning offsets (pdf-lib on its own uses advances only, which leaves
 * Arabic and Indic vowel signs in the wrong place).
 */
function drawShaped(page, font, text, { x, y, size, color, opacity, cos = 1, sin = 0 }) {
  const fk = font.embedder?.font;
  if (!fk?.layout) { page.drawText(text, { x, y, size, font, color }); return; }
  const run = fk.layout(text, font.embedder.fontFeatures);
  const hex = font.encodeText(text).asString(); // same layout, subset glyph codes
  const scale = size / fk.unitsPerEm;
  page.setFont(font);
  const key = page.fontKey;
  const gs = opacity !== undefined && page.maybeEmbedGraphicsState ? page.maybeEmbedGraphicsState({ opacity }) : null;
  const ops = [pushGraphicsState(), ...(gs ? [setGraphicsState(gs)] : []), ...(color ? [setFillingColor(color)] : []), beginText(), setFontAndSize(key, size)];
  let pen = 0;
  run.glyphs.forEach((g, i) => {
    const p = run.positions[i] || { xAdvance: g.advanceWidth, xOffset: 0, yOffset: 0 };
    const code = hex.slice(i * 4, i * 4 + 4);
    if (code) {
      const gx = (pen + p.xOffset) * scale;
      const gy = p.yOffset * scale;
      ops.push(setTextMatrix(cos, sin, -sin, cos, x + gx * cos - gy * sin, y + gx * sin + gy * cos));
      ops.push(showText(PDFHexString.of(code)));
    }
    pen += p.xAdvance;
  });
  ops.push(endText(), popGraphicsState());
  page.pushOperators(...ops);
}
