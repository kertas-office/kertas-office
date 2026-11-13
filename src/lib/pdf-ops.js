// PDF engine: builds the saved PDF from the page model with pdf-lib (MIT, the
// @cantoo fork which adds AES encryption). Pure module: no DOM, runs in Node tests.
//
// Page model entry:
//   { src, index, rotate, annots: [] }        a page taken from a source document
//   { kind: 'blank', width, height, rotate, annots }   a new empty page
//
// Annotations use the source page's user space (PDF points, origin bottom-left).
// `rot` is the page's on-screen rotation when the item was placed, so the item is
// drawn upright relative to how the user saw the page.
//   text      { x, y, size, color, text, rot }               baseline-left anchor
//   highlight { x1, y1, x2, y2, color }                      area highlight
//   rect      { x1, y1, x2, y2, color, width }
//   ellipse   { x1, y1, x2, y2, color, width }
//   line      { x1, y1, x2, y2, color, width, arrow }
//   ink       { points: [[x, y], ...], color, width }
//   image     { x, y, w, h, rot, png | img, fmt: 'png'|'jpg' } bottom-left anchor as seen
//   stamp     { x, y, w, h, rot, label, sub, color }         bottom-left anchor as seen
//   markup    { style: 'highlight'|'underline'|'strike', rects: [[x1,y1,x2,y2]], color }
//   note      { x, y, text, color, author }                  real PDF comment (top-left anchor)
//   redact    { x1, y1, x2, y2 }                             needs opts.rasters for that page
//   field     { kind, x1, y1, x2, y2, name, options, value, size, required, rot }   a new form field
//             kind: 'text'|'multiline'|'checkbox'|'radio'|'dropdown'|'list'; for radio, name is the
//             group and value this button's choice
//
// buildPdf(sources, pages, opts)
//   opts.title, opts.passwords { srcId: password }
//   opts.formValues { srcId: { fieldName: value } }, opts.flattenForms
//   opts.watermark { text, size, color, opacity, angle }
//   opts.headerFooter { header: [l,c,r], footer: [l,c,r], size, start, margin, file }
//   opts.outline [{ title, page, children }]                 page = output index
//   opts.rasters { pageIndexInModel: { bytes, fmt } }        flattened image of a redacted page
//   opts.encrypt { userPassword, ownerPassword, permissions }
//   opts.compress { images: 'strong'|'balanced'|'light'|null, recodeImage }   see pdf-compress.js
//   opts.layers { order: [key], existing: [{ key, src, id, name, origName, visible, deleted }], created: [{ key, name, visible }] }
//     items carry `layer: key`; a page's `layerMoves: { key: [a,b,c,d,e,f] }` moves an existing layer's content
import { createTextKit, isWinAnsi, loadFontBytes } from './pdf-text.js';
import { removeTextInBoxes } from './pdf-textedit.js';
import { fontInfo } from './pdf-convert.js';
import { findFont, facePath } from './font-catalog.js';
import { compressDocument } from './pdf-compress.js';
import { refFromId, createLayer, layerResourceName, beginLayer, endLayer, rewriteLayerContent, writeLayerCatalog, layersOnPages } from './pdf-layers.js';
import {
  pushGraphicsState, popGraphicsState, setTextRenderingMode, TextRenderingMode, setCharacterSqueeze,
  PDFDocument, StandardFonts, rgb, degrees, LineCapStyle, PDFName, PDFHexString, PDFString,
  PDFTextField, PDFCheckBox, PDFRadioGroup, PDFDropdown, PDFOptionList, PDFNumber, PDFArray, PDFDict,
} from '@cantoo/pdf-lib';

// ------------------------------------------------------------------ small helpers
export function hexToRgb(hex = '#000000') {
  const h = String(hex).replace('#', '');
  const n = parseInt(h.length === 3 ? h.split('').map((c) => c + c).join('') : h, 16) || 0;
  return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
}
const rgbArr = (hex) => { const c = hexToRgb(hex); return [c.red, c.green, c.blue]; };

// Standard fonts only cover WinAnsi; replace anything else so saving never fails.
// (Text we draw goes through pdf-text.js, which embeds fonts for other scripts.)
export function toWinAnsi(text) {
  return [...String(text)].map((ch) => {
    const code = ch.charCodeAt(0);
    if (ch === '\n' || (code >= 32 && code <= 126) || (code >= 160 && code <= 255)) return ch;
    const map = { '‘': "'", '’': "'", '“': '"', '”': '"', '–': '-', '\u2014': '-', '•': '*', '…': '...', '€': 'EUR' };
    return map[ch] ?? '?';
  }).join('');
}

/** "1-3, 5, 8-" -> [0,1,2,4,7,8,...] (0-based, in the order given, no duplicates) */
export function parseRanges(spec, total) {
  const out = [];
  const seen = new Set();
  const parts = String(spec).split(/[,;\s]+/).filter(Boolean);
  if (!parts.length) throw new Error('Enter pages, for example 1-3, 5');
  for (const part of parts) {
    const m = part.match(/^(\d*)\s*-\s*(\d*)$/) || part.match(/^(\d+)$/);
    if (!m) throw new Error(`"${part}" is not a page or range`);
    let a; let b;
    if (m.length === 2) { a = b = +m[1]; } else { a = m[1] ? +m[1] : 1; b = m[2] ? +m[2] : total; }
    if (a < 1 || b < 1 || a > total || b > total) throw new Error(`Page ${Math.max(a, b)} is outside 1-${total}`);
    const step = a <= b ? 1 : -1;
    for (let i = a; step > 0 ? i <= b : i >= b; i += step) if (!seen.has(i)) { seen.add(i); out.push(i - 1); }
  }
  return out;
}

/** Splits page indices 0..total-1 into chunks of n */
export function splitEvery(total, n) {
  const size = Math.max(1, Math.floor(n));
  const out = [];
  for (let i = 0; i < total; i += size) out.push(Array.from({ length: Math.min(size, total - i) }, (_, k) => i + k));
  return out;
}

/** Fills {n} {total} {date} {file} tokens */
export function fillTokens(tpl, { n, total, date, file }) {
  return String(tpl || '')
    .replace(/\{n\}/gi, n).replace(/\{total\}/gi, total)
    .replace(/\{date\}/gi, date).replace(/\{file\}/gi, file || '');
}

// Visual frame of a page: lets us place things "as seen on screen" whatever the /Rotate
function visualFrame(page) {
  const box = page.getCropBox();
  const r = ((page.getRotation().angle % 360) + 360) % 360;
  const { x: bx, y: by, width: W, height: H } = box;
  const swap = r === 90 || r === 270;
  const vw = swap ? H : W;
  const vh = swap ? W : H;
  const toUser = (vx, vy) => {
    switch (r) {
      case 90: return [bx + W - vy, by + vx];
      case 180: return [bx + W - vx, by + H - vy];
      case 270: return [bx + vy, by + H - vx];
      default: return [bx + vx, by + vy];
    }
  };
  return { vw, vh, rot: r, toUser };
}

const rotVec = (x, y, deg) => {
  const t = (deg * Math.PI) / 180;
  return [x * Math.cos(t) - y * Math.sin(t), x * Math.sin(t) + y * Math.cos(t)];
};

// ------------------------------------------------------------------ loading
// After decrypting, the old cross-reference stream and /Encrypt dictionary are still
// in the object table; written back out they make the file look encrypted again.
function scrubSecurity(doc) {
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    const dict = obj?.dict || obj;
    const type = dict?.get?.(PDFName.of('Type'))?.toString?.();
    const filter = dict?.get?.(PDFName.of('Filter'))?.toString?.();
    if (type === '/XRef' || (filter === '/Standard' && dict.has(PDFName.of('O')) && dict.has(PDFName.of('U')))) doc.context.delete(ref);
  }
  delete doc.context.trailerInfo.Encrypt;
}

async function loadOne(bytes, password) {
  const tries = password !== undefined ? [password] : [undefined, ''];
  let lastErr;
  for (const pw of tries) {
    try {
      const doc = await PDFDocument.load(bytes, { updateMetadata: false, ...(pw !== undefined ? { password: pw } : {}) });
      if (pw !== undefined) scrubSecurity(doc);
      return doc;
    } catch (err) { lastErr = err; }
  }
  if (/password/i.test(String(lastErr?.message))) throw new Error('The password for this PDF is incorrect.');
  if (/encrypt/i.test(String(lastErr?.message))) throw new Error('This PDF is password-protected. Reopen it with its password to save changes.');
  throw lastErr;
}

export async function loadSources(sources, passwords = {}) {
  const docs = {};
  for (const [id, bytes] of Object.entries(sources)) docs[id] = await loadOne(bytes, passwords[id]);
  return docs;
}

// ------------------------------------------------------------------ forms
/** Lists fillable fields with their widget positions */
export async function listFormFields(bytes, password) {
  const doc = await loadOne(bytes, password);
  let form;
  try { form = doc.getForm(); } catch { return []; }
  const pageIndex = new Map(doc.getPages().map((p, i) => [p.ref.toString(), i]));
  const out = [];
  for (const field of form.getFields()) {
    let type = null;
    if (field instanceof PDFTextField) type = 'text';
    else if (field instanceof PDFCheckBox) type = 'checkbox';
    else if (field instanceof PDFRadioGroup) type = 'radio';
    else if (field instanceof PDFDropdown) type = 'dropdown';
    else if (field instanceof PDFOptionList) type = 'list';
    if (!type) continue;
    const info = { name: field.getName(), type, readOnly: field.isReadOnly(), widgets: [] };
    if (type === 'text') { info.value = field.getText() ?? ''; info.multiline = field.isMultiline(); info.maxLength = field.getMaxLength(); }
    if (type === 'checkbox') info.value = field.isChecked();
    if (type === 'radio') { info.value = field.getSelected() ?? null; info.options = field.getOptions(); }
    if (type === 'dropdown' || type === 'list') { info.value = field.getSelected(); info.options = field.getOptions(); info.editable = type === 'dropdown' && field.isEditable(); }
    const widgets = field.acroField.getWidgets();
    widgets.forEach((w, wi) => {
      let page;
      try { page = form.findWidgetPage(w); } catch { page = undefined; }
      const idx = page ? pageIndex.get(page.ref.toString()) : undefined;
      if (idx === undefined) return;
      const r = w.getRectangle();
      const on = w.getOnValue?.();
      info.widgets.push({
        page: idx, rect: [r.x, r.y, r.x + r.width, r.y + r.height],
        onValue: type === 'radio' ? (info.options?.[wi] ?? on?.decodeText?.() ?? null) : undefined,
      });
    });
    if (info.widgets.length) out.push(info);
  }
  return out;
}

async function applyForm(doc, values, flatten) {
  let form;
  try { form = doc.getForm(); } catch { return; }
  let kit = null; // made on demand, for values in other scripts
  for (const [name, value] of Object.entries(values || {})) {
    let field;
    try { field = form.getField(name); } catch { continue; }
    if (!field || field.isReadOnly?.()) continue;
    try {
      if (field instanceof PDFTextField) {
        const text = String(value ?? '');
        if (isPlainText(text)) { field.setText(text); continue; }
        kit ||= createTextKit(doc, { regular: await doc.embedFont(StandardFonts.Helvetica), bold: await doc.embedFont(StandardFonts.HelveticaBold) });
        await kit.prepare([text], { boldToo: false });
        const font = kit.fontForText(text);
        field.setText(font ? text : toWinAnsi(text));
        if (font) field.updateAppearances(font);
      } else if (field instanceof PDFCheckBox) { if (value) field.check(); else field.uncheck(); }
      else if (field instanceof PDFRadioGroup) { if (value) field.select(value); else field.clear(); }
      else if (field instanceof PDFDropdown || field instanceof PDFOptionList) {
        const list = (Array.isArray(value) ? value : [value]).filter((v) => v !== null && v !== undefined && v !== '');
        if (list.length) field.select(list.length === 1 ? list[0] : list); else field.clear();
      }
    } catch (e) { console.warn(`Field ${name}:`, e); /* a malformed field should not block saving */ }
  }
  if (flatten) {
    try { form.flatten(); } catch { /* some forms cannot be flattened */ }
  }
}

const isPlainText = (t) => [...t].every((ch) => isWinAnsi(ch.codePointAt(0)));

// ------------------------------------------------------------------ annotations
const STAMP_PAD = 6;

function drawAnnot(doc, page, a, fonts, imageCache) {
  const color = hexToRgb(a.color || '#000000');
  switch (a.type) {
    case 'text': {
      const text = String(a.text || '');
      if (!text.trim()) return;
      text.split('\n').forEach((line, k) => {
        const [dx, dy] = rotVec(0, -k * a.size * 1.2, a.rot || 0);
        fonts.kit.draw(page, line, { x: a.x + dx, y: a.y + dy, size: a.size, color, rotate: a.rot || 0 });
      });
      return;
    }
    case 'ocrText': {
      // Recognised words, drawn invisibly over the scan so they can be searched and copied
      page.pushOperators(pushGraphicsState(), setTextRenderingMode(TextRenderingMode.Invisible));
      for (const w of a.words || []) {
        if (!w.str) continue;
        const natural = fonts.kit.width(w.str, w.size) || 1;
        const squeeze = Math.max(20, Math.min(400, (w.w / natural) * 100));
        page.pushOperators(setCharacterSqueeze(squeeze));
        fonts.kit.draw(page, w.str, { x: w.x, y: w.y, size: w.size, rotate: w.rot || 0 });
      }
      page.pushOperators(setCharacterSqueeze(100), popGraphicsState());
      return;
    }
    case 'highlight':
      page.drawRectangle({ x: Math.min(a.x1, a.x2), y: Math.min(a.y1, a.y2), width: Math.abs(a.x2 - a.x1), height: Math.abs(a.y2 - a.y1), color, opacity: 0.35 });
      return;
    case 'rect':
      page.drawRectangle({ x: Math.min(a.x1, a.x2), y: Math.min(a.y1, a.y2), width: Math.abs(a.x2 - a.x1), height: Math.abs(a.y2 - a.y1), borderColor: color, borderWidth: a.width || 2 });
      return;
    case 'ellipse':
      page.drawEllipse({ x: (a.x1 + a.x2) / 2, y: (a.y1 + a.y2) / 2, xScale: Math.abs(a.x2 - a.x1) / 2, yScale: Math.abs(a.y2 - a.y1) / 2, borderColor: color, borderWidth: a.width || 2 });
      return;
    case 'line': {
      const w = a.width || 2;
      page.drawLine({ start: { x: a.x1, y: a.y1 }, end: { x: a.x2, y: a.y2 }, thickness: w, color, lineCap: LineCapStyle.Round });
      if (a.arrow) {
        const ang = Math.atan2(a.y2 - a.y1, a.x2 - a.x1);
        const len = Math.max(8, w * 4);
        for (const d of [Math.PI * 0.82, -Math.PI * 0.82]) {
          page.drawLine({ start: { x: a.x2, y: a.y2 }, end: { x: a.x2 + len * Math.cos(ang + d), y: a.y2 + len * Math.sin(ang + d) }, thickness: w, color, lineCap: LineCapStyle.Round });
        }
      }
      return;
    }
    case 'ink': {
      const pts = a.points || [];
      for (let i = 1; i < pts.length; i++) {
        page.drawLine({ start: { x: pts[i - 1][0], y: pts[i - 1][1] }, end: { x: pts[i][0], y: pts[i][1] }, thickness: a.width || 2, color, lineCap: LineCapStyle.Round });
      }
      return;
    }
    case 'markup':
      for (const [x1, y1, x2, y2] of a.rects || []) {
        const x = Math.min(x1, x2); const y = Math.min(y1, y2); const w = Math.abs(x2 - x1); const h = Math.abs(y2 - y1);
        if (a.style === 'highlight') page.drawRectangle({ x, y, width: w, height: h, color, opacity: 0.38 });
        else {
          const ly = a.style === 'strike' ? y + h * 0.48 : y + h * 0.08;
          page.drawLine({ start: { x, y: ly }, end: { x: x + w, y: ly }, thickness: Math.max(0.8, h / 14), color });
        }
      }
      return;
    case 'stamp': {
      const rot = a.rot || 0;
      const [ox, oy] = rotVec(0, a.h, rot);
      const r = 5;
      const path = `M ${r} 0 H ${a.w - r} Q ${a.w} 0 ${a.w} ${r} V ${a.h - r} Q ${a.w} ${a.h} ${a.w - r} ${a.h} H ${r} Q 0 ${a.h} 0 ${a.h - r} V ${r} Q 0 0 ${r} 0 Z`;
      page.drawSvgPath(path, { x: a.x + ox, y: a.y + oy, rotate: degrees(rot), borderColor: color, borderWidth: 2.2, color, opacity: 0.06, borderOpacity: 0.9 });
      const label = String(a.label || '');
      const ls = a.labelSize || 16;
      const lw = fonts.kit.width(label, ls, true);
      const [lx, ly] = rotVec((a.w - lw) / 2, a.h - STAMP_PAD - ls * 0.74, rot);
      fonts.kit.draw(page, label, { x: a.x + lx, y: a.y + ly, size: ls, bold: true, color, rotate: rot, opacity: 0.9 });
      if (a.sub) {
        const sub = String(a.sub);
        const ss = ls * 0.5;
        const sw = fonts.kit.width(sub, ss);
        const [sx, sy] = rotVec((a.w - sw) / 2, STAMP_PAD + ss * 0.2, rot);
        fonts.kit.draw(page, sub, { x: a.x + sx, y: a.y + sy, size: ss, color, rotate: rot, opacity: 0.9 });
      }
      return;
    }
    case 'note':
      addNoteAnnotation(doc, page, a);
      return;
    case 'redact':
      // The page was replaced by a flattened image that already contains the black box;
      // draw it again so the mark is visible even if the raster is slightly offset.
      page.drawRectangle({ x: Math.min(a.x1, a.x2), y: Math.min(a.y1, a.y2), width: Math.abs(a.x2 - a.x1), height: Math.abs(a.y2 - a.y1), color: rgb(0, 0, 0) });
      return;
    case 'image': {
      const bytes = a.img || a.png;
      const key = bytes;
      return (async () => {
        let img = imageCache.get(key);
        if (!img) { img = a.fmt === 'jpg' ? await doc.embedJpg(bytes) : await doc.embedPng(bytes); imageCache.set(key, img); }
        page.drawImage(img, { x: a.x, y: a.y, width: a.w, height: a.h, rotate: degrees(a.rot || 0) });
      })();
    }
    default:
  }
  return undefined;
}

/** Stamp size for a label, so the UI and the PDF agree */
export function stampSize(label, sub, labelSize = 16, widthOf) {
  const lw = widthOf ? widthOf(label, labelSize) : label.length * labelSize * 0.68;
  const sw = sub ? (widthOf ? widthOf(sub, labelSize * 0.5, false) : sub.length * labelSize * 0.28) : 0;
  return { w: Math.max(lw, sw) + STAMP_PAD * 2 + 10, h: labelSize * 0.74 + (sub ? labelSize * 0.5 + 4 : 0) + STAMP_PAD * 2 };
}

// A real comment that shows in Acrobat, Preview, Edge and Chrome, with its own icon
function addNoteAnnotation(doc, page, a) {
  const size = 20;
  const [r, g, b] = rgbArr(a.color || '#facc15');
  const ap = doc.context.flateStream(
    `q ${r} ${g} ${b} rg 0.25 0.25 0.25 RG 0.8 w 1 1 18 18 re B 0.25 0.25 0.25 RG 1 w 4 13.5 m 16 13.5 l S 4 10 m 16 10 l S 4 6.5 m 12 6.5 l S Q`,
    { Type: 'XObject', Subtype: 'Form', BBox: [0, 0, size, size] },
  );
  const apRef = doc.context.register(ap);
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const stamp = `D:${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`;
  const annot = doc.context.obj({
    Type: 'Annot', Subtype: 'Text', Rect: [a.x, a.y - size, a.x + size, a.y],
    Contents: PDFHexString.fromText(a.text || ''), T: PDFHexString.fromText(a.author || 'Kertas Office'),
    Name: 'Comment', C: [r, g, b], F: 4 | 8 | 16, Open: false, M: PDFString.of(stamp),
    AP: { N: apRef },
  });
  page.node.addAnnot(doc.context.register(annot));
}

// ------------------------------------------------------------------ editing existing text
const STD_BY_FAMILY = { 'Times New Roman': ['TimesRoman', 'TimesRomanBold', 'TimesRomanItalic', 'TimesRomanBoldItalic'], 'Courier New': ['Courier', 'CourierBold', 'CourierOblique', 'CourierBoldOblique'] };
async function applyTextEdits(doc, page, edits, fonts) {
  const found = removeTextInBoxes(doc, page, edits.map((a) => a.box));
  const cache = fonts.editFonts || (fonts.editFonts = new Map());
  for (const [k, a] of edits.entries()) {
    const f = found[k];
    const info = fontInfo(f.font || '');
    const bold = a.bold ?? info.bold;
    const italic = a.italic ?? info.italic;
    const size = a.size || f.size || 11;
    const color = a.color ? hexToRgb(a.color) : f.color ? rgb(...f.color.map((v) => Math.max(0, Math.min(1, v)))) : rgb(0, 0, 0);
    const x = a.x ?? f.x ?? a.box.x0;
    const y = a.y ?? f.y ?? a.box.y0 + size * 0.2;
    const text = String(a.text ?? '');
    if (!text) continue;
    // The original font where we ship it (or its metric twin), else the closest standard font
    const family = a.family || info.family;
    const key = `${family}|${bold}|${italic}`;
    if (!cache.has(key)) {
      let font = null;
      const cat = family && findFont(family);
      if (cat) {
        const want = bold && italic ? ['boldItalic', 'bold', 'italic', 'regular'] : bold ? ['bold', 'regular'] : italic ? ['italic', 'regular'] : ['regular'];
        for (const face of want) {
          const bytes = await loadFontBytes(`face:${facePath(cat, face)}`).catch(() => null);
          if (bytes) { try { font = await doc.embedFont(bytes, { subset: true }); break; } catch { /* next */ } }
        }
      }
      if (!font) {
        const std = STD_BY_FAMILY[family] || ['Helvetica', 'HelveticaBold', 'HelveticaOblique', 'HelveticaBoldOblique'];
        font = await doc.embedFont(StandardFonts[std[(bold ? 1 : 0) + (italic ? 2 : 0)]]);
      }
      cache.set(key, font);
    }
    const font = cache.get(key);
    const plain = [...text].every((ch) => isWinAnsi(ch.codePointAt(0)));
    if (plain || font.name && !/Helvetica|Times|Courier/.test(font.name)) {
      try { page.drawText(text, { x, y, size, font, color, rotate: degrees(a.rot || 0) }); continue; } catch { /* glyphs missing: fall through */ }
    }
    await fonts.kit.prepare([text]);
    fonts.kit.draw(page, text, { x, y, size, color, bold, rotate: a.rot || 0 });
  }
}

// ------------------------------------------------------------------ page decorations
function drawWatermark(page, wm, kit) {
  const text = String(wm.text || '');
  if (!text.trim()) return;
  const { vw, vh, rot, toUser } = visualFrame(page);
  const size = wm.size || Math.min(vw, vh) / Math.max(4, [...text].length * 0.62);
  const angle = wm.angle ?? 45;
  const tw = kit.width(text, size, true);
  const cap = size * 0.7;
  // Start point so the text is centred on the page, in the visual frame
  const [dx, dy] = rotVec(-tw / 2, -cap / 2, angle);
  const [ux, uy] = toUser(vw / 2 + dx, vh / 2 + dy);
  kit.draw(page, text, { x: ux, y: uy, size, bold: true, color: hexToRgb(wm.color || '#94a3b8'), opacity: wm.opacity ?? 0.25, rotate: rot + angle });
}

function drawHeaderFooter(page, hf, kit, ctx) {
  const size = hf.size || 9;
  const margin = hf.margin ?? 28;
  const { vw, vh, rot, toUser } = visualFrame(page);
  const color = hexToRgb(hf.color || '#475569');
  const rows = [[hf.header, vh - margin - size * 0.7], [hf.footer, margin]];
  for (const [cells, vy] of rows) {
    (cells || []).forEach((tpl, i) => {
      const text = fillTokens(tpl, ctx);
      if (!text.trim()) return;
      const tw = kit.width(text, size);
      const vx = i === 0 ? margin : i === 1 ? (vw - tw) / 2 : vw - margin - tw;
      const [x, y] = toUser(vx, vy);
      kit.draw(page, text, { x, y, size, color, rotate: rot });
    });
  }
}

// ------------------------------------------------------------------ outline
function writeOutline(doc, items) {
  const pages = doc.getPages();
  const ctx = doc.context;
  const build = (list, parentRef) => {
    const refs = list.map(() => ctx.nextRef());
    let total = 0;
    list.forEach((it, i) => {
      const page = pages[Math.max(0, Math.min(pages.length - 1, it.page))];
      const dict = ctx.obj({ Title: PDFHexString.fromText(it.title || 'Untitled'), Parent: parentRef });
      dict.set(PDFName.of('Dest'), ctx.obj([page.ref, PDFName.of('XYZ'), null, null, null]));
      if (i > 0) dict.set(PDFName.of('Prev'), refs[i - 1]);
      if (i < list.length - 1) dict.set(PDFName.of('Next'), refs[i + 1]);
      const kids = it.children?.length ? build(it.children, refs[i]) : null;
      if (kids) {
        dict.set(PDFName.of('First'), kids.first);
        dict.set(PDFName.of('Last'), kids.last);
        dict.set(PDFName.of('Count'), PDFNumber.of(-kids.count)); // collapsed
      }
      ctx.assign(refs[i], dict);
      total += 1;
    });
    return { first: refs[0], last: refs[refs.length - 1], count: total };
  };
  const rootRef = ctx.nextRef();
  if (!items?.length) { doc.catalog.delete(PDFName.of('Outlines')); return; }
  const top = build(items, rootRef);
  ctx.assign(rootRef, ctx.obj({ Type: 'Outlines', First: top.first, Last: top.last, Count: PDFNumber.of(top.count) }));
  doc.catalog.set(PDFName.of('Outlines'), rootRef);
}

// ------------------------------------------------------------------ main
function isInPlace(pages, mainId) {
  let last = -1;
  for (const p of pages) {
    if (p.kind || p.src !== mainId || p.index <= last) return false;
    last = p.index;
  }
  return true;
}

async function replaceWithRaster(doc, at, original, raster) {
  const box = original.getCropBox();
  const angle = original.getRotation().angle;
  doc.removePage(at);
  const np = doc.insertPage(at, [box.width, box.height]);
  np.setMediaBox(box.x, box.y, box.width, box.height);
  const img = raster.fmt === 'png' ? await doc.embedPng(raster.bytes) : await doc.embedJpg(raster.bytes);
  np.drawImage(img, { x: box.x, y: box.y, width: box.width, height: box.height });
  np.setRotation(degrees(angle));
  return np;
}

export async function buildPdf(sources, pages, opts = {}) {
  const docs = await loadSources(sources, opts.passwords);
  const mainId = Object.keys(sources)[0];
  const inPlace = isInPlace(pages, mainId);
  const formValues = opts.formValues || {};

  let out;
  const outPages = [];
  if (inPlace) {
    // Edit the original document directly: keeps forms, links, bookmarks and metadata
    out = docs[mainId];
    await applyForm(out, formValues[mainId], !!opts.flattenForms);
    const keep = new Set(pages.map((p) => p.index));
    for (let i = out.getPageCount() - 1; i >= 0; i--) if (!keep.has(i)) out.removePage(i);
    out.getPages().forEach((pg) => outPages.push(pg));
  } else {
    // Pages are merged or reordered: assemble a new document. Form fields cannot
    // follow copied pages, so their values are flattened into the page first.
    for (const [id, d] of Object.entries(docs)) await applyForm(d, formValues[id], true);
    out = await PDFDocument.create({ updateMetadata: false });
    const first = docs[mainId];
    if (first) {
      const t = first.getTitle(); if (t) out.setTitle(t);
      const a = first.getAuthor(); if (a) out.setAuthor(a);
    }
    for (const p of pages) {
      if (p.kind === 'blank') {
        outPages.push(out.addPage([p.width || 595.28, p.height || 841.89]));
        continue;
      }
      const src = docs[p.src];
      if (!src) throw new Error(`Missing source ${p.src}`);
      const [page] = await out.copyPages(src, [p.index]);
      outPages.push(out.addPage(page));
    }
  }

  out.setCreationDate(out.getCreationDate?.() || new Date());
  out.setModificationDate(new Date());
  out.setProducer('Kertas Office');
  out.setCreator('Kertas Office');
  if (opts.title) out.setTitle(opts.title);

  const fonts = { r: await out.embedFont(StandardFonts.Helvetica), b: await out.embedFont(StandardFonts.HelveticaBold) };
  fonts.kit = createTextKit(out, { regular: fonts.r, bold: fonts.b });
  // Embed fonts for any non-Latin text we are about to draw
  const texts = [];
  for (const p of pages) for (const a of p.annots || []) texts.push(a.text, a.label, a.sub, ...(a.words || []).map((w) => w.str));
  if (opts.watermark?.text) texts.push(opts.watermark.text);
  for (const cell of [...(opts.headerFooter?.header || []), ...(opts.headerFooter?.footer || [])]) texts.push(cell, opts.headerFooter?.file);
  await fonts.kit.prepare(texts.filter(Boolean));
  const imageCache = new Map();
  const total = pages.length;

  // ---- layers: find (or make) each layer in the output document
  const L = opts.layers;
  const layerRef = {}; const layerRefKey = {};
  const removedLayers = [];
  if (L) {
    const byName = inPlace ? null : layersOnPages(out, outPages);
    for (const l of L.existing || []) {
      const ref = inPlace && l.src === mainId ? refFromId(out, l.id) : byName?.[l.origName ?? l.name];
      if (!ref) continue;
      if (l.deleted) removedLayers.push(ref); else layerRef[l.key] = ref;
      layerRefKey[l.key] = `${ref.objectNumber}R${ref.generationNumber || ''}`;
    }
    for (const l of L.created || []) layerRef[l.key] = createLayer(out, l.name);
  }
  const deletedKeys = new Set((L?.existing || []).filter((l) => l.deleted).map((l) => l.key));
  const formsDone = new Set(); // forms shared by several pages are changed once
  const newFields = [];
  const date = new Date().toLocaleDateString('en-GB');

  for (let i = 0; i < pages.length; i++) {
    const p = pages[i];
    let page = outPages[i];
    if (p.rotate) page.setRotation(degrees((((page.getRotation().angle + p.rotate) % 360) + 360) % 360));
    const raster = opts.rasters?.[i];
    if (raster) page = outPages[i] = await replaceWithRaster(out, i, page, raster);

    // Existing layers: deleted ones are taken out, moved ones get their transform
    if (L) {
      const changes = new Map();
      for (const key of deletedKeys) if (layerRefKey[key]) changes.set(layerRefKey[key], { remove: true });
      for (const [key, m] of Object.entries(p.layerMoves || {})) if (layerRefKey[key] && !deletedKeys.has(key)) changes.set(layerRefKey[key], { matrix: m });
      if (changes.size) rewriteLayerContent(out, page, changes, formsDone);
    }

    // Edited text: the old glyphs are taken out of the page, then the new words drawn in their place
    const edits = (p.annots || []).filter((a) => a.type === 'editText');
    if (edits.length) await applyTextEdits(out, page, edits, fonts);

    for (const a of p.annots || []) {
      if (a.type === 'editText' || (a.layer && deletedKeys.has(a.layer))) continue;
      if (a.type === 'field') { newFields.push({ page, a }); continue; }
      const ref = a.layer && layerRef[a.layer];
      if (ref) page.pushOperators(beginLayer(layerResourceName(out, page, ref)));
      const pending = drawAnnot(out, page, a, fonts, imageCache);
      if (pending) await pending;
      if (ref) {
        page.pushOperators(endLayer());
        if (a.type === 'note') { const annots = page.node.Annots(); const last = annots?.get(annots.size() - 1); const d = last && out.context.lookup(last); d?.set?.(PDFName.of('OC'), ref); }
      }
    }
    if (opts.watermark?.text) drawWatermark(page, opts.watermark, fonts.kit);
    if (opts.headerFooter) {
      const hf = opts.headerFooter;
      drawHeaderFooter(page, hf, fonts.kit, { n: i + (hf.start ?? 1), total: total + (hf.start ?? 1) - 1, date, file: hf.file });
    }
  }

  if (L) {
    const vis = Object.fromEntries([...(L.existing || []), ...(L.created || [])].map((l) => [l.key, l]));
    writeLayerCatalog(out, {
      layers: (L.order || []).filter((k) => layerRef[k]).map((k) => ({ ref: layerRef[k], name: vis[k]?.name, visible: vis[k]?.visible !== false })),
      removed: removedLayers,
    });
  }
  if (newFields.length) await addFormFields(out, newFields);
  if (opts.outline) writeOutline(out, opts.outline);
  if (Array.isArray(opts.attachments)) await writeAttachments(out, opts.attachments);
  if (opts.compress) await compressDocument(out, opts.compress);
  if (opts.encrypt && (opts.encrypt.userPassword || opts.encrypt.ownerPassword)) {
    const { userPassword = '', ownerPassword, permissions } = opts.encrypt;
    // No restrictions asked for means everything is allowed once the file is open
    const allowAll = { printing: 'highResolution', modifying: true, copying: true, annotating: true, fillingForms: true, contentAccessibility: true, documentAssembly: true };
    out.encrypt({ userPassword, ownerPassword: ownerPassword || userPassword, permissions: permissions || allowAll });
  }
  return out.save({ useObjectStreams: true });
}

// ------------------------------------------------------------------ attached files
const MIME_BY_EXT = { pdf: 'application/pdf', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', csv: 'text/csv', txt: 'text/plain', png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', zip: 'application/zip', xml: 'application/xml', json: 'application/json' };
/** Replaces the document's attached files with the list given ([{ name, data, description }]) */
async function writeAttachments(doc, list) {
  const names = doc.catalog.lookup(PDFName.of('Names'));
  if (names instanceof PDFDict) names.delete(PDFName.of('EmbeddedFiles'));
  doc.catalog.delete(PDFName.of('AF'));
  for (const a of list) {
    if (!a?.data || !a.name) continue;
    const ext = String(a.name).split('.').pop().toLowerCase();
    await doc.attach(a.data, a.name, { mimeType: MIME_BY_EXT[ext] || 'application/octet-stream', description: a.description || undefined, creationDate: a.created || new Date(), modificationDate: a.modified || new Date() });
  }
}

// ------------------------------------------------------------------ new form fields
export const FIELD_KINDS = ['text', 'multiline', 'checkbox', 'radio', 'dropdown', 'list'];

/** Makes real AcroForm fields that every PDF reader can fill in */
async function addFormFields(doc, list) {
  const form = doc.getForm();
  const taken = new Set(form.getFields().map((f) => f.getName()));
  const radios = new Map(); // group name -> PDFRadioGroup made in this save
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const unique = (base) => {
    const clean = toWinAnsi(String(base || 'Field').replace(/\./g, ' ').trim()) || 'Field';
    if (!taken.has(clean)) { taken.add(clean); return clean; }
    let n = 2;
    while (taken.has(`${clean} ${n}`)) n++;
    taken.add(`${clean} ${n}`);
    return `${clean} ${n}`;
  };
  const opts = (items) => (items || []).map((o) => toWinAnsi(String(o).trim())).filter(Boolean);
  for (const { page, a } of list) {
    const x = Math.min(a.x1, a.x2); const y = Math.min(a.y1, a.y2);
    const width = Math.max(4, Math.abs(a.x2 - a.x1)); const height = Math.max(4, Math.abs(a.y2 - a.y1));
    const place = { x, y, width, height, font, rotate: degrees(a.rot || 0), borderWidth: 1, borderColor: rgb(0.45, 0.5, 0.6), backgroundColor: rgb(1, 1, 1) };
    let field;
    try {
      switch (a.kind) {
        case 'checkbox':
          field = form.createCheckBox(unique(a.name || 'Check box'));
          field.addToPage(page, { ...place, backgroundColor: rgb(1, 1, 1) });
          if (a.checked) field.check();
          break;
        case 'radio': {
          const group = String(a.name || 'Choice').trim();
          let rg = radios.get(group);
          if (!rg) { rg = form.createRadioGroup(unique(group)); radios.set(group, rg); }
          const value = toWinAnsi(String(a.value || `Option ${rg.getOptions().length + 1}`));
          rg.addOptionToPage(value, page, place);
          field = rg;
          break;
        }
        case 'dropdown': case 'list': {
          field = a.kind === 'dropdown' ? form.createDropdown(unique(a.name || 'Dropdown')) : form.createOptionList(unique(a.name || 'List'));
          const items = opts(a.options);
          if (items.length) field.setOptions(items);
          if (a.value && items.includes(a.value)) field.select(a.value);
          field.addToPage(page, place);
          if (a.size) field.setFontSize(a.size);
          break;
        }
        default: {
          field = form.createTextField(unique(a.name || 'Text'));
          if (a.kind === 'multiline') field.enableMultiline();
          if (a.value) field.setText(toWinAnsi(a.value));
          field.addToPage(page, place);
          field.setFontSize(a.size || 0); // 0 = fit the box
        }
      }
      if (a.required) field?.enableRequired();
    } catch (e) { console.warn('New field:', e); }
  }
  try { form.updateFieldAppearances(font); } catch { /* appearances are rebuilt by readers anyway */ }
}

// ------------------------------------------------------------------ images to PDF
const PAGE_SIZES = { a4: [595.28, 841.89], letter: [612, 792] };

/** images: [{ bytes, fmt: 'png'|'jpg' }]; pageSize 'fit' | 'a4' | 'letter' */
export async function imagesToPdf(images, { pageSize = 'a4', margin = 24 } = {}) {
  const doc = await PDFDocument.create();
  doc.setProducer('Kertas Office');
  doc.setCreator('Kertas Office');
  for (const im of images) {
    const img = im.fmt === 'png' ? await doc.embedPng(im.bytes) : await doc.embedJpg(im.bytes);
    if (pageSize === 'fit') {
      const page = doc.addPage([img.width * 0.75, img.height * 0.75]);
      page.drawImage(img, { x: 0, y: 0, width: img.width * 0.75, height: img.height * 0.75 });
      continue;
    }
    let [w, h] = PAGE_SIZES[pageSize] || PAGE_SIZES.a4;
    if (img.width > img.height) [w, h] = [h, w];
    const s = Math.min((w - margin * 2) / img.width, (h - margin * 2) / img.height, 1.5);
    const page = doc.addPage([w, h]);
    page.drawImage(img, { x: (w - img.width * s) / 2, y: (h - img.height * s) / 2, width: img.width * s, height: img.height * s });
  }
  return doc.save();
}

/** Reads an outline from pdf-lib for tests and fallbacks: [{ title, page, children }] */
export async function readOutline(bytes, password) {
  const doc = await loadOne(bytes, password);
  const root = doc.catalog.lookup(PDFName.of('Outlines'));
  if (!root) return [];
  const pageIndex = new Map(doc.getPages().map((p, i) => [p.ref.toString(), i]));
  const walk = (first) => {
    const list = [];
    let ref = first;
    let guard = 0;
    while (ref && guard++ < 5000) {
      const item = doc.context.lookup(ref);
      const title = item.lookup(PDFName.of('Title'))?.decodeText?.() ?? '';
      const dest = item.lookup(PDFName.of('Dest'));
      const target = dest instanceof PDFArray ? dest.get(0) : undefined;
      const kids = item.get(PDFName.of('First'));
      list.push({ title, page: target ? pageIndex.get(target.toString()) ?? null : null, children: kids ? walk(kids) : [] });
      ref = item.get(PDFName.of('Next'));
    }
    return list;
  };
  return walk(root.get(PDFName.of('First')));
}
