// Editing the text that is already in a PDF. The page's content stream is
// parsed, every glyph is placed using the font's widths, and the glyphs that
// fall inside the edited boxes are removed for real (not just covered), so the
// old words are gone from copy, search and text extraction. The new text is
// then drawn in their place by pdf-ops.js. Pure module (pdf-lib only).
import { PDFName, PDFArray, PDFDict, PDFNumber, PDFRawStream, PDFRef, decodePDFRawStream } from '@cantoo/pdf-lib';

// ---------------------------------------------------------------- matrices
const mul = (m, n) => [
  m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const I = [1, 0, 0, 1, 0, 0];

// ---------------------------------------------------------------- tokenizer
const WS = new Set([0, 9, 10, 12, 13, 32]);
const DELIM = new Set([40, 41, 60, 62, 91, 93, 123, 125, 47, 37]);

/** Tokens of a content stream (latin1 string): { t: 'num'|'name'|'str'|'hex'|'arr'|'dict'|'op'|'bool'|'null', v, s, e } */
export function tokenize(src) {
  const out = [];
  const n = src.length;
  let i = 0;
  const code = (k) => src.charCodeAt(k);
  const skipWs = () => {
    while (i < n) {
      const c = code(i);
      if (WS.has(c)) i++;
      else if (c === 37) { while (i < n && code(i) !== 10 && code(i) !== 13) i++; } // % comment
      else break;
    }
  };
  const readValue = () => {
    skipWs();
    if (i >= n) return null;
    const s = i;
    const c = src[i];
    if (c === '(') {
      let depth = 0; let str = '';
      i++;
      while (i < n) {
        const ch = src[i];
        if (ch === '\\') {
          const nx = src[i + 1];
          const map = { n: '\n', r: '\r', t: '\t', b: '\b', f: '\f', '(': '(', ')': ')', '\\': '\\' };
          if (map[nx] !== undefined) { str += map[nx]; i += 2; }
          else if (/[0-7]/.test(nx)) { let o = ''; i++; while (o.length < 3 && /[0-7]/.test(src[i])) o += src[i++]; str += String.fromCharCode(parseInt(o, 8) & 255); }
          else if (nx === '\r' || nx === '\n') { i += 2; if (nx === '\r' && src[i] === '\n') i++; }
          else { str += nx || ''; i += 2; }
          continue;
        }
        if (ch === '(') depth++;
        if (ch === ')') { if (depth === 0) { i++; break; } depth--; }
        str += ch; i++;
      }
      return { t: 'str', v: str, s, e: i };
    }
    if (c === '<' && src[i + 1] === '<') {
      i += 2;
      const items = [];
      for (;;) { skipWs(); if (src[i] === '>' && src[i + 1] === '>') { i += 2; break; } const v = readValue(); if (!v) break; items.push(v); }
      return { t: 'dict', v: items, s, e: i };
    }
    if (c === '<') {
      const end = src.indexOf('>', i);
      let hex = src.slice(i + 1, end).replace(/\s+/g, '');
      if (hex.length % 2) hex += '0';
      let str = '';
      for (let k = 0; k < hex.length; k += 2) str += String.fromCharCode(parseInt(hex.slice(k, k + 2), 16));
      i = end + 1;
      return { t: 'str', v: str, s, e: i };
    }
    if (c === '[') {
      i++;
      const items = [];
      for (;;) { skipWs(); if (src[i] === ']') { i++; break; } const v = readValue(); if (!v) break; items.push(v); }
      return { t: 'arr', v: items, s, e: i };
    }
    if (c === '/') {
      i++;
      let name = '';
      while (i < n && !WS.has(code(i)) && !DELIM.has(code(i))) name += src[i++];
      return { t: 'name', v: name, s, e: i };
    }
    let word = '';
    while (i < n && !WS.has(code(i)) && !DELIM.has(code(i))) word += src[i++];
    if (!word) { i++; return readValue(); }
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) return { t: 'num', v: Number(word), s, e: i };
    if (word === 'true' || word === 'false') return { t: 'bool', v: word === 'true', s, e: i };
    if (word === 'null') return { t: 'null', v: null, s, e: i };
    return { t: 'op', v: word, s, e: i };
  };
  while (i < n) {
    const tok = readValue();
    if (!tok) break;
    out.push(tok);
    // Inline images carry binary data between ID and EI
    if (tok.t === 'op' && tok.v === 'ID') {
      i += 1;
      const m = /\sEI(?=[\s]|$)/g;
      m.lastIndex = i;
      const found = m.exec(src);
      const end = found ? found.index : n;
      out.push({ t: 'bin', v: null, s: i, e: end });
      i = end;
    }
  }
  return out;
}

// ---------------------------------------------------------------- fonts
const STD_WIDTH = (base = '') => (/courier/i.test(base) ? 600 : /times/i.test(base) ? 500 : 556);

function fontMetrics(doc, fontDict) {
  const get = (d, k) => { const v = d?.get(PDFName.of(k)); return v instanceof PDFRef ? doc.context.lookup(v) : v; };
  const num = (v) => (v instanceof PDFNumber ? v.asNumber() : Number(v?.asNumber?.() ?? v ?? 0));
  const subtype = get(fontDict, 'Subtype')?.asString?.() || '';
  const base = get(fontDict, 'BaseFont')?.asString?.()?.replace(/^\//, '') || '';
  if (subtype === '/Type0') {
    const desc = get(fontDict, 'DescendantFonts');
    const cid = desc instanceof PDFArray ? doc.context.lookup(desc.get(0)) : null;
    const dw = cid && get(cid, 'DW') ? num(get(cid, 'DW')) : 1000;
    const widths = new Map();
    const W = cid && get(cid, 'W');
    if (W instanceof PDFArray) {
      const arr = W.asArray().map((x) => (x instanceof PDFRef ? doc.context.lookup(x) : x));
      for (let k = 0; k < arr.length;) {
        const first = num(arr[k]);
        if (arr[k + 1] instanceof PDFArray) { arr[k + 1].asArray().forEach((w, j) => widths.set(first + j, num(w instanceof PDFRef ? doc.context.lookup(w) : w))); k += 2; }
        else { const last = num(arr[k + 1]); const w = num(arr[k + 2]); for (let c = first; c <= last; c++) widths.set(c, w); k += 3; }
      }
    }
    return { bytes: 2, width: (c) => widths.get(c) ?? dw, base };
  }
  const first = num(get(fontDict, 'FirstChar') ?? 0);
  const ws = get(fontDict, 'Widths');
  const list = ws instanceof PDFArray ? ws.asArray().map((w) => num(w instanceof PDFRef ? doc.context.lookup(w) : w)) : null;
  const fd = get(fontDict, 'FontDescriptor');
  const missing = fd ? num(get(fd, 'MissingWidth') ?? 0) : 0;
  let scale = 1;
  if (subtype === '/Type3') { const fm = get(fontDict, 'FontMatrix'); if (fm instanceof PDFArray) scale = num(fm.get(0)) * 1000; }
  const std = STD_WIDTH(base);
  return { bytes: 1, width: (c) => ((list ? (list[c - first] ?? (missing || std)) : std) * scale), base };
}

// ---------------------------------------------------------------- page content
function pageContent(doc, page) {
  const contents = page.node.get(PDFName.of('Contents'));
  const refs = contents instanceof PDFArray ? contents.asArray() : contents ? [contents] : [];
  const parts = [];
  for (const r of refs) {
    const s = r instanceof PDFRef ? doc.context.lookup(r) : r;
    if (!s) continue;
    const bytes = s instanceof PDFRawStream ? decodePDFRawStream(s).decode() : s.getContents?.() || new Uint8Array();
    let str = '';
    for (let k = 0; k < bytes.length; k += 0x8000) str += String.fromCharCode.apply(null, bytes.subarray(k, k + 0x8000));
    parts.push(str);
  }
  return parts.join('\n');
}
function setPageContent(doc, page, str) {
  const bytes = new Uint8Array(str.length);
  for (let k = 0; k < str.length; k++) bytes[k] = str.charCodeAt(k) & 255;
  const stream = doc.context.flateStream(bytes);
  page.node.set(PDFName.of('Contents'), doc.context.register(stream));
}

const hex = (codes, bytes) => `<${codes.map((c) => c.toString(16).padStart(bytes * 2, '0')).join('')}>`;
const fmt = (x) => (Math.abs(x) < 1e-6 ? '0' : String(+x.toFixed(3)));

/**
 * Removes the glyphs whose centres fall inside any of the boxes (PDF user space,
 * { x0, y0, x1, y1 }). Returns, per box, what was there: { removed, size, color, font, x, y }.
 */
export function removeTextInBoxes(doc, page, boxes) {
  const found = boxes.map(() => ({ removed: 0, size: 0, color: null, font: null, x: null, y: null }));
  if (!boxes.length) return found;
  const src = pageContent(doc, page);
  const toks = tokenize(src);
  const res = page.node.Resources?.();
  const fontsDict = res ? (() => { const f = res.get(PDFName.of('Font')); return f instanceof PDFRef ? doc.context.lookup(f) : f; })() : null;
  const metricsCache = new Map();
  const metricsFor = (name) => {
    if (!metricsCache.has(name)) {
      let fd = fontsDict instanceof PDFDict ? fontsDict.get(PDFName.of(name)) : null;
      if (fd instanceof PDFRef) fd = doc.context.lookup(fd);
      metricsCache.set(name, fd ? fontMetrics(doc, fd) : { bytes: 1, width: () => 556, base: '' });
    }
    return metricsCache.get(name);
  };

  let gs = { ctm: I, fill: [0, 0, 0], font: null, size: 12, Tc: 0, Tw: 0, Tz: 100, TL: 0, Ts: 0 };
  const stack = [];
  let Tm = I; let Tlm = I;
  const edits = []; // { s, e, text } replacements in src
  const inBox = (x, y) => boxes.findIndex((b) => x >= b.x0 && x <= b.x1 && y >= b.y0 && y <= b.y1);

  // Shows one string; returns the TJ elements to write back when something was removed
  const show = (str) => {
    const m = gs.font ? metricsFor(gs.font) : { bytes: 1, width: () => 556, base: '' };
    const codes = [];
    for (let k = 0; k < str.length; k += m.bytes) codes.push(m.bytes === 2 ? (str.charCodeAt(k) << 8) | (str.charCodeAt(k + 1) || 0) : str.charCodeAt(k));
    const th = gs.Tz / 100;
    const parts = []; // runs of kept codes and gaps
    let any = false;
    for (const c of codes) {
      const w0 = m.width(c) / 1000;
      const tx = (w0 * gs.size + gs.Tc + (m.bytes === 1 && c === 32 ? gs.Tw : 0)) * th;
      const trm = mul([gs.size * th, 0, 0, gs.size, 0, gs.Ts], mul(Tm, gs.ctm));
      const [cx, cy] = apply(trm, (tx / (gs.size * th || 1)) / 2, 0.3);
      const hit = inBox(cx, cy);
      if (hit >= 0) {
        any = true;
        const f = found[hit];
        f.removed += 1;
        if (!f.size) {
          const [ox, oy] = apply(mul(Tm, gs.ctm), 0, gs.Ts);
          f.size = Math.abs(gs.size * Math.hypot(trm[2] / (gs.size || 1), trm[3] / (gs.size || 1)));
          f.color = gs.fill; f.font = m.base; f.x = ox; f.y = oy;
        }
        parts.push({ gap: (-tx * 1000) / ((gs.size || 1) * th) });
      } else {
        const last = parts[parts.length - 1];
        if (last && last.codes) last.codes.push(c); else parts.push({ codes: [c] });
      }
      Tm = mul([1, 0, 0, 1, tx, 0], Tm);
    }
    return any ? parts.map((p) => (p.codes ? hex(p.codes, m.bytes) : fmt(p.gap))).join(' ') : null;
  };

  let operands = [];
  for (const tok of toks) {
    if (tok.t !== 'op') { operands.push(tok); continue; }
    const op = tok.v;
    const nums = operands.map((o) => o.v);
    switch (op) {
      case 'q': stack.push({ ...gs }); break;
      case 'Q': gs = stack.pop() || gs; break;
      case 'cm': if (nums.length === 6) gs.ctm = mul(nums, gs.ctm); break;
      case 'BT': Tm = I; Tlm = I; break;
      case 'Tf': gs.font = operands[0]?.v; gs.size = nums[1]; break;
      case 'Tc': gs.Tc = nums[0]; break;
      case 'Tw': gs.Tw = nums[0]; break;
      case 'Tz': gs.Tz = nums[0]; break;
      case 'TL': gs.TL = nums[0]; break;
      case 'Ts': gs.Ts = nums[0]; break;
      case 'Td': Tlm = mul([1, 0, 0, 1, nums[0], nums[1]], Tlm); Tm = Tlm; break;
      case 'TD': gs.TL = -nums[1]; Tlm = mul([1, 0, 0, 1, nums[0], nums[1]], Tlm); Tm = Tlm; break;
      case 'Tm': Tlm = nums.slice(0, 6); Tm = Tlm; break;
      case 'T*': Tlm = mul([1, 0, 0, 1, 0, -gs.TL], Tlm); Tm = Tlm; break;
      case 'g': gs.fill = [nums[0], nums[0], nums[0]]; break;
      case 'rg': gs.fill = nums.slice(0, 3); break;
      case 'k': { const [c, m2, y, k] = nums; gs.fill = [(1 - c) * (1 - k), (1 - m2) * (1 - k), (1 - y) * (1 - k)]; break; }
      case 'Tj': case "'": case '"': {
        const strTok = operands[operands.length - 1];
        if (op !== 'Tj') {
          if (op === '"') { gs.Tw = nums[0]; gs.Tc = nums[1]; }
          Tlm = mul([1, 0, 0, 1, 0, -gs.TL], Tlm); Tm = Tlm;
        }
        if (strTok?.t !== 'str') break;
        const arr = show(strTok.v);
        if (arr !== null) {
          const pre = op === "'" ? 'T* ' : op === '"' ? `${fmt(nums[0])} Tw ${fmt(nums[1])} Tc T* ` : '';
          edits.push({ s: operands[0].s, e: tok.e, text: `${pre}[${arr}] TJ` });
        }
        break;
      }
      case 'TJ': {
        const arrTok = operands[operands.length - 1];
        if (arrTok?.t !== 'arr') break;
        const out = [];
        let changedAny = false;
        for (const el of arrTok.v) {
          if (el.t === 'num') {
            out.push(fmt(el.v));
            Tm = mul([1, 0, 0, 1, ((-el.v / 1000) * gs.size * gs.Tz) / 100, 0], Tm);
          } else if (el.t === 'str') {
            const before = Tm;
            const r = show(el.v);
            if (r === null) {
              const m = gs.font ? metricsFor(gs.font) : { bytes: 1 };
              const codes = [];
              for (let k = 0; k < el.v.length; k += m.bytes) codes.push(m.bytes === 2 ? (el.v.charCodeAt(k) << 8) | (el.v.charCodeAt(k + 1) || 0) : el.v.charCodeAt(k));
              out.push(hex(codes, m.bytes));
            } else { out.push(r); changedAny = true; }
            void before;
          }
        }
        if (changedAny) edits.push({ s: arrTok.s, e: tok.e, text: `[${out.join(' ')}] TJ` });
        break;
      }
      default: break;
    }
    operands = [];
  }

  if (edits.length) {
    edits.sort((a, b) => a.s - b.s);
    let outStr = '';
    let pos = 0;
    for (const ed of edits) { outStr += src.slice(pos, ed.s) + ed.text; pos = ed.e; }
    outStr += src.slice(pos);
    setPageContent(doc, page, outStr);
  }
  return found;
}
