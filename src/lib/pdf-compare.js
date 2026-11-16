// Compare two PDFs: read the words and where they sit, line both documents up,
// and build a report PDF with the two versions side by side. Removed words are
// boxed in red on the original, added words in green on the revised copy, and a
// summary at the front lists every change.
//
// The pages come from extractPages() in pdf-convert.js: items in viewport space
// (scale 1, top-left origin, y on the baseline).
import { PDFDocument, PDFName, StandardFonts, rgb, degrees } from '@cantoo/pdf-lib';
import { createTextKit } from './pdf-text.js';

// ------------------------------------------------------------------ words

// Helvetica's character widths (per 1000), a good guide to where a word starts
// inside a run of proportional text
const HELV = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
  556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
  1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
  333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
  556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
const charWidth = (ch) => {
  const cp = ch.codePointAt(0);
  if (cp >= 32 && cp <= 126) return HELV[cp - 32];
  return cp >= 0x2e80 ? 1000 : 556;
};

// A word, keeping numbers like 4,500.00 and 12/03 whole; trailing punctuation is its own token
const TOKEN = /[\p{L}\p{N}\p{M}](?:[\p{L}\p{N}\p{M}'\u2019\-%]|[.,/:](?=[\p{L}\p{N}]))*|\S/gu;

/** Splits text items into words with estimated boxes, grouped into lines in reading order */
export function pageWords(page, pageNo) {
  const words = [];
  let line = -1;
  let last = null;
  for (const it of page.items) {
    const str = it.str || '';
    if (!str.trim()) continue;
    const size = it.size || 10;
    // A new line starts when the baseline moves or the text jumps back to the left
    const newLine = !last || Math.abs(it.y - last.y) > size * 0.5 || it.x < last.x1 - size * 2;
    if (newLine) line += 1;
    const chars = [...str];
    const offs = [0];
    for (const ch of chars) offs.push(offs[offs.length - 1] + charWidth(ch));
    const k = offs[offs.length - 1] ? it.w / offs[offs.length - 1] : 0;
    // Code-unit index -> character index, for surrogate pairs
    const ci = [];
    let n = 0;
    for (const ch of chars) { ci.push(n); for (let u = 1; u < ch.length; u++) ci.push(n); n++; }
    ci.push(n);
    let prevEnd = !newLine && last && it.x - last.x1 < size * 0.15 ? 0 : -1;
    for (const m of str.matchAll(TOKEN)) {
      const a = ci[m.index];
      const b = ci[m.index + m[0].length];
      words.push({
        text: m[0], page: pageNo, line,
        // Glued to the word before it (no space), such as a full stop
        glue: m.index === prevEnd,
        x: it.x + offs[a] * k, w: Math.max((offs[b] - offs[a]) * k, 1),
        top: it.y - size * 0.82, h: size * 1.04,
      });
      prevEnd = m.index + m[0].length;
    }
    last = { y: it.y, x1: it.x + it.w };
  }
  return words;
}

export function documentWords(pages) {
  const out = [];
  let base = 0;
  pages.forEach((p, i) => {
    const ws = pageWords(p, i);
    let maxLine = -1;
    for (const w of ws) { maxLine = Math.max(maxLine, w.line); out.push({ ...w, line: w.line + base }); }
    base += maxLine + 1;
  });
  return out;
}

// ------------------------------------------------------------------- diff

/**
 * Diffs two sequences of keys. Returns ops in order: ['=', i, j], ['-', i] or ['+', j].
 * Common ends are trimmed, short middles use an exact LCS table, and long ones are
 * split on lines that appear exactly once in each (patience anchors).
 */
export function diffSequences(a, b) {
  const ops = [];
  const LIMIT = 4e6;

  const lcs = (a0, a1, b0, b1) => {
    const n = a1 - a0;
    const m = b1 - b0;
    const w = m + 1;
    const L = new Uint32Array((n + 1) * w);
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        L[i * w + j] = a[a0 + i] === b[b0 + j] ? L[(i + 1) * w + j + 1] + 1 : Math.max(L[(i + 1) * w + j], L[i * w + j + 1]);
      }
    }
    let i = 0;
    let j = 0;
    while (i < n && j < m) {
      if (a[a0 + i] === b[b0 + j]) { ops.push(['=', a0 + i, b0 + j]); i++; j++; }
      else if (L[(i + 1) * w + j] >= L[i * w + j + 1]) { ops.push(['-', a0 + i]); i++; }
      else { ops.push(['+', b0 + j]); j++; }
    }
    for (; i < n; i++) ops.push(['-', a0 + i]);
    for (; j < m; j++) ops.push(['+', b0 + j]);
  };

  const anchors = (a0, a1, b0, b1) => {
    const ca = new Map();
    const cb = new Map();
    for (let i = a0; i < a1; i++) { const c = ca.get(a[i]); ca.set(a[i], c ? { n: c.n + 1, i } : { n: 1, i }); }
    for (let j = b0; j < b1; j++) { const c = cb.get(b[j]); cb.set(b[j], c ? { n: c.n + 1, i: j } : { n: 1, i: j }); }
    const pairs = [];
    for (const [k, c] of ca) {
      const d = cb.get(k);
      if (c.n === 1 && d && d.n === 1) pairs.push([c.i, d.i]);
    }
    pairs.sort((p, q) => p[0] - q[0]);
    // Longest increasing run of j (patience sorting)
    const tails = [];
    const prev = new Array(pairs.length);
    for (let k = 0; k < pairs.length; k++) {
      let lo = 0;
      let hi = tails.length;
      while (lo < hi) { const mid = (lo + hi) >> 1; if (pairs[tails[mid]][1] < pairs[k][1]) lo = mid + 1; else hi = mid; }
      prev[k] = lo ? tails[lo - 1] : -1;
      tails[lo] = k;
    }
    const out = [];
    for (let k = tails.length ? tails[tails.length - 1] : -1; k >= 0; k = prev[k]) out.push(pairs[k]);
    return out.reverse();
  };

  const rec = (a0, a1, b0, b1) => {
    while (a0 < a1 && b0 < b1 && a[a0] === b[b0]) { ops.push(['=', a0, b0]); a0++; b0++; }
    const tail = [];
    while (a1 > a0 && b1 > b0 && a[a1 - 1] === b[b1 - 1]) { a1--; b1--; tail.push(['=', a1, b1]); }
    if (a0 === a1 || b0 === b1) {
      for (let i = a0; i < a1; i++) ops.push(['-', i]);
      for (let j = b0; j < b1; j++) ops.push(['+', j]);
    } else if ((a1 - a0) * (b1 - b0) <= LIMIT) {
      lcs(a0, a1, b0, b1);
    } else {
      const an = anchors(a0, a1, b0, b1);
      if (!an.length) {
        for (let i = a0; i < a1; i++) ops.push(['-', i]);
        for (let j = b0; j < b1; j++) ops.push(['+', j]);
      } else {
        let pa = a0;
        let pb = b0;
        for (const [i, j] of an) {
          rec(pa, i, pb, j);
          ops.push(['=', i, j]);
          pa = i + 1; pb = j + 1;
        }
        rec(pa, a1, pb, b1);
      }
    }
    for (let k = tail.length - 1; k >= 0; k--) ops.push(tail[k]);
  };

  rec(0, a.length, 0, b.length);
  return ops;
}

const norm = (s) => s.normalize('NFC').replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/­/g, '');

/**
 * Compares the text of two documents.
 * Returns { changes: [{ kind: 'added'|'removed'|'replaced', removed: [word], added: [word] }],
 *           pairs: [[pageA|null, pageB|null]], wordsA, wordsB }
 */
export function compareText(pagesA, pagesB) {
  const wordsA = documentWords(pagesA);
  const wordsB = documentWords(pagesB);
  const keysA = wordsA.map((w) => norm(w.text));
  const keysB = wordsB.map((w) => norm(w.text));

  // Lines first, so a moved paragraph or a long document stays quick, then words in what differs
  const linesOf = (words, keys) => {
    const lines = [];
    words.forEach((w, i) => {
      let l = lines[lines.length - 1];
      if (!l || l.id !== w.line) { l = { id: w.line, from: i, to: i, page: w.page }; lines.push(l); }
      l.to = i + 1;
    });
    for (const l of lines) l.key = keys.slice(l.from, l.to).join(' ');
    return lines;
  };
  const linesA = linesOf(wordsA, keysA);
  const linesB = linesOf(wordsB, keysB);
  const lineOps = diffSequences(linesA.map((l) => l.key), linesB.map((l) => l.key));

  const changes = [];
  const together = new Map(); // 'pageA:pageB' -> matched words, for pairing pages
  const vote = (pa, pb, n) => { const k = `${pa}:${pb}`; together.set(k, (together.get(k) || 0) + n); };

  let pendA = [];
  let pendB = [];
  const flush = () => {
    if (!pendA.length && !pendB.length) return;
    const ia = pendA.flatMap((l) => range(l.from, l.to));
    const ib = pendB.flatMap((l) => range(l.from, l.to));
    const ops = diffSequences(ia.map((i) => keysA[i]), ib.map((j) => keysB[j]));
    let cur = null;
    for (const op of ops) {
      if (op[0] === '=') {
        cur = null;
        vote(wordsA[ia[op[1]]].page, wordsB[ib[op[2]]].page, 1);
        continue;
      }
      if (!cur) { cur = { removed: [], added: [] }; changes.push(cur); }
      if (op[0] === '-') cur.removed.push(wordsA[ia[op[1]]]);
      else cur.added.push(wordsB[ib[op[1]]]);
    }
    pendA = [];
    pendB = [];
  };
  for (const op of lineOps) {
    if (op[0] === '=') {
      flush();
      const la = linesA[op[1]];
      vote(la.page, linesB[op[2]].page, la.to - la.from);
    } else if (op[0] === '-') pendA.push(linesA[op[1]]);
    else pendB.push(linesB[op[1]]);
  }
  flush();
  for (const c of changes) c.kind = c.removed.length && c.added.length ? 'replaced' : c.removed.length ? 'removed' : 'added';

  return { changes, pairs: pairPages(pagesA.length, pagesB.length, together), wordsA, wordsB };
}

function range(a, b) { const out = []; for (let i = a; i < b; i++) out.push(i); return out; }

/** Lines the pages up in order, pairing pages that share the most text */
export function pairPages(n, m, together) {
  const w = (i, j) => together.get(`${i}:${j}`) || 0;
  const S = Array.from({ length: n + 1 }, () => new Float64Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) {
      const pair = w(i, j) ? S[i + 1][j + 1] + w(i, j) : -Infinity;
      S[i][j] = Math.max(pair, S[i + 1][j], S[i][j + 1]);
    }
  }
  const anchors = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (w(i, j) && S[i][j] === S[i + 1][j + 1] + w(i, j)) { anchors.push([i, j]); i++; j++; }
    else if (S[i + 1][j] >= S[i][j + 1]) i++;
    else j++;
  }
  // Pages between two anchors with nothing in common (a rewritten page) pair up in order
  const pairs = [];
  let pa = 0;
  let pb = 0;
  for (const [ai, bj] of [...anchors, [n, m]]) {
    const ga = ai - pa;
    const gb = bj - pb;
    for (let k = 0; k < Math.max(ga, gb); k++) pairs.push([k < ga ? pa + k : null, k < gb ? pb + k : null]);
    if (ai < n) pairs.push([ai, bj]);
    pa = ai + 1; pb = bj + 1;
  }
  return pairs;
}

// ----------------------------------------------------------------- report

const RED = rgb(0.86, 0.15, 0.15);
const GREEN = rgb(0.09, 0.6, 0.27);
const INK = rgb(0.07, 0.09, 0.12);
const MUTED = rgb(0.4, 0.43, 0.48);
const LINE = rgb(0.86, 0.88, 0.91);

export const joinWords = (words) => words.map((w, i) => (i && !w.glue ? ' ' : '') + w.text).join('');
const quote = (words, max = 90) => {
  const s = joinWords(words);
  return s.length > max ? `${s.slice(0, max - 1).trimEnd()}...` : s;
};

/** Boxes for a run of words: one per line, from the first word's left edge to the last word's right */
export function lineBoxes(words) {
  const by = new Map();
  for (const w of words) {
    const k = `${w.page}:${w.line}`;
    const b = by.get(k);
    if (!b) by.set(k, { page: w.page, x0: w.x, x1: w.x + w.w, top: w.top, bottom: w.top + w.h });
    else { b.x0 = Math.min(b.x0, w.x); b.x1 = Math.max(b.x1, w.x + w.w); b.top = Math.min(b.top, w.top); b.bottom = Math.max(b.bottom, w.top + w.h); }
  }
  return [...by.values()];
}

/**
 * Builds the report PDF.
 * a, b: { bytes, pages (from extractPages), name }
 * result: from compareText()
 */
export async function buildReport(a, b, result, { title = 'Comparison', date = new Date() } = {}) {
  const out = await PDFDocument.create();
  out.setTitle(title);
  out.setCreator('Kertas Office');
  const kit = createTextKit(out, { regular: await out.embedFont(StandardFonts.Helvetica), bold: await out.embedFont(StandardFonts.HelveticaBold) });
  const summaryTexts = result.changes.flatMap((c) => [quote(c.removed), quote(c.added)]);
  try { await kit.prepare([a.name, b.name, ...summaryTexts]); } catch { /* no extra fonts: unsupported characters show as ? */ }

  const srcA = await PDFDocument.load(a.bytes, { ignoreEncryption: true });
  const srcB = await PDFDocument.load(b.bytes, { ignoreEncryption: true });

  const counts = { added: 0, removed: 0, replaced: 0 };
  for (const c of result.changes) counts[c.kind] += 1;

  // Which report page shows each source page, for the summary's page numbers and links
  const rowOf = { a: new Map(), b: new Map() };
  result.pairs.forEach(([pa, pb], k) => { if (pa !== null) rowOf.a.set(pa, k); if (pb !== null) rowOf.b.set(pb, k); });

  // ---- summary pages
  const SW = 595;
  const SH = 842;
  const M = 48;
  const summaryPages = [];
  let sp = null;
  let y = 0;
  const newSummary = () => { sp = out.addPage([SW, SH]); summaryPages.push(sp); y = SH - M; };
  newSummary();
  y -= 20;
  for (const ln of wrapBold(kit, title, SW - M * 2, 20).slice(0, 3)) {
    kit.draw(sp, ln, { x: M, y, size: 20, color: INK, bold: true });
    y -= 26;
  }
  y -= 8;
  const meta = [
    ['Original', a.name],
    ['Revised', b.name],
    ['Compared', date.toLocaleString('en-GB', { dateStyle: 'medium', timeStyle: 'short' })],
  ];
  for (const [k, v] of meta) {
    kit.draw(sp, k, { x: M, y, size: 10, color: MUTED });
    kit.draw(sp, fit(kit, v, SW - M * 2 - 80, 10), { x: M + 80, y, size: 10, color: INK });
    y -= 16;
  }
  y -= 10;
  const total = result.changes.length;
  const head = total
    ? `${total} change${total === 1 ? '' : 's'}: ${counts.replaced} replaced, ${counts.added} added, ${counts.removed} removed`
    : 'No differences in the text';
  kit.draw(sp, head, { x: M, y, size: 13, color: INK, bold: true });
  y -= 14;
  sp.drawRectangle({ x: M, y: y - 4, width: 10, height: 10, color: RED, opacity: 0.25, borderColor: RED, borderWidth: 0.8 });
  kit.draw(sp, 'Removed, on the original (left)', { x: M + 16, y: y - 2, size: 9, color: MUTED });
  sp.drawRectangle({ x: M + 190, y: y - 4, width: 10, height: 10, color: GREEN, opacity: 0.25, borderColor: GREEN, borderWidth: 0.8 });
  kit.draw(sp, 'Added, on the revised copy (right)', { x: M + 206, y: y - 2, size: 9, color: MUTED });
  y -= 28;

  const links = []; // { page, rect, target row }
  const colW = SW - M * 2 - 70;
  result.changes.forEach((c, n) => {
    const textLines = [];
    if (c.removed.length) textLines.push({ label: c.kind === 'replaced' ? 'Was' : 'Removed', text: quote(c.removed, 400), color: RED });
    if (c.added.length) textLines.push({ label: c.kind === 'replaced' ? 'Now' : 'Added', text: quote(c.added, 400), color: GREEN });
    const wrapped = textLines.map((t) => ({ ...t, lines: wrap(kit, t.text, colW - 50, 10).slice(0, 4) }));
    const height = 16 + wrapped.reduce((h, t) => h + t.lines.length * 13, 0) + 8;
    if (y - height < M) newSummary();
    const top = y;
    const onA = c.removed.length ? c.removed[0].page : null;
    const onB = c.added.length ? c.added[0].page : null;
    const where = onB === null ? `original page ${onA + 1}`
      : onA === null ? `revised page ${onB + 1}`
        : onA === onB ? `page ${onA + 1}` : `page ${onA + 1}, now page ${onB + 1}`;
    kit.draw(sp, `${n + 1}. ${c.kind === 'replaced' ? 'Replaced' : c.kind === 'added' ? 'Added' : 'Removed'}`, { x: M, y: y - 10, size: 10, color: INK, bold: true });
    kit.draw(sp, where, { x: M + colW + 70 - kit.width(where, 9), y: y - 10, size: 9, color: MUTED });
    y -= 24;
    for (const t of wrapped) {
      t.lines.forEach((ln, i) => {
        if (!i) kit.draw(sp, t.label, { x: M + 12, y, size: 9, color: t.color, bold: true });
        kit.draw(sp, ln, { x: M + 62, y, size: 10, color: INK });
        y -= 13;
      });
    }
    y -= 6;
    sp.drawLine({ start: { x: M, y: y + 2 }, end: { x: SW - M, y: y + 2 }, thickness: 0.5, color: LINE });
    y -= 6;
    const row = onA !== null ? rowOf.a.get(onA) : rowOf.b.get(onB);
    if (row !== undefined) links.push({ page: sp, rect: [M, y + 4, SW - M, top], row });
  });

  // ---- side-by-side pages
  const GAP = 24;
  const HEAD = 34;
  const PAD = 20;
  const boxesA = groupBoxes(result.changes.flatMap((c) => lineBoxes(c.removed)));
  const boxesB = groupBoxes(result.changes.flatMap((c) => lineBoxes(c.added)));
  const embA = await embedAll(out, srcA, result.pairs.map((p) => p[0]));
  const embB = await embedAll(out, srcB, result.pairs.map((p) => p[1]));
  const rowPages = [];
  for (const [k, [pa, pb]] of result.pairs.entries()) {
    const da = pa !== null ? a.pages[pa] : null;
    const db = pb !== null ? b.pages[pb] : null;
    const wa = da ? da.width : (db ? db.width : 595);
    const wb = db ? db.width : wa;
    const h = Math.max(da ? da.height : 0, db ? db.height : 0) || 842;
    const W = PAD * 2 + wa + GAP + wb;
    const H = PAD + HEAD + h + PAD;
    const page = out.addPage([W, H]);
    rowPages.push(page);
    page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: rgb(0.95, 0.96, 0.97) });
    const sides = [
      { d: da, n: pa, emb: embA.get(pa), boxes: boxesA.get(pa), x: PAD, w: wa, label: 'Original', name: a.name, color: RED },
      { d: db, n: pb, emb: embB.get(pb), boxes: boxesB.get(pb), x: PAD + wa + GAP, w: wb, label: 'Revised', name: b.name, color: GREEN },
    ];
    for (const s of sides) {
      const top = H - PAD - HEAD;
      const label = s.d ? `${s.label}, page ${s.n + 1}` : `${s.label}: no matching page`;
      kit.draw(page, label, { x: s.x, y: H - PAD - 14, size: 11, color: INK, bold: true });
      kit.draw(page, fit(kit, s.name, s.w - kit.width(label, 11, true) - 12, 9), { x: s.x + kit.width(label, 11, true) + 10, y: H - PAD - 14, size: 9, color: MUTED });
      if (!s.d) {
        page.drawRectangle({ x: s.x, y: top - h, width: s.w, height: h, borderColor: LINE, borderWidth: 1, borderDashArray: [4, 4] });
        continue;
      }
      const ph = s.d.height;
      const y0 = top - ph; // bottom of the drawn page
      page.drawRectangle({ x: s.x, y: y0, width: s.w, height: ph, color: rgb(1, 1, 1), borderColor: LINE, borderWidth: 0.6 });
      if (s.emb) drawEmbedded(page, s.emb.emb, s.emb.rotate, s.x, y0);
      for (const bx of s.boxes || []) {
        const pad = 1.2;
        page.drawRectangle({
          x: s.x + bx.x0 - pad, y: y0 + (ph - bx.bottom) - pad,
          width: bx.x1 - bx.x0 + pad * 2, height: bx.bottom - bx.top + pad * 2,
          color: s.color, opacity: 0.2, borderColor: s.color, borderWidth: 0.8, borderOpacity: 0.9,
        });
        // A mark in the gutter, so changes are easy to spot when zoomed out
        const gx = s.label === 'Original' ? s.x - 8 : s.x + s.w + 4;
        page.drawRectangle({ x: gx, y: y0 + (ph - bx.bottom), width: 4, height: bx.bottom - bx.top, color: s.color });
      }
    }
  }

  // Summary entries link to the page that shows the change
  for (const l of links) addLink(out, l.page, l.rect, rowPages[l.row]);

  return out.save();
}

function groupBoxes(boxes) {
  const m = new Map();
  for (const b of boxes) { if (!m.has(b.page)) m.set(b.page, []); m.get(b.page).push(b); }
  return m;
}

async function embedAll(out, src, indices) {
  const map = new Map();
  const list = [...new Set(indices.filter((i) => i !== null))];
  const srcPages = src.getPages();
  const valid = list.filter((i) => srcPages[i]);
  if (!valid.length) return map;
  const boxes = valid.map((i) => {
    const c = srcPages[i].getCropBox();
    return { left: c.x, bottom: c.y, right: c.x + c.width, top: c.y + c.height };
  });
  const embedded = await out.embedPages(valid.map((i) => srcPages[i]), boxes);
  valid.forEach((i, k) => map.set(i, { emb: embedded[k], rotate: ((srcPages[i].getRotation().angle % 360) + 360) % 360 }));
  return map;
}

/** Draws an embedded page with its /Rotate applied, lower-left of the displayed page at (x, y) */
function drawEmbedded(page, emb, rotate, x, y) {
  const { width: W, height: H } = emb;
  if (rotate === 90) page.drawPage(emb, { x, y: y + W, rotate: degrees(-90) });
  else if (rotate === 180) page.drawPage(emb, { x: x + W, y: y + H, rotate: degrees(180) });
  else if (rotate === 270) page.drawPage(emb, { x: x + H, y, rotate: degrees(90) });
  else page.drawPage(emb, { x, y });
}

function addLink(doc, page, [x0, y0, x1, y1], target) {
  if (!target) return;
  const ctx = doc.context;
  const ref = ctx.register(ctx.obj({
    Type: 'Annot', Subtype: 'Link', Rect: [x0, y0, x1, y1], Border: [0, 0, 0],
    Dest: [target.ref, 'Fit'],
  }));
  const annots = page.node.Annots();
  if (annots) annots.push(ref);
  else page.node.set(PDFName.of('Annots'), ctx.obj([ref]));
}

function wrap(kit, text, width, size) {
  const words = String(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const next = cur ? `${cur} ${w}` : w;
    if (cur && kit.width(next, size) > width) { lines.push(cur); cur = w; } else cur = next;
  }
  if (cur) lines.push(cur);
  return lines.length ? lines : [''];
}

function wrapBold(kit, text, width, size) {
  const words = String(text).split(/\s+/).filter(Boolean);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const next = cur ? `${cur} ${w}` : w;
    if (cur && kit.width(next, size, true) > width) { lines.push(cur); cur = w; } else cur = next;
  }
  if (cur) lines.push(cur);
  return lines;
}

function fit(kit, text, width, size) {
  let s = String(text ?? '');
  if (kit.width(s, size) <= width) return s;
  while (s.length > 1 && kit.width(`${s}...`, size) > width) s = s.slice(0, -1);
  return `${s}...`;
}
