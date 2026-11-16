// Compare two documents: the differences between an original and a revised version, shown in a copy
// of the original as tracked changes that can be reviewed, accepted or rejected one by one.
//
// How it works: the revised document's paragraphs are read straight from its document.xml. The
// paragraphs of the copy are lined up with them (unchanged paragraphs first, then similar ones are
// paired), and each pair is compared word by word. The copy is then edited into the revised text
// with track changes on, so the editor records every edit as an ordinary insertion or deletion.
// This compares text: formatting-only changes and changed table structure are not reported.
import JSZip from 'jszip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** The paragraphs of a .docx body in reading order, as plain text (tables included) */
export async function readDocxParagraphs(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file('word/document.xml');
  if (!file) throw new Error('This file has no document body');
  const xml = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
  const body = xml.getElementsByTagNameNS(W, 'body')[0];
  if (!body) throw new Error('This file has no document body');
  const out = [];
  const walk = (node) => {
    for (const child of node.children) {
      const n = child.localName;
      if (child.namespaceURI !== W && n !== 'AlternateContent' && n !== 'Choice' && n !== 'Fallback' && n !== 'sdt') {
        // drawings and other foreign content: text boxes are not compared
        continue;
      }
      if (n === 'p') out.push({ text: paragraphText(child), styleId: paragraphStyle(child) });
      else if (n === 'del' || n === 'moveFrom' || n === 'Fallback' || n === 'txbxContent') continue;
      else walk(child);
    }
  };
  walk(body);
  return out;
}

/** The paragraph style of one w:p, if it has one */
function paragraphStyle(p) {
  const pPr = [...p.children].find((c) => c.namespaceURI === W && c.localName === 'pPr');
  const st = pPr && [...pPr.children].find((c) => c.namespaceURI === W && c.localName === 'pStyle');
  return st?.getAttributeNS(W, 'val') || st?.getAttribute('w:val') || null;
}

/** The text of one w:p: runs, insertions and field results, without deleted text or field codes */
function paragraphText(p) {
  let text = '';
  const walk = (node) => {
    for (const child of node.children) {
      if (child.namespaceURI !== W) continue;
      const n = child.localName;
      if (n === 't') text += child.textContent;
      else if (n === 'noBreakHyphen') text += '‑';
      else if (n === 'del' || n === 'moveFrom' || n === 'instrText' || n === 'delText' || n === 'pPr' || n === 'rPr' || n === 'txbxContent' || n === 'drawing' || n === 'pict') continue;
      else walk(child);
    }
  };
  walk(p);
  return text;
}

/** The text blocks of an editor document in reading order, with the document position of every character */
export function editorParagraphs(doc) {
  const out = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const chars = [];
    let text = '';
    node.descendants((child, offset) => {
      if (!child.isText) return true;
      // text already marked as deleted is not part of the document's current text
      if (child.marks.some((m) => m.type.name === 'trackDelete')) return false;
      for (let i = 0; i < child.text.length; i++) chars.push(pos + 1 + offset + i);
      text += child.text;
      return false;
    });
    out.push({ pos, size: node.nodeSize, text, chars });
    return false;
  });
  return out;
}

const norm = (s) => s.replace(/\s+/g, ' ').trim();
// words, numbers with their separators (412,300 or 3.5 or 15/11) as one piece, spaces, punctuation
const tokenize = (s) => s.match(/\s+|\p{N}+(?:[.,:/\-]\p{N}+)*|[\p{L}\p{N}_'’]+|[^\s\p{L}\p{N}_]/gu) || [];

/** Longest common subsequence of two arrays, as index pairs */
function lcsPairs(a, b, eq = (x, y) => x === y) {
  const n = a.length; const m = b.length;
  // trim matching ends first: most edits are local
  let start = 0;
  while (start < n && start < m && eq(a[start], b[start])) start++;
  let endA = n; let endB = m;
  while (endA > start && endB > start && eq(a[endA - 1], b[endB - 1])) { endA--; endB--; }
  const pairs = [];
  for (let i = 0; i < start; i++) pairs.push([i, i]);
  const rows = endA - start; const cols = endB - start;
  if (rows && cols) {
    if (rows * cols > 16e6) {
      // too big for a full table: match greedily in order
      let j = start;
      for (let i = start; i < endA; i++) {
        for (let k = j; k < Math.min(endB, j + 200); k++) if (eq(a[i], b[k])) { pairs.push([i, k]); j = k + 1; break; }
      }
    } else {
      const dp = new Uint32Array((rows + 1) * (cols + 1));
      const at = (i, j) => i * (cols + 1) + j;
      for (let i = rows - 1; i >= 0; i--) {
        for (let j = cols - 1; j >= 0; j--) {
          dp[at(i, j)] = eq(a[start + i], b[start + j]) ? dp[at(i + 1, j + 1)] + 1 : Math.max(dp[at(i + 1, j)], dp[at(i, j + 1)]);
        }
      }
      let i = 0; let j = 0;
      while (i < rows && j < cols) {
        if (eq(a[start + i], b[start + j])) { pairs.push([start + i, start + j]); i++; j++; }
        else if (dp[at(i + 1, j)] >= dp[at(i, j + 1)]) i++;
        else j++;
      }
    }
  }
  for (let i = 0; i < n - endA; i++) pairs.push([endA + i, endB + i]);
  return pairs;
}

/** How alike two paragraphs are, 0 to 1, by shared words */
function similarity(a, b) {
  const wa = norm(a).toLowerCase().split(' ').filter(Boolean);
  const wb = norm(b).toLowerCase().split(' ').filter(Boolean);
  if (!wa.length || !wb.length) return 0;
  const common = lcsPairs(wa, wb).length;
  return (2 * common) / (wa.length + wb.length);
}

/**
 * The edits that turn the original paragraphs into the revised ones:
 * { kind: 'words', para, ops } for a changed paragraph, { kind: 'delete', para } for a removed one,
 * and { kind: 'insert', after, before, text } for a new paragraph.
 */
export function planChanges(orig, rev) {
  const changes = [];
  const anchors = lcsPairs(orig, rev, (x, y) => norm(x.text) === norm(y.text) && (norm(x.text) !== '' || x.text === y.text));
  anchors.push([orig.length, rev.length]);
  let i = 0; let j = 0;
  for (const [ai, bj] of anchors) {
    // the unmatched stretch between two unchanged paragraphs
    const os = []; for (let k = i; k < ai; k++) os.push(k);
    const rs = []; for (let k = j; k < bj; k++) rs.push(k);
    let r = 0;
    let lastOrig = i - 1;
    for (const o of os) {
      // pair this original paragraph with the next similar revised one; revised ones skipped are new
      let found = -1;
      for (let k = r; k < Math.min(rs.length, r + 6); k++) if (similarity(orig[o].text, rev[rs[k]].text) >= 0.4) { found = k; break; }
      // nothing alike, but as many paragraphs left on each side: one stands in for the other in
      // the same place (a figure in a table cell, say), so show it as a replacement
      if (found < 0 && r < rs.length && os.length - os.indexOf(o) === rs.length - r) found = r;
      if (found < 0) { changes.push({ kind: 'delete', para: o }); continue; }
      for (let k = r; k < found; k++) changes.push({ kind: 'insert', after: lastOrig, before: o, text: rev[rs[k]].text, styleId: rev[rs[k]].styleId });
      changes.push({ kind: 'words', para: o, ops: wordOps(orig[o].text, rev[rs[found]].text) });
      r = found + 1;
      lastOrig = o;
    }
    for (let k = r; k < rs.length; k++) changes.push({ kind: 'insert', after: lastOrig >= i ? lastOrig : ai - 1, before: ai, text: rev[rs[k]].text, styleId: rev[rs[k]].styleId });
    i = ai + 1; j = bj + 1;
  }
  return changes.filter((c) => c.kind !== 'words' || c.ops.length);
}

/** Word-level edits inside one paragraph: { at, del, ins } in character offsets of the original */
export function wordOps(a, b) {
  const ta = tokenize(a); const tb = tokenize(b);
  const pairs = lcsPairs(ta, tb);
  pairs.push([ta.length, tb.length]);
  const offA = [0]; for (const t of ta) offA.push(offA[offA.length - 1] + t.length);
  const ops = [];
  let i = 0; let j = 0;
  for (const [pi, pj] of pairs) {
    if (pi > i || pj > j) ops.push({ at: offA[i], del: offA[pi] - offA[i], ins: tb.slice(j, pj).join('') });
    i = pi + 1; j = pj + 1;
  }
  return ops;
}

/** Gives the new paragraph at pos the revised paragraph's style, instead of its neighbour's */
function restyle(ed, pos, styleId) {
  const node = ed.state.doc.nodeAt(pos);
  if (!node?.isTextblock) return;
  const paragraphProperties = styleId ? { styleId } : {};
  ed.view.dispatch(ed.state.tr.setNodeMarkup(pos, null, { ...node.attrs, paragraphProperties }).setMeta('skipTrackChanges', true));
}

/**
 * Applies a comparison to the editor of the copy, which must be in suggesting mode.
 * Returns how many changes were made.
 */
export function applyChanges(getEditor, plan, orig) {
  // work from the end of the document backwards, so earlier positions stay valid
  const steps = [];
  for (const c of plan) {
    if (c.kind === 'words') for (const op of c.ops) steps.push({ para: c.para, ...op });
    else if (c.kind === 'delete') steps.push({ para: c.para, at: 0, del: orig[c.para].text.length, ins: '' });
  }
  // new paragraphs: after the paragraph before them, or before the first paragraph
  const inserts = plan.filter((c) => c.kind === 'insert');
  const sortKey = (s) => s.pos;
  const all = [];
  for (const s of steps) {
    const p = orig[s.para];
    const from = s.at < p.chars.length ? p.chars[s.at] : (p.chars.length ? p.chars[p.chars.length - 1] + 1 : p.pos + 1);
    const lastIdx = s.at + s.del - 1;
    const to = s.del ? p.chars[lastIdx] + 1 : from;
    all.push({ pos: from, order: 1, run: (ed) => {
      // insert after the deleted words, then mark the deleted words; both become tracked changes
      if (s.ins) ed.view.dispatch(ed.state.tr.insertText(s.ins, to));
      if (s.del) ed.view.dispatch(ed.state.tr.delete(from, to));
    } });
  }
  // each new paragraph lands next to its neighbour, ahead of any added there before it, so the
  // last one goes in first and they end up in order
  inserts.forEach((c, n) => {
    if (c.after >= 0) {
      const p = orig[c.after];
      const end = p.pos + p.size - 1;
      all.push({ pos: end, order: 2, seq: n, run: (ed) => {
        ed.view.dispatch(ed.state.tr.split(end));
        restyle(getEditor(), end + 1, c.styleId);
        const e2 = getEditor();
        if (c.text) e2.view.dispatch(e2.state.tr.insertText(c.text, end + 2));
      } });
    } else {
      const p = orig[c.before] || orig[0];
      const start = p.pos + 1;
      all.push({ pos: start, order: 0, seq: n, run: (ed) => {
        ed.view.dispatch(ed.state.tr.split(start));
        restyle(getEditor(), start - 1, c.styleId);
        const e2 = getEditor();
        if (c.text) e2.view.dispatch(e2.state.tr.insertText(c.text, start));
      } });
    }
  });
  all.sort((x, y) => sortKey(y) - sortKey(x) || y.order - x.order || (y.seq ?? 0) - (x.seq ?? 0));
  for (const step of all) step.run(getEditor());
  return steps.length + inserts.length;
}
