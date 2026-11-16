// Bookmarks and cross-references in Writer.
//
// Bookmarks are the editor's own bookmarkStart / bookmarkEnd nodes. A cross-reference is a REF or
// PAGEREF field (the editor's crossReference node) pointing at a bookmark; a heading or caption gets
// a hidden _Ref bookmark around its text the first time it is referred to, as Word does. The text a
// cross-reference shows (the target's text, its page number, or "above" / "below") is worked out
// here and kept up to date as the document changes.

/** Word's rules for bookmark names: a letter first, then letters, digits or underscores, up to 40 */
export const BOOKMARK_NAME = /^[A-Za-z][A-Za-z0-9_]{0,39}$/;

/** Every bookmark: { name, id, from, to, text, hidden }, in document order */
export function listBookmarks(doc) {
  const starts = [];
  const ends = new Map();
  doc.descendants((n, pos) => {
    if (n.type.name === 'bookmarkStart') starts.push({ name: n.attrs.name, id: String(n.attrs.id), pos });
    else if (n.type.name === 'bookmarkEnd') ends.set(String(n.attrs.id), pos);
    return true;
  });
  return starts.map((s) => {
    const from = s.pos + 1;
    const to = ends.get(s.id) ?? from;
    return { name: s.name, id: s.id, from, to, text: textBetween(doc, from, to), hidden: String(s.name || '').startsWith('_') };
  });
}

function textBetween(doc, from, to) {
  if (to <= from) return '';
  let text = '';
  doc.nodesBetween(from, to, (n, pos) => {
    if (n.isText) text += n.text.slice(Math.max(0, from - pos), to - pos);
    else if (n.type.name === 'crossReference') text += n.attrs.resolvedText || '';
    return true;
  });
  return text.replace(/\s+/g, ' ').trim();
}

/** A bookmark id not used yet */
function nextBookmarkId(doc) {
  let max = 0;
  doc.descendants((n) => { if (n.type.name === 'bookmarkStart' || n.type.name === 'bookmarkEnd') max = Math.max(max, Number(n.attrs.id) || 0); return true; });
  return String(max + 1);
}

/** Puts a bookmark called name around from..to (a point when they are equal) */
export function addBookmark(ed, name, from, to = from) {
  const { bookmarkStart, bookmarkEnd } = ed.schema.nodes;
  if (!bookmarkStart || !bookmarkEnd) throw new Error('Bookmarks are not available in this document');
  const id = nextBookmarkId(ed.state.doc);
  const tr = ed.state.tr;
  tr.insert(to, bookmarkEnd.create({ id }));
  tr.insert(from, bookmarkStart.create({ name, id }));
  ed.view.dispatch(tr.setMeta('skipTrackChanges', true));
}

/** Removes the bookmark called name; its text stays */
export function removeBookmark(ed, name) {
  const b = listBookmarks(ed.state.doc).find((x) => x.name === name);
  if (!b) return false;
  const tr = ed.state.tr;
  const endPos = b.to;
  const endNode = ed.state.doc.nodeAt(endPos);
  if (endNode?.type.name === 'bookmarkEnd') tr.delete(endPos, endPos + endNode.nodeSize);
  tr.delete(b.from - 1, b.from);
  ed.view.dispatch(tr.setMeta('skipTrackChanges', true));
  return true;
}

/** Selects the bookmark's text (or puts the cursor at it) and scrolls to it */
export function goToBookmark(ed, name) {
  const b = listBookmarks(ed.state.doc).find((x) => x.name === name);
  if (!b) return false;
  ed.commands.setTextSelection({ from: b.from, to: Math.max(b.from, b.to) });
  ed.view.focus();
  ed.commands.scrollIntoView?.();
  return true;
}

const HEADING = /^(heading|berschrift|titre|titolo|t[ií]tulo)\s?\d$/i;
const isHeading = (n) => {
  const pp = n.attrs?.paragraphProperties || {};
  return HEADING.test(pp.styleId || '') || (pp.outlineLevel != null && Number(pp.outlineLevel) < 9);
};
const isCaption = (n) => /^caption$/i.test(n.attrs?.paragraphProperties?.styleId || '');

/** What can be referred to: { kind: 'heading' | 'bookmark' | 'caption', label, pos, name?, level? } */
export function referenceTargets(doc) {
  const out = { heading: [], caption: [], bookmark: [] };
  doc.descendants((n, pos) => {
    if (!n.isTextblock) return true;
    const text = n.textContent.replace(/\s+/g, ' ').trim();
    if (text && isHeading(n)) {
      const level = Number(String(n.attrs.paragraphProperties?.styleId || '').match(/\d$/)?.[0] || (Number(n.attrs.paragraphProperties?.outlineLevel ?? 0) + 1));
      out.heading.push({ kind: 'heading', label: text, pos, level });
    } else if (text && isCaption(n)) out.caption.push({ kind: 'caption', label: text, pos });
    return false;
  });
  out.bookmark = listBookmarks(doc).filter((b) => !b.hidden).map((b) => ({ kind: 'bookmark', label: b.name, name: b.name, pos: b.from - 1, text: b.text }));
  return out;
}

/** The bookmark name to refer to a heading or caption paragraph by, adding a hidden one if needed */
export function bookmarkForParagraph(ed, paraPos) {
  const para = ed.state.doc.nodeAt(paraPos);
  if (!para) throw new Error('That paragraph is no longer there');
  const from = paraPos + 1;
  const to = paraPos + para.nodeSize - 1;
  // a _Ref bookmark already around this paragraph's text
  const existing = listBookmarks(ed.state.doc).find((b) => b.name.startsWith('_Ref') && b.from >= from - 1 && b.from <= from + 1 && b.to <= to + 1 && b.to >= to - 1);
  if (existing) return existing.name;
  const taken = new Set(listBookmarks(ed.state.doc).map((b) => b.name));
  let name;
  do name = `_Ref${String(Math.floor(1e8 + Math.random() * 9e8))}`; while (taken.has(name));
  addBookmark(ed, name, from, to);
  return name;
}

/**
 * Inserts a cross-reference at the cursor.
 * show: 'text' (the target's text), 'page' (its page number) or 'aboveBelow'.
 */
export function insertCrossReference(ed, { bookmark, show = 'text', link = true }, pageOf) {
  const type = ed.schema.nodes.crossReference;
  if (!type) throw new Error('Cross-references are not available in this document');
  const field = show === 'page' ? 'PAGEREF' : 'REF';
  const instruction = [field, bookmark, show === 'aboveBelow' ? '\\p' : '', link ? '\\h' : ''].filter(Boolean).join(' ');
  const at = ed.state.selection.from;
  const attrs = { instruction, fieldType: field, target: bookmark, display: show === 'aboveBelow' ? 'aboveBelow' : show === 'page' ? 'pageNumber' : 'content', resolvedText: '' };
  attrs.resolvedText = crossReferenceText(ed.state.doc, attrs, at, pageOf) || '?';
  ed.view.dispatch(ed.state.tr.replaceSelectionWith(type.create(attrs, null, ed.state.storedMarks || ed.state.selection.$from.marks()), false).scrollIntoView());
}

/**
 * Reads a REF or PAGEREF instruction: { field, name, switches }. Switch arguments (the
 * MERGEFORMAT after \*, say) are skipped, and quoted bookmark names lose their quotes.
 */
export function parseInstruction(instruction, fallbackField = 'REF') {
  const tokens = String(instruction || '').match(/"[^"]*"|\S+/g) || [];
  const field = (tokens[0] || fallbackField).toUpperCase();
  let name = '';
  const switches = [];
  for (let i = 1; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.startsWith('\\')) {
      const sw = t.toLowerCase();
      switches.push(sw);
      // general format switches (\* and \#, \@) take an argument
      if ((sw === '\\*' || sw === '\\#' || sw === '\\@') && i + 1 < tokens.length) i++;
    } else if (!name) name = t.replace(/^"(.*)"$/, '$1');
  }
  return { field, name, switches };
}

// The switches Kertas can work a result out for. Anything else, such as \r, \n or \w (the heading's
// number, "3.2"), \f or \d, is left showing the text Word saved, so opening a document never
// replaces a heading number with the heading's text.
const SUPPORTED_SWITCHES = new Set(['\\h', '\\p', '\\*']);

/** Whether Kertas can work out this cross-reference's text itself */
export function canResolve(attrs) {
  const { field, switches } = parseInstruction(attrs.instruction, attrs.fieldType);
  if (field !== 'REF' && field !== 'PAGEREF') return false;
  // PAGEREF \p reads "on page 3" or "above" depending on the layout; not worked out here
  if (field === 'PAGEREF' && switches.includes('\\p')) return false;
  return switches.every((s) => SUPPORTED_SWITCHES.has(s));
}

/**
 * What a cross-reference shows now; null when its bookmark cannot be found, undefined when it
 * uses a switch Kertas does not work out (its saved text should then be kept as it is)
 */
export function crossReferenceText(doc, attrs, at, pageOf) {
  if (!canResolve(attrs)) return undefined;
  const { field, name: parsed, switches } = parseInstruction(attrs.instruction, attrs.fieldType);
  const name = String(attrs.target || parsed || '').replace(/^"(.*)"$/, '$1');
  const b = listBookmarks(doc).find((x) => x.name === name);
  if (!b) return null;
  if (switches.includes('\\p')) return b.from < at ? 'above' : 'below';
  if (field === 'PAGEREF') { const page = pageOf?.(b.from); return page ? String(page) : null; }
  return b.text || null;
}

/**
 * Brings every cross-reference's text up to date. Returns how many changed; cross-references whose
 * bookmark has gone are left as they are and counted in missing, and ones Kertas cannot work out
 * (see canResolve) are left as they are and counted in kept.
 */
export function refreshCrossReferences(ed, pageOf) {
  const doc = ed.state.doc;
  const changes = [];
  let missing = 0;
  let kept = 0;
  let total = 0;
  doc.descendants((n, pos) => {
    if (n.type.name !== 'crossReference') return true;
    total++;
    const text = crossReferenceText(doc, n.attrs, pos, pageOf);
    if (text === undefined) { kept++; return false; }
    if (text == null) { missing++; return false; }
    if (text !== n.attrs.resolvedText) changes.push({ pos, attrs: { ...n.attrs, resolvedText: text, fieldResultContent: null } });
    return false;
  });
  if (changes.length) {
    const tr = ed.state.tr;
    for (const c of changes) tr.setNodeMarkup(c.pos, null, c.attrs);
    ed.view.dispatch(tr.setMeta('skipTrackChanges', true).setMeta('addToHistory', false));
  }
  return { changed: changes.length, missing, kept, total };
}

/** Whether the document has any cross-references (cheap enough to call after every change) */
export function hasCrossReferences(doc) {
  let found = false;
  doc.descendants((n) => { if (found) return false; if (n.type.name === 'crossReference') { found = true; return false; } return true; });
  return found;
}
