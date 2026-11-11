// PDF layers (optional content groups, the "layers" of Acrobat, Illustrator, InDesign and
// others). Pure module (pdf-lib only), used when saving:
//   - existing layers keep their identity; renames, visibility and deletions are written back
//   - new layers become real optional content groups, and the items drawn on them are wrapped
//     in marked content (/OC /name BDC ... EMC), so any PDF reader shows them as layers
//   - an existing layer's content on a page can be moved or resized: its marked-content blocks
//     (and images or forms marked with /OC) get a transform, nothing is redrawn or rasterised,
//     including blocks inside a text object, as Illustrator and InDesign write them
import { PDFName, PDFArray, PDFDict, PDFRef, PDFHexString, PDFString, PDFRawStream, PDFOperator, decodePDFRawStream } from '@cantoo/pdf-lib';
import { tokenize } from './pdf-textedit.js';

const N = (s) => PDFName.of(s);
const refKey = (r) => (r instanceof PDFRef ? `${r.objectNumber}R${r.generationNumber || ''}` : null);
/** pdf.js group ids look like "12R" (or "12R3" for generation 3) */
export function refFromId(doc, id) {
  const m = /^(\d+)R(\d*)$/.exec(String(id));
  return m ? PDFRef.of(Number(m[1]), m[2] ? Number(m[2]) : 0) : null;
}
const textOf = (v) => (v instanceof PDFHexString || v instanceof PDFString ? v.decodeText() : v ? String(v) : '');

/** The layers a document declares: [{ ref, id, name }] */
export function listLayers(doc) {
  const props = doc.catalog.lookupMaybe(N('OCProperties'), PDFDict);
  const ocgs = props?.lookupMaybe(N('OCGs'), PDFArray);
  if (!ocgs) return [];
  const out = [];
  for (let i = 0; i < ocgs.size(); i++) {
    const ref = ocgs.get(i);
    const d = doc.context.lookup(ref);
    if (d instanceof PDFDict) out.push({ ref, id: refKey(ref), name: textOf(d.get(N('Name'))) });
  }
  return out;
}

// ---------------------------------------------------------------- page resources
function resources(doc, page) {
  return page.node.Resources() || (() => { const r = doc.context.obj({}); page.node.set(N('Resources'), r); return r; })();
}
function subDict(doc, dict, key) {
  let d = dict.lookupMaybe(N(key), PDFDict);
  if (!d) { d = doc.context.obj({}); dict.set(N(key), d); }
  return d;
}
/** The resource name a page uses for a layer (added to /Properties when missing) */
export function layerResourceName(doc, page, ref) {
  const props = subDict(doc, resources(doc, page), 'Properties');
  for (const [k, v] of props.entries()) if (refKey(v) === refKey(ref)) return k.asString().slice(1);
  let i = 0;
  while (props.has(N(`KOC${i}`))) i++;
  props.set(N(`KOC${i}`), ref);
  return `KOC${i}`;
}
export const beginLayer = (name) => PDFOperator.of('BDC', [N('OC'), N(name)]);
export const endLayer = () => PDFOperator.of('EMC');

/** Which layers (by ref key) a /Properties entry or an /OC value stands for */
function groupsOf(doc, v) {
  const d = v instanceof PDFRef ? doc.context.lookup(v) : v;
  if (!(d instanceof PDFDict)) return [];
  const type = d.get(N('Type'))?.asString?.();
  if (type === '/OCMD') {
    const g = d.get(N('OCGs'));
    if (g instanceof PDFArray) return g.asArray().map(refKey).filter(Boolean);
    return [refKey(g)].filter(Boolean);
  }
  return [refKey(v)].filter(Boolean);
}

// ---------------------------------------------------------------- content rewrite
const mul = (m, n) => [
  m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
];
function inv(m) {
  const det = m[0] * m[3] - m[1] * m[2] || 1e-12;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}
const fmt = (x) => (Math.abs(x) < 1e-7 ? '0' : String(+x.toFixed(5)));
const isIdentity = (m) => Math.abs(m[0] - 1) < 1e-9 && Math.abs(m[3] - 1) < 1e-9 && [m[1], m[2], m[4], m[5]].every((v) => Math.abs(v) < 1e-9);

function pageContent(doc, page) {
  const contents = page.node.get(N('Contents'));
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
  page.node.set(N('Contents'), doc.context.register(doc.context.flateStream(bytes)));
}

/**
 * Edits one page's existing layer content.
 * changes: Map(layerRefKey -> { remove: true } | { matrix: [a,b,c,d,e,f] })   matrix in page user space
 * Returns { moved, removed, skipped } counts.
 */
export function rewriteLayerContent(doc, page, changes, done = new Set()) {
  const zero = { moved: 0, removed: 0, skipped: 0 };
  const res = page.node.Resources();
  if (!res || !changes.size) return zero;
  const r = rewriteStream(doc, pageContent(doc, page), res, changes, [1, 0, 0, 1, 0, 0], done, 0);
  if (r.out !== null) setPageContent(doc, page, r.out);
  return r.counts;
}

const decodeStream = (s) => {
  const bytes = s instanceof PDFRawStream ? decodePDFRawStream(s).decode() : s.getContents?.() || new Uint8Array();
  let str = '';
  for (let k = 0; k < bytes.length; k += 0x8000) str += String.fromCharCode.apply(null, bytes.subarray(k, k + 0x8000));
  return str;
};
/** Puts new content into a form XObject in place, keeping its other entries (BBox, Matrix, Resources...) */
function setFormContent(doc, ref, old, str) {
  const bytes = new Uint8Array(str.length);
  for (let k = 0; k < str.length; k++) bytes[k] = str.charCodeAt(k) & 255;
  const ns = doc.context.flateStream(bytes);
  for (const [k, v] of old.dict.entries()) if (!['/Filter', '/DecodeParms', '/Length', '/DL'].includes(k.asString())) ns.dict.set(k, v);
  doc.context.assign(ref, ns);
}

/**
 * One content stream (a page, or a form XObject drawn on it) with its resources.
 * base: the transform in force when the stream starts (a form's Matrix times the CTM at its Do).
 * Forms are followed (up to a few levels deep); a form shared by several pages is changed once.
 * Returns { out: new content or null when unchanged, counts }.
 */
function rewriteStream(doc, src, res, changes, base, done, depth) {
  const counts = { moved: 0, removed: 0, skipped: 0 };
  const props = res.lookupMaybe(N('Properties'), PDFDict);
  const xobjs = res.lookupMaybe(N('XObject'), PDFDict);
  const propTarget = new Map(); // resource name -> change
  const pick = (keys) => { for (const k of keys) if (changes.has(k)) return changes.get(k); return null; };
  if (props) for (const [k, v] of props.entries()) { const c = pick(groupsOf(doc, v)); if (c) propTarget.set(k.asString().slice(1), c); }
  const xoTarget = new Map();
  const forms = new Map(); // name -> ref, for forms that may hold layer content of their own
  if (xobjs) {
    for (const [k, v] of xobjs.entries()) {
      const st = v instanceof PDFRef ? doc.context.lookup(v) : v;
      const oc = st?.dict?.get?.(N('OC'));
      const c = oc ? pick(groupsOf(doc, oc)) : null;
      if (c) xoTarget.set(k.asString().slice(1), c);
      else if (v instanceof PDFRef && st?.dict?.get?.(N('Subtype'))?.asString?.() === '/Form' && depth < 6) forms.set(k.asString().slice(1), v);
    }
  }
  if (!propTarget.size && !xoTarget.size && !forms.size) return { out: null, counts };
  const formCalls = [];

  const toks = tokenize(src);
  // operators with their operands
  const ops = [];
  let operands = [];
  for (const t of toks) {
    if (t.t !== 'op') { operands.push(t); continue; }
    ops.push({ op: t.v, args: operands, start: operands[0] ? operands[0].s : t.s, opStart: t.s, end: t.e });
    operands = [];
  }
  const edits = []; // { at, end, text } replacing src[at, end)
  const insertCm = (m, c) => `q ${mul(mul(c, m), inv(c)).map(fmt).join(' ')} cm\n`;
  const raw = (o) => src.slice(o.start, o.end);
  const num = (a) => (typeof a?.v === 'number' ? a.v : Number(a?.v) || 0);

  // graphics and text state that a moved block may set and the content after it may rely on:
  // our own q/Q around the block would undo it, so it is set again after the Q
  const STATE = new Set(['w', 'J', 'j', 'M', 'd', 'ri', 'i', 'gs', 'CS', 'cs', 'SC', 'SCN', 'sc', 'scn', 'G', 'g', 'RG', 'rg', 'K', 'k', 'Tc', 'Tw', 'Tz', 'TL', 'Tf', 'Tr', 'Ts']);
  let ctm = base;
  const stack = [];
  let inText = false;
  // text position: line matrix, text matrix, and whether text was shown since the last line move
  let tlm = [1, 0, 0, 1, 0, 0]; let tm = tlm; let dirty = false; let leading = 0;
  const lineTo = (tx, ty) => { tlm = mul([1, 0, 0, 1, tx, ty], tlm); tm = tlm; dirty = false; };
  const marks = []; // open BDC/BMC: { change, start, bodyStart, ctm, inText, tm, dirty, state, depth } or null
  let removeDepth = 0; // inside a block being removed: ignore nested targets
  const open = () => marks.filter(Boolean).filter((m) => m.change.matrix);
  const SHOW = new Set(['Tj', 'TJ', "'", '"']);

  for (let k = 0; k < ops.length; k++) {
    const o = ops[k]; const op = o.op; const a = o.args;
    for (const m of open()) {
      if (op === 'q') m.depth++;
      else if (op === 'Q') m.depth--;
      else if (m.depth === 0 && STATE.has(op)) { m.state.delete(op); m.state.set(op, raw(o)); }
    }
    switch (op) {
      case 'q': stack.push(ctm); break;
      case 'Q': ctm = stack.pop() || ctm; break;
      case 'cm':
        if (a.length === 6) ctm = mul(a.map((x) => x.v), ctm);
        break;
      case 'BT': inText = true; tlm = [1, 0, 0, 1, 0, 0]; tm = tlm; dirty = false; break;
      case 'ET': inText = false; break;
      case 'Tm': if (a.length === 6) { tlm = a.map(num); tm = tlm; dirty = false; } break;
      case 'Td': lineTo(num(a[0]), num(a[1])); break;
      case 'TD': leading = -num(a[1]); lineTo(num(a[0]), num(a[1])); break;
      case 'TL': leading = num(a[0]); break;
      case 'T*': lineTo(0, -leading); break;
      case "'": case '"': lineTo(0, -leading); dirty = true; break;
      case 'Tj': case 'TJ': dirty = true; break;
      case 'BDC': case 'BMC': {
        let change = null;
        if (op === 'BDC' && a[0]?.t === 'name' && a[0].v === 'OC' && a[1]?.t === 'name') change = propTarget.get(a[1].v) || null;
        if (removeDepth) change = null;
        if (change?.matrix && isIdentity(change.matrix)) change = null;
        // in the middle of a line of text the position after the block cannot be known
        if (change?.matrix && inText && dirty) { counts.skipped++; change = null; }
        marks.push(change ? { change, start: o.start, bodyStart: o.end, ctm, inText, tm, state: new Map(), depth: 0 } : null);
        if (change?.remove) removeDepth++;
        break;
      }
      case 'EMC': {
        const m = marks.pop();
        if (!m) break;
        if (m.change.remove) {
          removeDepth--;
          edits.push({ at: m.start, end: o.end, text: ' ' });
          counts.removed++;
          break;
        }
        if (inText && dirty) {
          // text shown at the end of the block: fine only if what follows starts a new line
          let next = null;
          for (let n = k + 1; n < ops.length; n++) { const x = ops[n].op; if (SHOW.has(x) || ['Tm', 'Td', 'TD', 'T*', 'ET', 'BT'].includes(x)) { next = x; break; } }
          if (next === 'Tj' || next === 'TJ') { counts.skipped++; break; }
        }
        const head = m.inText ? `\nET\n${insertCm(m.change.matrix, m.ctm)}BT ${m.tm.map(fmt).join(' ')} Tm\n` : `\n${insertCm(m.change.matrix, m.ctm)}`;
        const restore = [...m.state.values()].join('\n');
        const tail = inText ? `ET Q\n${restore}${restore ? '\n' : ''}BT ${tlm.map(fmt).join(' ')} Tm\n` : `Q\n${restore}${restore ? '\n' : ''}`;
        edits.push({ at: m.bodyStart, end: m.bodyStart, text: head });
        edits.push({ at: o.start, end: o.start, text: tail });
        counts.moved++;
        break;
      }
      case 'Do': {
        const nm = a[0]?.t === 'name' ? a[0].v : null;
        const change = nm && !removeDepth ? xoTarget.get(nm) : null;
        if (!change && nm && forms.has(nm) && !removeDepth) formCalls.push({ name: nm, ctm });
        if (change?.remove) { edits.push({ at: o.start, end: o.end, text: ' ' }); counts.removed++; }
        else if (change?.matrix && !isIdentity(change.matrix)) {
          edits.push({ at: o.start, end: o.start, text: insertCm(change.matrix, ctm) });
          edits.push({ at: o.end, end: o.end, text: '\nQ' });
          counts.moved++;
        }
        break;
      }
      default:
    }
  }
  // layer content inside forms (some apps draw each layer as a reusable form)
  for (const call of formCalls) {
    const ref = forms.get(call.name);
    const key = `${ref.objectNumber} ${ref.generationNumber}`;
    if (done.has(key)) continue;
    done.add(key);
    const st = doc.context.lookup(ref);
    const fres = st.dict.lookupMaybe(N('Resources'), PDFDict) || res;
    const m = st.dict.lookupMaybe(N('Matrix'), PDFArray)?.asArray().map((x) => x.asNumber?.() ?? 0);
    const fbase = m?.length === 6 ? mul(m, call.ctm) : call.ctm;
    const sub = rewriteStream(doc, decodeStream(st), fres, changes, fbase, done, depth + 1);
    if (sub.out !== null) setFormContent(doc, ref, st, sub.out);
    counts.moved += sub.counts.moved; counts.removed += sub.counts.removed; counts.skipped += sub.counts.skipped;
  }
  if (!edits.length) return { out: null, counts };
  edits.sort((a, b) => b.at - a.at || b.end - a.end);
  let out = src;
  for (const e of edits) out = out.slice(0, e.at) + e.text + out.slice(e.end);
  return { out, counts };
}

// ---------------------------------------------------------------- catalog
/**
 * Writes the document's layer list.
 *   layers: [{ ref, name?, visible }]   every layer that should exist, in panel order (top first)
 *   removed: [ref]                      layers taken out of the document
 */
export function writeLayerCatalog(doc, { layers = [], removed = [] }) {
  if (!layers.length && !removed.length) return;
  const gone = new Set(removed.map(refKey));
  for (const l of layers) {
    if (l.name === undefined) continue;
    const d = doc.context.lookup(l.ref);
    if (d instanceof PDFDict) d.set(N('Name'), PDFHexString.fromText(l.name));
  }
  let props = doc.catalog.lookupMaybe(N('OCProperties'), PDFDict);
  if (!props) { props = doc.context.obj({}); doc.catalog.set(N('OCProperties'), props); }
  const known = new Map();
  const old = props.lookupMaybe(N('OCGs'), PDFArray);
  if (old) for (const r of old.asArray()) if (!gone.has(refKey(r))) known.set(refKey(r), r);
  for (const l of layers) known.set(refKey(l.ref), l.ref);
  props.set(N('OCGs'), doc.context.obj([...known.values()]));

  let d = props.lookupMaybe(N('D'), PDFDict);
  if (!d) { d = doc.context.obj({}); props.set(N('D'), d); }
  if (!d.has(N('Name'))) d.set(N('Name'), PDFHexString.fromText('Layers'));
  const visible = new Map(layers.map((l) => [refKey(l.ref), l.visible !== false]));
  const wasOff = new Set((d.lookupMaybe(N('OFF'), PDFArray)?.asArray() || []).map(refKey));
  const on = []; const off = [];
  for (const [k, r] of known) (visible.has(k) ? visible.get(k) : !wasOff.has(k)) ? on.push(r) : off.push(r);
  d.set(N('BaseState'), N('ON'));
  d.set(N('ON'), doc.context.obj(on));
  d.set(N('OFF'), doc.context.obj(off));
  // Order: the panel order, keeping any nested structure the file had for layers we did not touch
  const ordered = layers.map((l) => l.ref);
  const listed = new Set(ordered.map(refKey));
  const keepNested = (arr) => arr.asArray().map((v) => {
    const x = v instanceof PDFRef ? v : v;
    if (x instanceof PDFArray) { const inner = keepNested(x); return inner.length ? doc.context.obj(inner) : null; }
    if (x instanceof PDFRef && (gone.has(refKey(x)) || listed.has(refKey(x)))) return null;
    return x;
  }).filter(Boolean);
  const prev = d.lookupMaybe(N('Order'), PDFArray);
  d.set(N('Order'), doc.context.obj([...ordered, ...(prev ? keepNested(prev) : [])]));
  for (const key of ['RBGroups', 'Locked']) {
    const a = d.lookupMaybe(N(key), PDFArray);
    if (!a) continue;
    const clean = (arr) => arr.asArray().map((v) => (v instanceof PDFArray ? doc.context.obj(clean(v)) : gone.has(refKey(v)) ? null : v)).filter(Boolean);
    d.set(N(key), doc.context.obj(clean(a)));
  }
  // AS (auto state) entries can switch layers back on by themselves; keep them only for live layers
  if (d.has(N('AS'))) d.delete(N('AS'));
}

/** Creates a new layer and returns its ref */
export function createLayer(doc, name) {
  const dict = doc.context.obj({ Type: 'OCG', Name: PDFHexString.fromText(name), Intent: [N('View'), N('Design')] });
  dict.set(N('Usage'), doc.context.obj({ CreatorInfo: { Creator: PDFHexString.fromText('Kertas Office'), Subtype: N('Artwork') } }));
  return doc.context.register(dict);
}

/** Layers that a page's resources point at, by name: { name: ref } (used after pages are copied) */
export function layersOnPages(doc, pages) {
  const out = {};
  const note = (v) => {
    const d = v instanceof PDFRef ? doc.context.lookup(v) : null;
    if (!(d instanceof PDFDict)) return;
    const type = d.get(N('Type'))?.asString?.();
    if (type === '/OCG') out[textOf(d.get(N('Name')))] ||= v;
    else if (type === '/OCMD') { const g = d.get(N('OCGs')); (g instanceof PDFArray ? g.asArray() : [g]).forEach(note); }
  };
  for (const page of pages) {
    const res = page.node.Resources();
    const props = res?.lookupMaybe(N('Properties'), PDFDict);
    if (props) for (const [, v] of props.entries()) note(v);
    const xo = res?.lookupMaybe(N('XObject'), PDFDict);
    if (xo) for (const [, v] of xo.entries()) { const s = v instanceof PDFRef ? doc.context.lookup(v) : v; const oc = s?.dict?.get?.(N('OC')); if (oc) note(oc); }
  }
  return out;
}
