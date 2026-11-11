// PDF editor: added items stay objects (select, move, resize, delete, stacking order), and the
// document's layers (optional content groups) are shown and edited in the Layers panel.
//
// Items keep their PDF coordinates in the page model (see pdf-ops.js). Moving or resizing an
// item maps its points through the change seen on screen, so rotated pages work the same way.
// A layer that came with the file can have its content on a page lifted, moved and resized:
// it is shown as its own picture (the page drawn with and without the layer) and saved as a
// transform on the layer's own content, so nothing is flattened.
import { icon } from '../lib/icons.js';
import { openMenu } from '../lib/ribbon.js';
import { openDialog } from '../lib/dialog.js';

const SVGNS = 'http://www.w3.org/2000/svg';
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const MOVABLE = new Set(['image', 'stamp', 'text', 'rect', 'ellipse', 'line', 'ink', 'highlight', 'note', 'redact', 'field']);
const KEEP_ASPECT = new Set(['image', 'stamp', 'text', 'note']);
const ID = [1, 0, 0, 1, 0, 0];
// row-vector affine maths, as PDF uses: p' = p x M
export const mul = (m, n) => [
  m[0] * n[0] + m[1] * n[2], m[0] * n[1] + m[1] * n[3],
  m[2] * n[0] + m[3] * n[2], m[2] * n[1] + m[3] * n[3],
  m[4] * n[0] + m[5] * n[2] + n[4], m[4] * n[1] + m[5] * n[3] + n[5],
];
export function inv(m) {
  const det = m[0] * m[3] - m[1] * m[2] || 1e-12;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}
const applyM = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
const isId = (m) => !m || m.every((v, i) => Math.abs(v - ID[i]) < 1e-9);
export const isMovable = (a) => MOVABLE.has(a?.type);
const LABEL = { image: 'Image', stamp: 'Stamp', text: 'Text', rect: 'Box', ellipse: 'Ellipse', line: 'Line', ink: 'Drawing', highlight: 'Area', note: 'Comment', redact: 'Redaction', field: 'Form field', markup: 'Text markup', editText: 'Edited text' };
const ICON = { image: 'image', stamp: 'stamp', text: 'text', rect: 'rect', ellipse: 'ellipse', line: 'line', ink: 'ink', highlight: 'highlight', note: 'note', redact: 'redact', field: 'field-text', markup: 'highlight', editText: 'text-cursor' };
export const describe = (a) => (a.type === 'text' ? `Text: ${String(a.text).split('\n')[0].slice(0, 24)}` : a.type === 'stamp' ? `Stamp: ${a.label}` : a.type === 'line' && a.arrow ? 'Arrow' : a.type === 'field' ? `Field: ${a.name}` : LABEL[a.type] || a.type);

/** A copy of an item's geometry, for undo */
function snap(a) {
  return { x: a.x, y: a.y, w: a.w, h: a.h, x1: a.x1, y1: a.y1, x2: a.x2, y2: a.y2, size: a.size, labelSize: a.labelSize, points: a.points?.map((q) => [...q]), rects: a.rects?.map((r) => [...r]), layer: a.layer };
}
const restore = (a, s) => { for (const [k, v] of Object.entries(s)) if (v !== undefined || k === 'layer') a[k] = v; };

/**
 * Maps an item through a change seen on screen.
 * f(viewPoint) -> viewPoint; sx, sy: the change in size along the screen axes.
 */
export function mapAnnot(a, vp, f, sx = 1, sy = 1) {
  const P = (x, y) => vp.convertToPdfPoint(...f(vp.convertToViewportPoint(x, y)));
  const swap = Math.abs(((((vp.rotation - (a.rot || 0)) % 180) + 180) % 180) - 90) < 1;
  const [ax, ay] = swap ? [sy, sx] : [sx, sy];
  switch (a.type) {
    case 'image': case 'stamp': {
      // x, y is the box's lower-left corner as it was seen when placed
      [a.x, a.y] = P(a.x, a.y); a.w *= ax; a.h *= ay;
      if (a.labelSize) a.labelSize *= ay;
      break;
    }
    case 'text': [a.x, a.y] = P(a.x, a.y); a.size = Math.max(2, a.size * ay); break;
    case 'note': [a.x, a.y] = P(a.x, a.y); break;
    case 'ink': a.points = a.points.map(([x, y]) => P(x, y)); break;
    case 'markup': a.rects = a.rects.map(([x1, y1, x2, y2]) => [...P(x1, y1), ...P(x2, y2)]); break;
    default: [a.x1, a.y1] = P(a.x1, a.y1); [a.x2, a.y2] = P(a.x2, a.y2);
  }
}

export function createObjects(ctx) {
  const { pagesEl, sideEl, pages, rendered, sources, ribbon, toast, changed, undoStack } = ctx;
  let sel = null; // { p, a } | { p, layer: key }
  // ---------------------------------------------------------------- layers
  const layers = []; // panel order, top first: { key, src, id, name, visible, deleted, created }
  let active = null; // key that new items go on, or null for the page itself
  let newSeq = 1;
  const layerOf = (key) => layers.find((l) => l.key === key);
  const isVisible = (key) => { const l = key && layerOf(key); return !l || (l.visible && !l.deleted); };

  async function loadSource(srcId) {
    let cfg;
    try { cfg = await sources[srcId].doc.getOptionalContentConfig(); } catch { return; }
    const order = cfg?.getOrder?.();
    if (!order) return;
    const flat = [];
    const walk = (arr) => arr.forEach((x) => (typeof x === 'string' ? flat.push(x) : Array.isArray(x) ? walk(x) : x?.order && walk(x.order)));
    walk(order);
    for (const [gid] of cfg) if (!flat.includes(gid)) flat.push(gid);
    for (const gid of flat) {
      const g = cfg.getGroup(gid);
      if (!g || layers.some((l) => l.src === srcId && l.id === gid)) continue;
      layers.push({ key: `${srcId}|${gid}`, src: srcId, id: gid, name: g.name || 'Layer', origName: g.name || '', visible: g.visible, deleted: false, created: false });
    }
    renderPanel();
  }
  const hasPdfLayers = (src) => layers.some((l) => l.src === src && l.id && !l.created);
  /** pdf.js visibility for a page, with some layers hidden on top (the lifted ones) */
  async function configFor(p, hide) {
    if (!hasPdfLayers(p.src)) return null;
    const cfg = await sources[p.src].doc.getOptionalContentConfig();
    for (const l of layers) if (l.src === p.src && l.id) cfg.setVisibility(l.id, l.visible && !l.deleted && !hide.has(l.key), false);
    return cfg;
  }
  const liftedKeys = (p) => new Set(Object.keys(p.layerMoves || {}).filter((k) => isVisible(k) && layerOf(k)));

  /** Draws each lifted layer as its own picture over the page (the page with and without it) */
  async function renderOverlays(p, pdfPage, vp, div, base, entry, ratio, annotationMode) {
    entry.overlays = {};
    const lifted = liftedKeys(p);
    if (!lifted.size) return;
    const w = base.width; const h = base.height;
    const baseData = base.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
    for (const key of lifted) {
      const hide = new Set([...lifted].filter((k) => k !== key));
      const cfg = await configFor(p, hide);
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      await pdfPage.render({ canvas: c, viewport: vp, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : null, optionalContentConfigPromise: Promise.resolve(cfg), annotationMode }).promise;
      const cx = c.getContext('2d', { willReadFrequently: true });
      const img = cx.getImageData(0, 0, w, h);
      const d = img.data;
      let x0 = w; let y0 = h; let x1 = -1; let y1 = -1;
      for (let i = 0, px = 0; i < d.length; i += 4, px++) {
        const diff = Math.max(Math.abs(d[i] - baseData[i]), Math.abs(d[i + 1] - baseData[i + 1]), Math.abs(d[i + 2] - baseData[i + 2]));
        if (diff < 6) { d[i + 3] = 0; continue; }
        const x = px % w; const y = (px / w) | 0;
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
      }
      // pdf.js draws on an opaque canvas; the lifted picture needs one with transparency
      const lift = document.createElement('canvas');
      lift.width = w; lift.height = h;
      lift.getContext('2d').putImageData(img, 0, 0);
      c.width = 0; // free the scratch copy
      const cOut = lift;
      cOut.className = 'layer-lift';
      cOut.style.width = `${vp.width}px`; cOut.style.height = `${vp.height}px`;
      cOut.dataset.layer = key;
      base.after(cOut);
      entry.overlays[key] = { canvas: cOut, box: x1 < 0 ? null : { x: x0 / ratio, y: y0 / ratio, w: (x1 - x0 + 1) / ratio, h: (y1 - y0 + 1) / ratio } };
      placeOverlay(p, key, entry);
    }
  }
  /** The lifted layer's change as seen on screen */
  const viewMatrix = (vp, m) => mul(mul(inv(vp.transform), m), vp.transform);
  function placeOverlay(p, key, entry, extra = null) {
    const o = entry?.overlays?.[key];
    if (!o) return;
    let T = viewMatrix(entry.viewport, p.layerMoves[key] || ID);
    if (extra) T = mul(T, extra);
    o.canvas.style.transform = `matrix(${T.join(',')})`;
  }
  function overlayBox(p, key, entry) {
    const o = entry?.overlays?.[key];
    if (!o?.box) return null;
    const T = viewMatrix(entry.viewport, p.layerMoves[key] || ID);
    const { x, y, w, h } = o.box;
    const pts = [[x, y], [x + w, y], [x, y + h], [x + w, y + h]].map(([a, b]) => applyM(T, a, b));
    const xs = pts.map((q) => q[0]); const ys = pts.map((q) => q[1]);
    return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
  }

  // ---------------------------------------------------------------- selection frame
  const svgEl = (tag, attrs) => { const n = document.createElementNS(SVGNS, tag); for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, v); return n; };
  function nodeFor(p, a) {
    const svg = rendered.get(p.key)?.el.querySelector('svg.annots');
    const idx = p.annots.indexOf(a);
    return svg?.querySelector(`:scope > [data-idx="${idx}"]`) || null;
  }
  function boxOf(p, target) {
    const entry = rendered.get(p.key);
    if (!entry) return null;
    if (target.layer) return overlayBox(p, target.layer, entry);
    const node = nodeFor(p, target.a);
    if (!node) return null;
    const r = node.getBoundingClientRect(); const s = entry.el.getBoundingClientRect();
    const pad = target.a.type === 'line' || target.a.type === 'ink' ? 3 : 0;
    return { x: r.left - s.left - pad, y: r.top - s.top - pad, w: Math.max(4, r.width + pad * 2), h: Math.max(4, r.height + pad * 2) };
  }
  const HANDLES = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
  function drawFrame(p) {
    const entry = rendered.get(p.key);
    const svg = entry?.el.querySelector('svg.annots');
    if (!svg) return;
    svg.querySelector('.sel-frame')?.remove();
    if (!sel || sel.p !== p) return;
    const b = boxOf(p, sel);
    if (!b) return;
    const g = svgEl('g', { class: 'sel-frame' });
    g.appendChild(svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, class: 'sel-box', 'data-h': 'move' }));
    const noResize = sel.a?.type === 'note';
    if (!noResize) for (const [h, [fx, fy]] of Object.entries(HANDLES)) g.appendChild(svgEl('rect', { x: b.x + b.w * fx - 4, y: b.y + b.h * fy - 4, width: 8, height: 8, rx: 2, class: `sel-handle h-${h}`, 'data-h': h }));
    svg.appendChild(g);
  }

  function select(target) {
    const prev = sel?.p;
    sel = target;
    if (prev && prev !== target?.p) drawFrame(prev);
    if (target) drawFrame(target.p);
    ribbon.setTabVisible('object', !!target);
    // the Form field group only shows for fields
    const fg = ribbon.get('obj-field')?.closest('.rb-group');
    if (fg) { const on = target?.a?.type === 'field'; fg.hidden = !on; if (fg.previousElementSibling?.classList.contains('rb-sep')) fg.previousElementSibling.hidden = !on; }
    syncFields();
    renderPanel();
  }
  const clear = () => { if (sel) select(null); };
  function syncFields() {
    if (!sel) return;
    const entry = rendered.get(sel.p.key);
    const b = boxOf(sel.p, sel);
    const s = entry?.viewport?.scale || 1;
    ribbon.setValue('obj-w', b ? Math.round(b.w / s) : '');
    ribbon.setValue('obj-h', b ? Math.round(b.h / s) : '');
  }

  // ---------------------------------------------------------------- dragging
  function commit(p, target, M) {
    // M: the change on screen (row-vector affine), applied to the item
    const entry = rendered.get(p.key);
    const vp = entry.viewport;
    const sx = Math.hypot(M[0], M[1]); const sy = Math.hypot(M[2], M[3]);
    if (target.layer) {
      const before = p.layerMoves[target.layer] || ID;
      // the change in PDF space: P -> view -> M -> back to PDF
      const pdfChange = mul(mul(vp.transform, M), inv(vp.transform));
      p.layerMoves[target.layer] = mul(before, pdfChange);
      undoStack.push({ type: 'layermove', p, key: target.layer, before });
      placeOverlay(p, target.layer, entry);
    } else {
      undoStack.push({ type: 'modify', p, a: target.a, before: snap(target.a) });
      mapAnnot(target.a, vp, ([x, y]) => applyM(M, x, y), sx, sy);
      ctx.drawAnnots(p);
    }
    changed();
    drawFrame(p);
    syncFields();
  }

  function startDrag(e, p, target, handle) {
    const entry = rendered.get(p.key);
    const svg = entry.el.querySelector('svg.annots');
    const b0 = boxOf(p, target);
    if (!b0) return;
    svg.setPointerCapture(e.pointerId);
    const sx0 = e.clientX; const sy0 = e.clientY;
    // preview: the item (or the lifted layer) and the frame follow the pointer
    let wrap = null;
    if (!target.layer) {
      const node = nodeFor(p, target.a);
      if (node) { wrap = svgEl('g', {}); node.before(wrap); wrap.appendChild(node); }
    }
    const frame = svg.querySelector('.sel-frame');
    let M = ID;
    let moved = false;
    const keepAspect = !!target.layer || KEEP_ASPECT.has(target.a?.type);
    const move = (ev) => {
      const dx = ev.clientX - sx0; const dy = ev.clientY - sy0;
      if (!moved && Math.hypot(dx, dy) < 2) return;
      moved = true;
      if (handle === 'move') M = [1, 0, 0, 1, dx, dy];
      else {
        let { x, y, w, h } = b0;
        let x2 = x + w; let y2 = y + h;
        if (handle.includes('w')) x = Math.min(x2 - 6, b0.x + dx);
        if (handle.includes('e')) x2 = Math.max(x + 6, b0.x + b0.w + dx);
        if (handle.includes('n')) y = Math.min(y2 - 6, b0.y + dy);
        if (handle.includes('s')) y2 = Math.max(y + 6, b0.y + b0.h + dy);
        let ksx = (x2 - x) / b0.w; let ksy = (y2 - y) / b0.h;
        const corner = handle.length === 2;
        if (corner && keepAspect !== ev.shiftKey) {
          const k = Math.max(ksx, ksy); ksx = k; ksy = k;
          if (handle.includes('w')) x = b0.x + b0.w - b0.w * k; else x = b0.x;
          if (handle.includes('n')) y = b0.y + b0.h - b0.h * k; else y = b0.y;
        } else if (!corner && keepAspect && target.a?.type === 'text') {
          // text grows with its font size either way
          const k = handle === 'e' || handle === 'w' ? ksx : ksy; ksx = k; ksy = k;
        }
        M = [ksx, 0, 0, ksy, x - b0.x * ksx, y - b0.y * ksy];
      }
      const css = `matrix(${M.join(' ')})`;
      if (wrap) wrap.setAttribute('transform', css);
      if (target.layer) placeOverlay(p, target.layer, entry, M);
      frame?.setAttribute('transform', css);
      ctx.setStatus?.(`${Math.round((b0.w * Math.hypot(M[0], M[1])) / entry.viewport.scale)} x ${Math.round((b0.h * Math.hypot(M[2], M[3])) / entry.viewport.scale)} pt`);
    };
    const up = () => {
      svg.removeEventListener('pointermove', move);
      svg.removeEventListener('pointerup', up);
      svg.removeEventListener('pointercancel', up);
      if (moved) commit(p, target, M);
      else if (wrap) { wrap.replaceWith(...wrap.childNodes); }
    };
    svg.addEventListener('pointermove', move);
    svg.addEventListener('pointerup', up);
    svg.addEventListener('pointercancel', up);
  }

  /** Select tool: returns true when the pointer went to an item or the frame */
  function pointerDown(e, p) {
    const h = e.target.closest?.('[data-h]')?.dataset.h;
    if (h && sel?.p === p) { e.preventDefault(); e.stopPropagation(); startDrag(e, p, sel, h); return true; }
    const idx = e.target.closest?.('.annot')?.dataset.idx;
    const a = idx !== undefined ? p.annots[Number(idx)] : null;
    if (a && isMovable(a)) {
      e.preventDefault(); e.stopPropagation();
      const target = { p, a };
      if (sel?.a !== a) select(target);
      startDrag(e, p, sel, 'move');
      return true;
    }
    return false;
  }
  // capture phase: runs before a tool places something new (which may select it)
  pagesEl.addEventListener('pointerdown', (e) => {
    if (!sel || e.target.closest('.sel-frame, .annot.movable')) return;
    clear();
  }, true);

  // ---------------------------------------------------------------- commands
  function remove() {
    if (!sel) return;
    const { p } = sel;
    if (sel.layer) { resetLayer(p, sel.layer); return; }
    const idx = p.annots.indexOf(sel.a);
    if (idx < 0) return;
    p.annots.splice(idx, 1);
    undoStack.push({ type: 'remove', p, a: sel.a, idx });
    sel = null;
    ribbon.setTabVisible('object', false);
    ctx.drawAnnots(p); changed(); renderPanel();
  }
  function order(how) {
    if (!sel?.a) return;
    const { p, a } = sel;
    const list = p.annots;
    const before = [...list];
    const i = list.indexOf(a);
    // stacking stays inside the item's layer
    const same = list.map((x, k) => [x, k]).filter(([x]) => (x.layer || null) === (a.layer || null)).map(([, k]) => k);
    const at = same.indexOf(i);
    const target = how === 'front' ? same[same.length - 1] : how === 'back' ? same[0] : how === 'forward' ? same[Math.min(same.length - 1, at + 1)] : same[Math.max(0, at - 1)];
    if (target === undefined || target === i) return;
    list.splice(i, 1);
    list.splice(target, 0, a);
    undoStack.push({ type: 'order', p, before });
    ctx.drawAnnots(p); drawFrame(p); changed(); renderPanel();
  }
  function duplicate() {
    if (!sel?.a) return;
    const { p, a } = sel;
    const copy = { ...a, ...snap(a), img: a.img, png: a.png };
    const entry = rendered.get(p.key);
    mapAnnot(copy, entry.viewport, ([x, y]) => [x + 14, y + 14]);
    p.annots.splice(p.annots.indexOf(a) + 1, 0, copy);
    undoStack.push({ type: 'add', p, a: copy });
    ctx.drawAnnots(p); changed();
    select({ p, a: copy });
  }
  function nudge(dx, dy) {
    if (!sel) return;
    commit(sel.p, sel, [1, 0, 0, 1, dx, dy]);
  }
  function setSize(dim, pt) {
    if (!sel || !Number.isFinite(pt) || pt <= 0) return;
    const entry = rendered.get(sel.p.key);
    const b = boxOf(sel.p, sel);
    if (!b) return;
    const s = entry.viewport.scale;
    const keep = !sel.layer && KEEP_ASPECT.has(sel.a.type);
    let kx = dim === 'w' ? (pt * s) / b.w : 1; let ky = dim === 'h' ? (pt * s) / b.h : 1;
    if (keep) { const k = dim === 'w' ? kx : ky; kx = k; ky = k; }
    commit(sel.p, sel, [kx, 0, 0, ky, b.x - b.x * kx, b.y - b.y * ky]);
  }
  function moveToLayer(key) {
    if (!sel?.a) return;
    const { p, a } = sel;
    undoStack.push({ type: 'modify', p, a, before: snap(a) });
    a.layer = key || undefined;
    stack(p);
    ctx.drawAnnots(p); drawFrame(p); changed(); renderPanel();
    toast(key ? `Moved to ${layerOf(key)?.name}` : 'Moved to the page');
  }
  /** Keeps each page's items in layer order: the page's own items first, then layers bottom to top */
  function stack(p) {
    const rank = (a) => { if (!a.layer) return -1; const i = layers.findIndex((l) => l.key === a.layer); return i < 0 ? -1 : layers.length - i; };
    const sorted = p.annots.map((a, i) => [a, i]).sort((x, y) => rank(x[0]) - rank(y[0]) || x[1] - y[1]).map(([a]) => a);
    p.annots.splice(0, p.annots.length, ...sorted);
  }
  function undoStep(last) {
    if (last.type === 'modify') { restore(last.a, last.before); return true; }
    if (last.type === 'order') { last.p.annots.splice(0, last.p.annots.length, ...last.before); return true; }
    if (last.type === 'layermove') {
      if (isId(last.before)) delete last.p.layerMoves[last.key]; else last.p.layerMoves[last.key] = last.before;
      ctx.refreshPage(last.p);
      return true;
    }
    return false;
  }
  function afterUndo(p) { if (sel?.p === p) { if (sel.a && !p.annots.includes(sel.a)) clear(); else { drawFrame(p); syncFields(); } } renderPanel(); }

  function onKey(e) {
    if (!sel) return false;
    const mod = e.ctrlKey || e.metaKey;
    const step = e.shiftKey ? 10 : 1;
    switch (e.key) {
      case 'Delete': case 'Backspace': e.preventDefault(); remove(); return true;
      case 'Escape': clear(); return true;
      case 'ArrowLeft': e.preventDefault(); nudge(-step, 0); return true;
      case 'ArrowRight': e.preventDefault(); nudge(step, 0); return true;
      case 'ArrowUp': e.preventDefault(); nudge(0, -step); return true;
      case 'ArrowDown': e.preventDefault(); nudge(0, step); return true;
      default:
    }
    if (mod && e.key.toLowerCase() === 'd') { e.preventDefault(); duplicate(); return true; }
    if (mod && (e.code === 'BracketRight' || e.key === ']')) { e.preventDefault(); order(e.shiftKey ? 'front' : 'forward'); return true; }
    if (mod && (e.code === 'BracketLeft' || e.key === '[')) { e.preventDefault(); order(e.shiftKey ? 'back' : 'backward'); return true; }
    return false;
  }

  // ---------------------------------------------------------------- layer commands
  async function nameDialog(title, value, okLabel) {
    const res = await openDialog({
      title, width: 400,
      fields: [{ id: 'name', label: 'Layer name', value }],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: okLabel, primary: true }],
      validate: (v) => (v.name.trim() ? null : 'Enter a name'),
    });
    return res ? res.values.name.trim() : null;
  }
  async function newLayer() {
    const name = await nameDialog('New layer', `Layer ${layers.filter((l) => !l.deleted).length + 1}`, 'Add layer');
    if (!name) return;
    const key = `new${newSeq++}`;
    layers.unshift({ key, src: null, id: null, name, visible: true, deleted: false, created: true });
    active = key;
    changed(); renderPanel();
    showPanel();
    toast(`New items now go on "${name}"`);
  }
  async function rename(key) {
    const l = layerOf(key);
    const name = l && await nameDialog('Rename layer', l.name, 'Rename');
    if (!name) return;
    l.name = name; changed(); renderPanel();
  }
  async function deleteLayer(key) {
    const l = layerOf(key);
    if (!l) return;
    const own = pages.reduce((n, p) => n + p.annots.filter((a) => a.layer === key).length, 0);
    const res = await openDialog({
      title: `Delete "${l.name}"?`, width: 440,
      description: l.created ? `The ${own} item${own === 1 ? '' : 's'} on this layer are deleted with it.` : 'Everything on this layer is taken out of the PDF when you save. The rest of the page is not touched.',
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Delete layer', danger: true, primary: true }],
    });
    if (!res || res.action !== 'ok') return;
    if (sel?.layer === key || sel?.a?.layer === key) clear();
    for (const p of pages) { const keep = p.annots.filter((a) => a.layer !== key); if (keep.length !== p.annots.length) p.annots.splice(0, p.annots.length, ...keep); if (p.layerMoves) delete p.layerMoves[key]; }
    if (l.created) layers.splice(layers.indexOf(l), 1); else l.deleted = true;
    if (active === key) active = null;
    changed(); renderPanel(); ctx.layoutAll();
  }
  function toggleVisible(key) {
    const l = layerOf(key);
    if (!l) return;
    l.visible = !l.visible;
    if (!l.visible && (sel?.layer === key || sel?.a?.layer === key)) clear();
    changed(); renderPanel();
    if (l.created) pages.forEach((p) => ctx.drawAnnots(p)); else ctx.layoutAll();
  }
  function moveLayer(key, dir) {
    const i = layers.findIndex((l) => l.key === key);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= layers.length) return;
    [layers[i], layers[j]] = [layers[j], layers[i]];
    pages.forEach((p) => { stack(p); ctx.drawAnnots(p); });
    changed(); renderPanel();
  }
  /** Lifts a layer's content on a page so it can be moved and resized */
  async function liftLayer(key, p = ctx.currentPageModel()) {
    const l = layerOf(key);
    if (!l || !p) return;
    if (l.created) {
      const items = p.annots.filter((a) => a.layer === key && isMovable(a));
      if (!items.length) { toast('Nothing on this layer on this page'); return; }
      select({ p, a: items[items.length - 1] });
      return;
    }
    if (l.src !== p.src) { toast('This layer belongs to another document in this file'); return; }
    if (!l.visible) { l.visible = true; renderPanel(); }
    p.layerMoves ||= {};
    if (!p.layerMoves[key]) p.layerMoves[key] = [...ID];
    await ctx.refreshPage(p);
    const box = overlayBox(p, key, rendered.get(p.key));
    if (!box) { toast(`"${l.name}" has nothing on page ${pages.indexOf(p) + 1}`); delete p.layerMoves[key]; await ctx.refreshPage(p); return; }
    select({ p, layer: key });
    toast(`Drag "${l.name}" to move it, or a handle to resize`);
  }
  async function resetLayer(p, key) {
    if (!p.layerMoves?.[key]) return;
    undoStack.push({ type: 'layermove', p, key, before: p.layerMoves[key] });
    delete p.layerMoves[key];
    if (sel?.layer === key) clear();
    changed();
    await ctx.refreshPage(p);
    renderPanel();
  }

  // ---------------------------------------------------------------- panel
  const panel = document.createElement('div');
  panel.className = 'pdf-layers';
  panel.hidden = true;
  sideEl.appendChild(panel);
  function showPanel() { sideEl.querySelector('[data-side="layers"]')?.click(); }
  function renderPanel() {
    if (panel.hidden) return;
    const p = ctx.currentPageModel();
    const n = pages.indexOf(p) + 1;
    const itemRows = (key) => (p ? p.annots.map((a, i) => [a, i]).filter(([a]) => (a.layer || null) === key && a.type !== 'ocrText').reverse() : [])
      .map(([a, i]) => `<button class="ly-obj${sel?.a === a ? ' on' : ''}" data-obj="${i}" ${isMovable(a) ? '' : 'disabled'}>${icon(ICON[a.type] || 'rect', 13)}<span>${esc(describe(a))}</span></button>`).join('');
    const row = (l) => {
      const moved = p?.layerMoves?.[l.key] && !isId(p.layerMoves[l.key]);
      return `<div class="ly-layer${active === l.key ? ' active' : ''}${l.visible ? '' : ' off'}${sel?.layer === l.key ? ' sel' : ''}" data-key="${esc(l.key)}">
        <div class="ly-row">
          <button class="ly-eye" data-eye title="${l.visible ? 'Hide' : 'Show'}">${icon(l.visible ? 'view' : 'eye-off', 15)}</button>
          <span class="ly-name" title="Click to put new items on this layer. Double-click to rename.">${esc(l.name)}</span>
          <span>${l.created ? '<span class="ly-tag new">new</span>' : moved ? '<span class="ly-dot" title="Moved or resized on this page"></span>' : ''}</span>
          <button class="ly-move" data-lift title="Move or resize this layer on page ${n}">${icon('move', 14)}</button>
          <button class="ly-more" data-more title="More">${icon('more', 15)}</button>
        </div>
        <div class="ly-items">${itemRows(l.key)}</div>
      </div>`;
    };
    const live = layers.filter((l) => !l.deleted);
    panel.innerHTML = `
      <div class="ly-head"><span>${live.length ? `${live.length} layer${live.length === 1 ? '' : 's'}` : 'No layers yet'}</span><button class="chip" data-new>${icon('plus', 13)}New layer</button></div>
      <div class="ly-list">${live.map(row).join('')}
        <div class="ly-layer page${active === null ? ' active' : ''}" data-key="">
          <div class="ly-row"><span class="ly-eye ghost">${icon('file-plus', 14)}</span><span class="ly-name" title="Items that are not on a layer">Page (no layer)</span></div>
          <div class="ly-items">${itemRows(null)}</div>
        </div>
      </div>
      <p class="ly-hint">${live.length ? `New items go on the highlighted layer. Showing items on page ${n}.` : 'Layers keep parts of a page apart, so they can be shown, hidden and moved on their own. PDFs made with layers in other apps show them here.'}</p>`;
  }
  panel.addEventListener('click', (e) => {
    if (e.target.closest('[data-new]')) { newLayer(); return; }
    const obj = e.target.closest('[data-obj]');
    if (obj) {
      const p = ctx.currentPageModel();
      const a = p?.annots[Number(obj.dataset.obj)];
      if (a && isMovable(a)) select({ p, a });
      return;
    }
    const box = e.target.closest('.ly-layer');
    if (!box) return;
    const key = box.dataset.key || null;
    if (e.target.closest('[data-eye]')) { toggleVisible(key); return; }
    if (e.target.closest('[data-lift]')) { liftLayer(key); return; }
    if (e.target.closest('[data-more]')) {
      const l = layerOf(key);
      const p = ctx.currentPageModel();
      openMenu(e.target.closest('[data-more]'), [
        { label: 'Rename', icon: 'edit', action: () => rename(key) },
        { label: 'Put new items here', icon: 'check', action: () => { active = key; renderPanel(); } },
        { label: `Move or resize on page ${pages.indexOf(p) + 1}`, icon: 'move', action: () => liftLayer(key) },
        ...(!l.created && p?.layerMoves?.[key] ? [{ label: 'Put back where it was', icon: 'undo', action: () => resetLayer(p, key) }] : []),
        ...(l.created ? ['sep', { label: 'Move up', icon: 'up', action: () => moveLayer(key, -1) }, { label: 'Move down', icon: 'down', action: () => moveLayer(key, 1) }] : []),
        'sep',
        { label: 'Delete layer', icon: 'trash', action: () => deleteLayer(key) },
      ]);
      return;
    }
    if (e.target.closest('.ly-row')) { active = key; renderPanel(); }
  });
  panel.addEventListener('dblclick', (e) => { const k = e.target.closest('.ly-name') && e.target.closest('.ly-layer')?.dataset.key; if (k) rename(k); });

  // ---------------------------------------------------------------- saving
  function saveSpec() {
    if (!layers.length) return null;
    return {
      order: layers.filter((l) => !l.deleted).map((l) => l.key),
      existing: layers.filter((l) => !l.created).map((l) => ({ key: l.key, src: l.src, id: l.id, name: l.name, origName: l.origName, visible: l.visible, deleted: l.deleted })),
      created: layers.filter((l) => l.created).map((l) => ({ key: l.key, name: l.name, visible: l.visible })),
    };
  }

  return {
    get selected() { return sel; },
    loadSource, configFor, liftedKeys, renderOverlays, isVisible, drawFrame, pointerDown, onKey, undoStep, afterUndo,
    remove, order, duplicate, setSize, moveToLayer, newLayer, renderPanel, saveSpec, select, clear, stack,
    get active() { return active && isVisible(active) ? active : null; },
    get selection() { return sel; },
    layerChoices: () => [{ label: 'Page (no layer)', icon: 'file-plus', checked: !sel?.a?.layer, action: () => moveToLayer(null) }, ...layers.filter((l) => !l.deleted).map((l) => ({ label: l.name, checked: sel?.a?.layer === l.key, action: () => moveToLayer(l.key) }))],
    panel,
  };
}
