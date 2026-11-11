// Sparklines on the grid: each one is a small canvas floated exactly over its
// cell. They are redrawn when the data changes and re-placed when rows or
// columns move or change size.
import { createElement, useEffect, useRef } from 'react';
import { parseRef, quoteSheet } from '../lib/xlsx-charts.js';
import { rangeToA1 } from '../lib/xlsx-features.js';
import { drawSparkCanvas } from '../lib/sparklines.js';

const ROW_HEADER_W = 46;
const COL_HEADER_H = 20;
const DRAWING = 'SHEET_DRAWING_PLUGIN';
const MUTATION = 2;
const LAYOUT = /set-worksheet-(col-width|row-height|row-auto-height|row-is-auto-height)|insert-(row|col)|remove-(row|col)|move-(rows|cols)|set-(row|col)-(hidden|visible)|set-worksheet-(row|col)-hidden|set-worksheet-(row|col)-visible|set-worksheet-(row|col)s?-(hidden|visible)/;

export function createSparkLayer({ univerAPI, workbook, hooks = {} }) {
  const bySheet = new Map(); // sheetId -> [sp]
  const placed = new Map(); // floatId -> { sid, sp, el }
  const sheetById = (sid) => workbook.getSheets().find((s) => s.getSheetId() === sid);

  const valuesOf = (sid, sp) => {
    const p = parseRef(sp.ref);
    if (!p) return [];
    const ws = p.sheet ? workbook.getSheetByName(p.sheet) : sheetById(sid);
    if (!ws) return [];
    const r = ws.getRange(rangeToA1(p.range));
    return (r.getRawValues?.() || r.getValues()).flat().map((v) => (typeof v === 'number' ? v : v === '' || v === null || v === undefined || !Number.isFinite(Number(v)) ? null : Number(v)));
  };

  function paint(floatId) {
    const item = placed.get(floatId);
    const el = item?.el;
    if (!el) return;
    const canvas = el.querySelector('canvas');
    const w = el.clientWidth; const h = el.clientHeight;
    if (!w || !h) return;
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.round(w * dpr); canvas.height = Math.round(h * dpr);
    canvas.style.width = `${w}px`; canvas.style.height = `${h}px`;
    const ctx = canvas.getContext('2d');
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);
    drawSparkCanvas(ctx, item.sp, valuesOf(item.sid, item.sp), 0, 0, w, h, { dpr });
  }

  function Host(props) {
    const ref = useRef(null);
    const id = props?.data?.sparkId;
    useEffect(() => {
      const el = ref.current;
      if (!id || !el) return undefined;
      const item = placed.get(id);
      if (item) item.el = el;
      // Clicks go through to the cell underneath
      for (let n = el, k = 0; n && k < 4; n = n.parentElement, k++) n.style.pointerEvents = 'none';
      const ro = new ResizeObserver(() => paint(id));
      ro.observe(el);
      paint(id);
      return () => { ro.disconnect(); const it = placed.get(id); if (it?.el === el) it.el = null; };
    }, [id]);
    return createElement('div', { ref, className: 'k-spark' }, createElement('canvas'));
  }
  const KEY = `kertas-spark-${Math.random().toString(36).slice(2, 8)}`;
  univerAPI.registerComponent(KEY, Host);

  let seq = 0;
  function place(sid, sp) {
    const ws = sheetById(sid);
    if (!ws) return;
    const core = ws.getSheet?.();
    const rowShown = (i) => { try { return core ? core.getRowRawVisible(i) : true; } catch { return true; } };
    const colShown = (i) => { try { return core ? core.getColVisible(i) : true; } catch { return true; } };
    if (!rowShown(sp.r) || !colShown(sp.c)) return; // in a hidden row or column
    let x = ROW_HEADER_W; let y = COL_HEADER_H;
    for (let i = 0; i < sp.c; i++) if (colShown(i)) x += ws.getColumnWidth(i);
    for (let i = 0; i < sp.r; i++) if (rowShown(i)) y += ws.getRowHeight(i);
    const w = ws.getColumnWidth(sp.c); const h = ws.getRowHeight(sp.r);
    if (w < 4 || h < 4) return;
    const id = `spark-${++seq}-${Math.random().toString(36).slice(2, 6)}`;
    placed.set(id, { sid, sp, el: null });
    try {
      ws.addFloatDomToPosition({ componentKey: KEY, initPosition: { startX: x + 1, endX: x + w - 1, startY: y + 1, endY: y + h - 1 }, data: { sparkId: id }, allowTransform: false }, id);
    } catch (e) { placed.delete(id); console.warn('sparkline', e); }
  }
  function clearSheet(sid) {
    const ws = sheetById(sid);
    for (const [id, it] of [...placed]) {
      if (it.sid !== sid) continue;
      try { ws?.removeFloatDom?.(id); } catch { /* gone */ }
      placed.delete(id);
    }
  }
  function layout(sid) {
    clearSheet(sid);
    for (const sp of bySheet.get(sid) || []) place(sid, sp);
  }

  let paintTimer = 0; let layoutTimer = 0;
  const sub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => {
    if (!bySheet.size) return;
    if (LAYOUT.test(e.id) || /set-worksheet-active/.test(e.id)) {
      clearTimeout(layoutTimer);
      layoutTimer = setTimeout(() => { for (const sid of bySheet.keys()) layout(sid); }, 120);
      return;
    }
    if (e.type !== MUTATION) return;
    clearTimeout(paintTimer);
    paintTimer = setTimeout(() => { for (const id of placed.keys()) paint(id); }, 200);
  });

  return {
    load(map) {
      for (const [sid, list] of Object.entries(map || {})) { bySheet.set(sid, [...list]); layout(sid); }
    },
    /** Adds sparklines (replacing any in the same cells) */
    add(sid, list) {
      const cur = (bySheet.get(sid) || []).filter((s) => !list.some((n) => n.r === s.r && n.c === s.c));
      bySheet.set(sid, [...cur, ...list]);
      layout(sid);
      hooks.onChange?.();
    },
    /** Removes sparklines inside a range; returns how many */
    removeIn(sid, g) {
      const cur = bySheet.get(sid) || [];
      const keep = cur.filter((s) => !(s.r >= g.startRow && s.r <= g.endRow && s.c >= g.startColumn && s.c <= g.endColumn));
      if (keep.length === cur.length) return 0;
      bySheet.set(sid, keep);
      layout(sid);
      hooks.onChange?.();
      return cur.length - keep.length;
    },
    /** Restyles sparklines inside a range */
    styleIn(sid, g, patch) {
      const cur = bySheet.get(sid) || [];
      let n = 0;
      bySheet.set(sid, cur.map((s) => (s.r >= g.startRow && s.r <= g.endRow && s.c >= g.startColumn && s.c <= g.endColumn ? (n++, { ...s, ...patch }) : s)));
      if (n) { layout(sid); hooks.onChange?.(); }
      return n;
    },
    /** Rows or columns inserted (delta > 0) or removed (delta < 0) at an index */
    shift(sid, axis, at, delta) {
      const sheetName = sheetById(sid)?.getSheetName();
      let touched = false;
      for (const [id, list] of bySheet) {
        const next = [];
        for (const sp of list) {
          let { r, c, ref } = sp;
          if (id === sid) {
            const pos = axis === 'rows' ? r : c;
            if (delta < 0 && pos >= at && pos < at - delta) { touched = true; continue; }
            if (pos >= at) { if (axis === 'rows') r += delta; else c += delta; touched = true; }
          }
          // Data ranges on the changed sheet move too
          const p = parseRef(ref);
          if (p && (p.sheet || sheetName) === sheetName && sheetName) {
            const g = { ...p.range };
            const [a, b] = axis === 'rows' ? ['startRow', 'endRow'] : ['startColumn', 'endColumn'];
            if (g[a] >= at) g[a] = Math.max(at, g[a] + delta);
            if (g[b] >= at) g[b] = Math.max(g[a], g[b] + delta);
            const moved = `${quoteSheet(p.sheet || sheetName)}!${rangeToA1(g)}`;
            if (moved !== ref) { ref = moved; touched = true; }
          }
          next.push({ ...sp, r, c, ref });
        }
        bySheet.set(id, next);
      }
      if (touched) { for (const id of bySheet.keys()) layout(id); }
    },
    count: (sid) => (sid ? bySheet.get(sid)?.length || 0 : [...bySheet.values()].reduce((a, l) => a + l.length, 0)),
    all() { return Object.fromEntries([...bySheet].filter(([, l]) => l.length).map(([sid, l]) => [sid, l.map((s) => ({ ...s }))])); },
    stripFromSnapshot(snapshot) {
      const r = snapshot?.resources?.find((x) => x.name === DRAWING);
      if (!r) return snapshot;
      try {
        const d = JSON.parse(r.data);
        for (const s of Object.values(d)) {
          for (const [k, v] of Object.entries(s.data || {})) if (v.componentKey === KEY) { delete s.data[k]; s.order = (s.order || []).filter((x) => x !== k); }
        }
        r.data = JSON.stringify(d);
      } catch { /* leave as is */ }
      return snapshot;
    },
    dispose() {
      clearTimeout(paintTimer); clearTimeout(layoutTimer);
      try { sub?.dispose?.(); } catch { /* noop */ }
      placed.clear(); bySheet.clear();
    },
  };
}
