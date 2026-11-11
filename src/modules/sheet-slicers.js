// Slicers: a floating panel of buttons, one per value in a column of a table, that filters the
// table by those values. Several slicers on one table work together. The filtering itself is the
// sheet's normal filter, so it is saved as an Excel filter; the slicer panels are kept in a small
// Kertas part of the file (see xlsx-bridge.js) and come back when the workbook is opened again.
import { createElement, useEffect, useRef } from 'react';
import { icon } from '../lib/icons.js';
import { rangeToA1 } from '../lib/xlsx-features.js';

const ROW_HEADER_W = 46;
const COL_HEADER_H = 20;
const MUTATION = 2;
const DRAWING = 'SHEET_DRAWING_PLUGIN';
const MAX_ITEMS = 300;
const BLANK = '\u0000blank';
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
export const slicerId = () => `slicer-${Math.random().toString(36).slice(2, 10)}`;

/** The text the sheet's filter compares a cell by (Univer's extractPureTextFromCell) */
export function filterKey(v) {
  if (v === null || v === undefined || v === '') return BLANK;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return v.toString();
  return String(v).replace(/[\r\n]/g, '');
}

/** Distinct values of one column, in sheet order sorted naturally: [{ key, label }] */
export function distinctItems(raw, shown) {
  const seen = new Map();
  raw.forEach((v, i) => {
    const key = filterKey(v);
    if (!seen.has(key)) seen.set(key, { key, label: key === BLANK ? '(Blank)' : String(shown[i] ?? v), num: typeof v === 'number' ? v : null });
  });
  const list = [...seen.values()];
  const blank = list.filter((x) => x.key === BLANK);
  const rest = list.filter((x) => x.key !== BLANK).sort((a, b) => (a.num !== null && b.num !== null ? a.num - b.num : a.label.localeCompare(b.label, undefined, { numeric: true })));
  return [...rest, ...blank].slice(0, MAX_ITEMS);
}

/**
 * Manages the slicers of one workbook.
 * Model: { id, title, range: { startRow, startColumn, endRow, endColumn } (header row included),
 *          col (index inside the range), selected: [keys] (empty = everything), anchor }
 */
export function createSlicerLayer({ univerAPI, workbook, toast, hooks = {} }) {
  const slicers = new Map(); // id -> { model, sheetId, floatId, dispose }
  const mounts = new Map();
  const sheetById = (sid) => workbook.getSheets().find((s) => s.getSheetId() === sid);

  function columnData(s) {
    const ws = sheetById(s.sheetId);
    const r = s.model.range;
    if (!ws || r.endRow <= r.startRow) return { raw: [], shown: [] };
    const col = r.startColumn + s.model.col;
    const range = ws.getRange(r.startRow + 1, col, r.endRow - r.startRow, 1);
    const raw = (range.getRawValues?.() || range.getValues()).map((x) => x[0]);
    let shown;
    try { shown = (range.getDisplayValues?.() || range.getValues()).map((x) => x[0]); } catch { shown = raw; }
    return { raw, shown };
  }

  // The filter on the slicer's sheet, made over the slicer's table when there is none yet
  function filterFor(s, create) {
    const ws = sheetById(s.sheetId);
    if (!ws) return null;
    let f = ws.getFilter?.();
    const r = s.model.range;
    if (!f && create) {
      f = ws.getRange(r.startRow, r.startColumn, r.endRow - r.startRow + 1, r.endColumn - r.startColumn + 1).createFilter();
    }
    if (!f) return null;
    const fr = f.getRange().getRange();
    const col = r.startColumn + s.model.col;
    if (col < fr.startColumn || col > fr.endColumn || fr.startRow !== r.startRow) {
      if (create) toast?.('This sheet already has a filter on other cells. Clear it (Data, Filter) to use this slicer.', 'error');
      return null;
    }
    return f;
  }

  /** Reads the column's filter back, so the slicer follows changes made in the filter menu */
  function syncSelected(s) {
    const f = filterFor(s, false);
    if (!f) { s.model.selected = []; return; }
    const crit = f.getColumnFilterCriteria(s.model.range.startColumn + s.model.col);
    const list = crit?.filters?.filters || [];
    s.model.selected = [...list, ...(crit?.filters?.blank ? [BLANK] : [])];
  }

  function apply(s) {
    const col = s.model.range.startColumn + s.model.col;
    const sel = s.model.selected;
    const f = filterFor(s, sel.length > 0);
    if (!f) { if (sel.length) s.model.selected = []; return; }
    if (!sel.length) f.removeColumnFilterCriteria(col);
    else {
      const fr = f.getRange().getRange();
      const values = sel.filter((k) => k !== BLANK);
      f.setColumnFilterCriteria(col, { colId: col - fr.startColumn, filters: { filters: values, ...(sel.includes(BLANK) ? { blank: true } : {}) } });
    }
    hooks.onChange?.();
  }

  function draw(id) {
    const s = slicers.get(id);
    const el = mounts.get(id);
    if (!s || !el) return;
    syncSelected(s);
    const { raw, shown } = columnData(s);
    const items = distinctItems(raw, shown);
    const on = new Set(s.model.selected);
    el.querySelector('.k-slicer-title').textContent = s.model.title;
    el.querySelector('[data-act="clear"]').disabled = !on.size;
    el.querySelector('.k-slicer-items').innerHTML = items.length
      ? items.map((it) => `<button type="button" class="k-slicer-item${!on.size || on.has(it.key) ? ' on' : ''}" data-key="${esc(it.key)}" title="${esc(it.label)}">${esc(it.label)}</button>`).join('')
      : '<div class="k-slicer-empty">No values</div>';
  }

  function attach(id, el) {
    mounts.set(id, el);
    el.innerHTML = `<div class="k-slicer-head"><span class="k-slicer-title"></span>
        <button data-act="multi" title="Pick several (or hold Ctrl)">${icon('field-list', 14)}</button>
        <button data-act="clear" title="Clear this slicer's filter">${icon('filter-x', 14)}</button>
        <button data-act="delete" title="Remove slicer">${icon('close', 14)}</button></div>
      <div class="k-slicer-items"></div>`;
    // buttons inside the panel act on click; dragging the header moves the panel
    el.addEventListener('pointerdown', (e) => { if (e.target.closest('button')) e.stopPropagation(); }, true);
    el.addEventListener('click', (e) => {
      const s = slicers.get(id);
      if (!s) return;
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (act === 'delete') { remove(id); return; }
      if (act === 'clear') { s.model.selected = []; apply(s); draw(id); return; }
      if (act === 'multi') { s.multi = !s.multi; el.querySelector('[data-act="multi"]').classList.toggle('on', s.multi); return; }
      const btn = e.target.closest('.k-slicer-item');
      if (!btn) return;
      const key = btn.dataset.key;
      const cur = new Set(s.model.selected);
      if (e.ctrlKey || e.metaKey || s.multi) { if (cur.has(key)) cur.delete(key); else cur.add(key); }
      else if (cur.size === 1 && cur.has(key)) cur.clear();
      else { cur.clear(); cur.add(key); }
      s.model.selected = [...cur];
      apply(s);
      draw(id);
    });
    draw(id);
    return () => { if (mounts.get(id) === el) mounts.delete(id); };
  }

  function SlicerHost(props) {
    const ref = useRef(null);
    const id = props?.data?.slicerId;
    useEffect(() => (id && ref.current ? attach(id, ref.current) : undefined), [id]);
    return createElement('div', { ref, className: 'k-slicer' });
  }
  const KEY = `kertas-slicer-${Math.random().toString(36).slice(2, 8)}`;
  univerAPI.registerComponent(KEY, SlicerHost);

  function place(sid, model) {
    const ws = sheetById(sid);
    if (!ws) return null;
    const a = model.anchor;
    let x = ROW_HEADER_W; let y = COL_HEADER_H;
    for (let i = 0; i < a.col; i++) x += ws.getColumnWidth(i);
    for (let i = 0; i < a.row; i++) y += ws.getRowHeight(i);
    x += a.colOff || 0; y += a.rowOff || 0;
    return ws.addFloatDomToPosition({
      componentKey: KEY,
      initPosition: { startX: x, endX: x + a.width, startY: y, endY: y + a.height },
      data: { slicerId: model.id },
      allowTransform: true,
    }, model.id);
  }

  function add(sid, model, { quiet = false } = {}) {
    const m = { selected: [], ...model, id: model.id || slicerId() };
    const res = place(sid, m);
    if (!res) return null;
    const s = { model: m, sheetId: sid, floatId: res.id, dispose: res.dispose };
    slicers.set(m.id, s);
    // a saved selection is put back into the filter when the file opens
    if (m.selected.length) apply(s);
    if (!quiet) hooks.onChange?.();
    return m;
  }

  function remove(id) {
    const s = slicers.get(id);
    if (!s) return;
    if (s.model.selected.length) { s.model.selected = []; apply(s); }
    try { sheetById(s.sheetId)?.removeFloatDom?.(s.floatId); } catch { /* already gone */ }
    try { s.dispose?.(); } catch { /* noop */ }
    slicers.delete(id);
    mounts.delete(id);
    hooks.onChange?.();
  }

  // Values and filters change with the cells
  let timer = 0;
  const sub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => {
    if (!slicers.size) return;
    if (e.type !== MUTATION && !/filter/i.test(e.id || '')) return;
    clearTimeout(timer);
    timer = setTimeout(() => { for (const id of mounts.keys()) draw(id); }, 120);
  });

  return {
    add, remove,
    get size() { return slicers.size; },
    /** Saved form, with positions read back from the grid: { sheetId: [model] } */
    all(snapshot) {
      let drawings = {};
      try { drawings = JSON.parse(snapshot?.resources?.find((r) => r.name === DRAWING)?.data || '{}'); } catch { drawings = {}; }
      const out = {};
      for (const s of slicers.values()) {
        if (!sheetById(s.sheetId)) continue;
        const d = drawings[s.sheetId]?.data?.[s.floatId];
        const from = d?.sheetTransform?.from;
        if (from && d.transform) s.model.anchor = { col: from.column, colOff: Math.round(from.columnOffset), row: from.row, rowOff: Math.round(from.rowOffset), width: Math.round(d.transform.width), height: Math.round(d.transform.height) };
        (out[s.sheetId] ||= []).push({ ...s.model, range: { ...s.model.range }, selected: [...s.model.selected] });
      }
      return out;
    },
    /** Removes the panels from a snapshot (they are saved in their own part) */
    stripFromSnapshot(snapshot) {
      const r = snapshot?.resources?.find((x) => x.name === DRAWING);
      if (!r) return snapshot;
      try {
        const d = JSON.parse(r.data);
        for (const sh of Object.values(d)) {
          for (const [k, v] of Object.entries(sh.data || {})) if (v.componentKey === KEY) { delete sh.data[k]; sh.order = (sh.order || []).filter((x) => x !== k); }
        }
        r.data = JSON.stringify(d);
      } catch { /* leave as is */ }
      return snapshot;
    },
    load(bySheet) {
      for (const [sid, list] of Object.entries(bySheet || {})) for (const m of list) add(sid, m, { quiet: true });
    },
    rangeLabel: (m) => rangeToA1(m.range),
    dispose() { try { sub?.dispose?.(); } catch { /* noop */ } clearTimeout(timer); },
  };
}
