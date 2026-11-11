// Live charts for Sheets. Each chart is a floating card over the grid (a
// Univer "float DOM"), drawn with Chart.js from the cells it points at and
// redrawn when those cells change. Saved into .xlsx as real Excel charts.
import { createElement, useEffect, useRef } from 'react';
import { Chart, registerables } from 'chart.js';
import { PALETTE, resolveChart, parseRef, chartId } from '../lib/xlsx-charts.js';
import { rangeToA1 } from '../lib/xlsx-features.js';
import { icon } from '../lib/icons.js';

Chart.register(...registerables);

const ROW_HEADER_W = 46;
const COL_HEADER_H = 20;
const MUTATION = 2;
const DRAWING = 'SHEET_DRAWING_PLUGIN';
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function alpha(hex, a) {
  const h = hex.replace('#', '');
  const n = parseInt(h, 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

/** Chart.js configuration for one of our models */
export function chartJsConfig(model, data, dark) {
  const text = dark ? '#cbd5e1' : '#475569';
  const grid = dark ? 'rgba(148,163,184,0.16)' : 'rgba(15,23,42,0.07)';
  const type = model.type;
  const round = type === 'pie' || type === 'doughnut';
  const datasets = data.datasets.map((d, i) => {
    const base = { label: d.name, data: d.values };
    if (round) {
      return { ...base, backgroundColor: d.values.map((_, k) => PALETTE[k % PALETTE.length]), borderColor: dark ? '#1e293b' : '#ffffff', borderWidth: 2, hoverOffset: 6 };
    }
    if (type === 'scatter') {
      const xs = data.labels.map(Number);
      return { ...base, data: d.values.map((y, k) => ({ x: Number.isFinite(xs[k]) ? xs[k] : k + 1, y })), backgroundColor: d.color, pointRadius: 5, pointHoverRadius: 7 };
    }
    if (type === 'line' || type === 'area') {
      return {
        ...base, borderColor: d.color, backgroundColor: type === 'area' ? alpha(d.color, model.stacked ? 0.75 : 0.28) : d.color,
        fill: type === 'area' ? (model.stacked && i > 0 ? '-1' : 'origin') : false,
        cubicInterpolationMode: 'monotone', tension: 0.3, pointRadius: type === 'line' ? 3 : 0, pointHoverRadius: 5, borderWidth: 2.5,
      };
    }
    return { ...base, backgroundColor: d.color, borderRadius: 5, borderSkipped: false, maxBarThickness: 46 };
  });
  const legendPos = model.legend === 'none' ? null : model.legend || 'bottom';
  const scales = round ? {} : {
    x: { stacked: !!model.stacked && type !== 'line', grid: { display: type === 'bar' || type === 'scatter', color: grid }, ticks: { color: text }, border: { display: false }, ...(type === 'scatter' ? { type: 'linear' } : {}) },
    y: { stacked: !!model.stacked && type !== 'line', grid: { display: type !== 'bar', color: grid }, ticks: { color: text }, border: { display: false }, beginAtZero: true },
  };
  return {
    type: type === 'column' || type === 'bar' ? 'bar' : type === 'area' ? 'line' : type,
    data: { labels: data.labels, datasets },
    options: {
      responsive: true, maintainAspectRatio: false, animation: { duration: 350 },
      indexAxis: type === 'bar' ? 'y' : 'x',
      layout: { padding: { top: 4, right: 10, bottom: 4, left: 6 } },
      plugins: {
        title: { display: !!model.title, text: model.title, color: dark ? '#f1f5f9' : '#0f172a', font: { size: 15, weight: '600', family: 'Inter Variable, Inter, system-ui, sans-serif' }, padding: { bottom: 10 } },
        legend: { display: !!legendPos, position: legendPos || 'bottom', labels: { color: text, usePointStyle: true, pointStyle: 'circle', boxWidth: 8, boxHeight: 8, padding: 14, font: { family: 'Inter Variable, Inter, system-ui, sans-serif' } } },
        tooltip: { backgroundColor: dark ? '#0f172a' : 'rgba(15,23,42,0.92)', padding: 10, cornerRadius: 8, titleFont: { weight: '600' } },
      },
      ...(round ? { cutout: type === 'doughnut' ? '58%' : 0 } : { scales }),
    },
  };
}

/**
 * Manages the charts of one workbook.
 * hooks: { onEdit(model), onChange() }  (onEdit opens the chart dialog)
 */
export function createChartLayer({ univerAPI, workbook, dark = false, hooks = {} }) {
  const charts = new Map(); // id -> { model, sheetId, chart, el, floatId }
  const mounts = new Map(); // id -> host element waiting for its chart
  let isDark = dark;

  const sheetById = (sid) => workbook.getSheets().find((s) => s.getSheetId() === sid);
  const sheetName = (sid) => sheetById(sid)?.getSheetName();

  /** Reads a reference from the live workbook */
  const reader = (fallbackSheet) => (ref) => {
    const p = parseRef(ref);
    if (!p) return [];
    const ws = p.sheet ? workbook.getSheetByName(p.sheet) : workbook.getSheetByName(fallbackSheet);
    if (!ws) return [];
    const range = ws.getRange(rangeToA1(p.range));
    let rows;
    try { rows = range.getDisplayValues?.() || range.getValues(); } catch { rows = range.getValues(); }
    // getValues gives what the cell shows ("38,000"); the raw values are the numbers
    const raw = range.getRawValues?.() || range.getValues();
    // Numbers come from the raw values, text (like dates) from what the cell shows
    return rows.flatMap((r, i) => r.map((v, j) => {
      const x = raw[i]?.[j];
      return typeof x === 'number' ? x : (v ?? x ?? '');
    }));
  };

  function draw(id) {
    const c = charts.get(id);
    const el = mounts.get(id);
    if (!c || !el) return;
    const data = resolveChart(c.model, reader(sheetName(c.sheetId)));
    const cfg = chartJsConfig(c.model, data, isDark);
    el.classList.toggle('dark', isDark);
    if (c.chart && c.chart.config.type === cfg.type && !!c.chart.options.cutout === !!cfg.options.cutout) {
      c.chart.data = cfg.data;
      c.chart.options = cfg.options;
      c.chart.update();
      return;
    }
    c.chart?.destroy();
    const canvas = el.querySelector('canvas');
    c.chart = new Chart(canvas, cfg);
  }

  function attach(id, el) {
    mounts.set(id, el);
    el.innerHTML = `<div class="k-chart-tools">
        <button data-act="edit" title="Edit chart">${icon('edit', 14)}</button>
        <button data-act="delete" title="Delete chart">${icon('trash', 14)}</button>
      </div><div class="k-chart-canvas"><canvas></canvas></div>`;
    el.addEventListener('pointerdown', (e) => { if (e.target.closest('.k-chart-tools')) e.stopPropagation(); }, true);
    el.addEventListener('click', (e) => {
      const act = e.target.closest('[data-act]')?.dataset.act;
      if (act === 'edit') hooks.onEdit?.(charts.get(id)?.model);
      if (act === 'delete') remove(id);
    });
    draw(id);
    return () => {
      const c = charts.get(id);
      c?.chart?.destroy();
      if (c) c.chart = null;
      if (mounts.get(id) === el) mounts.delete(id);
    };
  }

  // React host registered with Univer; the chart itself is plain DOM + Chart.js
  function ChartHost(props) {
    const ref = useRef(null);
    const id = props?.data?.chartId;
    useEffect(() => (id && ref.current ? attach(id, ref.current) : undefined), [id]);
    return createElement('div', { ref, className: `k-chart${isDark ? ' dark' : ''}` });
  }
  const KEY = `kertas-chart-${Math.random().toString(36).slice(2, 8)}`;
  univerAPI.registerComponent(KEY, ChartHost);

  function place(sid, model) {
    const ws = sheetById(sid);
    if (!ws) return null;
    const a = model.anchor;
    let x = ROW_HEADER_W;
    let y = COL_HEADER_H;
    for (let i = 0; i < a.col; i++) x += ws.getColumnWidth(i);
    for (let i = 0; i < a.row; i++) y += ws.getRowHeight(i);
    x += a.colOff || 0;
    y += a.rowOff || 0;
    const res = ws.addFloatDomToPosition({
      componentKey: KEY,
      initPosition: { startX: x, endX: x + a.width, startY: y, endY: y + a.height },
      data: { chartId: model.id },
      allowTransform: true,
    }, model.id);
    return res;
  }

  function add(sid, model, { quiet = false } = {}) {
    const m = { ...model, id: model.id || chartId() };
    const res = place(sid, m);
    if (!res) return null;
    charts.set(m.id, { model: m, sheetId: sid, floatId: res.id, dispose: res.dispose });
    if (!quiet) hooks.onChange?.();
    return m;
  }

  function remove(id) {
    const c = charts.get(id);
    if (!c) return;
    c.chart?.destroy();
    try { sheetById(c.sheetId)?.removeFloatDom?.(c.floatId); } catch { /* already gone */ }
    try { c.dispose?.(); } catch { /* noop */ }
    charts.delete(id);
    mounts.delete(id);
    hooks.onChange?.();
  }

  function update(id, patch) {
    const c = charts.get(id);
    if (!c) return;
    c.model = { ...c.model, ...patch, id };
    draw(id);
    hooks.onChange?.();
  }

  // Redraw when cells change (debounced)
  let timer = 0;
  const sub = univerAPI.addEvent(univerAPI.Event.CommandExecuted, (e) => {
    if (e.type !== MUTATION || !charts.size) return;
    clearTimeout(timer);
    timer = setTimeout(() => { for (const id of charts.keys()) draw(id); }, 250);
  });

  return {
    load(bySheet) {
      for (const [sid, list] of Object.entries(bySheet || {})) for (const m of list) add(sid, m, { quiet: true });
    },
    add,
    update,
    remove,
    has: (id) => charts.has(id),
    count: () => charts.size,
    /** Chart models with positions read back from the grid, keyed by sheet id */
    models(snapshot) {
      let drawings = {};
      try { drawings = JSON.parse(snapshot?.resources?.find((r) => r.name === DRAWING)?.data || '{}'); } catch { drawings = {}; }
      const out = {};
      for (const [id, c] of charts) {
        const d = drawings[c.sheetId]?.data?.[c.floatId];
        const from = d?.sheetTransform?.from;
        const anchor = from && d.transform
          ? { col: from.column, colOff: Math.round(from.columnOffset), row: from.row, rowOff: Math.round(from.rowOffset), width: Math.round(d.transform.width), height: Math.round(d.transform.height) }
          : c.model.anchor;
        c.model.anchor = anchor;
        (out[c.sheetId] ||= []).push({ ...c.model, id, anchor });
      }
      return out;
    },
    /** Removes our float DOMs from a snapshot (they are saved as charts instead) */
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
    setTheme(d) { isDark = d; for (const id of charts.keys()) draw(id); },
    /** A snapshot-based description for the PDF exporter */
    dispose() {
      clearTimeout(timer);
      try { sub?.dispose?.(); } catch { /* noop */ }
      for (const c of charts.values()) c.chart?.destroy();
      charts.clear();
    },
    escHtml: esc,
  };
}

/** Draws a chart off screen (light theme, 2x) and returns PNG bytes, for printing */
export async function renderChartPng(model, read) {
  const w = Math.max(80, Math.round(model.anchor?.width || 480));
  const h = Math.max(60, Math.round(model.anchor?.height || 300));
  const box = document.createElement('div');
  box.style.cssText = `position:fixed;left:-10000px;top:0;width:${w}px;height:${h}px;background:#fff`;
  const canvas = document.createElement('canvas');
  canvas.width = w;
  canvas.height = h;
  canvas.style.cssText = `width:${w}px;height:${h}px`;
  box.appendChild(canvas);
  document.body.appendChild(box);
  try {
    const cfg = chartJsConfig(model, resolveChart(model, read), false);
    cfg.options = { ...cfg.options, responsive: false, animation: false, devicePixelRatio: 2 };
    // White card behind the chart so it prints like the one on screen
    cfg.plugins = [{ id: 'bg', beforeDraw: (c) => { const { ctx } = c; ctx.save(); ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, c.width, c.height); ctx.restore(); } }];
    const chart = new Chart(canvas, cfg);
    const url = chart.toBase64Image('image/png');
    chart.destroy();
    const bin = atob(url.split(',')[1]);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  } catch (e) {
    console.error(e);
    return null;
  } finally {
    box.remove();
  }
}
