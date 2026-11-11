// Conditional formatting, evaluated outside the grid (for Export as PDF).
// Takes Univer's rule list for a sheet and the cell values in the snapshot,
// and answers, per cell: extra style, a data bar, or an icon. Pure module.
//
// Rules come first-to-last in priority order. Highlight rules merge their styles
// with earlier (higher priority) rules winning per property, as Excel does.
// Formula rules are skipped: their results live in the grid's calculation engine.

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const blank = (v) => v === '' || v === null || v === undefined;
const isErr = (v) => typeof v === 'string' && /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!)$/.test(v);

function cellValue(sh, r, c) {
  const cd = sh.cellData?.[r]?.[c];
  if (!cd) return null;
  const v = cd.v;
  if (v === undefined || v === null) return null;
  if (cd.t === 2 || (typeof v === 'number' && cd.t !== 1)) return Number(v);
  if (typeof v === 'string' && cd.t !== 1 && v.trim() !== '' && Number.isFinite(Number(v))) return Number(v);
  return v;
}
function inRanges(ranges, r, c) {
  return ranges.some((g) => r >= g.startRow && r <= g.endRow && c >= g.startColumn && c <= g.endColumn);
}
function valuesIn(sh, ranges) {
  const out = [];
  for (const g of ranges) {
    const r1 = Math.min(g.endRow, g.startRow + 20000);
    for (let r = g.startRow; r <= r1; r++) for (let c = g.startColumn; c <= g.endColumn; c++) out.push(cellValue(sh, r, c));
  }
  return out;
}
const numbers = (list) => list.filter(isNum).sort((a, b) => a - b);
function percentile(sorted, p) {
  if (!sorted.length) return 0;
  const k = (sorted.length - 1) * Math.max(0, Math.min(1, p));
  const lo = Math.floor(k); const hi = Math.ceil(k);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (k - lo);
}
/** A colour scale / data bar threshold */
function threshold(v = {}, sorted, fallback) {
  const val = Number(String(v.value ?? '').replace(/^=/, ''));
  switch (v.type) {
    case 'min': return sorted[0] ?? 0;
    case 'max': return sorted.at(-1) ?? 0;
    case 'percent': { const lo = sorted[0] ?? 0; const hi = sorted.at(-1) ?? 0; return lo + (hi - lo) * ((Number.isFinite(val) ? val : 0) / 100); }
    case 'percentile': return percentile(sorted, (Number.isFinite(val) ? val : 50) / 100);
    case 'num': case 'formula': return Number.isFinite(val) ? val : fallback;
    default: return fallback;
  }
}
const hex = (s) => { const m = String(s || '').replace('#', ''); const n = parseInt(m.length === 3 ? m.replace(/./g, '$&$&') : m.slice(0, 6), 16); return Number.isFinite(n) ? [(n >> 16) & 255, (n >> 8) & 255, n & 255] : [255, 255, 255]; };
const toHex = ([r, g, b]) => `#${[r, g, b].map((x) => Math.round(x).toString(16).padStart(2, '0')).join('')}`;
const mix = (a, b, t) => toHex(hex(a).map((x, i) => x + (hex(b)[i] - x) * t));

function dayOf(serial) { return Math.floor(serial); }
function todaySerial(now) { return Math.floor((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - EXCEL_EPOCH) / 86400000); }
function timeMatch(op, v, now) {
  if (!isNum(v)) return false;
  const d = dayOf(v); const t = todaySerial(now);
  const date = new Date(EXCEL_EPOCH + d * 86400000);
  const ym = (dt) => dt.getUTCFullYear() * 12 + dt.getUTCMonth();
  const nowDate = new Date(EXCEL_EPOCH + t * 86400000);
  const weekStart = t - nowDate.getUTCDay();
  switch (op) {
    case 'today': return d === t;
    case 'yesterday': return d === t - 1;
    case 'tomorrow': return d === t + 1;
    case 'last7Days': return d <= t && d > t - 7;
    case 'thisWeek': return d >= weekStart && d < weekStart + 7;
    case 'lastWeek': return d >= weekStart - 7 && d < weekStart;
    case 'nextWeek': return d >= weekStart + 7 && d < weekStart + 14;
    case 'thisMonth': return ym(date) === ym(nowDate);
    case 'lastMonth': return ym(date) === ym(nowDate) - 1;
    case 'nextMonth': return ym(date) === ym(nowDate) + 1;
    default: return false;
  }
}

function highlightMatch(rule, v, ctx) {
  const op = rule.operator;
  const val = rule.value;
  switch (rule.subType) {
    case 'number': {
      if (!isNum(v)) return false;
      const a = Number(Array.isArray(val) ? val[0] : val); const b = Number(Array.isArray(val) ? val[1] : val);
      const [lo, hi] = a <= b ? [a, b] : [b, a];
      switch (op) {
        case 'between': return v >= lo && v <= hi;
        case 'notBetween': return v < lo || v > hi;
        case 'equal': return v === a;
        case 'notEqual': return v !== a;
        case 'greaterThan': return v > a;
        case 'greaterThanOrEqual': return v >= a;
        case 'lessThan': return v < a;
        case 'lessThanOrEqual': return v <= a;
        default: return false;
      }
    }
    case 'text': {
      const s = blank(v) ? '' : String(v).toLowerCase();
      const t = String(val ?? '').toLowerCase();
      switch (op) {
        case 'containsText': return s.includes(t);
        case 'notContainsText': return !s.includes(t);
        case 'beginsWith': return s.startsWith(t);
        case 'endsWith': return s.endsWith(t);
        case 'equal': return s === t;
        case 'notEqual': return s !== t;
        case 'containsBlanks': return s.trim() === '';
        case 'notContainsBlanks': return s.trim() !== '';
        case 'containsErrors': return isErr(v);
        case 'notContainsErrors': return !isErr(v);
        default: return false;
      }
    }
    case 'timePeriod': return timeMatch(op, v, ctx.now);
    case 'average': {
      if (!isNum(v)) return false;
      const avg = ctx.stats().avg;
      switch (op) {
        case 'greaterThan': return v > avg;
        case 'greaterThanOrEqual': return v >= avg;
        case 'lessThan': return v < avg;
        case 'lessThanOrEqual': return v <= avg;
        default: return false;
      }
    }
    case 'rank': {
      if (!isNum(v)) return false;
      const sorted = ctx.stats().sorted;
      const n = Math.max(1, rule.isPercent ? Math.floor((sorted.length * (Number(val) || 10)) / 100) : Number(val) || 10);
      if (rule.isBottom) return v <= sorted[Math.min(sorted.length - 1, n - 1)];
      return v >= sorted[Math.max(0, sorted.length - n)];
    }
    case 'duplicateValues': case 'uniqueValues': {
      if (blank(v)) return false;
      const count = ctx.stats().counts.get(typeof v === 'string' ? v.toLowerCase() : v) || 0;
      return rule.subType === 'duplicateValues' ? count > 1 : count === 1;
    }
    default: return false; // formula rules need the calculation engine
  }
}

const STYLE_KEYS = ['bg', 'cl', 'bl', 'it', 'ul', 'st'];

/**
 * Builds the evaluator for one sheet.
 * rules: Univer's list for the sheet [{ ranges, rule, stopIfTrue }]
 * Returns (r, c) => null | { style, bar: { ratio, negative, color, gradient, showValue, axis }, icon: { set, index, count }, hideValue }
 */
export function cfEvaluator(sh, rules = [], { now = new Date() } = {}) {
  if (!rules.length) return () => null;
  const prepared = rules.map((x) => {
    let cache = null;
    const stats = () => {
      if (!cache) {
        const all = valuesIn(sh, x.ranges || []);
        const sorted = numbers(all);
        const counts = new Map();
        for (const v of all) if (!blank(v)) { const k = typeof v === 'string' ? v.toLowerCase() : v; counts.set(k, (counts.get(k) || 0) + 1); }
        cache = { sorted, counts, avg: sorted.length ? sorted.reduce((a, b) => a + b, 0) / sorted.length : 0 };
      }
      return cache;
    };
    return { ...x, stats };
  });
  const memo = new Map();
  return (r, c) => {
    const key = `${r}:${c}`;
    if (memo.has(key)) return memo.get(key);
    let out = null;
    const v = cellValue(sh, r, c);
    for (const p of prepared) {
      if (!inRanges(p.ranges || [], r, c)) continue;
      const rule = p.rule || {};
      let hit = false;
      if (rule.type === 'highlightCell') {
        if (highlightMatch(rule, v, { now, stats: p.stats })) {
          hit = true;
          out ||= { style: {} };
          out.style ||= {};
          for (const k of STYLE_KEYS) if (rule.style?.[k] !== undefined && out.style[k] === undefined) out.style[k] = rule.style[k];
        }
      } else if (rule.type === 'colorScale' && isNum(v)) {
        const cfg = [...(rule.config || [])].sort((a, b) => a.index - b.index);
        const { sorted } = p.stats();
        if (cfg.length >= 2 && sorted.length) {
          const stops = cfg.map((s, i) => ({ at: threshold(s.value, sorted, i === 0 ? sorted[0] : sorted.at(-1)), color: s.color }));
          let color = stops[0].color;
          if (v >= stops.at(-1).at) color = stops.at(-1).color;
          else if (v > stops[0].at) {
            for (let i = 0; i < stops.length - 1; i++) {
              const a = stops[i]; const b = stops[i + 1];
              if (v >= a.at && v <= b.at) { color = mix(a.color, b.color, b.at === a.at ? 0 : (v - a.at) / (b.at - a.at)); break; }
            }
          }
          out ||= { style: {} };
          if (out.style.bg === undefined) out.style.bg = { rgb: color };
          hit = true;
        }
      } else if (rule.type === 'dataBar' && isNum(v) && !out?.bar) {
        const cfg = rule.config || {};
        const { sorted } = p.stats();
        const lo = threshold(cfg.min, sorted, Math.min(0, sorted[0] ?? 0));
        const hi = threshold(cfg.max, sorted, sorted.at(-1) ?? 1);
        const span = hi - lo || 1;
        const hasNeg = lo < 0;
        const axis = hasNeg ? Math.max(0, Math.min(1, -lo / span)) : 0;
        const ratio = hasNeg ? Math.max(-1, Math.min(1, v / span)) : Math.max(0, Math.min(1, (v - lo) / span));
        out ||= { style: {} };
        out.bar = { ratio, axis, negative: v < 0, color: v < 0 ? cfg.nativeColor || '#FF0000' : cfg.positiveColor || '#638EC6', gradient: cfg.isGradient !== false };
        if (rule.isShowValue === false) out.hideValue = true;
        hit = true;
      } else if (rule.type === 'iconSet' && isNum(v) && !out?.icon) {
        const cfg = rule.config || [];
        const { sorted } = p.stats();
        let index = cfg.length - 1;
        for (let i = 0; i < cfg.length - 1; i++) {
          const t = threshold(cfg[i].value, sorted, 0);
          const ok = cfg[i].operator === 'greaterThan' ? v > t : v >= t;
          if (ok) { index = i; break; }
        }
        out ||= { style: {} };
        out.icon = { set: cfg[index]?.iconType || cfg[0]?.iconType || '3TrafficLights1', id: Number(cfg[index]?.iconId ?? index), index, count: cfg.length };
        if (rule.isShowValue === false) out.hideValue = true;
        hit = true;
      }
      if (hit && p.stopIfTrue) break;
    }
    memo.set(key, out);
    return out;
  };
}

/** Icon colour and shape for drawing: { shape: 'up'|'down'|'side'|'circle'|'flag'|'star', color } */
export function iconLook(icon) {
  const set = String(icon.set || '');
  const n = icon.count || 3;
  // id 0 is the "best" icon in Univer's order
  const rank = Math.max(0, Math.min(n - 1, icon.id ?? icon.index));
  const palette = n <= 3 ? ['#16A34A', '#EAB308', '#DC2626'] : n === 4 ? ['#16A34A', '#84CC16', '#EAB308', '#DC2626'] : ['#16A34A', '#84CC16', '#EAB308', '#F97316', '#DC2626'];
  const color = palette[rank] || '#64748B';
  if (/Arrow/i.test(set)) return { shape: rank === 0 ? 'up' : rank === n - 1 ? 'down' : 'side', color };
  if (/Flag/i.test(set)) return { shape: 'flag', color };
  if (/Star/i.test(set)) return { shape: 'star', color, fill: 1 - rank / (n - 1 || 1) };
  return { shape: 'circle', color };
}
