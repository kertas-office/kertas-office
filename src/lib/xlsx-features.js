// Maps the "extra" spreadsheet features between Excel and Univer:
// filters (with criteria), conditional formatting, data validation,
// hyperlinks, notes and pictures. Used by xlsx-bridge.js. Pure module.
//
// Univer keeps these in snapshot.resources as [{ name, data: JSON string }],
// each keyed by sheet id. ExcelJS covers most of them; filter criteria are
// not written by ExcelJS, so they are patched into the sheet XML with JSZip.

export const RES = {
  filter: 'SHEET_FILTER_PLUGIN',
  cf: 'SHEET_CONDITIONAL_FORMATTING_PLUGIN',
  dv: 'SHEET_DATA_VALIDATION_PLUGIN',
  note: 'SHEET_NOTE_PLUGIN',
  drawing: 'SHEET_DRAWING_PLUGIN',
};

const EMU = 9525; // EMUs per pixel
const ROW_HEADER_W = 46;
const COL_HEADER_H = 20;

let uid = 0;
const rid = (p = '') => `${p}${Math.random().toString(36).slice(2, 8)}${(uid++).toString(36)}`;

// ---------------------------------------------------------------- A1 helpers

export function colName(c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
const colIndex = (letters) => letters.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

export function rangeToA1(r, abs = false) {
  const d = abs ? '$' : '';
  const a = `${d}${colName(r.startColumn)}${d}${r.startRow + 1}`;
  const b = `${d}${colName(r.endColumn)}${d}${r.endRow + 1}`;
  return a === b ? a : `${a}:${b}`;
}

/** 'B2' | '$B$2:$D$9' | 'Sheet1!A1:B2' -> { startRow, startColumn, endRow, endColumn } */
export function parseA1(ref) {
  const clean = String(ref).replace(/^.*!/, '').replace(/\$/g, '');
  const m = clean.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
  if (!m) return null;
  const r = { startRow: +m[2] - 1, startColumn: colIndex(m[1]), endRow: +(m[4] || m[2]) - 1, endColumn: colIndex(m[3] || m[1]), rangeType: 0 };
  return r;
}

const clampRange = (r, sh) => ({
  startRow: Number.isFinite(r.startRow) ? r.startRow : 0,
  startColumn: Number.isFinite(r.startColumn) ? r.startColumn : 0,
  endRow: Number.isFinite(r.endRow) && r.rangeType !== 2 && r.rangeType !== 3 ? r.endRow : (sh.rowCount || 1000) - 1,
  endColumn: Number.isFinite(r.endColumn) && r.rangeType !== 1 && r.rangeType !== 3 ? r.endColumn : (sh.columnCount || 26) - 1,
});

const quoteSheet = (name) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
const xmlEsc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const xmlUnesc = (s) => String(s).replace(/&(amp|lt|gt|quot|apos|#\d+|#x[0-9a-f]+);/gi, (m, e) => {
  const map = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  if (map[e.toLowerCase()]) return map[e.toLowerCase()];
  return String.fromCodePoint(e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
});
const attr = (tag, name) => { const m = tag.match(new RegExp(`\\b${name}="([^"]*)"`)); return m ? xmlUnesc(m[1]) : undefined; };

// ---------------------------------------------------------------- colours

function toHex(c) {
  if (!c) return undefined;
  if (typeof c === 'object') c = c.rgb || c.argb;
  if (!c) return undefined;
  const s = String(c).trim();
  let m = s.match(/^#?([0-9a-f]{6})$/i);
  if (m) return `#${m[1].toUpperCase()}`;
  m = s.match(/^#?[0-9a-f]{2}([0-9a-f]{6})$/i);
  if (m) return `#${m[1].toUpperCase()}`;
  m = s.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (m) return `#${m.slice(1, 4).map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
  return undefined;
}
const argb = (c) => { const h = toHex(c); return h ? { argb: `FF${h.slice(1)}` } : undefined; };

function styleToDxf(st = {}) {
  const out = {};
  const font = {};
  if (st.bl) font.bold = true;
  if (st.it) font.italic = true;
  if (st.ul?.s) font.underline = true;
  if (st.st?.s) font.strike = true;
  const fc = argb(st.cl);
  if (fc) font.color = fc;
  if (Object.keys(font).length) out.font = font;
  const bg = argb(st.bg);
  if (bg) out.fill = { type: 'pattern', pattern: 'solid', fgColor: bg, bgColor: bg };
  return out;
}
function dxfToStyle(d = {}) {
  const st = {};
  const f = d.font || {};
  if (f.bold) st.bl = 1;
  if (f.italic) st.it = 1;
  if (f.underline) st.ul = { s: 1 };
  if (f.strike) st.st = { s: 1 };
  const fc = toHex(f.color);
  if (fc) st.cl = { rgb: fc };
  const fill = d.fill;
  const bg = fill && (toHex(fill.bgColor) || toHex(fill.fgColor));
  if (bg) st.bg = { rgb: bg };
  return st;
}

// ---------------------------------------------------------------- conditional formatting

const num = (v) => (v !== '' && v !== null && v !== undefined && Number.isFinite(Number(v)) ? Number(v) : v);
const qstr = (s) => `"${String(s).replace(/"/g, '""')}"`;
const CMP = { greaterThan: '>', greaterThanOrEqual: '>=', lessThan: '<', lessThanOrEqual: '<=', equal: '=', notEqual: '<>' };

const cfvoOut = (v = {}) => {
  const type = v.type === 'num' ? 'num' : v.type || 'min';
  const o = { type };
  if (type !== 'min' && type !== 'max') o.value = num(String(v.value ?? 0).replace(/^=/, ''));
  return o;
};
const cfvoIn = (v = {}) => {
  const type = v.type || 'min';
  const o = { type };
  if (type !== 'min' && type !== 'max') o.value = type === 'formula' ? `=${v.value}` : num(v.value);
  return o;
};

/** One Univer rule -> one ExcelJS rule (or null when it cannot be expressed) */
function cfRuleOut(rule, first, abs) {
  const tl = rangeToA1({ ...first, endRow: first.startRow, endColumn: first.startColumn });
  const style = styleToDxf(rule.style);
  if (rule.type === 'highlightCell') {
    const op = rule.operator;
    const v = rule.value;
    switch (rule.subType) {
      case 'number': {
        const between = op === 'between' || op === 'notBetween';
        const formulae = between ? [String(num(v?.[0] ?? 0)), String(num(v?.[1] ?? 0))] : [String(num(v ?? 0))];
        return { type: 'cellIs', operator: op, formulae, style };
      }
      case 'text': {
        const t = String(v ?? '');
        if (op === 'containsText' || /Blanks|Errors/.test(op)) return { type: 'containsText', operator: op, text: t, style };
        if (op === 'notContainsText') return { type: 'expression', formulae: [`ISERROR(SEARCH(${qstr(t)},${tl}))`], style };
        if (op === 'beginsWith') return { type: 'expression', formulae: [`LEFT(${tl},${t.length})=${qstr(t)}`], style };
        if (op === 'endsWith') return { type: 'expression', formulae: [`RIGHT(${tl},${t.length})=${qstr(t)}`], style };
        if (op === 'equal' || op === 'notEqual') return { type: 'cellIs', operator: op, formulae: [qstr(t)], style };
        return null;
      }
      case 'timePeriod':
        return { type: 'timePeriod', timePeriod: op, style };
      case 'average':
        return { type: 'aboveAverage', aboveAverage: !/^less/.test(op), style };
      case 'rank':
        return { type: 'top10', rank: Number(v) || 10, percent: !!rule.isPercent, bottom: !!rule.isBottom, style };
      case 'formula':
        return { type: 'expression', formulae: [String(v || '').replace(/^=/, '')], style };
      case 'duplicateValues':
      case 'uniqueValues':
        return { type: 'expression', formulae: [`COUNTIF(${abs},${tl})${rule.subType === 'duplicateValues' ? '>1' : '=1'}`], style };
      default:
        return null;
    }
  }
  if (rule.type === 'dataBar') {
    const c = rule.config || {};
    return { type: 'dataBar', cfvo: [cfvoOut(c.min), cfvoOut(c.max)], color: argb(c.positiveColor) || { argb: 'FF638EC6' }, gradient: !!c.isGradient, showValue: rule.isShowValue !== false };
  }
  if (rule.type === 'colorScale') {
    const cfg = [...(rule.config || [])].sort((a, b) => a.index - b.index);
    return { type: 'colorScale', cfvo: cfg.map((x) => cfvoOut(x.value)), color: cfg.map((x) => argb(x.color) || { argb: 'FFFFFFFF' }) };
  }
  if (rule.type === 'iconSet') {
    // Univer lists thresholds from the top icon down; Excel lists them from the bottom up
    const cfg = rule.config || [];
    const n = cfg.length;
    const cfvo = [{ type: 'percent', value: 0 }];
    for (let i = n - 2; i >= 0; i--) cfvo.push(cfvoOut(cfg[i].value));
    return { type: 'iconSet', iconSet: cfg[0]?.iconType || '3Arrows', cfvo, showValue: rule.isShowValue !== false };
  }
  return null;
}

function cfRuleIn(r, firstTl, absRef) {
  const style = dxfToStyle(r.style);
  const hl = (subType, operator, value, extra = {}) => ({ type: 'highlightCell', subType, operator, value, style, ...extra });
  switch (r.type) {
    case 'cellIs': {
      const f = (r.formulae || []).map((x) => String(x));
      const isNum = f.every((x) => Number.isFinite(Number(x)) && x !== '');
      if (isNum) {
        const between = r.operator === 'between' || r.operator === 'notBetween';
        return hl('number', r.operator, between ? [Number(f[0]), Number(f[1])] : Number(f[0]));
      }
      if ((r.operator === 'equal' || r.operator === 'notEqual') && /^".*"$/.test(f[0])) return hl('text', r.operator, f[0].slice(1, -1).replace(/""/g, '"'));
      if (r.operator === 'between' || r.operator === 'notBetween') {
        const inner = `AND(${firstTl}>=${f[0]},${firstTl}<=${f[1]})`;
        return hl('formula', 'containsText', `=${r.operator === 'between' ? inner : `NOT(${inner})`}`);
      }
      return CMP[r.operator] ? hl('formula', 'containsText', `=${firstTl}${CMP[r.operator]}${f[0]}`) : null;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      if (r.text !== undefined || /Blanks|Errors/.test(r.operator || '')) return hl('text', r.operator || r.type, r.text ?? '');
      return r.formulae?.[0] ? cfRuleIn({ ...r, type: 'expression' }, firstTl, absRef) : null;
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      return hl('text', r.type, '');
    case 'timePeriod':
      return r.timePeriod ? hl('timePeriod', r.timePeriod, '') : null;
    case 'aboveAverage':
      return hl('average', r.aboveAverage === false ? (r.equalAverage ? 'lessThanOrEqual' : 'lessThan') : (r.equalAverage ? 'greaterThanOrEqual' : 'greaterThan'), '');
    case 'top10':
      return hl('rank', '', r.rank || 10, { isBottom: !!r.bottom, isPercent: !!r.percent });
    case 'duplicateValues':
    case 'uniqueValues':
      return hl(r.type, '', '');
    case 'expression': {
      const f = String(r.formulae?.[0] || '');
      const dup = f.match(/^COUNTIF\((\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+),([A-Z]+\d+)\)(>1|=1)$/i);
      if (dup && dup[1].replace(/\$/g, '') === absRef.replace(/\$/g, '')) return hl(dup[3] === '>1' ? 'duplicateValues' : 'uniqueValues', '', '');
      // Excel writes "contains" rules with their formula; turn the common shapes back into text rules
      const contains = f.match(/^NOT\(ISERROR\(SEARCH\("((?:[^"]|"")*)",([A-Z]+\d+)\)\)\)$/i);
      if (contains && contains[2] === firstTl) return hl('text', 'containsText', contains[1].replace(/""/g, '"'));
      const notContains = f.match(/^ISERROR\(SEARCH\("((?:[^"]|"")*)",([A-Z]+\d+)\)\)$/i);
      if (notContains && notContains[2] === firstTl) return hl('text', 'notContainsText', notContains[1].replace(/""/g, '"'));
      const edge = f.match(/^(LEFT|RIGHT)\(([A-Z]+\d+),(?:LEN\("(?:[^"]|"")*"\)|\d+)\)="((?:[^"]|"")*)"$/i);
      if (edge && edge[2] === firstTl) return hl('text', edge[1].toUpperCase() === 'LEFT' ? 'beginsWith' : 'endsWith', edge[3].replace(/""/g, '"'));
      return f ? hl('formula', 'containsText', `=${f}`) : null;
    }
    case 'dataBar':
      return {
        type: 'dataBar', isShowValue: r.showValue !== false,
        config: { min: cfvoIn(r.cfvo?.[0]), max: cfvoIn(r.cfvo?.[1] || { type: 'max' }), positiveColor: toHex(r.color) || '#638EC6', nativeColor: '#FF0000', isGradient: r.gradient !== false },
      };
    case 'colorScale':
      return {
        type: 'colorScale',
        config: (r.cfvo || []).map((v, i) => ({ index: i, color: toHex(r.color?.[i]) || '#FFFFFF', value: cfvoIn(v) })),
      };
    case 'iconSet': {
      const cfvo = r.cfvo || [];
      const n = cfvo.length || 3;
      const iconType = r.iconSet || '3TrafficLights1';
      const config = [];
      for (let i = 0; i < n; i++) {
        const k = n - 1 - i; // Excel index for this icon
        const last = i === n - 1;
        const v = last ? cfvo[1] : cfvo[k];
        config.push({ iconType, iconId: String(r.reverse ? n - 1 - i : i), operator: last ? 'lessThanOrEqual' : 'greaterThanOrEqual', value: cfvoIn(v || { type: 'percent', value: 0 }) });
      }
      return { type: 'iconSet', isShowValue: r.showValue !== false, config };
    }
    default:
      return null;
  }
}

// ---------------------------------------------------------------- data validation

const toIsoDate = (d) => (d instanceof Date ? d.toISOString().slice(0, 10) : String(d));
const SERIAL_EPOCH = Date.UTC(1899, 11, 30);

function dvOut(dv) {
  const base = { allowBlank: dv.allowBlank !== false, showErrorMessage: !!dv.showErrorMessage, ...(dv.error ? { error: dv.error } : {}) };
  const op = dv.operator || 'between';
  const two = op === 'between' || op === 'notBetween';
  switch (dv.type) {
    case 'list':
    case 'listMultiple': {
      const f = String(dv.formula1 || '');
      if (f.startsWith('=')) return { ...base, type: 'list', formulae: [f.slice(1).replace(/^\[[^\]]*\]/, '')] };
      let items;
      try { items = JSON.parse(f); } catch { items = null; }
      if (!Array.isArray(items)) items = f.split(',').map((s) => s.trim());
      return { ...base, type: 'list', formulae: [qstr(items.filter(Boolean).join(','))] };
    }
    case 'checkbox':
      return { ...base, type: 'list', formulae: [qstr(`${dv.formula1 ?? 1},${dv.formula2 ?? 0}`)] };
    case 'whole':
    case 'decimal':
    case 'textLength':
      return { ...base, type: dv.type, operator: op, formulae: two ? [num(dv.formula1), num(dv.formula2)] : [num(dv.formula1)] };
    case 'date': {
      const d = (s) => { const t = Date.parse(s); return Number.isFinite(t) ? new Date(t) : s; };
      return { ...base, type: 'date', operator: op, formulae: two ? [d(dv.formula1), d(dv.formula2)] : [d(dv.formula1)] };
    }
    case 'custom':
      return { ...base, type: 'custom', formulae: [String(dv.formula1 || '').replace(/^=/, '')] };
    default:
      return null;
  }
}

function dvIn(m, unitId) {
  // Excel never marks empty cells as invalid, whatever "ignore blank" says; Univer would
  const base = { uid: rid('dv'), allowBlank: true, ...(m.showErrorMessage ? { showErrorMessage: true } : {}), ...(m.error ? { error: m.error } : {}) };
  const f = m.formulae || [];
  switch (m.type) {
    case 'list': {
      const s = String(f[0] ?? '');
      if (/^".*"$/.test(s)) {
        const items = s.slice(1, -1).replace(/""/g, '"').split(',').map((x) => x.trim());
        const up = items.map((x) => x.toUpperCase());
        if (items.length === 2 && ((up[0] === 'TRUE' && up[1] === 'FALSE') || (up[0] === '1' && up[1] === '0'))) {
          return { ...base, type: 'checkbox', formula1: items[0], formula2: items[1] };
        }
        return { ...base, type: 'list', formula1: JSON.stringify(items.filter(Boolean)), showDropDown: true };
      }
      if (!s) return null;
      const ref = s.replace(/^=/, '');
      return { ...base, type: 'list', formula1: `=${ref.includes('!') ? `[${unitId}]` : ''}${ref}`, showDropDown: true };
    }
    case 'whole':
    case 'decimal':
    case 'textLength':
      return { ...base, type: m.type, operator: m.operator || 'between', formula1: String(f[0] ?? ''), ...(f[1] !== undefined ? { formula2: String(f[1]) } : {}) };
    case 'time': {
      // Univer has no time rule. A time is a fraction of a day, so the same check is a number
      // rule between those fractions (it is saved back as a number rule, which Excel enforces alike)
      const t = (x) => {
        if (x instanceof Date) return String((x.getUTCHours() * 3600 + x.getUTCMinutes() * 60 + x.getUTCSeconds()) / 86400);
        const hm = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(String(x ?? '').trim());
        return hm ? String((+hm[1] * 3600 + +hm[2] * 60 + +(hm[3] || 0)) / 86400) : String(x ?? '');
      };
      return { ...base, type: 'decimal', operator: m.operator || 'between', formula1: t(f[0]), ...(f[1] !== undefined ? { formula2: t(f[1]) } : {}) };
    }
    case 'date': {
      const d = (x) => (x instanceof Date ? toIsoDate(x) : Number.isFinite(Number(x)) ? new Date(SERIAL_EPOCH + Number(x) * 86400000).toISOString().slice(0, 10) : String(x ?? ''));
      return { ...base, type: 'date', operator: m.operator || 'between', formula1: d(f[0]), ...(f[1] !== undefined ? { formula2: d(f[1]) } : {}) };
    }
    case 'custom':
      return f[0] ? { ...base, type: 'custom', formula1: `=${String(f[0]).replace(/^=/, '')}` } : null;
    default:
      return null;
  }
}

/** Groups single cells into as few rectangles as possible (rows first, then columns) */
function cellsToRanges(cells) {
  const set = new Set(cells.map(([r, c]) => `${r},${c}`));
  const done = new Set();
  const out = [];
  const sorted = [...cells].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (const [r, c] of sorted) {
    if (done.has(`${r},${c}`)) continue;
    let w = 1;
    while (set.has(`${r},${c + w}`) && !done.has(`${r},${c + w}`)) w++;
    let h = 1;
    const rowOk = (rr) => { for (let i = 0; i < w; i++) if (!set.has(`${rr},${c + i}`) || done.has(`${rr},${c + i}`)) return false; return true; };
    while (rowOk(r + h)) h++;
    for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) done.add(`${r + i},${c + j}`);
    out.push({ startRow: r, startColumn: c, endRow: r + h - 1, endColumn: c + w - 1, rangeType: 0 });
  }
  return out;
}

// ---------------------------------------------------------------- filters (XML)

function filterXml(f) {
  const ref = f.ref;
  const cols = (f.filterColumns || []).map((fc) => {
    const colId = fc.colId - ref.startColumn;
    if (colId < 0) return '';
    if (fc.filters) {
      const vals = (fc.filters.filters || []).map((v) => `<filter val="${xmlEsc(v)}"/>`).join('');
      return `<filterColumn colId="${colId}"><filters${fc.filters.blank ? ' blank="1"' : ''}>${vals}</filters></filterColumn>`;
    }
    if (fc.customFilters) {
      const list = (fc.customFilters.customFilters || []).map((c) => `<customFilter${c.operator ? ` operator="${xmlEsc(c.operator)}"` : ''} val="${xmlEsc(c.val ?? '')}"/>`).join('');
      return `<filterColumn colId="${colId}"><customFilters${fc.customFilters.and ? ' and="1"' : ''}>${list}</customFilters></filterColumn>`;
    }
    return '';
  }).join('');
  return `<autoFilter ref="${rangeToA1(ref)}">${cols}</autoFilter>`;
}

function parseFilterXml(xml) {
  const m = xml.match(/<(?:\w+:)?autoFilter\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?autoFilter>)/);
  if (!m) return null;
  const ref = parseA1(attr(m[1], 'ref') || '');
  if (!ref) return null;
  const filterColumns = [];
  const body = m[2] || '';
  for (const fc of body.matchAll(/<(?:\w+:)?filterColumn\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?filterColumn>/g)) {
    const colId = ref.startColumn + Number(attr(fc[1], 'colId') || 0);
    const inner = fc[2];
    const fl = inner.match(/<(?:\w+:)?filters\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?filters>)/);
    if (fl) {
      const vals = [...(fl[2] || '').matchAll(/<(?:\w+:)?filter\b([^>]*)\/>/g)].map((x) => attr(x[1], 'val')).filter((x) => x !== undefined);
      const filters = { filters: vals };
      if (attr(fl[1], 'blank') === '1' || attr(fl[1], 'blank') === 'true') filters.blank = true;
      filterColumns.push({ colId, filters });
      continue;
    }
    const cf = inner.match(/<(?:\w+:)?customFilters\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?customFilters>/);
    if (cf) {
      const list = [...cf[2].matchAll(/<(?:\w+:)?customFilter\b([^>]*)\/>/g)].map((x) => {
        const op = attr(x[1], 'operator');
        return { ...(op && op !== 'equal' ? { operator: op } : {}), val: attr(x[1], 'val') ?? '' };
      });
      const and = attr(cf[1], 'and');
      filterColumns.push({ colId, customFilters: { ...(and === '1' || and === 'true' ? { and: 1 } : {}), customFilters: list } });
    }
  }
  return { ref, filterColumns };
}

/** Maps sheet names to their XML part paths in an xlsx zip */
export async function sheetParts(zip) {
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!wbXml || !relsXml) return [];
  const rels = {};
  for (const m of relsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) rels[attr(m[1], 'Id')] = attr(m[1], 'Target');
  const out = [];
  for (const m of wbXml.matchAll(/<(?:\w+:)?sheet\b([^>]*)\/?>/g)) {
    const target = rels[attr(m[1], 'r:id')];
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//, '')}`;
    out.push({ name: attr(m[1], 'name'), path });
  }
  return out;
}

/**
 * Some writers (openpyxl, Google Sheets exports) use absolute part paths in
 * relationship files. ExcelJS only follows relative ones, so rewrite them.
 * Returns true when anything changed.
 */
export async function normaliseRels(zip) {
  let changed = false;
  const rels = Object.keys(zip.files).filter((n) => /(^|\/)_rels\/[^/]+\.rels$/.test(n));
  for (const name of rels) {
    const xml = await zip.file(name).async('string');
    const baseDir = name.replace(/_rels\/[^/]+\.rels$/, ''); // e.g. 'xl/worksheets/'
    const fixed = xml.replace(/<Relationship\b[^>]*>/g, (tag) => {
      if (/TargetMode="External"/.test(tag)) return tag;
      return tag.replace(/Target="\/([^"]+)"/, (m, abs) => {
        if (!baseDir) return `Target="${abs}"`;
        const from = baseDir.split('/').filter(Boolean);
        const to = abs.split('/');
        let i = 0;
        while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
        return `Target="${[...from.slice(i).map(() => '..'), ...to.slice(i)].join('/')}"`;
      });
    });
    if (fixed !== xml) { zip.file(name, fixed); changed = true; }
  }
  return changed;
}

function resolvePart(baseDir, target) {
  const parts = (target.startsWith('/') ? target.slice(1) : baseDir + target).split('/');
  const out = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); }
  return out.join('/');
}

/** Relationships of a part: [{ id, type, target, external }] with targets resolved to zip paths */
export async function relsOf(zip, partPath) {
  const dir = partPath.replace(/[^/]+$/, '');
  const file = zip.file(`${dir}_rels/${partPath.slice(dir.length)}.rels`);
  if (!file) return [];
  const xml = await file.async('string');
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/g)].map((m) => {
    const external = attr(m[1], 'TargetMode') === 'External';
    const t = attr(m[1], 'Target') || '';
    return { id: attr(m[1], 'Id'), type: (attr(m[1], 'Type') || '').replace(/^.*\//, ''), target: external ? t : resolvePart(dir, t), external };
  });
}

const tagText = (xml, tag) => { const m = xml.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([^<]*)<`)); return m ? Number(m[1]) : 0; };
function anchorPoint(xml, tag) {
  const m = xml.match(new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*>([\\s\\S]*?)</(?:\\w+:)?${tag}>`));
  if (!m) return null;
  return { col: tagText(m[1], 'col'), colOff: tagText(m[1], 'colOff'), row: tagText(m[1], 'row'), rowOff: tagText(m[1], 'rowOff') };
}

/**
 * Reads the anchors of a drawing part: pictures and charts with their position.
 * Each: { kind: 'picture'|'chart', from, to?, ext?, target }  (offsets in EMU)
 */
export async function readDrawing(zip, drawingPath) {
  const xml = await zip.file(drawingPath)?.async('string');
  if (!xml) return [];
  const rels = await relsOf(zip, drawingPath);
  const out = [];
  for (const m of xml.matchAll(/<(?:\w+:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/(?:\w+:)?\1>/g)) {
    const body = m[2];
    const from = anchorPoint(body, 'from');
    const to = anchorPoint(body, 'to');
    const extM = body.match(/<(?:\w+:)?ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
    const ext = extM ? { cx: +extM[1], cy: +extM[2] } : null;
    const pos = body.match(/<(?:\w+:)?pos\b[^>]*\bx="(\d+)"[^>]*\by="(\d+)"/);
    const blip = body.match(/<(?:\w+:)?blip\b[^>]*\br:embed="([^"]+)"/);
    const chart = body.match(/<(?:\w+:)?chart\b[^>]*\br:id="([^"]+)"/);
    const ref = blip || chart;
    if (!ref) continue;
    const rel = rels.find((r) => r.id === ref[1]);
    if (!rel || rel.external) continue;
    out.push({ kind: blip ? 'picture' : 'chart', from, to, ext, pos: pos ? { x: +pos[1], y: +pos[2] } : null, target: rel.target });
  }
  return out;
}

/** Notes (legacy comments) of a sheet part: [{ row, col, text }] */
export async function readNotes(zip, commentsPath) {
  const xml = await zip.file(commentsPath)?.async('string');
  if (!xml) return [];
  const out = [];
  for (const m of xml.matchAll(/<(?:\w+:)?comment\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?comment>/g)) {
    const r = parseA1(attr(m[1], 'ref') || '');
    if (!r) continue;
    const text = [...m[2].matchAll(/<(?:\w+:)?t\b[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g)].map((t) => xmlUnesc(t[1])).join('');
    out.push({ row: r.startRow, col: r.startColumn, text });
  }
  return out;
}

/** Drawing and comment parts are read by us, not ExcelJS (it chokes on many real files); remove them from its copy */
export async function stripDrawings(zip) {
  for (const { path } of await sheetParts(zip)) {
    const xml = await zip.file(path)?.async('string');
    if (xml && /<(?:\w+:)?(drawing|legacyDrawing)\b/.test(xml)) zip.file(path, xml.replace(/<(?:\w+:)?(?:drawing|legacyDrawing)\b[^>]*\/>/g, ''));
    const relPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    const rx = await zip.file(relPath)?.async('string');
    // Notes are read by us as well, so their parts go too
    if (rx) zip.file(relPath, rx.replace(/<Relationship\b[^>]*\/relationships\/(?:drawing|comments|vmlDrawing)"[^>]*\/>/g, ''));
  }
  for (const n of Object.keys(zip.files)) if (/^xl\/drawings\/(_rels\/)?[^/]+\.xml(\.rels)?$/.test(n)) zip.remove(n);
}

// ---------------------------------------------------------------- geometry

const colW = (sh, c) => sh.columnData?.[c]?.w ?? sh.defaultColumnWidth ?? 88;
const rowH = (sh, r) => sh.rowData?.[r]?.h ?? sh.defaultRowHeight ?? 24;
export function xOf(sh, col, off = 0) { let x = 0; for (let i = 0; i < col; i++) x += colW(sh, i); return x + off; }
export function yOf(sh, row, off = 0) { let y = 0; for (let i = 0; i < row; i++) y += rowH(sh, i); return y + off; }
export function cellAt(sh, axis, px) {
  let i = 0;
  let pos = 0;
  const size = axis === 'x' ? (k) => colW(sh, k) : (k) => rowH(sh, k);
  while (pos + size(i) <= px && i < 20000) { pos += size(i); i++; }
  return { index: i, offset: Math.round(px - pos) };
}

function toBase64(u8) {
  if (typeof Buffer !== 'undefined' && Buffer.from) return Buffer.from(u8).toString('base64');
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s);
}
const MIME = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif' };

/** Builds a Univer drawing entry for a picture at a pixel box */
export function makeDrawing(sh, unitId, { source, col, colOff = 0, row, rowOff = 0, width, height, id }) {
  const drawingId = id || rid('img');
  const left = xOf(sh, col, colOff);
  const top = yOf(sh, row, rowOff);
  const endX = cellAt(sh, 'x', left + width);
  const endY = cellAt(sh, 'y', top + height);
  const from = { column: col, columnOffset: colOff, row, rowOffset: rowOff };
  const to = { column: endX.index, columnOffset: endX.offset, row: endY.index, rowOffset: endY.offset };
  const flat = { flipY: false, flipX: false, angle: 0, skewX: 0, skewY: 0 };
  return {
    drawingId, drawingType: 0, imageSourceType: 'BASE64', source, unitId, subUnitId: sh.id,
    column: col, columnOffset: colOff, row, rowOffset: rowOff, width, height,
    sheetTransform: { from, to, ...flat },
    axisAlignSheetTransform: { from, to, ...flat },
    transform: { left: left + ROW_HEADER_W, top: top + COL_HEADER_H, width, height, ...flat },
  };
}

// ---------------------------------------------------------------- hyperlinks

/** A cell whose text is a link, in Univer's rich-text form */
export function linkCell(text, url) {
  const t = String(text ?? url);
  return {
    v: t, t: 1,
    p: {
      id: rid('d'), documentStyle: {},
      body: {
        dataStream: `${t}\r\n`,
        textRuns: [],
        paragraphs: [{ startIndex: t.length }],
        customRanges: [{ rangeType: 0, rangeId: rid('lnk'), properties: { url }, startIndex: 0, endIndex: t.length - 1 }],
      },
    },
  };
}

/** The link inside a Univer cell, if any: { text, url } */
export function cellLink(cd) {
  const body = cd?.p?.body;
  const r = body?.customRanges?.find((x) => x.rangeType === 0 && x.properties?.url);
  if (!r) return null;
  const text = (body.dataStream || '').replace(/[\r\n]+$/, '');
  return { text: text || r.properties.url, url: r.properties.url };
}

// ---------------------------------------------------------------- import

/**
 * Reads the extra features from a loaded ExcelJS workbook into Univer resources.
 * sheets: [{ ws, sh }] in order, where sh is the Univer sheet already built.
 * Also mutates sh (filter rows, link cells).
 */
export async function importFeatures({ wb, zip, sheets, unitId }) {
  const res = { filter: {}, cf: {}, dv: {}, note: {}, drawing: {} };
  const charts = {}; // sheet id -> chart anchors, returned for the chart reader
  const parts = zip ? await sheetParts(zip) : [];
  const idByName = Object.fromEntries(sheets.map(({ ws, sh }) => [ws.name, sh.id]));

  for (const { ws, sh } of sheets) {
    const id = sh.id;

    // Filters: ExcelJS gives the range, the criteria come from the sheet XML
    const part = parts.find((p) => p.name === ws.name);
    let filter = null;
    if (part && zip) {
      const xml = await zip.file(part.path)?.async('string');
      if (xml) filter = parseFilterXml(xml);
    }
    if (filter) {
      const hidden = [];
      if (filter.filterColumns.length) {
        for (let r = filter.ref.startRow + 1; r <= filter.ref.endRow; r++) {
          if (sh.rowData?.[r]?.hd) { hidden.push(r); delete sh.rowData[r].hd; }
        }
      }
      res.filter[id] = { ref: { ...filter.ref, rangeType: 0 }, filterColumns: filter.filterColumns, cachedFilteredOut: hidden };
    }

    // Conditional formatting
    const cfs = [];
    for (const cf of ws.conditionalFormattings || []) {
      const ranges = String(cf.ref || '').split(/\s+/).map(parseA1).filter(Boolean);
      if (!ranges.length) continue;
      const tl = rangeToA1({ ...ranges[0], endRow: ranges[0].startRow, endColumn: ranges[0].startColumn });
      const abs = rangeToA1(ranges[0], true);
      const rules = [...(cf.rules || [])].sort((a, b) => (a.priority || 0) - (b.priority || 0));
      for (const r of rules) {
        const rule = cfRuleIn(r, tl, abs);
        if (rule) cfs.push({ cfId: rid('cf'), ranges, rule, stopIfTrue: !!r.stopIfTrue });
      }
    }
    if (cfs.length) res.cf[id] = cfs;

    // Data validation (ExcelJS expands ranges to single cells; group them back)
    const groups = new Map();
    for (const [addr, m] of Object.entries(ws.dataValidations?.model || {})) {
      const r = parseA1(addr);
      if (!r) continue;
      if (!groups.has(m)) groups.set(m, []);
      for (let i = r.startRow; i <= r.endRow; i++) for (let j = r.startColumn; j <= r.endColumn; j++) groups.get(m).push([i, j]);
    }
    const dvs = [];
    for (const [m, cells] of groups) {
      const rule = dvIn(m, unitId);
      if (rule) dvs.push({ ...rule, ranges: cellsToRanges(cells) });
    }
    if (dvs.length) res.dv[id] = dvs;

    const sheetRels = part && zip ? await relsOf(zip, part.path) : [];

    // Notes, read straight from the comments part
    const notes = {};
    for (const rel of sheetRels.filter((x) => x.type === 'comments' && !x.external)) {
      for (const n of await readNotes(zip, rel.target)) {
        if (n.text) (notes[n.row] ||= {})[n.col] = { id: rid('n'), row: n.row, col: n.col, note: n.text, width: 220, height: 92, show: false };
      }
    }

    // Hyperlinks, cell by cell
    ws.eachRow({ includeEmpty: true }, (row, rn) => {
      row.eachCell({ includeEmpty: true }, (cell, cn) => {
        const r = rn - 1;
        const c = cn - 1;
        const v = cell.value;
        if (v && typeof v === 'object' && v.hyperlink) {
          let url = String(v.hyperlink);
          const internal = url.match(/^#?(?:'((?:[^']|'')+)'|([^!#]+))!\$?([A-Z]+)\$?(\d+)/i);
          if (internal && (url.startsWith('#') || !/^[a-z]+:/i.test(url))) {
            const sheetName = (internal[1] || internal[2]).replace(/''/g, "'");
            const target = idByName[sheetName];
            if (target) url = `#gid=${target}&range=${internal[3]}${internal[4]}`;
          }
          const text = typeof v.text === 'object' ? (v.text.richText || []).map((x) => x.text).join('') : v.text;
          const existing = sh.cellData?.[r]?.[c] || {};
          (sh.cellData[r] ||= {})[c] = { ...existing, ...linkCell(text, url) };
        }
      });
    });
    if (Object.keys(notes).length) res.note[id] = notes;

    // Pictures (and, from the same drawing parts, charts)
    const data = {};
    const order = [];
    const px = (emu) => Math.round((emu || 0) / EMU);
    for (const rel of sheetRels.filter((x) => x.type === 'drawing' && !x.external)) {
      for (const a of await readDrawing(zip, rel.target)) {
        if (a.kind === 'chart') { (charts[id] ||= []).push(a); continue; }
        const ext = a.target.split('.').pop().toLowerCase();
        if (!MIME[ext]) continue;
        const bytes = await zip.file(a.target)?.async('uint8array');
        if (!bytes) continue;
        let col; let row; let colOff; let rowOff;
        if (a.from) ({ col, row } = a.from), colOff = px(a.from.colOff), rowOff = px(a.from.rowOff);
        else if (a.pos) {
          const cx = cellAt(sh, 'x', px(a.pos.x)); const cy = cellAt(sh, 'y', px(a.pos.y));
          col = cx.index; colOff = cx.offset; row = cy.index; rowOff = cy.offset;
        } else continue;
        let width; let height;
        if (a.to) {
          width = xOf(sh, a.to.col, px(a.to.colOff)) - xOf(sh, col, colOff);
          height = yOf(sh, a.to.row, px(a.to.rowOff)) - yOf(sh, row, rowOff);
        } else if (a.ext) { width = px(a.ext.cx); height = px(a.ext.cy); }
        if (!(width > 0 && height > 0)) continue;
        const source = `data:${MIME[ext]};base64,${toBase64(bytes)}`;
        const d = makeDrawing(sh, unitId, { source, col, colOff, row, rowOff, width, height });
        data[d.drawingId] = d;
        order.push(d.drawingId);
      }
    }
    if (order.length) res.drawing[id] = { data, order };
  }

  const resources = [];
  for (const [k, name] of Object.entries(RES)) {
    if (Object.keys(res[k]).length) resources.push({ name, data: JSON.stringify(res[k]) });
  }
  return { resources, charts };
}

// ---------------------------------------------------------------- export

export function readResources(snapshot) {
  const out = {};
  for (const [k, name] of Object.entries(RES)) {
    const r = (snapshot.resources || []).find((x) => x.name === name);
    if (!r || !r.data) { out[k] = {}; continue; }
    try { out[k] = typeof r.data === 'string' ? JSON.parse(r.data) : r.data; } catch { out[k] = {}; }
  }
  return out;
}

/**
 * Writes the extra features of one Univer sheet onto an ExcelJS worksheet.
 * Returns XML patches that ExcelJS cannot make itself ({ autoFilter }).
 */
export function exportSheetFeatures({ wb, ws, sh, res, sheetNameById }) {
  const id = sh.id;
  const patch = {};

  // Filter: range through ExcelJS, criteria patched in afterwards, filtered rows hidden
  const f = res.filter?.[id];
  if (f?.ref) {
    ws.autoFilter = rangeToA1(f.ref);
    if (f.filterColumns?.length) patch.autoFilter = filterXml(f);
    for (const r of f.cachedFilteredOut || []) ws.getRow(r + 1).hidden = true;
  }

  // Conditional formatting: one block per Univer rule, in priority order
  for (const cf of res.cf?.[id] || []) {
    const ranges = (cf.ranges || []).map((r) => clampRange(r, sh));
    if (!ranges.length) continue;
    const rule = cfRuleOut(cf.rule || {}, ranges[0], rangeToA1(ranges[0], true));
    if (!rule) continue;
    if (cf.stopIfTrue) rule.stopIfTrue = true;
    ws.addConditionalFormatting({ ref: ranges.map((r) => rangeToA1(r)).join(' '), rules: [rule] });
  }

  // Data validation
  for (const dv of res.dv?.[id] || []) {
    const rule = dvOut(dv);
    if (!rule) continue;
    for (const r of dv.ranges || []) ws.dataValidations.add(rangeToA1(clampRange(r, sh)), rule);
  }

  // Notes
  for (const [r, cols] of Object.entries(res.note?.[id] || {})) {
    for (const [c, n] of Object.entries(cols || {})) {
      if (n?.note) ws.getCell(+r + 1, +c + 1).note = String(n.note);
    }
  }

  // Hyperlinks live in the cells
  for (const [r, row] of Object.entries(sh.cellData || {})) {
    for (const [c, cd] of Object.entries(row || {})) {
      const link = cellLink(cd);
      if (!link) continue;
      let target = link.url;
      const internal = target.match(/^#gid=([^&]+)(?:&range=([A-Z]+\d+))?/i);
      if (internal) {
        const name = sheetNameById[internal[1]];
        if (!name) continue;
        target = `${quoteSheet(name)}!${(internal[2] || 'A1').toUpperCase()}`;
      }
      const cell = ws.getCell(+r + 1, +c + 1);
      cell.value = { text: link.text, hyperlink: target };
      if (!cell.font) cell.font = { color: { argb: 'FF0563C1' }, underline: true };
    }
  }

  // Pictures
  const dr = res.drawing?.[id];
  for (const key of dr?.order || Object.keys(dr?.data || {})) {
    const d = dr.data?.[key];
    if (!d || d.drawingType !== 0 || typeof d.source !== 'string') continue;
    const m = d.source.match(/^data:image\/(png|jpe?g|gif);base64,/i);
    if (!m) continue;
    const imageId = wb.addImage({ base64: d.source, extension: m[1].toLowerCase() === 'jpg' ? 'jpeg' : m[1].toLowerCase() });
    const from = d.sheetTransform?.from || { column: d.column, columnOffset: d.columnOffset, row: d.row, rowOffset: d.rowOffset };
    const width = Math.round(d.transform?.width || d.width || 100);
    const height = Math.round(d.transform?.height || d.height || 100);
    // A two-cell anchor that moves with cells (oneCellAnchor may not carry editAs)
    const x0 = xOf(sh, from.column || 0, from.columnOffset || 0);
    const y0 = yOf(sh, from.row || 0, from.rowOffset || 0);
    const ex = cellAt(sh, 'x', x0 + width);
    const ey = cellAt(sh, 'y', y0 + height);
    ws.addImage(imageId, {
      tl: { nativeCol: from.column || 0, nativeColOff: Math.round((from.columnOffset || 0) * EMU), nativeRow: from.row || 0, nativeRowOff: Math.round((from.rowOffset || 0) * EMU) },
      br: { nativeCol: ex.index, nativeColOff: ex.offset * EMU, nativeRow: ey.index, nativeRowOff: ey.offset * EMU },
      editAs: 'oneCell',
    });
  }

  return patch;
}

/** Applies XML patches to the written xlsx (filter criteria). patches: [{ name, autoFilter }] */
export async function applyPatches(zip, patches) {
  const parts = await sheetParts(zip);
  for (const p of patches) {
    if (!p.autoFilter) continue;
    const part = parts.find((x) => x.name === p.name);
    if (!part) continue;
    let xml = await zip.file(part.path).async('string');
    xml = xml.replace(/<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/, p.autoFilter);
    zip.file(part.path, xml);
  }
}
