// PivotTable summaries. A pivot is written as ordinary cells with
// SUMIFS / COUNTIFS / AVERAGEIFS / MAXIFS / MINIFS formulas that point back at
// the source data, so it stays live when the numbers change and opens the same
// in Excel, WPS and LibreOffice. Pure module.
import { colName } from './xlsx-features.js';
import { quoteSheet } from './xlsx-charts.js';

export const AGGREGATES = [
  { id: 'sum', label: 'Sum' },
  { id: 'count', label: 'Count' },
  { id: 'average', label: 'Average' },
  { id: 'max', label: 'Max' },
  { id: 'min', label: 'Min' },
];

const isNum = (v) => typeof v === 'number' || (v !== '' && v !== null && v !== undefined && typeof v !== 'boolean' && Number.isFinite(Number(v)));
const blank = (v) => v === '' || v === null || v === undefined;

/** Names of a table's columns: the header text, or "Column F" for a blank header */
export const fieldNames = (head, startColumn = 0) => head.map((h, i) => (blank(h) ? `Column ${colName(startColumn + i)}` : String(h)));

/** Sorted distinct non-blank values: numbers first (ascending), then text A-Z */
export function distinct(values) {
  const seen = new Map();
  for (const v of values) {
    if (blank(v)) continue;
    const key = isNum(v) ? `n:${Number(v)}` : `s:${String(v).toLowerCase()}`;
    if (!seen.has(key)) seen.set(key, isNum(v) ? Number(v) : String(v));
  }
  return [...seen.values()].sort((a, b) => {
    const na = typeof a === 'number';
    const nb = typeof b === 'number';
    if (na && nb) return a - b;
    if (na !== nb) return na ? -1 : 1;
    return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
  });
}

/**
 * A pivot's settings in one shape: { rows: [col], columns: [col], values: [{ fld, agg }] }.
 * Accepts the single-field form too: { rows, columns, values, agg }.
 */
export function normSpec(spec) {
  const list = (x) => (Array.isArray(x) ? x : x === null || x === undefined ? [] : [x]);
  const values = Array.isArray(spec.values)
    ? spec.values.map((v) => (typeof v === 'object' ? { fld: v.fld, agg: v.agg || 'sum' } : { fld: v, agg: spec.agg || 'sum' }))
    : list(spec.values).map((fld) => ({ fld, agg: spec.agg || 'sum' }));
  return { rows: list(spec.rows), columns: list(spec.columns), values };
}
export const isSimpleSpec = (n) => n.rows.length === 1 && n.columns.length <= 1 && n.values.length === 1;
export const aggLabel = (agg) => AGGREGATES.find((a) => a.id === agg)?.label || 'Sum';

/** Distinct combinations of several columns, in sorted order */
function tuples(data, cols) {
  if (!cols.length) return [[]];
  const seen = new Map();
  for (const r of data) {
    const t = cols.map((c) => r[c]);
    if (t.some(blank)) continue;
    const key = t.map((v) => (isNum(v) ? `n:${Number(v)}` : `s:${String(v).toLowerCase()}`)).join('\u0001');
    if (!seen.has(key)) seen.set(key, t.map((v) => (isNum(v) ? Number(v) : String(v))));
  }
  const cmp = (a, b) => {
    const na = typeof a === 'number'; const nb = typeof b === 'number';
    if (na && nb) return a - b;
    if (na !== nb) return na ? -1 : 1;
    return a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
  };
  return [...seen.values()].sort((x, y) => { for (let k = 0; k < x.length; k++) { const c = cmp(x[k], y[k]); if (c) return c; } return 0; });
}

/**
 * grid: the source block including its header row (2D values)
 * src: { sheet: 'Data', startRow, startColumn }  (top-left of grid on its sheet, 0-based)
 * spec: see normSpec (one or several row, column and value fields)
 * Returns { cells: 2D array of { v } | { f } | null, label, rows, columns, headerRows, rowFields }
 *
 * Layout (tabular, as Excel's "Show in tabular form"): a title row, one header row per column
 * field (plus one naming the values when there are several), then one row per combination of
 * row fields with a column for each row field, and a Grand Total row and column.
 */
export function buildPivot(grid, src, spec) {
  const n = normSpec(spec);
  if (!n.rows.length || !n.values.length) throw new Error('A PivotTable needs a row field and a value field');
  const [head, ...data] = grid;
  const name = (i) => (blank(head[i]) ? `Column ${colName(src.startColumn + i)}` : String(head[i]));
  const firstRow = src.startRow + 2; // 1-based row of the first data row
  const lastRow = src.startRow + grid.length;
  const col = (i) => `${quoteSheet(src.sheet)}!$${colName(src.startColumn + i)}$${firstRow}:$${colName(src.startColumn + i)}$${lastRow}`;
  const ifs = (v, crit) => {
    const valRange = col(v.fld);
    const pairs = crit.map(([range, ref]) => `${range},${ref}`).join(',');
    switch (v.agg) {
      case 'count': return `COUNTIFS(${pairs})`;
      case 'average': return `IFERROR(AVERAGEIFS(${valRange},${pairs}),"")`;
      case 'max': return `MAXIFS(${valRange},${pairs})`;
      case 'min': return `MINIFS(${valRange},${pairs})`;
      default: return `SUMIFS(${valRange},${pairs})`;
    }
  };
  const vLabel = (v) => `${aggLabel(v.agg)} of ${name(v.fld)}`;
  const label = n.values.length === 1 ? vLabel(n.values[0]) : 'Values';
  const rowTuples = tuples(data, n.rows);
  const colTuples = n.columns.length ? tuples(data, n.columns) : [[]];
  const nr = n.rows.length;
  const many = n.values.length > 1;
  const headerRows = n.columns.length + (many && n.columns.length ? 1 : 0) || 1;
  const HEAD = 1 + headerRows; // title row + header rows

  // data columns: [{ tuple, v }] then grand totals per value
  const dataCols = [];
  if (n.columns.length) for (const t of colTuples) for (const v of n.values) dataCols.push({ t, v });
  else for (const v of n.values) dataCols.push({ t: [], v });
  const totalCols = n.columns.length ? n.values.map((v) => ({ v })) : [];

  const cells = [];
  cells.push([{ v: label }]);
  if (n.columns.length) {
    // one row per column field; the last header row also names the row fields
    n.columns.forEach((cf, k) => {
      const row = Array.from({ length: nr }, (_, j) => (k === headerRows - 1 ? { v: name(n.rows[j]) } : (j === 0 && k === 0 ? { v: name(cf) } : null)));
      dataCols.forEach((dc) => row.push({ v: dc.t[k] }));
      totalCols.forEach((tc, j) => row.push(k === 0 ? { v: many ? `Total ${vLabel(tc.v)}` : 'Grand Total' } : null));
      cells.push(row);
    });
    if (many) cells.push([...n.rows.map((r) => ({ v: name(r) })), ...dataCols.map((dc) => ({ v: vLabel(dc.v) })), ...totalCols.map(() => null)]);
  } else {
    cells.push([...n.rows.map((r) => ({ v: name(r) })), ...n.values.map((v) => ({ v: vLabel(v) }))]);
  }
  // a column's criteria refer to its header cells (row k + 2, 1-based)
  const colCrit = (dc, c) => n.columns.map((cf, k) => [col(cf), `${c}$${2 + k}`]);
  rowTuples.forEach((t, i) => {
    const r = HEAD + i + 1;
    const rowCrit = n.rows.map((rf, j) => [col(rf), `$${colName(j)}${r}`]);
    const row = t.map((v) => ({ v }));
    dataCols.forEach((dc, j) => {
      const c = colName(nr + j);
      row.push({ f: `=${ifs(dc.v, [...rowCrit, ...colCrit(dc, c)])}` });
    });
    totalCols.forEach((tc) => row.push({ f: `=${ifs(tc.v, rowCrit)}` }));
    cells.push(row);
  });
  const total = [{ v: 'Grand Total' }, ...Array.from({ length: nr - 1 }, () => null)];
  const any = [[col(n.rows[0]), '"<>"']];
  dataCols.forEach((dc, j) => {
    const c = colName(nr + j);
    total.push({ f: `=${ifs(dc.v, n.columns.length ? colCrit(dc, c) : any)}` });
  });
  totalCols.forEach((tc) => total.push({ f: `=${ifs(tc.v, any)}` }));
  cells.push(total);
  // which value each data column shows (for number formats)
  const valueCols = [...dataCols.map((dc) => dc.v), ...totalCols.map((tc) => tc.v)];
  return { cells, label, rows: rowTuples.length, columns: n.columns.length ? colTuples.length : 0, headerRows: HEAD, rowFields: nr, valueCols, rowCols: n.rows, colRows: n.columns };
}
