// Dialogs for the Sheets module
import { openDialog } from '../lib/dialog.js';
import { CHART_TYPES } from '../lib/xlsx-charts.js';
import { AGGREGATES } from '../lib/pivot.js';

const CHART_ICONS = { column: 'chart-column', bar: 'chart-bar', line: 'chart-line', area: 'chart-area', pie: 'chart-pie', doughnut: 'chart-donut', scatter: 'chart-scatter' };

/** Insert or edit a chart. init: { type, title, range, byRows, headers, labels, legend, stacked } */
export async function chartDialog(init = {}, { editing = false } = {}) {
  const res = await openDialog({
    title: editing ? 'Edit chart' : 'Insert chart',
    description: 'Pick a chart type and the cells to plot. The chart updates when the cells change.',
    width: 640,
    fields: [
      { id: 'type', label: 'Chart type', type: 'segmented', value: init.type || 'column', options: CHART_TYPES.map((t) => ({ value: t.id, label: t.label, icon: CHART_ICONS[t.id] })) },
      { id: 'title', label: 'Title', value: init.title ?? '', placeholder: 'Chart title (optional)' },
      { id: 'range', label: 'Data', value: init.range || '', placeholder: "Sheet1!A1:C10", hint: 'Include the header row and the label column.' },
      { id: 'byRows', label: 'Series in', type: 'segmented', value: init.byRows ? 'rows' : 'cols', options: [{ value: 'cols', label: 'Columns' }, { value: 'rows', label: 'Rows' }] },
      { id: 'headers', type: 'checkbox', label: 'First row has names', value: init.headers ?? true },
      { id: 'labels', type: 'checkbox', label: 'First column has labels', value: init.labels ?? true },
      { id: 'legend', label: 'Legend', type: 'segmented', value: init.legend || 'bottom', options: [{ value: 'bottom', label: 'Bottom' }, { value: 'right', label: 'Right' }, { value: 'top', label: 'Top' }, { value: 'none', label: 'None' }] },
      { id: 'stacked', type: 'checkbox', label: 'Stack the series', value: !!init.stacked, show: (v) => ['column', 'bar', 'area'].includes(v.type) },
    ],
    validate: (v) => (/^(?:(?:'(?:[^']|'')+'|[^!]+)!)?\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i.test(v.range.trim()) ? null : 'Enter a cell range such as Sheet1!A1:C10'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: editing ? 'Update chart' : 'Insert chart', primary: true }],
  });
  if (!res) return null;
  const v = res.values;
  return { ...v, range: v.range.trim(), byRows: v.byRows === 'rows' };
}

/** PivotTable: fields are the source header names. Up to two row fields and two values. */
export async function pivotDialog(fields, rangeText) {
  const opts = fields.map((f, i) => ({ value: String(i), label: f }));
  const none = { value: '', label: '(none)' };
  const numericGuess = Math.max(0, fields.length - 1);
  const aggs = AGGREGATES.map((a) => ({ value: a.id, label: a.label }));
  const res = await openDialog({
    title: 'PivotTable',
    description: `Summarises ${rangeText} on a new sheet. The summary uses live formulas, so it follows changes to the numbers; Refresh picks up new rows and categories.`,
    width: 540,
    fields: [
      { id: 'rows', label: 'Rows', type: 'select', value: '0', options: opts },
      { id: 'rows2', label: 'Then by', type: 'select', value: '', options: [none, ...opts], hint: 'A second row field, shown beside the first' },
      { id: 'columns', label: 'Columns', type: 'select', value: '', options: [none, ...opts] },
      { id: 'values', label: 'Values', type: 'select', value: String(numericGuess), options: opts },
      { id: 'agg', label: 'Summarise by', type: 'segmented', value: 'sum', options: aggs },
      { id: 'values2', label: 'Also show', type: 'select', value: '', options: [none, ...opts], hint: 'A second value, in its own column' },
      { id: 'agg2', label: 'Summarise it by', type: 'segmented', value: 'count', options: aggs, show: (v) => v.values2 !== '' },
    ],
    validate: (v) => {
      const used = [v.rows, v.rows2, v.columns].filter((x) => x !== '');
      if (new Set(used).size !== used.length) return 'Pick different fields for rows and columns';
      return null;
    },
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Create', primary: true }],
  });
  if (!res) return null;
  const v = res.values;
  if (v.rows2 === '' && v.values2 === '') return { rows: Number(v.rows), columns: v.columns === '' ? null : Number(v.columns), values: Number(v.values), agg: v.agg };
  return {
    rows: [Number(v.rows), ...(v.rows2 !== '' ? [Number(v.rows2)] : [])],
    columns: v.columns === '' ? [] : [Number(v.columns)],
    values: [{ fld: Number(v.values), agg: v.agg }, ...(v.values2 !== '' ? [{ fld: Number(v.values2), agg: v.agg2 }] : [])],
  };
}

export const FUNCTIONS = [
  { cat: 'Common', items: [
    ['SUM', 'number1, number2, ...', 'Adds numbers'], ['AVERAGE', 'number1, number2, ...', 'Average of numbers'], ['COUNT', 'value1, value2, ...', 'Counts cells with numbers'],
    ['MAX', 'number1, number2, ...', 'Largest value'], ['MIN', 'number1, number2, ...', 'Smallest value'], ['IF', 'test, value_if_true, value_if_false', 'Chooses a value by a condition'],
  ] },
  { cat: 'Math', items: [
    ['ROUND', 'number, digits', 'Rounds to a number of digits'], ['SUMIF', 'range, criteria, [sum_range]', 'Adds cells that meet a condition'], ['SUMIFS', 'sum_range, range1, criteria1, ...', 'Adds cells that meet several conditions'],
    ['SUMPRODUCT', 'array1, array2, ...', 'Sum of products'], ['ABS', 'number', 'Absolute value'], ['MOD', 'number, divisor', 'Remainder'], ['ROUNDUP', 'number, digits', 'Rounds up'], ['ROUNDDOWN', 'number, digits', 'Rounds down'],
  ] },
  { cat: 'Logical', items: [
    ['IFS', 'test1, value1, test2, value2, ...', 'First value whose test is true'], ['AND', 'logical1, logical2, ...', 'True if all are true'], ['OR', 'logical1, logical2, ...', 'True if any is true'],
    ['NOT', 'logical', 'Reverses a condition'], ['IFERROR', 'value, value_if_error', 'Replaces errors'], ['SWITCH', 'expression, value1, result1, ...', 'Matches a value against a list'],
  ] },
  { cat: 'Lookup', items: [
    ['XLOOKUP', 'lookup, lookup_array, return_array, [if_not_found]', 'Finds a value and returns its match'], ['VLOOKUP', 'lookup, table, col_index, [range_lookup]', 'Looks down the first column'],
    ['HLOOKUP', 'lookup, table, row_index, [range_lookup]', 'Looks across the first row'], ['INDEX', 'array, row, [column]', 'Value at a position'], ['MATCH', 'lookup, lookup_array, [match_type]', 'Position of a value'],
  ] },
  { cat: 'Text', items: [
    ['CONCAT', 'text1, text2, ...', 'Joins text'], ['TEXTJOIN', 'delimiter, ignore_empty, text1, ...', 'Joins text with a delimiter'], ['LEFT', 'text, [num_chars]', 'Characters from the start'],
    ['RIGHT', 'text, [num_chars]', 'Characters from the end'], ['MID', 'text, start, num_chars', 'Characters from the middle'], ['LEN', 'text', 'Length of text'], ['TRIM', 'text', 'Removes extra spaces'],
    ['UPPER', 'text', 'Upper case'], ['LOWER', 'text', 'Lower case'], ['PROPER', 'text', 'Title Case'], ['TEXT', 'value, format', 'Formats a number as text'], ['SUBSTITUTE', 'text, old, new, [instance]', 'Replaces text'],
  ] },
  { cat: 'Date', items: [
    ['TODAY', '', "Today's date"], ['NOW', '', 'Date and time now'], ['DATE', 'year, month, day', 'Makes a date'], ['YEAR', 'date', 'Year of a date'], ['MONTH', 'date', 'Month of a date'],
    ['DAY', 'date', 'Day of a date'], ['EDATE', 'start, months', 'Date months away'], ['EOMONTH', 'start, months', 'End of a month'], ['NETWORKDAYS', 'start, end, [holidays]', 'Working days between dates'], ['DATEDIF', 'start, end, unit', 'Difference between dates'],
  ] },
  { cat: 'Statistical', items: [
    ['COUNTA', 'value1, value2, ...', 'Counts non-empty cells'], ['COUNTIF', 'range, criteria', 'Counts cells that meet a condition'], ['COUNTIFS', 'range1, criteria1, ...', 'Counts with several conditions'],
    ['AVERAGEIF', 'range, criteria, [average_range]', 'Average with a condition'], ['MEDIAN', 'number1, number2, ...', 'Middle value'], ['STDEV.S', 'number1, number2, ...', 'Standard deviation (sample)'],
    ['LARGE', 'array, k', 'k-th largest'], ['SMALL', 'array, k', 'k-th smallest'], ['RANK.EQ', 'number, ref, [order]', 'Rank of a number'],
  ] },
  { cat: 'Financial', items: [
    ['PMT', 'rate, nper, pv, [fv], [type]', 'Loan payment'], ['FV', 'rate, nper, pmt, [pv], [type]', 'Future value'], ['PV', 'rate, nper, pmt, [fv], [type]', 'Present value'],
    ['NPV', 'rate, value1, value2, ...', 'Net present value'], ['IRR', 'values, [guess]', 'Internal rate of return'],
  ] },
];

export async function functionDialog(cellA1, suggested = '', preset = 'SUM') {
  const all = FUNCTIONS.flatMap((g) => g.items.map(([name, args, desc]) => ({ value: name, label: `${name}  -  ${desc}`, args, cat: g.cat })));
  const res = await openDialog({
    title: 'Insert function',
    description: `Writes the formula into ${cellA1}.`,
    width: 560,
    fields: [
      { id: 'fn', label: 'Function', type: 'select', value: preset || 'SUM', options: FUNCTIONS.flatMap((g) => g.items.map(([name, , desc]) => ({ value: name, label: `${g.cat}: ${name} - ${desc}` }))) },
      { id: 'args', label: 'Arguments', value: suggested, placeholder: 'e.g. B2:B10', hint: 'Separate arguments with commas. Use cell ranges like A1:A10.' },
      { id: 'info', type: 'info', html: '' },
    ],
    validate: () => null,
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    onChange: (v, dlg) => {
      const f = all.find((x) => x.value === v.fn);
      const box = dlg?.querySelector('[data-row="info"] .dlg-info');
      if (box && f) box.innerHTML = `<code>${f.value}(${f.args})</code>`;
    },
  });
  if (!res) return null;
  return `=${res.values.fn}(${res.values.args.trim()})`;
}

export async function sizeDialog(kind, current) {
  const res = await openDialog({
    title: kind === 'row' ? 'Row height' : 'Column width',
    fields: [{ id: 'px', label: 'Size in pixels', type: 'number', min: 2, max: 2000, value: current }],
    validate: (v) => (v.px >= 2 && v.px <= 2000 ? null : 'Use 2 to 2000 pixels'),
    width: 340,
  });
  return res ? Number(res.values.px) : null;
}

const CF_TITLES = {
  greaterThan: 'Greater than', lessThan: 'Less than', between: 'Between', equal: 'Equal to', textContains: 'Text that contains',
};
export async function cfRuleDialog(kind) {
  const two = kind === 'between';
  const text = kind === 'textContains';
  const res = await openDialog({
    title: `Highlight cells: ${CF_TITLES[kind]}`,
    description: 'Formats the selected cells that match.',
    width: 440,
    fields: [
      { id: 'a', label: text ? 'Text' : two ? 'From' : 'Value', type: text ? 'text' : 'number', value: '' },
      ...(two ? [{ id: 'b', label: 'To', type: 'number', value: '' }] : []),
      { id: 'style', label: 'Format with', type: 'select', value: 'red', options: [
        { value: 'red', label: 'Light red fill with dark red text' }, { value: 'yellow', label: 'Yellow fill with dark yellow text' },
        { value: 'green', label: 'Green fill with dark green text' }, { value: 'redText', label: 'Red text' }, { value: 'bold', label: 'Bold' },
      ] },
    ],
    validate: (v) => (text ? (v.a.trim() ? null : 'Enter some text') : (v.a === null || (two && v.b === null) ? 'Enter a number' : null)),
  });
  return res?.values || null;
}

export async function removeDuplicatesDialog(columns) {
  const res = await openDialog({
    title: 'Remove duplicates',
    description: 'Rows that repeat in the ticked columns are removed; the first one is kept.',
    width: 420,
    fields: [
      { id: 'headers', type: 'checkbox', label: 'My data has headers', value: true },
      ...columns.map((c, i) => ({ id: `c${i}`, type: 'checkbox', label: c, value: true })),
    ],
    validate: (v) => (columns.some((_, i) => v[`c${i}`]) ? null : 'Tick at least one column'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Remove duplicates', primary: true }],
  });
  if (!res) return null;
  return { headers: res.values.headers, cols: columns.map((_, i) => i).filter((i) => res.values[`c${i}`]) };
}
