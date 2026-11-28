// Spreadsheets in the page: any workbook or table read with SheetJS, written as XLSX, ODS,
// CSV, TSV, JSON or HTML. A CSV, TSV or JSON of a workbook with several sheets is one file
// per sheet.
import { extOf } from './formats.js';
import { htmlPage, textOf } from './doc.js';

async function lib() { return import('@e965/xlsx'); }

export async function readWorkbook(bytes, name) {
  const XLSX = await lib();
  const ext = extOf(name);
  if (ext === 'json') {
    const data = JSON.parse(textOf(bytes));
    const wb = XLSX.utils.book_new();
    const add = (rows, sheet) => {
      const list = Array.isArray(rows) ? rows : [rows];
      const ws = list.every((r) => Array.isArray(r)) ? XLSX.utils.aoa_to_sheet(list) : XLSX.utils.json_to_sheet(list.map((r) => (r && typeof r === 'object' ? r : { value: r })));
      XLSX.utils.book_append_sheet(wb, ws, String(sheet).slice(0, 31) || 'Sheet1');
    };
    // { "Sheet1": [...], "Sheet2": [...] } is a workbook; anything else is one sheet
    if (data && !Array.isArray(data) && typeof data === 'object' && Object.values(data).every(Array.isArray)) Object.entries(data).forEach(([k, v]) => add(v, k));
    else add(data, 'Sheet1');
    return wb;
  }
  if (ext === 'csv' || ext === 'tsv') {
    return XLSX.read(textOf(bytes), { type: 'string', FS: ext === 'tsv' ? '\t' : undefined, raw: false, cellDates: true });
  }
  return XLSX.read(bytes, { type: 'array', cellDates: true, cellStyles: true });
}

/** [{ suffix, data }]: one entry, or one per sheet for the single-sheet formats */
export async function writeWorkbook(wb, to) {
  const XLSX = await lib();
  const enc = new TextEncoder();
  const perSheet = (fn) => wb.SheetNames.map((sheet) => ({ suffix: wb.SheetNames.length > 1 ? ` - ${sheet}` : '', data: fn(wb.Sheets[sheet], sheet) }));
  switch (to) {
    case 'xlsx':
    case 'ods':
      return [{ suffix: '', data: new Uint8Array(XLSX.write(wb, { bookType: to, type: 'array', compression: true })) }];
    case 'csv':
      // a mark at the start so Excel reads accents and other scripts correctly
      return perSheet((ws) => enc.encode(`﻿${XLSX.utils.sheet_to_csv(ws, { blankrows: false })}`));
    case 'tsv':
      return perSheet((ws) => enc.encode(XLSX.utils.sheet_to_csv(ws, { FS: '\t', blankrows: false })));
    case 'json': {
      const all = Object.fromEntries(wb.SheetNames.map((s) => [s, XLSX.utils.sheet_to_json(wb.Sheets[s], { defval: null, raw: true })]));
      const value = wb.SheetNames.length === 1 ? all[wb.SheetNames[0]] : all;
      return [{ suffix: '', data: enc.encode(`${JSON.stringify(value, null, 2)}\n`) }];
    }
    case 'html':
      return [{ suffix: '', data: enc.encode(await workbookHtml(wb)) }];
    default: throw new Error(`Spreadsheets cannot be saved as ${to}`);
  }
}

/** Every sheet as a table on a page, for HTML and PDF */
export async function workbookHtml(wb, title = 'Workbook') {
  const XLSX = await lib();
  const parts = wb.SheetNames.map((s) => {
    const table = XLSX.utils.sheet_to_html(wb.Sheets[s], { header: '', footer: '' });
    return `${wb.SheetNames.length > 1 ? `<h2>${s.replace(/[&<>]/g, '')}</h2>` : ''}${table}`;
  });
  return htmlPage(`<style>body{max-width:none} table{font-size:9pt} td{white-space:nowrap}</style>${parts.join('\n')}`, title);
}
