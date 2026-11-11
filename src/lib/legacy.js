// Older and OpenDocument formats: .doc, .rtf, .odt, .xls, .xlsb, .ods.
// They open as Kertas documents and are saved as .docx / .xlsx.
//  - Desktop with LibreOffice installed: LibreOffice converts them, formatting kept.
//  - Spreadsheets otherwise: read with SheetJS (values, formulas, merges, widths, number formats).
//  - .doc otherwise: its text is extracted (paragraphs, tables, notes) on the desktop.
import { read as readBook, write as writeBook } from '@e965/xlsx';

export const LEGACY = { doc: 'writer', rtf: 'writer', odt: 'writer', xls: 'sheets', xlsb: 'sheets', ods: 'sheets' };
export const isLegacy = (name) => !!LEGACY[String(name).toLowerCase().split('.').pop()];

/** .xls / .xlsb / .ods bytes -> .xlsx bytes */
export function spreadsheetToXlsx(bytes) {
  const wb = readBook(bytes, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true, cellDates: false });
  const out = writeBook(wb, { type: 'array', bookType: 'xlsx', cellStyles: true, compression: true });
  return new Uint8Array(out);
}

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** word-extractor text parts -> HTML for the Writer (tables from tab-separated rows) */
export function docTextToHtml({ body = '', footnotes = '', endnotes = '', comments = '' } = {}) {
  const html = [];
  const lines = String(body).replace(/\r\n?/g, '\n').replace(/\n{3,}/g, '\n\n').split('\n');
  let table = null;
  const flush = () => {
    if (!table) return;
    const cols = Math.max(...table.map((r) => r.length));
    html.push(`<table>${table.map((r) => `<tr>${Array.from({ length: cols }, (_, i) => `<td>${esc(r[i] ?? '')}</td>`).join('')}</tr>`).join('')}</table>`);
    table = null;
  };
  for (const line of lines) {
    if (line.includes('\t') && line.replace(/\t/g, '').trim()) {
      (table ||= []).push(line.replace(/\t+$/, '').split('\t').map((c) => c.trim()));
      continue;
    }
    flush();
    html.push(line.trim() ? `<p>${esc(line)}</p>` : '<p></p>');
  }
  flush();
  while (html.length && html[html.length - 1] === '<p></p>') html.pop();
  const section = (title, text) => {
    const t = String(text || '').trim();
    if (t) html.push(`<h2>${esc(title)}</h2>`, ...t.split(/\n+/).map((l) => `<p>${esc(l)}</p>`));
  };
  section('Footnotes', footnotes);
  section('Endnotes', endnotes);
  section('Comments', comments);
  return html.join('');
}

/**
 * Turns an opened legacy file into something a module can load.
 * rec: { name, data }; native: platform convert function or null.
 * Returns { data } (docx/xlsx bytes) or { html } plus a note for the user.
 */
export async function openLegacy(rec, convertNative) {
  const ext = rec.name.toLowerCase().split('.').pop();
  const kind = LEGACY[ext];
  if (convertNative) {
    const res = await convertNative(rec);
    if (res?.data) return { data: res.data, note: `Converted from .${ext} with LibreOffice. Saving makes a .${res.ext} copy; the original is not changed.` };
    if (res?.text) return { html: docTextToHtml(res.text), note: `Opened the text of this .doc file. Install LibreOffice to keep its formatting too. Saving makes a .docx copy.` };
  }
  if (kind === 'sheets') {
    return { data: spreadsheetToXlsx(rec.data), note: `Opened .${ext} with values, formulas and layout. ${convertNative ? 'Install LibreOffice to keep colours and fonts too. ' : ''}Saving makes a .xlsx copy.` };
  }
  throw new Error(convertNative ? `Opening .${ext} files needs LibreOffice (free) installed` : `Open .${ext} files in the Kertas desktop app`);
}
