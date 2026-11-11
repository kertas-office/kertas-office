// Mail merge: recipients from a workbook or CSV, merge fields filled in a .docx.
// Works on the saved Word XML, so the letters keep every bit of formatting.
// Pure module (Node and browser).
import JSZip from 'jszip';
import { read as readBook, utils } from '@e965/xlsx';

/** { columns: [name], rows: [{ [name]: text }], sheets: [name] } from .xlsx/.xls/.ods/.csv bytes */
export function readRecipients(bytes, { sheet } = {}) {
  const wb = readBook(bytes, { type: 'array', cellDates: true, raw: false });
  const name = sheet && wb.SheetNames.includes(sheet) ? sheet : wb.SheetNames[0];
  const grid = utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '', blankrows: false });
  const head = (grid[0] || []).map((h, i) => String(h || '').trim() || `Column ${i + 1}`);
  // Duplicate headings get a number so every field name is unique
  const seen = {};
  const columns = head.map((h) => { seen[h] = (seen[h] || 0) + 1; return seen[h] > 1 ? `${h} ${seen[h]}` : h; });
  const rows = grid.slice(1)
    .filter((r) => r.some((v) => String(v).trim() !== ''))
    .map((r) => Object.fromEntries(columns.map((c, i) => [c, String(r[i] ?? '')])));
  return { columns, rows, sheets: wb.SheetNames, sheet: name };
}

/** Word's field name for a column: letters, digits and underscores */
export const fieldName = (column) => String(column).trim().replace(/[^\p{L}\p{N}_]+/gu, '_').replace(/^_+|_+$/g, '') || 'Field';

/** "MERGEFIELD  "First Name" \* MERGEFORMAT" -> "First Name" */
export function mergeFieldName(instr) {
  const m = String(instr).trim().match(/^MERGEFIELD\s+(?:"([^"]+)"|(\S+))/i);
  return m ? (m[1] ?? m[2]) : null;
}

const xmlEsc = (s) => String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const textRun = (rPr, value) => `<w:r>${rPr}${String(value).split(/\r?\n/).map((line, i) => `${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${xmlEsc(line)}</w:t>`).join('')}</w:r>`;
// Removes the invisible field anchors. Text such as "\u200B is due" keeps its leading
// space only if the <w:t> says so, so xml:space="preserve" is added where needed.
export const stripAnchors = (xml) => xml.replace(/<w:t(\s[^>]*)?>([^<]*)<\/w:t>/g, (m, attrs = '', text) => {
  if (!text.includes('\u200B')) return m;
  const t = text.replace(/\u200B/g, '');
  const keep = /^\s|\s$/.test(t) && !/xml:space=/.test(attrs) ? `${attrs} xml:space="preserve"` : attrs;
  return `<w:t${keep}>${t}</w:t>`;
}).replace(/\u200B/g, '');
const firstRPr = (xml) => xml.match(/<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>/)?.[0] || '';
const shownTextOf = (xml) => [...xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
// Text typed straight after a field can end up inside its result ("«Name», thanks"):
// whatever follows the «Name» placeholder is kept after the merged value.
function withTail(value, shown, name) {
  const ph = `«${name}»`;
  const k = shown.indexOf(ph);
  return k >= 0 ? `${shown.slice(0, k).replace(/^\s+$/, '')}${value}${shown.slice(k + ph.length)}`.replace(/\u200B/g, '') : value;
}
const instrTextOf = (xml) => [...xml.matchAll(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/g)].map((m) => m[1]).join('').replace(/&quot;/g, '"').replace(/&amp;/g, '&');

/**
 * Replaces MERGEFIELDs in one Word XML part. lookup(name) -> value or undefined (left as is).
 * Handles simple fields (w:fldSimple) and complex fields (fldChar begin/separate/end).
 */
export function fillFields(xml, lookup) {
  // Simple fields
  xml = xml.replace(/<w:fldSimple\b[^>]*w:instr="([^"]*)"[^>]*?(?:\/>|>([\s\S]*?)<\/w:fldSimple>)/g, (all, instr, inner = '') => {
    const name = mergeFieldName(instr.replace(/&quot;/g, '"'));
    const value = name != null ? lookup(name) : undefined;
    return value === undefined ? all : textRun(firstRPr(inner), withTail(value, shownTextOf(inner), name));
  });
  // Complex fields: find each begin, walk to its matching end
  let out = '';
  let pos = 0;
  const begin = /<w:fldChar\b[^>]*w:fldCharType="begin"[^>]*\/?>/g;
  let m;
  while ((m = begin.exec(xml))) {
    const runStart = Math.max(xml.lastIndexOf('<w:r>', m.index), xml.lastIndexOf('<w:r ', m.index));
    if (runStart < pos) continue;
    // Walk fldChars, tracking nesting
    const walker = /<w:fldChar\b[^>]*w:fldCharType="(begin|separate|end)"[^>]*\/?>/g;
    walker.lastIndex = m.index;
    let depth = 0; let sep = -1; let endAt = -1; let w;
    while ((w = walker.exec(xml))) {
      if (w[1] === 'begin') depth++;
      else if (w[1] === 'separate' && depth === 1 && sep < 0) sep = w.index;
      else if (w[1] === 'end' && --depth === 0) { endAt = w.index; break; }
    }
    if (endAt < 0) break;
    const runEnd = xml.indexOf('</w:r>', endAt) + '</w:r>'.length;
    const span = xml.slice(runStart, runEnd);
    if (/<\/w:p>/.test(span)) { begin.lastIndex = endAt; continue; } // spans paragraphs: leave it
    const instr = instrTextOf(xml.slice(m.index, sep > 0 ? sep : endAt));
    const name = mergeFieldName(instr);
    const value = name != null ? lookup(name) : undefined;
    if (value === undefined) { begin.lastIndex = endAt; continue; }
    const shown = sep > 0 ? xml.slice(sep, endAt) : '';
    out += xml.slice(pos, runStart) + textRun(firstRPr(shown) || firstRPr(span), withTail(value, shownTextOf(shown), name));
    pos = runEnd;
    begin.lastIndex = runEnd;
  }
  return out + xml.slice(pos);
}

/** Lists the merge fields used in a .docx (names, in order of first use) */
export async function fieldsInDocx(docx) {
  const zip = await JSZip.loadAsync(docx);
  const names = [];
  for (const path of Object.keys(zip.files).filter((p) => /^word\/(document|header\d*|footer\d*)\.xml$/.test(p))) {
    // Document order: simple fields carry the instruction in w:instr, complex ones in w:instrText
    const xml = await zip.file(path).async('string');
    for (const m of xml.matchAll(/w:instr="([^"]*)"|<w:instrText[^>]*>([^<]*)</g)) {
      const n = mergeFieldName((m[1] ?? m[2]).replace(/&quot;/g, '"').replace(/&amp;/g, '&'));
      if (n && !names.includes(n)) names.push(n);
    }
  }
  return names;
}

function lookupFor(record) {
  const byField = new Map(Object.entries(record).map(([k, v]) => [fieldName(k).toLowerCase(), v]));
  for (const [k, v] of Object.entries(record)) byField.set(k.toLowerCase(), v);
  return (name) => {
    const v = byField.get(String(name).toLowerCase()) ?? byField.get(fieldName(name).toLowerCase());
    return v === undefined ? '' : v; // unknown fields become blank in merged letters
  };
}
const PARTS = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

/** One merged .docx per record: [{ data, record }] */
export async function mergeSeparate(docx, records) {
  const src = await JSZip.loadAsync(docx);
  const parts = {};
  for (const p of Object.keys(src.files).filter((f) => PARTS.test(f))) parts[p] = await src.file(p).async('string');
  const out = [];
  for (const record of records) {
    const zip = await JSZip.loadAsync(docx);
    const look = lookupFor(record);
    for (const [p, xml] of Object.entries(parts)) zip.file(p, stripAnchors(fillFields(xml, look)));
    out.push({ record, data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) });
  }
  return out;
}

/** All records in one .docx, each starting on a new page */
export async function mergeSingle(docx, records) {
  const zip = await JSZip.loadAsync(docx);
  const xml = await zip.file('word/document.xml').async('string');
  const bodyOpen = xml.indexOf('>', xml.indexOf('<w:body')) + 1;
  const bodyClose = xml.lastIndexOf('</w:body>');
  let body = xml.slice(bodyOpen, bodyClose);
  // The last section's settings close the body; they are written once at the end
  const lastSect = body.lastIndexOf('<w:sectPr');
  const sectPr = lastSect >= 0 && !body.slice(lastSect).includes('</w:p>') ? body.slice(lastSect) : '';
  if (sectPr) body = body.slice(0, lastSect);
  // Ids that must be unique in a document get renumbered in each copy
  body = body.replace(/\sw14:(?:paraId|textId)="[^"]*"/g, '');
  const copies = records.map((record, i) => {
    let copy = stripAnchors(fillFields(body, lookupFor(record)));
    copy = copy.replace(/(<wp:docPr\b[^>]*\bid=")(\d+)"/g, (all, a, id) => `${a}${Number(id) + i * 10000}"`);
    copy = copy.replace(/(<w:bookmark(?:Start|End)\b[^>]*\bw:id=")(\d+)"/g, (all, a, id) => `${a}${Number(id) + i * 10000}"`);
    copy = copy.replace(/(<w:bookmarkStart\b[^>]*\bw:name=")([^"]*)"/g, (all, a, n) => (i ? `${a}${n}_${i + 1}"` : all));
    return copy;
  });
  const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  zip.file('word/document.xml', xml.slice(0, bodyOpen) + copies.join(pageBreak) + sectPr + xml.slice(bodyClose));
  const look = lookupFor(records[0] || {});
  for (const p of Object.keys(zip.files).filter((f) => PARTS.test(f) && f !== 'word/document.xml')) {
    zip.file(p, fillFields(await zip.file(p).async('string'), look));
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** File-name-safe text from a record, e.g. "Letter - Aminah Yusof" */
export function recordLabel(record, columns, i) {
  const pick = columns.find((c) => /name/i.test(c)) || columns[0];
  const extra = /first/i.test(pick) ? columns.find((c) => /last|surname|family/i.test(c)) : null;
  const text = [record[pick], extra && record[extra]].filter(Boolean).join(' ').trim();
  return (text || `Record ${i + 1}`).replace(/[\\/:*?"<>|]+/g, ' ').slice(0, 60).trim();
}
