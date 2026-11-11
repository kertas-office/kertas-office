// Parts of an Excel workbook that Kertas carries through from opening to saving:
//   - named ranges (Formulas > Name Manager in Excel), which formulas such as =SUM(Sales) use;
//     they are given to Univer so those formulas work while editing, and written back on save
//   - sheet protection, kept exactly as it came (including its password hash)
//   - Excel tables (Format as Table), kept with their style and their header names brought up
//     to date, since Excel rejects a table whose header text no longer matches the cells
import { sheetParts, relsOf } from './xlsx-features.js';

export const NAMES_RESOURCE = 'SHEET_DEFINED_NAME_PLUGIN';
const WORKBOOK_SCOPE = 'AllDefaultWorkbook';
const attr = (s, n) => new RegExp(`\\b${n}="([^"]*)"`).exec(s)?.[1];
const unXml = (t) => String(t).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const xml = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const nameId = () => Math.random().toString(36).slice(2, 12);

// ------------------------------------------------------------------ named ranges
/** Named ranges from workbook.xml as Univer's resource data; sheetIds[i] is sheet i's id */
export async function readDefinedNames(zip, sheetIds) {
  const wbXml = await zip?.file('xl/workbook.xml')?.async('string');
  const out = {};
  if (!wbXml) return out;
  for (const m of wbXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const name = attr(m[1], 'name');
    if (!name || /^_xlnm\./i.test(name) || /^_xl/i.test(name)) continue; // print areas and Excel's own
    const ref = unXml(m[2]).trim();
    if (!ref || /#REF!/.test(ref)) continue;
    const local = attr(m[1], 'localSheetId');
    const id = nameId();
    out[id] = {
      id, name, formulaOrRefString: ref.replace(/^=/, ''),
      localSheetId: local !== undefined && sheetIds[+local] ? sheetIds[+local] : WORKBOOK_SCOPE,
      ...(attr(m[1], 'hidden') === '1' ? { hidden: true } : {}),
      ...(attr(m[1], 'comment') ? { comment: unXml(attr(m[1], 'comment')) } : {}),
    };
  }
  return out;
}

/** Writes Univer's named ranges into a saved workbook.xml */
export async function writeDefinedNames(zip, names, sheetOrderIds) {
  const list = Object.values(names || {}).filter((n) => n?.name && n.formulaOrRefString);
  if (!list.length) return;
  let wbXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!wbXml) return;
  const have = new Set([...wbXml.matchAll(/<definedName\b[^>]*\bname="([^"]*)"/g)].map((m) => m[1].toLowerCase()));
  const tags = list.filter((n) => !have.has(n.name.toLowerCase())).map((n) => {
    const local = n.localSheetId && n.localSheetId !== WORKBOOK_SCOPE ? sheetOrderIds.indexOf(n.localSheetId) : -1;
    const ref = String(n.formulaOrRefString).replace(/^=/, '');
    return `<definedName name="${xml(n.name)}"${local >= 0 ? ` localSheetId="${local}"` : ''}${n.hidden ? ' hidden="1"' : ''}${n.comment ? ` comment="${xml(n.comment)}"` : ''}>${xml(ref)}</definedName>`;
  });
  if (!tags.length) return;
  if (/<definedNames\b[^>]*>/.test(wbXml)) wbXml = wbXml.replace(/<\/definedNames>/, `${tags.join('')}</definedNames>`);
  else if (/<definedNames\s*\/>/.test(wbXml)) wbXml = wbXml.replace(/<definedNames\s*\/>/, `<definedNames>${tags.join('')}</definedNames>`);
  else wbXml = wbXml.replace(/<\/sheets>/, `</sheets><definedNames>${tags.join('')}</definedNames>`);
  zip.file('xl/workbook.xml', wbXml);
}

// ------------------------------------------------------------------ protection and tables
/** Per sheet (by name): { protection: '<sheetProtection .../>', tables: [{ xml }] } */
export async function readSheetKeep(zip) {
  const out = {};
  if (!zip) return out;
  for (const part of await sheetParts(zip)) {
    const sx = await zip.file(part.path)?.async('string');
    if (!sx) continue;
    const keep = {};
    const prot = /<sheetProtection\b[^>]*\/>/.exec(sx)?.[0];
    if (prot) keep.protection = prot;
    const tables = [];
    for (const r of await relsOf(zip, part.path)) {
      if (r.type !== 'table' || r.external) continue;
      const tx = await zip.file(r.target)?.async('string');
      if (tx) tables.push({ xml: tx });
    }
    if (tables.length) keep.tables = tables;
    if (keep.protection || keep.tables) out[part.name] = keep;
  }
  return out;
}

const colIndex = (letters) => [...letters].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const parseRef = (ref) => { const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref || ''); return m ? { c1: colIndex(m[1]), r1: +m[2] - 1, c2: colIndex(m[3]), r2: +m[4] - 1 } : null; };

/** Header text of a table, made unique and never empty, as Excel requires */
function headerNames(cells, count) {
  const seen = new Set();
  return Array.from({ length: count }, (_, i) => {
    let base = String(cells[i] ?? '').replace(/[\r\n]+/g, ' ').trim() || `Column${i + 1}`;
    let name = base; let k = 2;
    while (seen.has(name.toLowerCase())) name = `${base}${k++}`;
    seen.add(name.toLowerCase());
    return name;
  });
}

/**
 * Puts kept protection and tables back into a saved workbook.
 * bySheet: { sheetName: { protection, tables } }; headerOf(sheetName, row, col) -> cell text
 */
export async function writeSheetKeep(zip, bySheet, headerOf) {
  const entries = Object.entries(bySheet || {});
  if (!entries.length) return;
  const parts = await sheetParts(zip);
  let types = await zip.file('[Content_Types].xml').async('string');
  let tableNo = Object.keys(zip.files).filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n)).length;
  const usedTableNames = new Set();
  for (const [sheetName, keep] of entries) {
    const part = parts.find((p) => p.name === sheetName);
    if (!part) continue; // the sheet was deleted or renamed
    let sx = await zip.file(part.path).async('string');
    if (keep.protection && !/<sheetProtection\b/.test(sx)) {
      // sheetProtection sits right after sheetData (and sheetCalcPr) in the element order
      sx = /<sheetCalcPr\b[^>]*\/>/.test(sx) ? sx.replace(/(<sheetCalcPr\b[^>]*\/>)/, `$1${keep.protection}`)
        : sx.replace(/(<\/sheetData>|<sheetData\s*\/>)/, `$1${keep.protection}`);
    }
    if (keep.tables?.length && !/<tableParts\b/.test(sx)) {
      const relsPath = part.path.replace(/([^/]+)$/, '_rels/$1.rels');
      let rels = await zip.file(relsPath)?.async('string')
        || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
      const ids = [];
      for (const t of keep.tables) {
        const ref = parseRef(attr(t.xml, 'ref'));
        if (!ref) continue;
        tableNo += 1;
        let tx = t.xml.replace(/(<table\b[^>]*\bid=")\d+(")/, `$1${tableNo}$2`);
        // a name used twice would make Excel refuse the file
        let tname = attr(tx, 'name') || `Table${tableNo}`;
        while (usedTableNames.has(tname.toLowerCase())) tname = `${tname}_${tableNo}`;
        usedTableNames.add(tname.toLowerCase());
        tx = tx.replace(/(<table\b[^>]*\bname=")[^"]*(")/, `$1${xml(tname)}$2`).replace(/(<table\b[^>]*\bdisplayName=")[^"]*(")/, `$1${xml(tname)}$2`);
        // header names follow the cells as they are now
        const count = ref.c2 - ref.c1 + 1;
        const heads = headerNames(Array.from({ length: count }, (_, i) => headerOf(sheetName, ref.r1, ref.c1 + i)), count);
        let k = 0;
        tx = tx.replace(/<tableColumn\b([^>]*?)\bname="[^"]*"/g, (m, pre) => `<tableColumn${pre}name="${xml(heads[k++] ?? `Column${k}`)}"`);
        const path = `xl/tables/table${tableNo}.xml`;
        zip.file(path, tx);
        if (!types.includes(`PartName="/${path}"`)) types = types.replace('</Types>', `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`);
        let n = 1; while (rels.includes(`Id="rIdT${n}"`)) n++;
        rels = rels.replace('</Relationships>', `<Relationship Id="rIdT${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tableNo}.xml"/></Relationships>`);
        ids.push(`rIdT${n}`);
      }
      if (ids.length) {
        zip.file(relsPath, rels);
        // a table brings its own filter: a sheet filter over the same cells would clash
        sx = sx.replace(/<autoFilter\b[^>]*?(?:\/>|>[\s\S]*?<\/autoFilter>)/, '');
        const tp = `<tableParts count="${ids.length}">${ids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>`;
        sx = /<extLst>/.test(sx) ? sx.replace(/<extLst>/, `${tp}<extLst>`) : sx.replace(/<\/worksheet>/, `${tp}</worksheet>`);
        if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
      }
    }
    zip.file(part.path, sx);
  }
  zip.file('[Content_Types].xml', types);
}
