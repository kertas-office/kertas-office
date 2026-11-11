// Real PivotTables in .xlsx files. Kertas keeps a pivot as live formulas on its
// own sheet; when saving, it also writes Excel's pivot cache and pivot table
// parts over those cells, marked to refresh on open, so Excel, WPS and
// LibreOffice show a genuine PivotTable with its field list. Opening a file
// reads simple pivots (one row field, optional column field, one value) back.
// Pure module.
import { colName, parseA1, rangeToA1 } from './xlsx-features.js';
import { distinct, normSpec, isSimpleSpec, aggLabel } from './pivot.js';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = {
  cacheDef: `${R_NS}/pivotCacheDefinition`,
  records: `${R_NS}/pivotCacheRecords`,
  table: `${R_NS}/pivotTable`,
};
const CT = {
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  cacheDef: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  records: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
};
const SUBTOTAL = { sum: null, count: 'count', average: 'average', max: 'max', min: 'min' };
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const blank = (v) => v === '' || v === null || v === undefined;

function sharedItems(values, axis) {
  const present = values.filter((v) => !blank(v));
  const nums = present.filter(isNum);
  const hasBlank = present.length < values.length;
  if (!axis) {
    if (present.length && nums.length === present.length) return `<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"${nums.every(Number.isInteger) ? ' containsInteger="1"' : ''} minValue="${Math.min(...nums)}" maxValue="${Math.max(...nums)}"${hasBlank ? ' containsBlank="1"' : ''}/>`;
    return `<sharedItems${hasBlank ? ' containsBlank="1"' : ''}/>`;
  }
  const keys = distinct(values);
  const items = keys.map((k) => (isNum(k) ? `<n v="${k}"/>` : `<s v="${esc(k)}"/>`)).join('');
  const allNum = keys.length && keys.every(isNum);
  const attrs = allNum ? ` containsSemiMixedTypes="0" containsString="0" containsNumber="1"${keys.every(Number.isInteger) ? ' containsInteger="1"' : ''} minValue="${Math.min(...keys)}" maxValue="${Math.max(...keys)}"` : keys.some(isNum) ? ' containsMixedTypes="1" containsNumber="1"' : '';
  return `<sharedItems${attrs}${hasBlank ? ' containsBlank="1"' : ''} count="${keys.length}">${items}</sharedItems>`;
}

/**
 * The three XML parts for one pivot.
 * p: { name, source: { sheet, range }, grid (source values incl. header), spec: { rows, columns, values, agg }, label }
 */
export function pivotXml(p, cacheId) {
  const n = normSpec(p.spec);
  if (!isSimpleSpec(n)) return pivotXmlMulti(p, cacheId, n);
  p = { ...p, spec: { rows: n.rows[0], columns: n.columns.length ? n.columns[0] : null, values: n.values[0].fld, agg: n.values[0].agg } };
  const [head, ...data] = p.grid;
  const fields = head.map((h, i) => (blank(h) ? `Column ${colName(p.source.range.startColumn + i)}` : String(h)));
  const col = (i) => data.map((r) => r[i]);
  const { rows, columns, values } = p.spec;
  const hasCol = columns !== null && columns !== undefined;
  const rowKeys = distinct(col(rows));
  const colKeys = hasCol ? distinct(col(columns)) : [];
  const agg = p.spec.agg || 'sum';

  const cacheFields = fields.map((f, i) => `<cacheField name="${esc(f)}" numFmtId="0">${sharedItems(col(i), i === rows || (hasCol && i === columns))}</cacheField>`).join('');
  const cacheDef = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${NS}" xmlns:r="${R_NS}" r:id="rId1" refreshOnLoad="1" refreshedBy="Kertas Office" refreshedDate="${(Date.now() / 86400000 + 25569).toFixed(6)}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="${rangeToA1(p.source.range)}" sheet="${esc(p.source.sheet)}"/></cacheSource><cacheFields count="${fields.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
  const records = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="${NS}" xmlns:r="${R_NS}" count="0"/>`;

  const items = (keys) => `<items count="${keys.length + 1}">${keys.map((_, k) => `<item x="${k}"/>`).join('')}<item t="default"/></items>`;
  const pivotFields = fields.map((_, i) => {
    if (i === rows) return `<pivotField axis="axisRow" showAll="0">${items(rowKeys)}</pivotField>`;
    if (hasCol && i === columns) return `<pivotField axis="axisCol" showAll="0">${items(colKeys)}</pivotField>`;
    if (i === values) return '<pivotField dataField="1" showAll="0"/>';
    return '<pivotField showAll="0"/>';
  }).join('');
  const axisItems = (keys) => `${keys.map((_, k) => (k ? `<i><x v="${k}"/></i>` : '<i><x/></i>')).join('')}<i t="grand"><x/></i>`;
  // Kertas puts the value caption in A1; Excel's compact layout matches that
  // when the table starts at A1 with a column field, or at A2 without one.
  const top = hasCol ? 0 : 1;
  const width = hasCol ? colKeys.length + 2 : 2;
  const height = rowKeys.length + (hasCol ? 3 : 2);
  const ref = rangeToA1({ startRow: top, startColumn: 0, endRow: top + height - 1, endColumn: width - 1 });
  const sub = SUBTOTAL[agg];
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${NS}" name="${esc(p.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0" rowHeaderCaption="${esc(fields[rows])}"${hasCol ? ` colHeaderCaption="${esc(fields[columns])}"` : ''}>`
    + `<location ref="${ref}" firstHeaderRow="1" firstDataRow="${hasCol ? 2 : 1}" firstDataCol="1"/>`
    + `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>`
    + `<rowFields count="1"><field x="${rows}"/></rowFields><rowItems count="${rowKeys.length + 1}">${axisItems(rowKeys)}</rowItems>`
    + (hasCol ? `<colFields count="1"><field x="${columns}"/></colFields><colItems count="${colKeys.length + 1}">${axisItems(colKeys)}</colItems>` : '<colItems count="1"><i/></colItems>')
    + `<dataFields count="1"><dataField name="${esc(p.label)}" fld="${values}"${sub ? ` subtotal="${sub}"` : ''} baseField="0" baseItem="0"/></dataFields>`
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
    + '</pivotTableDefinition>';
  return { cacheDef, records, table, ref };
}

/**
 * Several row, column or value fields: a tabular PivotTable (row fields side by side), the same
 * layout Kertas writes in the cells. Item lists are left for Excel to fill in, as the table is
 * marked to refresh when it opens.
 */
function pivotXmlMulti(p, cacheId, n) {
  const [head, ...data] = p.grid;
  const fields = head.map((h, i) => (blank(h) ? `Column ${colName(p.source.range.startColumn + i)}` : String(h)));
  const col = (i) => data.map((r) => r[i]);
  const axis = new Set([...n.rows, ...n.columns]);
  const cacheFields = fields.map((f, i) => `<cacheField name="${esc(f)}" numFmtId="0">${sharedItems(col(i), axis.has(i))}</cacheField>`).join('');
  const cacheDef = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${NS}" xmlns:r="${R_NS}" r:id="rId1" refreshOnLoad="1" refreshedBy="Kertas Office" refreshedDate="${(Date.now() / 86400000 + 25569).toFixed(6)}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="${rangeToA1(p.source.range)}" sheet="${esc(p.source.sheet)}"/></cacheSource><cacheFields count="${fields.length}">${cacheFields}</cacheFields></pivotCacheDefinition>`;
  const records = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="${NS}" xmlns:r="${R_NS}" count="0"/>`;
  const items = (keys) => `<items count="${keys.length + 1}">${keys.map((_, k) => `<item x="${k}"/>`).join('')}<item t="default"/></items>`;
  const dataSet = new Set(n.values.map((v) => v.fld));
  const pivotFields = fields.map((_, i) => {
    const flat = ' compact="0" outline="0" showAll="0"';
    if (n.rows.includes(i)) return `<pivotField axis="axisRow"${flat} defaultSubtotal="0">${items(distinct(col(i)))}</pivotField>`;
    if (n.columns.includes(i)) return `<pivotField axis="axisCol"${flat} defaultSubtotal="0">${items(distinct(col(i)))}</pivotField>`;
    return `<pivotField${dataSet.has(i) ? ' dataField="1"' : ''}${flat}/>`;
  }).join('');
  const many = n.values.length > 1;
  const colFieldList = [...n.columns, ...(many ? [-2] : [])];
  const headerRows = Math.max(1, n.columns.length + (many && n.columns.length ? 1 : 0));
  // the area the cells use: title row, header rows, rows, total
  const rowCount = distinctTuples(data, n.rows);
  const colCount = n.columns.length ? distinctTuples(data, n.columns) : 1;
  const width = n.rows.length + (n.columns.length ? colCount * n.values.length + n.values.length : n.values.length);
  const height = 1 + headerRows + rowCount + 1;
  const ref = rangeToA1({ startRow: 1, startColumn: 0, endRow: height - 1, endColumn: width - 1 });
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${NS}" name="${esc(p.name)}" cacheId="${cacheId}" applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0" compact="0" compactData="0" outline="0" outlineData="0" gridDropZones="1" multipleFieldFilters="0">`
    + `<location ref="${ref}" firstHeaderRow="${headerRows - 1 || 1}" firstDataRow="${headerRows}" firstDataCol="${n.rows.length}"/>`
    + `<pivotFields count="${fields.length}">${pivotFields}</pivotFields>`
    + `<rowFields count="${n.rows.length}">${n.rows.map((x) => `<field x="${x}"/>`).join('')}</rowFields>`
    + (colFieldList.length ? `<colFields count="${colFieldList.length}">${colFieldList.map((x) => `<field x="${x}"/>`).join('')}</colFields>` : '')
    + `<dataFields count="${n.values.length}">${n.values.map((v) => `<dataField name="${esc(`${aggLabel(v.agg)} of ${fields[v.fld]}`)}" fld="${v.fld}"${SUBTOTAL[v.agg] ? ` subtotal="${SUBTOTAL[v.agg]}"` : ''} baseField="0" baseItem="0"/>`).join('')}</dataFields>`
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
    + '</pivotTableDefinition>';
  return { cacheDef, records, table, ref };
}
function distinctTuples(data, cols) {
  const seen = new Set();
  for (const r of data) { const t = cols.map((c) => r[c]); if (t.some(blank)) continue; seen.add(t.map((v) => (isNum(v) ? `n:${v}` : `s:${String(v).toLowerCase()}`)).join('\u0001')); }
  return seen.size;
}

async function readRels(zip, path) {
  const xml = await zip.file(path)?.async('string');
  return xml || `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
}
function addRel(relsXml, type, target) {
  const used = new Set([...relsXml.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]));
  let n = 1; while (used.has(`rId${n}`)) n++;
  const id = `rId${n}`;
  return { xml: relsXml.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`), id };
}
const relsPathOf = (part) => part.replace(/([^/]+)$/, '_rels/$1.rels');

/**
 * Writes pivots into a saved workbook. list: [{ sheetPath (where the pivot sits), ...pivotXml input }]
 */
export async function injectPivots(zip, list) {
  if (!list.length) return;
  let ct = await zip.file('[Content_Types].xml').async('string');
  let wbXml = await zip.file('xl/workbook.xml').async('string');
  let wbRels = await readRels(zip, 'xl/_rels/workbook.xml.rels');
  const caches = [];
  let n = 0;
  for (const p of list) {
    n++;
    const cacheId = 100 + n;
    const { cacheDef, records, table } = pivotXml(p, cacheId);
    const defPath = `xl/pivotCache/pivotCacheDefinition${n}.xml`;
    const recPath = `xl/pivotCache/pivotCacheRecords${n}.xml`;
    const tablePath = `xl/pivotTables/pivotTable${n}.xml`;
    zip.file(defPath, cacheDef);
    zip.file(recPath, records);
    zip.file(tablePath, table);
    zip.file(relsPathOf(defPath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL.records}" Target="pivotCacheRecords${n}.xml"/></Relationships>`);
    zip.file(relsPathOf(tablePath), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL.cacheDef}" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`);
    const sheetRelsPath = relsPathOf(p.sheetPath);
    const sr = addRel(await readRels(zip, sheetRelsPath), REL.table, `../pivotTables/pivotTable${n}.xml`);
    zip.file(sheetRelsPath, sr.xml);
    const wr = addRel(wbRels, REL.cacheDef, `pivotCache/pivotCacheDefinition${n}.xml`);
    wbRels = wr.xml;
    caches.push(`<pivotCache cacheId="${cacheId}" r:id="${wr.id}"/>`);
    for (const [path, type] of [[tablePath, CT.table], [defPath, CT.cacheDef], [recPath, CT.records]]) {
      if (!ct.includes(`PartName="/${path}"`)) ct = ct.replace('</Types>', `<Override PartName="/${path}" ContentType="${type}"/></Types>`);
    }
  }
  if (!/xmlns:r=/.test(wbXml.slice(0, 600))) wbXml = wbXml.replace('<workbook ', `<workbook xmlns:r="${R_NS}" `);
  const block = `<pivotCaches>${caches.join('')}</pivotCaches>`;
  // pivotCaches comes after calcPr and before any extLst
  wbXml = /<extLst>/.test(wbXml) ? wbXml.replace('<extLst>', `${block}<extLst>`) : wbXml.replace('</workbook>', `${block}</workbook>`);
  zip.file('xl/workbook.xml', wbXml);
  zip.file('xl/_rels/workbook.xml.rels', wbRels);
  zip.file('[Content_Types].xml', ct);
}

const attr = (s, n) => s.match(new RegExp(`\\b${n}="([^"]*)"`))?.[1];
const unesc = (s) => String(s ?? '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const resolve = (base, target) => {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); }
  return parts.join('/');
};

/**
 * Simple pivots in an opened workbook: [{ sheetPath, name, location, source: { sheet, range }, spec, fields }]
 * Several row, column and value fields come back in the list form of normSpec; pivots built
 * on a named range or with other summary types are reported with supported: false.
 */
export async function readPivots(zip) {
  const out = [];
  const tables = Object.keys(zip.files).filter((f) => /^xl\/pivotTables\/[^/]+\.xml$/.test(f));
  if (!tables.length) return out;
  // Which sheet holds each pivot table
  const holder = new Map();
  for (const relsPath of Object.keys(zip.files).filter((f) => /^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(f))) {
    const sheetPath = relsPath.replace('_rels/', '').replace(/\.rels$/, '');
    const xml = await zip.file(relsPath).async('string');
    for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) if (/pivotTable"/.test(attr(m[0], 'Type') + '"')) holder.set(resolve(sheetPath, attr(m[0], 'Target')), sheetPath);
  }
  for (const t of tables) {
    const xml = await zip.file(t).async('string');
    const rels = await zip.file(relsPathOf(t))?.async('string') || '';
    const defTarget = [...rels.matchAll(/<Relationship\b[^>]*>/g)].map((m) => m[0]).find((r) => /pivotCacheDefinition"/.test(`${attr(r, 'Type')}"`));
    const defPath = defTarget && resolve(t, attr(defTarget, 'Target'));
    const def = defPath && await zip.file(defPath)?.async('string');
    if (!def) continue;
    const src = def.match(/<worksheetSource\b[^>]*>/)?.[0] || '';
    const range = parseA1(attr(src, 'ref') || '');
    const fields = [...def.matchAll(/<cacheField\b([^>]*)/g)].map((m) => unesc(attr(m[1], 'name')));
    const fieldList = (tag) => [...(xml.match(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`))?.[1] || '').matchAll(/<field\b[^>]*x="(-?\d+)"/g)].map((m) => Number(m[1])).filter((x) => x >= 0);
    const rowsF = fieldList('rowFields');
    const colsF = fieldList('colFields');
    const dataF = [...xml.matchAll(/<dataField\b([^>]*)\/?>/g)].map((m) => ({ fld: Number(attr(m[1], 'fld')), subtotal: attr(m[1], 'subtotal') || 'sum', name: unesc(attr(m[1], 'name')) }));
    const loc = xml.match(/<location\b[^>]*>/)?.[0] || '';
    // several row, column and value fields are fine; the Values pseudo-field (-2) is dropped above
    const supported = !!range && rowsF.length >= 1 && dataF.length >= 1 && dataF.every((d) => ['sum', 'count', 'average', 'max', 'min'].includes(d.subtotal)) && !attr(src, 'name');
    out.push({
      sheetPath: holder.get(t) || null,
      name: unesc(attr(xml.match(/<pivotTableDefinition\b[^>]*>/)?.[0] || '', 'name')) || 'PivotTable',
      location: parseA1(attr(loc, 'ref') || ''),
      source: { sheet: unesc(attr(src, 'sheet') || ''), range },
      fields,
      spec: !supported ? null
        : rowsF.length === 1 && colsF.length <= 1 && dataF.length === 1
          ? { rows: rowsF[0], columns: colsF.length ? colsF[0] : null, values: dataF[0].fld, agg: dataF[0].subtotal }
          : { rows: rowsF, columns: colsF, values: dataF.map((d) => ({ fld: d.fld, agg: d.subtotal })) },
      supported,
    });
  }
  return out;
}
