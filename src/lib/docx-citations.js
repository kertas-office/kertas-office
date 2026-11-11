// In-text citations as real Word citation fields.
//
// In the editor a citation is plain text (SuperDoc keeps the cursor out of citation fields,
// which makes typing around them awkward), and a small registry part remembers which sources
// each one stands for. These two passes move between that and Word's own form:
//   - on save, every registered citation in the text becomes a CITATION field, so Word's
//     References tab, Update Field and bibliography see it as a citation
//   - on open, CITATION fields (and Word's citation content controls) become plain text, and
//     the caller registers what they were, so a change of citation style still rewrites them
import JSZip from 'jszip';

export const REG_NS = 'urn:kertas:citations:1';
const unXml = (t) => String(t).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const xml = (t) => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const BODY_PARTS = /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/;

/** The CITATION instruction for a list of source tags, with an optional page */
export function citationInstr(sourceIds, page) {
  const [first, ...rest] = sourceIds;
  return ` CITATION ${first}${page ? ` \\p ${String(page).replace(/[^\w.,:-]/g, '')}` : ''}${rest.map((s) => ` \\m ${s}`).join('')} \\l 1033 `;
}
/** Source tags and page from a CITATION instruction */
export function parseInstr(instr) {
  const m = /^\s*CITATION\s+(\S+)(.*)$/i.exec(instr || '');
  if (!m) return null;
  const ids = [m[1]];
  for (const x of m[2].matchAll(/\\m\s+(\S+)/g)) ids.push(x[1]);
  const page = /\\p\s+(\S+)/.exec(m[2])?.[1] || '';
  return { sourceIds: ids, page };
}

// ---------------------------------------------------------------- save: text -> fields
const RUN = /<w:r(?:\s[^>]*)?>((?:(?!<\/w:r>)[\s\S])*)<\/w:r>/g;

function fieldRuns(rPr, instr, text) {
  const r = (inner) => `<w:r>${rPr}${inner}</w:r>`;
  return r('<w:fldChar w:fldCharType="begin"/>')
    + r(`<w:instrText xml:space="preserve">${xml(instr)}</w:instrText>`)
    + r('<w:fldChar w:fldCharType="separate"/>')
    + r(`<w:t xml:space="preserve">${xml(text)}</w:t>`)
    + r('<w:fldChar w:fldCharType="end"/>');
}

/** Rewrites registered citations in one part's XML as fields. Returns { xml, count } */
export function textToFields(partXml, entries) {
  const list = entries.filter((e) => e.text && e.sourceIds?.length).sort((a, b) => b.text.length - a.text.length);
  if (!list.length) return { xml: partXml, count: 0 };
  let count = 0;
  const out = partXml.replace(RUN, (whole, inner) => {
    // plain text runs only: one <w:t>, nothing else but run properties
    const rPr = /^\s*(<w:rPr>[\s\S]*?<\/w:rPr>)?/.exec(inner)?.[1] || '';
    const rest = inner.slice(inner.indexOf(rPr) + rPr.length).trim();
    const t = /^<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>$/.exec(rest);
    if (!t) return whole;
    let text = unXml(t[1]);
    const pieces = [];
    for (;;) {
      let best = null;
      for (const e of list) { const k = text.indexOf(e.text); if (k >= 0 && (!best || k < best.k)) best = { k, e }; }
      if (!best) break;
      if (best.k) pieces.push({ text: text.slice(0, best.k) });
      pieces.push({ field: best.e });
      text = text.slice(best.k + best.e.text.length);
    }
    if (!pieces.some((p) => p.field)) return whole;
    if (text) pieces.push({ text });
    count += pieces.filter((p) => p.field).length;
    return pieces.map((p) => (p.field
      ? fieldRuns(rPr, citationInstr(p.field.sourceIds, p.field.page), p.field.text)
      : `<w:r>${rPr}<w:t xml:space="preserve">${xml(p.text)}</w:t></w:r>`)).join('');
  });
  return { xml: out, count };
}

/** Reads the registry entries from a .docx's custom XML parts */
async function registryOf(zip) {
  for (const name of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/.test(name)) continue;
    const text = await zip.file(name).async('string');
    if (!text.includes(REG_NS)) continue;
    const body = /<data>([\s\S]*?)<\/data>/.exec(text)?.[1];
    try { return body ? JSON.parse(unXml(body)) : []; } catch { return []; }
  }
  return [];
}

/** Save pass: turns registered citation text into CITATION fields. Returns the new bytes. */
export async function citationsToFields(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const entries = await registryOf(zip);
  if (!entries.length) return bytes;
  let total = 0;
  for (const name of Object.keys(zip.files)) {
    if (!BODY_PARTS.test(name)) continue;
    const src = await zip.file(name).async('string');
    const { xml: out, count } = textToFields(src, entries);
    if (count) { zip.file(name, out); total += count; }
  }
  if (!total) return bytes;
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

// ---------------------------------------------------------------- open: fields -> text
/** Turns CITATION fields in one part's XML into their result text. Returns { xml, found } */
export function fieldsToText(partXml) {
  const found = [];
  let out = partXml;
  // Word's citation content control: keep what is inside
  out = out.replace(/<w:sdt>(?:(?!<w:sdt>)[\s\S])*?<w:citation\/>[\s\S]*?<w:sdtContent>([\s\S]*?)<\/w:sdtContent>\s*<\/w:sdt>/g, (_, content) => content);
  // simple fields
  out = out.replace(/<w:fldSimple\s[^>]*w:instr="([^"]*)"[^>]*>([\s\S]*?)<\/w:fldSimple>/g, (whole, instr, content) => {
    const p = parseInstr(unXml(instr));
    if (!p) return whole;
    const text = [...content.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => unXml(m[1])).join('');
    found.push({ ...p, text });
    return content;
  });
  // complex fields: begin, instruction runs, separate, result runs, end (one paragraph)
  const runs = [...out.matchAll(RUN)].map((m) => ({ s: m.index, e: m.index + m[0].length, inner: m[1] }));
  const cuts = [];
  for (let i = 0; i < runs.length; i++) {
    if (!/<w:fldChar\b[^>]*w:fldCharType="begin"/.test(runs[i].inner)) continue;
    let instr = ''; let sep = -1; let end = -1; let depth = 0;
    for (let j = i + 1; j < runs.length; j++) {
      const r = runs[j].inner;
      if (/w:fldCharType="begin"/.test(r)) depth++;
      else if (/w:fldCharType="end"/.test(r)) { if (depth) depth--; else { end = j; break; } } else if (/w:fldCharType="separate"/.test(r) && !depth) sep = j;
      else if (sep < 0 && !depth) instr += [...r.matchAll(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g)].map((m) => unXml(m[1])).join('');
    }
    const p = end > 0 ? parseInstr(instr) : null;
    if (!p || sep < 0) continue;
    const text = runs.slice(sep + 1, end).map((r) => [...r.inner.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => unXml(m[1])).join('')).join('');
    found.push({ ...p, text });
    cuts.push({ from: runs[i].s, to: runs[sep].e }, { from: runs[end].s, to: runs[end].e });
    i = end;
  }
  for (const c of cuts.sort((a, b) => b.from - a.from)) out = out.slice(0, c.from) + out.slice(c.to);
  return { xml: out, found };
}

/** Open pass: citation fields become text. Returns { bytes, found: [{ text, sourceIds, page }] } */
export async function fieldsToCitations(bytes) {
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { return { bytes, found: [] }; }
  const found = [];
  let changed = false;
  for (const name of Object.keys(zip.files)) {
    if (!BODY_PARTS.test(name)) continue;
    const src = await zip.file(name).async('string');
    if (!/CITATION|<w:citation\/>/.test(src)) continue;
    const r = fieldsToText(src);
    if (r.xml !== src) { zip.file(name, r.xml); changed = true; found.push(...r.found); }
  }
  if (!changed) return { bytes, found };
  return { bytes: new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })), found };
}
