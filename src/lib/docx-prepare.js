// Small fixes to a .docx before the editor reads it, for content the editor would otherwise drop.
//
// Point comments: LibreOffice (and older Word files) can attach a comment to a single spot,
// with only a <w:commentReference> and no commented range. The editor only keeps comments that
// cover some text, so on save the comment lost its place in the document and Word no longer
// showed it. Such a comment is given a range over the words just before the spot (or just after,
// at the start of a paragraph), which Word and LibreOffice display the same way.
import JSZip from 'jszip';

const RUN = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;

/** Gives point comments a range in one part's XML; returns the XML (unchanged if nothing to do) */
export function anchorPointComments(xml) {
  if (!xml.includes('commentReference')) return xml;
  const ranged = new Set([...xml.matchAll(/<w:commentRangeStart\b[^>]*\bw:id="([^"]+)"/g)].map((m) => m[1]));
  return xml.replace(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g, (para) => {
    if (!para.includes('commentReference')) return para;
    // innermost paragraphs only (a text box paragraph sits inside another one)
    if (para.slice(4).includes('<w:p ') || para.slice(4).includes('<w:p>')) return para;
    const runs = [...para.matchAll(RUN)].map((m) => ({ text: m[0], at: m.index }));
    const hasText = (r) => /<w:t[\s>]/.test(r.text) && !/commentReference/.test(r.text);
    const wraps = new Map(); // run index -> comment ids to put around it
    for (let i = 0; i < runs.length; i++) {
      const id = /<w:commentReference\b[^>]*\bw:id="([^"]+)"/.exec(runs[i].text)?.[1];
      if (id === undefined || ranged.has(id)) continue;
      let j = i - 1;
      while (j >= 0 && !hasText(runs[j])) j--;
      if (j < 0) { j = i + 1; while (j < runs.length && !hasText(runs[j])) j++; }
      if (j < 0 || j >= runs.length) continue;
      if (!wraps.has(j)) wraps.set(j, []);
      wraps.get(j).push(id);
      ranged.add(id);
    }
    if (!wraps.size) return para;
    let out = '';
    let pos = 0;
    runs.forEach((r, k) => {
      const ids = wraps.get(k);
      if (!ids) return;
      out += para.slice(pos, r.at);
      out += ids.map((id) => `<w:commentRangeStart w:id="${id}"/>`).join('') + r.text + ids.map((id) => `<w:commentRangeEnd w:id="${id}"/>`).join('');
      pos = r.at + r.text.length;
    });
    out += para.slice(pos);
    return out;
  });
}

/** Returns prepared bytes, or the same bytes when nothing needed changing */
export async function prepareDocx(bytes) {
  let zip;
  try { zip = await JSZip.loadAsync(bytes); } catch { return bytes; }
  const doc = await zip.file('word/document.xml')?.async('string');
  if (!doc) return bytes;
  const fixed = anchorPointComments(doc);
  if (fixed === doc) return bytes;
  zip.file('word/document.xml', fixed);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
