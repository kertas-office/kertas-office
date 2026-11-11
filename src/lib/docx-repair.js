// A last pass over a saved .docx, comparing it with the file as it was opened.
//
// The editor rebuilds each part's relationships when it saves, and a few it does not know
// (such as the link from an embedded spreadsheet's icon to the spreadsheet itself) can go
// missing while the text still points at them. Word reports a file like that as damaged.
// This puts such relationships back from the original, with the parts and content types
// they need, and removes macro code that nothing refers to any more (a .docx may not hold
// macros, and Word will not open one that does).
import JSZip from 'jszip';

const REL_ATTR = /\br:[A-Za-z]+="([^"]+)"/g;
const relsPathOf = (part) => part.replace(/(^|\/)([^/]+)$/, '$1_rels/$2.rels');
const dirOf = (part) => part.replace(/[^/]*$/, '');
const RELATIONSHIP = /<Relationship\b[^>]*\/>/g;
const attr = (tag, name) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

/** Where a relationship's Target points, as a path inside the package */
function resolve(fromPart, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (dirOf(fromPart) + target).split('/');
  const out = [];
  for (const p of parts) { if (p === '..') out.pop(); else if (p !== '.' && p !== '') out.push(p); }
  return out.join('/');
}

/** Returns repaired bytes, or the same bytes when nothing needed fixing */
export async function repairDocx(savedBytes, originalBytes) {
  let saved; let orig;
  try { saved = await JSZip.loadAsync(savedBytes); } catch { return savedBytes; }
  try { orig = originalBytes ? await JSZip.loadAsync(originalBytes) : null; } catch { orig = null; }
  let changed = false;
  let types = await saved.file('[Content_Types].xml')?.async('string');
  const origTypes = orig ? await orig.file('[Content_Types].xml')?.async('string') : '';

  // 1. relationships the text still uses
  if (orig) {
    for (const part of Object.keys(saved.files)) {
      if (!/^word\/[^/]+\.xml$/.test(part)) continue;
      const xml = await saved.file(part).async('string');
      const used = new Set([...xml.matchAll(REL_ATTR)].map((m) => m[1]).filter(Boolean));
      if (!used.size) continue;
      const relsPath = relsPathOf(part);
      let rels = await saved.file(relsPath)?.async('string');
      const origRels = await orig.file(relsPath)?.async('string');
      if (!origRels) continue;
      const have = new Set((rels || '').match(RELATIONSHIP)?.map((t) => attr(t, 'Id')) || []);
      const add = [];
      for (const tag of origRels.match(RELATIONSHIP) || []) {
        const id = attr(tag, 'Id');
        if (!used.has(id) || have.has(id)) continue;
        const external = attr(tag, 'TargetMode') === 'External';
        const target = attr(tag, 'Target');
        if (!external) {
          const path = resolve(part, target);
          if (!saved.file(path)) {
            const src = orig.file(path);
            if (!src) continue; // nothing to point at
            saved.file(path, await src.async('uint8array'));
          }
          // its content type: an override from the original, unless the extension already has one
          const ext = path.split('.').pop().toLowerCase();
          const hasDefault = new RegExp(`<Default\\b[^>]*Extension="${ext}"`, 'i').test(types);
          const hasOverride = types.includes(`PartName="/${path}"`);
          if (!hasDefault && !hasOverride) {
            const ov = new RegExp(`<Override\\b[^>]*PartName="/${path.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`).exec(origTypes)?.[0]
              || new RegExp(`<Default\\b[^>]*Extension="${ext}"[^>]*/>`, 'i').exec(origTypes)?.[0];
            if (ov) types = types.replace('</Types>', `${ov}</Types>`);
          }
        }
        add.push(tag);
      }
      if (add.length) {
        rels = rels || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        saved.file(relsPath, rels.replace('</Relationships>', `${add.join('')}</Relationships>`));
        changed = true;
      }
    }
  }

  // 2. macro code with nothing pointing at it
  const vba = Object.keys(saved.files).filter((n) => /^word\/vbaProject\.bin$/i.test(n));
  if (vba.length) {
    const docRels = await saved.file('word/_rels/document.xml.rels')?.async('string') || '';
    if (!/vbaProject\.bin/i.test(docRels)) {
      for (const n of ['word/vbaProject.bin', 'word/_rels/vbaProject.bin.rels', 'word/vbaData.xml']) saved.remove(n);
      types = types.replace(/<Override\b[^>]*PartName="\/word\/(vbaProject\.bin|vbaData\.xml)"[^>]*\/>/gi, '');
      changed = true;
    }
  }

  if (!changed) return savedBytes;
  saved.file('[Content_Types].xml', types);
  return new Uint8Array(await saved.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
