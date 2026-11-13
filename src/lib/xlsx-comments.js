// Threaded comments (Excel 365's comments with replies) between .xlsx and Univer.
//
// Excel keeps them in xl/threadedComments/threadedCommentN.xml (one per sheet), with the authors
// in xl/persons/person.xml. Older readers (LibreOffice, WPS, Excel 2016) only know notes, so each
// conversation is also written as a note holding the whole thread, whose author is "tc={id}":
// that is how Excel ties the note to the conversation and hides it. Pure module.

export const COMMENT_RES = 'SHEET_UNIVER_THREAD_COMMENT_PLUGIN';
const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const REL_THREAD = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const REL_PERSON = 'http://schemas.microsoft.com/office/2017/10/relationships/person';
const CT_THREAD = 'application/vnd.ms-excel.threadedcomments+xml';
const CT_PERSON = 'application/vnd.ms-excel.person+xml';

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unesc = (s) => String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_m, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_m, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');
const attr = (s, name) => { const m = s.match(new RegExp(`\\b${name}="([^"]*)"`)); return m ? unesc(m[1]) : null; };
const guid = () => {
  const h = (n) => Array.from({ length: n }, () => '0123456789ABCDEF'[Math.floor(Math.random() * 16)]).join('');
  return `{${h(8)}-${h(4)}-4${h(3)}-${'89AB'[Math.floor(Math.random() * 4)]}${h(3)}-${h(12)}}`;
};
/** Comment text as Univer's document body, and back */
export function toBody(text) {
  // each line ends in \r with a paragraph; the body ends in \r\n (paragraph, then section break)
  const dataStream = `${String(text).replace(/\r\n?/g, '\n').replace(/\n/g, '\r')}\r\n`;
  const paragraphs = [];
  for (let i = 0; i < dataStream.length - 1; i++) if (dataStream[i] === '\r') paragraphs.push({ startIndex: i, paragraphId: `p${i}`, paragraphStyle: {} });
  return { dataStream, textRuns: [], paragraphs, sectionBreaks: [{ startIndex: dataStream.length - 1, sectionId: 's0' }], customRanges: [] };
}
export const fromBody = (body) => String(body?.dataStream ?? '').replace(/\r\n$/, '').replace(/\r/g, '\n');
/** Excel's date stamp (no time zone, two decimals) */
const stamp = (d) => { const t = new Date(d); return Number.isNaN(+t) ? new Date().toISOString().slice(0, 22) : t.toISOString().slice(0, 22); };

/** Relationships of a part as [{ id, type, target }] (targets relative, as written) */
function parseRels(xml) {
  return [...String(xml || '').matchAll(/<Relationship\b([^>]*)\/?>/g)].map((m) => ({ id: attr(m[1], 'Id'), type: attr(m[1], 'Type') || '', target: attr(m[1], 'Target') || '' }));
}
const resolve = (dir, target) => {
  const out = [];
  for (const p of (target.startsWith('/') ? target.slice(1) : dir + target).split('/')) { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); }
  return out.join('/');
};
const relsPath = (part) => part.replace(/([^/]+)$/, '_rels/$1.rels');

/** The workbook's people: id -> display name */
async function readPersons(zip) {
  const rels = parseRels(await zip.file('xl/_rels/workbook.xml.rels')?.async('string'));
  const rel = rels.find((r) => r.type === REL_PERSON);
  const xml = rel ? await zip.file(resolve('xl/', rel.target))?.async('string') : null;
  const out = {};
  for (const m of String(xml || '').matchAll(/<(?:\w+:)?person\b([^>]*)\/?>/g)) out[attr(m[1], 'id')] = attr(m[1], 'displayName') || 'Someone';
  return out;
}

/**
 * Threaded comments of one sheet part, as Univer threads:
 * [{ id, threadId, ref, dT, personId, authorName, text, resolved, children: [...] }]
 */
export async function readThreadedComments(zip, sheetPath, persons) {
  persons ||= await readPersons(zip);
  const rels = parseRels(await zip.file(relsPath(sheetPath))?.async('string'));
  const dir = sheetPath.replace(/[^/]+$/, '');
  const roots = new Map(); const replies = [];
  for (const rel of rels.filter((r) => r.type === REL_THREAD)) {
    const xml = await zip.file(resolve(dir, rel.target))?.async('string');
    for (const m of String(xml || '').matchAll(/<(?:\w+:)?threadedComment\b([^>]*)>([\s\S]*?)<\/(?:\w+:)?threadedComment>/g)) {
      const a = m[1];
      const textM = m[2].match(/<(?:\w+:)?text\b[^>]*>([\s\S]*?)<\/(?:\w+:)?text>/);
      const id = attr(a, 'id'); const parentId = attr(a, 'parentId'); const personId = attr(a, 'personId') || '';
      const c = { id, ref: (attr(a, 'ref') || '').toUpperCase(), dT: attr(a, 'dT') || stamp(Date.now()), personId, authorName: persons[personId] || 'Someone', text: toBody(unesc(textM ? textM[1] : '')) };
      if (!id || !c.ref) continue;
      if (parentId) replies.push({ ...c, parentId });
      else roots.set(id, { ...c, threadId: id, resolved: attr(a, 'done') === '1', children: [] });
    }
  }
  for (const r of replies) { const root = roots.get(r.parentId); if (root) root.children.push({ ...r, threadId: root.id }); }
  return [...roots.values()];
}

/** The note written for older readers: the whole conversation as text */
export function threadAsNote(t) {
  const lines = [`${t.authorName || 'Someone'}: ${fromBody(t.text)}`];
  for (const c of t.children || []) lines.push(`${c.authorName || 'Someone'}: ${fromBody(c.text)}`);
  return lines.join('\n\n');
}

/**
 * Writes the conversations into a saved workbook. bySheetName: { 'Sheet1': [threads] }.
 * The notes for older readers are already on the cells (written by ExcelJS); their authors
 * become "tc={id}" here so Excel 365 shows the conversation instead.
 */
export async function writeThreadedComments(zip, bySheetName, sheetPathByName) {
  const people = new Map(); // display name -> person id
  const person = (name) => { const n = name || 'Someone'; if (!people.has(n)) people.set(n, guid()); return people.get(n); };
  let ct = await zip.file('[Content_Types].xml').async('string');
  let n = 0;
  for (const [name, threads] of Object.entries(bySheetName)) {
    const sheetPath = sheetPathByName[name];
    if (!sheetPath || !threads?.length) continue;
    n++;
    const ids = new Map(); // univer id -> excel id
    const xid = (id) => { if (!ids.has(id)) ids.set(id, /^\{[0-9A-F-]{36}\}$/i.test(id) ? id.toUpperCase() : guid()); return ids.get(id); };
    let body = '';
    for (const t of threads) {
      const rootId = xid(t.id);
      body += `<threadedComment ref="${esc(t.ref)}" dT="${esc(stamp(t.dT))}" personId="${person(t.authorName)}" id="${rootId}"${t.resolved ? ' done="1"' : ''}><text>${esc(fromBody(t.text))}</text></threadedComment>`;
      for (const c of t.children || []) body += `<threadedComment ref="${esc(t.ref)}" dT="${esc(stamp(c.dT))}" personId="${person(c.authorName)}" id="${xid(c.id)}" parentId="${rootId}"><text>${esc(fromBody(c.text))}</text></threadedComment>`;
    }
    const part = `xl/threadedComments/threadedComment${n}.xml`;
    zip.file(part, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<ThreadedComments xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${body}</ThreadedComments>`);
    ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${CT_THREAD}"/></Types>`);
    // the sheet points at its conversations
    const rp = relsPath(sheetPath);
    let rels = await zip.file(rp)?.async('string') || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const rid = `rIdTc${n}`;
    rels = rels.replace('</Relationships>', `<Relationship Id="${rid}" Type="${REL_THREAD}" Target="../threadedComments/threadedComment${n}.xml"/></Relationships>`);
    zip.file(rp, rels);
    // the matching notes: author tc={id} (ExcelJS wrote one author for all of them)
    const commentsRel = parseRels(rels).find((r) => /\/comments$/.test(r.type));
    if (commentsRel) {
      const cp = resolve(sheetPath.replace(/[^/]+$/, ''), commentsRel.target);
      let cx = await zip.file(cp)?.async('string');
      if (cx) {
        const authors = [...cx.matchAll(/<author>([\s\S]*?)<\/author>/g)].map((m) => m[1]);
        const byRef = Object.fromEntries(threads.filter((t) => !t.ownNote).map((t) => [t.ref, `tc=${xid(t.id)}`]));
        for (const a of Object.values(byRef)) if (!authors.includes(a)) authors.push(a);
        cx = cx.replace(/<authors>[\s\S]*?<\/authors>/, `<authors>${authors.map((a) => `<author>${esc(a)}</author>`).join('')}</authors>`);
        cx = cx.replace(/<comment\b([^>]*)>/g, (tag, a) => {
          const ref = (attr(a, 'ref') || '').toUpperCase();
          if (!byRef[ref]) return tag;
          return `<comment${a.replace(/\sauthorId="\d+"/, '')} authorId="${authors.indexOf(byRef[ref])}">`;
        });
        zip.file(cp, cx);
      }
    }
  }
  if (!n) return;
  // the people
  zip.file('xl/persons/person.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<personList xmlns="${NS}" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${[...people].map(([name, id]) => `<person displayName="${esc(name)}" id="${id}" userId="${esc(name)}" providerId="None"/>`).join('')}</personList>`);
  ct = ct.replace('</Types>', `<Override PartName="/xl/persons/person.xml" ContentType="${CT_PERSON}"/></Types>`);
  zip.file('[Content_Types].xml', ct);
  const wr = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  if (!wr.includes(REL_PERSON)) zip.file('xl/_rels/workbook.xml.rels', wr.replace('</Relationships>', `<Relationship Id="rIdPersons" Type="${REL_PERSON}" Target="persons/person.xml"/></Relationships>`));
}

/** Reads the comment resource of a Univer snapshot: { sheetId: [threads] } */
export function commentsOf(snapshot) {
  try {
    const r = (snapshot.resources || []).find((x) => x.name === COMMENT_RES);
    const v = r ? (typeof r.data === 'string' ? JSON.parse(r.data) : r.data) : {};
    return v && typeof v === 'object' ? v : {};
  } catch { return {}; }
}
