// Documents in the page: Word to text, HTML and Markdown (mammoth), Markdown to HTML (marked),
// HTML to Markdown (turndown), and HTML, Markdown or text to Word with Kertas's own writer.

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const decodeText = (bytes) => {
  // UTF-8 with or without a mark; anything that is not UTF-8 is read as Windows-1252
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, ''); } catch { return new TextDecoder('windows-1252').decode(bytes); }
};
export const textOf = decodeText;

/** A whole page around a piece of HTML, styled for reading and printing */
export function htmlPage(body, title = 'Document') {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>
body { font: 11pt/1.5 "Segoe UI", Inter, Arial, sans-serif; color: #111; max-width: 46em; margin: 0 auto; padding: 1em; }
h1, h2, h3, h4 { line-height: 1.25; margin: 1.2em 0 .5em; }
table { border-collapse: collapse; margin: 1em 0; }
td, th { border: 1px solid #bbb; padding: 4px 8px; vertical-align: top; }
th { background: #f1f5f9; }
img { max-width: 100%; }
pre, code { font-family: Consolas, "Courier New", monospace; font-size: 10pt; }
pre { background: #f6f8fa; padding: 8px; white-space: pre-wrap; }
blockquote { border-left: 3px solid #cbd5e1; margin: 1em 0; padding-left: 1em; color: #334155; }
@page { margin: 2cm; }
@media print { body { max-width: none; padding: 0; } }
</style></head><body>${body}</body></html>`;
}

export async function docxToHtml(bytes) {
  const mammoth = (await import('mammoth/mammoth.browser.js')).default;
  const { value } = await mammoth.convertToHtml({ arrayBuffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
  return value;
}
export async function docxToText(bytes) {
  const mammoth = (await import('mammoth/mammoth.browser.js')).default;
  const { value } = await mammoth.extractRawText({ arrayBuffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) });
  return value;
}
export async function markdownToHtml(md) {
  const { marked } = await import('marked');
  return marked.parse(md, { gfm: true, async: false });
}
export async function htmlToMarkdown(html) {
  const TurndownService = (await import('turndown')).default;
  const td = new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' });
  return td.turndown(bodyOf(html));
}
export const textToHtml = (text) => text.split(/\r?\n\r?\n/).map((p) => `<p>${esc(p).replace(/\r?\n/g, '<br>')}</p>`).join('\n');
export function htmlToText(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script, style').forEach((n) => n.remove());
  // blocks on lines of their own
  doc.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, li, tr, br, pre, blockquote').forEach((n) => n.after(doc.createTextNode('\n')));
  return (doc.body?.textContent || '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}
/** The body of a whole page, or the HTML as it is */
function bodyOf(html) {
  const m = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
  return m ? m[1] : html;
}
export const titleOf = (html) => (/<title[^>]*>([^<]*)<\/title>/i.exec(html)?.[1] || '').trim();

// ---------------------------------------------------------------- HTML to Word
const HEADINGS = { H1: 'Heading1', H2: 'Heading2', H3: 'Heading3', H4: 'Heading4', H5: 'Heading5', H6: 'Heading6' };

function dataUrlBytes(src) {
  const m = /^data:image\/(png|jpe?g|gif|bmp);base64,(.+)$/i.exec(src || '');
  if (!m) return null;
  const bin = atob(m[2]);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return { data: out, type: m[1].toLowerCase().replace('jpeg', 'jpg') };
}
async function imageSize(bytes, type) {
  try {
    const bmp = await createImageBitmap(new Blob([bytes], { type: `image/${type === 'jpg' ? 'jpeg' : type}` }));
    const s = { w: bmp.width, h: bmp.height };
    bmp.close();
    return s;
  } catch { return { w: 400, h: 300 }; }
}

/** HTML (a whole page or a piece) as a .docx: headings, paragraphs, bold, italic, underline,
 *  strikethrough, links, lists, tables, code, quotes and embedded pictures */
export async function htmlToDocx(html, title = '') {
  const d = await import('docx');
  const { Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell, WidthType, ImageRun, AlignmentType } = d;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  doc.querySelectorAll('script, style, noscript').forEach((n) => n.remove());

  // runs for the inline content of a node, with the formatting it inherits
  const runs = async (node, fmt = {}) => {
    const out = [];
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        const text = fmt.pre ? n.textContent : n.textContent.replace(/\s+/g, ' ');
        if (!text) continue;
        if (fmt.pre) {
          text.split('\n').forEach((line, i) => out.push(new TextRun({ text: line, font: 'Consolas', break: i > 0 ? 1 : 0, size: 20 })));
        } else out.push(new TextRun({ text, bold: fmt.b, italics: fmt.i, underline: fmt.u ? {} : undefined, strike: fmt.s, font: fmt.code ? 'Consolas' : undefined, superScript: fmt.sup, subScript: fmt.sub, ...(fmt.link ? { style: 'Hyperlink' } : {}) }));
        continue;
      }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (tag === 'BR') { out.push(new TextRun({ text: '', break: 1 })); continue; }
      if (tag === 'IMG') {
        const img = dataUrlBytes(n.getAttribute('src'));
        if (img && img.type !== 'gif' && img.type !== 'bmp') {
          const { w, h } = await imageSize(img.data, img.type);
          const k = Math.min(1, 600 / w);
          out.push(new ImageRun({ data: img.data, type: img.type, transformation: { width: Math.round(w * k), height: Math.round(h * k) } }));
        }
        continue;
      }
      const next = { ...fmt };
      if (tag === 'B' || tag === 'STRONG') next.b = true;
      if (tag === 'I' || tag === 'EM') next.i = true;
      if (tag === 'U' || tag === 'INS') next.u = true;
      if (tag === 'S' || tag === 'DEL' || tag === 'STRIKE') next.s = true;
      if (tag === 'CODE' || tag === 'KBD' || tag === 'SAMP') next.code = true;
      if (tag === 'SUP') next.sup = true;
      if (tag === 'SUB') next.sub = true;
      if (tag === 'A' && /^(https?:|mailto:)/i.test(n.getAttribute('href') || '')) {
        const children = await runs(n, { ...next, link: true });
        if (children.length) out.push(new ExternalHyperlink({ link: n.getAttribute('href'), children }));
        continue;
      }
      out.push(...await runs(n, next));
    }
    return out;
  };

  const BLOCK = /^(P|DIV|H[1-6]|UL|OL|LI|TABLE|PRE|BLOCKQUOTE|HR|SECTION|ARTICLE|HEADER|FOOTER|MAIN|NAV|ASIDE|FIGURE|FIGCAPTION|DL|DT|DD|ADDRESS)$/;
  const hasBlocks = (n) => [...n.children].some((c) => BLOCK.test(c.tagName));

  const blocks = async (node, ctx = {}) => {
    const out = [];
    let inline = [];
    const flush = async () => {
      if (!inline.length) return;
      const holder = doc.createElement('span');
      inline.forEach((n) => holder.appendChild(n.cloneNode(true)));
      inline = [];
      const children = await runs(holder, ctx.fmt);
      if (children.some((r) => r)) out.push(new Paragraph({ children, ...(ctx.quote ? { indent: { left: 567 } } : {}) }));
    };
    for (const n of [...node.childNodes]) {
      if (n.nodeType === 3 || (n.nodeType === 1 && !BLOCK.test(n.tagName))) {
        if (n.nodeType === 3 && !n.textContent.trim() && !inline.length) continue;
        inline.push(n);
        continue;
      }
      if (n.nodeType !== 1) continue;
      await flush();
      const tag = n.tagName;
      if (HEADINGS[tag]) out.push(new Paragraph({ heading: HEADINGS[tag], children: await runs(n) }));
      else if (tag === 'P' || tag === 'DT' || tag === 'ADDRESS' || tag === 'FIGCAPTION') {
        const align = (n.getAttribute('align') || n.style?.textAlign || '').toLowerCase();
        out.push(new Paragraph({
          children: await runs(n, ctx.fmt),
          ...(ctx.quote ? { indent: { left: 567 } } : {}),
          ...(align === 'center' ? { alignment: AlignmentType.CENTER } : align === 'right' ? { alignment: AlignmentType.RIGHT } : {}),
        }));
      } else if (tag === 'PRE') out.push(new Paragraph({ children: await runs(n, { pre: true }), shading: { fill: 'F6F8FA' } }));
      else if (tag === 'HR') out.push(new Paragraph({ border: { bottom: { style: 'single', size: 6, color: 'BBBBBB', space: 1 } }, children: [] }));
      else if (tag === 'UL' || tag === 'OL') {
        const level = ctx.level ?? 0;
        for (const li of n.children) {
          if (li.tagName !== 'LI') continue;
          const nested = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
          const clone = li.cloneNode(true);
          clone.querySelectorAll(':scope > ul, :scope > ol').forEach((c) => c.remove());
          const children = hasBlocks(clone)
            ? await runs(Object.assign(doc.createElement('span'), { textContent: clone.textContent }))
            : await runs(clone);
          out.push(new Paragraph({
            children,
            ...(tag === 'OL' ? { numbering: { reference: 'kertas-numbers', level: Math.min(level, 8) } } : { bullet: { level: Math.min(level, 8) } }),
          }));
          for (const sub of nested) {
            const holder = doc.createElement('div');
            holder.appendChild(sub.cloneNode(true));
            out.push(...await blocks(holder, { ...ctx, level: level + 1 }));
          }
        }
      } else if (tag === 'TABLE') {
        const rows = [...n.querySelectorAll('tr')].filter((tr) => tr.closest('table') === n);
        const width = Math.max(1, ...rows.map((r) => r.children.length));
        const tableRows = [];
        for (const tr of rows) {
          const cells = [];
          for (const td of tr.children) {
            const header = td.tagName === 'TH';
            const children = await runs(td, header ? { b: true } : {});
            cells.push(new TableCell({
              children: [new Paragraph({ children })],
              columnSpan: Number(td.getAttribute('colspan')) || 1,
              ...(header ? { shading: { fill: 'F1F5F9' } } : {}),
            }));
          }
          if (cells.length) tableRows.push(new TableRow({ children: cells }));
        }
        if (tableRows.length) {
          out.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE }, columnWidths: Array(width).fill(Math.floor(9000 / width)) }));
          out.push(new Paragraph({ children: [] }));
        }
      } else if (tag === 'BLOCKQUOTE') out.push(...await blocks(n, { ...ctx, quote: true, fmt: { ...ctx.fmt, i: true } }));
      else out.push(...await blocks(n, ctx));
    }
    await flush();
    return out;
  };

  const children = await blocks(doc.body || doc.documentElement);
  const file = new Document({
    creator: 'Kertas Office',
    title: title || titleOf(html) || undefined,
    numbering: { config: [{ reference: 'kertas-numbers', levels: Array.from({ length: 9 }, (_, level) => ({ level, format: 'decimal', text: `%${level + 1}.`, alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720 * (level + 1), hanging: 360 } } } })) }] },
    styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
    sections: [{ children: children.length ? children : [new Paragraph({ children: [] })] }],
  });
  return new Uint8Array(await (await Packer.toBlob(file)).arrayBuffer());
}
