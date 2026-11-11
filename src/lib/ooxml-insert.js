// Inserting Word XML (OOXML) into the open document.
// The snippet is wrapped in a tiny .docx and read by the editor's own importer,
// the same code that opens documents, so anything Word can express (equations,
// framed text boxes, fields, styled paragraphs) arrives as real editor content.
import JSZip from 'jszip';
import { Editor, BLANK_DOCX_BASE64 } from 'superdoc/super-editor';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  w14: 'http://schemas.microsoft.com/office/word/2010/wordml',
};

let template = null;
async function blankParts() {
  if (!template) {
    const zip = await JSZip.loadAsync(BLANK_DOCX_BASE64, { base64: true });
    template = { zip, doc: await zip.file('word/document.xml').async('string') };
  }
  return template;
}

/** Body XML (one or more <w:p>/<w:tbl>) -> ProseMirror JSON nodes */
export async function ooxmlToNodes(bodyXml) {
  const { zip, doc } = await blankParts();
  let xml = doc;
  for (const [p, uri] of Object.entries(NS)) if (!new RegExp(`xmlns:${p}=`).test(xml)) xml = xml.replace(/<w:document\b/, `<w:document xmlns:${p}="${uri}"`);
  const open = xml.indexOf('>', xml.indexOf('<w:body')) + 1;
  const sect = xml.lastIndexOf('<w:sectPr');
  const close = xml.lastIndexOf('</w:body>');
  xml = xml.slice(0, open) + bodyXml + (sect > open ? xml.slice(sect, close) : '') + xml.slice(close);
  const copy = new JSZip();
  for (const [name, f] of Object.entries(zip.files)) if (!f.dir) copy.file(name, await f.async('uint8array'));
  copy.file('word/document.xml', xml);
  const blob = await copy.generateAsync({ type: 'blob' });
  const ed = await Editor.open(blob, { isHeadless: true, mode: 'docx', telemetry: { enabled: false } });
  try {
    const own = { fontFamily: /<w:rFonts\b/.test(bodyXml), fontSize: /<w:sz\b/.test(bodyXml) };
    return (ed.getJSON().content || []).map((n) => stripInherited(n, null, own));
  } finally {
    try { ed.destroy(); } catch { /* already gone */ }
  }
}

// The helper document's defaults (Arial 12) must not travel with the snippet:
// keep only the run formatting the snippet itself states, so the text takes the
// document's own font where it lands.
const freshParaId = () => (Math.floor(Math.random() * 0x7ffffffe) + 1).toString(16).toUpperCase().padStart(8, '0');
const FONT_MARK_KEYS = { fontFamily: 'fontFamily', fontSize: 'fontSize', eastAsiaFontFamily: 'fontFamily', csFontFamily: 'fontFamily' };
function stripInherited(node, inlineKeys = null, own = {}) {
  if (!node || typeof node !== 'object') return node;
  const out = { ...node };
  let keys = inlineKeys;
  // Fresh identities, so blocks never collide with ones already in the document
  if (node.attrs && ('paraId' in node.attrs || 'sdBlockId' in node.attrs)) {
    out.attrs = { ...node.attrs };
    if ('paraId' in out.attrs) out.attrs.paraId = freshParaId();
    if ('textId' in out.attrs) out.attrs.textId = null;
    if ('sdBlockId' in out.attrs) out.attrs.sdBlockId = crypto.randomUUID();
  }
  if (node.type === 'run' && node.attrs) {
    // A property counts as the snippet's own only when it is both listed and actually present
    const listed = new Set(node.attrs.runPropertiesInlineKeys || []);
    const rp = {};
    for (const [k, v] of Object.entries(node.attrs.runProperties || {})) {
      if (!listed.has(k)) continue;
      if ((k === 'fontFamily' && !own.fontFamily) || ((k === 'fontSize' || k === 'fontSizeCs') && !own.fontSize)) continue;
      rp[k] = v;
    }
    keys = new Set(Object.keys(rp));
    out.attrs = { ...(out.attrs || node.attrs), runProperties: rp, runPropertiesInlineKeys: [...keys] };
  }
  if (Array.isArray(node.marks) && keys) {
    out.marks = node.marks.map((m) => {
      if (m.type !== 'textStyle') return m;
      const attrs = { ...m.attrs };
      for (const [a, k] of Object.entries(FONT_MARK_KEYS)) if (!keys.has(k)) attrs[a] = null;
      return { ...m, attrs };
    });
  }
  if (Array.isArray(node.content)) out.content = node.content.map((c) => stripInherited(c, keys, own));
  return out;
}

/** Position just after the block that holds the cursor (or the end of the document) */
export function afterCaretBlock(editor) {
  const { state } = editor;
  const $from = state.selection.$from;
  for (let d = $from.depth; d > 0; d--) {
    const node = $from.node(d);
    if (node.isBlock && $from.node(d - 1).type.name === 'doc') return $from.after(d);
  }
  return state.doc.content.size;
}

/** Inserts block XML after the cursor's paragraph; returns the inserted range */
export async function insertBlocksXml(editor, bodyXml) {
  const nodes = await ooxmlToNodes(bodyXml);
  if (!nodes.length) throw new Error('Nothing to insert');
  const pos = afterCaretBlock(editor);
  const before = editor.state.doc.content.size;
  editor.commands.insertContentAt(pos, nodes);
  const end = pos + (editor.state.doc.content.size - before);
  return { from: pos, to: end };
}

/** Inserts the inline content of a one-paragraph XML snippet at the cursor */
export async function insertInlineXml(editor, paragraphXml) {
  const nodes = await ooxmlToNodes(paragraphXml);
  const caretStyle = editor.state.selection.$from.marks().find((m) => m.type.name === 'textStyle')?.attrs;
  const adopt = (n) => {
    const o = { ...n };
    if (Array.isArray(n.marks) && caretStyle) o.marks = n.marks.map((m) => (m.type === 'textStyle' ? { ...m, attrs: Object.fromEntries(Object.entries(m.attrs || {}).map(([k, v]) => [k, v ?? caretStyle[k] ?? null])) } : m));
    else if (n.type === 'text' && caretStyle) o.marks = [{ type: 'textStyle', attrs: { ...caretStyle } }];
    if (Array.isArray(n.content)) o.content = n.content.map(adopt);
    return o;
  };
  const inline = (nodes[0]?.content || []).map(adopt);
  if (!inline.length) throw new Error('Nothing to insert');
  const pos = editor.state.selection.from;
  const before = editor.state.doc.content.size;
  editor.commands.insertContentAt(pos, inline);
  return { from: pos, to: pos + (editor.state.doc.content.size - before) };
}
