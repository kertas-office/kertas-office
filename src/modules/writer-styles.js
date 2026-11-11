// Paragraph styles: list every style in the document, and create a new one from the
// formatting at the cursor (Word: Styles > Create a Style). The style is written as real
// Word XML (styles.xml), so it travels with the document.

/** Paragraph styles in the document, quick styles first */
export function paragraphStyles(ed) {
  const all = ed?.converter?.translatedLinkedStyles?.styles || {};
  return Object.values(all)
    .filter((s) => s?.type === 'paragraph' && !s.semiHidden && !s.hidden && s.styleId)
    .map((s) => ({ id: s.styleId, name: displayName(s.name || s.styleId), quick: !!s.qFormat, priority: s.uiPriority ?? 99 }))
    .sort((a, b) => (b.quick - a.quick) || (a.priority - b.priority) || a.name.localeCompare(b.name));
}
const displayName = (n) => String(n).replace(/^heading (\d)$/i, 'Heading $1').replace(/^(\w)/, (c) => c.toUpperCase());

/** Formatting at the cursor: run (font, size, colour, bold, italic) and paragraph (spacing, alignment, indent) */
function formattingAtCursor(ed) {
  const $f = ed.state.selection.$from;
  let para = null; let run = null;
  for (let d = $f.depth; d > 0; d--) { const n = $f.node(d); if (!run && n.type.name === 'run') run = n; if (n.isTextblock) { para = n; break; } }
  const marks = $f.marks();
  const ts = marks.find((m) => m.type.name === 'textStyle')?.attrs || {};
  const rp = run?.attrs?.runProperties || {};
  const pp = para?.attrs?.paragraphProperties || {};
  const family = String(ts.fontFamily || '').split(',')[0].replace(/["']/g, '').trim() || null;
  const sizePt = ts.fontSize ? parseFloat(ts.fontSize) : (rp.fontSize ? rp.fontSize / 2 : null);
  const color = (ts.color || (rp.color?.val ? `#${rp.color.val}` : '') || '').replace('#', '') || null;
  return {
    family, sizePt, color: color && color !== 'auto' ? color.toUpperCase() : null,
    bold: marks.some((m) => m.type.name === 'bold') || !!rp.bold,
    italic: marks.some((m) => m.type.name === 'italic') || !!rp.italic,
    underline: marks.some((m) => m.type.name === 'underline') || !!rp.underline,
    spacing: pp.spacing || null, justification: pp.justification || null, indent: pp.indent || null,
    basedOn: pp.styleId || 'Normal',
  };
}

const attr = (o) => Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined && v !== null).map(([k, v]) => [k, String(v)]));
const el = (name, attributes, elements) => ({ type: 'element', name, ...(attributes ? { attributes } : {}), ...(elements ? { elements } : {}) });

export function styleIdFor(ed, name) {
  const base = String(name).replace(/[^A-Za-z0-9]/g, '') || 'Style';
  const all = ed.converter.translatedLinkedStyles?.styles || {};
  let id = base; let i = 2;
  while (all[id]) id = `${base}${i++}`;
  return id;
}

/** Creates a paragraph style from the formatting at the cursor and returns its id */
export function createStyleFromCursor(ed, name) {
  const c = ed.converter;
  if (!c?.convertedXml) throw new Error('The document is still opening');
  const f = formattingAtCursor(ed);
  const id = styleIdFor(ed, name);
  const sp = f.spacing || {};
  // ---- Word XML in styles.xml
  const pPr = [];
  if (sp.before != null || sp.after != null || sp.line != null) pPr.push(el('w:spacing', attr({ 'w:before': sp.before, 'w:after': sp.after, 'w:line': sp.line, 'w:lineRule': sp.lineRule })));
  if (f.indent && (f.indent.left || f.indent.right || f.indent.firstLine || f.indent.hanging)) pPr.push(el('w:ind', attr({ 'w:left': f.indent.left, 'w:right': f.indent.right, 'w:firstLine': f.indent.firstLine, 'w:hanging': f.indent.hanging })));
  if (f.justification) pPr.push(el('w:jc', { 'w:val': f.justification }));
  const rPr = [];
  if (f.family) rPr.push(el('w:rFonts', { 'w:ascii': f.family, 'w:hAnsi': f.family, 'w:cs': f.family }));
  if (f.bold) rPr.push(el('w:b'));
  if (f.italic) rPr.push(el('w:i'));
  if (f.underline) rPr.push(el('w:u', { 'w:val': 'single' }));
  if (f.color) rPr.push(el('w:color', { 'w:val': f.color }));
  if (f.sizePt) { const hp = String(Math.round(f.sizePt * 2)); rPr.push(el('w:sz', { 'w:val': hp }), el('w:szCs', { 'w:val': hp })); }
  const xml = el('w:style', { 'w:type': 'paragraph', 'w:customStyle': '1', 'w:styleId': id }, [
    el('w:name', { 'w:val': name }), el('w:basedOn', { 'w:val': f.basedOn }), el('w:next', { 'w:val': 'Normal' }), el('w:qFormat'),
    ...(pPr.length ? [el('w:pPr', null, pPr)] : []), ...(rPr.length ? [el('w:rPr', null, rPr)] : []),
  ]);
  const root = c.convertedXml['word/styles.xml']?.elements?.find((e) => e.name === 'w:styles');
  if (!root) throw new Error('This document has no style sheet');
  root.elements = root.elements || [];
  root.elements.push(xml);
  // ---- the editor's two parsed copies
  const runProperties = {};
  if (f.family) runProperties.fontFamily = { ascii: f.family, hAnsi: f.family, cs: f.family };
  if (f.bold) runProperties.bold = true;
  if (f.italic) runProperties.italic = true;
  if (f.underline) runProperties.underline = { val: 'single' };
  if (f.color) runProperties.color = { val: f.color };
  if (f.sizePt) { runProperties.fontSize = Math.round(f.sizePt * 2); runProperties.fontSizeCs = runProperties.fontSize; }
  const paragraphProperties = {};
  if (f.spacing) paragraphProperties.spacing = { ...f.spacing };
  if (f.indent) paragraphProperties.indent = { ...f.indent };
  if (f.justification) paragraphProperties.justification = f.justification;
  c.translatedLinkedStyles.styles[id] = { type: 'paragraph', styleId: id, name, basedOn: f.basedOn, next: 'Normal', qFormat: true, customStyle: true, uiPriority: 1, paragraphProperties, runProperties };
  const css = {};
  if (f.family) css['font-family'] = `${f.family}, sans-serif`;
  if (f.sizePt) css['font-size'] = `${f.sizePt}pt`;
  if (f.color) css.color = `#${f.color}`;
  if (f.bold) css['font-weight'] = 'bold';
  if (f.italic) css['font-style'] = 'italic';
  c.linkedStyles = c.linkedStyles || [];
  c.linkedStyles.push({ id, type: 'paragraph', definition: { attrs: { name, basedOn: f.basedOn, qFormat: true }, styles: { spacing: {}, indent: {}, ...css } }, attributes: {} });
  return id;
}
