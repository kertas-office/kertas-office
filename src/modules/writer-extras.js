// Writer: shapes, text boxes, equations, citations and bibliography, mail merge.
// Built on SuperDoc's public Document API (create, insert, fields, citations, images).
import * as ooxml from '../lib/ooxml-insert.js';
import { openDialog } from '../lib/dialog.js';
import { icon } from '../lib/icons.js';
import { openFiles, saveFile, saveMany } from '../lib/platform.js';
import { SHAPES, SHAPE_COLORS, shapeIcon, shapePng } from '../lib/shapes.js';
export { SHAPES };
import { latexToOmml, EQUATION_PRESETS } from '../lib/omml.js';
import temml from 'temml';
import { CITATION_STYLES, SOURCE_TYPES, parseAuthors, formatAuthors, inText, numbering, bibliographyXml, entryText, ordered } from '../lib/citations.js';
import { readRecipients, fieldName, mergeSingle, mergeSeparate, recordLabel } from '../lib/mailmerge.js';

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
};
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const xmlEsc = (s) => String(s ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const store = {
  get(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* private mode */ } },
};

export function createExtras(ctx) {
  const { api, act, toast, caretBlock, selection, onBlock, failed, reason, exportDocx, openDocument, docName, setStatus } = ctx;

  // ---------------------------------------------------------------- shared helpers
  async function caretTextTarget() {
    const s = await selection().catch(() => null);
    const seg = s?.target?.segments?.[0];
    if (!seg) throw new Error('Click in the document where it should go');
    const at = seg.range.end ?? seg.range.start;
    return { kind: 'text', segments: [{ blockId: seg.blockId, range: { start: at, end: at } }] };
  }
  /** Puts the caret at a text offset (best effort; it is only a convenience) */
  function caretAt(blockId, offset) {
    try {
      ctx.caretTo?.(blockId, offset);
      ctx.focus?.();
    } catch { /* not available */ }
  }
  const blockItems = async () => { const r = await api().blocks.list({}); return r?.items || r?.blocks || r || []; };
  /** Inserts Word paragraphs (OOXML) after the paragraph with the caret, or at the end */
  /** Inserts Word paragraphs (OOXML) after the paragraph with the caret, then leaves the caret after them */
  async function insertParagraphsXml(xml) {
    const ed = ctx.editor?.();
    if (!ed) throw new Error('The document is still opening');
    const r = await ooxml.insertBlocksXml(ed, xml);
    try {
      const end = Math.max(r.from + 1, Math.min(r.to - 1, ed.state.doc.content.size));
      ed.commands.setTextSelection({ from: end, to: end });
      ctx.focus?.();
    } catch { /* caret placement is a convenience */ }
    return r;
  }
  const insertXmlBlocks = (xml) => ooxml.insertBlocksXml(ctx.editor(), xml);

  // ---------------------------------------------------------------- shapes
  function shapeMenu() {
    return { grid: SHAPES.map((s) => s.id), html: SHAPES.map(shapeIcon), titles: SHAPES.map((s) => s.label), columns: 7, cell: 44, onPick: (id) => insertShape(SHAPES.find((s) => s.id === id)) };
  }
  async function insertShape(shape) {
    const last = store.get('kertas-shape', { fill: SHAPE_COLORS[0].fill, stroke: SHAPE_COLORS[0].stroke, wrap: 'Square' });
    const res = await openDialog({
      title: `Insert ${shape.label.toLowerCase()}`,
      width: 460,
      fields: [
        { id: 'preview', type: 'info', html: '' },
        { id: 'fill', label: 'Fill', type: 'color', value: last.fill },
        { id: 'stroke', label: 'Outline', type: 'color', value: last.stroke },
        ...(shape.line ? [] : [{ id: 'text', label: 'Text inside (optional)', value: '', placeholder: 'e.g. Step 1' }]),
        { id: 'width', label: 'Width (cm)', type: 'number', value: shape.line ? 8 : 4, min: 0.5, max: 17, step: 0.5 },
        { id: 'height', label: 'Height (cm)', type: 'number', value: shape.line ? 1 : shape.id === 'circle' || shape.id === 'star5' ? 4 : 3, min: 0.3, max: 24, step: 0.5 },
        { id: 'wrap', label: 'Text around it', type: 'segmented', value: last.wrap, options: [{ value: 'Inline', label: 'In line' }, { value: 'Square', label: 'Wrap around' }, { value: 'None', label: 'In front' }] },
        { id: 'align', label: 'Position', type: 'segmented', value: last.align || 'right', options: [{ value: 'left', label: 'Left' }, { value: 'center', label: 'Centre' }, { value: 'right', label: 'Right' }], show: (v) => v.wrap !== 'Inline' },
      ],
      onChange: (v, form) => {
        const box = form.querySelector('[data-row="preview"] .dlg-info');
        const ratio = (Number(v.height) || 1) / (Number(v.width) || 1);
        box.style.cssText = 'display:grid;place-items:center;background:transparent;padding:4px 0 0';
        box.innerHTML = shapeSvgPreview(shape, { fill: v.fill, stroke: v.stroke, text: v.text, ratio });
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    store.set('kertas-shape', { fill: v.fill, stroke: v.stroke, wrap: v.wrap, align: v.align });
    const px = (cm) => Math.round((Number(cm) || 1) * 37.795);
    const width = px(v.width); const height = px(v.height);
    const src = await shapePng(shape, { width, height, fill: v.fill, stroke: v.stroke, text: v.text || '', textColor: contrast(v.fill), strokeWidth: shape.line ? 3 : 2 });
    const before = new Set(((await api().images.list({}))?.items || []).map((i) => i.sdImageId));
    // A clean paragraph to sit before, so the picture does not inherit a text box's frame
    let spot = null;
    try {
      const was = new Set((await blockItems()).map((x) => x.nodeId));
      await insertXmlBlocks('<w:p><w:pPr/></w:p>');
      spot = (await blockItems()).find((x) => !was.has(x.nodeId));
    } catch { /* fall back to the caret */ }
    const b = spot ? { kind: 'block', nodeType: spot.nodeType, nodeId: spot.nodeId } : await caretBlock().catch(() => null);
    const r = await act('Shape', () => api().create.image({ src, alt: shape.label, size: { width, height, unit: 'px' }, ...(b ? { at: { kind: spot ? 'before' : 'after', target: b } } : { at: { kind: 'documentEnd' } }) }));
    if (spot) caretAt(spot.nodeId, 0);
    if (!r || v.wrap === 'Inline') return;
    await new Promise((ok) => setTimeout(ok, 120));
    const added = ((await api().images.list({}))?.items || []).find((i) => !before.has(i.sdImageId));
    if (!added) return;
    await act('Shape layout', async () => {
      const f = await api().images.convertToFloating({ imageId: added.sdImageId });
      if (failed(f)) return f;
      const w = await api().images.setWrapType({ imageId: added.sdImageId, type: v.wrap });
      if (failed(w)) return w;
      // Beside the paragraph it belongs to, at the chosen side of the column
      return api().images.setPosition({ imageId: added.sdImageId, position: { hRelativeFrom: 'column', alignH: v.align || 'right', vRelativeFrom: 'paragraph', marginOffset: { top: 0 } } });
    });
  }
  function shapeSvgPreview(shape, { fill, stroke, text, ratio }) {
    const w = 200; const h = Math.max(20, Math.min(160, Math.round(w * ratio)));
    return `<svg width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><g transform="scale(${w / 100} ${h / 100})"><path d="${shape.d}" fill="${shape.line ? 'none' : esc(fill)}" stroke="${esc(stroke)}" stroke-width="${shape.line ? 3 : 2}" vector-effect="non-scaling-stroke" stroke-linejoin="round" stroke-linecap="round"/></g>${text ? `<text x="${w / 2}" y="${h / 2}" fill="${contrast(fill)}" font-family="Inter, Arial" font-weight="600" font-size="${Math.min(h * 0.18, 320 / Math.max(4, text.length))}" text-anchor="middle" dominant-baseline="central">${esc(text)}</text>` : ''}</svg>`;
  }
  function contrast(hex) {
    const n = parseInt(String(hex || '#000').slice(1), 16);
    const l = (0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
    return l > 0.6 ? '#111827' : '#FFFFFF';
  }

  // ---------------------------------------------------------------- text boxes
  const TEXTBOX_KINDS = {
    right: { label: 'Text box (right)', pos: 'right', fill: 'EEF2FE', line: '4F46E5', width: 3600 },
    left: { label: 'Text box (left)', pos: 'left', fill: 'F1F5F9', line: '64748B', width: 3600 },
    sidebar: { label: 'Sidebar note', pos: 'right', fill: 'FEF9C3', line: 'CA8A04', width: 2880 },
    quote: { label: 'Pull quote', pos: null },
  };
  function textBoxMenu() {
    return Object.entries(TEXTBOX_KINDS).map(([k, t]) => ({ label: t.label, icon: k === 'quote' ? 'quote' : 'text-box', hint: k === 'quote' ? 'A large centred quotation' : 'Text beside the main text', action: () => insertTextBox(k) }));
  }
  async function insertTextBox(kind = 'right') {
    const t = TEXTBOX_KINDS[kind];
    const res = await openDialog({
      title: t.label,
      width: 460,
      fields: [{ id: 'text', label: 'Text', type: 'textarea', rows: 4, value: '', placeholder: kind === 'quote' ? 'A sentence worth repeating' : 'Type the text for the box' }],
      validate: (v) => (v.text.trim() ? null : 'Type some text for the box'),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    const paras = res.values.text.trim().split(/\n+/);
    let xml;
    if (kind === 'quote') {
      const pPr = '<w:pPr><w:pBdr><w:top w:val="single" w:sz="12" w:space="8" w:color="4F46E5"/><w:bottom w:val="single" w:sz="12" w:space="8" w:color="4F46E5"/></w:pBdr><w:spacing w:before="240" w:after="240"/><w:ind w:left="720" w:right="720"/><w:jc w:val="center"/></w:pPr>';
      xml = paras.map((p) => `<w:p>${pPr}<w:r><w:rPr><w:i/><w:color w:val="3730A3"/><w:sz w:val="32"/></w:rPr><w:t xml:space="preserve">${xmlEsc(p)}</w:t></w:r></w:p>`).join('');
    } else {
      xml = textBoxXml(paras, t);
    }
    await act('Text box', () => insertParagraphsXml(xml), kind === 'quote' ? null : 'Text box added. Click inside it to edit; the text around it wraps.');
  }
  /** A real Word text box (DrawingML, wrap square), anchored to the paragraph at the caret */
  function textBoxXml(paras, t) {
    const EMU = 635; // per twip
    const cx = t.width * EMU;
    // Height from the text: about 13 characters per inch of width at 11pt, 0.2in per line
    const perLine = Math.max(8, Math.floor((t.width / 1440 - 0.25) * 13));
    const lines = paras.reduce((n, p) => n + Math.max(1, Math.ceil(p.length / perLine)), 0);
    const cy = Math.round((lines * 0.205 + 0.2) * 914400);
    const id = 1000 + Math.floor(Math.random() * 1e6);
    const body = paras.map((p) => `<w:p><w:pPr><w:spacing w:after="60"/></w:pPr><w:r><w:t xml:space="preserve">${xmlEsc(p)}</w:t></w:r></w:p>`).join('');
    return '<w:p><w:r><w:drawing>'
      + `<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="${251659264 + (id % 1000)}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">`
      + '<wp:simplePos x="0" y="0"/>'
      + `<wp:positionH relativeFrom="margin"><wp:align>${t.pos}</wp:align></wp:positionH>`
      + '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
      + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>`
      + `<wp:docPr id="${id}" name="Text Box ${id}"/><wp:cNvGraphicFramePr/>`
      + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
      + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/>'
      + `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
      + `<a:solidFill><a:srgbClr val="${t.fill}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="${t.line}"/></a:solidFill></a:ln></wps:spPr>`
      + `<wps:txbx><w:txbxContent>${body}</w:txbxContent></wps:txbx>`
      + '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:spAutoFit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>';
  }

  // ---------------------------------------------------------------- equations
  async function insertEquation(preset) {
    const last = store.get('kertas-eq', EQUATION_PRESETS[0].tex);
    const res = await openDialog({
      title: 'Insert equation',
      description: 'Type the equation in LaTeX, or start from an example. Word opens it as a real, editable equation.',
      width: 600,
      fields: [
        { id: 'preset', label: 'Examples', type: 'select', value: '', options: [{ value: '', label: 'Choose an example...' }, ...EQUATION_PRESETS.map((p, i) => ({ value: String(i), label: p.label }))] },
        { id: 'tex', label: 'LaTeX', type: 'textarea', rows: 3, value: preset?.tex || last },
        { id: 'preview', type: 'info', html: '' },
        { id: 'help', type: 'info', html: '<span class="muted">\\frac{a}{b} fraction, x^2 power, x_i subscript, \\sqrt{x} root, \\sum_{i=1}^{n} sum, \\int_a^b integral, \\alpha \\beta \\pi Greek, \\left( \\right) brackets</span>' },
      ],
      validate: (v) => { try { latexToOmml(v.tex); return null; } catch (e) { return `Not quite right: ${e.message}`; } },
      onChange: (v, form) => {
        const ta = form.querySelector('[data-f="tex"]');
        if (v.preset !== '' && form.dataset.preset !== v.preset) {
          form.dataset.preset = v.preset;
          ta.value = EQUATION_PRESETS[Number(v.preset)].tex;
          v.tex = ta.value;
        }
        const box = form.querySelector('[data-row="preview"] .dlg-info');
        box.classList.add('eq-preview');
        try { box.innerHTML = temml.renderToString(v.tex, { displayMode: true, throwOnError: true }); } catch (e) { box.innerHTML = `<span class="muted">${esc(String(e.message).replace(/^Temml parse error:\s*/i, ''))}</span>`; }
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    store.set('kertas-eq', res.values.tex);
    const { paragraph } = latexToOmml(res.values.tex);
    await act('Equation', () => insertParagraphsXml(paragraph));
  }

  // ---------------------------------------------------------------- citations
  let citeStyle = store.get('kertas-cite-style', 'apa');
  const sourcesList = async () => ((await api().citations.sources.list({}))?.items || []).map((s) => ({ sourceId: s.sourceId, type: s.type, fields: s.fields || {} }));
  // In-text citations are plain text, so typing around them behaves normally. A small
  // registry saved inside the .docx (a custom XML part) remembers which sources each
  // rendered citation stands for, so a change of style can find and rewrite them.
  const REG_NS = 'urn:kertas:citations:1';
  const xmlText = (t) => String(t).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const unXml = (t) => String(t).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
  async function readRegistry() {
    try {
      const l = await api().customXml.parts.list({ rootNamespace: REG_NS });
      const part = (l?.items || l?.parts || [])[0];
      if (!part) return { target: null, entries: [] };
      const target = part.id ? { id: part.id } : { partName: part.partName };
      const info = await api().customXml.parts.get({ target });
      const body = String(info?.content ?? info?.xml ?? '').match(/<data>([\s\S]*?)<\/data>/)?.[1];
      return { target, entries: body ? JSON.parse(unXml(body)) : [] };
    } catch { return { target: null, entries: [] }; }
  }
  async function writeRegistry(reg) {
    const content = `<citations xmlns="${REG_NS}"><data>${xmlText(JSON.stringify(reg.entries))}</data></citations>`;
    try {
      if (reg.target) await api().customXml.parts.patch({ target: reg.target, content });
      else await api().customXml.parts.create({ content });
    } catch (e) { console.warn('citation registry', e); }
  }
  const sameEntry = (a, b) => a.text === b.text && a.page === b.page && a.sourceIds.join() === b.sourceIds.join();
  async function remember(entry) {
    const reg = await readRegistry();
    if (!reg.entries.some((e) => sameEntry(e, entry))) { reg.entries.push(entry); await writeRegistry(reg); }
  }
  async function bodyBlocks() {
    const r = await api().blocks.list({ includeText: true });
    return (r?.blocks || r?.items || []).map((x) => ({ id: x.nodeId, ordinal: x.ordinal, text: String(x.text ?? x.textPreview ?? '') }));
  }
  /** Every place a registered citation appears: [{ entry, blockId, ordinal, start, end }] in reading order */
  function occurrences(entries, blocks) {
    const out = [];
    for (const b of blocks) {
      for (const entry of entries) {
        if (!entry.text) continue;
        for (let k = b.text.indexOf(entry.text); k >= 0; k = b.text.indexOf(entry.text, k + entry.text.length)) {
          out.push({ entry, blockId: b.id, ordinal: b.ordinal, start: k, end: k + entry.text.length });
        }
      }
    }
    return out.sort((x, y) => x.ordinal - y.ordinal || x.start - y.start);
  }
  async function citedOrder() {
    const { entries } = await readRegistry();
    const ids = [];
    for (const o of occurrences(entries, await bodyBlocks())) for (const sid of o.entry.sourceIds) if (!ids.includes(sid)) ids.push(sid);
    return ids;
  }
  async function blockText(blockId) {
    return (await bodyBlocks()).find((x) => x.id === blockId)?.text || '';
  }
  const point = (blockId, offset) => ({ kind: 'text', blockId, offset });
  const range = (blockId, from, to) => ({ kind: 'selection', start: point(blockId, from), end: point(blockId, to) });

  async function sourceDialog(src) {
    const f = src?.fields || {};
    const res = await openDialog({
      title: src ? 'Edit source' : 'New source',
      width: 540,
      fields: [
        { id: 'type', label: 'Type', type: 'select', value: src?.type || 'book', options: SOURCE_TYPES },
        { id: 'authors', label: 'Authors', value: formatAuthors(f.authors), placeholder: 'Newport, Cal; Kaplan, Rachel', hint: 'Last name, first name. Separate people with a semicolon.' },
        { id: 'title', label: 'Title', value: f.title || '' },
        { id: 'year', label: 'Year', value: f.year || '', placeholder: '2024' },
        { id: 'journalName', label: 'Journal', value: f.journalName || '', show: (v) => v.type === 'journalArticle' || v.type === 'conferenceProceedings' },
        { id: 'volume', label: 'Volume', value: f.volume || '', show: (v) => v.type === 'journalArticle' },
        { id: 'issue', label: 'Issue', value: f.issue || '', show: (v) => v.type === 'journalArticle' },
        { id: 'pages', label: 'Pages', value: f.pages || '', placeholder: '169-182', show: (v) => v.type === 'journalArticle' || v.type === 'conferenceProceedings' },
        { id: 'publisher', label: 'Publisher or website name', value: f.publisher || '', show: (v) => v.type !== 'journalArticle' },
        { id: 'city', label: 'City', value: f.city || '', show: (v) => v.type === 'book' || v.type === 'report' || v.type === 'thesis' },
        { id: 'edition', label: 'Edition', value: f.edition || '', placeholder: '2nd', show: (v) => v.type === 'book' },
        { id: 'url', label: 'Web address', value: f.url || '', placeholder: 'https://' },
        { id: 'doi', label: 'DOI', value: f.doi || '', placeholder: '10.1000/xyz123', show: (v) => v.type === 'journalArticle' || v.type === 'conferenceProceedings' || v.type === 'report' },
        { id: 'preview', type: 'info', html: '' },
      ],
      validate: (v) => (v.title.trim() || v.authors.trim() ? null : 'Give at least a title or an author'),
      onChange: (v, form) => {
        const s = { sourceId: 'x', type: v.type, fields: toFields(v) };
        form.querySelector('[data-row="preview"] .dlg-info').innerHTML = `<span class="muted">${esc(CITATION_STYLES.find((c) => c.id === citeStyle).label)}:</span> ${esc(entryText(s, citeStyle, 1)).replace(/\t/, ' ')}`;
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: src ? 'Save' : 'Add source', primary: true }],
    });
    return res ? { type: res.values.type, fields: toFields(res.values) } : null;
  }
  function toFields(v) {
    const f = { title: v.title.trim(), authors: parseAuthors(v.authors), year: v.year.trim() };
    for (const k of ['journalName', 'volume', 'issue', 'pages', 'publisher', 'city', 'edition', 'url', 'doi']) if (v[k]?.trim?.()) f[k] = v[k].trim();
    return f;
  }

  async function insertCitation() {
    let target;
    try { target = await caretTextTarget(); } catch (e) { toast(e.message, true); return; }
    let list = await sourcesList();
    if (!list.length) {
      const s = await sourceDialog();
      if (!s) return;
      const r = await act('Source', () => api().citations.sources.insert(s));
      if (!r) return;
      list = await sourcesList();
    }
    const res = await openDialog({
      title: 'Insert citation',
      width: 560,
      fields: [
        { id: 'style', label: 'Style', type: 'select', value: citeStyle, options: CITATION_STYLES.map((s) => ({ value: s.id, label: s.label })) },
        { id: 'head', type: 'info', html: '<b>Sources to cite</b>' },
        ...ordered(list, 'apa').map((s) => ({ id: `s-${s.sourceId}`, type: 'checkbox', label: entryText(s, 'apa').slice(0, 110), value: false })),
        { id: 'page', label: 'Page (optional)', value: '', placeholder: '12 or 12-14' },
      ],
      validate: (v, a) => (a === 'new' || list.some((s) => v[`s-${s.sourceId}`]) ? null : 'Tick at least one source'),
      actions: [{ id: 'new', label: 'New source...' }, { id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    if (res.action === 'new') {
      const s = await sourceDialog();
      if (s && await act('Source', () => api().citations.sources.insert(s))) insertCitation();
      return;
    }
    if (res.values.style !== citeStyle) { citeStyle = res.values.style; store.set('kertas-cite-style', citeStyle); }
    const chosen = list.filter((s) => res.values[`s-${s.sourceId}`]);
    const page = res.values.page.trim();
    const cited = await citedOrder();
    for (const s of chosen) if (!cited.includes(s.sourceId)) cited.push(s.sourceId);
    const text = inText(chosen, citeStyle, { numbers: numbering(list, cited), page });
    const seg = target.segments[0];
    const at = seg.range.start;
    const before = (await blockText(seg.blockId)).slice(0, at);
    const lead = before && !/\s$/.test(before) ? ' ' : '';
    const ok = await act('Citation', () => api().insert({ value: lead + text, target: range(seg.blockId, at, at) }));
    if (!ok) return;
    await remember({ text, sourceIds: chosen.map((s) => s.sourceId), page });
    caretAt(seg.blockId, at + lead.length + text.length);
    if (citeStyle === 'ieee') await updateCitations(true);
  }

  /** Re-renders every citation in the current style (and IEEE numbers) */
  async function updateCitations(quiet = false) {
    const list = await sourcesList();
    const byId = new Map(list.map((s) => [s.sourceId, s]));
    const reg = await readRegistry();
    const found = occurrences(reg.entries, await bodyBlocks());
    const cited = [];
    for (const o of found) for (const sid of o.entry.sourceIds) if (!cited.includes(sid)) cited.push(sid);
    const nums = numbering(list, cited);
    const fresh = new Map(reg.entries.map((e) => {
      const srcs = e.sourceIds.map((id) => byId.get(id)).filter(Boolean);
      return [e, srcs.length ? inText(srcs, citeStyle, { numbers: nums, page: e.page }) : e.text];
    }));
    let n = 0;
    // From the end backwards, so earlier offsets in a paragraph stay valid
    for (const o of [...found].reverse()) {
      const text = fresh.get(o.entry);
      if (text === o.entry.text) continue;
      const r = await api().replace({ target: range(o.blockId, o.start, o.end), text });
      if (!failed(r)) n++;
    }
    const entries = [];
    for (const e of reg.entries) { const next = { ...e, text: fresh.get(e) }; if (!entries.some((x) => sameEntry(x, next))) entries.push(next); }
    await writeRegistry({ target: reg.target, entries });
    if (n) ctx.setDirty(true);
    if (!quiet) toast(n ? `Updated ${n} citation${n === 1 ? '' : 's'}` : 'Citations are up to date');
  }

  async function manageSources() {
    const list = await sourcesList();
    const res = await openDialog({
      title: 'Sources',
      description: list.length ? 'The sources saved in this document. Word shows them in its Source Manager too.' : 'No sources yet.',
      width: 600,
      fields: [
        { id: 'style', label: 'Citation style', type: 'select', value: citeStyle, options: CITATION_STYLES.map((s) => ({ value: s.id, label: s.label })) },
        ...ordered(list, citeStyle).map((s, k) => ({ id: `pick-${s.sourceId}`, type: 'checkbox', label: entryText(s, citeStyle, k + 1).replace(/^\[\d+\]\t/, '').slice(0, 140), value: false })),
      ],
      actions: [
        { id: 'add', label: 'New' }, { id: 'edit', label: 'Edit ticked' }, { id: 'remove', label: 'Remove ticked', danger: true },
        { id: 'done', label: 'Close', primary: true },
      ],
    });
    if (!res) return;
    if (res.values.style !== citeStyle) {
      citeStyle = res.values.style; store.set('kertas-cite-style', citeStyle);
      await updateCitations(true);
      toast(`Citations now use ${CITATION_STYLES.find((s) => s.id === citeStyle).label}. Insert the bibliography again to match.`);
    }
    const picked = list.filter((s) => res.values[`pick-${s.sourceId}`]);
    if (res.action === 'add') {
      const s = await sourceDialog();
      if (s) await act('Source', () => api().citations.sources.insert(s), 'Source added');
      return manageSources();
    }
    if (res.action === 'edit' && picked[0]) {
      const s = await sourceDialog(picked[0]);
      if (s) { await act('Source', () => api().citations.sources.update({ target: { kind: 'entity', entityType: 'citationSource', sourceId: picked[0].sourceId }, patch: s })); await updateCitations(true); }
      return manageSources();
    }
    if (res.action === 'remove' && picked.length) {
      for (const s of picked) await act('Remove source', () => api().citations.sources.remove({ target: { kind: 'entity', entityType: 'citationSource', sourceId: s.sourceId } }));
      return manageSources();
    }
    return undefined;
  }

  function bibliographyMenu() {
    return [
      { header: 'Insert bibliography' },
      ...CITATION_STYLES.map((s) => ({ label: s.label, hint: s.heading, checked: s.id === citeStyle, action: () => insertBibliography(s.id) })),
      'sep',
      { label: 'Update citations', icon: 'refresh', action: () => updateCitations() },
    ];
  }
  async function insertBibliography(style = citeStyle) {
    const list = await sourcesList();
    if (!list.length) { toast('Add a source first (Insert citation or Sources)'); return; }
    if (style !== citeStyle) { citeStyle = style; store.set('kertas-cite-style', style); await updateCitations(true); }
    await act('Bibliography', async () => insertParagraphsXml(bibliographyXml(list, style, await citedOrder())));
  }

  // ---------------------------------------------------------------- mail merge
  let recipients = null; // { columns, rows, sheet, file }
  let currentRecord = 0;
  async function selectRecipients() {
    const [file] = await openFiles('sheets', false);
    if (!file) return;
    try {
      const r = readRecipients(file.data);
      if (!r.rows.length) { toast('That file has no rows under its headings', true); return; }
      recipients = { ...r, file: file.name };
      currentRecord = 0;
      ctx.onRecipients?.(recipients);
      toast(`${r.rows.length} recipient${r.rows.length === 1 ? '' : 's'} from ${file.name} (${r.columns.length} fields)`);
    } catch (e) {
      toast(`Could not read ${file.name}: ${e.message}`, true);
    }
  }
  function mergeFieldMenu() {
    if (!recipients) return [{ label: 'Select recipients first', icon: 'sheet', action: () => selectRecipients() }, 'sep', { label: 'Type a field name...', icon: 'mail', action: () => customField() }];
    return [{ header: recipients.file }, ...recipients.columns.map((c) => ({ label: c, icon: 'mail', action: () => insertMergeField(fieldName(c)) })), 'sep', { label: 'Type a field name...', icon: 'mail', action: () => customField() }];
  }
  /** Merge field from the slash menu: pick a column if recipients are loaded */
  async function promptMergeField() {
    if (!recipients) return customField();
    const res = await openDialog({ title: 'Insert merge field', fields: [{ id: 'col', label: `Field from ${recipients.file}`, type: 'select', value: recipients.columns[0], options: recipients.columns.map((c) => ({ value: c, label: c })) }], actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }] });
    if (res) insertMergeField(fieldName(res.values.col));
  }
  async function customField() {
    const res = await openDialog({ title: 'Merge field', fields: [{ id: 'name', label: 'Field name', value: '', placeholder: 'FirstName' }], validate: (v) => (fieldName(v.name) && v.name.trim() ? null : 'Type a name'), actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }] });
    if (res) insertMergeField(fieldName(res.values.name));
  }
  async function insertMergeField(name) {
    const ed = ctx.editor?.();
    if (!ed) return;
    const shown = `«${name}»`;
    const safe = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;');
    // A complex field (begin / instruction / separate / result / end), exactly as Word writes it.
    // The invisible character after it gives the caret a place outside the field,
    // so the next words typed are ordinary text.
    const xml = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      + `<w:r><w:instrText xml:space="preserve"> MERGEFIELD ${safe(name)} \\* MERGEFORMAT </w:instrText></w:r>`
      + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      + `<w:r><w:t>${safe(shown)}</w:t></w:r>`
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>\u200B</w:t></w:r></w:p>';
    const r = await act('Merge field', () => ooxml.insertInlineXml(ed, xml));
    if (r) {
      try { ed.commands.setTextSelection({ from: r.to, to: r.to }); ctx.focus?.(); } catch { /* fine */ }
    }
  }
  function mergeValue(record, name) {
    const key = Object.keys(record).find((k) => fieldName(k).toLowerCase() === name.toLowerCase() || k.toLowerCase() === name.toLowerCase());
    return key ? record[key] : '';
  }
  async function previewMerge() {
    if (!recipients) { await selectRecipients(); if (!recipients) return; }
    let html;
    try { html = await api().getHtml({}); } catch (e) { toast(`Preview is not available: ${e.message}`, true); return; }
    html = String(html?.html ?? html);
    const overlay = document.createElement('div');
    overlay.className = 'modal-backdrop';
    overlay.innerHTML = `<div class="modal glass merge-preview"><div class="mp-head"><h3>${icon('mails', 18)}Preview</h3><div class="mp-nav"><button class="chip" data-k="first" title="First">|‹</button><button class="chip" data-k="prev">‹ Previous</button><span class="mp-count"></span><button class="chip" data-k="next">Next ›</button><button class="chip" data-k="last" title="Last">›|</button></div><button class="chip" data-k="close">Close</button></div><div class="mp-page"></div></div>`;
    document.body.appendChild(overlay);
    const page = overlay.querySelector('.mp-page');
    const render = () => {
      const rec = recipients.rows[currentRecord];
      page.innerHTML = html.replace(/«([^»<]{1,80})»/g, (m, n) => `<mark class="mp-val">${esc(mergeValue(rec, n)) || '<i>(blank)</i>'}</mark>`);
      overlay.querySelector('.mp-count').textContent = `${currentRecord + 1} of ${recipients.rows.length}: ${recordLabel(rec, recipients.columns, currentRecord)}`;
    };
    render();
    const close = () => { overlay.remove(); document.removeEventListener('keydown', keys, true); };
    const keys = (e) => { if (e.key === 'Escape') close(); if (e.key === 'ArrowRight') step(1); if (e.key === 'ArrowLeft') step(-1); };
    const step = (d) => { currentRecord = Math.max(0, Math.min(recipients.rows.length - 1, currentRecord + d)); render(); };
    document.addEventListener('keydown', keys, true);
    overlay.addEventListener('click', (e) => {
      const k = e.target.closest('[data-k]')?.dataset.k;
      if (e.target === overlay || k === 'close') close();
      if (k === 'prev') step(-1);
      if (k === 'next') step(1);
      if (k === 'first') step(-1e9);
      if (k === 'last') step(1e9);
    });
  }
  function finishMenu() {
    return [
      { label: 'One document with every letter', icon: 'doc', hint: 'Opens in a new tab, each letter on its own page', action: () => finishMerge('tab') },
      { label: 'Save one Word file', icon: 'save', hint: 'Every letter in a single .docx', action: () => finishMerge('file') },
      { label: 'Save a file per recipient', icon: 'files', hint: 'Named after each recipient', action: () => finishMerge('separate') },
    ];
  }
  async function finishMerge(mode) {
    if (!recipients) { await selectRecipients(); if (!recipients) return; }
    const rows = recipients.rows;
    setStatus?.(`Merging ${rows.length} letters...`);
    try {
      const docx = await exportDocx();
      const base = String(docName || 'Letters.docx').replace(/\.docx$/i, '');
      if (mode === 'separate') {
        const files = await mergeSeparate(docx, rows);
        const used = new Set();
        const out = files.map(({ data, record }, i) => {
          let n = `${base} - ${recordLabel(record, recipients.columns, i)}`;
          while (used.has(n)) n += ` ${i + 1}`;
          used.add(n);
          return { name: `${n}.docx`, data };
        });
        const res = await saveMany(out);
        if (res) toast(`Saved ${out.length} letters${res.folder ? ` to ${res.folder}` : ''}`);
        return;
      }
      const data = await mergeSingle(docx, rows);
      if (mode === 'tab') { await openDocument?.('writer', { name: `${base} (merged).docx`, data }); return; }
      const res = await saveFile({ data, name: `${base} (merged).docx`, kind: 'writer', saveAs: true });
      if (res) toast(`Saved ${rows.length} letters in one file`);
    } catch (e) {
      console.error(e);
      toast(`Merge failed: ${e.message}`, true);
    } finally {
      setStatus?.('');
    }
  }

  return {
    shapeMenu, insertShape, textBoxMenu, insertTextBox, insertEquation, promptMergeField,
    insertCitation, manageSources, bibliographyMenu, insertBibliography, updateCitations,
    /** Registers citations that came in as Word citation fields (now plain text) */
    async registerCitations(found) {
      const reg = await readRegistry();
      let added = 0;
      for (const f of found) {
        const entry = { text: f.text, sourceIds: f.sourceIds, page: f.page || '' };
        if (entry.text && !reg.entries.some((e) => sameEntry(e, entry))) { reg.entries.push(entry); added++; }
      }
      if (added) await writeRegistry(reg);
      return added;
    },
    selectRecipients, mergeFieldMenu, previewMerge, finishMenu, finishMerge,
    get recipients() { return recipients; },
  };
}
