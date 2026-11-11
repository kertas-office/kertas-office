// Writer: .docx editing powered by SuperDoc (AGPL-3.0).
// SuperDoc's own toolbar handles fonts, colours, lists and tables; the ribbon above it
// adds the rest of Word's tabs on top of SuperDoc's Document API.
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import { createRibbon, openMenu, closeOpenMenu } from '../lib/ribbon.js';
import { icon } from '../lib/icons.js';
import { openFiles, editCommand } from '../lib/platform.js';
import { superdocFonts } from '../lib/font-assets.js';
import { createSpellProvider, personalWords, addToPersonal } from '../lib/spell/provider.js';
import { LANGUAGES as SPELL_LANGUAGES } from '../lib/spell/words.js';
import { openDialog } from '../lib/dialog.js';
import { createExtras, SHAPES } from './writer-extras.js';
import { createSlashMenu } from './writer-slash.js';
import * as ooxml from '../lib/ooxml-insert.js';
import { enableClickableBoxes } from './writer-checkbox.js';
import { createDecorations } from './writer-decorations.js';
import { tableContext, moveRow, moveColumn, tableNodeId, sortRows } from './writer-table.js';
import { paragraphStyles, createStyleFromCursor } from './writer-styles.js';
import { TABLE_DESIGNS, ensureTableStyle, designPreview } from './writer-table-styles.js';
import { enablePictureDrag } from './writer-picture-drag.js';
import { fieldsToCitations, citationsToFields } from '../lib/docx-citations.js';
import { repairDocx } from '../lib/docx-repair.js';
import { prepareDocx } from '../lib/docx-prepare.js';
import { userName, setUserName, editorUser } from '../lib/user.js';
import {
  pageNumberDialog, marginsDialog, paperSizeDialog, paragraphDialog, tableDialog, linkDialog, noteTextDialog,
  captionDialog, dateTimeDialog, wordCountDialog, watermarkTextDialog, restrictDialog, commentDialog,
} from './writer-dialogs.js';

let seq = 0;
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const CM = 567; // twips per centimetre
const PT = 20; // twips per point

const STYLE_TILES = [
  { id: 'Normal', label: 'Normal', style: 'font-size:13px' },
  { id: 'NoSpacing', label: 'No Spacing', style: 'font-size:13px' },
  { id: 'Title', label: 'Title', style: 'font-size:19px;font-weight:300;letter-spacing:-.02em' },
  { id: 'Subtitle', label: 'Subtitle', style: 'font-size:13px;color:#64748b;font-style:italic' },
  { id: 'Heading1', label: 'Heading 1', style: 'font-size:15px;color:#2f5496;font-weight:600' },
  { id: 'Heading2', label: 'Heading 2', style: 'font-size:13.5px;color:#2f5496;font-weight:600' },
  { id: 'Heading3', label: 'Heading 3', style: 'font-size:12.5px;color:#1f3763;font-weight:600' },
  { id: 'Quote', label: 'Quote', style: 'font-size:13px;font-style:italic;color:#475569' },
  { id: 'IntenseQuote', label: 'Intense Quote', style: 'font-size:13px;font-style:italic;color:#2f5496' },
  { id: 'ListParagraph', label: 'List Paragraph', style: 'font-size:13px;padding-left:8px' },
];

const PAPER = [
  { label: 'Letter', hint: '21.59 x 27.94 cm', w: 8.5, h: 11 },
  { label: 'A4', hint: '21 x 29.7 cm', w: 8.27, h: 11.69 },
  { label: 'Legal', hint: '21.59 x 35.56 cm', w: 8.5, h: 14 },
  { label: 'A5', hint: '14.8 x 21 cm', w: 5.83, h: 8.27 },
  { label: 'A3', hint: '29.7 x 42 cm', w: 11.69, h: 16.54 },
  { label: 'Executive', hint: '18.42 x 26.67 cm', w: 7.25, h: 10.5 },
];
const MARGINS = [
  { label: 'Normal', hint: '2.54 cm all round', m: { top: 1, bottom: 1, left: 1, right: 1 } },
  { label: 'Narrow', hint: '1.27 cm all round', m: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 } },
  { label: 'Moderate', hint: '2.54 cm top/bottom, 1.91 cm sides', m: { top: 1, bottom: 1, left: 0.75, right: 0.75 } },
  { label: 'Wide', hint: '2.54 cm top/bottom, 5.08 cm sides', m: { top: 1, bottom: 1, left: 2, right: 2 } },
  { label: 'Office 2003 default', hint: '2.54 cm top/bottom, 3.18 cm sides', m: { top: 1, bottom: 1, left: 1.25, right: 1.25 } },
];
const SYMBOLS = [...'©®™€£¥¢$§¶•…–°±×÷≠≈≤≥∞√∑πΩµαβγΔ←→↑↓↔⇒✓✗★☆♥♦♣♠☎✉✂☺½¼¾¹²³«»‘’“”¿¡'];
const WATERMARKS = ['CONFIDENTIAL', 'DO NOT COPY', 'DRAFT', 'SAMPLE', 'URGENT', 'ASAP', 'PERSONAL'];

export async function mount(el, { data, name, html, setDirty, setStatus, toast, fileMenu, printToPdf, openDocument }) {
  const id = `writer-${++seq}`;
  el.innerHTML = `
    <div data-ribbon></div>
    <div class="writer-sd-toolbar" id="${id}-toolbar"></div>
    <div class="writer-main">
      <div class="writer-body"><div class="writer-loading" aria-busy="true" aria-label="Opening ${esc(name || 'document')}"><div class="k-skel-ruler"></div><div class="k-skel-page">${data ? '<div class="k-skel-line w60"></div><div class="k-skel-line"></div><div class="k-skel-line"></div><div class="k-skel-line w80"></div><div class="k-skel-gap"></div><div class="k-skel-line"></div><div class="k-skel-line w90"></div><div class="k-skel-line w70"></div>' : ''}</div></div><div id="${id}-editor"></div></div>
      <aside class="writer-side" hidden>
        <div class="side-head">
          <div class="side-tabs">
            <button data-side="nav">${icon('layers', 14)}Navigation</button>
            <button data-side="comments">${icon('note', 14)}Comments</button>
            <button data-side="changes">${icon('track', 14)}Changes</button>
          </div>
          <button class="icon-btn small" data-side-close title="Close">${icon('close', 14)}</button>
        </div>
        <div class="side-body"></div>
      </aside>
    </div>
    <footer class="writer-foot">
      <span class="wf-words"></span>
      <div class="zoom-ctl">
        <button type="button" data-z="fit" title="Page width">${icon('fit', 14)}</button>
        <button type="button" data-z="out" title="Zoom out">${icon('minus', 15)}</button>
        <input type="range" min="0" max="100" step="1" value="50" aria-label="Zoom" title="Zoom">
        <button type="button" data-z="in" title="Zoom in">${icon('plus', 15)}</button>
        <button type="button" data-z="menu" class="zoom-pct" title="Zoom (Ctrl + mouse wheel)"><span>100%</span>${icon('chevron', 14)}</button>
      </div>
    </footer>
    <div class="find-panel glass" hidden>
      <div class="find-row">
        <label class="rb-search">${icon('search', 15)}<input data-find placeholder="Find"></label>
        <span class="find-count"></span>
        <button class="icon-btn small" data-find-prev title="Previous (Shift+Enter)">${icon('up', 15)}</button>
        <button class="icon-btn small" data-find-next title="Next (Enter)">${icon('down', 15)}</button>
        <button class="icon-btn small" data-find-close title="Close (Esc)">${icon('close', 15)}</button>
      </div>
      <div class="find-row replace-row">
        <label class="rb-search">${icon('edit', 15)}<input data-replace placeholder="Replace with"></label>
        <button class="chip" data-replace-one>Replace</button>
        <button class="chip" data-replace-all>Replace all</button>
      </div>
    </div>
    <button class="exit-reading chip primary" hidden>${icon('close', 14)}Exit reading mode</button>`;

  let stopBoxes = null;
  let decorations = null;
  let unwire = null;
  let ready = false;
  let suppressDirty = true; // the initial import fires update events
  let pages = 0;
  let superdoc;
  let statsTimer;
  let sideTimer;
  let tracking = false;
  let sidePanel = null;

  // ---------------------------------------------------------------- helpers
  const api = () => superdoc.activeEditor.doc;
  // The author name on changes and comments follows the setting, even in documents already open
  function applyUser() {
    const u = editorUser();
    try { superdoc.config.user = { ...(superdoc.config.user || {}), ...u }; superdoc.user = { ...(superdoc.user || {}), ...u }; } catch { /* not ready */ }
    for (const ed of [superdoc?.activeEditor, ...(superdoc?.superdocStore?.documents || []).map((d) => d.getEditor?.())]) { if (ed?.options) ed.options.user = { ...(ed.options.user || {}), ...u }; }
  }
  const onUser = () => applyUser();
  window.addEventListener('kertas-user', onUser);
  /** Asks for the user's name once, before their first tracked change or comment */
  async function ensureUserName() {
    if (userName()) return true;
    const res = await openDialog({
      title: 'Your name',
      description: 'Shown on your tracked changes and comments, so others know who made them. You can change it later in About.',
      width: 420,
      fields: [{ id: 'name', label: 'Name', value: '', placeholder: 'e.g. Aisha Rahman' }],
      validate: (v) => (v.name.trim() ? null : 'Type your name'),
      actions: [{ id: 'skip', label: 'Not now' }, { id: 'ok', label: 'Save', primary: true }],
    });
    if (res?.action === 'ok') { setUserName(res.values.name); applyUser(); }
    return true;
  }
  window.__kertasOoxml = { ...ooxml, editor: () => superdoc?.activeEditor }; // for tests
  /** Puts the cursor at a text offset inside a paragraph (found by its block id) */
  function caretTo(blockId, offset) {
    const ed = superdoc?.activeEditor;
    if (!ed?.state) return false;
    let at = null;
    ed.state.doc.descendants((node, pos) => {
      if (at !== null) return false;
      const a = node.attrs || {};
      if (node.isTextblock && (a.sdBlockId === blockId || a.paraId === blockId || a.blockId === blockId || a.id === blockId)) {
        // Walk the paragraph's inline content to the requested text offset
        let left = offset; let p = pos + 1;
        node.forEach((child, childOffset) => {
          if (left < 0) return;
          const len = child.isText ? child.text.length : child.textContent.length;
          if (left <= len) { p = pos + 1 + childOffset + (child.isText ? left : child.nodeSize); left = -1; return; }
          left -= len;
        });
        at = left < 0 ? p : pos + node.nodeSize - 1;
        return false;
      }
      return true;
    });
    if (at === null) return false;
    try { ed.commands.setTextSelection({ from: at, to: at }); ed.view?.focus?.(); superdoc.focus?.(); return true; } catch { return false; }
  }
  const editor = () => superdoc.activeEditor;
  // "Already set" (NO_OP) is a success: the document is already as asked
  const failed = (r) => r && (r.success === false || r.status === 'rejected') && r.failure?.code !== 'NO_OP' && !/\bis already\b/.test(r.failure?.message || '');
  const reason = (r) => r?.failure?.message || r?.rejection?.detail || 'not possible here';
  async function act(label, fn, okMsg) {
    try {
      const r = await fn();
      if (failed(r)) throw new Error(reason(r));
      if (okMsg) toast(okMsg);
      setDirty(true);
      scheduleSide();
      return r;
    } catch (e) {
      console.error(label, e);
      toast(`${label}: ${e.message}`, true);
      return null;
    } finally {
      superdoc?.focus?.();
    }
  }
  const selection = async () => api().selection.current();
  /** The selection as the SelectionTarget that format.apply and replace expect */
  function toSelectionTarget(s) {
    const t = s?.selectionTarget || s?.target;
    if (t?.kind === 'selection') return t;
    const segs = t?.segments || [];
    if (!segs.length) throw new Error('Select some text first');
    const a = segs[0]; const z = segs[segs.length - 1];
    return { kind: 'selection', start: { kind: 'text', blockId: a.blockId, offset: a.range.start }, end: { kind: 'text', blockId: z.blockId, offset: z.range.end } };
  }
  async function selectedBlocks() {
    // Paragraphs in the selection, by their Word paragraph id (what blocks.list reports);
    // the editor's own block id is the fallback for paragraphs that have none yet
    const ed = superdoc?.activeEditor;
    if (ed?.state) {
      const { from, to } = ed.state.selection;
      const ids = [];
      ed.state.doc.nodesBetween(from, to, (node) => {
        if (!node.isTextblock) return true;
        const id = node.attrs?.paraId || node.attrs?.sdBlockId;
        if (id && !ids.includes(id)) ids.push(id);
        return false;
      });
      if (ids.length) return ids.map((nodeId) => ({ kind: 'block', nodeType: 'paragraph', nodeId }));
    }
    const s = await selection();
    return [...new Set((s?.target?.segments || []).map((g) => g.blockId))].map((nodeId) => ({ kind: 'block', nodeType: 'paragraph', nodeId }));
  }
  async function caretBlock() {
    const b = (await selectedBlocks())[0];
    if (!b) throw new Error('Click in the document first');
    return b;
  }
  async function textSelection(required = true) {
    const s = await selection();
    if (required && (!s || s.empty)) throw new Error('Select some text first');
    return s;
  }
  // Paragraph-level operations accept paragraphs, headings and list items; try each kind
  async function onBlock(fn, b) {
    let last;
    for (const nodeType of ['paragraph', 'heading', 'listItem']) {
      try {
        const r = await fn({ ...b, nodeType });
        if (!failed(r)) return r;
        last = r;
        if (!/target|type/i.test(`${r.failure?.code} ${r.failure?.message}`)) return r;
      } catch (e) {
        last = { success: false, failure: { message: e.message } };
        if (!/target|type/i.test(`${e.code} ${e.message}`)) throw e;
      }
    }
    return last;
  }
  async function onSelectedBlocks(fn) {
    const list = await selectedBlocks();
    if (!list.length) throw new Error('Click in a paragraph first');
    let r;
    for (const b of list) { r = await onBlock(fn, b); if (failed(r)) return r; }
    return r;
  }
  async function sections() { return (await api().sections.list()).items || []; }
  async function onSections(fn) {
    let r;
    for (const s of await sections()) { r = await fn(s.address, s); if (failed(r)) return r; }
    return r;
  }
  async function insertTextAtCaret(text) {
    const ed = superdoc?.activeEditor;
    if (!ed?.state) throw new Error('Click in the document first');
    ed.commands.insertContent(text);
    ed.view?.focus?.();
    return { success: true };
  }
  async function newParagraphAfter(text = '') {
    const b = await caretBlock();
    const r = await onBlock((t) => api().create.paragraph({ at: { kind: 'after', target: t }, text }), b);
    if (failed(r)) throw new Error(reason(r));
    return r.paragraph;
  }

  const extras = createExtras({
    api, act, toast, caretBlock, selection, onBlock, failed, reason, setStatus, openDocument, docName: name, caretTo, focus: () => superdoc?.focus?.(), editor: () => superdoc?.activeEditor,
    setDirty: () => { setDirty(true); scheduleSide(); },
    exportDocx: async () => new Uint8Array(await (await superdoc.export({ exportType: ['docx'], commentsType: 'external', triggerDownload: false })).arrayBuffer()),
  });

  // ---------------------------------------------------------------- slash commands
  async function makeList(kind) {
    await act(kind === 'bullet' ? 'Bulleted list' : 'Numbered list', async () => {
      const b = await caretBlock();
      return onBlock((t) => api().lists.create({ mode: 'fromParagraphs', target: t, kind }), b);
    });
  }
  const shape = (id) => () => extras.insertShape(SHAPES.find((s) => s.id === id));
  const SLASH = [
    { id: 'text', group: 'Text', label: 'Text', keys: ['normal', 'paragraph', 'body', 'p'], icon: 'text', hint: 'Plain body text', run: () => applyStyle('Normal', 'Normal') },
    { id: 'h1', group: 'Text', label: 'Heading 1', keys: ['h1', 'title section'], icon: 'heading-1', hint: 'Big section heading', run: () => applyStyle('Heading1', 'Heading 1') },
    { id: 'h2', group: 'Text', label: 'Heading 2', keys: ['h2', 'subheading'], icon: 'heading-2', hint: 'Medium heading', run: () => applyStyle('Heading2', 'Heading 2') },
    { id: 'h3', group: 'Text', label: 'Heading 3', keys: ['h3'], icon: 'heading-3', hint: 'Small heading', run: () => applyStyle('Heading3', 'Heading 3') },
    { id: 'title', group: 'Text', label: 'Title', keys: ['document title'], icon: 'heading', hint: 'Document title', run: () => applyStyle('Title', 'Title') },
    { id: 'quote', group: 'Text', label: 'Quote', keys: ['blockquote', 'citation text'], icon: 'quote', hint: 'Quoted passage', run: () => applyStyle('Quote', 'Quote') },
    { id: 'bullets', group: 'Lists', label: 'Bulleted list', keys: ['ul', 'bullet', 'unordered', '-'], icon: 'list', hint: 'A simple list', run: () => makeList('bullet') },
    { id: 'numbers', group: 'Lists', label: 'Numbered list', keys: ['ol', 'ordered', '1.', 'steps'], icon: 'line-numbers', hint: '1, 2, 3', run: () => makeList('ordered') },
    { id: 'tick', group: 'Lists', label: 'Tick box', keys: ['checkbox', 'todo', 'checklist', 'task'], icon: 'check-square', hint: 'A ☐ to tick later', run: () => insertTextAtCaret('☐ ') },
    { id: 'table', group: 'Insert', label: 'Table', keys: ['grid', 'rows', 'columns'], icon: 'table', hint: 'Rows and columns', run: () => insertTable() },
    { id: 'picture', group: 'Insert', label: 'Picture', keys: ['image', 'photo', 'img'], icon: 'image', hint: 'From your computer', run: () => insertPicture() },
    { id: 'divider', group: 'Insert', label: 'Divider', keys: ['line', 'hr', 'rule', 'separator'], icon: 'minus', hint: 'Horizontal line', run: () => horizontalLine() },
    { id: 'pagebreak', group: 'Insert', label: 'Page break', keys: ['new page', 'break'], icon: 'page-break', hint: 'Start a new page', shortcut: 'Ctrl+Enter', run: () => pageBreak() },
    { id: 'link', group: 'Insert', label: 'Link', keys: ['url', 'hyperlink', 'web'], icon: 'link', hint: 'Web address or place in the document', shortcut: 'Ctrl+K', run: () => insertLink() },
    { id: 'date', group: 'Insert', label: 'Date', keys: ['today', 'time'], icon: 'calendar', hint: "Today's date", run: () => insertDate() },
    { id: 'comment', group: 'Insert', label: 'Comment', keys: ['note', 'review'], icon: 'note-add', hint: 'Comment on the selection', run: () => newComment() },
    { id: 'equation', group: 'Insert', label: 'Equation', keys: ['math', 'formula', 'latex'], icon: 'equation', hint: 'Fractions, roots, sums', run: () => extras.insertEquation() },
    { id: 'textbox', group: 'Insert', label: 'Text box', keys: ['box', 'aside', 'sidebar'], icon: 'text-box', hint: 'Text beside the main text', run: () => extras.insertTextBox('right') },
    { id: 'pullquote', group: 'Insert', label: 'Pull quote', keys: ['callout', 'highlight quote'], icon: 'quote', hint: 'A large centred quotation', run: () => extras.insertTextBox('quote') },
    { id: 'shape-rect', group: 'Shapes', label: 'Rectangle', keys: ['shape', 'box', 'square'], icon: 'shapes', hint: 'Shape', run: shape('rounded') },
    { id: 'shape-circle', group: 'Shapes', label: 'Circle', keys: ['shape', 'oval', 'round'], icon: 'shapes', hint: 'Shape', run: shape('circle') },
    { id: 'shape-arrow', group: 'Shapes', label: 'Arrow', keys: ['shape', 'right arrow'], icon: 'arrow-shape', hint: 'Shape', run: shape('arrow-right') },
    { id: 'shape-callout', group: 'Shapes', label: 'Speech bubble', keys: ['shape', 'callout', 'bubble'], icon: 'callout', hint: 'Shape', run: shape('callout') },
    { id: 'shape-star', group: 'Shapes', label: 'Star', keys: ['shape'], icon: 'star', hint: 'Shape', run: shape('star5') },
    { id: 'toc', group: 'References', label: 'Table of contents', keys: ['toc', 'contents', 'index'], icon: 'toc', hint: 'From your headings', run: () => insertToc() },
    { id: 'footnote', group: 'References', label: 'Footnote', keys: ['note', 'fn'], icon: 'footnote', hint: 'Note at the bottom of the page', run: () => insertNote('footnote') },
    { id: 'endnote', group: 'References', label: 'Endnote', keys: ['note'], icon: 'footnote', hint: 'Note at the end', run: () => insertNote('endnote') },
    { id: 'cite', group: 'References', label: 'Citation', keys: ['cite', 'reference', 'source'], icon: 'citation', hint: 'Cite a source here', run: () => extras.insertCitation() },
    { id: 'bib', group: 'References', label: 'Bibliography', keys: ['references', 'works cited', 'sources list'], icon: 'bibliography', hint: 'Reference list in the current style', run: () => extras.insertBibliography() },
    { id: 'caption', group: 'References', label: 'Caption', keys: ['figure', 'table caption'], icon: 'caption', hint: 'Figure or table caption', run: () => insertCaption() },
    { id: 'merge', group: 'Mailings', label: 'Merge field', keys: ['mail merge', 'field', 'placeholder'], icon: 'mail', hint: 'Filled from each recipient', run: () => extras.promptMergeField() },
  ];
  // "/" opens Kertas's own command menu at the cursor; Ctrl+/ opens it anywhere.
  const slash = createSlashMenu({ root: el, api, superdoc: () => superdoc, commands: SLASH, toast, onSlash: true, caretTo });

  // ---------------------------------------------------------------- ribbon
  const big = (bid, ic, label, action, title) => ({ id: bid, kind: 'big', icon: ic, label, action, title });
  const small = (bid, ic, label, action, title) => ({ id: bid, icon: ic, label, action, title });
  const menu = (bid, ic, label, items, title) => ({ id: bid, kind: 'menu', icon: ic, label, items, title });

  // Picture wrapping and position: on the Picture tab, the Layout tab and the right-click menu
  const WRAP_ITEMS = [
    { label: 'In line with text', icon: 'text', action: () => setWrap('Inline') },
    { label: 'Square', icon: 'wrap', hint: 'Text flows around the box', action: () => setWrap('Square') },
    { label: 'Tight', icon: 'wrap', hint: 'Text follows the picture closely', action: () => setWrap('Tight') },
    { label: 'Top and bottom', icon: 'rows', hint: 'Text above and below only', action: () => setWrap('TopAndBottom') },
    'sep',
    { label: 'Behind text', icon: 'back', action: () => setWrap('None', true) },
    { label: 'In front of text', icon: 'front', action: () => setWrap('None', false) },
  ];
  const POS_ITEMS = [
    { label: 'Left', icon: 'align-left', action: () => alignPicture('left') },
    { label: 'Centre', icon: 'align-center', action: () => alignPicture('center') },
    { label: 'Right', icon: 'align-right', action: () => alignPicture('right') },
  ];
  // Right-click on a picture: the same wrapping choices
  function pictureMenuItems() {
    const onPicture = (ctx) => ctx.trigger === 'click' && (ctx.node?.type?.name === 'image' || ctx.currentNodeType === 'image');
    // the picture that was right-clicked, whatever is selected
    const choose = (ed, ctx, fn) => { const id = ctx?.node?.attrs?.sdImageId || null; setTimeout(() => fn(id), 30); };
    const opts = [
      ['in-line', 'In line with text', 'text', (id) => setWrap('Inline', false, id)],
      ['square', 'Wrap text: Square', 'wrap', (id) => setWrap('Square', false, id)],
      ['tight', 'Wrap text: Tight', 'wrap', (id) => setWrap('Tight', false, id)],
      ['top-bottom', 'Wrap text: Top and bottom', 'rows', (id) => setWrap('TopAndBottom', false, id)],
      ['behind', 'Behind text', 'back', (id) => setWrap('None', true, id)],
      ['front', 'In front of text', 'front', (id) => setWrap('None', false, id)],
    ];
    return [{ id: 'kertas-picture', items: opts.map(([key, label, ic, fn]) => ({ id: `k-wrap-${key}`, label, icon: icon(ic, 16), showWhen: onPicture, action: (ed, ctx) => choose(ed, ctx, fn) })) }];
  }
  const ribbon = createRibbon(el.querySelector('[data-ribbon]'), {
    app: 'writer',
    fileMenu,
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Undo', items: [
          small('undo', 'undo', '', () => fmt('undo'), 'Undo (Ctrl+Z)'),
          small('redo', 'redo', '', () => fmt('redo'), 'Redo (Ctrl+Y)'),
        ] },
        { label: 'Clipboard', items: [
          big('paste', 'paste', 'Paste', () => clip('paste'), 'Paste (Ctrl+V)'),
          small('cut', 'cut', '', () => clip('cut'), 'Cut (Ctrl+X)'),
          small('copy', 'copy', '', () => clip('copy'), 'Copy (Ctrl+C)'),
          small('paste-plain', 'paste-plain', '', () => clip('pasteAndMatchStyle'), 'Paste as plain text (Ctrl+Shift+V)'),
          small('painter', 'painter', '', () => fmt('copy-format'), 'Format painter: copy the formatting here, then select text to apply it'),
        ] },
        { label: 'Font', rows: [
          [
            { id: 'fontname', kind: 'combo', width: 150, title: 'Font', items: () => kFonts.options.map((o) => ({ label: o.label, value: o.label, style: `font-family:${o.props.style.fontFamily};font-size:14px` })), action: (v) => v && fmt('font-family', String(v).trim()) },
            { id: 'fontsize', kind: 'combo', width: 54, title: 'Font size', items: () => FONT_SIZES.map((n) => ({ label: String(n), value: n })), action: (v) => setFontSize(v) },
            small('grow', 'grow', '', () => growFont(1), 'Increase font size'),
            small('shrink', 'shrink', '', () => growFont(-1), 'Decrease font size'),
            menu('case', 'change-case', '', () => caseMenu(), 'Capitals and strikethrough'),
            small('clear', 'clear-format', '', () => clearFormatting(), 'Clear all formatting'),
          ],
          [
            small('b', 'bold', '', () => fmt('bold'), 'Bold (Ctrl+B)'),
            small('i', 'italic', '', () => fmt('italic'), 'Italic (Ctrl+I)'),
            small('u', 'underline', '', () => fmt('underline'), 'Underline (Ctrl+U)'),
            small('s', 'strike', '', () => fmt('strikethrough'), 'Strikethrough'),
            small('sub', 'subscript', '', () => toggleVert('subscript'), 'Subscript (Ctrl+=)'),
            small('sup', 'superscript', '', () => toggleVert('superscript'), 'Superscript (Ctrl+Shift+=)'),
            { id: 'hl', kind: 'swatch', icon: 'highlighter', palette: 'highlight', value: '#FFFF00', title: 'Text highlight colour', action: (c) => fmt('highlight-color', c) },
            { id: 'fc', kind: 'swatch', icon: 'font-color', palette: 'text', value: '#C00000', title: 'Font colour', action: (c) => fmt('text-color', c) },
          ],
        ] },
        { label: 'Paragraph', rows: [
          [
            small('ul', 'list-ul', '', () => fmt('bullet-list'), 'Bullets'),
            small('ol', 'list-ol', '', () => fmt('numbered-list'), 'Numbering'),
            menu('numbering', 'line-numbers', '', () => numberingMenu(), 'Restart or continue list numbering'),
            small('outdent', 'indent-less', '', () => fmt('indent-decrease'), 'Decrease indent'),
            small('indent', 'indent-more', '', () => fmt('indent-increase'), 'Increase indent'),
            small('ltr', 'ltr', '', () => setDirection('ltr'), 'Left-to-right text'),
            small('rtl', 'rtl', '', () => setDirection('rtl'), 'Right-to-left text (Arabic, Hebrew, Urdu)'),
            small('marks', 'pilcrow', '', () => toggleMarks(), 'Show formatting marks'),
          ],
          [
            small('al', 'align-left', '', () => fmt('text-align', 'left'), 'Align left (Ctrl+L)'),
            small('ac', 'align-center', '', () => fmt('text-align', 'center'), 'Centre (Ctrl+E)'),
            small('ar', 'align-right', '', () => fmt('text-align', 'right'), 'Align right (Ctrl+R)'),
            small('aj', 'justify', '', () => fmt('text-align', 'justify'), 'Justify (Ctrl+J)'),
            menu('spacing', 'line-spacing', '', () => spacingMenu(), 'Line and paragraph spacing'),
            { id: 'shading', kind: 'swatch', icon: 'fill', palette: 'fill', value: '#FFF2CC', title: 'Shading', action: (c) => shade(c) },
            menu('borders', 'border', '', () => borderMenu(), 'Borders'),
            small('para', 'paragraph', '', () => paragraphSettings(), 'Paragraph settings: indents, spacing and line breaks'),
          ],
        ] },
        { label: 'Editing', items: [
          small('find', 'search', 'Find', () => openFind(false), 'Find (Ctrl+F)'),
          small('replace', 'replace', 'Replace', () => openFind(true), 'Replace (Ctrl+H)'),
          small('selall', 'select-all', 'Select all', () => clip('selectAll'), 'Select all (Ctrl+A)'),
        ] },
        { label: 'Styles', wide: true, items: [{ id: 'styles', kind: 'gallery', tiles: STYLE_TILES.map((t) => ({ ...t, action: () => applyStyle(t.id, t.label) })) }, menu('style-list', 'list', 'All styles', () => allStylesMenu(), 'Every style in this document, and Create a style')] },
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Pages', items: [
          big('blank', 'file-plus', 'Blank page', () => blankPage()),
          big('pgbreak', 'page-break', 'Page break', () => pageBreak(), 'Page break (Ctrl+Enter)'),
          menu('secbreak', 'section', 'Section break', [
            { label: 'Next page', hint: 'New section starts on the next page', icon: 'page-break', action: () => sectionBreak('nextPage') },
            { label: 'Continuous', hint: 'New section starts on the same page', icon: 'section', action: () => sectionBreak('continuous') },
            { label: 'Even page', icon: 'section', action: () => sectionBreak('evenPage') },
            { label: 'Odd page', icon: 'section', action: () => sectionBreak('oddPage') },
          ]),
        ] },
        { label: 'Tables', items: [big('table', 'table', 'Table', () => insertTable())] },
        { label: 'Illustrations', items: [
          big('picture', 'image', 'Picture', () => insertPicture()),
          menu('shapes', 'shapes', 'Shapes', () => extras.shapeMenu(), 'Rectangles, arrows, stars, callouts and more'),
          menu('textbox', 'text-box', 'Text box', () => extras.textBoxMenu(), 'Text beside the main text, or a pull quote'),
          big('hr', 'minus', 'Line', () => horizontalLine(), 'Horizontal line'),
        ] },
        { label: 'Maths', items: [big('equation', 'equation', 'Equation', () => extras.insertEquation(), 'Fractions, roots, sums and more (LaTeX)')] },
        { label: 'Links', items: [big('link', 'link', 'Link', () => insertLink()), big('comment', 'note', 'Comment', () => newComment())] },
        { label: 'Header & Footer', items: [
          big('header', 'header', 'Header', () => editFurniture('header'), 'Edit the page header'),
          big('footer', 'footer', 'Footer', () => editFurniture('footer'), 'Edit the page footer'),
          big('pgnum', 'hash', 'Page number', () => pageNumbers()),
        ] },
        { label: 'Text', items: [
          big('date', 'calendar', 'Date & time', () => insertDate()),
          menu('symbol', 'omega', 'Symbol', () => symbolMenu()),
        ] },
      ] },
      { id: 'layout', label: 'Layout', groups: [
        { label: 'Page setup', items: [
          menu('margins', 'margins', 'Margins', () => [
            ...MARGINS.map((m) => ({ label: m.label, hint: m.hint, icon: 'margins', action: () => setMargins(m.m) })),
            'sep', { label: 'Custom margins...', icon: 'edit', action: () => customMargins() },
          ]),
          menu('orient', 'orientation', 'Orientation', [
            { label: 'Portrait', icon: 'portrait', action: () => setOrientation('portrait') },
            { label: 'Landscape', icon: 'landscape', action: () => setOrientation('landscape') },
          ]),
          menu('size', 'paper', 'Size', () => [
            ...PAPER.map((p) => ({ label: p.label, hint: p.hint, icon: 'paper', action: () => setPaper(p) })),
            'sep', { label: 'Custom size...', icon: 'edit', action: () => customPaper() },
          ]),
          menu('columns', 'columns', 'Columns', [
            { label: 'One', icon: 'col1', action: () => setColumns(1) },
            { label: 'Two', icon: 'columns', action: () => setColumns(2) },
            { label: 'Three', icon: 'col3', action: () => setColumns(3) },
          ]),
          menu('breaks', 'page-break', 'Breaks', [
            { header: 'Page breaks' },
            { label: 'Page', hint: 'Start a new page', icon: 'page-break', action: () => pageBreak() },
            { header: 'Section breaks' },
            { label: 'Next page', icon: 'section', action: () => sectionBreak('nextPage') },
            { label: 'Continuous', icon: 'section', action: () => sectionBreak('continuous') },
          ]),
          menu('linenum', 'line-numbers', 'Line numbers', [
            { label: 'None', icon: 'close', action: () => lineNumbers(false) },
            { label: 'Continuous', icon: 'line-numbers', action: () => lineNumbers(true, 'continuous') },
            { label: 'Restart each page', icon: 'line-numbers', action: () => lineNumbers(true, 'newPage') },
            { label: 'Restart each section', icon: 'line-numbers', action: () => lineNumbers(true, 'newSection') },
          ]),
        ] },
        { label: 'Page background', items: [
          menu('wm', 'droplet', 'Watermark', () => [
            ...WATERMARKS.map((w) => ({ label: w, icon: 'droplet', action: () => setWatermark({ text: w }) })),
            'sep', { label: 'Custom watermark...', icon: 'edit', action: () => customWatermark() },
            { label: 'Remove watermark', icon: 'trash', action: async () => { const r = await removeWatermarks(); toast(r.removed ? 'Watermark removed' : 'This document has no watermark'); } },
          ]),
          menu('pgborder', 'border', 'Page borders', [
            { label: 'Box', hint: 'Thin line around every page', icon: 'border', action: () => pageBorder('single', 4) },
            { label: 'Thick box', icon: 'border', action: () => pageBorder('single', 18) },
            { label: 'Double', icon: 'border', action: () => pageBorder('double', 4) },
            { label: 'None', icon: 'close', action: () => pageBorder(null) },
          ]),
        ] },
        { label: 'Arrange', items: [
          menu('lay-wrap', 'wrap', 'Wrap text', WRAP_ITEMS, 'How text flows around the selected picture'),
          menu('lay-pos', 'align-center', 'Position', POS_ITEMS, 'Where the selected picture sits'),
          big('lay-front', 'front', 'Bring forward', () => picOrder(true), 'Put the selected picture in front'),
          big('lay-back', 'back', 'Send backward', () => picOrder(false), 'Put the selected picture behind'),
        ] },
        { label: 'Indent', rows: true, items: [
          { id: 'ind-left', kind: 'number', label: 'Left', unit: 'cm', step: 0.25, min: -5, max: 20, title: 'Indent from the left margin', action: (v) => setIndent('left', v) },
          { id: 'ind-right', kind: 'number', label: 'Right', unit: 'cm', step: 0.25, min: -5, max: 20, title: 'Indent from the right margin', action: (v) => setIndent('right', v) },
        ] },
        { label: 'Spacing', rows: true, items: [
          { id: 'sp-before', kind: 'number', label: 'Before', unit: 'pt', step: 6, min: 0, max: 400, title: 'Space before the paragraph', action: (v) => setSpace('before', v) },
          { id: 'sp-after', kind: 'number', label: 'After', unit: 'pt', step: 6, min: 0, max: 400, title: 'Space after the paragraph', action: (v) => setSpace('after', v) },
        ] },
        { label: 'Paragraph', items: [
          big('para2', 'paragraph', 'Indents & spacing', () => paragraphSettings()),
          small('ltr2', 'ltr', 'Left to right', () => setDirection('ltr')),
          small('rtl2', 'rtl', 'Right to left', () => setDirection('rtl')),
        ] },
      ] },
      { id: 'references', label: 'References', groups: [
        { label: 'Table of contents', items: [
          big('toc', 'toc', 'Table of contents', () => insertToc(), 'Built from your headings'),
          small('toc-update', 'refresh', 'Update table', () => updateToc()),
          small('toc-remove', 'trash', 'Remove', () => removeToc()),
        ] },
        { label: 'Footnotes', items: [
          big('fn', 'footnote', 'Insert footnote', () => insertNote('footnote')),
          big('en', 'endnote', 'Insert endnote', () => insertNote('endnote')),
        ] },
        { label: 'Citations & Bibliography', items: [
          big('cite', 'citation', 'Insert citation', () => extras.insertCitation(), 'Cite a source at the cursor'),
          big('sources', 'book', 'Sources', () => extras.manageSources(), 'Add, edit or remove sources and pick the style'),
          menu('bib', 'bibliography', 'Bibliography', () => extras.bibliographyMenu(), 'Insert the reference list'),
        ] },
        { label: 'Captions', items: [big('caption', 'caption', 'Insert caption', () => insertCaption())] },
        { label: 'Fields', items: [big('fields-update', 'refresh', 'Update fields', () => updateFields(), 'Refresh cross-references, dates and other fields (F9)')] },
      ] },
      { id: 'mailings', label: 'Mailings', groups: [
        { label: 'Start', items: [big('mm-recipients', 'sheet', 'Select recipients', () => extras.selectRecipients(), 'Use a list from Excel or a CSV file')] },
        { label: 'Write', items: [menu('mm-field', 'mail', 'Insert merge field', () => extras.mergeFieldMenu(), 'A placeholder filled from each row')] },
        { label: 'Preview', items: [big('mm-preview', 'view', 'Preview results', () => extras.previewMerge(), 'See each letter with real data')] },
        { label: 'Finish', items: [menu('mm-finish', 'mails', 'Finish & merge', () => extras.finishMenu(), 'Make every letter')] },
      ] },
      { id: 'review', label: 'Review', groups: [
        { label: 'Proofing', items: [
          big('spell', 'spell', 'Spelling', () => toggleSpelling(), 'Red underlines under misspelled words. Right-click one for suggestions.'),
          menu('spell-lang', 'languages', 'Language', () => spellLanguageMenu()),
          small('spell-add', 'plus', 'Add to dictionary', () => addWordToDictionary(), 'Add the selected word to your dictionary'),
          small('spell-dict', 'book', 'My dictionary', () => manageDictionary()),
          big('wc', 'word-count', 'Word count', () => wordCount()),
        ] },
        { label: 'Comments', items: [
          big('rv-comment', 'note', 'New comment', () => newComment()),
          big('rv-comments', 'messages', 'Show comments', () => showSide('comments')),
        ] },
        { label: 'Tracking', items: [
          big('track', 'track', 'Track changes', () => setTracking(!tracking)),
          menu('markup', 'view', 'Markup', [
            { label: 'All markup', hint: 'Show insertions and deletions', icon: 'view', action: () => markupView('review') },
            { label: 'No markup', hint: 'Show the document as if accepted', icon: 'view', action: () => markupView('final') },
            { label: 'Original', hint: 'Show the document before changes', icon: 'view', action: () => markupView('original') },
          ]),
          big('rv-changes', 'list', 'Changes pane', () => showSide('changes')),
        ] },
        { label: 'Changes', items: [
          big('acc-all', 'check', 'Accept all', () => decideAll('accept')),
          big('rej-all', 'close', 'Reject all', () => decideAll('reject')),
        ] },
        { label: 'Protect', items: [big('restrict', 'lock', 'Restrict editing', () => restrictEditing())] },
      ] },
      { id: 'table', label: 'Table', contextual: true, groups: [
        { label: 'Table', items: [
          menu('t-select', 'select', 'Select', [
            { label: 'Cell', icon: 'rect', action: () => selectInTable('cell') },
            { label: 'Row', icon: 'rows', action: () => selectInTable('row') },
            { label: 'Column', icon: 'columns', action: () => selectInTable('column') },
            { label: 'Whole table', icon: 'table', action: () => selectInTable('table') },
          ]),
          big('t-del-table', 'trash', 'Delete table', () => tcmd('Delete table', 'deleteTable')),
        ] },
        { label: 'Design', items: [
          { id: 't-designs', kind: 'big', icon: 'table-style', label: 'Table design', title: 'Ready-made table designs', action: () => designMenu(ribbon.get('t-designs')) },
          big('t-opt-header', 'heading', 'Header row', () => toggleLook('t-opt-header'), 'Style the first row as a header'),
          big('t-opt-banded', 'rows', 'Banded rows', () => toggleLook('t-opt-banded'), 'Shade every other row'),
          big('t-opt-first', 'columns', 'First column', () => toggleLook('t-opt-first'), 'Style the first column'),
          big('t-opt-total', 'sigma', 'Total row', () => toggleLook('t-opt-total'), 'Style the last row as a total'),
        ] },
        { label: 'Rows & Columns', items: [
          big('t-row-above', 'insert-row', 'Insert above', () => tcmd('Insert row', 'addRowBefore')),
          big('t-row-below', 'insert-row', 'Insert below', () => tcmd('Insert row', 'addRowAfter')),
          big('t-col-left', 'insert-col', 'Insert left', () => tcmd('Insert column', 'addColumnBefore')),
          big('t-col-right', 'insert-col', 'Insert right', () => tcmd('Insert column', 'addColumnAfter')),
          menu('t-delete', 'trash', 'Delete', [
            { label: 'Delete row', icon: 'trash', action: () => tcmd('Delete row', 'deleteRow') },
            { label: 'Delete column', icon: 'trash', action: () => tcmd('Delete column', 'deleteColumn') },
            'sep',
            { label: 'Delete table', icon: 'trash', action: () => tcmd('Delete table', 'deleteTable') },
          ]),
        ] },
        { label: 'Move', items: [
          small('t-up', 'up', 'Row up', () => act('Move row', () => moveRow(editor(), -1)), 'Move this row up (Alt+Shift+Up)'),
          small('t-down', 'down', 'Row down', () => act('Move row', () => moveRow(editor(), 1)), 'Move this row down (Alt+Shift+Down)'),
          small('t-left', 'left', 'Column left', () => act('Move column', () => moveColumn(editor(), -1))),
          small('t-right', 'right', 'Column right', () => act('Move column', () => moveColumn(editor(), 1))),
        ], rows: true },
        { label: 'Merge', items: [
          big('t-merge', 'merge', 'Merge cells', () => tcmd('Merge cells', 'mergeCells'), 'Select several cells first'),
          big('t-split', 'split', 'Split cell', () => tcmd('Split cell', 'splitCell')),
        ] },
        { label: 'Size', items: [
          small('t-dist-rows', 'rows', 'Distribute rows', () => tapi('Distribute rows', (id) => api().tables.distributeRows({ nodeId: id }))),
          small('t-dist-cols', 'columns', 'Distribute columns', () => tapi('Distribute columns', (id) => api().tables.distributeColumns({ nodeId: id }))),
          small('t-header', 'heading', 'Repeat header row', () => tapi('Header row', (id) => api().tables.setRowOptions({ nodeId: id, rowIndex: 0, repeatHeader: true }), 'The first row repeats at the top of each page'), 'Repeat the first row on every page'),
        ], rows: true },
        { label: 'Design', items: [
          menu('t-borders', 'border', 'Borders', [
            { label: 'All borders', icon: 'grid', action: () => tapi('Borders', (id) => api().tables.applyBorderPreset({ nodeId: id, preset: 'all' }), 'All borders') },
            { label: 'Outside only', icon: 'rect', action: () => tapi('Borders', (id) => api().tables.applyBorderPreset({ nodeId: id, preset: 'box' }), 'Outside border only') },
            { label: 'Grid', icon: 'table', action: () => tapi('Borders', (id) => api().tables.applyBorderPreset({ nodeId: id, preset: 'grid' }), 'Grid borders') },
            { label: 'No borders', icon: 'eye-off', action: () => tapi('Borders', (id) => api().tables.applyBorderPreset({ nodeId: id, preset: 'none' }), 'Borders removed') },
          ]),
          { id: 't-shade', kind: 'color', icon: 'paint', label: 'Shading', value: '#deeaf6', title: 'Cell shading', action: (v) => tcmd('Cell shading', 'setCellBackground', v) },
        ] },
        { label: 'Data', items: [
          small('t-sort-az', 'sort-az', 'Sort A to Z', () => sortTable('ascending'), 'Sort rows by this column'),
          small('t-sort-za', 'sort-za', 'Sort Z to A', () => sortTable('descending')),
          small('t-sort', 'sort', 'Sort...', () => sortDialog(), 'Choose the column, the order and whether there is a header row'),
          small('t-to-text', 'text', 'Convert to text', () => tapi('Convert to text', (id) => api().tables.convertToText({ nodeId: id, delimiter: 'tab' }))),
        ], rows: true },
      ] },
      { id: 'picture', label: 'Picture', contextual: true, groups: [
        { label: 'Arrange', items: [
          menu('pic-wrap', 'wrap', 'Wrap text', WRAP_ITEMS, 'How text flows around the picture'),
          menu('pic-pos', 'align-center', 'Position', POS_ITEMS),
          big('pic-front', 'front', 'Bring to front', () => picOrder(true)),
          big('pic-back', 'back', 'Send to back', () => picOrder(false)),
        ] },
        { label: 'Size', items: [
          { id: 'pic-w', kind: 'number', label: 'Width', unit: 'px', step: 10, min: 4, title: 'Width (keeps the proportions)', action: (v) => picSize('w', v) },
          { id: 'pic-h', kind: 'number', label: 'Height', unit: 'px', step: 10, min: 4, title: 'Height (keeps the proportions)', action: (v) => picSize('h', v) },
        ] },
        { label: 'Picture', items: [
          big('pic-alt', 'caption', 'Alt text', () => picAlt(), 'Describe the picture for screen readers'),
          big('pic-del', 'trash', 'Delete', () => picAct('Delete picture', (imageId) => api().images.delete({ imageId }))),
        ] },
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Views', items: [
          big('read', 'book', 'Read mode', () => readingMode(true)),
          big('focus', 'focus', 'Focus', () => focusMode(true), 'Hide the ribbon (Esc to exit)'),
        ] },
        { label: 'Show', items: [
          big('ruler', 'ruler', 'Ruler', () => { superdoc.toggleRuler(); rulerOn = !rulerOn; ribbon.setOn('ruler', rulerOn); }),
          menu('ruler-unit', 'ruler', 'Units', () => [
            { label: 'Centimetres', icon: rulerUnit() === 'cm' ? 'check' : 'ruler', action: () => setRulerUnit('cm') },
            { label: 'Inches', icon: rulerUnit() === 'in' ? 'check' : 'ruler', action: () => setRulerUnit('in') },
          ], 'Ruler units'),
          big('marks2', 'pilcrow', 'Formatting marks', () => toggleMarks()),
          big('navpane', 'layers', 'Navigation pane', () => showSide('nav')),
          big('slash-hint', 'command', 'Slash hint', () => { slash.setHint(!slash.hintEnabled()); ribbon.setOn('slash-hint', slash.hintEnabled()); }, 'Show "Type / for commands" on empty lines (Ctrl+/ opens the menu anywhere)'),
        ] },
        { label: 'Zoom', items: [
          big('z-out', 'zoom-out', 'Zoom out', () => setZoom(currentZoom() - 10)),
          big('z-100', 'percent', '100%', () => setZoom(100)),
          big('z-in', 'zoom-in', 'Zoom in', () => setZoom(currentZoom() + 10)),
          big('z-fit', 'fit', 'Page width', () => { superdoc.setZoomMode?.('fit-width'); setTimeout(syncZoom, 200); }),
        ] },
      ] },
    ],
    inline: [
      {
        id: 'mode', kind: 'select', title: 'Editing mode', value: 'editing',
        options: [{ value: 'editing', label: 'Editing' }, { value: 'suggesting', label: 'Suggesting' }, { value: 'viewing', label: 'Viewing' }],
        action: (v) => { tracking = v === 'suggesting'; applyMode(v); ribbon.setOn('track', tracking); },
      },
      { id: 'side', icon: 'sidebar', title: 'Show side panel', action: () => (sidePanel ? hideSide() : showSide('nav')) },
    ],
  });
  let rulerOn = true;
  globalThis.__kertasRulerUnit = rulerUnit();
  let marksOn = false;
  ribbon.setOn('ruler', true);

  // ---------------------------------------------------------------- start SuperDoc
  // Word citation fields become editable text here and fields again on save
  let openedCitations = [];
  const openedBytes = data || null; // the file as opened, for the repair pass on save
  if (data) {
    try { data = await prepareDocx(data); } catch (e) { console.warn('prepare', e); }
    try { const r = await fieldsToCitations(data); data = r.bytes; openedCitations = r.found; } catch (e) { console.warn('citations', e); }
  }
  const file = data
    ? new File([data], name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
    : null;

  const kFonts = superdocFonts();
  const spellPrefs = () => { try { return JSON.parse(localStorage.getItem('kertas-spell') || '{}'); } catch { return {}; } };
  const saveSpellPrefs = (p) => { try { localStorage.setItem('kertas-spell', JSON.stringify({ ...spellPrefs(), ...p })); } catch { /* storage off */ } };
  let spellOn = spellPrefs().on !== false;
  const spellProvider = createSpellProvider(() => spellPrefs().lang || null);
  // The editor also reports updates that leave the text as it was (for example when a font
  // finishes loading and the page is laid out again); only a real change marks the file unsaved.
  const seenDocs = new WeakMap(); // editor -> the document as it was last seen
  function contentChanged(payload) {
    // headers, footers and notes have editors of their own: a change there always counts
    if (payload?.surface && payload.surface !== 'body') return true;
    const ed = payload?.editor || superdoc?.activeEditor;
    const doc = ed?.state?.doc;
    if (!ed || !doc) return true;
    const prev = seenDocs.get(ed);
    seenDocs.set(ed, doc);
    return !prev || !prev.eq(doc);
  }
  superdoc = await new Promise((resolve, reject) => {
    let sd;
    try {
      sd = new SuperDoc({
        selector: `#${id}-editor`,
        toolbar: `#${id}-toolbar`,
        modules: {
          toolbar: { selector: `#${id}-toolbar`, excludeItems: ['documentMode', 'zoom', 'search', 'ruler', 'measurementUnit'], fonts: kFonts.options },
          contextMenu: { includeDefaultItems: true, noSlash: true, customItems: pictureMenuItems() }, // "/" opens Kertas's own menu
        },
        fonts: { families: kFonts.families },
        proofing: {
          enabled: spellOn,
          provider: spellProvider,
          defaultLanguage: 'en-US',
          allowIgnoreWord: true,
          ignoredWords: personalWords(),
          maxSuggestions: 6,
          onProofingError: (e) => console.warn('spelling', e?.message),
        },
        document: file, // null opens a blank document
        documentMode: 'editing',
        pagination: true,
        rulers: true,
        user: editorUser(),
        telemetry: { enabled: false }, // files stay on this computer
        onReady: () => { ready = true; try { const ed = sd.activeEditor; if (ed?.state?.doc) seenDocs.set(ed, ed.state.doc); } catch { /* not ready */ } el.querySelector('.writer-loading')?.remove(); stopBoxes = enableClickableBoxes(el.querySelector('.writer-body'), () => superdoc?.activeEditor, () => setDirty(true)); decorations = createDecorations(el.querySelector('.writer-body'), () => superdoc?.activeEditor); setTimeout(wireEditorEvents, 0); resolve(sd); },
        onEditorUpdate: (payload) => { if (ready && !suppressDirty && contentChanged(payload)) setDirty(true); scheduleStats(); scheduleSide(); },
        onContentError: (e) => { console.warn('content error', e); toast('Some content in this document could not be read', true); },
        onException: (e) => { console.error(e); if (!ready) reject(e?.error || new Error('Failed to open document')); },
        onPaginationUpdate: (p) => { pages = p?.totalPages ?? p?.pageCount ?? pages; scheduleStats(); },
      });
    } catch (err) { reject(err); }
    setTimeout(() => { if (!ready) reject(new Error('Timed out opening document')); }, 60000);
  });

  ribbon.setOn('spell', spellOn);

  // New documents get Word's modern defaults: Calibri 11, 8pt after paragraphs, 1.08 lines
  if (!data) {
    try {
      await api().styles.apply({ target: { scope: 'docDefaults', channel: 'run' }, patch: { fontFamily: { ascii: 'Calibri', hAnsi: 'Calibri', cs: 'Calibri', eastAsia: 'Calibri' }, fontSize: 22 } });
      await api().styles.apply({ target: { scope: 'docDefaults', channel: 'paragraph' }, patch: { spacing: { before: 0, after: 160, line: 259, lineRule: 'auto' } } });
    } catch (e) { console.warn('defaults', e); }
  }

  // Text imported from an older format (.doc without LibreOffice)
  if (!data && html) {
    try { await api().insert({ type: 'html', value: html }); } catch (e) { console.warn('import', e); toast('Some of the text could not be imported', true); }
  }
  setTimeout(() => {
    suppressDirty = false;
    updateStats();
  }, 400);

  // ---------------------------------------------------------------- Home
  async function clip(action) {
    superdoc.focus?.();
    try { await editCommand(action); } catch { toast('Use the keyboard shortcut for this in the browser version', true); }
  }
  /** A run property across the selection: its value when every run agrees, otherwise undefined */
  function selectionRunProp(key) {
    const ed = superdoc?.activeEditor;
    if (!ed?.state) return undefined;
    const { from, to, empty } = ed.state.selection;
    const vals = new Set();
    const read = (run) => vals.add(run.attrs?.runProperties?.[key] ?? null);
    if (empty) { const $f = ed.state.selection.$from; for (let d = $f.depth; d > 0; d--) if ($f.node(d).type.name === 'run') { read($f.node(d)); break; } }
    else ed.state.doc.nodesBetween(from, to, (n) => { if (n.type.name === 'run') { if (n.textContent) read(n); return false; } return true; });
    return vals.size === 1 ? [...vals][0] : undefined;
  }
  async function toggleInline(key) {
    await act('Format', async () => {
      const s = await textSelection();
      // Caps is stored as an uppercase text transform; "off" is false (1.47.2 ignores null here)
      const on = key === 'caps' ? selectionRunProp('textTransform') === 'uppercase' : !!selectionRunProp(key);
      return api().format.apply({ target: toSelectionTarget(s), inline: { [key]: on ? false : true } });
    });
  }
  async function toggleVert(kind) {
    await act('Format', async () => {
      const s = await textSelection();
      const cur = selectionRunProp('vertAlign');
      return api().format.apply({ target: toSelectionTarget(s), inline: { vertAlign: cur === kind ? null : kind } });
    });
  }
  async function clearFormatting() {
    await act('Clear formatting', async () => {
      const s = await textSelection();
      const keys = ['bold', 'italic', 'underline', 'strike', 'color', 'highlight', 'fontFamily', 'fontSize', 'vertAlign', 'smallCaps', 'caps', 'letterSpacing', 'dstrike'];
      return api().format.apply({ target: toSelectionTarget(s), inline: Object.fromEntries(keys.map((k) => [k, null])) });
    });
  }
  async function applyStyle(styleId, label) {
    await act(`${label} style`, () => onSelectedBlocks((t) => api().styles.paragraph.setStyle({ target: t, styleId })));
  }
  function allStylesMenu() {
    const list = paragraphStyles(editor());
    return [
      { label: 'Create a style...', icon: 'plus', hint: 'From the formatting at the cursor', action: () => newStyle() },
      'sep',
      ...list.slice(0, 40).map((st) => ({ label: st.name, icon: 'paragraph', action: () => applyStyle(st.id, st.name) })),
    ];
  }
  async function newStyle() {
    const res = await openDialog({
      title: 'Create a style',
      description: 'The new style takes the font, size, colour, spacing and alignment of the paragraph with the cursor. It is saved in the document and works in Word.',
      width: 420,
      fields: [{ id: 'name', label: 'Name', value: '', placeholder: 'e.g. Client note' }],
      validate: (v) => (!v.name.trim() ? 'Type a name' : paragraphStyles(editor()).some((st) => st.name.toLowerCase() === v.name.trim().toLowerCase()) ? 'A style with that name exists' : null),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Create', primary: true }],
    });
    if (!res) return;
    const name = res.values.name.trim();
    let id;
    try { id = createStyleFromCursor(editor(), name); } catch (e) { toast(`Create a style: ${e.message}`, true); return; }
    setDirty(true);
    await applyStyle(id, name);
    toast(`Style "${name}" created. Find it under All styles.`);
  }
  async function shade(color) {
    const fill = !color || color === 'none' ? 'auto' : color.replace('#', '').toUpperCase();
    await act('Shading', () => onSelectedBlocks((t) => api().format.paragraph.setShading({ target: t, fill })));
  }
  // ---------------------------------------------------------------- Home tab formatting
  // The Font and Paragraph controls drive SuperDoc's own formatting commands (the same ones its
  // toolbar uses), and follow the cursor through the toolbar's live state.
  const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const tbc = () => superdoc?.toolbar?.controller || null;
  function fmt(cmd, payload) {
    const c = tbc();
    if (!c) return;
    // a typed font or size moved focus to the ribbon box; the editor takes it back first
    try { editor()?.focus?.(); } catch { /* fine */ }
    try { c.execute(cmd, payload); } catch (e) { toast(String(e?.message || e), true); }
    setTimeout(syncHome, 0);
  }
  const currentSize = () => parseFloat(String(tbc()?.getSnapshot?.()?.commands?.['font-size']?.value || '11')) || 11;
  function setFontSize(v) {
    const n = parseFloat(String(v).replace(',', '.'));
    if (!Number.isFinite(n) || n < 1 || n > 1638) { toast('Enter a size between 1 and 1638', true); syncHome(); return; }
    fmt('font-size', `${n}pt`);
  }
  function growFont(dir) {
    const cur = currentSize();
    const next = dir > 0 ? (FONT_SIZES.find((n) => n > cur) ?? cur + 10) : ([...FONT_SIZES].reverse().find((n) => n < cur) ?? Math.max(1, cur - 1));
    fmt('font-size', `${next}pt`);
  }
  function caseMenu() {
    const marks = editor()?.state?.selection?.$from?.marks?.() || [];
    const has = (k) => marks.some((m) => m.type.name === 'textStyle' && m.attrs?.[k]);
    return [
      { label: 'All caps', checked: has('textTransform') || has('caps'), action: () => toggleInline('caps') },
      { label: 'Small caps', checked: has('smallCaps'), action: () => toggleInline('smallCaps') },
      { label: 'Double strikethrough', checked: has('dstrike'), action: () => toggleInline('dstrike') },
    ];
  }
  function spacingMenu() {
    const cur = Number(tbc()?.getSnapshot?.()?.commands?.['line-height']?.value) || null;
    return [
      { header: 'Line spacing' },
      ...[1, 1.08, 1.15, 1.5, 2, 2.5, 3].map((v) => ({ label: v === 1.08 ? '1.08 (Word default)' : String(v.toFixed(v % 1 ? 2 : 1)).replace(/0$/, ''), checked: cur !== null && Math.abs(cur - v) < 0.01, action: () => fmt('line-height', v) })),
      'sep',
      { label: 'Paragraph settings...', icon: 'paragraph', action: () => paragraphSettings() },
    ];
  }
  let homeCtl = null; let homeUnsub = null;
  function syncHome() {
    const c = tbc();
    if (!c) return;
    if (c !== homeCtl) { try { homeUnsub?.(); } catch { /* old one gone */ } homeCtl = c; homeUnsub = c.subscribe?.(() => syncHomeSoon()); }
    const cmd = c.getSnapshot?.()?.commands || {};
    const on = (k) => !!cmd[k]?.active;
    ribbon.setOn('b', on('bold')); ribbon.setOn('i', on('italic')); ribbon.setOn('u', on('underline')); ribbon.setOn('s', on('strikethrough'));
    ribbon.setOn('ul', on('bullet-list')); ribbon.setOn('ol', on('numbered-list'));
    const align = cmd['text-align']?.value || 'left';
    ribbon.setOn('al', align === 'left'); ribbon.setOn('ac', align === 'center'); ribbon.setOn('ar', align === 'right'); ribbon.setOn('aj', align === 'justify' || align === 'both');
    ribbon.setOn('painter', on('copy-format'));
    ribbon.setDisabled('undo', !!cmd.undo?.disabled); ribbon.setDisabled('redo', !!cmd.redo?.disabled);
    const fam = String(cmd['font-family']?.value || '').split(',')[0].replace(/["']/g, '').trim();
    ribbon.setValue('fontname', fam);
    const size = cmd['font-size']?.value;
    ribbon.setValue('fontsize', size ? String(parseFloat(size)) : '');
  }
  let homeTimer = null;
  const syncHomeSoon = () => { if (homeTimer) return; homeTimer = requestAnimationFrame(() => { homeTimer = null; syncHome(); }); };
  function borderMenu() {
    const set = (side) => act('Borders', () => onSelectedBlocks((t) => api().format.paragraph.setBorder({ target: t, side, style: 'single', size: 6, space: 4, color: '000000' })));
    return [
      { label: 'Bottom border', icon: 'border-bottom', action: () => set('bottom') },
      { label: 'Top border', icon: 'border-top', action: () => set('top') },
      { label: 'Box', icon: 'border', action: async () => { for (const s of ['top', 'bottom', 'left', 'right']) await set(s); } },
      'sep',
      { label: 'No border', icon: 'close', action: () => act('Borders', () => onSelectedBlocks((t) => api().format.paragraph.clearBorder({ target: t, side: 'all' }))) },
    ];
  }
  function toggleMarks() {
    superdoc.toggleFormattingMarks();
    marksOn = !marksOn;
    ribbon.setOn('marks', marksOn);
    ribbon.setOn('marks2', marksOn);
  }
  // ---------------------------------------------------------------- Spelling
  function toggleSpelling() {
    spellOn = !spellOn;
    saveSpellPrefs({ on: spellOn });
    ribbon.setOn('spell', spellOn);
    toast(spellOn ? 'Spell check is on for documents you open from now on' : 'Spell check is off for documents you open from now on');
  }
  function spellLanguageMenu() {
    const cur = spellPrefs().lang || null;
    return [
      { label: 'Follow the document', checked: !cur, action: () => { saveSpellPrefs({ lang: null }); toast('Spelling follows each paragraph\'s language'); } },
      ...SPELL_LANGUAGES.map((l) => ({ label: l.label, checked: cur === l.id, action: () => { saveSpellPrefs({ lang: l.id }); toast(`Checking spelling in ${l.label}`); } })),
    ];
  }
  async function addWordToDictionary() {
    const s = await selection().catch(() => null);
    let word = '';
    try { word = String((await api().getText({ target: (s?.selectionTarget || s?.target) }))?.text ?? '').trim(); } catch { word = ''; }
    if (!word || /\s/.test(word)) { toast('Select one word first'); return; }
    addToPersonal(word);
    toast(`"${word}" added to your dictionary. It is no longer marked in documents you open.`);
  }
  async function manageDictionary() {
    const words = personalWords();
    const res = await openDialog({
      title: 'My dictionary',
      description: 'Words you added are never marked as misspelled. One word per line.',
      fields: [{ id: 'words', label: 'Words', type: 'textarea', rows: 10, value: words.join('\n') }],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Save', primary: true }],
      width: 420,
    });
    if (!res) return;
    const list = res.values.words.split(/\n+/).map((w) => w.trim()).filter(Boolean);
    try { localStorage.setItem('kertas-personal-dictionary', JSON.stringify([...new Set(list)].sort())); } catch { /* storage off */ }
    toast(`${list.length} word${list.length === 1 ? '' : 's'} in your dictionary`);
  }

  async function setIndent(side, cm) {
    if (cm === null || !Number.isFinite(cm)) return;
    await act('Indent', () => onSelectedBlocks((t) => api().format.paragraph.setIndentation({ target: t, [side]: Math.round(cm * CM) })));
  }
  async function setSpace(which, pt) {
    if (pt === null || !Number.isFinite(pt)) return;
    await act('Spacing', () => onSelectedBlocks((t) => api().format.paragraph.setSpacing({ target: t, [which]: Math.round(pt * PT) })));
  }
  async function setDirection(direction) {
    // Right-to-left paragraphs also take right alignment, as in Word
    await act('Text direction', () => {
      const ed = superdoc?.activeEditor;
      if (!ed?.state) throw new Error('The document is still opening');
      const { state } = ed; const { from, to } = state.selection; const rtl = direction === 'rtl';
      const tr = state.tr; let n = 0;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isTextblock || !node.type.spec.attrs?.paragraphProperties) return true;
        const pp = { ...(node.attrs.paragraphProperties || {}) };
        pp.rightToLeft = rtl;
        const j = pp.justification;
        // Left/right alignment follows the direction (Word reads it as start/end), so it is cleared
        if (j === 'left' || j === 'right' || j === 'start' || j === 'end') delete pp.justification;
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, paragraphProperties: pp });
        n++; return false;
      });
      if (!n) throw new Error('Click in a paragraph first');
      ed.view.dispatch(tr);
      return { success: true };
    });
    ribbon.setOn('rtl', direction === 'rtl'); ribbon.setOn('rtl2', direction === 'rtl');
    ribbon.setOn('ltr', direction === 'ltr'); ribbon.setOn('ltr2', direction === 'ltr');
  }
  async function paragraphSettings() {
    const list = await selectedBlocks().catch(() => []);
    if (!list.length) { toast('Click in a paragraph first'); return; }
    const v = await paragraphDialog();
    if (!v) return;
    await act('Paragraph', () => onSelectedBlocks(async (t) => {
      const p = api().format.paragraph;
      if (v.align !== 'keep') await p.setAlignment({ target: t, alignment: v.align });
      const indent = { target: t, left: Math.round((v.indentLeft || 0) * CM), right: Math.round((v.indentRight || 0) * CM) };
      if (v.special !== 'none') indent[v.special] = Math.round((v.specialBy || 0) * CM);
      await p.setIndentation(indent);
      await p.setSpacing({ target: t, before: Math.round((v.before || 0) * PT), after: Math.round((v.after || 0) * PT), line: Math.round(Number(v.line) * 240), lineRule: 'auto' });
      await p.setKeepOptions({ target: t, keepNext: v.keepNext, keepLines: v.keepLines });
      return p.setFlowOptions({ target: t, pageBreakBefore: v.breakBefore });
    }));
  }

  // Find and replace
  const findEl = el.querySelector('.find-panel');
  const findInput = findEl.querySelector('[data-find]');
  const replaceInput = findEl.querySelector('[data-replace]');
  // Find and replace on the editor's own search commands (matches are highlighted in the page)
  const findState = { total: 0, activeIndex: 0 };
  const search = () => {
    const ed = superdoc.activeEditor;
    const take = (r) => { if (r && Array.isArray(r.matches)) findState.total = r.matches.length; if (r && Number.isInteger(r.activeMatchIndex)) findState.activeIndex = Math.max(0, r.activeMatchIndex); return { ...findState }; };
    return {
      query: ({ query }) => {
        if (!query) { ed.commands.clearSearchSession?.(); findState.total = 0; return { total: 0 }; }
        const r = ed.commands.setSearchSession(query, { highlight: true });
        if (r?.matches?.[0]) ed.commands.goToSearchResult(r.matches[0]);
        return take(r);
      },
      next: () => take(ed.commands.nextSearchMatch()),
      previous: () => take(ed.commands.previousSearchMatch()),
      getState: () => ({ ...findState }),
      replace: ({ replacement }) => take(ed.commands.replaceSearchMatch(replacement)),
      replaceAll: ({ query, replacement }) => {
        const r = ed.commands.setSearchSession(query, { highlight: true });
        const n = r?.matches?.length || 0;
        ed.commands.replaceAllSearchMatches(replacement);
        findState.total = 0;
        return { replaced: n };
      },
    };
  };
  function openFind(withReplace) {
    findEl.hidden = false;
    findEl.classList.toggle('with-replace', withReplace);
    const s = window.getSelection()?.toString();
    if (s && s.length < 80) findInput.value = s;
    findInput.focus();
    findInput.select();
    if (findInput.value) runFind();
  }
  function closeFind() { findEl.hidden = true; try { search().query({ query: '' }); } catch { /* noop */ } superdoc.focus?.(); }
  async function runFind() {
    const q = findInput.value;
    if (!q) { findEl.querySelector('.find-count').textContent = ''; return; }
    const r = await search().query({ query: q });
    updateFindCount(r);
  }
  function updateFindCount(r) {
    const st = r?.total !== undefined ? r : search().getState();
    findEl.querySelector('.find-count').textContent = st.total ? `${(st.activeIndex ?? 0) + 1} of ${st.total}` : 'No results';
  }
  let findTimer;
  findInput.addEventListener('input', () => { clearTimeout(findTimer); findTimer = setTimeout(runFind, 200); });
  findEl.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') closeFind();
    if (e.key === 'Enter' && e.target === findInput) { e.preventDefault(); (e.shiftKey ? search().previous() : search().next()); updateFindCount(); }
    if (e.key === 'Enter' && e.target === replaceInput) { e.preventDefault(); replaceOne(); }
  });
  findEl.querySelector('[data-find-next]').addEventListener('click', () => { search().next(); updateFindCount(); });
  findEl.querySelector('[data-find-prev]').addEventListener('click', () => { search().previous(); updateFindCount(); });
  findEl.querySelector('[data-find-close]').addEventListener('click', closeFind);
  async function replaceOne() {
    if (!findInput.value) return;
    await search().replace({ replacement: replaceInput.value });
    setDirty(true);
    await runFind();
  }
  findEl.querySelector('[data-replace-one]').addEventListener('click', replaceOne);
  findEl.querySelector('[data-replace-all]').addEventListener('click', async () => {
    if (!findInput.value) return;
    const r = await search().replaceAll({ query: findInput.value, replacement: replaceInput.value });
    setDirty(true);
    toast(`Replaced ${r?.replaced ?? 0} occurrence${r?.replaced === 1 ? '' : 's'}`);
    await runFind();
  });

  // ---------------------------------------------------------------- Insert
  async function pageBreak() {
    await act('Page break', async () => {
      const p = await newParagraphAfter('');
      return onBlock((t) => api().format.paragraph.setFlowOptions({ target: t, pageBreakBefore: true }), p);
    });
  }
  async function blankPage() {
    await act('Blank page', async () => {
      const a = await newParagraphAfter('');
      await onBlock((t) => api().format.paragraph.setFlowOptions({ target: t, pageBreakBefore: true }), a);
      const b = await onBlock((t) => api().create.paragraph({ at: { kind: 'after', target: t }, text: '' }), a);
      return onBlock((t) => api().format.paragraph.setFlowOptions({ target: t, pageBreakBefore: true }), b.paragraph);
    });
  }
  async function sectionBreak(breakType) {
    await act('Section break', async () => {
      const b = await caretBlock();
      return onBlock((t) => api().create.sectionBreak({ at: { kind: 'after', target: t }, breakType }), b);
    }, 'Section break inserted');
  }
  async function insertTable() {
    const b = await caretBlock().catch(() => null);
    const v = await tableDialog();
    if (!v) return;
    await act('Table', () => api().create.table({ rows: v.rows, columns: v.columns, ...(b ? { at: { kind: 'after', nodeId: b.nodeId } } : {}) }));
  }
  async function insertPicture() {
    const b = await caretBlock().catch(() => null);
    const [img] = await openFiles('image', false);
    if (!img) return;
    const type = /\.png$/i.test(img.name) ? 'image/png' : 'image/jpeg';
    const src = await new Promise((r) => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(new Blob([img.data], { type })); });
    const [w, h] = await new Promise((r) => { const im = new Image(); im.onload = () => r([im.naturalWidth, im.naturalHeight]); im.onerror = () => r([400, 300]); im.src = src; });
    const s = Math.min(1, 576 / w); // fit within 6 inches
    await act('Picture', () => api().create.image({ src, alt: img.name, size: { width: Math.round(w * s), height: Math.round(h * s), unit: 'px' }, ...(b ? { at: { kind: 'after', target: b } } : {}) }));
  }
  async function horizontalLine() {
    await act('Line', async () => {
      const ed = superdoc.activeEditor;
      if (!ed?.commands?.insertHorizontalRule) throw new Error('Not available');
      // Below the current paragraph, never splitting it (as in Word)
      const $f = ed.state.selection.$from;
      for (let d = $f.depth; d > 0; d--) if ($f.node(d).isTextblock) { const end = $f.end(d); ed.commands.setTextSelection({ from: end, to: end }); break; }
      ed.commands.insertHorizontalRule();
      ed.view?.focus?.();
      return { success: true };
    });
  }
  async function insertLink() {
    const s = await selection().catch(() => null);
    if (!s?.target) { toast('Click in the document first'); return; }
    const seg = s.target.segments?.[0];
    const v = await linkDialog('');
    if (!v) return;
    const link = { destination: { href: v.href.trim() } };
    await act('Link', () => {
      if (!s.empty && seg && !v.text.trim()) return api().hyperlinks.wrap({ target: { kind: 'text', blockId: seg.blockId, range: seg.range }, link });
      return api().hyperlinks.insert({ target: { kind: 'text', blockId: seg.blockId, range: { start: seg.range.end, end: seg.range.end } }, text: v.text.trim() || v.href.trim(), link });
    });
  }
  async function newComment() {
    const s = await selection().catch(() => null);
    if (!s || s.empty) { toast('Select the text you want to comment on'); return; }
    await ensureUserName();
    const text = await commentDialog();
    if (!text) return;
    const r = await act('Comment', () => api().comments.create({ target: s.target, text }));
    if (r) showSide('comments');
  }
  // Enter Word-style header/footer editing on the page in view
  async function editFurniture(kind) {
    const pagesEls = [...el.querySelectorAll('.superdoc-page')];
    const body = el.querySelector('.writer-body').getBoundingClientRect();
    const page = pagesEls.find((p) => { const r = p.getBoundingClientRect(); return r.bottom > body.top + 80 && r.top < body.bottom - 80; }) || pagesEls[0];
    if (!page) return;
    page.scrollIntoView({ block: kind === 'header' ? 'start' : 'end' });
    await new Promise((r) => setTimeout(r, 120));
    const r = page.getBoundingClientRect();
    const x = r.left + r.width / 2;
    const y = kind === 'header' ? r.top + 36 : r.bottom - 36;
    const target = document.elementFromPoint(x, y) || page;
    for (const type of ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click', 'pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click', 'dblclick']) {
      const Ctor = type.startsWith('pointer') ? PointerEvent : MouseEvent;
      target.dispatchEvent(new Ctor(type, { bubbles: true, cancelable: true, clientX: x, clientY: y, detail: type === 'dblclick' ? 2 : 1, button: 0, buttons: type.endsWith('down') ? 1 : 0, pointerType: 'mouse', isPrimary: true }));
    }
    setStatus(`Editing the ${kind}. Press Esc or double-click the page body to return.`);
  }
  // Page numbers are live PAGE / NUMPAGES fields in the header or footer, as Word writes them
  const hfSlot = (section, kind) => ({ kind: 'headerFooterSlot', section, headerFooterKind: kind, variant: 'default' });
  const hfStory = (section, kind) => ({ kind: 'story', storyType: 'headerFooterSlot', section, headerFooterKind: kind, variant: 'default' });
  async function pageNumbers() {
    const v = await pageNumberDialog();
    if (!v) return;
    if (v === 'remove') {
      await act('Page numbers', () => onSections(async (t) => {
        for (const kind of ['header', 'footer']) {
          try { api().headerFooters.refs.clear({ target: hfSlot(t, kind) }); } catch { /* none there */ }
        }
        return { success: true };
      }), 'Page numbers removed');
      return;
    }
    await act('Page numbers', () => onSections(async (t) => {
      const kind = v.kind === 'header' ? 'header' : 'footer';
      const num = (type) => ({ kind: 'field', field: { fieldType: type, instruction: type, resultText: '1' } });
      const run = (text) => ({ kind: 'run', run: { text } });
      const inlines = v.style === 'of' ? [run('Page '), num('PAGE'), run(' of '), num('NUMPAGES')]
        : v.style === 'page' ? [run('Page '), num('PAGE')] : [num('PAGE')];
      const r0 = api().sections.setPageNumbering({ target: t, format: v.format, start: Number.isFinite(v.start) ? v.start : 1 });
      if (failed(r0)) return r0;
      // A fresh header/footer holding just the number (Word's built-in page numbers do the same)
      try { api().headerFooters.refs.clear({ target: hfSlot(t, kind) }); } catch { /* nothing to clear */ }
      const r = api().insert({ in: hfStory(t, kind), content: { kind: 'paragraph', paragraph: { props: { alignment: v.alignment || 'center' }, inlines } } });
      if (failed(r)) return r;
      // The new part starts with an empty line; the number takes its place
      const empty = r?.resolution?.target?.blockId;
      if (empty) { try { api().blocks.delete({ target: { kind: 'block', nodeType: 'paragraph', nodeId: empty, story: hfStory(t, kind) }, in: hfStory(t, kind) }); } catch { /* an empty line is harmless */ } }
      try { api().sections.setTitlePage({ target: t, enabled: !v.first }); } catch { /* fine */ }
      return { success: true };
    }), 'Page numbers added');
  }
  async function insertDate() {
    const text = await dateTimeDialog();
    if (text) await act('Date', () => insertTextAtCaret(text));
  }
  function symbolMenu() {
    return { grid: SYMBOLS, onPick: (c) => act('Symbol', () => insertTextAtCaret(c)) };
  }

  // ---------------------------------------------------------------- Layout
  async function getSection() { return (await sections())[0]; }
  async function setMargins(m) { await act('Margins', () => onSections((t) => api().sections.setPageMargins({ target: t, ...m }))); }
  async function customMargins() {
    const s = await getSection();
    const v = await marginsDialog(s?.margins);
    if (v) await setMargins(v);
  }
  async function setOrientation(orientation) {
    await act('Orientation', () => onSections((t, s) => {
      const { width = 8.5, height = 11 } = s.pageSetup || {};
      const [w, h] = orientation === 'landscape' ? [Math.max(width, height), Math.min(width, height)] : [Math.min(width, height), Math.max(width, height)];
      return api().sections.setPageSetup({ target: t, orientation, width: w, height: h });
    }));
  }
  async function setPaper(p) {
    await act('Paper size', () => onSections((t, s) => {
      const land = s.pageSetup?.orientation === 'landscape';
      return api().sections.setPageSetup({ target: t, width: land ? p.h : p.w, height: land ? p.w : p.h, orientation: land ? 'landscape' : 'portrait' });
    }), `${p.label} paper`);
  }
  async function customPaper() {
    const v = await paperSizeDialog();
    if (v) await setPaper({ label: 'Custom', w: Math.min(v.width, v.height), h: Math.max(v.width, v.height) });
  }
  async function setColumns(count) { await act('Columns', () => onSections((t) => api().sections.setColumns({ target: t, count, gap: 0.5, equalWidth: true }))); }
  async function lineNumbers(enabled, restart) {
    await act('Line numbers', () => onSections((t) => api().sections.setLineNumbering({ target: t, enabled, ...(enabled ? { countBy: 1, start: 1, restart } : {}) })));
  }
  // Watermarks are a pale picture centred on the page behind the text, placed in the header
  // (Word's "picture watermark"), so Word, LibreOffice and PDF all show them.
  const watermarkIds = new Set();
  function watermarkPng({ text, orientation, color, opacity }) {
    const W = 1400; const H = orientation === 'diagonal' ? 1400 : 420;
    const c = document.createElement('canvas'); c.width = W; c.height = H;
    const g = c.getContext('2d');
    g.translate(W / 2, H / 2);
    if (orientation === 'diagonal') g.rotate(-Math.PI / 4);
    let size = 260;
    g.font = `bold ${size}px Calibri, Carlito, Arial, sans-serif`;
    const maxW = orientation === 'diagonal' ? W * 1.2 : W * 0.95;
    while (g.measureText(text).width > maxW && size > 40) { size -= 10; g.font = `bold ${size}px Calibri, Carlito, Arial, sans-serif`; }
    g.globalAlpha = Math.max(0.05, Math.min(1, opacity));
    g.fillStyle = color;
    g.textAlign = 'center'; g.textBaseline = 'middle';
    g.fillText(text, 0, 0);
    return { src: c.toDataURL('image/png'), ratio: H / W };
  }
  /** A watermark picture: Word's text watermark (a VML shape), its picture watermark, or one made here */
  const isWatermark = (n) => {
    if (n.type.name !== 'image') return false;
    const a = n.attrs || {};
    return !!a.vmlWatermark || /^Watermark/i.test(a.alt || '') || /PowerPlusWaterMark|WordPictureWatermark/i.test(`${a.vmlAttributes?.id || ''} ${a.name || ''} ${a.title || ''}`);
  };
  /** Takes watermarks out of every header, including ones that came with the file */
  function scrubHeaderWatermarks() {
    const ed = superdoc?.activeEditor;
    const pe = ed?.presentationEditor;
    let removed = 0;
    for (const kind of ['header', 'footer']) {
      for (const refId of Object.keys(ed?.converter?.[`${kind}s`] || {})) {
        let he = null;
        try { const x = pe?.ensureHeaderFooterEditor?.(kind, refId); he = x?.editor || x; } catch { /* no such part */ }
        if (!he?.state || !he.view) continue;
        const spots = [];
        he.state.doc.descendants((n, pos) => { if (isWatermark(n)) spots.push([pos, n.nodeSize]); });
        if (!spots.length) continue;
        let tr = he.state.tr;
        for (const [pos, size] of spots.reverse()) tr = tr.delete(pos, pos + size);
        he.view.dispatch(tr);
        removed += spots.length;
      }
    }
    return removed;
  }
  async function removeWatermarks() {
    let removed = 0;
    for (const s of await sections()) {
      const story = hfStory(s.address, 'header');
      for (const id of [...watermarkIds]) {
        try { const r = api().blocks.delete({ target: { kind: 'block', nodeType: 'paragraph', nodeId: id, story }, in: story }); if (!failed(r)) { removed++; watermarkIds.delete(id); } } catch { /* not in this section */ }
      }
    }
    removed += scrubHeaderWatermarks();
    if (removed) setDirty(true);
    return { success: true, removed };
  }
  async function setWatermark({ text, orientation = 'diagonal', color = '#C0C0C0', opacity = 0.5 }) {
    await act('Watermark', async () => {
      await removeWatermarks();
      const { src, ratio } = watermarkPng({ text, orientation, color, opacity });
      const width = 470; const height = Math.round(width * ratio); // points-ish at 96 dpi: about 6.5 inches wide
      for (const s of await sections()) {
        const id = `kertas-watermark-${Math.random().toString(36).slice(2, 10)}`;
        const image = { src, alt: `Watermark: ${text}`, geometry: { width, height }, pmAttrs: { isAnchor: true, anchorData: { hRelativeFrom: 'page', vRelativeFrom: 'page', alignH: 'center', alignV: 'center' }, wrap: { type: 'None', attrs: { behindDoc: true } }, size: { width, height, unit: 'px' } } };
        const r = api().insert({ in: hfStory(s.address, 'header'), content: { kind: 'paragraph', id, paragraph: { inlines: [{ kind: 'image', image }] } } });
        if (failed(r)) return r;
        watermarkIds.add(id);
      }
      return { success: true };
    }, `Watermark "${text}" added`);
  }
  async function customWatermark() {
    const v = await watermarkTextDialog();
    if (v) await setWatermark({ text: v.text.trim(), orientation: v.orientation, color: v.color, opacity: v.opacity / 100 });
  }
  async function pageBorder(style, size) {
    await act('Page borders', () => onSections((t) => {
      if (!style) return api().sections.clearPageBorders({ target: t });
      const side = { style, size, space: 24, color: '000000' };
      return api().sections.setPageBorders({ target: t, borders: { display: 'allPages', offsetFrom: 'page', top: side, bottom: side, left: side, right: side } });
    }));
  }

  // ---------------------------------------------------------------- References
  async function insertToc() {
    const b = await caretBlock().catch(() => null);
    // the cursor may be in a heading or a list item: the target has to name the right kind
    await act('Table of contents', () => (b ? onBlock((t) => api().create.tableOfContents({ at: { kind: 'before', target: t } }), b) : api().create.tableOfContents({ at: { kind: 'documentStart' } })), 'Table of contents inserted');
  }
  async function updateToc() {
    const list = (await api().toc.list()).items || [];
    if (!list.length) { toast('There is no table of contents yet'); return; }
    await act('Update table', async () => { let r; for (const t of list) r = await api().toc.update({ target: t.address, mode: 'all' }); return r; }, 'Table of contents updated');
  }
  async function removeToc() {
    const list = (await api().toc.list()).items || [];
    if (!list.length) { toast('There is no table of contents'); return; }
    await act('Remove table', async () => { let r; for (const t of list) r = await api().toc.remove({ target: t.address }); return r; });
  }
  async function insertNote(type) {
    const s = await selection().catch(() => null);
    if (!s?.target) { toast('Click where the note reference should go'); return; }
    const text = await noteTextDialog(type);
    if (text) await act(type === 'endnote' ? 'Endnote' : 'Footnote', () => api().footnotes.insert({ type, content: text }));
  }
  async function insertCaption() {
    const b = await caretBlock().catch(() => null);
    if (!b) { toast('Click next to the table or picture first'); return; }
    const v = await captionDialog();
    if (v) await act('Caption', () => onBlock((t) => api().captions.insert({ adjacentTo: t, position: v.position, label: v.label, text: v.text ? ` ${v.text}` : '' }), b));
  }

  // ---------------------------------------------------------------- Review
  async function wordCount() {
    const info = await api().info();
    await wordCountDialog(info.counts || {}, pages || el.querySelectorAll('.superdoc-page').length);
  }
  /** "January → March" for a replacement; the excerpt for anything else */
  function changeText(c) {
    const plain = esc(String(c.excerpt || c.text || '').slice(0, 160));
    if (c.type !== 'replacement') return plain;
    const ed = superdoc?.activeEditor; let del = ''; let ins = '';
    ed?.state.doc.descendants((n) => {
      if (!n.isText) return true;
      for (const m of n.marks) {
        if (m.attrs?.id !== c.id) continue;
        if (m.type.name === 'trackDelete') del += n.text;
        if (m.type.name === 'trackInsert') ins += n.text;
      }
      return true;
    });
    return del || ins ? `<s class="tc-old">${esc(del.slice(0, 80))}</s> → <span class="tc-new">${esc(ins.slice(0, 80))}</span>` : plain;
  }
  async function setTracking(on) {
    if (on) await ensureUserName();
    tracking = on;
    applyMode(on ? 'suggesting' : 'editing');
    const sel = ribbon.get('mode');
    if (sel) sel.value = on ? 'suggesting' : 'editing';
    ribbon.setOn('track', on);
    toast(on ? 'Track changes is on. Your edits are recorded as suggestions.' : 'Track changes is off');
  }
  function markupView(mode) {
    superdoc.setTrackedChangesPreferences({ mode });
    toast({ review: 'Showing all markup', final: 'Showing the document as if every change were accepted', original: 'Showing the original document' }[mode]);
  }
  function wireEditorEvents() {
    setTimeout(syncHome, 0);
    if (openedCitations.length && superdoc?.activeEditor) { const f = openedCitations; openedCitations = []; setTimeout(() => extras.registerCitations(f).catch(() => {}), 300); }
    // The editor reports ready before the variable that holds it is assigned
    if (!superdoc?.activeEditor) { setTimeout(wireEditorEvents, 50); return; }
    const body = el.querySelector('.writer-body');
    if (!body.dataset.picDrag) { body.dataset.picDrag = '1'; enablePictureDrag(body, { editor, api, toast: (m, err) => toast(m, err), onMoved: (id) => setTimeout(() => { reselectPicture(id); updateTableTab(); }, 150) }); }
    const ed = editor();
    try { ed.on?.('selectionUpdate', updateTableTab); ed.on?.('transaction', updateTableTab); } catch { /* older editor */ }
    syncZoom();
    // page-width zoom changes with the window
    window.addEventListener('resize', () => setTimeout(syncZoom, 250));
    body.addEventListener('mouseup', () => setTimeout(updateTableTab, 0));
    body.addEventListener('keyup', () => setTimeout(updateTableTab, 0));
    // Ctrl + mouse wheel (or a trackpad pinch) zooms the document
    let wheelAcc = 0;
    body.addEventListener('wheel', (e) => {
      if (!(e.ctrlKey || e.metaKey)) return;
      e.preventDefault();
      wheelAcc += e.deltaY;
      if (Math.abs(wheelAcc) < 40) return;
      setZoom(currentZoom() + (wheelAcc < 0 ? 10 : -10));
      wheelAcc = 0;
    }, { passive: false });
    // Ctrl+click (Cmd+click on Mac) follows a link, as in Word
    body.addEventListener('pointerdown', (e) => {
      if (!(e.ctrlKey || e.metaKey) || e.button !== 0) return;
      const href = linkAt(e.clientX, e.clientY);
      if (!href) return;
      e.preventDefault(); e.stopPropagation();
      followLink(href);
    }, true);
    body.addEventListener('mousemove', (e) => { body.classList.toggle('k-link-hover', (e.ctrlKey || e.metaKey) && !!linkAt(e.clientX, e.clientY)); }, { passive: true });
    const keys = (e) => {
      if (!el.isConnected || !el.closest('.view')?.classList.contains('active')) return;
      if (document.querySelector('.modal')) return;
      if (e.key === 'F9' && !e.ctrlKey && !e.altKey) { e.preventDefault(); updateFields(); }
      if (e.altKey && e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) { e.preventDefault(); e.stopPropagation(); act('Move', () => moveBlock(e.key === 'ArrowUp' ? -1 : 1)); }
    };
    document.addEventListener('keydown', keys, true);
    unwire = () => document.removeEventListener('keydown', keys, true);
  }

  // ---------------------------------------------------------------- Table tab
  async function tcmd(label, name, ...args) {
    await act(label, () => {
      const ed = editor();
      if (!tableContext(ed)) throw new Error('Click in a table first');
      if (typeof ed.commands[name] !== 'function') throw new Error('Not available');
      if (ed.commands[name](...args) === false) throw new Error(name === 'mergeCells' ? 'Select two or more cells to merge' : 'Not possible here');
      ed.view?.focus?.();
      return { success: true };
    });
  }
  async function tapi(label, fn, okMsg) {
    await act(label, () => {
      const id = tableNodeId(editor(), api);
      if (!id) throw new Error('Click in a table first');
      return fn(id);
    }, okMsg);
  }
  async function sortTable(direction, opts) {
    const r = await act('Sort', () => sortRows(editor(), direction, opts));
    if (r) toast(`Sorted ${direction === 'ascending' ? 'A to Z' : 'Z to A'}${r.header ? ', keeping the header row on top' : ''}`);
  }
  async function sortDialog() {
    const ctx = tableContext(editor());
    if (!ctx) { toast('Click in a table first'); return; }
    const first = ctx.table.node.child(0);
    const cols = []; first.forEach((c, _o, i) => cols.push({ value: String(i), label: c.textContent.trim() ? `Column ${i + 1}: ${c.textContent.trim().slice(0, 24)}` : `Column ${i + 1}` }));
    const res = await openDialog({
      title: 'Sort table',
      width: 420,
      fields: [
        { id: 'col', label: 'Sort by', type: 'select', value: String(ctx.cellIndex), options: cols },
        { id: 'dir', label: 'Order', type: 'segmented', value: 'ascending', options: [{ value: 'ascending', label: 'A to Z, smallest first' }, { value: 'descending', label: 'Z to A, largest first' }] },
        { id: 'header', type: 'checkbox', label: 'My table has a header row', value: true },
      ],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Sort', primary: true }],
    });
    if (res) await sortTable(res.values.dir, { column: Number(res.values.col), header: !!res.values.header });
  }
  let lastPicKey = null;
  function updateTableTab() {
    const ed = editor();
    syncHomeSoon();
    const inside = !!tableContext(ed);
    ribbon.setTabVisible('table', inside);
    const node = ed?.state?.selection?.node;
    const pic = node?.type?.name === 'image' ? node : null;
    ribbon.setTabVisible('picture', !!pic);
    // clicking a picture brings up its tab, as long as it is a new picture being picked
    const picKey = pic ? pic.attrs?.sdImageId || ed.state.selection.from : null;
    if (pic && picKey !== lastPicKey) ribbon.activateTab('picture');
    lastPicKey = picKey;
    if (pic) syncPictureTab(pic);
    if (inside) syncTableLook();
  }
  // ---- tables: selecting, designs, style options
  function cellPositions() {
    const ctx = tableContext(editor());
    if (!ctx) return null;
    const rows = [];
    let p = ctx.table.pos + 1;
    ctx.table.node.forEach((row) => { const cells = []; let q = p + 1; row.forEach((cell) => { cells.push(q); q += cell.nodeSize; }); rows.push(cells); p += row.nodeSize; });
    return { ctx, rows };
  }
  function selectInTable(what) {
    const t = cellPositions();
    if (!t) { toast('Click in a table first', true); return; }
    const { ctx, rows } = t;
    const r = ctx.rowIndex; const c = Math.min(ctx.cellIndex, rows[r].length - 1);
    const last = rows.length - 1;
    const pick = { cell: [rows[r][c], rows[r][c]], row: [rows[r][0], rows[r][rows[r].length - 1]], column: [rows[0][Math.min(c, rows[0].length - 1)], rows[last][Math.min(c, rows[last].length - 1)]], table: [rows[0][0], rows[last][rows[last].length - 1]] }[what];
    const ed = editor();
    if (ed.commands.setCellSelection?.({ anchorCell: pick[0], headCell: pick[1] }) === false) toast('Could not select', true);
    ed.view?.focus?.();
  }
  async function applyDesign(d) {
    await tapi('Table design', (nodeId) => {
      ensureTableStyle(editor(), d);
      return api().tables.applyStyle({ nodeId, styleId: d.id, styleOptions: { headerRow: true, bandedRows: d.kind === 'header' || d.kind === 'light' } });
    });
    syncTableLook();
  }
  function designMenu(anchor) {
    const list = TABLE_DESIGNS;
    const m = openMenu(anchor, { grid: list.map((d) => d.id), html: list.map(designPreview), titles: list.map((d) => d.name), columns: 4, cell: 84, onPick: (id) => { closeOpenMenu(); applyDesign(list.find((d) => d.id === id)); } });
    m.classList.add('table-designs');
  }
  const LOOK = { 't-opt-header': 'headerRow', 't-opt-banded': 'bandedRows', 't-opt-first': 'firstColumn', 't-opt-total': 'lastRow' };
  let tableLook = {};
  async function syncTableLook() {
    try {
      const id = tableNodeId(editor(), api);
      const props = id ? await api().tables.getProperties({ nodeId: id }) : null;
      tableLook = props?.styleOptions || props?.properties?.styleOptions || {};
    } catch { tableLook = {}; }
    for (const [bid, flag] of Object.entries(LOOK)) ribbon.setOn(bid, !!tableLook[flag]);
  }
  async function toggleLook(bid) {
    const flag = LOOK[bid];
    const on = !tableLook[flag];
    await tapi('Table style options', (nodeId) => api().tables.applyStyle({ nodeId, styleOptions: { [flag]: on } }));
    syncTableLook();
  }
  // ---- pictures
  let picInfo = null;
  let NodeSel = null; // ProseMirror's NodeSelection, taken from a picture selection
  async function syncPictureTab(node) {
    NodeSel = editor().state.selection.constructor;
    const id = node.attrs?.sdImageId;
    try { picInfo = id ? await api().images.get({ imageId: id }) : null; } catch { picInfo = null; }
    const pr = picInfo?.properties || {};
    ribbon.setValue('pic-w', pr.size?.width ? Math.round(pr.size.width) : '');
    ribbon.setValue('pic-h', pr.size?.height ? Math.round(pr.size.height) : '');
  }
  const picId = () => { const n = editor()?.state?.selection?.node; return n?.type?.name === 'image' ? n.attrs.sdImageId : null; };
  async function picAct(label, fn, okMsg, given) {
    const id = given || picId();
    if (!id) { toast('Click a picture first', true); return; }
    await act(label, () => fn(id), okMsg);
    setTimeout(() => { reselectPicture(id); updateTableTab(); }, 120);
  }
  function reselectPicture(id) {
    const ed = editor(); let at = null;
    ed.state.doc.descendants((n, pos) => { if (at === null && n.type.name === 'image' && n.attrs.sdImageId === id) at = pos; return at === null; });
    if (at === null) return;
    if (!NodeSel) return;
    try { ed.view.dispatch(ed.state.tr.setSelection(NodeSel.create(ed.state.doc, at))); } catch { /* keep */ }
  }
  async function setWrap(type, behind = false, imageIdGiven = null) {
    await picAct('Wrap text', async (imageId) => {
      const cur = (await api().images.get({ imageId }))?.properties || {};
      if (type === 'Inline') return cur.placement === 'inline' ? { success: true } : api().images.convertToInline({ imageId });
      if (cur.placement === 'inline') {
        const f = await api().images.convertToFloating({ imageId });
        if (failed(f)) return f;
        await api().images.setPosition({ imageId, position: { hRelativeFrom: 'column', alignH: 'left', vRelativeFrom: 'paragraph', marginOffset: { top: 0 } } });
      }
      const w = await api().images.setWrapType({ imageId, type });
      if (failed(w)) return w;
      return api().images.setAnchorOptions({ imageId, options: { behindDoc: !!behind, allowOverlap: true } });
    }, undefined, imageIdGiven);
  }
  async function alignPicture(alignH) {
    await picAct('Position', async (imageId) => {
      const cur = (await api().images.get({ imageId }))?.properties || {};
      if (cur.placement === 'inline') {
        const f = await api().images.convertToFloating({ imageId }); if (failed(f)) return f;
        const w = await api().images.setWrapType({ imageId, type: 'Square' }); if (failed(w)) return w;
      }
      return api().images.setPosition({ imageId, position: { hRelativeFrom: 'column', alignH, vRelativeFrom: 'paragraph', marginOffset: { top: 0 } } });
    });
  }
  async function picSize(dim, v) {
    if (!Number.isFinite(v) || v <= 0) return;
    const pr = picInfo?.properties || {};
    const w0 = pr.size?.width || 1; const h0 = pr.size?.height || 1;
    const size = dim === 'w' ? { width: v, height: Math.round((v * h0) / w0) } : { height: v, width: Math.round((v * w0) / h0) };
    await picAct('Picture size', (imageId) => api().images.setSize({ imageId, size: { ...size, unit: 'px' } }));
  }
  async function picOrder(front) {
    await picAct(front ? 'Bring to front' : 'Send to back', async (imageId) => {
      const all = ((await api().images.list({}))?.items || []).map((i) => i.properties?.relativeHeight || 0);
      const z = front ? Math.max(251658240, ...all) + 1024 : Math.max(0, Math.min(251658240, ...all.filter(Boolean)) - 1024);
      return api().images.setZOrder({ imageId, zOrder: { relativeHeight: z } });
    });
  }
  async function picAlt() {
    const id = picId(); if (!id) return;
    const cur = (await api().images.get({ imageId: id }))?.properties || {};
    const res = await openDialog({ title: 'Alt text', description: 'Read aloud by screen readers in place of the picture.', fields: [{ id: 'alt', label: 'Description', type: 'textarea', rows: 3, value: cur.description || cur.alt || '' }], actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Save', primary: true }] });
    if (res) await picAct('Alt text', (imageId) => api().images.setAltText({ imageId, description: res.values.alt }));
  }

  // ---------------------------------------------------------------- List numbering
  function listTarget() {
    const ed = editor(); const $f = ed?.state?.selection?.$from;
    if (!$f) return null;
    for (let d = $f.depth; d > 0; d--) {
      const n = $f.node(d);
      if (n.isTextblock) {
        const isList = n.attrs?.paragraphProperties?.numberingProperties || n.attrs?.numberingProperties || n.attrs?.listRendering;
        if (!isList) return null;
        return { kind: 'block', nodeType: 'listItem', nodeId: n.attrs.paraId || n.attrs.sdBlockId };
      }
    }
    return null;
  }
  function numberingMenu() {
    const run = (label, fn, ok) => act(label, () => { const t = listTarget(); if (!t) throw new Error('Click in a numbered list first'); return fn(t); }, ok);
    return [
      { label: 'Restart at 1', icon: 'refresh', hint: 'This item starts a new count', action: () => run('Restart numbering', (t) => api().lists.restartAt({ target: t, startAt: 1 }), 'Numbering restarts at 1') },
      { label: 'Continue numbering', icon: 'line-numbers', hint: 'Carry on from the list above', action: () => run('Continue numbering', (t) => api().lists.continuePrevious({ target: t }), 'Numbering continues from the list above') },
      { label: 'Set numbering value...', icon: 'hash', action: async () => {
        const t = listTarget(); if (!t) { toast('Click in a numbered list first'); return; }
        const res = await openDialog({ title: 'Set numbering value', width: 360, fields: [{ id: 'n', label: 'Set value to', type: 'number', value: 1, min: 0, max: 9999 }], actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'OK', primary: true }] });
        if (res) await run('Numbering value', (tt) => api().lists.restartAt({ target: tt, startAt: Math.max(0, Math.round(Number(res.values.n) || 1)) }));
      } },
    ];
  }

  // ---------------------------------------------------------------- Fields, ruler units, moving paragraphs, links
  async function updateFields() {
    let done = 0; let failedCount = 0;
    await act('Update fields', () => {
      const items = (() => { try { return api().fields.list({}).items || []; } catch { return []; } })();
      for (const it of items) {
        const target = it.address || it.domain?.address || it.target;
        if (!target || it.nested) continue;
        try { const r = api().fields.rebuild({ target }); if (failed(r)) failedCount++; else done++; } catch { failedCount++; }
      }
      try { for (const t of api().toc.list({}).items || []) { const target = t.address || t.target; if (target) { api().toc.update({ target }); done++; } } } catch { /* no table of contents */ }
      return { success: true };
    });
    toast(done ? `Updated ${done} field${done === 1 ? '' : 's'}` : failedCount ? 'These fields cannot be updated here' : 'There are no fields to update');
  }
  function rulerUnit() {
    try { return localStorage.getItem('kertas-ruler-unit') || ((navigator.language || 'en-US') === 'en-US' ? 'in' : 'cm'); } catch { return 'cm'; }
  }
  function setRulerUnit(u) {
    try { localStorage.setItem('kertas-ruler-unit', u); } catch { /* private mode */ }
    globalThis.__kertasRulerUnit = u;
    if (rulerOn) { superdoc.toggleRuler(); setTimeout(() => superdoc.toggleRuler(), 30); }
    toast(u === 'cm' ? 'Ruler in centimetres' : 'Ruler in inches');
  }
  /** Moves the paragraph at the cursor (or the table row) up or down, like Word's Alt+Shift+Up/Down */
  function moveBlock(dir) {
    const ed = editor();
    if (tableContext(ed)) return moveRow(ed, dir);
    const $f = ed.state.selection.$from;
    let d = $f.depth; while (d > 1 && $f.node(d - 1).type.name !== 'doc') d--;
    if (d < 1) throw new Error('Nothing to move');
    const index = $f.index(d - 1); const doc = ed.state.doc;
    const to = index + dir;
    if (to < 0 || to >= doc.childCount) throw new Error(dir < 0 ? 'Already at the top' : 'Already at the end');
    const here = doc.child(index); const other = doc.child(to);
    const start = $f.before(d);
    const offset = ed.state.selection.from - start;
    let tr = ed.state.tr;
    if (dir < 0) {
      const otherStart = start - other.nodeSize;
      tr = tr.replaceWith(otherStart, start + here.nodeSize, [here, other]);
      tr.setSelection(ed.state.selection.constructor.near(tr.doc.resolve(otherStart + offset)));
    } else {
      const otherEnd = start + here.nodeSize + other.nodeSize;
      tr = tr.replaceWith(start, otherEnd, [other, here]);
      tr.setSelection(ed.state.selection.constructor.near(tr.doc.resolve(start + other.nodeSize + offset)));
    }
    ed.view.dispatch(tr.scrollIntoView());
    return { success: true };
  }
  /** The link under the pointer in the painted page, if any */
  function linkAt(x, y) {
    const span = document.elementsFromPoint(x, y).find((e) => e.matches?.('[data-pm-start]'));
    if (!span) return null;
    const ed = editor(); const pos = Number(span.dataset.pmStart);
    let href = null;
    try {
      const end = Number(span.dataset.pmEnd || pos + 1);
      ed.state.doc.nodesBetween(pos, Math.max(pos + 1, end), (n) => { const m = n.marks?.find((mk) => mk.type.name === 'link'); if (m && !href) href = m.attrs.href; });
    } catch { /* outside the document */ }
    return href;
  }
  function followLink(href) {
    if (href.startsWith('#')) {
      const name = href.slice(1);
      const ed = editor(); let at = null;
      ed.state.doc.descendants((n, p) => { if (at === null && (n.type.name === 'bookmarkStart' && n.attrs?.name === name)) at = p; });
      if (at !== null) { ed.commands.setTextSelection({ from: at + 1, to: at + 1 }); ed.view.focus(); ed.commands.scrollIntoView?.(); }
      return;
    }
    if (/^(https?:|mailto:)/i.test(href)) window.open(href, '_blank', 'noopener');
  }

  async function decideAll(decision) {
    const total = (await api().trackChanges.list()).total || 0;
    if (!total) { toast('There are no tracked changes'); return; }
    await act(decision === 'accept' ? 'Accept all' : 'Reject all', () => api().trackChanges.decide({ decision, target: { kind: 'all' } }),
      `${decision === 'accept' ? 'Accepted' : 'Rejected'} ${total} change${total === 1 ? '' : 's'}`);
  }
  async function restrictEditing() {
    const v = await restrictDialog();
    if (!v) return;
    if (v === 'off') await act('Stop protection', () => api().protection.clearEditingRestriction({}), 'Protection stopped');
    else await act('Restrict editing', () => api().protection.setEditingRestriction({ mode: v }), 'Editing restriction is on');
  }

  // ---------------------------------------------------------------- View
  const currentZoom = () => superdoc?.getZoom?.() || 100;
  function setZoom(z) {
    const clamped = Math.min(300, Math.max(30, Math.round(z)));
    superdoc.setZoomMode?.('manual');
    superdoc.setZoom?.(clamped);
    syncZoom();
  }
  // The slider puts 100% in the middle: 30-100% on the left half, 100-300% on the right
  const zoomToSlider = (z) => (z <= 100 ? ((z - 30) / 70) * 50 : 50 + ((z - 100) / 200) * 50);
  const sliderToZoom = (v) => (v <= 50 ? 30 + (v / 50) * 70 : 100 + ((v - 50) / 50) * 200);
  function syncZoom() {
    const z = Math.round(currentZoom());
    const foot = el.querySelector('.writer-foot');
    if (!foot) return;
    foot.querySelector('.zoom-pct span').textContent = `${z}%`;
    const range = foot.querySelector('input[type=range]');
    if (document.activeElement !== range) range.value = zoomToSlider(z);
    range.style.setProperty('--fill', `${range.value}%`);
  }
  {
    const foot = el.querySelector('.writer-foot');
    const range = foot.querySelector('input[type=range]');
    range.addEventListener('input', () => {
      // snap to 100% near the middle, as the sheet's slider does
      const v = Math.abs(range.value - 50) < 2 ? 50 : +range.value;
      setZoom(sliderToZoom(v));
    });
    foot.addEventListener('click', (e) => {
      const b = e.target.closest('[data-z]');
      if (!b || !superdoc) return;
      const z = Math.round(currentZoom());
      if (b.dataset.z === 'in') setZoom(z + 10);
      else if (b.dataset.z === 'out') setZoom(z - 10);
      else if (b.dataset.z === 'fit') { superdoc.setZoomMode?.('fit-width'); setTimeout(syncZoom, 200); }
      else if (b.dataset.z === 'menu') openMenu(b, zoomMenu());
    });
  }
  function zoomMenu() {
    const z = Math.round(currentZoom());
    return [
      ...[50, 75, 100, 125, 150, 200].map((v) => ({ label: `${v}%`, checked: v === z, action: () => setZoom(v) })),
      'sep',
      { label: 'Zoom in', icon: 'zoom-in', shortcut: 'Ctrl+Wheel', action: () => setZoom(z + 10) },
      { label: 'Zoom out', icon: 'zoom-out', action: () => setZoom(z - 10) },
      { label: 'Page width', icon: 'fit', action: () => { superdoc.setZoomMode?.('fit-width'); setTimeout(syncZoom, 200); } },
    ];
  }
  /** Editing, Suggesting or Viewing, with the formatting toolbar following along */
  function applyMode(v) {
    superdoc?.setDocumentMode(v);
    const sel = el.querySelector('.ribbon select[data-id="mode"]');
    if (sel && sel.value !== v) sel.value = v;
    // the toolbar forgets the editor while viewing; give it back once the mode has settled
    // SuperDoc's toolbar learns the new mode a little later; once it has, hand the editor back
    // and send a selection transaction so it works out which tools are available again
    const started = Date.now();
    const refresh = () => {
      const tb = superdoc?.toolbar; const ed = superdoc?.activeEditor;
      if (!tb || !ed) return;
      const seen = tb.snapshot?.commands?.['document-mode']?.value;
      if (seen && seen !== v && Date.now() - started < 3000) { setTimeout(refresh, 60); return; }
      if (v !== 'viewing') {
        tb.setActiveEditor?.(ed);
        try { ed.view.focus(); } catch { /* not mounted */ }
        try { ed.view.dispatch(ed.state.tr.setSelection(ed.state.selection).setMeta('addToHistory', false)); } catch { /* nothing to refresh */ }
      }
      tb.updateToolbarState?.();
    };
    setTimeout(refresh, 0);
  }
  function readingMode(on) {
    el.classList.toggle('reading', on);
    applyMode(on ? 'viewing' : (tracking ? 'suggesting' : 'editing'));
    el.querySelector('.exit-reading').hidden = !on;
    if (on) toast('Reading mode. Press Esc to go back to editing.');
  }
  function focusMode(on) {
    el.classList.toggle('focus-mode', on);
    el.querySelector('.exit-reading').hidden = !on;
    el.querySelector('.exit-reading').lastChild.textContent = on ? 'Exit focus mode' : 'Exit reading mode';
  }
  el.querySelector('.exit-reading').addEventListener('click', () => { readingMode(false); focusMode(false); });

  // ---------------------------------------------------------------- side panel
  const sideEl = el.querySelector('.writer-side');
  const sideBody = sideEl.querySelector('.side-body');
  function showSide(which) {
    sidePanel = which;
    sideEl.hidden = false;
    sideEl.querySelectorAll('[data-side]').forEach((b) => b.classList.toggle('on', b.dataset.side === which));
    ribbon.setOn('side', true);
    renderSide();
  }
  function hideSide() { sidePanel = null; sideEl.hidden = true; ribbon.setOn('side', false); }
  sideEl.querySelector('.side-tabs').addEventListener('click', (e) => { const b = e.target.closest('[data-side]'); if (b) showSide(b.dataset.side); });
  sideEl.querySelector('[data-side-close]').addEventListener('click', hideSide);
  sideEl.addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });
  function scheduleSide() { if (!sidePanel) return; clearTimeout(sideTimer); sideTimer = setTimeout(() => (typeof requestIdleCallback === 'function' ? requestIdleCallback(renderSide, { timeout: 1500 }) : renderSide()), 600); }

  async function renderSide() {
    if (!sidePanel || !superdoc) return;
    try {
      if (sidePanel === 'nav') await renderNav();
      if (sidePanel === 'comments') await renderComments();
      if (sidePanel === 'changes') await renderChanges();
    } catch (e) { console.warn('side panel', e); }
  }
  async function renderNav() {
    const outline = (await api().info()).outline || [];
    sideBody.innerHTML = outline.length
      ? `<div class="nav-list">${outline.map((h, i) => `<button class="nav-item" data-i="${i}" style="--lvl:${Math.min(h.level, 6) - 1}">${esc(h.text.replace(/￼/g, '').trim() || '(empty heading)')}</button>`).join('')}</div>`
      : `<div class="empty small">${icon('layers', 22)}<br>No headings yet.<br>Use the Heading styles on the Home tab to build an outline.</div>`;
    sideBody.querySelectorAll('.nav-item').forEach((b) => b.addEventListener('click', () => goToBlock(outline[Number(b.dataset.i)].nodeId)));
  }
  async function goToBlock(nodeId) {
    for (const nodeType of ['heading', 'paragraph']) {
      try { if (await superdoc.navigateTo({ kind: 'block', nodeType, nodeId })) return; } catch { /* try next */ }
    }
    try {
      const pm = editor().pageMetrics;
      const idx = await pm.pageIndexForBodyTarget({ kind: 'block', nodeType: 'heading', nodeId });
      if (typeof idx === 'number') pm.scrollToPage(idx);
    } catch { toast('Could not jump to that heading'); }
  }
  async function renderComments() {
    const items = ((await api().comments.list()).items || []).map((c) => ({ ...c, commentId: c.commentId ?? c.id ?? c.importedId }));
    const ids = new Set(items.map((c) => c.commentId).filter((x) => x != null));
    // a comment whose parent is missing (or is itself) is shown as a thread of its own
    const parentOf = (c) => (c.parentCommentId != null && c.parentCommentId !== c.commentId && ids.has(c.parentCommentId) ? c.parentCommentId : null);
    const roots = items.filter((c) => parentOf(c) === null);
    const replies = (id) => (id == null ? [] : items.filter((c) => parentOf(c) === id));
    const shown = new Set();
    const card = (c, isReply) => {
      if (shown.has(c)) return ''; // never loop on a broken reply chain
      shown.add(c);
      return `
      <div class="cm${isReply ? ' reply' : ''}${c.status === 'resolved' ? ' resolved' : ''}" data-id="${esc(c.commentId)}">
        <div class="cm-head"><span class="avatar">${esc((c.creatorName || '?').slice(0, 1))}</span><b>${esc(c.creatorName || 'Unknown')}</b>${c.createdTime ? `<time>${new Date(c.createdTime).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</time>` : ''}</div>
        <div class="cm-text">${esc(c.text)}</div>
        ${isReply ? '' : `<div class="cm-actions">
          <button class="chip" data-a="reply">${icon('reply', 12)}Reply</button>
          <button class="chip" data-a="${c.status === 'resolved' ? 'reopen' : 'resolve'}">${icon('check', 12)}${c.status === 'resolved' ? 'Reopen' : 'Resolve'}</button>
          <button class="chip" data-a="delete">${icon('trash', 12)}</button>
        </div>`}
        ${replies(c.commentId).map((r) => card(r, true)).join('')}
      </div>`;
    };
    sideBody.innerHTML = `
      <div class="side-actions"><button class="chip primary" data-new>${icon('plus', 13)}New comment</button><span class="muted">${roots.length} thread${roots.length === 1 ? '' : 's'}</span></div>
      ${roots.length ? roots.map((c) => card(c, false)).join('') : `<div class="empty small">${icon('note', 22)}<br>No comments.<br>Select text and choose New comment.</div>`}`;
    sideBody.querySelector('[data-new]').addEventListener('click', newComment);
    sideBody.querySelectorAll('.cm:not(.reply)').forEach((n) => n.addEventListener('click', async (e) => {
      const cid = n.dataset.id;
      const a = e.target.closest('[data-a]')?.dataset.a;
      if (!a) { try { superdoc.scrollToComment(cid, { behavior: 'smooth', block: 'center' }); } catch { /* noop */ } return; }
      if (a === 'reply') { const t = await commentDialog('Reply'); if (t) await act('Reply', () => api().comments.create({ text: t, parentCommentId: cid })); }
      if (a === 'resolve' || a === 'reopen') await act('Comment', () => api().comments.patch({ commentId: cid, status: a === 'resolve' ? 'resolved' : 'active' }));
      if (a === 'delete') await act('Delete comment', () => api().comments.delete({ commentId: cid }));
      renderComments();
    }));
  }
  async function renderChanges() {
    const res = await api().trackChanges.list();
    const items = res.items || [];
    const kind = (t) => (/^ins/.test(t) ? 'insert' : /^del/.test(t) ? 'delete' : /format/.test(t) ? 'format' : /move/.test(t) ? 'move' : 'change');
    const label = (t) => ({ insert: 'Inserted', delete: 'Deleted', format: 'Formatted', move: 'Moved' }[kind(t)] || 'Changed');
    sideBody.innerHTML = `
      <div class="side-actions">
        <button class="chip ${tracking ? 'primary' : ''}" data-track>${icon('track', 13)}Tracking ${tracking ? 'on' : 'off'}</button>
        ${items.length ? `<button class="chip" data-all="accept">${icon('check', 12)}Accept all</button><button class="chip" data-all="reject">${icon('close', 12)}Reject all</button>` : ''}
      </div>
      ${items.length ? items.map((c) => `
        <div class="tc tc-${kind(c.type || '')}" data-id="${esc(c.id)}">
          <div class="cm-head"><span class="tc-kind">${label(c.type)}</span><b>${esc(c.author || c.authorName || '')}</b>${c.date ? `<time>${new Date(c.date).toLocaleString('en-GB', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' })}</time>` : ''}</div>
          <div class="cm-text">${changeText(c)}</div>
          <div class="cm-actions"><button class="chip" data-a="accept">${icon('check', 12)}Accept</button><button class="chip" data-a="reject">${icon('close', 12)}Reject</button></div>
        </div>`).join('') : `<div class="empty small">${icon('track', 22)}<br>No tracked changes.<br>Turn on Track changes to record edits as suggestions.</div>`}`;
    sideBody.querySelector('[data-track]').addEventListener('click', () => { setTracking(!tracking).then(renderChanges); });
    sideBody.querySelectorAll('[data-all]').forEach((b) => b.addEventListener('click', async () => { await decideAll(b.dataset.all); renderChanges(); }));
    sideBody.querySelectorAll('.tc').forEach((n) => n.addEventListener('click', async (e) => {
      const a = e.target.closest('[data-a]')?.dataset.a;
      const tid = n.dataset.id;
      if (!a) { try { await superdoc.navigateTo({ kind: 'entity', entityType: 'trackedChange', entityId: tid }); } catch { /* noop */ } return; }
      await act(a === 'accept' ? 'Accept' : 'Reject', () => api().trackChanges.decide({ decision: a, target: { kind: 'id', id: tid } }));
      renderChanges();
    }));
  }

  // ---------------------------------------------------------------- stats
  // Word count and page count update once typing pauses, when the browser is idle
  function scheduleStats() { clearTimeout(statsTimer); statsTimer = setTimeout(() => (typeof requestIdleCallback === 'function' ? requestIdleCallback(updateStats, { timeout: 1500 }) : updateStats()), 800); }
  async function updateStats() {
    if (!superdoc) return;
    let words = 0;
    try { words = (await api().info()).counts?.words ?? 0; } catch { /* still loading */ }
    const pageCount = pages || el.querySelectorAll('.superdoc-page').length || 1;
    const wf = el.querySelector('.wf-words');
    if (wf) wf.textContent = `${pageCount} page${pageCount === 1 ? '' : 's'} · ${words.toLocaleString()} word${words === 1 ? '' : 's'}`;
    syncZoom();
  }

  // ---------------------------------------------------------------- keyboard
  const onKey = (e) => {
    if (!el.classList.contains('active')) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key.toLowerCase();
    if (mod && !e.shiftKey && k === 'f') { e.preventDefault(); e.stopPropagation(); openFind(false); }
    else if (mod && k === 'h') { e.preventDefault(); e.stopPropagation(); openFind(true); }
    else if (mod && e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); pageBreak(); }
    else if (mod && e.shiftKey && (e.key === '+' || e.key === '=')) { e.preventDefault(); toggleVert('superscript'); }
    else if (mod && !e.shiftKey && e.key === '=') { e.preventDefault(); toggleVert('subscript'); }
    else if (e.key === 'Escape' && (el.classList.contains('reading') || el.classList.contains('focus-mode'))) { readingMode(false); focusMode(false); }
  };
  window.addEventListener('keydown', onKey, true);

  setStatus(data ? name : 'New document');

  return {
    async save(ext) {
      if (ext !== 'docx') throw new Error(`Writer can only save .docx (asked for .${ext})`);
      const blob = await superdoc.export({ exportType: ['docx'], commentsType: 'external', triggerDownload: false });
      let bytes = new Uint8Array(await blob.arrayBuffer());
      try { bytes = await citationsToFields(bytes); } catch (e) { console.warn('citation fields', e); }
      try { bytes = await repairDocx(bytes, openedBytes); } catch (e) { console.warn('repair', e); }
      return bytes;
    },

    /** Prints the paginated layout to PDF using the window's print stylesheet */
    async exportPdf() {
      const zoom = currentZoom();
      if (zoom !== 100) superdoc.setZoom?.(100);
      await new Promise((r) => setTimeout(r, 250));
      const first = el.querySelector('.superdoc-page');
      const w = first ? first.offsetWidth / 96 : 8.5;
      const h = first ? first.offsetHeight / 96 : 11;
      const style = document.createElement('style');
      style.textContent = `@page { size: ${w.toFixed(3)}in ${h.toFixed(3)}in; margin: 0; }`;
      document.head.appendChild(style);
      document.body.classList.add('print-writer');
      try {
        return await printToPdf({ pageSize: { width: w, height: h } });
      } finally {
        document.body.classList.remove('print-writer');
        style.remove();
        if (zoom !== 100) superdoc.setZoom?.(zoom);
      }
    },

    markSaved() { setDirty(false); },
    dispose() {
      clearTimeout(statsTimer);
      clearTimeout(sideTimer);
      window.removeEventListener('keydown', onKey, true);
      slash.dispose();

      ribbon.dispose();
      try { stopBoxes?.(); decorations?.destroy(); unwire?.(); } catch { /* fine */ }
      window.removeEventListener('kertas-user', onUser);
      try { superdoc.destroy(); } catch { /* already gone */ }
    },
    superdoc,
  };
}
