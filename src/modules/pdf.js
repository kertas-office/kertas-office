// PDF: viewing with pdf.js (Apache-2.0), editing and saving with pdf-lib (MIT, @cantoo fork)
import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { PDFDocument } from '@cantoo/pdf-lib';
import { buildPdf, listFormFields, imagesToPdf } from '../lib/pdf-ops.js';
import { openFiles, saveMany, saveFile, saveAnyFile } from '../lib/platform.js';
import { extractPages as extractTextPages, pdfToDocx, pdfToSnapshot, fontInfo } from '../lib/pdf-convert.js';
import { findFont, cssStack } from '../lib/font-catalog.js';
import { snapshotToXlsx } from '../lib/xlsx-bridge.js';
import { OCR_LANGUAGES, recognise, stopOcr } from '../lib/ocr.js';
import { openDialog } from '../lib/dialog.js';
import { readCertificate, signPdf } from '../lib/pdf-sign.js';
import { createRibbon } from '../lib/ribbon.js';
import { icon } from '../lib/icons.js';
import { createObjects, isMovable } from './pdf-objects.js';
import {
  STAMPS, askPassword, protectDialog, watermarkDialog, headerFooterDialog, splitDialog,
  noteDialog, bookmarkDialog, imagesDialog,
} from './pdf-dialogs.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

const TOOLS = {
  select: 'Select text, click links and comments',
  text: 'Click where you want to type',
  'edit-text': 'Click a line of the PDF to change its words',
  date: "Click to insert today's date",
  check: 'Click to add a tick',
  cross: 'Click to add a cross',
  highlight: 'Drag over an area to highlight it',
  'mark-hl': 'Select text to highlight it',
  'mark-ul': 'Select text to underline it',
  'mark-st': 'Select text to strike it through',
  rect: 'Drag to draw a box',
  ellipse: 'Drag to draw an ellipse',
  line: 'Drag to draw a line',
  arrow: 'Drag to draw an arrow',
  ink: 'Draw freehand',
  note: 'Click to add a comment',
  stamp: 'Click to place the stamp',
  sign: 'Click to place your signature',
  image: 'Click to place the image',
  redact: 'Drag over content to redact it. It is removed for good when you save.',
  erase: 'Click an annotation to remove it',
  'field-text': 'Drag to draw a text field, or click to place one',
  'field-para': 'Drag to draw a box for several lines of text',
  'field-check': 'Click to place a check box',
  'field-radio': 'Click to place a choice button. Buttons with the same group name work together.',
  'field-drop': 'Drag to draw a drop-down list',
  'field-list': 'Drag to draw a list box',
};
// New form fields: the tool, the kind saved in the PDF and the size when placed with a click
const FIELD_TOOLS = {
  'field-text': { kind: 'text', label: 'Text field', w: 180, h: 22 },
  'field-para': { kind: 'multiline', label: 'Paragraph field', w: 240, h: 72 },
  'field-check': { kind: 'checkbox', label: 'Check box', w: 14, h: 14 },
  'field-radio': { kind: 'radio', label: 'Choice button', w: 14, h: 14 },
  'field-drop': { kind: 'dropdown', label: 'Drop-down list', w: 160, h: 22 },
  'field-list': { kind: 'list', label: 'List box', w: 160, h: 64 },
};
const FIELD_LABEL = Object.fromEntries(Object.values(FIELD_TOOLS).map((f) => [f.kind, f.label]));
const TEXT_SELECT_TOOLS = new Set(['select', 'mark-hl', 'mark-ul', 'mark-st']);
const DRAG_TOOLS = new Set(['highlight', 'rect', 'ellipse', 'line', 'arrow', 'redact']);
const DEFAULT_COLORS = {
  highlight: '#facc15', 'mark-hl': '#facc15', 'mark-ul': '#2563eb', 'mark-st': '#e11d48', note: '#facc15',
  text: '#111827', date: '#111827', check: '#111827', cross: '#111827', redact: '#000000',
};
const SVGNS = 'http://www.w3.org/2000/svg';
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const today = () => new Date().toLocaleDateString('en-GB');

export async function mount(el, { data, name, setDirty, setStatus, toast, fileMenu, openDocument }) {
  if (!data) throw new Error('Open an existing PDF to edit it');

  el.innerHTML = `
    <div data-ribbon></div>
    <div class="pdf-layout">
      <aside class="pdf-side">
        <div class="side-tabs">
          <button class="on" data-side="pages" title="Pages">${icon('grid', 14)}<span>Pages</span></button>
          <button data-side="outline" title="Bookmarks">${icon('bookmark', 14)}<span>Bookmarks</span></button>
          <button data-side="layers" title="Layers">${icon('layers', 14)}<span>Layers</span></button>
          <button data-side="attach" title="Files attached to this PDF">${icon('attach', 14)}<span>Files</span></button>
        </div>
        <div class="pdf-thumbs"></div>
        <div class="pdf-outline" hidden></div>
        <div class="pdf-attach" hidden></div>
      </aside>
      <div class="pdf-pages" tabindex="0"></div>
    </div>`;

  // ---------------------------------------------------------------- ribbon
  const SIZES = [8, 10, 12, 14, 18, 24, 32, 48].map((v) => ({ value: v, label: `${v} pt` }));
  const toolBtn = (id, ic, label) => ({ id: `tool-${id}`, kind: 'big', icon: ic, label, title: TOOLS[id], action: () => setTool(id) });
  const toolSmall = (id, ic, label) => ({ id: `tool-${id}`, icon: ic, label, title: TOOLS[id], action: () => setTool(id) });
  const styleGroup = () => ({ label: 'Style', items: [
    { id: `color-${Math.random().toString(36).slice(2, 7)}`, kind: 'color', value: '#e11d48', label: 'Colour', title: 'Colour', action: (v) => setColor(v) },
    { id: `size-${Math.random().toString(36).slice(2, 7)}`, kind: 'select', value: 14, options: SIZES, label: 'Size', title: 'Text size and line width', action: (v) => setSize(v) },
  ] });
  const stampSelect = { id: 'stamp-kind', kind: 'select', value: 'approved', label: 'Stamp type', title: 'Stamp', options: STAMPS.map((s) => ({ value: s.id, label: s.label })), action: (v) => { stampKind = v; setTool('stamp'); } };

  const ribbon = createRibbon(el.querySelector('[data-ribbon]'), {
    app: 'pdf',
    fileMenu: [
      ...fileMenu.slice(0, 6),
      { id: 'extract', label: 'Extract pages...', icon: 'split', action: () => extractPages('extract') },
      { id: 'protect', label: 'Protect with password...', icon: 'lock', action: () => protect() },
      { id: 'compress', label: 'Save a smaller copy...', icon: 'compress', action: () => compressCopy() },
      ...fileMenu.slice(6),
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Tools', items: [toolBtn('select', 'select', 'Select'), toolBtn('text', 'text', 'Add text'), toolBtn('mark-hl', 'highlight', 'Highlight'), toolBtn('note', 'note', 'Comment'), toolBtn('sign', 'sign', 'Sign')] },
        { label: 'Pages', items: [
          { id: 'h-merge', kind: 'big', icon: 'merge', label: 'Merge', title: 'Append pages from other PDFs', action: () => mergeFiles() },
          { id: 'h-split', kind: 'big', icon: 'split', label: 'Split', title: 'Split into several files', action: () => extractPages('split') },
        ] },
        { label: 'Zoom', items: [
          { id: 'h-zoom-out', kind: 'big', icon: 'zoom-out', label: 'Zoom out', action: () => zoomBy(1 / 1.2) },
          { id: 'zoom', kind: 'label', label: '100%', caption: 'Zoom' },
          { id: 'h-zoom-in', kind: 'big', icon: 'zoom-in', label: 'Zoom in', action: () => zoomBy(1.2) },
          { id: 'h-fit', kind: 'big', icon: 'fit', label: 'Fit width', action: () => fitAndLayout() },
        ] },
      ] },
      { id: 'edit', label: 'Edit', groups: [
        { label: 'Text', items: [toolBtn('edit-text', 'text-cursor', 'Edit text')] },
        { label: 'Insert', items: [
          toolBtn('image', 'image', 'Image'),
          { id: 'e-blank', kind: 'big', icon: 'file-plus', label: 'Blank page', title: 'Insert a blank page after this one', action: () => insertBlank() },
          { id: 'e-pdf', kind: 'big', icon: 'merge', label: 'Insert PDF', action: () => mergeFiles() },
          { id: 'e-attach', kind: 'big', icon: 'attach', label: 'Attach file', title: 'Attach any file to this PDF (it travels inside the PDF)', action: () => attachFiles() },
        ] },
        { label: 'Page design', items: [
          { id: 'e-watermark', kind: 'big', icon: 'droplet', label: 'Watermark', action: () => editWatermark() },
          { id: 'e-hf', kind: 'big', icon: 'header', label: 'Header & footer', action: () => editHeaderFooter() },
          { id: 'e-num', kind: 'big', icon: 'hash', label: 'Page numbers', action: () => editHeaderFooter(true) },
        ] },
        { label: 'Bookmarks', items: [{ id: 'e-bm', kind: 'big', icon: 'bookmark', label: 'Add bookmark', action: () => addBookmark() }] },
        { label: 'Layers', items: [
          { id: 'e-layers', kind: 'big', icon: 'layers', label: 'Layers', title: 'Show the layers of this PDF', action: () => showLayers() },
          { id: 'e-newlayer', kind: 'big', icon: 'layer-plus', label: 'New layer', action: () => objects.newLayer() },
        ] },
      ] },
      { id: 'organize', label: 'Organize', groups: [
        { label: 'Current page', items: [
          { id: 'p-rotl', kind: 'big', icon: 'rotate-l', label: 'Rotate left', action: () => pageOp('rotl', currentPage) },
          { id: 'p-rotr', kind: 'big', icon: 'rotate-r', label: 'Rotate right', action: () => pageOp('rotr', currentPage) },
          { id: 'p-up', kind: 'big', icon: 'up', label: 'Move up', action: () => pageOp('up', currentPage) },
          { id: 'p-down', kind: 'big', icon: 'down', label: 'Move down', action: () => pageOp('down', currentPage) },
          { id: 'p-dup', kind: 'big', icon: 'files', label: 'Duplicate', action: () => pageOp('dup', currentPage) },
          { id: 'p-del', kind: 'big', icon: 'trash', label: 'Delete', action: () => pageOp('del', currentPage) },
        ] },
        { label: 'Document', items: [
          { id: 'o-blank', kind: 'big', icon: 'file-plus', label: 'Blank page', action: () => insertBlank() },
          { id: 'o-merge', kind: 'big', icon: 'merge', label: 'Insert PDF', action: () => mergeFiles() },
          { id: 'o-extract', kind: 'big', icon: 'export', label: 'Extract', action: () => extractPages('extract') },
          { id: 'o-split', kind: 'big', icon: 'split', label: 'Split', action: () => extractPages('split') },
        ] },
        { label: 'Optimize', items: [
          { id: 'o-compress', kind: 'big', icon: 'compress', label: 'Compress', title: 'Save a smaller copy of this PDF', action: () => compressCopy() },
        ] },
      ] },
      { id: 'comment', label: 'Comment', groups: [
        { label: 'Text markup', items: [toolBtn('mark-hl', 'highlight', 'Highlight'), toolBtn('mark-ul', 'underline', 'Underline'), toolBtn('mark-st', 'strike', 'Strikeout')] },
        { label: 'Notes', items: [toolBtn('note', 'note', 'Comment'), toolBtn('text', 'text', 'Text box'), toolBtn('stamp', 'stamp', 'Stamp'), stampSelect] },
        { label: 'Drawing', items: [toolSmall('highlight', 'highlight', 'Area'), toolSmall('rect', 'rect', 'Box'), toolSmall('ellipse', 'ellipse', 'Ellipse'), toolSmall('line', 'line', 'Line'), toolSmall('arrow', 'arrow', 'Arrow'), toolSmall('ink', 'ink', 'Pen')] },
        styleGroup(),
        { label: 'Edit', items: [toolBtn('edit-text', 'text-cursor', 'Edit text'), toolBtn('erase', 'erase', 'Erase'), { id: 'undo', kind: 'big', icon: 'undo', label: 'Undo', title: 'Undo (Ctrl+Z)', action: () => undo() }] },
      ] },
      { id: 'fill-sign', label: 'Fill & Sign', groups: [
        { label: 'Fill in', items: [toolBtn('text', 'text', 'Text'), toolBtn('date', 'calendar', 'Date'), toolBtn('check', 'check', 'Tick'), toolBtn('cross', 'close', 'Cross')] },
        { label: 'Sign', items: [toolBtn('sign', 'sign', 'Sign'), { id: 'resign', kind: 'big', icon: 'edit', label: 'New signature', title: 'Draw a different signature', action: () => openSignaturePad() }] },
        { label: 'Create form', items: [
          toolSmall('field-text', 'field-text', 'Text field'), toolSmall('field-para', 'field-para', 'Paragraph'), toolSmall('field-check', 'field-check', 'Check box'),
          toolSmall('field-radio', 'field-radio', 'Choice'), toolSmall('field-drop', 'field-drop', 'Drop-down'), toolSmall('field-list', 'field-list', 'List box'),
        ] },
        { label: 'Form', items: [
          { id: 'form-hl', kind: 'big', icon: 'form', label: 'Highlight fields', title: 'Show or hide form field shading', action: () => toggleFieldShading() },
          { id: 'form-clear', kind: 'big', icon: 'eraser-form', label: 'Clear form', title: 'Reset every field to its original value', action: () => clearForm() },
          { id: 'form-flatten', kind: 'big', icon: 'flatten', label: 'Flatten', title: 'Lock filled values into the page when saving', action: () => { flattenForms = !flattenForms; ribbon.setOn('form-flatten', flattenForms); changed(); } },
        ] },
        styleGroup(),
      ] },
      { id: 'protect', label: 'Protect', groups: [
        { label: 'Security', items: [
          { id: 'x-protect', kind: 'big', icon: 'lock', label: 'Password', title: 'Encrypt with a password and set permissions', action: () => protect() },
          { id: 'x-unlock', kind: 'big', icon: 'unlock', label: 'Remove password', title: 'Save without encryption', action: () => removeProtection() },
        ] },
        { label: 'Certificate', items: [
          { id: 'x-digisign', kind: 'big', icon: 'file-sign', label: 'Digitally sign', title: 'Sign with a certificate (.p12 or .pfx) and save a signed copy', action: () => digitalSign() },
        ] },
        { label: 'Redaction', items: [toolBtn('redact', 'redact', 'Redact')] },
        { label: 'Marking', items: [{ id: 'x-wm', kind: 'big', icon: 'droplet', label: 'Watermark', action: () => editWatermark() }] },
      ] },
      { id: 'convert', label: 'Convert', groups: [
        { label: 'Export to', items: [
          { id: 'c-word', kind: 'big', icon: 'doc', label: 'Document', title: 'Turn this PDF into an editable Word document', action: () => convertTo('word') },
          { id: 'c-excel', kind: 'big', icon: 'sheet', label: 'Workbook', title: 'Put the tables and text of this PDF into a workbook', action: () => convertTo('excel') },
          { id: 'c-images', kind: 'menu', icon: 'image', label: 'Images', items: [
            { label: 'PNG, one per page (150 dpi)', icon: 'image', action: () => pagesToImages('png', 150) },
            { label: 'PNG, high quality (300 dpi)', icon: 'image', action: () => pagesToImages('png', 300) },
            { label: 'JPG, smaller files (150 dpi)', icon: 'image', action: () => pagesToImages('jpg', 150) },
          ] },
        ] },
        { label: 'Scanned documents', items: [
          { id: 'c-ocr', kind: 'big', icon: 'scan-text', label: 'Recognise text', title: 'Read the text in scanned pages (OCR) so it can be searched, copied and converted', action: () => runOcr() },
        ] },
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Zoom', items: [
          { id: 'v-zoom-out', kind: 'big', icon: 'zoom-out', label: 'Zoom out', action: () => zoomBy(1 / 1.2) },
          { id: 'v-zoom-in', kind: 'big', icon: 'zoom-in', label: 'Zoom in', action: () => zoomBy(1.2) },
          { id: 'v-fit', kind: 'big', icon: 'fit', label: 'Fit width', action: () => fitAndLayout() },
          { id: 'v-page', kind: 'big', icon: 'fit-page', label: 'Whole page', action: () => fitPage() },
        ] },
        { label: 'Panels', items: [{ id: 'v-thumbs', kind: 'big', icon: 'sidebar', label: 'Sidebar', action: () => toggleSidebar() }, { id: 'v-layers', kind: 'big', icon: 'layers', label: 'Layers', action: () => showLayers() }] },
      ] },
      { id: 'object', label: 'Object', contextual: true, groups: [
        { label: 'Arrange', items: [
          { id: 'obj-front', kind: 'big', icon: 'front', label: 'Bring to front', title: 'Bring to front (Ctrl+Shift+])', action: () => objects.order('front') },
          { id: 'obj-fwd', kind: 'big', icon: 'up', label: 'Forward', title: 'Bring forward (Ctrl+])', action: () => objects.order('forward') },
          { id: 'obj-bwd', kind: 'big', icon: 'down', label: 'Backward', title: 'Send backward (Ctrl+[)', action: () => objects.order('backward') },
          { id: 'obj-back', kind: 'big', icon: 'back', label: 'Send to back', title: 'Send to back (Ctrl+Shift+[)', action: () => objects.order('back') },
        ] },
        { label: 'Size', items: [
          { id: 'obj-w', kind: 'number', label: 'Width', unit: 'pt', step: 1, min: 1, title: 'Width in points', action: (v) => objects.setSize('w', v) },
          { id: 'obj-h', kind: 'number', label: 'Height', unit: 'pt', step: 1, min: 1, title: 'Height in points', action: (v) => objects.setSize('h', v) },
        ] },
        { label: 'Form field', items: [
          { id: 'obj-field', kind: 'big', icon: 'field-props', label: 'Field options', title: 'Name, choices and other options of this field (double-click the field)', action: () => { const a = objects.selected?.a; if (a?.type === 'field') editField(a); } },
        ] },
        { label: 'Layer', items: [
          { id: 'obj-layer', kind: 'menu', icon: 'layers', label: 'Move to layer', items: () => objects.layerChoices() },
        ] },
        { label: 'Edit', items: [
          { id: 'obj-dup', kind: 'big', icon: 'duplicate', label: 'Duplicate', title: 'Duplicate (Ctrl+D)', action: () => objects.duplicate() },
          { id: 'obj-del', kind: 'big', icon: 'trash', label: 'Delete', title: 'Delete (Del)', action: () => objects.remove() },
        ] },
      ] },
    ],
    inline: [
      { id: 'secure', kind: 'label', label: '' },
      { id: 'find', kind: 'search', label: 'Find in document' },
      { id: 'pageinfo', kind: 'label', label: '' },
    ],
  });

  const pagesEl = el.querySelector('.pdf-pages');
  const sideEl = el.querySelector('.pdf-side');
  const thumbsEl = el.querySelector('.pdf-thumbs');
  const outlineEl = el.querySelector('.pdf-outline');
  const attachEl = el.querySelector('.pdf-attach');

  // ---------------------------------------------------------------- state
  const sources = {};      // id -> { bytes, doc (pdfjs), password }
  const pages = [];        // page model, see pdf-ops.js
  let scale = 1;
  let tool = 'select';
  let userColor = null;    // null until the user picks a colour
  let userSize = 14;
  let stampKind = 'approved';
  let signaturePng = null; // { bytes, dataUrl, aspect }
  let placeImage = null;   // { bytes, fmt, dataUrl, aspect }
  const undoStack = [];
  let srcSeq = 0;
  let disposed = false;
  let findQuery = '';
  let currentPage = 0;
  // A page picked in the thumbnails (or just moved, rotated or deleted) stays current until
  // the person scrolls the pages themselves: a short last page may never reach the line
  // that marks the current page while scrolling
  let pickedPage = null;
  let keySeq = 0;
  // document-level settings
  let security = null;     // { userPassword, ownerPassword, permissions, restrict }
  let watermark = null;
  let headerFooter = null;
  let flattenForms = false;
  let fieldShading = true;
  let outline = [];        // [{ title, src, index, children }]
  let outlineTouched = false;
  let attachments = [];    // files attached to the PDF: [{ name, data, description }]
  const forms = {};        // srcId -> fields (from listFormFields)
  const formValues = {};   // srcId -> { name: value }
  const formInitial = {};

  const changed = () => setDirty(true);
  const colorFor = (t) => userColor || DEFAULT_COLORS[t] || '#e11d48';
  const size = () => userSize;
  function setColor(v) {
    userColor = v;
    el.querySelectorAll('.ribbon input[type=color]').forEach((n) => { n.value = v; });
  }
  function setSize(v) {
    userSize = Number(v);
    el.querySelectorAll('.ribbon select[data-id^="size-"]').forEach((n) => { n.value = v; });
  }
  const newKey = (src, index) => `${src}:${index}:${keySeq++}`;

  // ---------------------------------------------------------------- sources
  async function addSource(bytes, fileName, isMain = false) {
    let password;
    let wrong = false;
    for (;;) {
      try {
        const doc = await pdfjsLib.getDocument({ data: bytes.slice(), password, isEvalSupported: false, fontExtraProperties: true }).promise;
        const id = `s${srcSeq++}`;
        sources[id] = { bytes, doc, password };
        if (password && isMain) security = { userPassword: password, ownerPassword: password, restrict: false, inherited: true };
        return { id, doc };
      } catch (err) {
        if (err?.name !== 'PasswordException') throw err;
        password = await askPassword(fileName, wrong || err.code === 2);
        if (password === null) throw new Error('A password is needed to open this PDF');
        wrong = true;
      }
    }
  }

  async function blankSource(width, height) {
    const d = await PDFDocument.create();
    d.addPage([width, height]);
    return addSource(await d.save(), 'blank');
  }

  const main = await addSource(data, name, true);
  for (let i = 0; i < main.doc.numPages; i++) pages.push({ src: main.id, index: i, rotate: 0, annots: [], key: newKey(main.id, i) });

  // Form fields (read with pdf-lib so saving and filling use the same field names)
  try {
    const fields = await listFormFields(data, sources[main.id].password);
    if (fields.length) {
      forms[main.id] = fields;
      formValues[main.id] = Object.fromEntries(fields.map((f) => [f.name, f.value]));
      formInitial[main.id] = structuredClone(formValues[main.id]);
    }
  } catch (err) { console.warn('form read failed', err); }

  // Bookmarks
  async function loadOutline() {
    const doc = sources[main.id].doc;
    const raw = await doc.getOutline().catch(() => null);
    if (!raw?.length) return [];
    const resolve = async (dest) => {
      try {
        const d = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
        if (!Array.isArray(d)) return null;
        return typeof d[0] === 'number' ? d[0] : await doc.getPageIndex(d[0]);
      } catch { return null; }
    };
    const walk = async (items) => Promise.all(items.map(async (it) => ({
      title: it.title, src: main.id, index: it.dest ? await resolve(it.dest) : null, url: it.url || null, children: await walk(it.items || []),
    })));
    return walk(raw);
  }
  outline = await loadOutline();

  // ---------------------------------------------------------------- rendering
  const rendered = new Map(); // key -> { el, viewport, renderTask }
  const objects = createObjects({
    el, pagesEl, sideEl, pages, rendered, sources, ribbon, toast, changed, undoStack, setStatus,
    drawAnnots: (p) => drawAnnots(p),
    refreshPage: (p) => refreshPage(p),
    layoutAll: () => layout(),
    currentPageModel: () => pages[currentPage] || pages[0],
  });
  await objects.loadSource(main.id);
  /** Draws one page again (after a layer is shown, hidden or lifted) */
  async function refreshPage(p) {
    const old = rendered.get(p.key);
    old?.renderTask?.cancel();
    rendered.delete(p.key);
    const div = pagesEl.querySelector(`.pdf-page[data-key="${p.key}"]`);
    if (div) div.innerHTML = '';
    await renderPage(p.key);
  }
  function showLayers() {
    if (sideEl.classList.contains('collapsed')) toggleSidebar();
    sideEl.querySelector('[data-side="layers"]').click();
  }

  async function pdfPageFor(p) { return sources[p.src].doc.getPage(p.index + 1); }
  function viewportFor(pdfPage, p, s = scale) {
    return pdfPage.getViewport({ scale: s, rotation: (pdfPage.rotate + p.rotate) % 360 });
  }

  async function fitWidth() {
    if (!pages.length) return;
    const first = await pdfPageFor(pages[0]);
    const vp = viewportFor(first, pages[0], 1);
    scale = Math.max(0.25, Math.min(1.5, (pagesEl.clientWidth - 48) / vp.width));
  }
  async function fitPage() {
    const p = pages[currentPage] || pages[0];
    const vp = viewportFor(await pdfPageFor(p), p, 1);
    scale = Math.max(0.25, Math.min((pagesEl.clientWidth - 48) / vp.width, (pagesEl.clientHeight - 40) / vp.height));
    await layout();
    pagesEl.querySelectorAll('.pdf-page')[currentPage]?.scrollIntoView({ block: 'center' });
  }

  const observer = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) renderPage(e.target.dataset.key);
  }, { root: pagesEl, rootMargin: '600px 0px' });

  async function layout() {
    observer.disconnect();
    for (const r of rendered.values()) r.renderTask?.cancel();
    rendered.clear();
    pagesEl.innerHTML = '';
    ribbon.setLabel('zoom', `${Math.round(scale * 100)}%`);
    for (const p of pages) {
      const vp = viewportFor(await pdfPageFor(p), p);
      const div = document.createElement('div');
      div.className = 'pdf-page';
      div.dataset.key = p.key;
      div.style.width = `${vp.width}px`;
      div.style.height = `${vp.height}px`;
      div.style.setProperty('--scale-factor', vp.scale);
      pagesEl.appendChild(div);
      observer.observe(div);
    }
    applyToolClass();
    updatePageInfo();
  }

  const fieldsOn = (p) => (forms[p.src] || []).filter((f) => f.widgets.some((w) => w.page === p.index));

  async function renderPage(key) {
    if (rendered.has(key) || disposed) return;
    const p = pages.find((x) => x.key === key);
    const div = pagesEl.querySelector(`.pdf-page[data-key="${key}"]`);
    if (!p || !div) return;
    const entry = { el: div, viewport: null, renderTask: null };
    rendered.set(key, entry);

    const pdfPage = await pdfPageFor(p);
    const vp = viewportFor(pdfPage, p);
    entry.viewport = vp;
    const hasFields = fieldsOn(p).length > 0;

    const ratio = window.devicePixelRatio || 1;
    const canvas = document.createElement('canvas');
    canvas.className = 'render';
    canvas.width = Math.floor(vp.width * ratio);
    canvas.height = Math.floor(vp.height * ratio);
    canvas.style.width = `${vp.width}px`;
    canvas.style.height = `${vp.height}px`;
    div.appendChild(canvas);

    // Layers: hidden ones stay off, lifted ones (being moved) are drawn on their own afterwards
    const lifted = objects.liftedKeys(p);
    const ocConfig = await objects.configFor(p, lifted);
    const annotationMode = hasFields ? pdfjsLib.AnnotationMode.ENABLE_FORMS : pdfjsLib.AnnotationMode.ENABLE;
    entry.renderTask = pdfPage.render({
      canvas, viewport: vp, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : null,
      // Form widgets are drawn by our own inputs, so leave their appearance off the canvas
      annotationMode,
      ...(ocConfig ? { optionalContentConfigPromise: Promise.resolve(ocConfig) } : {}),
    });
    try { await entry.renderTask.promise; } catch (e) { if (e?.name !== 'RenderingCancelledException') console.warn(e); return; }
    if (disposed || !div.isConnected || rendered.get(key) !== entry) return;
    if (lifted.size) await objects.renderOverlays(p, pdfPage, vp, div, canvas, entry, ratio, annotationMode);

    const textDiv = document.createElement('div');
    textDiv.className = 'textLayer';
    div.appendChild(textDiv);
    const textLayer = new pdfjsLib.TextLayer({ textContentSource: pdfPage.streamTextContent(), container: textDiv, viewport: vp });
    await textLayer.render().catch(() => {});
    // Remember each span's real font (by family name) so Edit text can preview in it
    pdfPage.getTextContent().then((tc) => {
      const items = tc.items.filter((it) => it.str !== undefined);
      textLayer.textDivs.forEach((sp, i) => {
        const id = items[i]?.fontName;
        if (!id) return;
        try { const f = pdfPage.commonObjs.get(id); const fam = fontInfo(f?.name || '').family; if (fam) sp.dataset.family = fam; } catch { /* font not loaded */ }
      });
    }).catch(() => {});

    renderLinksAndComments(div, pdfPage, vp);
    if (hasFields) renderFormLayer(div, p, vp);

    const svg = document.createElementNS(SVGNS, 'svg');
    svg.classList.add('annots');
    div.appendChild(svg);
    bindPointer(div, svg, p, entry);
    drawAnnots(p);
    if (findQuery) highlightMatches(p, entry);
  }

  // Links and comments that already exist in the file
  async function renderLinksAndComments(div, pdfPage, vp) {
    const annots = await pdfPage.getAnnotations({ intent: 'display' }).catch(() => []);
    if (!annots.length || !div.isConnected) return;
    const layer = document.createElement('div');
    layer.className = 'link-layer';
    for (const a of annots) {
      const [x1, y1] = vp.convertToViewportPoint(a.rect[0], a.rect[1]);
      const [x2, y2] = vp.convertToViewportPoint(a.rect[2], a.rect[3]);
      const box = `left:${Math.min(x1, x2)}px;top:${Math.min(y1, y2)}px;width:${Math.abs(x2 - x1)}px;height:${Math.abs(y2 - y1)}px`;
      if (a.subtype === 'Link' && (a.url || a.dest)) {
        const link = document.createElement('a');
        link.className = 'pdf-link';
        link.style.cssText = box;
        if (a.url) { link.href = a.url; link.target = '_blank'; link.rel = 'noopener'; link.title = a.url; } else {
          link.href = '#';
          link.title = 'Go to page';
          link.addEventListener('click', async (e) => {
            e.preventDefault();
            const src = pages.find((pp) => pp.key === div.dataset.key)?.src;
            const pdf = sources[src]?.doc;
            if (!pdf) return;
            try {
              const d = typeof a.dest === 'string' ? await pdf.getDestination(a.dest) : a.dest;
              const idx = typeof d[0] === 'number' ? d[0] : await pdf.getPageIndex(d[0]);
              goToSourcePage(src, idx);
            } catch { /* broken link */ }
          });
        }
        layer.appendChild(link);
      } else if ((a.subtype === 'Text' || a.subtype === 'FreeText' || a.subtype === 'Highlight' || a.subtype === 'Underline' || a.subtype === 'StrikeOut' || a.subtype === 'Square' || a.subtype === 'Circle' || a.subtype === 'Ink') && a.contentsObj?.str) {
        const tip = document.createElement('div');
        tip.className = 'file-comment';
        tip.style.cssText = box;
        tip.title = `${a.titleObj?.str ? `${a.titleObj.str}: ` : ''}${a.contentsObj.str}`;
        layer.appendChild(tip);
      }
    }
    div.appendChild(layer);
  }

  function goToSourcePage(src, index) {
    const i = pages.findIndex((p) => p.src === src && p.index === index);
    if (i >= 0) pagesEl.querySelectorAll('.pdf-page')[i]?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ---------------------------------------------------------------- forms
  function renderFormLayer(div, p, vp) {
    const layer = document.createElement('div');
    layer.className = `form-layer${fieldShading ? ' shaded' : ''}`;
    const values = formValues[p.src];
    for (const f of fieldsOn(p)) {
      f.widgets.forEach((w) => {
        if (w.page !== p.index) return;
        const [x1, y1] = vp.convertToViewportPoint(w.rect[0], w.rect[1]);
        const [x2, y2] = vp.convertToViewportPoint(w.rect[2], w.rect[3]);
        const left = Math.min(x1, x2); const top = Math.min(y1, y2);
        const width = Math.abs(x2 - x1); const height = Math.abs(y2 - y1);
        let input;
        if (f.type === 'text') {
          input = document.createElement(f.multiline ? 'textarea' : 'input');
          input.value = values[f.name] ?? '';
          if (f.maxLength) input.maxLength = f.maxLength;
          input.style.fontSize = `${Math.max(8, Math.min(height * (f.multiline ? 0.3 : 0.62), 14 * vp.scale))}px`;
          input.addEventListener('input', () => setField(p.src, f.name, input.value));
        } else if (f.type === 'checkbox') {
          input = document.createElement('input');
          input.type = 'checkbox';
          input.checked = !!values[f.name];
          input.addEventListener('change', () => setField(p.src, f.name, input.checked));
        } else if (f.type === 'radio') {
          input = document.createElement('input');
          input.type = 'radio';
          input.name = `${div.dataset.key}-${f.name}`;
          input.checked = values[f.name] === w.onValue;
          input.addEventListener('change', () => { if (input.checked) setField(p.src, f.name, w.onValue); });
        } else {
          input = document.createElement('select');
          if (f.type === 'list') input.multiple = true;
          input.innerHTML = `${f.type === 'dropdown' ? '<option value=""></option>' : ''}${(f.options || []).map((o) => `<option ${(values[f.name] || []).includes(o) ? 'selected' : ''}>${esc(o)}</option>`).join('')}`;
          input.style.fontSize = `${Math.max(8, Math.min(height * 0.6, 14 * vp.scale))}px`;
          input.addEventListener('change', () => setField(p.src, f.name, [...input.selectedOptions].map((o) => o.value).filter(Boolean)));
        }
        input.className = `ff ff-${f.type}`;
        input.dataset.field = f.name;
        input.title = f.name;
        input.disabled = f.readOnly;
        input.style.cssText += `;left:${left}px;top:${top}px;width:${width}px;height:${height}px`;
        input.addEventListener('keydown', (e) => e.stopPropagation());
        layer.appendChild(input);
      });
    }
    div.appendChild(layer);
  }

  function setField(src, fieldName, value) {
    formValues[src][fieldName] = value;
    // Keep every widget of the same field in sync (fields can repeat across pages)
    pagesEl.querySelectorAll(`.ff[data-field="${CSS.escape(fieldName)}"]`).forEach((n) => {
      if (n.type === 'checkbox') n.checked = !!value;
      else if (n.type === 'radio') { /* handled by name groups */ } else if (n.tagName === 'SELECT') [...n.options].forEach((o) => { o.selected = (Array.isArray(value) ? value : [value]).includes(o.value); });
      else if (n.value !== value) n.value = value;
    });
    changed();
  }

  function toggleFieldShading() {
    fieldShading = !fieldShading;
    ribbon.setOn('form-hl', fieldShading);
    pagesEl.querySelectorAll('.form-layer').forEach((l) => l.classList.toggle('shaded', fieldShading));
  }

  async function clearForm() {
    if (!forms[main.id]) { toast('This PDF has no form fields'); return; }
    formValues[main.id] = structuredClone(formInitial[main.id]);
    changed();
    await layout();
  }

  // ---------------------------------------------------------------- annotation drawing
  function svgEl(tag, attrs) {
    const n = document.createElementNS(SVGNS, tag);
    for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, v);
    return n;
  }

  function drawAnnots(p) {
    const entry = rendered.get(p.key);
    if (!entry?.viewport) return;
    const svg = entry.el.querySelector('svg.annots');
    const vp = entry.viewport;
    const s = vp.scale;
    svg.innerHTML = '';
    const toView = (x, y) => vp.convertToViewportPoint(x, y);
    const viewRot = vp.rotation;
    const box = (a) => {
      const [ax, ay] = toView(a.x1, a.y1);
      const [bx, by] = toView(a.x2, a.y2);
      return { x: Math.min(ax, bx), y: Math.min(ay, by), w: Math.abs(bx - ax), h: Math.abs(by - ay) };
    };

    drawDecorations(svg, p, vp);

    p.annots.forEach((a, idx) => {
      if (a.layer && !objects.isVisible(a.layer)) return;
      let node;
      switch (a.type) {
        case 'highlight': { const b = box(a); node = svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: a.color, 'fill-opacity': 0.35 }); break; }
        case 'rect': { const b = box(a); node = svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: 'transparent', stroke: a.color, 'stroke-width': a.width * s }); break; }
        case 'ellipse': { const b = box(a); node = svgEl('ellipse', { cx: b.x + b.w / 2, cy: b.y + b.h / 2, rx: b.w / 2, ry: b.h / 2, fill: 'transparent', stroke: a.color, 'stroke-width': a.width * s }); break; }
        case 'redact': { const b = box(a); node = svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: '#000', stroke: '#e11d48', 'stroke-width': 1.5, 'stroke-dasharray': '4 3' }); break; }
        case 'line': {
          const [x1, y1] = toView(a.x1, a.y1); const [x2, y2] = toView(a.x2, a.y2);
          node = svgEl('g', {});
          node.appendChild(svgEl('line', { x1, y1, x2, y2, stroke: a.color, 'stroke-width': a.width * s, 'stroke-linecap': 'round' }));
          if (a.arrow) {
            const ang = Math.atan2(y2 - y1, x2 - x1); const len = Math.max(8, a.width * 4) * s;
            for (const d of [Math.PI * 0.82, -Math.PI * 0.82]) node.appendChild(svgEl('line', { x1: x2, y1: y2, x2: x2 + len * Math.cos(ang + d), y2: y2 + len * Math.sin(ang + d), stroke: a.color, 'stroke-width': a.width * s, 'stroke-linecap': 'round' }));
          }
          break;
        }
        case 'ink':
          node = svgEl('polyline', { points: a.points.map(([x, y]) => toView(x, y).join(',')).join(' '), fill: 'none', stroke: a.color, 'stroke-width': a.width * s, 'stroke-linecap': 'round', 'stroke-linejoin': 'round' });
          break;
        case 'markup':
          node = svgEl('g', {});
          for (const r of a.rects) {
            const b = box({ x1: r[0], y1: r[1], x2: r[2], y2: r[3] });
            if (a.style === 'highlight') node.appendChild(svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: a.color, 'fill-opacity': 0.38 }));
            else {
              const ly = a.style === 'strike' ? b.y + b.h * 0.52 : b.y + b.h * 0.92;
              node.appendChild(svgEl('line', { x1: b.x, y1: ly, x2: b.x + b.w, y2: ly, stroke: a.color, 'stroke-width': Math.max(1, b.h / 14) }));
            }
            // a transparent box makes thin underlines easy to hit with the eraser
            node.appendChild(svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: 'transparent' }));
          }
          break;
        case 'text': {
          const [x, y] = toView(a.x, a.y);
          node = svgEl('text', { x, y, fill: a.color, 'font-family': 'Helvetica, Arial, sans-serif', 'font-size': a.size * s, transform: `rotate(${viewRot - (a.rot || 0)} ${x} ${y})` });
          a.text.split('\n').forEach((line, i) => {
            const t = svgEl('tspan', { x, y: y + i * a.size * 1.2 * s });
            t.textContent = line || ' ';
            node.appendChild(t);
          });
          break;
        }
        case 'editText': {
          const b = box({ x1: a.box.x0, y1: a.box.y0, x2: a.box.x1, y2: a.box.y1 });
          node = svgEl('g', {});
          node.appendChild(svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, fill: '#fff' }));
          const t = svgEl('text', { x: b.x + 1, y: b.y + b.h * 0.78, fill: a.color || a.seenColor || '#111', 'font-family': a.css || 'Helvetica, Arial, sans-serif', 'font-size': (a.size || a.seenSize) * s, 'font-weight': a.bold ? 700 : 400 });
          t.textContent = a.text;
          node.appendChild(t);
          node.appendChild(svgEl('rect', { x: b.x, y: b.y, width: Math.max(b.w, t.getComputedTextLength?.() || 0), height: b.h, fill: 'none', stroke: '#6366f1', 'stroke-width': 1, 'stroke-dasharray': '3 2', class: 'edit-outline' }));
          break;
        }
        case 'image': {
          const [x, y] = toView(a.x, a.y);
          node = svgEl('image', { href: a.dataUrl, x, y: y - a.h * s, width: a.w * s, height: a.h * s, transform: `rotate(${viewRot - (a.rot || 0)} ${x} ${y})` });
          break;
        }
        case 'stamp': {
          const [x, y] = toView(a.x, a.y);
          node = svgEl('g', { transform: `rotate(${viewRot - (a.rot || 0)} ${x} ${y})`, opacity: 0.9 });
          node.appendChild(svgEl('rect', { x, y: y - a.h * s, width: a.w * s, height: a.h * s, rx: 5 * s, fill: a.color, 'fill-opacity': 0.06, stroke: a.color, 'stroke-width': 2.2 * s }));
          const ls = (a.labelSize || 16) * s;
          const label = svgEl('text', { x: x + (a.w * s) / 2, y: y - a.h * s + 6 * s + ls * 0.74, fill: a.color, 'font-family': 'Helvetica, Arial, sans-serif', 'font-weight': 700, 'font-size': ls, 'text-anchor': 'middle' });
          label.textContent = a.label;
          node.appendChild(label);
          if (a.sub) {
            const sub = svgEl('text', { x: x + (a.w * s) / 2, y: y - 6 * s - ls * 0.1, fill: a.color, 'font-family': 'Helvetica, Arial, sans-serif', 'font-size': ls * 0.5, 'text-anchor': 'middle' });
            sub.textContent = a.sub;
            node.appendChild(sub);
          }
          break;
        }
        case 'field': {
          const b = box(a);
          node = svgEl('g', { class: 'annot-field' });
          const round = a.kind === 'radio';
          node.appendChild(round
            ? svgEl('ellipse', { cx: b.x + b.w / 2, cy: b.y + b.h / 2, rx: b.w / 2, ry: b.h / 2, fill: '#dbeafe', 'fill-opacity': 0.7, stroke: '#3b82f6', 'stroke-width': 1.2 })
            : svgEl('rect', { x: b.x, y: b.y, width: b.w, height: b.h, rx: 2, fill: '#dbeafe', 'fill-opacity': 0.7, stroke: '#3b82f6', 'stroke-width': 1.2 }));
          if (a.kind === 'dropdown') node.appendChild(svgEl('path', { d: `M${b.x + b.w - 14} ${b.y + b.h / 2 - 2} l4 4 l4 -4`, fill: 'none', stroke: '#1d4ed8', 'stroke-width': 1.4 }));
          if (a.kind === 'checkbox' && a.checked) node.appendChild(svgEl('path', { d: `M${b.x + b.w * 0.2} ${b.y + b.h * 0.55} L${b.x + b.w * 0.42} ${b.y + b.h * 0.78} L${b.x + b.w * 0.82} ${b.y + b.h * 0.25}`, fill: 'none', stroke: '#1d4ed8', 'stroke-width': 1.6 }));
          if (b.w > 40 && b.h > 10) {
            const fs = Math.min(11, b.h * 0.6);
            const t = svgEl('text', { x: b.x + 4, y: b.y + Math.min(b.h / 2 + fs * 0.35, fs + 3), fill: '#1e40af', 'font-family': 'Inter, Helvetica, Arial, sans-serif', 'font-size': fs, 'fill-opacity': 0.85 });
            t.textContent = a.value && a.kind !== 'radio' ? a.value : a.name;
            node.appendChild(t);
          }
          const tip = svgEl('title', {}); tip.textContent = `${FIELD_LABEL[a.kind]}: ${a.name}${a.kind === 'radio' ? ` (${a.value})` : ''}`; node.appendChild(tip);
          break;
        }
        case 'note': {
          const [x, y] = toView(a.x, a.y);
          const n = 20 * s;
          node = svgEl('g', { class: 'annot-note' });
          node.appendChild(svgEl('rect', { x, y, width: n, height: n, rx: 3 * s, fill: a.color, stroke: '#3f3f46', 'stroke-width': 0.8 * s }));
          for (const [fy, fx] of [[0.32, 0.8], [0.5, 0.8], [0.68, 0.6]]) node.appendChild(svgEl('line', { x1: x + n * 0.2, y1: y + n * fy, x2: x + n * fx, y2: y + n * fy, stroke: '#3f3f46', 'stroke-width': 1 * s }));
          const t = svgEl('title', {}); t.textContent = a.text; node.appendChild(t);
          break;
        }
        default:
      }
      if (node) {
        // lines and pen strokes are thin: a wider invisible copy makes them easy to pick up
        if (a.type === 'line' || a.type === 'ink') {
          const g = node.tagName === 'g' ? node : svgEl('g', {});
          const hit = (node.tagName === 'g' ? node.firstChild : node).cloneNode();
          hit.setAttribute('stroke', 'transparent'); hit.setAttribute('stroke-width', 12); hit.classList.add('hit');
          if (g !== node) g.appendChild(node);
          g.appendChild(hit);
          node = g;
        }
        node.classList.add('annot');
        if (isMovable(a)) node.classList.add('movable');
        node.dataset.idx = idx; node.querySelectorAll('*').forEach((c) => { c.dataset.idx = idx; });
        svg.appendChild(node);
      }
    });
    objects.drawFrame(p);
  }

  // Watermark and header/footer previews, matching the saved output
  function drawDecorations(svg, p, vp) {
    const s = vp.scale;
    const vw = vp.width; const vh = vp.height;
    if (watermark?.text) {
      const fs = (watermark.size || Math.min(vw / s, vh / s) / Math.max(4, watermark.text.length * 0.62)) * s;
      const t = svgEl('text', { x: vw / 2, y: vh / 2, fill: watermark.color, 'fill-opacity': watermark.opacity, 'font-family': 'Helvetica, Arial, sans-serif', 'font-weight': 700, 'font-size': fs, 'text-anchor': 'middle', 'dominant-baseline': 'central', transform: `rotate(${-(watermark.angle ?? 45)} ${vw / 2} ${vh / 2})`, class: 'decor' });
      t.textContent = watermark.text;
      svg.appendChild(t);
    }
    if (headerFooter) {
      const hf = headerFooter;
      const fs = (hf.size || 9) * s;
      const m = (hf.margin ?? 28) * s;
      const n = pages.indexOf(p) + (hf.start ?? 1);
      const total = pages.length + (hf.start ?? 1) - 1;
      const fill = (tpl) => String(tpl || '').replace(/\{n\}/gi, n).replace(/\{total\}/gi, total).replace(/\{date\}/gi, today()).replace(/\{file\}/gi, name);
      const rows = [[hf.header, m + fs * 0.7], [hf.footer, vh - m]];
      for (const [cells, y] of rows) {
        (cells || []).forEach((tpl, i) => {
          const text = fill(tpl);
          if (!text.trim()) return;
          const anchor = ['start', 'middle', 'end'][i];
          const x = [m, vw / 2, vw - m][i];
          const t = svgEl('text', { x, y, fill: '#475569', 'font-family': 'Helvetica, Arial, sans-serif', 'font-size': fs, 'text-anchor': anchor, class: 'decor' });
          t.textContent = text;
          svg.appendChild(t);
        });
      }
    }
  }
  const redrawAll = () => { for (const p of pages) drawAnnots(p); };

  // ---------------------------------------------------------------- tools
  function applyToolClass() {
    for (const t of Object.keys(TOOLS)) el.querySelectorAll(`[data-id="tool-${t}"]`).forEach((b) => b.classList.toggle('on', t === tool || (t === 'sign' && tool === 'place-sig')));
    pagesEl.querySelectorAll('.pdf-page').forEach((d) => {
      d.classList.toggle('tool-active', !TEXT_SELECT_TOOLS.has(tool));
      d.classList.toggle('tool-markup', tool.startsWith('mark-'));
      d.classList.toggle('tool-erase', tool === 'erase');
    });
  }

  async function setTool(t) {
    if (t === 'sign') {
      if (!signaturePng) { openSignaturePad(); return; }
      t = 'place-sig';
    }
    if (t === 'image') {
      const [img] = await openFiles('image', false);
      if (!img) return;
      const fmt = /\.png$/i.test(img.name) ? 'png' : 'jpg';
      const dataUrl = URL.createObjectURL(new Blob([img.data], { type: fmt === 'png' ? 'image/png' : 'image/jpeg' }));
      const dims = await new Promise((r) => { const im = new Image(); im.onload = () => r([im.naturalWidth, im.naturalHeight]); im.onerror = () => r([4, 3]); im.src = dataUrl; });
      placeImage = { bytes: img.data, fmt, dataUrl, aspect: dims[0] / dims[1] };
    }
    tool = t;
    applyToolClass();
    setStatus(TOOLS[t === 'place-sig' ? 'sign' : t] || '');
    if (t === 'redact') toast('Redacted areas are removed for good when you save. Those pages are saved as images.');
  }

  function addAnnot(p, a) {
    const layer = objects.active;
    if (layer && a.layer === undefined) a.layer = layer;
    p.annots.push(a);
    if (a.layer) objects.stack(p);
    undoStack.push({ type: 'add', p, a });
    drawAnnots(p);
    changed();
    objects.renderPanel();
    // a placed picture or signature is selected straight away, ready to move and resize
    if (a.type === 'image') objects.select({ p, a });
  }

  // ---------------------------------------------------------------- new form fields
  const allFields = () => pages.flatMap((pg) => pg.annots.filter((x) => x.type === 'field'));
  function suggestName(kind) {
    const base = { text: 'Text', multiline: 'Paragraph', checkbox: 'Check box', dropdown: 'Drop-down', list: 'List' }[kind] || 'Field';
    const used = new Set([...allFields().map((f) => f.name), ...(forms[main.id] || []).map((f) => f.name)]);
    let n = 1;
    while (used.has(`${base} ${n}`)) n++;
    return `${base} ${n}`;
  }
  async function placeField(p, draft, entry, sx, sy) {
    const ft = Object.values(FIELD_TOOLS).find((f) => f.kind === draft.kind);
    const small = Math.abs(draft.x2 - draft.x1) < 6 || Math.abs(draft.y2 - draft.y1) < 6;
    if (small || draft.kind === 'checkbox' || draft.kind === 'radio') {
      // a click places a field of the usual size with its top-left corner at the pointer
      const vs = entry.viewport.scale;
      const side = Math.max(Math.abs(draft.x2 - draft.x1), Math.abs(draft.y2 - draft.y1));
      const [w, h] = (draft.kind === 'checkbox' || draft.kind === 'radio') ? [small ? ft.w : side, small ? ft.h : side] : [ft.w, ft.h];
      const [x1, y1] = entry.viewport.convertToPdfPoint(sx, sy);
      const [x2, y2] = entry.viewport.convertToPdfPoint(sx + w * vs, sy + h * vs);
      Object.assign(draft, { x1, y1, x2, y2 });
    }
    if (draft.kind === 'radio') {
      const last = [...allFields()].reverse().find((f) => f.kind === 'radio');
      draft.name = last?.name || 'Choice 1';
      draft.value = `Option ${allFields().filter((f) => f.kind === 'radio' && f.name === draft.name).length + 1}`;
    } else draft.name = suggestName(draft.kind);
    if (draft.kind === 'dropdown' || draft.kind === 'list') {
      draft.options = ['Option 1', 'Option 2', 'Option 3'];
      if (!(await editField(draft, true))) { drawAnnots(p); return; }
    }
    addAnnot(p, draft);
    // choice buttons usually come in sets, so that tool stays on; other fields are ready to move
    if (draft.kind !== 'radio') setTool('select');
    objects.select({ p, a: draft });
  }
  /** Name, choices and other options of a new field. Returns false when cancelled. */
  async function editField(a, isNew = false) {
    const choice = a.kind === 'dropdown' || a.kind === 'list';
    const texty = a.kind === 'text' || a.kind === 'multiline';
    const groups = [...new Set(allFields().filter((f) => f.kind === 'radio').map((f) => f.name))];
    const res = await openDialog({
      title: `${FIELD_LABEL[a.kind]} options`,
      description: a.kind === 'radio' ? 'Choice buttons with the same group name work together: picking one clears the others.' : 'The name is what the field is called when the form is filled in or its data is read.',
      fields: [
        { id: 'name', label: a.kind === 'radio' ? 'Group name' : 'Name', value: a.name, hint: a.kind === 'radio' && groups.length > 1 ? `Groups in this document: ${groups.join(', ')}` : '' },
        ...(a.kind === 'radio' ? [{ id: 'value', label: 'This button\'s choice', value: a.value }] : []),
        ...(choice ? [{ id: 'options', label: 'Choices, one per line', type: 'textarea', rows: 5, value: (a.options || []).join('\n') }] : []),
        ...(choice ? [{ id: 'def', label: 'Chosen at first', type: 'select', value: a.value || '', options: [{ value: '', label: 'Nothing' }, ...(a.options || []).map((o) => ({ value: o, label: o }))] }] : []),
        ...(texty ? [{ id: 'def', label: 'Text at first', value: a.value || '', placeholder: 'Leave empty for a blank field' }] : []),
        ...(a.kind === 'checkbox' ? [{ id: 'checked', label: 'Ticked at first', type: 'checkbox', value: !!a.checked }] : []),
        ...(texty || choice ? [{ id: 'size', label: 'Text size', type: 'select', value: a.size || 0, options: [{ value: 0, label: 'Fit the box' }, ...[8, 9, 10, 11, 12, 14, 16, 18].map((v) => ({ value: v, label: `${v} pt` }))] }] : []),
        { id: 'required', label: 'Must be filled in', type: 'checkbox', value: !!a.required },
      ],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: isNew ? 'Add field' : 'OK', primary: true }],
      validate: (v) => (!String(v.name || '').trim() ? 'Give the field a name' : choice && !String(v.options || '').trim() ? 'Add at least one choice' : null),
      onChange: (v, form) => {
        if (!choice) return;
        // keep "Chosen at first" in step with the list of choices
        const sel = form.querySelector('[data-f="def"]');
        const items = String(v.options || '').split('\n').map((x) => x.trim()).filter(Boolean);
        const cur = sel.value;
        sel.innerHTML = ['<option value="">Nothing</option>', ...items.map((o) => `<option${o === cur ? ' selected' : ''}>${esc(o)}</option>`)].join('');
      },
    });
    if (!res || res.action !== 'ok') return false;
    const v = res.values;
    a.name = String(v.name).trim();
    if (a.kind === 'radio') a.value = String(v.value || '').trim() || 'Option';
    if (choice) { a.options = String(v.options).split('\n').map((x) => x.trim()).filter(Boolean); a.value = a.options.includes(v.def) ? v.def : ''; }
    if (texty) a.value = v.def || '';
    if (a.kind === 'checkbox') a.checked = !!v.checked;
    if (texty || choice) a.size = Number(v.size) || 0;
    a.required = !!v.required;
    if (!isNew) { const p = pages.find((pg) => pg.annots.includes(a)); if (p) drawAnnots(p); changed(); }
    return true;
  }

  // ---------------------------------------------------------------- compress
  const mb = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(n / 1024))} KB`);
  async function recodeImage({ data, kind, width, height, channels, maxSide, quality }) {
    const s = Math.min(1, maxSide / Math.max(width, height));
    const w = Math.max(1, Math.round(width * s)); const h = Math.max(1, Math.round(height * s));
    let source;
    if (kind === 'jpeg') source = await createImageBitmap(new Blob([data], { type: 'image/jpeg' }));
    else {
      const rgba = new Uint8ClampedArray(width * height * 4);
      for (let i = 0, j = 0; i < width * height; i++, j += channels) {
        const r = data[j]; const g = channels === 3 ? data[j + 1] : r; const b = channels === 3 ? data[j + 2] : r;
        rgba[i * 4] = r; rgba[i * 4 + 1] = g; rgba[i * 4 + 2] = b; rgba[i * 4 + 3] = 255;
      }
      source = document.createElement('canvas');
      source.width = width; source.height = height;
      source.getContext('2d').putImageData(new ImageData(rgba, width, height), 0, 0);
    }
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    const ctx = c.getContext('2d');
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(source, 0, 0, w, h);
    source.close?.();
    const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', quality));
    return blob ? { bytes: new Uint8Array(await blob.arrayBuffer()), width: w, height: h } : null;
  }
  async function compressCopy() {
    const before = sources[main.id]?.bytes?.length || 0;
    const res = await openDialog({
      title: 'Compress PDF',
      description: 'Saves a smaller copy. Pictures are resized and re-saved; text and drawings stay sharp. This document is not changed.',
      fields: [
        { id: 'level', type: 'segmented', label: 'Pictures', value: 'balanced', options: [
          { value: 'strong', label: 'Smallest file' }, { value: 'balanced', label: 'Balanced' }, { value: 'light', label: 'Best quality' }, { value: 'none', label: 'Unchanged' },
        ] },
        { id: 'info', type: 'info', label: `This file is ${mb(before)} now.` },
      ],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Compress and save', primary: true }],
    });
    if (!res || res.action !== 'ok') return;
    setStatus('Compressing...');
    try {
      const level = res.values.level === 'none' ? null : res.values.level;
      const data = await buildOutput(pages, { compress: { images: level, recodeImage } });
      const base = name.replace(/\.pdf$/i, '');
      const out = await saveFile({ data, name: `${base} (compressed).pdf`, kind: 'pdf', saveAs: true });
      if (out) {
        const pct = before ? Math.round((1 - data.length / before) * 100) : 0;
        toast(pct > 0 ? `Saved ${out.name}: ${mb(before)} to ${mb(data.length)}, ${pct}% smaller` : `Saved ${out.name} (${mb(data.length)}). This PDF was already compact.`);
      }
    } catch (err) {
      toast(`Could not compress: ${err.message}`, true);
    }
    setStatus(`${pages.length} pages`);
  }

  function measure(text, px, bold) {
    const ctx = measure.ctx || (measure.ctx = document.createElement('canvas').getContext('2d'));
    ctx.font = `${bold ? 'bold ' : ''}${px}px Helvetica, Arial, sans-serif`;
    return ctx.measureText(text).width;
  }

  function bindPointer(div, svg, p, entry) {
    const local = (e) => { const r = div.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
    const toPdf = (x, y) => entry.viewport.convertToPdfPoint(x, y);
    const vs = () => entry.viewport.scale;
    const rot = () => entry.viewport.rotation;

    // Notes are clickable in every mode except erase
    svg.addEventListener('click', async (e) => {
      const g = e.target.closest('.annot-note');
      if (!g || tool === 'erase') return;
      const a = p.annots[Number(g.dataset.idx)];
      const res = await noteDialog(a.text, true);
      if (res === 'delete') { p.annots.splice(p.annots.indexOf(a), 1); undoStack.push({ type: 'remove', p, a, idx: Number(g.dataset.idx) }); }
      else if (res) a.text = res.text;
      if (res) { drawAnnots(p); changed(); }
    });

    svg.addEventListener('dblclick', (e) => {
      // the drag that selected the field holds the pointer, so look at what is under it
      const under = document.elementsFromPoint(e.clientX, e.clientY);
      const g = under.find((n) => n.closest?.('.annot-field'))?.closest('.annot-field');
      const onFrame = under.some((n) => n.closest?.('.sel-frame'));
      const a = g ? p.annots[Number(g.dataset.idx)] : onFrame && objects.selected?.p === p ? objects.selected.a : null;
      if (a?.type === 'field') editField(a);
    });
    svg.addEventListener('mousedown', (e) => { if (!TEXT_SELECT_TOOLS.has(tool)) e.preventDefault(); });
    svg.addEventListener('pointerdown', async (e) => {
      if (tool === 'select' && objects.pointerDown(e, p)) return;
      if (TEXT_SELECT_TOOLS.has(tool) || e.target.closest('.annot-note')) return;
      e.preventDefault();
      const [sx, sy] = local(e);

      if (tool === 'erase') {
        const idx = e.target?.dataset?.idx;
        if (idx !== undefined) {
          const [a] = p.annots.splice(Number(idx), 1);
          undoStack.push({ type: 'remove', p, a, idx: Number(idx) });
          drawAnnots(p);
          changed();
        }
        return;
      }
      if (tool === 'text') { startTextEntry(div, p, entry, sx, sy); return; }
      if (tool === 'edit-text') { startTextEdit(div, p, entry, e.clientX, e.clientY); return; }
      if (tool === 'date') {
        const fs = size();
        const [ax, ay] = toPdf(sx, sy + fs * vs() * 0.35);
        addAnnot(p, { type: 'text', x: ax, y: ay, size: fs, color: colorFor('date'), text: today(), rot: rot() });
        return;
      }
      if (tool === 'check' || tool === 'cross') {
        const d = size() * 1.1 * vs();
        const pts = (list) => list.map(([fx, fy]) => toPdf(sx - d / 2 + fx * d, sy - d / 2 + fy * d));
        const w = Math.max(1.2, size() / 8);
        if (tool === 'check') addAnnot(p, { type: 'ink', points: pts([[0.05, 0.55], [0.38, 0.88], [0.95, 0.12]]), color: colorFor('check'), width: w });
        else {
          addAnnot(p, { type: 'ink', points: pts([[0.1, 0.1], [0.9, 0.9]]), color: colorFor('cross'), width: w });
          addAnnot(p, { type: 'ink', points: pts([[0.9, 0.1], [0.1, 0.9]]), color: colorFor('cross'), width: w });
        }
        return;
      }
      if (tool === 'note') {
        const res = await noteDialog('');
        if (!res) return;
        const [ax, ay] = toPdf(sx, sy);
        addAnnot(p, { type: 'note', x: ax, y: ay, text: res.text, color: colorFor('note'), author: 'Kertas Office' });
        return;
      }
      if (tool === 'stamp') {
        const st = STAMPS.find((x) => x.id === stampKind) || STAMPS[0];
        const ls = 16;
        const sub = `${today()}`;
        const w = Math.max(measure(st.label, ls, true) * 1.06, measure(sub, ls * 0.5)) + 22;
        const h = ls * 0.74 + ls * 0.5 + 4 + 12;
        const [ax, ay] = toPdf(sx - (w * vs()) / 2, sy + (h * vs()) / 2);
        addAnnot(p, { type: 'stamp', x: ax, y: ay, w, h, rot: rot(), label: st.label, sub, color: userColor || st.color, labelSize: ls });
        return;
      }
      if ((tool === 'place-sig' && signaturePng) || (tool === 'image' && placeImage)) {
        const src = tool === 'image' ? placeImage : signaturePng;
        const wView = tool === 'image' ? Math.min(260 * vs(), div.clientWidth * 0.6) : 180;
        const hView = wView / src.aspect;
        // keep the whole picture on the page
        const px = Math.max(0, Math.min(sx, div.clientWidth - wView));
        const py = Math.max(0, Math.min(sy, div.clientHeight - hView));
        const [ax, ay] = toPdf(px, py + hView); // bottom-left of the box as seen on screen
        addAnnot(p, { type: 'image', x: ax, y: ay, w: wView / vs(), h: hView / vs(), rot: rot(), img: src.bytes, fmt: src.fmt || 'png', dataUrl: src.dataUrl });
        if (tool === 'image') setTool('select');
        return;
      }

      // Drag tools
      const fieldTool = FIELD_TOOLS[tool];
      if (!DRAG_TOOLS.has(tool) && tool !== 'ink' && !fieldTool) return;
      svg.setPointerCapture(e.pointerId);
      const start = toPdf(sx, sy);
      const draft = fieldTool
        ? { type: 'field', kind: fieldTool.kind, x1: start[0], y1: start[1], x2: start[0], y2: start[1], rot: rot() }
        : tool === 'ink'
        ? { type: 'ink', points: [start], color: colorFor('ink'), width: Math.max(1, size() / 6) }
        : tool === 'arrow' || tool === 'line'
          ? { type: 'line', arrow: tool === 'arrow', x1: start[0], y1: start[1], x2: start[0], y2: start[1], color: colorFor(tool), width: Math.max(1, size() / 7) }
          : { type: tool, x1: start[0], y1: start[1], x2: start[0], y2: start[1], color: colorFor(tool), width: Math.max(1, size() / 7) };
      p.annots.push(draft);
      const move = (ev) => {
        const pt = toPdf(...local(ev));
        if (draft.type === 'ink') draft.points.push(pt); else { draft.x2 = pt[0]; draft.y2 = pt[1]; }
        drawAnnots(p);
      };
      const up = () => {
        svg.removeEventListener('pointermove', move);
        svg.removeEventListener('pointerup', up);
        p.annots.pop();
        if (draft.type === 'field') { placeField(p, draft, entry, sx, sy); return; }
        const tiny = draft.type === 'ink' ? draft.points.length < 2
          : draft.type === 'line' ? Math.hypot(draft.x2 - draft.x1, draft.y2 - draft.y1) < 3
            : Math.abs(draft.x2 - draft.x1) < 2 || Math.abs(draft.y2 - draft.y1) < 2;
        if (!tiny) addAnnot(p, draft); else drawAnnots(p);
      };
      svg.addEventListener('pointermove', move);
      svg.addEventListener('pointerup', up);
    });
  }

  // Text markup: select text with a markup tool, then the selection becomes the markup
  pagesEl.addEventListener('mouseup', () => {
    if (!tool.startsWith('mark-')) return;
    setTimeout(applySelectionMarkup, 0);
  });
  function applySelectionMarkup() {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed || !sel.rangeCount) return;
    const style = { 'mark-hl': 'highlight', 'mark-ul': 'underline', 'mark-st': 'strike' }[tool];
    const byPage = new Map();
    for (let i = 0; i < sel.rangeCount; i++) {
      for (const r of sel.getRangeAt(i).getClientRects()) {
        if (r.width < 1 || r.height < 1) continue;
        const pageDiv = document.elementsFromPoint(r.left + r.width / 2, r.top + r.height / 2).map((n) => n.closest?.('.pdf-page')).find(Boolean);
        if (!pageDiv || !el.contains(pageDiv)) continue;
        const entry = rendered.get(pageDiv.dataset.key);
        if (!entry?.viewport) continue;
        const pr = pageDiv.getBoundingClientRect();
        const [x1, y1] = entry.viewport.convertToPdfPoint(r.left - pr.left, r.top - pr.top);
        const [x2, y2] = entry.viewport.convertToPdfPoint(r.right - pr.left, r.bottom - pr.top);
        if (!byPage.has(pageDiv.dataset.key)) byPage.set(pageDiv.dataset.key, []);
        byPage.get(pageDiv.dataset.key).push([Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)]);
      }
    }
    sel.removeAllRanges();
    for (const [key, rects] of byPage) {
      const p = pages.find((x) => x.key === key);
      if (p) addAnnot(p, { type: 'markup', style, rects: mergeRects(rects), color: colorFor(tool) });
    }
  }
  // pdf.js splits text into many spans; join pieces that sit on the same line
  function mergeRects(rects) {
    rects.sort((a, b) => b[1] - a[1] || a[0] - b[0]);
    const out = [];
    for (const r of rects) {
      const last = out[out.length - 1];
      const sameLine = last && Math.abs(last[1] - r[1]) < (last[3] - last[1]) * 0.5 && r[0] <= last[2] + 4;
      if (sameLine) { last[0] = Math.min(last[0], r[0]); last[2] = Math.max(last[2], r[2]); last[1] = Math.min(last[1], r[1]); last[3] = Math.max(last[3], r[3]); } else out.push([...r]);
    }
    return out;
  }

  function startTextEntry(div, p, entry, x, y) {
    const fs = size();
    const ta = document.createElement('textarea');
    ta.className = 'pdf-text-input';
    ta.style.left = `${x}px`;
    ta.style.top = `${y}px`;
    ta.style.fontSize = `${fs * entry.viewport.scale}px`;
    ta.style.color = colorFor('text');
    ta.rows = 1;
    div.appendChild(ta);
    ta.focus();
    let done = false;
    const commit = () => {
      if (done) return;
      done = true;
      const text = ta.value.replace(/\s+$/, '');
      ta.remove();
      if (!text) return;
      const [ax, ay] = entry.viewport.convertToPdfPoint(x + 7, y + 3 + fs * entry.viewport.scale * 0.95);
      addAnnot(p, { type: 'text', x: ax, y: ay, size: fs, color: colorFor('text'), text, rot: entry.viewport.rotation });
    };
    ta.addEventListener('blur', commit);
    ta.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); ta.blur(); }
      if (e.key === 'Escape') { done = true; ta.remove(); }
      e.stopPropagation();
    });
    ta.addEventListener('input', () => { ta.rows = ta.value.split('\n').length; });
  }

  function cssForSpan(span) {
    const fam = span.dataset.family;
    const f = fam && findFont(fam);
    return f ? cssStack(f) : null;
  }

  // Changing the PDF's own words: pick the line under the pointer from the text layer,
  // type over it, and the old glyphs are removed from the page when it is saved
  function startTextEdit(div, p, entry, cx, cy) {
    const vp = entry.viewport;
    const pr = div.getBoundingClientRect();
    const inside = (r) => cx >= r.left - 2 && cx <= r.right + 2 && cy >= r.top - 2 && cy <= r.bottom + 2;
    // An earlier edit at this spot is edited again
    const prevIdx = p.annots.findIndex((a) => {
      if (a.type !== 'editText') return false;
      const [x1, y1] = vp.convertToViewportPoint(a.box.x0, a.box.y0);
      const [x2, y2] = vp.convertToViewportPoint(a.box.x1, a.box.y1);
      return inside({ left: pr.left + Math.min(x1, x2), right: pr.left + Math.max(x1, x2), top: pr.top + Math.min(y1, y2), bottom: pr.top + Math.max(y1, y2) });
    });
    const prev = prevIdx >= 0 ? p.annots[prevIdx] : null;
    let rect; let text; let fontPx; let family;
    if (prev) {
      const [x1, y1] = vp.convertToViewportPoint(prev.box.x0, prev.box.y0);
      const [x2, y2] = vp.convertToViewportPoint(prev.box.x1, prev.box.y1);
      rect = { left: pr.left + Math.min(x1, x2), top: pr.top + Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
      text = prev.text; fontPx = (prev.size || prev.seenSize) * vp.scale; family = prev.css;
    } else {
      const span = [...div.querySelectorAll('.textLayer span')].find((sp) => sp.textContent.trim() && inside(sp.getBoundingClientRect()));
      if (!span) { toast('There is no text there. Click on a line of the PDF\'s own text (scans need Recognise text first).'); return; }
      const r = span.getBoundingClientRect();
      rect = { left: r.left, top: r.top, width: r.width, height: r.height };
      text = span.textContent; fontPx = r.height; family = getComputedStyle(span).fontFamily;
      // The text layer only knows serif / sans-serif; the page's real font gives a truer preview
      const known = cssForSpan(span);
      if (known) family = known;
    }
    if (vp.rotation % 180 !== 0) { toast('Turn the page upright to edit its text.'); return; }
    const ta = document.createElement('textarea');
    ta.className = 'pdf-text-input pdf-edit-input';
    ta.style.left = `${rect.left - pr.left - 3}px`;
    ta.style.top = `${rect.top - pr.top - 3}px`;
    ta.style.minWidth = `${rect.width + 24}px`;
    ta.style.fontSize = `${fontPx}px`;
    ta.style.fontFamily = /g_d\d/.test(family || '') ? 'Helvetica, Arial, sans-serif' : family || '';
    ta.rows = 1;
    ta.value = text;
    div.appendChild(ta);
    ta.focus();
    ta.select();
    let done = false;
    const commit = () => {
      if (done) return;
      done = true;
      const value = ta.value.replace(/\n+/g, ' ');
      ta.remove();
      if (value === text) return;
      if (prev) {
        prev.text = value;
        undoStack.push({ type: 'edit', p, a: prev, old: text });
        drawAnnots(p);
        changed();
        return;
      }
      const l = rect.left - pr.left; const t = rect.top - pr.top;
      const [ax, ay] = vp.convertToPdfPoint(l - 1, t + rect.height + 1);
      const [bx, by] = vp.convertToPdfPoint(l + rect.width + 1, t - 1);
      const boxPdf = { x0: Math.min(ax, bx), y0: Math.min(ay, by), x1: Math.max(ax, bx), y1: Math.max(ay, by) };
      addAnnot(p, { type: 'editText', box: boxPdf, text: value, seenSize: fontPx / vp.scale, css: ta.style.fontFamily, original: text });
    };
    ta.addEventListener('blur', commit);
    ta.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); ta.blur(); }
      if (e.key === 'Escape') { done = true; ta.remove(); }
      e.stopPropagation();
    });
  }

  function openSignaturePad() {
    const overlay = document.createElement('div');
    overlay.className = 'modal-backdrop';
    overlay.innerHTML = `
      <div class="modal glass">
        <h3>Draw your signature</h3>
        <p>Use your mouse, trackpad or pen. It is reused until you draw a new one.</p>
        <canvas width="1040" height="360"></canvas>
        <div class="row">
          <button class="chip" data-s="clear">Clear</button>
          <button class="chip" data-s="cancel">Cancel</button>
          <button class="chip primary" data-s="use">${icon('check', 14)}Use signature</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    const cv = overlay.querySelector('canvas');
    const ctx = cv.getContext('2d');
    ctx.lineWidth = 5; ctx.lineCap = 'round'; ctx.lineJoin = 'round'; ctx.strokeStyle = '#1d2433';
    let drawing = false; let drew = false;
    const pos = (e) => { const r = cv.getBoundingClientRect(); return [(e.clientX - r.left) * (cv.width / r.width), (e.clientY - r.top) * (cv.height / r.height)]; };
    cv.addEventListener('pointerdown', (e) => { drawing = true; drew = true; cv.setPointerCapture(e.pointerId); ctx.beginPath(); ctx.moveTo(...pos(e)); });
    cv.addEventListener('pointermove', (e) => { if (drawing) { ctx.lineTo(...pos(e)); ctx.stroke(); } });
    cv.addEventListener('pointerup', () => { drawing = false; });
    overlay.addEventListener('keydown', (e) => { if (e.key === 'Escape') overlay.remove(); e.stopPropagation(); });
    overlay.addEventListener('click', async (e) => {
      const act = e.target.closest('[data-s]')?.dataset.s;
      if (act === 'clear') { ctx.clearRect(0, 0, cv.width, cv.height); drew = false; }
      if (act === 'cancel') overlay.remove();
      if (act === 'use') {
        if (!drew) { toast('Draw a signature first'); return; }
        const cropped = cropCanvas(cv);
        const dataUrl = cropped.toDataURL('image/png');
        const bytes = new Uint8Array(await (await fetch(dataUrl)).arrayBuffer());
        signaturePng = { bytes, dataUrl, aspect: cropped.width / cropped.height, fmt: 'png' };
        overlay.remove();
        tool = 'place-sig';
        applyToolClass();
        setStatus('Click on a page to place your signature');
      }
    });
  }

  function cropCanvas(cv) {
    const { width, height } = cv;
    const px = cv.getContext('2d').getImageData(0, 0, width, height).data;
    let minX = width; let minY = height; let maxX = 0; let maxY = 0;
    for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
      if (px[(y * width + x) * 4 + 3] > 0) { minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y); }
    }
    const pad = 6;
    minX = Math.max(0, minX - pad); minY = Math.max(0, minY - pad);
    maxX = Math.min(width - 1, maxX + pad); maxY = Math.min(height - 1, maxY + pad);
    const out = document.createElement('canvas');
    out.width = maxX - minX + 1; out.height = maxY - minY + 1;
    out.getContext('2d').drawImage(cv, minX, minY, out.width, out.height, 0, 0, out.width, out.height);
    return out;
  }

  function undo() {
    const last = undoStack.pop();
    if (!last) return;
    if (last.type === 'add') last.p.annots.splice(last.p.annots.indexOf(last.a), 1);
    if (last.type === 'remove') last.p.annots.splice(last.idx, 0, last.a);
    if (last.type === 'edit') last.a.text = last.old;
    objects.undoStep(last);
    drawAnnots(last.p);
    objects.afterUndo(last.p);
    changed();
  }

  // ---------------------------------------------------------------- sidebar
  function toggleSidebar() {
    sideEl.classList.toggle('collapsed');
    ribbon.setOn('v-thumbs', !sideEl.classList.contains('collapsed'));
  }
  sideEl.querySelector('.side-tabs').addEventListener('click', (e) => {
    const b = e.target.closest('[data-side]');
    if (!b) return;
    sideEl.querySelectorAll('[data-side]').forEach((x) => x.classList.toggle('on', x === b));
    thumbsEl.hidden = b.dataset.side !== 'pages';
    outlineEl.hidden = b.dataset.side !== 'outline';
    attachEl.hidden = b.dataset.side !== 'attach';
    if (b.dataset.side === 'attach') renderAttachments();
    objects.panel.hidden = b.dataset.side !== 'layers';
    objects.renderPanel();
  });

  async function renderThumbs() {
    thumbsEl.innerHTML = '';
    for (let i = 0; i < pages.length; i++) {
      const p = pages[i];
      const pdfPage = await pdfPageFor(p);
      const tvp = viewportFor(pdfPage, p, 140 / viewportFor(pdfPage, p, 1).width);
      const wrap = document.createElement('div');
      wrap.className = 'thumb';
      wrap.dataset.key = p.key;
      wrap.draggable = true;
      wrap.innerHTML = `
        <div class="ops">
          <button data-op="rotl" title="Rotate left">${icon('rotate-l', 14)}</button>
          <button data-op="rotr" title="Rotate right">${icon('rotate-r', 14)}</button>
          <button data-op="dup" title="Duplicate">${icon('files', 14)}</button>
          <button data-op="del" title="Delete page">${icon('trash', 14)}</button>
        </div>
        <canvas width="${Math.floor(tvp.width)}" height="${Math.floor(tvp.height)}"></canvas>
        <div class="num">${i + 1}</div>`;
      thumbsEl.appendChild(wrap);
      pdfPage.render({ canvas: wrap.querySelector('canvas'), viewport: tvp }).promise.catch(() => {});
    }
    updatePageInfo();
  }

  thumbsEl.addEventListener('click', async (e) => {
    const wrap = e.target.closest('.thumb');
    if (!wrap) return;
    const i = pages.findIndex((p) => p.key === wrap.dataset.key);
    const op = e.target.closest('[data-op]')?.dataset.op;
    if (!op) {
      // the picked page is the current one at once, not when the smooth scroll gets there
      // (Delete or Rotate pressed straight after acted on the page that was in view before)
      pickedPage = i;
      updatePageInfo();
      pagesEl.querySelector(`.pdf-page[data-key="${wrap.dataset.key}"]`)?.scrollIntoView({ behavior: 'smooth' });
      return;
    }
    await pageOp(op, i);
  });

  // Drag thumbnails to reorder pages
  let dragKey = null;
  thumbsEl.addEventListener('dragstart', (e) => { dragKey = e.target.closest('.thumb')?.dataset.key; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'page'); });
  thumbsEl.addEventListener('dragover', (e) => {
    if (!dragKey && e.dataTransfer?.types.includes('Files')) { markFileDrop(e); return; }
    if (!dragKey) return;
    e.preventDefault();
    e.stopPropagation();
    const over = e.target.closest('.thumb');
    thumbsEl.querySelectorAll('.thumb').forEach((t) => t.classList.remove('drop-before', 'drop-after'));
    if (over && over.dataset.key !== dragKey) {
      const r = over.getBoundingClientRect();
      over.classList.add(e.clientY < r.top + r.height / 2 ? 'drop-before' : 'drop-after');
    }
  });
  thumbsEl.addEventListener('drop', async (e) => {
    if (!dragKey) return;
    e.preventDefault();
    e.stopPropagation();
    const over = thumbsEl.querySelector('.drop-before, .drop-after');
    const from = pages.findIndex((p) => p.key === dragKey);
    dragKey = null;
    if (!over || from < 0) { renderThumbs(); return; }
    const [moved] = pages.splice(from, 1);
    let to = pages.findIndex((p) => p.key === over.dataset.key);
    if (over.classList.contains('drop-after')) to += 1;
    pages.splice(to, 0, moved);
    undoStack.length = 0;
    changed();
    await layout();
    await renderThumbs();
  });
  thumbsEl.addEventListener('dragend', () => { dragKey = null; clearDropMarks(); });

  // Drop PDFs or pictures between the thumbnails to insert their pages there
  const clearDropMarks = () => thumbsEl.querySelectorAll('.thumb').forEach((t) => t.classList.remove('drop-before', 'drop-after'));
  function markFileDrop(e) {
    const thumbs = [...thumbsEl.querySelectorAll('.thumb')];
    if (!thumbs.length) return;
    // the thumbnail under the pointer, or the last one when below them all
    const over = e.target.closest('.thumb') || thumbs.find((t) => e.clientY < t.getBoundingClientRect().bottom) || thumbs[thumbs.length - 1];
    const r = over.getBoundingClientRect();
    const before = e.clientY < r.top + r.height / 2;
    if (over.classList.contains(before ? 'drop-before' : 'drop-after')) return;
    clearDropMarks();
    over.classList.add(before ? 'drop-before' : 'drop-after');
  }
  thumbsEl.addEventListener('dragleave', (e) => { if (!thumbsEl.contains(e.relatedTarget)) clearDropMarks(); });
  thumbsEl.dataset.fileDrop = '';
  thumbsEl.kertasFileDrop = async (files) => {
    const mark = thumbsEl.querySelector('.drop-before, .drop-after');
    let at = pages.length;
    if (mark) {
      at = pages.findIndex((p) => p.key === mark.dataset.key);
      if (mark.classList.contains('drop-after')) at += 1;
    }
    clearDropMarks();
    await insertFiles(files, at);
  };

  // Drop any files on the Files tab to attach them
  attachEl.dataset.fileDrop = '';
  attachEl.kertasFileDrop = async (files) => { attachEl.classList.remove('file-over'); await addAttachments(files); };
  attachEl.addEventListener('dragover', (e) => { if (e.dataTransfer?.types.includes('Files')) attachEl.classList.add('file-over'); });
  attachEl.addEventListener('dragleave', (e) => { if (!attachEl.contains(e.relatedTarget)) attachEl.classList.remove('file-over'); });

  function renderOutline() {
    const items = (list, depth) => list.map((it, i) => `
      <div class="ol-item" style="--d:${depth}">
        <button class="ol-link" data-path="${depth}-${i}" ${it.index === null && !it.url ? 'disabled' : ''}>${it.children.length ? icon('chevron', 12) : '<span class="ol-dot"></span>'}<span>${esc(it.title)}</span></button>
        ${it.children.length ? `<div class="ol-kids">${items(it.children, depth + 1)}</div>` : ''}
      </div>`).join('');
    outlineEl.innerHTML = outline.length
      ? `<div class="ol">${items(outline, 0)}</div>`
      : `<div class="empty small">${icon('bookmark', 22)}<br>No bookmarks yet.<br><button class="chip" data-add-bm>${icon('plus', 13)}Add bookmark</button></div>`;
    // map buttons back to items in document order
    const flat = [];
    const walk = (list) => list.forEach((it) => { flat.push(it); walk(it.children); });
    walk(outline);
    outlineEl.querySelectorAll('.ol-link').forEach((b, i) => { b._item = flat[i]; });
  }
  outlineEl.addEventListener('click', (e) => {
    if (e.target.closest('[data-add-bm]')) { addBookmark(); return; }
    const b = e.target.closest('.ol-link');
    if (!b?._item) return;
    if (b._item.children.length && e.target.closest('svg')) { b.parentElement.classList.toggle('open'); return; }
    if (b._item.url) window.open(b._item.url, '_blank');
    else if (b._item.index !== null) goToSourcePage(b._item.src, b._item.index);
  });

  // ---------------------------------------------------------------- attached files
  try {
    // pdf.js lists the files by name and hands over each one's bytes on request
    const doc = sources[main.id].doc;
    const found = await doc.getAttachments();
    const list = found instanceof Map ? [...found] : Object.entries(found || {});
    for (const [id, a] of list) {
      const data = a.content || await doc.getAttachmentContent?.(id).catch(() => null);
      if (data?.length) attachments.push({ name: a.filename || id || 'attachment', data, description: a.description || '' });
    }
  } catch (err) { console.warn('attachments', err); }
  const sizeText = (n) => (n < 1024 ? `${n} bytes` : n < 1048576 ? `${Math.round(n / 1024)} KB` : `${(n / 1048576).toFixed(1)} MB`);
  const OPENABLE = /\.(pdf|docx|doc|rtf|odt|xlsx|xlsm|xls|ods|csv)$/i;
  function renderAttachments() {
    attachEl.innerHTML = `
      <div class="att-head"><button class="chip" data-att-add>${icon('plus', 13)}Attach file</button></div>
      ${attachments.length ? attachments.map((a, i) => `
        <div class="att-row" data-i="${i}">
          <div class="att-ic">${icon('attach', 16)}</div>
          <div class="att-text"><strong title="${esc(a.name)}">${esc(a.name)}</strong><span>${sizeText(a.data.length)}${a.description ? ` · ${esc(a.description)}` : ''}</span></div>
          <div class="att-ops">
            ${OPENABLE.test(a.name) ? `<button data-att="open" title="Open in Kertas">${icon('open', 14)}</button>` : ''}
            <button data-att="save" title="Save a copy">${icon('save', 14)}</button>
            <button data-att="del" title="Remove from the PDF">${icon('trash', 14)}</button>
          </div>
        </div>`).join('') : `<div class="empty small">${icon('attach', 22)}<br>No attached files.<br>Drop files here, or use Attach file, to send spreadsheets, documents or pictures inside this PDF.</div>`}`;
  }
  attachEl.addEventListener('click', async (e) => {
    if (e.target.closest('[data-att-add]')) { attachFiles(); return; }
    const op = e.target.closest('[data-att]')?.dataset.att;
    const i = Number(e.target.closest('.att-row')?.dataset.i);
    const a = attachments[i];
    if (!op || !a) return;
    if (op === 'open') openDocument?.({ name: a.name, data: a.data.slice() });
    if (op === 'save') { const r = await saveAnyFile({ data: a.data, name: a.name }).catch((err) => { toast(String(err.message || err), true); return null; }); if (r) toast(`Saved ${r.name}`); }
    if (op === 'del') { attachments.splice(i, 1); changed(); renderAttachments(); toast(`${a.name} will be removed when you save`); }
  });
  async function attachFiles() { await addAttachments(await openFiles('any', true)); }
  async function addAttachments(files) {
    if (!files?.length) return;
    const big = files.find((f) => f.data.length > 100 * 1048576);
    if (big) { toast(`${big.name} is over 100 MB; attach a smaller file`, true); return; }
    for (const f of files) {
      const same = attachments.findIndex((a) => a.name === f.name);
      if (same >= 0) attachments[same] = { ...attachments[same], data: f.data };
      else attachments.push({ name: f.name, data: f.data, description: '' });
    }
    changed();
    sideEl.querySelector('[data-side="attach"]').click();
    toast(files.length === 1 ? `${files[0].name} attached` : `${files.length} files attached`);
  }

  async function addBookmark() {
    const p = pages[currentPage];
    if (!p) return;
    const title = await bookmarkDialog(`Page ${currentPage + 1}`);
    if (!title) return;
    outline.push({ title, src: p.src, index: p.index, url: null, children: [] });
    outlineTouched = true;
    changed();
    renderOutline();
    sideEl.querySelector('[data-side="outline"]').click();
    toast('Bookmark added');
  }

  // ---------------------------------------------------------------- page operations
  async function pageOp(op, i) {
    if (i < 0 || i >= pages.length) return;
    let focus = i;
    if (op === 'del') {
      if (pages.length === 1) { toast('A PDF needs at least one page'); return; }
      pages.splice(i, 1);
      focus = Math.min(i, pages.length - 1);
    }
    if (op === 'up' && i > 0) { [pages[i - 1], pages[i]] = [pages[i], pages[i - 1]]; focus = i - 1; }
    if (op === 'down' && i < pages.length - 1) { [pages[i + 1], pages[i]] = [pages[i], pages[i + 1]]; focus = i + 1; }
    if (op === 'rotl' || op === 'rotr') pages[i].rotate = (pages[i].rotate + (op === 'rotr' ? 90 : 270)) % 360;
    if (op === 'dup') {
      const src = pages[i];
      pages.splice(i + 1, 0, { ...src, annots: src.annots.map((a) => ({ ...a })), layerMoves: src.layerMoves ? structuredClone(src.layerMoves) : undefined, key: newKey(src.src, src.index) });
      focus = i + 1;
    }
    undoStack.length = 0; // page edits are not undoable in this version
    changed();
    pickedPage = focus;
    await layout();
    await renderThumbs();
    pagesEl.querySelectorAll('.pdf-page')[focus]?.scrollIntoView({ block: 'start' });
    updatePageInfo();
  }

  async function insertBlank() {
    const ref = pages[currentPage] || pages[0];
    const vp = viewportFor(await pdfPageFor(ref), ref, 1);
    const { id } = await blankSource(vp.width, vp.height);
    pages.splice(currentPage + 1, 0, { src: id, index: 0, rotate: 0, annots: [], key: newKey(id, 0) });
    changed();
    await layout();
    await renderThumbs();
    pagesEl.querySelectorAll('.pdf-page')[currentPage + 1]?.scrollIntoView({ block: 'start' });
    toast('Blank page inserted');
  }

  /** Adds the pages of PDFs (and pictures, one page each) at position at, by default the end */
  async function insertFiles(files, at = pages.length) {
    if (!files?.length) return;
    const isPdf = (f) => /\.pdf$/i.test(f.name);
    const isPicture = (f) => /\.(png|jpe?g)$/i.test(f.name);
    const skipped = files.filter((f) => !isPdf(f) && !isPicture(f));
    let added = 0;
    let used = 0;
    const first = at;
    for (const f of files) {
      if (!isPdf(f) && !isPicture(f)) continue;
      try {
        const data = isPdf(f) ? f.data : await imagesToPdf([{ bytes: f.data, fmt: /\.png$/i.test(f.name) ? 'png' : 'jpg' }]);
        const { id, doc } = await addSource(data, f.name);
        await objects.loadSource(id);
        const added1 = [];
        for (let i = 0; i < doc.numPages; i++) added1.push({ src: id, index: i, rotate: 0, annots: [], key: newKey(id, i) });
        pages.splice(at, 0, ...added1);
        at += added1.length;
        added += added1.length;
        used++;
      } catch (err) { toast(`Could not read ${f.name}: ${err.message}`, true); }
    }
    if (skipped.length) toast(`Only PDFs and pictures (PNG, JPG) can be added as pages; ${skipped.map((f) => f.name).join(', ')} left out`, true);
    if (added) {
      undoStack.length = 0;
      toast(`Added ${added} page${added === 1 ? '' : 's'} from ${used} file${used === 1 ? '' : 's'}`);
      changed(); pickedPage = first; await layout(); await renderThumbs();
      pagesEl.querySelectorAll('.pdf-page')[first]?.scrollIntoView({ block: 'start' });
      updatePageInfo();
    }
  }
  async function mergeFiles() { await insertFiles(await openFiles('pdf', true)); }

  // ---------------------------------------------------------------- OCR
  // Recognised words per source page, '<src>:<index>' -> { items (pdf-layout format), words (PDF space) }
  const ocrText = new Map();
  async function pageHasText(p) {
    const page = await sources[p.src].doc.getPage(p.index + 1);
    const tc = await page.getTextContent();
    return tc.items.some((it) => it.str && it.str.trim());
  }
  async function runOcr() {
    const list = pages.filter((p) => p.kind !== 'blank' && sources[p.src]);
    const res = await openDialog({
      title: 'Recognise text (OCR)',
      description: 'Reads the text in scanned pages on this computer. Afterwards the PDF can be searched, its text copied, and it converts to Word or Excel. Saving keeps the recognised text, hidden under the page image.',
      width: 500,
      fields: [
        { id: 'pages', label: 'Pages', type: 'segmented', value: 'scanned', options: [{ value: 'scanned', label: 'Pages without text' }, { value: 'all', label: 'All pages' }, { value: 'current', label: 'This page' }] },
        { id: 'info', type: 'info', html: '<b>Languages in the document</b>' },
        ...OCR_LANGUAGES.map((l) => ({ id: `lang-${l.id}`, type: 'checkbox', label: l.label, value: l.id === 'eng' })),
      ],
      validate: (v) => (OCR_LANGUAGES.some((l) => v[`lang-${l.id}`]) ? null : 'Pick at least one language'),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Recognise', primary: true, icon: 'scan-text' }],
    });
    if (!res) return;
    const langs = OCR_LANGUAGES.filter((l) => res.values[`lang-${l.id}`]).map((l) => l.id);
    let targets = res.values.pages === 'current' ? [pages[currentPage]].filter(Boolean) : list;
    if (res.values.pages === 'scanned') {
      const keep = [];
      for (const p of targets) if (!(await pageHasText(p))) keep.push(p);
      targets = keep;
      if (!targets.length) { toast('Every page already has text. Choose All pages to read them again.'); return; }
    }
    let done = 0;
    let words = 0;
    try {
      for (const p of targets) {
        const label = `Recognising page ${pages.indexOf(p) + 1} (${done + 1} of ${targets.length})`;
        setStatus(`${label}...`);
        const pdfPage = await sources[p.src].doc.getPage(p.index + 1);
        const SCALE = 300 / 72;
        const vp = pdfPage.getViewport({ scale: SCALE, rotation: pdfPage.rotate });
        const canvas = document.createElement('canvas');
        canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
        await pdfPage.render({ canvas, viewport: vp }).promise;
        const found = await recognise(canvas, langs, (f) => setStatus(`${label}: ${Math.round(f * 100)}%`));
        const items = found.map((w) => ({ str: w.str, x: w.x0 / SCALE, y: w.base / SCALE, w: (w.x1 - w.x0) / SCALE, size: Math.max(4, ((w.y1 - w.y0) / SCALE) * 0.95), bold: false, italic: false, font: null }))
          // Words on the same line become one line of text, with spaces between them
          .flatMap((it, i, arr) => (i && Math.abs(arr[i - 1].y - it.y) < it.size * 0.4 ? [{ ...it, str: ` ${it.str}` }] : [it]));
        const pdfWords = found.map((w) => {
          const [ax, ay] = vp.convertToPdfPoint(w.x0, w.base);
          const [bx, by] = vp.convertToPdfPoint(w.x1, w.base);
          return { str: w.str, x: ax, y: ay, w: Math.hypot(bx - ax, by - ay), size: Math.max(4, ((w.y1 - w.y0) / SCALE) * 0.9), rot: (Math.atan2(by - ay, bx - ax) * 180) / Math.PI };
        });
        ocrText.set(`${p.src}:${p.index}`, { items, words: pdfWords });
        words += found.length;
        done += 1;
      }
      changed();
      setStatus('');
      toast(`Recognised ${words.toLocaleString()} words on ${done} page${done === 1 ? '' : 's'}. Save to keep the text, or use Convert to Word or Excel.`);
    } catch (err) {
      console.error(err);
      setStatus('');
      toast(`Text recognition stopped: ${err.message || err}`, true);
    }
  }

  // ---------------------------------------------------------------- convert
  /** Positioned text for the pages as they are now (order, deletions, OCR text included) */
  async function currentText(onProgress) {
    const out = [];
    let n = 0;
    for (const p of pages) {
      n += 1;
      if (p.kind === 'blank' || !sources[p.src]) continue;
      const [pg] = await extractTextPages(sources[p.src].doc, { Util: pdfjsLib.Util, pages: [p.index] });
      if (!pg) continue;
      // Words recognised by OCR count as text too
      const ocr = ocrText.get(`${p.src}:${p.index}`);
      if (ocr && !pg.items.some((it) => it.str.trim())) pg.items = ocr.items;
      out.push({ ...pg, index: n - 1 });
      onProgress?.(n, pages.length);
    }
    return out;
  }

  async function convertTo(kind) {
    const base = name.replace(/\.pdf$/i, '');
    try {
      setStatus('Reading the PDF...');
      const text = await currentText((n, t) => setStatus(`Reading page ${n} of ${t}...`));
      const blankPages = text.filter((p) => !p.items.some((it) => it.str.trim())).length;
      if (blankPages && blankPages === text.length) {
        toast('This PDF has no text layer, so it looks like a scan. Run Recognise text first, then convert.', true);
        setStatus('');
        return;
      }
      if (kind === 'word') {
        setStatus('Building the Word document...');
        const bytes = await pdfToDocx(text, { title: base });
        await openDocument('writer', { name: `${base}.docx`, data: bytes });
        toast(blankPages ? `Converted to Word. ${blankPages} scanned page${blankPages === 1 ? ' has' : 's have'} no text yet; use Recognise text for those.` : 'Converted to Word. Check the layout, then save.');
      } else {
        setStatus('Building the workbook...');
        const bytes = await snapshotToXlsx(pdfToSnapshot(text, { name: base }));
        await openDocument('sheets', { name: `${base}.xlsx`, data: bytes });
        toast('Converted to Excel: tables keep their columns and numbers stay numbers.');
      }
      setStatus('');
    } catch (err) {
      console.error(err);
      toast(`Could not convert: ${err.message || err}`, true);
    }
  }

  async function pagesToImages(fmt, dpi) {
    const files = [];
    const base = name.replace(/\.pdf$/i, '');
    for (let i = 0; i < pages.length; i++) {
      const p = pages[i];
      if (!sources[p.src]) continue;
      setStatus(`Rendering page ${i + 1} of ${pages.length}...`);
      const pdfPage = await sources[p.src].doc.getPage(p.index + 1);
      const vp = pdfPage.getViewport({ scale: dpi / 72, rotation: (pdfPage.rotate + p.rotate) % 360 });
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      await pdfPage.render({ canvas, viewport: vp }).promise;
      const blob = await new Promise((r) => canvas.toBlob(r, fmt === 'jpg' ? 'image/jpeg' : 'image/png', 0.88));
      files.push({ name: `${base} - page ${i + 1}.${fmt}`, data: new Uint8Array(await blob.arrayBuffer()) });
    }
    setStatus('');
    const out = await saveMany(files);
    if (out) toast(`Saved ${files.length} image${files.length === 1 ? '' : 's'}`);
  }

  async function extractPages(mode) {
    const res = await splitDialog(pages.length, currentPage, mode);
    if (!res) return;
    const base = name.replace(/\.pdf$/i, '');
    setStatus('Preparing files...');
    try {
      const files = [];
      for (const [gi, group] of res.groups.entries()) {
        const subset = group.map((i) => pages[i]);
        const label = group.length === 1 ? `p${group[0] + 1}` : `p${group[0] + 1}-${group[group.length - 1] + 1}`;
        files.push({ name: res.groups.length > 1 || mode === 'extract' ? `${base} ${label}.pdf` : `${base} part ${gi + 1}.pdf`, data: await buildOutput(subset) });
      }
      if (files.length === 1) {
        const out = await saveFile({ data: files[0].data, name: files[0].name, kind: 'pdf', saveAs: true });
        if (out) toast(`Saved ${out.name}`);
      } else {
        const out = await saveMany(files);
        if (out) toast(`Saved ${out.count} files${out.folder ? ` to ${out.folder}` : ''}`);
      }
    } catch (err) {
      toast(`Could not create the files: ${err.message}`, true);
    }
    setStatus(`${pages.length} pages`);
  }

  // ---------------------------------------------------------------- document settings
  // ---------------------------------------------------------------- certificate signatures
  const escapeHtml = (t) => String(t ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  async function digitalSign() {
    if (security && (security.userPassword || security.ownerPassword)) {
      toast('A password-protected PDF cannot be signed here. Remove the password first, sign, then protect the signed copy in your PDF reader.');
      return;
    }
    const [cert] = await openFiles('cert', false);
    if (!cert) return;
    let certificate = null;
    const last = (() => { try { return JSON.parse(localStorage.getItem('kertas-digisign') || '{}'); } catch { return {}; } })();
    const res = await openDialog({
      title: 'Digitally sign',
      description: `Signs with ${cert.name}. A signed copy is saved; any later change to it shows as an invalid signature in PDF readers.`,
      width: 500,
      fields: [
        { id: 'password', label: 'Certificate password', type: 'password', value: '', autocomplete: 'off' },
        { id: 'reason', label: 'Reason', value: last.reason ?? 'I approve this document', placeholder: 'Optional' },
        { id: 'location', label: 'Location', value: last.location ?? '', placeholder: 'Optional' },
        { id: 'place', label: 'Show the signature', type: 'segmented', value: last.place || 'last', options: [{ value: 'last', label: 'Last page' }, { value: 'current', label: 'This page' }, { value: 'none', label: 'Hidden' }] },
        { id: 'who', type: 'info', html: '' },
      ],
      validate: (v) => {
        try { certificate = readCertificate(cert.data, v.password); return null; } catch (e) { return e.message; }
      },
      onChange: (v, form) => {
        const info = form.querySelector('[data-row="who"] .dlg-info');
        try {
          const c = readCertificate(cert.data, v.password);
          info.innerHTML = `${icon('verified', 14)} <b>${escapeHtml(c.name)}</b>${c.organisation ? `, ${escapeHtml(c.organisation)}` : ''}<br><span class="muted">Issued by ${escapeHtml(c.issuer || 'unknown')}${c.selfSigned ? ' (self-signed: readers will show it as not trusted)' : ''}. Valid until ${c.validTo.toLocaleDateString()}.</span>`;
        } catch { info.textContent = ''; }
      },
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Sign and save', primary: true }],
    });
    if (!res || !certificate) return;
    const v = res.values;
    try { localStorage.setItem('kertas-digisign', JSON.stringify({ reason: v.reason, location: v.location, place: v.place })); } catch { /* private mode */ }
    setStatus('Signing...');
    try {
      const list = pages;
      const bytes = await buildOutput(list);
      let visible = null;
      if (v.place !== 'none') {
        const idx = v.place === 'current' ? Math.max(0, Math.min(list.length - 1, currentPage)) : list.length - 1;
        const doc = await PDFDocument.load(bytes, { updateMetadata: false });
        const pg = doc.getPage(idx);
        const { width, height } = pg.getSize();
        const box = pg.getCropBox?.() || { x: 0, y: 0, width, height };
        const w = Math.min(220, box.width * 0.42); const h = 58;
        visible = { page: idx, x: box.x + box.width - w - 36, y: box.y + 36, width: w, height: h };
      }
      const signed = await signPdf(bytes, { certificate, reason: v.reason.trim(), location: v.location.trim(), visible });
      const base = String(name || 'document.pdf').replace(/\.pdf$/i, '');
      const saved = await saveFile({ data: signed, name: `${base} (signed).pdf`, kind: 'pdf', saveAs: true });
      if (saved) toast(`Signed by ${certificate.name} and saved.`);
    } catch (e) {
      toast(e.message || 'Signing failed');
    } finally {
      setStatus('');
    }
  }

  async function protect() {
    const res = await protectDialog(security);
    if (!res) return;
    if (res === 'remove') { removeProtection(); return; }
    security = res;
    changed();
    updateSecurityBadge();
    toast('Protection will be applied when you save');
  }
  function removeProtection() {
    if (!security) { toast('This document is not protected'); return; }
    security = null;
    changed();
    updateSecurityBadge();
    toast('The password will be removed when you save');
  }
  function updateSecurityBadge() {
    ribbon.setLabel('secure', security ? 'Protected' : '');
    ribbon.get('secure')?.classList.toggle('badge-secure', !!security);
  }

  async function editWatermark() {
    const res = await watermarkDialog(watermark);
    if (!res) return;
    watermark = res === 'remove' ? null : res;
    changed();
    redrawAll();
  }
  async function editHeaderFooter(numbersOnly) {
    const preset = numbersOnly && !headerFooter ? { header: ['', '', ''], footer: ['', 'Page {n} of {total}', ''] } : null;
    const res = await headerFooterDialog(headerFooter, preset);
    if (!res) return;
    headerFooter = res === 'remove' ? null : res;
    changed();
    redrawAll();
  }

  // ---------------------------------------------------------------- find
  async function find(q) {
    findQuery = q.trim().toLowerCase();
    pagesEl.querySelectorAll('.search-hit').forEach((n) => n.remove());
    if (!findQuery) { setStatus(''); return; }
    let total = 0;
    let firstKey = null;
    for (const p of pages) {
      const tc = await (await pdfPageFor(p)).getTextContent();
      const n = tc.items.filter((it) => it.str?.toLowerCase().includes(findQuery)).length;
      if (n && !firstKey) firstKey = p.key;
      total += n;
      const entry = rendered.get(p.key);
      if (entry?.viewport) highlightMatches(p, entry, tc);
    }
    setStatus(total ? `${total} match${total === 1 ? '' : 'es'}` : 'No matches');
    if (firstKey) pagesEl.querySelector(`.pdf-page[data-key="${firstKey}"]`)?.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  async function highlightMatches(p, entry, tc) {
    if (!findQuery) return;
    tc ||= await (await pdfPageFor(p)).getTextContent();
    for (const it of tc.items) {
      if (!it.str?.toLowerCase().includes(findQuery)) continue;
      const [, , , , x, y] = it.transform;
      const h = it.height || Math.hypot(it.transform[2], it.transform[3]);
      const [x1, y1] = entry.viewport.convertToViewportPoint(x, y);
      const [x2, y2] = entry.viewport.convertToViewportPoint(x + it.width, y + h);
      const hit = document.createElement('div');
      hit.className = 'search-hit';
      hit.style.cssText = `left:${Math.min(x1, x2)}px;top:${Math.min(y1, y2)}px;width:${Math.abs(x2 - x1)}px;height:${Math.abs(y2 - y1)}px`;
      entry.el.appendChild(hit);
    }
  }

  // ---------------------------------------------------------------- page info
  // a page picked in the thumbnails stays current until the person scrolls the pages
  for (const ev of ['wheel', 'touchstart', 'keydown']) pagesEl.addEventListener(ev, () => { pickedPage = null; }, { passive: true });
  // working on a page makes it the current one, even one that cannot scroll up to the line
  pagesEl.addEventListener('pointerdown', (e) => {
    const div = e.target.closest?.('.pdf-page');
    const i = div ? [...pagesEl.querySelectorAll('.pdf-page')].indexOf(div) : -1;
    pickedPage = i >= 0 ? i : null;
    updatePageInfo();
  }, { passive: true });
  function updatePageInfo() {
    const top = pagesEl.scrollTop + pagesEl.clientHeight / 3;
    const divs = [...pagesEl.querySelectorAll('.pdf-page')];
    let current = 0;
    divs.forEach((d, i) => { if (d.offsetTop <= top) current = i; });
    if (pickedPage !== null && pickedPage < divs.length) current = pickedPage;
    ribbon.setLabel('pageinfo', `Page ${current + 1} / ${pages.length}`);
    if (current !== currentPage || pages.length !== updatePageInfo.last) setStatus(`${pages.length} page${pages.length === 1 ? '' : 's'}`);
    updatePageInfo.last = pages.length;
    if (current !== currentPage) { currentPage = current; objects.renderPanel(); }
    thumbsEl.querySelectorAll('.thumb').forEach((t, i) => t.classList.toggle('active', i === current));
  }
  pagesEl.addEventListener('scroll', () => requestAnimationFrame(updatePageInfo));

  // ---------------------------------------------------------------- saving
  // Renders a redacted page into a flat image, with the black boxes burnt in
  async function rasterize(p) {
    const pdfPage = await pdfPageFor(p);
    const vp = pdfPage.getViewport({ scale: 2.5, rotation: 0 });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(vp.width);
    canvas.height = Math.ceil(vp.height);
    await pdfPage.render({ canvas, viewport: vp }).promise;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#000';
    for (const a of p.annots) {
      if (a.type !== 'redact') continue;
      const [x1, y1] = vp.convertToViewportPoint(a.x1, a.y1);
      const [x2, y2] = vp.convertToViewportPoint(a.x2, a.y2);
      ctx.fillRect(Math.min(x1, x2) - 1, Math.min(y1, y2) - 1, Math.abs(x2 - x1) + 2, Math.abs(y2 - y1) + 2);
    }
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.92));
    return { bytes: new Uint8Array(await blob.arrayBuffer()), fmt: 'jpg' };
  }

  function outlineFor(subset) {
    if (!outlineTouched) {
      const unchanged = subset.length === sources[main.id].doc.numPages && subset.every((p, i) => p.src === main.id && p.index === i);
      if (unchanged) return undefined; // the original bookmarks are kept exactly as they are
    }
    const map = (list) => list.flatMap((it) => {
      const kids = map(it.children);
      const at = subset.findIndex((p) => p.src === it.src && p.index === it.index);
      if (at < 0) return kids; // page gone: keep its children
      return [{ title: it.title, page: at, children: kids }];
    });
    return outline.length || outlineTouched ? map(outline) : undefined;
  }

  async function buildOutput(subset, extra = {}) {
    const srcBytes = {};
    const passwords = {};
    for (const [id, s] of Object.entries(sources)) {
      srcBytes[id] = s.bytes;
      if (s.password) passwords[id] = s.password;
    }
    const rasters = {};
    for (const [i, p] of subset.entries()) if (p.annots.some((a) => a.type === 'redact')) rasters[i] = await rasterize(p);
    // Recognised text goes in as an invisible layer, so the saved PDF is searchable
    subset = subset.map((p) => {
      const ocr = ocrText.get(`${p.src}:${p.index}`);
      return ocr ? { ...p, annots: [...p.annots, { type: 'ocrText', words: ocr.words }] } : p;
    });
    return buildPdf(srcBytes, subset, {
      passwords, rasters,
      formValues, flattenForms,
      watermark, headerFooter: headerFooter ? { ...headerFooter, file: name } : null,
      outline: outlineFor(subset),
      attachments: attachments.map((a) => ({ name: a.name, data: a.data, description: a.description })),
      layers: objects.saveSpec(),
      encrypt: security && (security.userPassword || security.ownerPassword) ? security : null,
      ...extra,
    });
  }

  // ---------------------------------------------------------------- wiring
  async function zoomBy(f) {
    const ratio = pagesEl.scrollTop / Math.max(1, pagesEl.scrollHeight);
    scale = Math.max(0.25, Math.min(4, scale * f));
    await layout();
    pagesEl.scrollTop = ratio * pagesEl.scrollHeight;
  }
  async function fitAndLayout() { await fitWidth(); await layout(); }
  let findTimer;
  ribbon.get('find').addEventListener('input', (e) => { clearTimeout(findTimer); findTimer = setTimeout(() => find(e.target.value), 250); });
  pagesEl.addEventListener('wheel', (e) => { if (e.ctrlKey) { e.preventDefault(); zoomBy(e.deltaY < 0 ? 1.1 : 1 / 1.1); } }, { passive: false });
  const onKey = (e) => {
    if (!el.classList.contains('active') || e.target.closest?.('input,textarea,select')) return;
    if (objects.onKey(e)) return;
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'z') { e.preventDefault(); undo(); }
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); ribbon.get('find').focus(); }
    if (e.key === 'Escape') setTool('select');
  };
  window.addEventListener('keydown', onKey);

  // Initial view: wait until the tab is visible so we can measure its width
  await new Promise((r) => requestAnimationFrame(r));
  await fitWidth();
  await layout();
  renderThumbs();
  renderOutline();
  setTool('select');
  ribbon.setOn('form-hl', fieldShading);
  ribbon.setOn('v-thumbs', true);
  updateSecurityBadge();
  const bits = [`${pages.length} page${pages.length === 1 ? '' : 's'}`];
  if (forms[main.id]) bits.push(`${forms[main.id].length} form field${forms[main.id].length === 1 ? '' : 's'}`);
  if (outline.length) bits.push('bookmarks');
  if (attachments.length) bits.push(`${attachments.length} attached file${attachments.length === 1 ? '' : 's'}`);
  setStatus(bits.join(' · '));
  if (forms[main.id]) toast('This PDF has a form you can fill in');

  return {
    async save() { return buildOutput(pages); },
    /** Export and print use the same output as save */
    async exportPdf() { return buildOutput(pages); },
    appendFiles: (files) => insertFiles(files),
    startSignature() { setTool('sign'); },
    markSaved() {
      setDirty(false);
      // after the first save with a fresh password, treat it as this file's own protection
      if (security) security.inherited = false;
    },
    onShow() { updatePageInfo(); },
    dispose() {
      disposed = true;
      observer.disconnect();
      window.removeEventListener('keydown', onKey);
      ribbon.dispose();
      stopOcr();
      Object.values(sources).forEach((s) => s.doc.destroy());
    },
    _state: { pages, sources, formValues, get security() { return security; } },
  };
}

/** Builds a PDF from images, used by the "Images to PDF" command */
export async function createFromImages(images) {
  const size = await imagesDialog(images.length);
  if (!size) return null;
  return imagesToPdf(images.map((im) => ({ bytes: im.data, fmt: /\.png$/i.test(im.name) ? 'png' : 'jpg' })), { pageSize: size });
}
