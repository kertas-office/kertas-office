// Table borders, row height and cell margins, the Font dialog, and the formatting bar that
// opens with the right-click menu (like Word's mini toolbar).
import { openDialog } from '../lib/dialog.js';
import { openMenu, closeOpenMenu, openPalette } from '../lib/ribbon.js';
import { icon } from '../lib/icons.js';
import { tableContext, tableNodeId } from './writer-table.js';

export const BORDER_TARGETS = [
  { id: 'bottom', label: 'Bottom border', icon: 'border-bottom' },
  { id: 'top', label: 'Top border', icon: 'border-top' },
  { id: 'left', label: 'Left border', icon: 'border' },
  { id: 'right', label: 'Right border', icon: 'border' },
  { id: 'none', label: 'No border', icon: 'eye-off' },
  { id: 'all', label: 'All borders', icon: 'grid' },
  { id: 'outside', label: 'Outside borders', icon: 'rect' },
  { id: 'inside', label: 'Inside borders', icon: 'table' },
  { id: 'insideH', label: 'Inside horizontal', icon: 'rows' },
  { id: 'insideV', label: 'Inside vertical', icon: 'columns' },
];
export const LINE_STYLES = [
  { id: 'single', label: 'Single' },
  { id: 'double', label: 'Double' },
  { id: 'dashed', label: 'Dashed' },
  { id: 'dotted', label: 'Dotted' },
  { id: 'dotDash', label: 'Dot and dash' },
  { id: 'triple', label: 'Triple' },
];
export const LINE_WEIGHTS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6];
const UNDERLINES = [
  { value: 'none', label: 'None' }, { value: 'single', label: 'Single' }, { value: 'double', label: 'Double' },
  { value: 'thick', label: 'Thick' }, { value: 'dotted', label: 'Dotted' }, { value: 'dash', label: 'Dashed' }, { value: 'wave', label: 'Wavy' },
];
const weightLabel = (w) => `${String(w).replace(/^0\./, '.')} pt`.replace(/^\./, '0.');
const hex = (c) => String(c || '#000000').replace('#', '').toUpperCase();

/**
 * deps: { editor, api, act, toast, fmt, setFontSize, growFont, fonts(), designMenu(anchor),
 *         paragraphSettings(), runProp(key), textSelection(), toSelectionTarget(s), rulerUnit(), onChange() }
 */
export function createFormatTools(deps) {
  const { editor, api, act, toast } = deps;

  // ---- the pen: line style, weight and colour used by the border buttons (remembered)
  const PEN_KEY = 'kertas-border-pen';
  let pen = { style: 'single', weight: 0.5, color: '000000' };
  try { pen = { ...pen, ...JSON.parse(localStorage.getItem(PEN_KEY) || '{}') }; } catch { /* defaults */ }
  const savePen = () => { try { localStorage.setItem(PEN_KEY, JSON.stringify(pen)); } catch { /* private mode */ } };
  const spec = () => ({ lineStyle: pen.style, lineWeightPt: pen.weight, color: hex(pen.color) });
  function setPen(patch) { pen = { ...pen, ...patch }; savePen(); }

  // ---- the selected cells, as the table's cell list (row and column on the table grid)
  function selectedCells() {
    const ed = editor();
    const sel = ed?.state?.selection;
    if (!sel?.$anchorCell || typeof sel.forEachCell !== 'function') return null;
    const ctx = tableContext(ed);
    const nodeId = tableNodeId(ed, api);
    if (!ctx || !nodeId) return null;
    // the table's cells in document order, matched to the API's list (same order)
    const order = [];
    ctx.table.node.descendants((n, pos) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') { order.push(ctx.table.pos + 1 + pos); return false; } return true; });
    const all = api().tables.getCells({ nodeId })?.cells || [];
    const picked = []; const positions = [];
    sel.forEachCell((_n, pos) => { const i = order.indexOf(pos); if (i >= 0 && all[i]) { picked.push(all[i]); positions.push(pos); } });
    if (!picked.length) return null;
    const r0 = Math.min(...picked.map((c) => c.rowIndex));
    const r1 = Math.max(...picked.map((c) => c.rowIndex + (c.rowspan || 1) - 1));
    const c0 = Math.min(...picked.map((c) => c.columnIndex));
    const c1 = Math.max(...picked.map((c) => c.columnIndex + (c.colspan || 1) - 1));
    return { nodeId, cells: picked, positions, r0, r1, c0, c1 };
  }
  /** The selected cells, or else the cell with the cursor */
  function targetCells() {
    const block = selectedCells();
    if (block) return block;
    const ed = editor();
    const ctx = tableContext(ed);
    const nodeId = ctx && tableNodeId(ed, api);
    if (!nodeId) return null;
    const order = [];
    ctx.table.node.descendants((n, pos) => { if (n.type.name === 'tableCell' || n.type.name === 'tableHeader') { order.push(ctx.table.pos + 1 + pos); return false; } return true; });
    const i = order.indexOf(ctx.cell.pos);
    const cell = (api().tables.getCells({ nodeId })?.cells || [])[i];
    return cell ? { nodeId, cells: [cell], positions: [ctx.cell.pos] } : null;
  }

  // ---- cell alignment: top, middle or bottom, and left, centre or right
  const ALIGN_V = ['top', 'middle', 'bottom'];
  const ALIGN_H = ['left', 'center', 'right'];
  const alignLabel = (v, h) => `${{ top: 'Top', middle: 'Middle', bottom: 'Bottom' }[v]} ${{ left: 'left', center: 'centre', right: 'right' }[h]}`;
  async function alignCells(v, h) {
    const t = targetCells();
    if (!t) { toast('Click in a table first', true); return; }
    const ed = editor();
    // the paragraphs in those cells, read before anything changes
    const paras = [];
    for (const pos of t.positions) {
      ed.state.doc.nodeAt(pos)?.descendants((n) => { if (!n.isTextblock) return true; const id = n.attrs?.paraId || n.attrs?.sdBlockId; if (id && !paras.includes(id)) paras.push(id); return false; });
    }
    await act('Cell alignment', async () => {
      let r = { success: true };
      for (const c of t.cells) { r = await api().tables.setCellProperties({ nodeId: c.nodeId, verticalAlign: v === 'middle' ? 'center' : v }); if (r?.success === false) return r; }
      for (const nodeId of paras) {
        r = await deps.onBlock((target) => api().format.paragraph.setAlignment({ target, alignment: h }), { kind: 'block', nodeType: 'paragraph', nodeId });
        if (r?.success === false) return r;
      }
      return r;
    });
    deps.onChange?.();
  }
  /** Which of the nine alignments the cell with the cursor has */
  function currentAlign() {
    const ctx = tableContext(editor());
    if (!ctx) return null;
    const va = ctx.cell.node.attrs?.verticalAlign;
    const v = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
    // the paragraph's own alignment (the toolbar's value can lag behind the cursor)
    const $f = editor().state.selection.$from;
    let ha = 'left';
    for (let d = $f.depth; d > 0; d--) { const n = $f.node(d); if (n.isTextblock) { ha = n.attrs?.paragraphProperties?.justification || 'left'; break; } }
    const h = ha === 'center' ? 'center' : ha === 'right' ? 'right' : 'left';
    return { v, h };
  }
  /** The nine alignments as a 3 by 3 picker, opened from the right-click menu */
  function alignPickerAt(x, y) {
    const keys = []; const html = []; const titles = [];
    for (const v of ALIGN_V) for (const h of ALIGN_H) { keys.push(`${v}|${h}`); html.push(icon(`cell-${v}-${h}`, 20)); titles.push(alignLabel(v, h)); }
    const m = openMenu(anchorAt(x, y), { grid: keys, html, titles, columns: 3, cell: 38, onPick: (k) => { closeOpenMenu(); const [v, h] = k.split('|'); alignCells(v, h); } });
    m.classList.add('cell-align-picker');
  }

  /** Which edges of one cell a target touches, inside the selected block of cells */
  function cellEdges(c, b, target) {
    const top = c.rowIndex === b.r0; const bottom = c.rowIndex + (c.rowspan || 1) - 1 === b.r1;
    const left = c.columnIndex === b.c0; const right = c.columnIndex + (c.colspan || 1) - 1 === b.c1;
    const e = new Set();
    const outside = () => { if (top) e.add('top'); if (bottom) e.add('bottom'); if (left) e.add('left'); if (right) e.add('right'); };
    const insideH = () => { if (!top) e.add('top'); if (!bottom) e.add('bottom'); };
    const insideV = () => { if (!left) e.add('left'); if (!right) e.add('right'); };
    switch (target) {
      case 'top': if (top) e.add('top'); break;
      case 'bottom': if (bottom) e.add('bottom'); break;
      case 'left': if (left) e.add('left'); break;
      case 'right': if (right) e.add('right'); break;
      case 'outside': outside(); break;
      case 'insideH': insideH(); break;
      case 'insideV': insideV(); break;
      case 'inside': insideH(); insideV(); break;
      case 'all': case 'none': ['top', 'bottom', 'left', 'right'].forEach((x) => e.add(x)); break;
      default: break;
    }
    return [...e];
  }

  /** Applies a border target with the pen to the selected cells, or to the whole table */
  async function applyBorder(target, opts = {}) {
    const label = BORDER_TARGETS.find((t) => t.id === target)?.label || 'Borders';
    const border = target === 'none' ? null : (opts.spec || spec());
    const block = opts.wholeTable ? null : selectedCells();
    await act(label, async () => {
      if (block) {
        let res = { success: true };
        for (const c of block.cells) {
          for (const edge of cellEdges(c, block, target)) {
            res = border
              ? await api().tables.setBorder({ nodeId: c.nodeId, edge, ...border })
              : await api().tables.setBorder({ nodeId: c.nodeId, edge, lineStyle: 'nil', lineWeightPt: 0.5, color: 'auto' });
            if (res?.success === false) return res;
          }
        }
        return res;
      }
      const nodeId = tableNodeId(editor(), api);
      if (!nodeId) throw new Error('Click in a table first');
      if (target === 'none') return api().tables.setBorders({ nodeId, mode: 'applyTo', applyTo: 'all', border: null });
      return api().tables.setBorders({ nodeId, mode: 'applyTo', applyTo: target, border });
    });
    deps.onChange?.();
  }

  function bordersMenu() {
    return [
      ...BORDER_TARGETS.map((t) => ({ label: t.label, icon: t.icon, action: () => applyBorder(t.id) })),
      'sep',
      { label: 'Borders and shading...', icon: 'border', hint: 'Line style, width, colour and shading', action: () => bordersDialog() },
    ];
  }
  function lineMenu() {
    return [
      { header: 'Line style' },
      ...LINE_STYLES.map((s) => ({ label: s.label, checked: pen.style === s.id, action: () => { setPen({ style: s.id }); toast(`Borders: ${s.label.toLowerCase()} line`); } })),
      { header: 'Width' },
      ...LINE_WEIGHTS.map((w) => ({ label: weightLabel(w), checked: pen.weight === w, action: () => { setPen({ weight: w }); toast(`Borders: ${weightLabel(w)}`); } })),
    ];
  }
  const penColor = () => `#${hex(pen.color)}`;
  function setPenColor(c) { setPen({ color: hex(c) }); }

  async function bordersDialog() {
    const ctx = tableContext(editor());
    if (!ctx) { toast('Click in a table first', true); return; }
    const block = selectedCells();
    const res = await openDialog({
      title: 'Borders and shading',
      width: 460,
      fields: [
        block
          ? { id: 'scope', label: 'Apply to', type: 'segmented', value: 'cells', options: [{ value: 'cells', label: 'Selected cells' }, { value: 'table', label: 'Whole table' }] }
          : { id: 'scopeInfo', type: 'info', label: 'Applies to the whole table. Select cells first to change only those.' },
        { id: 'target', label: 'Borders', type: 'select', value: 'all', options: BORDER_TARGETS.map((t) => ({ value: t.id, label: t.label })) },
        { id: 'style', label: 'Line style', type: 'select', value: pen.style, options: LINE_STYLES.map((s) => ({ value: s.id, label: s.label })), show: (v) => v.target !== 'none' },
        { id: 'weight', label: 'Width', type: 'select', value: String(pen.weight), options: LINE_WEIGHTS.map((w) => ({ value: String(w), label: weightLabel(w) })), show: (v) => v.target !== 'none' },
        { id: 'color', label: 'Colour', type: 'color', value: penColor(), show: (v) => v.target !== 'none' },
        { id: 'shade', type: 'checkbox', label: 'Shade the cells', value: false },
        { id: 'fill', label: 'Shading colour', type: 'color', value: '#DEEAF6', show: (v) => v.shade },
      ],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    if (v.target !== 'none') setPen({ style: v.style, weight: Number(v.weight), color: hex(v.color) });
    const wholeTable = !block || v.scope === 'table';
    await applyBorder(v.target, { wholeTable });
    if (v.shade) {
      await act('Shading', async () => {
        if (!wholeTable) {
          let r = { success: true };
          for (const c of block.cells) { r = await api().tables.setShading({ nodeId: c.nodeId, color: hex(v.fill) }); if (r?.success === false) return r; }
          return r;
        }
        const nodeId = tableNodeId(editor(), api);
        const cells = api().tables.getCells({ nodeId })?.cells || [];
        let r = { success: true };
        for (const c of cells) { r = await api().tables.setShading({ nodeId: c.nodeId, color: hex(v.fill) }); if (r?.success === false) return r; }
        return r;
      });
    }
  }

  // ---- row height and cell margins
  const unit = () => (deps.rulerUnit?.() === 'in' ? 'in' : 'cm');
  const PT = { cm: 72 / 2.54, in: 72 };
  const toUnit = (pt) => Math.round((pt / PT[unit()]) * 100) / 100;
  const fromUnit = (v) => Number(v) * PT[unit()];

  /** The rows the selection covers: { nodeId, rows: [index...], first: row node } */
  function selectedRows() {
    const ed = editor();
    const ctx = tableContext(ed);
    if (!ctx) return null;
    const nodeId = tableNodeId(ed, api);
    const block = selectedCells();
    const rows = block ? Array.from({ length: block.r1 - block.r0 + 1 }, (_x, i) => block.r0 + i) : [ctx.rowIndex];
    return { nodeId, rows, first: ctx.table.node.child(rows[0]) };
  }
  async function rowHeightDialog() {
    const sr = selectedRows();
    if (!sr?.nodeId) { toast('Click in a table first', true); return; }
    const h = sr.first?.attrs?.tableRowProperties?.rowHeight;
    const curPt = h?.value ? h.value / 20 : 0;
    const res = await openDialog({
      title: 'Row height',
      description: sr.rows.length > 1 ? `Applies to the ${sr.rows.length} selected rows.` : 'Applies to the row with the cursor. Select cells in several rows to change them together.',
      width: 420,
      fields: [
        { id: 'rule', label: 'Height is', type: 'segmented', value: h?.rule === 'exact' ? 'exact' : curPt ? 'atLeast' : 'auto', options: [{ value: 'auto', label: 'Automatic' }, { value: 'atLeast', label: 'At least' }, { value: 'exact', label: 'Exactly' }] },
        { id: 'height', label: `Height (${unit()})`, type: 'number', value: curPt ? toUnit(curPt) : (unit() === 'cm' ? 1 : 0.4), show: (v) => v.rule !== 'auto' },
      ],
      validate: (v) => (v.rule !== 'auto' && (!v.height || v.height <= 0 || fromUnit(v.height) > 1584) ? 'Enter a height above 0' : null),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    await act('Row height', async () => {
      let r = { success: true };
      for (const rowIndex of sr.rows) {
        r = await api().tables.setRowHeight({ nodeId: sr.nodeId, rowIndex, heightPt: v.rule === 'auto' ? 1 : Math.round(fromUnit(v.height) * 100) / 100, rule: v.rule });
        if (r?.success === false) return r;
      }
      return r;
    });
    deps.onChange?.();
  }
  async function cellMarginsDialog() {
    const nodeId = tableNodeId(editor(), api);
    if (!nodeId) { toast('Click in a table first', true); return; }
    let m = {};
    try { m = api().tables.getProperties({ nodeId })?.defaultCellMargins || {}; } catch { /* none */ }
    const u = unit();
    const res = await openDialog({
      title: 'Cell margins',
      description: `Space between the cell edges and the text, in ${u === 'cm' ? 'centimetres' : 'inches'}. Applies to every cell in the table.`,
      width: 420,
      fields: [
        { id: 'top', label: 'Top', type: 'number', value: toUnit(m.topPt ?? 0) },
        { id: 'bottom', label: 'Bottom', type: 'number', value: toUnit(m.bottomPt ?? 0) },
        { id: 'left', label: 'Left', type: 'number', value: toUnit(m.leftPt ?? 5.4) },
        { id: 'right', label: 'Right', type: 'number', value: toUnit(m.rightPt ?? 5.4) },
      ],
      validate: (v) => (['top', 'bottom', 'left', 'right'].some((k) => v[k] === null || v[k] < 0 || fromUnit(v[k]) > 288) ? 'Use values from 0 up to 4 inches (10 cm)' : null),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    const pt = (k) => Math.round(fromUnit(v[k]) * 100) / 100;
    const r = await act('Cell margins', () => api().tables.setTablePadding({ nodeId, topPt: pt('top'), rightPt: pt('right'), bottomPt: pt('bottom'), leftPt: pt('left') }));
    // Each cell also carries the margins it was opened with, and the page is drawn from those;
    // they follow the new table margins unless the cell set its own in the file
    if (r !== null && r !== undefined) {
      try {
        const ed = editor(); const ctx = tableContext(ed);
        if (ctx) {
          const px = (k) => Math.round(fromUnit(v[k]) * (96 / 72) * 100) / 100;
          const margins = { top: px('top'), bottom: px('bottom'), left: px('left'), right: px('right') };
          const tr = ed.state.tr;
          ctx.table.node.descendants((n, pos) => {
            if (n.type.name !== 'tableCell' && n.type.name !== 'tableHeader') return true;
            if (!(n.attrs.tableCellPropertiesInlineKeys || []).includes('cellMargins')) tr.setNodeMarkup(ctx.table.pos + 1 + pos, null, { ...n.attrs, cellMargins: margins });
            return false;
          });
          if (tr.docChanged) ed.view.dispatch(tr);
        }
      } catch { /* the table margins are still set */ }
    }
    deps.onChange?.();
  }

  // ---- Font dialog
  async function fontDialog() {
    let s;
    try { s = await deps.textSelection(); } catch { toast('Select some text first', true); return; }
    const rp = deps.runProp;
    const u = rp('underline');
    // font and size as the Home tab shows them (they may come from the paragraph style)
    const snap = deps.snapshot?.() || {};
    const snapFont = String(snap['font-family']?.value || '').split(',')[0].replace(/["']/g, '').trim();
    const snapSize = parseFloat(String(snap['font-size']?.value || '')) || (Number(rp('fontSize')) ? Number(rp('fontSize')) / 2 : '');
    const cur = {
      font: snapFont || '', size: snapSize || '', bold: !!rp('bold'), italic: !!rp('italic'),
      underline: !u ? 'none' : typeof u === 'object' ? (u.val || u.style || 'single') : 'single',
      color: typeof rp('color') === 'string' && /^#?[0-9a-f]{6}$/i.test(rp('color')) ? `#${hex(rp('color'))}` : '#000000',
      strike: !!rp('strike'), dstrike: !!rp('dstrike'), sup: rp('vertAlign') === 'superscript', sub: rp('vertAlign') === 'subscript',
      smallCaps: !!rp('smallCaps'), caps: !!rp('caps') || rp('textTransform') === 'uppercase', spacing: (Number(rp('letterSpacing')) || 0) / 20,
    };
    const fontOptions = [{ value: '', label: '(Keep as is)' }, ...deps.fonts().map((f) => ({ value: f, label: f }))];
    if (cur.font && !fontOptions.some((o) => o.value === cur.font)) fontOptions.push({ value: cur.font, label: cur.font });
    const res = await openDialog({
      title: 'Font',
      description: 'Applies to the selected text.',
      width: 520,
      fields: [
        { id: 'font', label: 'Font', type: 'select', value: cur.font, options: fontOptions },
        { id: 'size', label: 'Size (pt)', type: 'number', value: cur.size || null },
        { id: 'bold', type: 'checkbox', label: 'Bold', value: cur.bold },
        { id: 'italic', type: 'checkbox', label: 'Italic', value: cur.italic },
        { id: 'underline', label: 'Underline', type: 'select', value: UNDERLINES.some((o) => o.value === cur.underline) ? cur.underline : 'single', options: UNDERLINES },
        { id: 'color', label: 'Font colour', type: 'color', value: cur.color },
        { id: 'strike', type: 'checkbox', label: 'Strikethrough', value: cur.strike },
        { id: 'dstrike', type: 'checkbox', label: 'Double strikethrough', value: cur.dstrike },
        { id: 'sup', type: 'checkbox', label: 'Superscript', value: cur.sup },
        { id: 'sub', type: 'checkbox', label: 'Subscript', value: cur.sub },
        { id: 'smallCaps', type: 'checkbox', label: 'Small caps', value: cur.smallCaps },
        { id: 'caps', type: 'checkbox', label: 'All caps', value: cur.caps },
        { id: 'spacing', label: 'Character spacing (pt, below 0 is condensed)', type: 'number', value: cur.spacing },
      ],
      validate: (v) => (v.size !== null && (v.size < 1 || v.size > 1638) ? 'Enter a size between 1 and 1638' : v.sup && v.sub ? 'Choose superscript or subscript, not both' : null),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    // only what was changed, so mixed formatting in the selection stays as it was
    const inline = {};
    if (v.font && v.font !== cur.font) inline.fontFamily = v.font;
    if (v.size && v.size !== Number(cur.size)) inline.fontSize = v.size;
    for (const k of ['bold', 'italic', 'strike', 'dstrike', 'smallCaps', 'caps']) if (v[k] !== cur[k]) inline[k] = v[k];
    if (v.underline !== cur.underline) inline.underline = v.underline === 'none' ? false : v.underline === 'single' ? true : { style: v.underline };
    if (hex(v.color) !== hex(cur.color)) inline.color = `#${hex(v.color)}`;
    if (v.sup !== cur.sup || v.sub !== cur.sub) inline.vertAlign = v.sup ? 'superscript' : v.sub ? 'subscript' : null;
    if (Number(v.spacing || 0) !== cur.spacing) inline.letterSpacing = Number(v.spacing || 0) || null;
    if (!Object.keys(inline).length) return;
    await act('Font', () => api().format.apply({ target: deps.toSelectionTarget(s), inline }));
    deps.onChange?.();
  }

  // ---- the formatting bar shown with the right-click menu
  let bar = null; let barCleanup = null; let anchorEl = null;
  function hideBar() { barCleanup?.(); barCleanup = null; bar?.remove(); bar = null; closeOpenMenu(); }
  /** An invisible element at a screen point, for menus that open from the bar or the right-click menu */
  function anchorAt(x, y) {
    anchorEl?.remove();
    anchorEl = document.createElement('div');
    anchorEl.style.cssText = `position:fixed;left:${x}px;top:${y}px;width:1px;height:1px;pointer-events:none`;
    document.body.appendChild(anchorEl);
    return anchorEl;
  }
  const btn = (id, ic, title, extra = '') => `<button class="fbar-btn${extra}" data-f="${id}" title="${title}" aria-label="${title}">${icon(ic, 16)}</button>`;
  function showBar(menuEl) {
    hideBar();
    const snap = deps.snapshot?.() || {};
    const fontName = String(snap['font-family']?.value || deps.runProp('fontFamily') || 'Font').split(',')[0].replace(/["']/g, '');
    const size = String(snap['font-size']?.value || '').replace(/pt$/, '') || '';
    bar = document.createElement('div');
    bar.className = 'fbar';
    bar.innerHTML = `
      <button class="fbar-pick fbar-font" data-f="font" title="Font"><span>${fontName.replace(/[<>&]/g, '')}</span>${icon('chevron', 12)}</button>
      <button class="fbar-pick fbar-size" data-f="size" title="Font size"><span>${size.replace(/[<>&]/g, '') || '11'}</span>${icon('chevron', 12)}</button>
      ${btn('grow', 'grow', 'Increase font size')}${btn('shrink', 'shrink', 'Decrease font size')}
      <span class="fbar-sep"></span>
      ${btn('bold', 'bold', 'Bold', snap.bold?.active ? ' on' : '')}${btn('italic', 'italic', 'Italic', snap.italic?.active ? ' on' : '')}${btn('underline', 'underline', 'Underline', snap.underline?.active ? ' on' : '')}
      <button class="fbar-btn fbar-sw" data-f="highlight" title="Highlight colour">${icon('highlighter', 16)}<i style="background:${deps.lastColor?.('hl') || '#FFFF00'}"></i></button>
      <button class="fbar-btn fbar-sw" data-f="color" title="Font colour">${icon('font-color', 16)}<i style="background:${deps.lastColor?.('fc') || '#C00000'}"></i></button>
      <span class="fbar-sep"></span>
      ${btn('left', 'align-left', 'Align left')}${btn('center', 'align-center', 'Centre')}${btn('right', 'align-right', 'Align right')}${btn('justify', 'justify', 'Justify')}
      <span class="fbar-sep"></span>
      ${btn('ul', 'list-ul', 'Bullets')}${btn('ol', 'list-ol', 'Numbering')}${btn('outdent', 'indent-less', 'Decrease indent')}${btn('indent', 'indent-more', 'Increase indent')}`;
    document.body.appendChild(bar);
    place(menuEl);
    // keeps the selection: the right-click menu moves the cursor to any click outside it
    for (const t of ['pointerdown', 'mousedown']) bar.addEventListener(t, (e) => { e.preventDefault(); e.stopPropagation(); });
    bar.addEventListener('click', (e) => {
      const b = e.target.closest('[data-f]');
      if (!b) return;
      const r = b.getBoundingClientRect();
      const a = anchorAt(r.left, r.bottom - 1);
      const f = b.dataset.f;
      deps.closeContextMenu?.();
      switch (f) {
        case 'font': openMenu(a, deps.fonts().map((name) => ({ label: name, style: `font-family:'${name}'`, action: () => deps.fmt('font-family', name) }))); break;
        case 'size': openMenu(a, [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72].map((n) => ({ label: String(n), action: () => deps.setFontSize(n) }))); break;
        case 'grow': deps.growFont(1); break;
        case 'shrink': deps.growFont(-1); break;
        case 'bold': case 'italic': case 'underline': deps.fmt(f); b.classList.toggle('on'); break;
        case 'highlight': openPalette(a, { palette: 'highlight', value: deps.lastColor?.('hl') }, (c) => { deps.fmt('highlight-color', c); deps.setLastColor?.('hl', c); }); break;
        case 'color': openPalette(a, { palette: 'text', value: deps.lastColor?.('fc') }, (c) => { deps.fmt('text-color', c); deps.setLastColor?.('fc', c); }); break;
        case 'left': case 'center': case 'right': case 'justify': deps.fmt('text-align', f); break;
        case 'ul': deps.fmt('bullet-list'); break;
        case 'ol': deps.fmt('numbered-list'); break;
        case 'outdent': deps.fmt('indent-decrease'); break;
        case 'indent': deps.fmt('indent-increase'); break;
        default: break;
      }
    });
    // the bar goes away with a click elsewhere, a key, or a scroll
    const away = (e) => { if (bar && !bar.contains(e.target) && !e.target.closest?.('.menu')) hideBar(); };
    const key = (e) => { if (e.key === 'Escape' || (!e.ctrlKey && !e.metaKey && e.key.length === 1)) hideBar(); };
    const scroll = (e) => { if (!e.target?.closest?.('.menu, .context-menu')) hideBar(); };
    setTimeout(() => {
      document.addEventListener('mousedown', away, true);
      document.addEventListener('keydown', key, true);
      document.addEventListener('scroll', scroll, true);
    }, 0);
    barCleanup = () => {
      document.removeEventListener('mousedown', away, true);
      document.removeEventListener('keydown', key, true);
      document.removeEventListener('scroll', scroll, true);
    };
  }
  function place(menuEl) {
    if (!bar) return;
    const m = menuEl?.getBoundingClientRect();
    const w = bar.offsetWidth; const h = bar.offsetHeight;
    let x = m ? m.left : 100; let y = m ? m.top - h - 8 : 100;
    if (m && y < 8) y = m.bottom + 8; // no room above the menu: below it
    x = Math.max(8, Math.min(x, window.innerWidth - w - 8));
    y = Math.max(8, Math.min(y, window.innerHeight - h - 8));
    bar.style.left = `${x}px`; bar.style.top = `${y}px`;
  }

  /** The table design gallery, opened from the right-click menu at the click point */
  function designAt(x, y) { deps.designMenu(anchorAt(x, y)); }

  return {
    alignCells, currentAlign, alignPickerAt, alignLabel, ALIGN_V, ALIGN_H,
    applyBorder, bordersMenu, lineMenu, bordersDialog, penColor, setPenColor,
    rowHeightDialog, cellMarginsDialog, fontDialog, showBar, hideBar, designAt, selectedCells,
  };
}
