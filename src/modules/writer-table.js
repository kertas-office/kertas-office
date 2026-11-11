// Table tools: the contextual Table tab (shown while the cursor is in a table).
// Row and column moves are Kertas's own; the rest uses the editor's table commands
// and the document API.

/** The table, row and cell around the cursor, with their positions */
export function tableContext(ed) {
  const $f = ed?.state?.selection?.$from;
  if (!$f) return null;
  let table = null; let row = null; let cell = null;
  for (let d = $f.depth; d > 0; d--) {
    const n = $f.node(d); const name = n.type.name;
    if (!cell && (name === 'tableCell' || name === 'tableHeader')) cell = { node: n, pos: $f.before(d), depth: d };
    else if (!row && name === 'tableRow') row = { node: n, pos: $f.before(d), depth: d };
    else if (name === 'table') { table = { node: n, pos: $f.before(d), depth: d }; break; }
  }
  if (!table || !row || !cell) return null;
  const rowIndex = $f.index(row.depth - 1);
  const cellIndex = $f.index(cell.depth - 1);
  return { table, row, cell, rowIndex, cellIndex };
}

const spans = (cell) => ({ col: cell.attrs?.colspan || 1, row: cell.attrs?.rowspan || 1, cont: cell.attrs?.vMerge === 'continue' });
function rowsAreSimple(table, a, b) {
  for (const i of [a, b]) {
    const r = table.child(i);
    let bad = false;
    r.forEach((c) => { const s = spans(c); if (s.row > 1 || s.cont) bad = true; });
    if (bad) return false;
  }
  return true;
}
function tableIsGrid(table) {
  let ok = true; let width = -1;
  table.forEach((r) => {
    let w = 0;
    r.forEach((c) => { const s = spans(c); if (s.col > 1 || s.row > 1 || s.cont) ok = false; w += 1; });
    if (width === -1) width = w; else if (w !== width) ok = false;
  });
  return ok;
}

/** Moves the cursor's row up (-1) or down (+1) */
export function moveRow(ed, dir) {
  const ctx = tableContext(ed);
  if (!ctx) throw new Error('Click in a table first');
  const { table, rowIndex } = ctx;
  const to = rowIndex + dir;
  if (to < 0 || to >= table.node.childCount) throw new Error(dir < 0 ? 'This is already the first row' : 'This is already the last row');
  if (!rowsAreSimple(table.node, rowIndex, to)) throw new Error('Rows with cells merged across them cannot be moved');
  const rows = []; table.node.forEach((r) => rows.push(r));
  [rows[rowIndex], rows[to]] = [rows[to], rows[rowIndex]];
  const offsetInRow = ed.state.selection.from - (ctx.row.pos + 1);
  const next = ed.state.schema.nodes.table.create(table.node.attrs, rows, table.node.marks);
  const tr = ed.state.tr.replaceWith(table.pos, table.pos + table.node.nodeSize, next);
  // Keep the cursor in the moved row
  let rowPos = table.pos + 1; for (let i = 0; i < to; i++) rowPos += rows[i].nodeSize;
  try { tr.setSelection(ed.state.selection.constructor.near(tr.doc.resolve(Math.min(rowPos + 1 + offsetInRow, tr.doc.content.size)))); } catch { /* cursor stays */ }
  ed.view.dispatch(tr.scrollIntoView());
  return { success: true };
}

/** Moves the cursor's column left (-1) or right (+1) */
export function moveColumn(ed, dir) {
  const ctx = tableContext(ed);
  if (!ctx) throw new Error('Click in a table first');
  const { table, cellIndex } = ctx;
  if (!tableIsGrid(table.node)) throw new Error('Columns in a table with merged cells cannot be moved');
  const to = cellIndex + dir;
  const width = table.node.child(0).childCount;
  if (to < 0 || to >= width) throw new Error(dir < 0 ? 'This is already the first column' : 'This is already the last column');
  const rows = [];
  table.node.forEach((r) => {
    const cells = []; r.forEach((c) => cells.push(c));
    [cells[cellIndex], cells[to]] = [cells[to], cells[cellIndex]];
    rows.push(r.copy(r.type.create(r.attrs, cells).content));
  });
  // Column widths in the grid travel with their column
  const attrs = { ...table.node.attrs };
  for (const key of ['grid', 'colWidths']) {
    if (Array.isArray(attrs[key]) && attrs[key].length === width) { const g = [...attrs[key]]; [g[cellIndex], g[to]] = [g[to], g[cellIndex]]; attrs[key] = g; }
  }
  const next = ed.state.schema.nodes.table.create(attrs, rows, table.node.marks);
  const tr = ed.state.tr.replaceWith(table.pos, table.pos + table.node.nodeSize, next);
  // Cursor into the moved cell
  let p = table.pos + 1;
  for (let i = 0; i < ctx.rowIndex; i++) p += rows[i].nodeSize;
  p += 1; for (let i = 0; i < to; i++) p += rows[ctx.rowIndex].child(i).nodeSize;
  try { tr.setSelection(ed.state.selection.constructor.near(tr.doc.resolve(p + 2))); } catch { /* fine */ }
  ed.view.dispatch(tr.scrollIntoView());
  return { success: true };
}

/** The document-API id of the cursor's table (its position among the document's tables) */
export function tableNodeId(ed, api) {
  const ctx = tableContext(ed);
  if (!ctx) return null;
  let ordinal = -1; let i = 0;
  ed.state.doc.descendants((n, pos) => {
    if (n.type.name !== 'table') return true;
    if (pos === ctx.table.pos) ordinal = i;
    i++;
    return true;
  });
  const direct = ctx.table.node.attrs?.sdBlockId || ctx.table.node.attrs?.paraId;
  try {
    const all = (api().blocks.list({}).blocks || []).filter((b) => b.nodeType === 'table');
    if (direct && all.some((b) => b.nodeId === direct)) return direct;
    return all[ordinal]?.nodeId || direct;
  } catch { return direct; }
}

/** Sorts the table's rows by the cursor's column. A header row stays on top (Word detects it the same way). */
export function sortRows(ed, direction, opts = {}) {
  const ctx = tableContext(ed);
  if (!ctx) throw new Error('Click in a table first');
  const { table } = ctx; const cellIndex = Number.isInteger(opts.column) ? opts.column : ctx.cellIndex;
  const rows = []; table.node.forEach((r) => rows.push(r));
  for (const r of rows) { let bad = false; r.forEach((c) => { const sp = spans(c); if (sp.row > 1 || sp.cont) bad = true; }); if (bad) throw new Error('Tables with cells merged down cannot be sorted'); }
  const text = (r) => (r.childCount > cellIndex ? r.child(cellIndex).textContent.replace(/\u200B/g, '').trim() : '');
  const num = (t) => Number(t.replace(/[,%\s]|RM|\$/gi, ''));
  const isNum = (t) => t !== '' && Number.isFinite(num(t));
  const first = rows[0];
  const bold = (() => { let b = false; first.descendants((n) => { if (n.marks?.some((m) => m.type.name === 'bold')) b = true; }); return b; })();
  const headerish = first.attrs?.repeatHeader || first.attrs?.tableRowProperties?.repeatHeader || first.child(0)?.type.name === 'tableHeader' || bold
    || (rows.slice(1).filter((r) => text(r) !== '').every((r) => isNum(text(r))) && !isNum(text(first)) && rows.length > 2)
    // A first row of labels (every cell filled, none a number) reads as a header, as most tables have one
    || (rows.length > 2 && (() => { let all = true; first.forEach((c) => { const t = c.textContent.replace(/\u200B/g, '').trim(); if (!t || isNum(t)) all = false; }); return all; })());
  const head = (opts.header ?? headerish) ? [first] : [];
  const body = rows.slice(head.length);
  const numeric = body.filter((r) => text(r) !== '').every((r) => isNum(text(r)));
  const sign = direction === 'descending' ? -1 : 1;
  const filled = body.filter((r) => text(r) !== ''); const empty = body.filter((r) => text(r) === '');
  filled.sort((a, b) => sign * (numeric ? num(text(a)) - num(text(b)) : text(a).localeCompare(text(b), undefined, { sensitivity: 'base', numeric: true })));
  const next = ed.state.schema.nodes.table.create(table.node.attrs, [...head, ...filled, ...empty], table.node.marks);
  ed.view.dispatch(ed.state.tr.replaceWith(table.pos, table.pos + table.node.nodeSize, next));
  return { success: true, header: !!head.length };
}
