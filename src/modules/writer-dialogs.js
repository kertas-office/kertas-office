// Dialogs for the Writer module
import { openDialog } from '../lib/dialog.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export async function pageNumberDialog() {
  const res = await openDialog({
    title: 'Page numbers',
    description: 'Adds a live page number field that Word updates automatically.',
    fields: [
      { id: 'kind', label: 'Position', type: 'segmented', value: 'footer', options: [{ value: 'header', label: 'Top of page' }, { value: 'footer', label: 'Bottom of page' }] },
      { id: 'style', label: 'Show as', type: 'segmented', value: 'plain', options: [{ value: 'plain', label: '1' }, { value: 'page', label: 'Page 1' }, { value: 'of', label: 'Page 1 of 3' }] },
      { id: 'alignment', label: 'Alignment', type: 'segmented', value: 'center', options: [{ value: 'left', label: 'Left' }, { value: 'center', label: 'Centre' }, { value: 'right', label: 'Right' }] },
      { id: 'format', label: 'Number format', type: 'select', value: 'decimal', options: [
        { value: 'decimal', label: '1, 2, 3' }, { value: 'lowerRoman', label: 'i, ii, iii' }, { value: 'upperRoman', label: 'I, II, III' },
        { value: 'lowerLetter', label: 'a, b, c' }, { value: 'upperLetter', label: 'A, B, C' }, { value: 'numberInDash', label: '- 1 -, - 2 -' },
      ] },
      { id: 'start', label: 'Start at', type: 'number', min: 0, value: 1 },
      { id: 'first', type: 'checkbox', label: 'Show the number on the first page', value: true },
    ],
    actions: [{ id: 'remove', label: 'Remove page numbers', danger: true }, { id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
  });
  if (!res) return null;
  return res.action === 'remove' ? 'remove' : res.values;
}

export async function marginsDialog(current = {}) {
  const res = await openDialog({
    title: 'Custom margins',
    description: 'In inches. Applies to the whole document.',
    fields: [
      { id: 'top', label: 'Top', type: 'number', value: current.top ?? 1 },
      { id: 'bottom', label: 'Bottom', type: 'number', value: current.bottom ?? 1 },
      { id: 'left', label: 'Left', type: 'number', value: current.left ?? 1 },
      { id: 'right', label: 'Right', type: 'number', value: current.right ?? 1 },
      { id: 'gutter', label: 'Gutter', type: 'number', value: current.gutter ?? 0 },
    ],
    validate: (v) => (['top', 'bottom', 'left', 'right'].some((k) => v[k] === null || v[k] < 0 || v[k] > 4) ? 'Margins must be between 0 and 4 inches' : null),
    width: 420,
  });
  return res?.values || null;
}

export async function paperSizeDialog() {
  const res = await openDialog({
    title: 'Custom paper size',
    description: 'In inches.',
    fields: [{ id: 'width', label: 'Width', type: 'number', value: 8.5 }, { id: 'height', label: 'Height', type: 'number', value: 11 }],
    validate: (v) => (!v.width || !v.height || v.width < 1 || v.height < 1 || v.width > 22 || v.height > 22 ? 'Use sizes between 1 and 22 inches' : null),
    width: 380,
  });
  return res?.values || null;
}

export async function paragraphDialog() {
  const res = await openDialog({
    title: 'Paragraph',
    description: 'Applies to the selected paragraphs.',
    width: 520,
    fields: [
      { id: 'align', label: 'Alignment', type: 'segmented', value: 'keep', options: [{ value: 'keep', label: 'Keep' }, { value: 'left', label: 'Left' }, { value: 'center', label: 'Centre' }, { value: 'right', label: 'Right' }, { value: 'justify', label: 'Justify' }] },
      { id: 'indentLeft', label: 'Indent left (cm)', type: 'number', value: 0 },
      { id: 'indentRight', label: 'Indent right (cm)', type: 'number', value: 0 },
      { id: 'special', label: 'First line', type: 'segmented', value: 'none', options: [{ value: 'none', label: 'None' }, { value: 'firstLine', label: 'First line' }, { value: 'hanging', label: 'Hanging' }] },
      { id: 'specialBy', label: 'By (cm)', type: 'number', value: 1.27, show: (v) => v.special !== 'none' },
      { id: 'before', label: 'Spacing before (pt)', type: 'number', value: 0 },
      { id: 'after', label: 'Spacing after (pt)', type: 'number', value: 8 },
      { id: 'line', label: 'Line spacing', type: 'select', value: '1.08', options: [{ value: '1', label: 'Single' }, { value: '1.08', label: '1.08 (Word default)' }, { value: '1.15', label: '1.15' }, { value: '1.5', label: '1.5 lines' }, { value: '2', label: 'Double' }] },
      { id: 'keepNext', type: 'checkbox', label: 'Keep with next', value: false },
      { id: 'keepLines', type: 'checkbox', label: 'Keep lines together', value: false },
      { id: 'breakBefore', type: 'checkbox', label: 'Page break before', value: false },
    ],
  });
  return res?.values || null;
}

export async function tableDialog() {
  const res = await openDialog({
    title: 'Insert table',
    fields: [
      { id: 'columns', label: 'Columns', type: 'number', min: 1, max: 20, value: 3 },
      { id: 'rows', label: 'Rows', type: 'number', min: 1, max: 100, value: 3 },
    ],
    validate: (v) => (!v.columns || !v.rows || v.columns < 1 || v.rows < 1 || v.columns > 20 || v.rows > 100 ? 'Use 1 to 20 columns and 1 to 100 rows' : null),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    width: 380,
  });
  return res?.values || null;
}

export async function linkDialog(text = '') {
  const res = await openDialog({
    title: 'Insert link',
    fields: [
      { id: 'text', label: 'Text to display', value: text, placeholder: 'Leave empty to use the selected text' },
      { id: 'href', label: 'Address', value: 'https://', placeholder: 'https://example.com or mailto:you@email.com' },
    ],
    validate: (v) => (/^(https?:\/\/[^\s]+\.[^\s]+|mailto:\S+@\S+)$/i.test(v.href.trim()) ? null : 'Enter a web address starting with https:// or a mailto: link'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert link', primary: true }],
  });
  return res?.values || null;
}

export async function noteTextDialog(kind) {
  const res = await openDialog({
    title: kind === 'endnote' ? 'Insert endnote' : 'Insert footnote',
    description: `A numbered reference is added at the cursor and the note goes at the ${kind === 'endnote' ? 'end of the document' : 'bottom of the page'}.`,
    fields: [{ id: 'text', label: 'Note text', type: 'textarea', rows: 3 }],
    validate: (v) => (v.text.trim() ? null : 'Write the note text'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
  });
  return res ? res.values.text.trim() : null;
}

export async function captionDialog() {
  const res = await openDialog({
    title: 'Insert caption',
    description: 'Numbered automatically, for the table or picture next to the cursor.',
    fields: [
      { id: 'label', label: 'Label', type: 'segmented', value: 'Figure', options: [{ value: 'Figure', label: 'Figure' }, { value: 'Table', label: 'Table' }, { value: 'Equation', label: 'Equation' }] },
      { id: 'position', label: 'Position', type: 'segmented', value: 'below', options: [{ value: 'above', label: 'Above' }, { value: 'below', label: 'Below' }] },
      { id: 'text', label: 'Caption text', placeholder: 'Sales by region' },
    ],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
  });
  return res?.values || null;
}

export async function dateTimeDialog() {
  const d = new Date();
  const fmts = [
    d.toLocaleDateString('en-GB'),
    d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }),
    d.toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
    d.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }),
    d.toISOString().slice(0, 10),
    d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' }),
    `${d.toLocaleDateString('en-GB')} ${d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' })}`,
  ];
  const res = await openDialog({
    title: 'Date and time',
    fields: [{ id: 'fmt', label: 'Format', type: 'select', value: fmts[1], options: fmts.map((f) => ({ value: f, label: f })) }],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    width: 400,
  });
  return res ? res.values.fmt : null;
}

export async function wordCountDialog(counts, pages) {
  const rows = [
    ['Pages', pages], ['Words', counts.words], ['Characters', counts.characters], ['Paragraphs', counts.paragraphs],
    ['Headings', counts.headings], ['Tables', counts.tables], ['Images', counts.images], ['Comments', counts.comments], ['Tracked changes', counts.trackedChanges],
  ];
  await openDialog({
    title: 'Word count',
    fields: [{ id: 'info', type: 'info', html: `<table class="stat-table">${rows.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${Number(v || 0).toLocaleString()}</td></tr>`).join('')}</table>` }],
    actions: [{ id: 'cancel', label: 'Close', primary: true }],
    width: 360,
  });
}

export async function watermarkTextDialog() {
  const res = await openDialog({
    title: 'Custom watermark',
    description: 'A real Word watermark, behind the text on every page.',
    fields: [
      { id: 'text', label: 'Text', value: 'CONFIDENTIAL' },
      { id: 'orientation', label: 'Layout', type: 'segmented', value: 'diagonal', options: [{ value: 'diagonal', label: 'Diagonal' }, { value: 'horizontal', label: 'Horizontal' }] },
      { id: 'color', label: 'Colour', type: 'color', value: '#c0c0c0' },
      { id: 'opacity', label: 'Opacity', type: 'range', min: 10, max: 100, value: 50, format: (n) => `${n}%` },
    ],
    validate: (v) => (v.text.trim() ? null : 'Enter the watermark text'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Apply', primary: true }],
  });
  return res?.values || null;
}

export async function restrictDialog() {
  const res = await openDialog({
    title: 'Restrict editing',
    description: 'Saved in the .docx, so Word enforces it too.',
    fields: [{ id: 'mode', label: 'Allow only', type: 'select', value: 'readOnly', options: [
      { value: 'readOnly', label: 'No changes (read only)' },
      { value: 'comments', label: 'Comments' },
      { value: 'trackedChanges', label: 'Tracked changes' },
      { value: 'forms', label: 'Filling in forms' },
    ] }],
    actions: [{ id: 'off', label: 'Stop protection', danger: true }, { id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Start enforcing', primary: true, icon: 'lock' }],
    width: 420,
  });
  if (!res) return null;
  return res.action === 'off' ? 'off' : res.values.mode;
}

export async function commentDialog(title = 'New comment', text = '') {
  const res = await openDialog({
    title,
    fields: [{ id: 'text', label: 'Comment', type: 'textarea', rows: 4, value: text, placeholder: 'Type your comment' }],
    validate: (v) => (v.text.trim() ? null : 'Write something first'),
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Post', primary: true }],
  });
  return res ? res.values.text.trim() : null;
}
