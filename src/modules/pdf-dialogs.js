// Dialogs for the PDF module
import { openDialog } from '../lib/dialog.js';
import { parseRanges } from '../lib/pdf-ops.js';

export const STAMPS = [
  { id: 'approved', label: 'APPROVED', color: '#16a34a' },
  { id: 'rejected', label: 'REJECTED', color: '#dc2626' },
  { id: 'paid', label: 'PAID', color: '#16a34a' },
  { id: 'received', label: 'RECEIVED', color: '#2563eb' },
  { id: 'draft', label: 'DRAFT', color: '#2563eb' },
  { id: 'confidential', label: 'CONFIDENTIAL', color: '#dc2626' },
  { id: 'for-review', label: 'FOR REVIEW', color: '#d97706' },
  { id: 'final', label: 'FINAL', color: '#7c3aed' },
  { id: 'copy', label: 'COPY', color: '#475569' },
];

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

export async function askPassword(name, wrong) {
  const res = await openDialog({
    title: 'This PDF is password protected',
    description: wrong ? 'That password did not work. Try again.' : `Enter the password to open <b>${esc(name)}</b>.`,
    fields: [{ id: 'pw', label: 'Password', type: 'password', autocomplete: 'current-password' }],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Open', primary: true, icon: 'lock' }],
    validate: (v) => (v.pw ? null : 'Enter the password'),
    width: 420,
  });
  return res ? res.values.pw : null;
}

/** Returns { userPassword, ownerPassword, permissions } | 'remove' | null */
export async function protectDialog(current) {
  const p = current?.permissions || {};
  const res = await openDialog({
    title: 'Protect with password',
    description: current
      ? 'This document is protected. Change the settings or remove the protection.'
      : 'Encrypted with AES-256. Opens in Acrobat, Chrome, Edge, Preview and every modern reader.',
    fields: [
      { id: 'requireOpen', type: 'checkbox', label: 'Require a password to open the document', value: current ? !!current.userPassword : true },
      { id: 'user', label: 'Open password', type: 'password', autocomplete: 'new-password', value: current?.userPassword || '', show: (v) => v.requireOpen },
      { id: 'user2', label: 'Confirm open password', type: 'password', autocomplete: 'new-password', value: current?.userPassword || '', show: (v) => v.requireOpen },
      { id: 'restrict', type: 'checkbox', label: 'Restrict printing, copying and editing', value: !!current?.restrict },
      { id: 'printing', type: 'checkbox', label: 'Allow printing', value: p.printing !== false, show: (v) => v.restrict },
      { id: 'copying', type: 'checkbox', label: 'Allow copying text and images', value: p.copying !== false, show: (v) => v.restrict },
      { id: 'editing', type: 'checkbox', label: 'Allow editing, comments and form filling', value: p.modifying !== false, show: (v) => v.restrict },
      { id: 'owner', label: 'Permissions password', type: 'password', autocomplete: 'new-password', value: current?.ownerPassword && current.ownerPassword !== current.userPassword ? current.ownerPassword : '', hint: 'Needed to change these restrictions later. Must differ from the open password.', show: (v) => v.restrict },
    ],
    actions: [
      ...(current ? [{ id: 'remove', label: 'Remove protection', danger: true }] : []),
      { id: 'cancel', label: 'Cancel' },
      { id: 'ok', label: 'Protect', primary: true, icon: 'lock' },
    ],
    validate: (v, action) => {
      if (action === 'remove') return null;
      if (!v.requireOpen && !v.restrict) return 'Choose at least one kind of protection';
      if (v.requireOpen) {
        if (v.user.length < 4) return 'Use at least 4 characters for the open password';
        if (v.user !== v.user2) return 'The open passwords do not match';
      }
      if (v.restrict) {
        if (v.owner.length < 4) return 'Use at least 4 characters for the permissions password';
        if (v.requireOpen && v.owner === v.user) return 'The permissions password must differ from the open password';
      }
      return null;
    },
  });
  if (!res) return null;
  if (res.action === 'remove') return 'remove';
  const v = res.values;
  const permissions = v.restrict
    ? { printing: v.printing ? 'highResolution' : false, copying: v.copying, modifying: v.editing, annotating: v.editing, fillingForms: v.editing, documentAssembly: v.editing, contentAccessibility: true }
    : undefined;
  return {
    userPassword: v.requireOpen ? v.user : '',
    ownerPassword: v.restrict ? v.owner : (v.user || undefined),
    permissions,
    restrict: v.restrict,
  };
}

/** Returns watermark settings | 'remove' | null */
export async function watermarkDialog(current) {
  const res = await openDialog({
    title: 'Watermark',
    description: 'Added to every page when you save or export.',
    fields: [
      { id: 'text', label: 'Text', value: current?.text || 'CONFIDENTIAL', placeholder: 'CONFIDENTIAL' },
      { id: 'angle', label: 'Direction', type: 'segmented', value: String(current?.angle ?? 45), options: [{ value: '45', label: 'Diagonal' }, { value: '0', label: 'Horizontal' }] },
      { id: 'color', label: 'Colour', type: 'color', value: current?.color || '#94a3b8' },
      { id: 'opacity', label: 'Opacity', type: 'range', min: 5, max: 80, value: Math.round((current?.opacity ?? 0.25) * 100), format: (n) => `${n}%` },
      { id: 'size', label: 'Size', type: 'range', min: 20, max: 140, value: current?.size || 72, format: (n) => `${n} pt` },
    ],
    actions: [
      ...(current ? [{ id: 'remove', label: 'Remove watermark', danger: true }] : []),
      { id: 'cancel', label: 'Cancel' },
      { id: 'ok', label: 'Apply', primary: true },
    ],
    validate: (v, a) => (a !== 'remove' && !v.text.trim() ? 'Enter the watermark text' : null),
  });
  if (!res) return null;
  if (res.action === 'remove') return 'remove';
  const v = res.values;
  return { text: v.text.trim(), angle: Number(v.angle), color: v.color, opacity: v.opacity / 100, size: v.size };
}

/** Returns header/footer settings | 'remove' | null */
export async function headerFooterDialog(current, preset) {
  const c = current || preset || { header: ['', '', ''], footer: ['', 'Page {n} of {total}', ''] };
  const res = await openDialog({
    title: 'Header, footer and page numbers',
    description: 'Use <code>{n}</code> page number, <code>{total}</code> page count, <code>{date}</code> today, <code>{file}</code> file name.',
    width: 560,
    fields: [
      { id: 'hl', label: 'Header left', value: c.header[0] },
      { id: 'hc', label: 'Header centre', value: c.header[1] },
      { id: 'hr', label: 'Header right', value: c.header[2] },
      { id: 'fl', label: 'Footer left', value: c.footer[0] },
      { id: 'fc', label: 'Footer centre', value: c.footer[1] },
      { id: 'fr', label: 'Footer right', value: c.footer[2] },
      { id: 'start', label: 'First page number', type: 'number', min: 0, value: c.start ?? 1 },
      { id: 'size', label: 'Font size', type: 'range', min: 6, max: 20, value: c.size || 9, format: (n) => `${n} pt` },
    ],
    actions: [
      ...(current ? [{ id: 'remove', label: 'Remove', danger: true }] : []),
      { id: 'cancel', label: 'Cancel' },
      { id: 'ok', label: 'Apply', primary: true },
    ],
    validate: (v, a) => (a !== 'remove' && ![v.hl, v.hc, v.hr, v.fl, v.fc, v.fr].some((s) => s.trim()) ? 'Fill in at least one box' : null),
  });
  if (!res) return null;
  if (res.action === 'remove') return 'remove';
  const v = res.values;
  return { header: [v.hl, v.hc, v.hr], footer: [v.fl, v.fc, v.fr], start: v.start ?? 1, size: v.size, margin: 28 };
}

/** Returns { groups: [[pageIndex...]], single: bool } | null */
export async function splitDialog(total, current, mode = 'extract') {
  const res = await openDialog({
    title: mode === 'extract' ? 'Extract pages' : 'Split document',
    description: `${total} page${total === 1 ? '' : 's'}. Your annotations and settings are included in the new files.`,
    fields: [
      { id: 'mode', type: 'segmented', value: mode === 'extract' ? 'ranges' : 'every', options: [
        { value: 'ranges', label: 'Pick pages' }, { value: 'every', label: 'Every N pages' }, { value: 'single', label: 'One file per page' },
      ] },
      { id: 'ranges', label: 'Pages', value: `${current + 1}`, placeholder: 'e.g. 1-3, 5, 8-', hint: 'Commas separate files only in "Separate files" mode below.', show: (v) => v.mode === 'ranges' },
      { id: 'separate', type: 'checkbox', label: 'Save each range as a separate file', value: false, show: (v) => v.mode === 'ranges' },
      { id: 'n', label: 'Pages per file', type: 'number', min: 1, max: total, value: Math.max(1, Math.ceil(total / 2)), show: (v) => v.mode === 'every' },
    ],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: mode === 'extract' ? 'Extract' : 'Split', primary: true, icon: 'split' }],
    validate: (v) => {
      if (v.mode === 'ranges') { try { parseRanges(v.ranges, total); } catch (e) { return e.message; } }
      if (v.mode === 'every' && (!v.n || v.n < 1)) return 'Enter how many pages go in each file';
      return null;
    },
  });
  if (!res) return null;
  const v = res.values;
  if (v.mode === 'single') return { groups: Array.from({ length: total }, (_, i) => [i]) };
  if (v.mode === 'every') {
    const n = Math.max(1, Math.floor(v.n));
    const groups = [];
    for (let i = 0; i < total; i += n) groups.push(Array.from({ length: Math.min(n, total - i) }, (_, k) => i + k));
    return { groups };
  }
  if (v.separate) return { groups: v.ranges.split(',').map((part) => parseRanges(part, total)).filter((g) => g.length) };
  return { groups: [parseRanges(v.ranges, total)] };
}

/** Returns { text } | 'delete' | null */
export async function noteDialog(text = '', existing = false) {
  const res = await openDialog({
    title: existing ? 'Comment' : 'Add a comment',
    description: 'Saved as a real PDF comment, visible in Acrobat and other readers.',
    fields: [{ id: 'text', label: 'Comment', type: 'textarea', value: text, rows: 5, placeholder: 'Type your comment' }],
    actions: [
      ...(existing ? [{ id: 'delete', label: 'Delete', danger: true }] : []),
      { id: 'cancel', label: 'Cancel' },
      { id: 'ok', label: existing ? 'Save' : 'Add comment', primary: true },
    ],
    validate: (v, a) => (a !== 'delete' && !v.text.trim() ? 'Write something first' : null),
  });
  if (!res) return null;
  if (res.action === 'delete') return 'delete';
  return { text: res.values.text.trim() };
}

export async function bookmarkDialog(title = '') {
  const res = await openDialog({
    title: 'Add bookmark',
    description: 'Links to the page you are viewing.',
    fields: [{ id: 'title', label: 'Title', value: title, placeholder: 'Chapter 1' }],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Add', primary: true }],
    validate: (v) => (v.title.trim() ? null : 'Enter a title'),
    width: 420,
  });
  return res ? res.values.title.trim() : null;
}

export async function imagesDialog(count) {
  const res = await openDialog({
    title: 'Images to PDF',
    description: `${count} image${count === 1 ? '' : 's'}, one per page, in the order you picked them.`,
    fields: [{ id: 'size', label: 'Page size', type: 'segmented', value: 'a4', options: [{ value: 'a4', label: 'A4' }, { value: 'letter', label: 'Letter' }, { value: 'fit', label: 'Same as image' }] }],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Create PDF', primary: true }],
    width: 420,
  });
  return res ? res.values.size : null;
}
