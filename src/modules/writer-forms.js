// Form fields in documents (Word content controls): text box, tick box, dropdown list and date
// picker. They are saved as Word's own controls, so Word and LibreOffice can fill them in too.
import { openDialog } from '../lib/dialog.js';
import { openMenu, closeOpenMenu } from '../lib/ribbon.js';

const DATE_FORMATS = [
  { value: 'd MMMM yyyy', label: '13 November 2026' },
  { value: 'dd/MM/yyyy', label: '13/11/2026' },
  { value: 'MM/dd/yyyy', label: '11/13/2026' },
  { value: 'yyyy-MM-dd', label: '2026-11-13' },
  { value: 'MMMM d, yyyy', label: 'November 13, 2026' },
  { value: 'dddd, d MMMM yyyy', label: 'Friday, 13 November 2026' },
];
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** A yyyy-mm-dd date written in a Word date pattern (d, dd, M, MM, MMM, MMMM, yy, yyyy, ddd, dddd) */
export function formatDate(iso, pattern) {
  const [y, m, d] = iso.split('-').map(Number);
  const wd = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  const pad = (n) => String(n).padStart(2, '0');
  return pattern.replace(/dddd|ddd|dd|d|MMMM|MMM|MM|M|yyyy|yy/g, (t) => ({
    dddd: DAYS[wd], ddd: DAYS[wd].slice(0, 3), dd: pad(d), d: String(d), MMMM: MONTHS[m - 1], MMM: MONTHS[m - 1].slice(0, 3), MM: pad(m), M: String(m), yyyy: String(y), yy: String(y).slice(2),
  }[t]));
}

/** deps: { editor, api, act, toast, selection, toSelectionTarget, failed } */
export function createForms(deps) {
  const { api, act, toast } = deps;
  const target = (nodeId) => ({ kind: 'inline', nodeType: 'sdt', nodeId: String(nodeId) });
  const info = (nodeId) => { try { return api().contentControls.get({ target: target(nodeId) }); } catch { return null; } };
  /** Runs a change on a field locked against deletion: Word still lets it be filled in, the editor does not */
  async function filling(nodeId, fn) {
    const t = target(nodeId);
    const locked = info(nodeId)?.lockMode === 'sdtLocked';
    if (locked) { const u = await api().contentControls.setLockMode({ target: t, lockMode: 'unlocked' }); if (deps.failed(u)) return u; }
    try { return await fn(t); } finally { if (locked) { try { await api().contentControls.setLockMode({ target: t, lockMode: 'sdtLocked' }); } catch { /* stays unlocked */ } } }
  }

  async function insertField(type) {
    const titles = { text: 'Text field', checkbox: 'Tick box', dropDownList: 'Dropdown list', date: 'Date picker' };
    const fields = [{ id: 'title', label: 'Title (shown when the field is selected)', value: titles[type] }];
    if (type === 'text') fields.push({ id: 'placeholder', label: 'Text shown before it is filled in', value: 'Click here to type' });
    if (type === 'dropDownList') fields.push({ id: 'choices', label: 'Choices, one per line', type: 'textarea', rows: 5, value: 'First choice\nSecond choice\nThird choice' });
    if (type === 'date') fields.push({ id: 'format', label: 'Date format', type: 'select', value: 'd MMMM yyyy', options: DATE_FORMATS });
    const res = await openDialog({
      title: `Insert ${titles[type].toLowerCase()}`,
      description: 'A form field that can be filled in here, in Word and in LibreOffice.',
      width: 440,
      fields,
      validate: (v) => (type === 'dropDownList' && !v.choices.split('\n').some((c) => c.trim()) ? 'Add at least one choice' : null),
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Insert', primary: true }],
    });
    if (!res) return;
    const v = res.values;
    const content = type === 'checkbox' ? '☐' : type === 'dropDownList' ? 'Choose an item' : type === 'date' ? 'Pick a date' : (v.placeholder || 'Click here to type');
    await act(titles[type], async () => {
      const s = await deps.selection();
      const r = await api().create.contentControl({ kind: 'inline', controlType: type, at: deps.toSelectionTarget(s), content, alias: v.title || titles[type] });
      if (deps.failed(r)) return r;
      const t = r.contentControl;
      if (type === 'dropDownList') {
        const items = v.choices.split('\n').map((c) => c.trim()).filter(Boolean).map((c) => ({ displayText: c, value: c }));
        return api().contentControls.choiceList.setItems({ target: t, items });
      }
      if (type === 'date') return api().contentControls.date.setDisplayFormat({ target: t, format: v.format });
      return r;
    });
  }

  // ---- filling in: a click on a dropdown, date or tick box
  async function pickChoice(nodeId, anchor) {
    const c = info(nodeId);
    const items = c?.properties?.items || [];
    openMenu(anchor, [
      ...items.map((it) => ({ label: it.displayText, checked: c.properties.selectedValue === it.value, action: () => act('Choose', () => filling(nodeId, (t) => api().contentControls.choiceList.setSelected({ target: t, value: it.value }))) })),
      ...(items.length ? ['sep'] : []),
      { label: 'Edit choices...', icon: 'edit', action: () => fieldSettings(nodeId) },
    ]);
  }
  async function pickDate(nodeId) {
    const c = info(nodeId);
    const fmt = c?.properties?.dateFormat || 'd MMMM yyyy';
    const today = new Date(); const iso = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    const res = await openDialog({
      title: c?.properties?.alias || 'Pick a date',
      width: 360,
      fields: [{ id: 'date', label: 'Date', type: 'date', value: iso }],
      validate: (v) => (/^\d{4}-\d{2}-\d{2}$/.test(v.date || '') ? null : 'Pick a date'),
      actions: [{ id: 'clear', label: 'Clear' }, { id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'OK', primary: true }],
    });
    if (!res) return;
    if (res.action === 'clear') { await act('Date', () => filling(nodeId, (t) => api().contentControls.replaceContent({ target: t, content: 'Pick a date' }))); return; }
    await act('Date', () => filling(nodeId, async (t) => {
      const r = await api().contentControls.date.setValue({ target: t, value: res.values.date });
      if (deps.failed(r)) return r;
      return api().contentControls.replaceContent({ target: t, content: formatDate(res.values.date, fmt) });
    }));
  }
  async function toggleBox(nodeId) {
    await act('Tick box', () => filling(nodeId, (t) => api().contentControls.checkbox.toggle({ target: t })));
  }
  /** Handles a click on a form field; returns true when it did something */
  function onClick(e) {
    const hit = document.elementsFromPoint(e.clientX, e.clientY).find((n) => n.dataset?.sdtId && n.classList?.contains('superdoc-structured-content-inline'));
    if (!hit) return false;
    const c = info(hit.dataset.sdtId);
    if (!c || c.lockMode === 'contentLocked' || c.lockMode === 'sdtContentLocked') return false;
    if (c.controlType === 'checkbox') { toggleBox(hit.dataset.sdtId); return true; }
    if (c.controlType === 'dropDownList' || c.controlType === 'comboBox') { pickChoice(hit.dataset.sdtId, hit); return true; }
    if (c.controlType === 'date') { pickDate(hit.dataset.sdtId); return true; }
    return false;
  }

  // ---- the field at the cursor: settings and removal
  function fieldAtCursor() {
    const $f = deps.editor()?.state?.selection?.$from;
    for (let d = $f?.depth ?? 0; d > 0; d--) { const n = $f.node(d); if (/structuredContent/i.test(n.type.name) && n.attrs?.id != null) return String(n.attrs.id); }
    const node = deps.editor()?.state?.selection?.node;
    if (node && /structuredContent/i.test(node.type.name) && node.attrs?.id != null) return String(node.attrs.id);
    return null;
  }
  async function fieldSettings(given) {
    const nodeId = given || fieldAtCursor();
    if (!nodeId) { toast('Click in a form field first', true); return; }
    const c = info(nodeId);
    if (!c) { toast('Click in a form field first', true); return; }
    const p = c.properties || {};
    const fields = [{ id: 'title', label: 'Title', value: p.alias || '' }];
    if (c.controlType === 'dropDownList' || c.controlType === 'comboBox') fields.push({ id: 'choices', label: 'Choices, one per line', type: 'textarea', rows: 6, value: (p.items || []).map((i) => i.displayText).join('\n') });
    if (c.controlType === 'date') fields.push({ id: 'format', label: 'Date format', type: 'select', value: DATE_FORMATS.some((f) => f.value === p.dateFormat) ? p.dateFormat : 'd MMMM yyyy', options: DATE_FORMATS });
    fields.push({ id: 'locked', type: 'checkbox', label: 'Can be filled in but not deleted', value: c.lockMode === 'sdtLocked' });
    const res = await openDialog({
      title: 'Form field settings',
      width: 440,
      fields,
      actions: [{ id: 'remove', label: 'Remove field', danger: true }, { id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Save', primary: true }],
    });
    if (!res) return;
    const t = target(nodeId);
    if (res.action === 'remove') { await act('Remove field', () => api().contentControls.unwrap({ target: t }), 'Field removed; its text stays'); return; }
    const v = res.values;
    await act('Form field', async () => {
      // unlocked while it changes, locked again last if asked
      let r = c.lockMode !== 'unlocked' ? await api().contentControls.setLockMode({ target: t, lockMode: 'unlocked' }) : { success: true };
      if (deps.failed(r)) return r;
      r = await api().contentControls.patch({ target: t, alias: v.title || null });
      if (deps.failed(r)) return r;
      if (v.choices !== undefined) {
        const items = v.choices.split('\n').map((x) => x.trim()).filter(Boolean).map((x) => ({ displayText: x, value: x }));
        r = await api().contentControls.choiceList.setItems({ target: t, items });
        if (deps.failed(r)) return r;
      }
      if (v.format) { r = await api().contentControls.date.setDisplayFormat({ target: t, format: v.format }); if (deps.failed(r)) return r; }
      return v.locked ? api().contentControls.setLockMode({ target: t, lockMode: 'sdtLocked' }) : r;
    });
  }

  function menu() {
    return [
      { label: 'Text field', icon: 'field-text', hint: 'A box to type in', action: () => insertField('text') },
      { label: 'Tick box', icon: 'field-check', hint: 'Click to tick or clear', action: () => insertField('checkbox') },
      { label: 'Dropdown list', icon: 'field-drop', hint: 'Pick one of your choices', action: () => insertField('dropDownList') },
      { label: 'Date picker', icon: 'calendar', hint: 'Pick a date from a calendar', action: () => insertField('date') },
      'sep',
      { label: 'Field settings...', icon: 'field-props', hint: 'For the field at the cursor', action: () => fieldSettings() },
    ];
  }

  return { insertField, onClick, fieldSettings, fieldAtCursor, menu, closeOpenMenu };
}
