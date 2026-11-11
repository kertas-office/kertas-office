// Small form-dialog builder shared by every module.
//
// openDialog({
//   title, description, width,
//   fields: [{ id, label, type: 'text'|'password'|'number'|'textarea'|'select'|'checkbox'|'color'|'range'|'segmented'|'info',
//              value, options: [{value,label}], placeholder, hint, min, max, step, show(values) }],
//   actions: [{ id, label, primary, danger }],      // last one is used for Escape when it is 'cancel'
//   validate(values, action) -> error string or null,
//   onChange(values, formEl)                          // after every edit, and once on open
// }) -> Promise<{ action, values } | null>
import { icon } from './icons.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function fieldHtml(f) {
  const id = `dlg-${f.id}`;
  const hint = f.hint ? `<div class="dlg-hint">${esc(f.hint)}</div>` : '';
  switch (f.type) {
    case 'info':
      return `<div class="dlg-row" data-row="${f.id}"><div class="dlg-info">${f.html ?? esc(f.label)}</div></div>`;
    case 'checkbox':
      return `<div class="dlg-row" data-row="${f.id}"><label class="dlg-check"><input type="checkbox" id="${id}" data-f="${f.id}" ${f.value ? 'checked' : ''}><span class="box">${icon('check', 12)}</span><span>${esc(f.label)}</span></label>${hint}</div>`;
    case 'select':
      return `<div class="dlg-row" data-row="${f.id}"><label for="${id}">${esc(f.label)}</label><select class="rb-select dlg-input" id="${id}" data-f="${f.id}">${f.options.map((o) => `<option value="${esc(o.value)}" ${String(o.value) === String(f.value) ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}</select>${hint}</div>`;
    case 'segmented':
      return `<div class="dlg-row" data-row="${f.id}">${f.label ? `<label>${esc(f.label)}</label>` : ''}<div class="segmented" data-f="${f.id}" data-value="${esc(f.value)}">${f.options.map((o) => `<button type="button" data-v="${esc(o.value)}" class="${String(o.value) === String(f.value) ? 'on' : ''}">${o.icon ? icon(o.icon, 15) : ''}${esc(o.label)}</button>`).join('')}</div>${hint}</div>`;
    case 'textarea':
      return `<div class="dlg-row" data-row="${f.id}"><label for="${id}">${esc(f.label)}</label><textarea class="dlg-input" id="${id}" data-f="${f.id}" rows="${f.rows || 4}" placeholder="${esc(f.placeholder)}">${esc(f.value)}</textarea>${hint}</div>`;
    case 'color':
      return `<div class="dlg-row inline" data-row="${f.id}"><label for="${id}">${esc(f.label)}</label><span class="rb-color"><input type="color" id="${id}" data-f="${f.id}" value="${esc(f.value || '#000000')}"></span>${hint}</div>`;
    case 'range':
      return `<div class="dlg-row" data-row="${f.id}"><label for="${id}">${esc(f.label)} <output>${esc(f.format ? f.format(f.value) : f.value)}</output></label><input type="range" id="${id}" data-f="${f.id}" min="${f.min}" max="${f.max}" step="${f.step || 1}" value="${esc(f.value)}">${hint}</div>`;
    default:
      return `<div class="dlg-row" data-row="${f.id}"><label for="${id}">${esc(f.label)}</label><input class="dlg-input" id="${id}" data-f="${f.id}" type="${f.type || 'text'}" value="${esc(f.value)}" placeholder="${esc(f.placeholder)}" ${f.min !== undefined ? `min="${f.min}"` : ''} ${f.max !== undefined ? `max="${f.max}"` : ''} ${f.autocomplete ? `autocomplete="${f.autocomplete}"` : ''}>${hint}</div>`;
  }
}

export function openDialog({ title, description, fields = [], actions, validate, width = 480, onChange }) {
  const acts = actions || [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'OK', primary: true }];
  return new Promise((resolve) => {
    const back = document.createElement('div');
    back.className = 'modal-backdrop';
    back.innerHTML = `
      <form class="modal glass dialog" style="width:min(${width}px,92vw)" novalidate>
        <h3>${esc(title)}</h3>
        ${description ? `<p>${description}</p>` : ''}
        <div class="dlg-fields">${fields.map(fieldHtml).join('')}</div>
        <div class="dlg-error" role="alert"></div>
        <div class="row">${acts.map((a) => `<button type="${a.primary ? 'submit' : 'button'}" class="chip${a.primary ? ' primary' : ''}${a.danger ? ' danger' : ''}" data-a="${a.id}">${a.icon ? icon(a.icon, 14) : ''}${esc(a.label)}</button>`).join('')}</div>
      </form>`;
    document.body.appendChild(back);
    const form = back.querySelector('form');
    const err = back.querySelector('.dlg-error');

    const values = () => {
      const v = {};
      for (const f of fields) {
        if (f.type === 'info') continue;
        const n = form.querySelector(`[data-f="${f.id}"]`);
        if (!n) continue;
        if (f.type === 'checkbox') v[f.id] = n.checked;
        else if (f.type === 'segmented') v[f.id] = n.dataset.value;
        else if (f.type === 'number' || f.type === 'range') v[f.id] = n.value === '' ? null : Number(n.value);
        else v[f.id] = n.value;
      }
      return v;
    };
    const refresh = () => {
      const v = values();
      for (const f of fields) {
        if (!f.show) continue;
        const row = form.querySelector(`[data-row="${f.id}"]`);
        if (row) row.style.display = f.show(v) ? '' : 'none';
      }
      form.querySelectorAll('input[type=range]').forEach((r) => {
        const f = fields.find((x) => x.id === r.dataset.f);
        const out = r.parentElement.querySelector('output');
        if (out) out.textContent = f?.format ? f.format(Number(r.value)) : r.value;
      });
      onChange?.(v, form);
    };
    const close = (result) => { back.remove(); resolve(result); };
    const finish = (action) => {
      const v = values();
      const cancel = action === 'cancel';
      const msg = !cancel && validate ? validate(v, action) : null;
      if (msg) { err.textContent = msg; err.classList.add('show'); return; }
      close(cancel ? null : { action, values: v });
    };

    form.addEventListener('input', refresh);
    form.addEventListener('change', refresh);
    form.addEventListener('click', (e) => {
      const seg = e.target.closest('.segmented button');
      if (seg) {
        const group = seg.parentElement;
        group.dataset.value = seg.dataset.v;
        group.querySelectorAll('button').forEach((b) => b.classList.toggle('on', b === seg));
        refresh();
        return;
      }
      const a = e.target.closest('[data-a]');
      if (a && a.type !== 'submit') finish(a.dataset.a);
    });
    form.addEventListener('submit', (e) => { e.preventDefault(); finish(e.submitter?.dataset.a || acts.find((x) => x.primary)?.id || 'ok'); });
    form.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); close(null); }
      e.stopPropagation();
    });
    back.addEventListener('mousedown', (e) => { if (e.target === back) close(null); });
    refresh();
    (form.querySelector('.dlg-input, input') || form.querySelector('button'))?.focus();
  });
}
