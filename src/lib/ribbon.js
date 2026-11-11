// Ribbon toolbar: a File button with a dropdown, tab labels, and grouped controls.
//
// createRibbon(container, {
//   app: 'writer' | 'sheets' | 'pdf',          // colours the File button
//   fileMenu: [{ id, label, icon, shortcut, action } | 'sep'],
//   tabs: [{ id, label, groups: [{ label, items: [Item] }] }],
//   inline: [Item],                            // controls shown in the tab row
// })
// Item kinds:
//   big | small (default)   button: { id, icon, label, title, action() }
//   menu                    button that opens a dropdown: { id, icon, label, items: [{ label, icon, hint, action, checked } | 'sep'] | () => items }
//   gallery                 strip of tiles: { id, tiles: [{ id, label, style, action }] }
//   color | select | search | label
//   number                  labelled number box: { id, label, unit, step, min, max, value, action(value) }
//   combo                   editable box with a dropdown (fonts, sizes): { id, value, width, items: () => [{ label, value, style }], action(value) }
//   swatch                  colour button with a bar (text colour, highlight): { id, icon, title, value, palette: 'text' | 'fill' | 'highlight', action(colour or 'none') }
// A group with rows: [[Item], [Item]] lays out two neat rows of small controls (Word's Font and Paragraph groups).
import { icon } from './icons.js';

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function itemHtml(it) {
  const title = esc(it.title || it.label || '');
  switch (it.kind) {
    case 'big':
      return `<button class="rb-big" data-id="${it.id}" title="${title}">${icon(it.icon, 22)}<span>${esc(it.label)}</span></button>`;
    case 'menu':
      return it.small
        ? `<button class="rb-small rb-menu" data-id="${it.id}" title="${title}">${icon(it.icon, 16)}${it.label ? `<span>${esc(it.label)}</span>` : ''}${icon('chevron', 12)}</button>`
        : `<button class="rb-big rb-menu" data-id="${it.id}" title="${title}">${icon(it.icon, 22)}<span>${esc(it.label)}${icon('chevron', 11)}</span></button>`;
    case 'gallery':
      return `<div class="rb-gallery" data-gallery="${it.id}">${it.tiles.map((t) => `<button class="rb-tile" data-id="${it.id}:${t.id}" title="${esc(t.title || t.label)}"><span class="tile-sample" style="${esc(t.style || '')}">${esc(t.sample || 'AaBbCc')}</span><span class="tile-name">${esc(t.label)}</span></button>`).join('')}</div>`;
    case 'color':
      return `<label class="rb-color" title="${title}">${it.icon ? icon(it.icon, 16) : ''}<input type="color" data-id="${it.id}" value="${it.value || '#000000'}"><span>${esc(it.label || '')}</span></label>`;
    case 'select':
      return `<select class="rb-select" data-id="${it.id}" title="${title}">${it.options.map((o) => `<option value="${esc(o.value)}" ${String(o.value) === String(it.value) ? 'selected' : ''}>${esc(o.label)}</option>`).join('')}</select>`;
    case 'search':
      return `<label class="rb-search">${icon('search', 15)}<input type="search" data-id="${it.id}" placeholder="${esc(it.label)}"></label>`;
    case 'label':
      return `<span class="rb-label" data-id="${it.id}">${esc(it.label || '')}</span>`;
    case 'number':
      return `<label class="rb-num" title="${title}">${it.icon ? icon(it.icon, 15) : ''}<span class="rb-num-label">${esc(it.label || '')}</span><input type="number" data-id="${it.id}" step="${it.step ?? 0.1}" ${it.min !== undefined ? `min="${it.min}"` : ''} ${it.max !== undefined ? `max="${it.max}"` : ''} value="${esc(it.value ?? '')}" placeholder="${esc(it.placeholder ?? '0')}"><span class="rb-num-unit">${esc(it.unit || '')}</span></label>`;
    default:
      return `<button class="rb-small" data-id="${it.id}" title="${title}">${icon(it.icon, 16)}${it.label ? `<span>${esc(it.label)}</span>` : ''}</button>`;
  }
}

// ---- Small controls used in rows
function comboHtml(it) {
  return `<span class="rb-combo" style="width:${it.width || 120}px" title="${esc(it.title || '')}"><input data-id="${it.id}" value="${esc(it.value ?? '')}" spellcheck="false" autocomplete="off"><button class="rb-combo-btn" data-combo="${it.id}" tabindex="-1">${icon('chevron', 12)}</button></span>`;
}
function swatchHtml(it) {
  return `<span class="rb-swatch" title="${esc(it.title || '')}"><button class="rb-ico rb-swatch-main" data-id="${it.id}">${icon(it.icon, 16)}<i class="rb-bar" style="background:${esc(it.value || '#000')}"></i></button><button class="rb-swatch-more" data-swatch="${it.id}" tabindex="-1">${icon('chevron', 10)}</button></span>`;
}
function rowItemHtml(it) {
  if (it.kind === 'combo') return comboHtml(it);
  if (it.kind === 'swatch') return swatchHtml(it);
  if (it.kind === 'sep') return '<span class="rb-rowsep"></span>';
  return iconCellHtml(it);
}
const PALETTES = {
  text: ['#000000', '#404040', '#7F7F7F', '#BFBFBF', '#FFFFFF', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0', '#1F3864', '#2F5496', '#833C0B', '#385623', '#7B2C8F'],
  fill: ['#FFFFFF', '#F2F2F2', '#D9D9D9', '#BFBFBF', '#808080', '#FDE9D9', '#FCE4D6', '#FFF2CC', '#E2EFDA', '#DDEBF7', '#D9E1F2', '#EDE7F6', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B0F0', '#0070C0', '#7030A0', '#000000'],
  highlight: ['#FFFF00', '#00FF00', '#00FFFF', '#FF00FF', '#0000FF', '#FF0000', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0', '#000000'],
};
/** Colour palette menu for a swatch button */
function openPalette(anchor, it, onPick) {
  const colours = PALETTES[it.palette || 'text'] || PALETTES.text;
  const grid = [...colours, 'none', 'more'];
  const html = grid.map((c) => (c === 'none' ? `<span class="pal-none">${icon('close', 14)}</span>` : c === 'more' ? `<span class="pal-more">${icon('palette', 15)}</span>` : `<span class="pal-chip" style="background:${c}"></span>`));
  const titles = grid.map((c) => (c === 'none' ? (it.palette === 'text' ? 'Automatic' : 'No colour') : c === 'more' ? 'More colours...' : c));
  const m = openMenu(anchor, { grid, html, titles, columns: 5, cell: 30, onPick: (c) => {
    if (c === 'more') {
      const inp = document.createElement('input');
      inp.type = 'color'; inp.value = it.value && it.value.startsWith('#') ? it.value : '#000000';
      inp.style.cssText = 'position:fixed;left:-100px;top:0;opacity:0';
      document.body.appendChild(inp);
      inp.addEventListener('change', () => { onPick(inp.value.toUpperCase()); inp.remove(); });
      inp.addEventListener('blur', () => setTimeout(() => inp.remove(), 500));
      closeOpenMenu();
      inp.click();
      return;
    }
    closeOpenMenu();
    onPick(c);
  } });
  m.classList.add('palette');
  return m;
}

// ---- Panel layout: every labelled option is the same tile (icon on top, label below);
// options shown as an icon alone sit together in an evenly spaced two-row grid.
const isIconOnly = (it) => (!it.kind || it.kind === 'small' || it.kind === 'menu') && !it.label && it.icon;
function tileHtml(it) {
  const title = esc(it.title || it.label || '');
  const label = esc(it.label || it.caption || it.title || '');
  switch (it.kind) {
    case 'menu':
      return `<button class="rb-big rb-menu" data-id="${it.id}" title="${title}"><span class="rb-ic">${icon(it.icon, 22)}<i class="rb-caret">${icon('chevron', 10)}</i></span><span class="rb-lb">${label}</span></button>`;
    case 'gallery': case 'search':
      return itemHtml(it);
    case 'color':
      return `<label class="rb-big rb-color-tile" title="${title}"><span class="rb-ic">${it.icon ? icon(it.icon, 20) : ''}<input type="color" data-id="${it.id}" value="${it.value || '#000000'}"></span><span class="rb-lb">${label || 'Colour'}</span></label>`;
    case 'select':
      return `<label class="rb-field" title="${title}"><span class="rb-ctl">${itemHtml(it)}</span><span class="rb-lb">${label}</span></label>`;
    case 'number':
      return `<label class="rb-field" title="${title}"><span class="rb-ctl rb-num-box"><input type="number" data-id="${it.id}" step="${it.step ?? 0.1}" ${it.min !== undefined ? `min="${it.min}"` : ''} ${it.max !== undefined ? `max="${it.max}"` : ''} value="${esc(it.value ?? '')}" placeholder="${esc(it.placeholder ?? '0')}"><span class="rb-num-unit">${esc(it.unit || '')}</span></span><span class="rb-lb">${it.icon ? icon(it.icon, 12) : ''}${esc(it.label || '')}</span></label>`;
    case 'label':
      return `<div class="rb-big rb-readout" data-id="${it.id}" title="${title}"><span>${esc(it.label || '')}</span><small>${esc(it.caption || '')}</small></div>`;
    default: // big, small with a label
      return `<button class="rb-big" data-id="${it.id}" title="${title}"><span class="rb-ic">${icon(it.icon, 22)}</span><span class="rb-lb">${label}</span></button>`;
  }
}
function iconCellHtml(it) {
  const title = esc(it.title || '');
  return it.kind === 'menu'
    ? `<button class="rb-ico rb-menu" data-id="${it.id}" title="${title}">${icon(it.icon, 16)}${icon('chevron', 9)}</button>`
    : `<button class="rb-ico" data-id="${it.id}" title="${title}">${icon(it.icon, 16)}</button>`;
}
function groupRowsHtml(rows) {
  return `<div class="rb-rows">${rows.map((r) => `<div class="rb-row2">${r.map(rowItemHtml).join('')}</div>`).join('')}</div>`;
}
function groupItemsHtml(items) {
  const icons = items.filter(isIconOnly);
  let gridDone = false;
  return items.map((it) => {
    if (!isIconOnly(it)) return tileHtml(it);
    if (gridDone) return '';
    gridDone = true;
    const rows = icons.length > 1 ? 2 : 1;
    return `<div class="rb-icogrid" style="--rows:${rows}">${icons.map(iconCellHtml).join('')}</div>`;
  }).join('');
}

let openMenuEl = null;
/** Opens a dropdown under an anchor. entries: [{ label, icon, hint, shortcut, checked, action } | 'sep' | { header }] */
export function openMenu(anchor, entries, { onClose } = {}) {
  closeOpenMenu();
  const menu = document.createElement('div');
  menu.className = 'menu glass';
  if (entries.grid) {
    // A grid of single characters, used for the symbol picker
    menu.classList.add('symbol-grid');
    // or of pictures (entries.html) with names (entries.titles), e.g. the shape gallery
    if (entries.columns) menu.style.gridTemplateColumns = `repeat(${entries.columns}, ${entries.cell || 34}px)`;
    if (entries.html) menu.classList.add('picture-grid');
    menu.innerHTML = entries.grid.map((c, i) => `<button class="sym" data-g="${i}" title="${esc(entries.titles?.[i] ?? `U+${String(c).codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`)}">${entries.html ? entries.html[i] : esc(c)}</button>`).join('');
    document.body.appendChild(menu);
    const r = anchor.getBoundingClientRect();
    menu.style.left = `${Math.max(8, Math.min(r.left, window.innerWidth - menu.offsetWidth - 8))}px`;
    menu.style.top = `${r.bottom + 6 + menu.offsetHeight > window.innerHeight - 8 && r.top - 6 - menu.offsetHeight > 8 ? r.top - 6 - menu.offsetHeight : r.bottom + 6}px`; // opens upward near the bottom of the window
    anchor.classList.add('open');
    menu.addEventListener('mousedown', (ev) => ev.preventDefault());
    menu.addEventListener('click', (ev) => {
      const b = ev.target.closest('[data-g]');
      if (b) entries.onPick(entries.grid[Number(b.dataset.g)]);
    });
    const away = (ev) => { if (!menu.contains(ev.target) && !anchor.contains(ev.target)) closeOpenMenu(); };
    setTimeout(() => document.addEventListener('mousedown', away, true), 0);
    openMenuEl = { menu, anchor, away, onClose };
    return menu;
  }
  menu.innerHTML = entries.map((m, i) => {
    if (m === 'sep') return '<div class="menu-sep"></div>';
    if (m.header) return `<div class="menu-header">${esc(m.header)}</div>`;
    const lead = m.checked !== undefined ? (m.checked ? icon('check', 15) : '<span style="width:15px"></span>') : icon(m.icon, 16);
    return `<button class="menu-item${m.disabled ? ' disabled' : ''}" data-m="${i}" ${m.disabled ? 'disabled' : ''}>${lead}<span${m.style ? ` style="${esc(m.style)}"` : ''}>${esc(m.label)}${m.hint ? `<small>${esc(m.hint)}</small>` : ''}</span><kbd>${esc(m.shortcut || '')}</kbd></button>`;
  }).join('');
  document.body.appendChild(menu);
  const r = anchor.getBoundingClientRect();
  const left = Math.min(r.left, window.innerWidth - menu.offsetWidth - 8);
  menu.style.left = `${Math.max(8, left)}px`;
  menu.style.top = `${r.bottom + 6 + menu.offsetHeight > window.innerHeight - 8 && r.top - 6 - menu.offsetHeight > 8 ? r.top - 6 - menu.offsetHeight : r.bottom + 6}px`; // opens upward near the bottom of the window
  anchor.classList.add('open');
  menu.addEventListener('mousedown', (ev) => ev.preventDefault()); // keep the editor's selection
  menu.addEventListener('click', (ev) => {
    const i = ev.target.closest('[data-m]')?.dataset.m;
    if (i === undefined) return;
    ev.stopPropagation();
    closeOpenMenu();
    entries[Number(i)]?.action?.();
  });
  const away = (ev) => { if (!menu.contains(ev.target) && !anchor.contains(ev.target)) closeOpenMenu(); };
  setTimeout(() => document.addEventListener('mousedown', away, true), 0);
  openMenuEl = { menu, anchor, away, onClose };
  return menu;
}
export function closeOpenMenu() {
  document.querySelectorAll('.rb-combo-btn.open, .rb-swatch-more.open').forEach((n) => n.classList.remove('open'));
  if (!openMenuEl) return;
  const { menu, anchor, away, onClose } = openMenuEl;
  openMenuEl = null;
  menu.remove();
  anchor.classList.remove('open');
  document.removeEventListener('mousedown', away, true);
  onClose?.();
}

export function createRibbon(container, { app, fileMenu = [], tabs = [], inline = [] }) {
  const root = document.createElement('div');
  root.className = `ribbon ribbon-${app}${tabs.length ? '' : ' ribbon-compact'}`;
  root.innerHTML = `
    <div class="rb-row">
      <button class="rb-file" data-file>${icon(app === 'sheets' ? 'sheet' : app === 'pdf' ? 'pdf' : 'doc', 16)}<span>File</span>${icon('chevron', 14)}</button>
      <div class="rb-tabs">${tabs.map((t, i) => `<button class="rb-tab${i === 0 ? ' active' : ''}${t.contextual ? ' contextual' : ''}" data-tab="${t.id}"${t.contextual ? ' hidden' : ''}>${esc(t.label)}</button>`).join('')}</div>
      <div class="rb-inline">${inline.map(itemHtml).join('')}</div>
    </div>
    ${tabs.length ? `<div class="rb-panel">${tabs.map((t, i) => `
      <div class="rb-pane${i === 0 ? ' active' : ''}" data-pane="${t.id}">
        ${t.groups.map((g) => `<div class="rb-group${g.wide ? ' wide' : ''}"><div class="rb-items">${Array.isArray(g.rows) ? groupRowsHtml(g.rows) + (g.items ? groupItemsHtml(g.items) : '') : groupItemsHtml(g.items)}</div><div class="rb-gname">${esc(g.label)}</div></div>`).join('<div class="rb-sep"></div>')}
      </div>`).join('')}</div>` : ''}`;
  container.appendChild(root);

  const items = new Map();
  let lostTab = null; // a contextual tab that was open when it had to hide
  const add = (it) => {
    items.set(it.id, it);
    if (it.kind === 'gallery') for (const t of it.tiles) items.set(`${it.id}:${t.id}`, t);
  };
  for (const t of tabs) for (const g of t.groups) { for (const it of g.items || []) add(it); for (const r of Array.isArray(g.rows) ? g.rows : []) for (const it of r) if (it.id) add(it); }
  for (const it of inline) add(it);

  // Tabs
  const activateTab = (id) => {
    root.querySelectorAll('.rb-tab').forEach((x) => x.classList.toggle('active', x.dataset.tab === id));
    root.querySelectorAll('.rb-pane').forEach((p) => p.classList.toggle('active', p.dataset.pane === id));
  };
  root.querySelectorAll('.rb-tab').forEach((b) => b.addEventListener('click', () => activateTab(b.dataset.tab)));
  // Double-click a tab to collapse the ribbon, like Office
  root.querySelectorAll('.rb-tab').forEach((b) => b.addEventListener('dblclick', () => root.classList.toggle('collapsed')));

  // Buttons never take focus, so the document keeps its selection
  root.addEventListener('mousedown', (e) => { if (e.target.closest('button')) e.preventDefault(); });

  // Controls
  root.addEventListener('click', (e) => {
    const cb = e.target.closest('[data-combo]');
    if (cb) {
      const it = items.get(cb.dataset.combo);
      if (cb.classList.contains('open')) { closeOpenMenu(); return; }
      const list = typeof it.items === 'function' ? it.items() : it.items;
      const cur = String(root.querySelector(`input[data-id="${it.id}"]`)?.value || '');
      const m = openMenu(cb.closest('.rb-combo'), list.map((o) => (o === 'sep' || o.header ? o : { label: o.label, style: o.style, checked: String(o.value).toLowerCase() === cur.toLowerCase(), action: () => { setValue(it.id, o.label); it.action?.(o.value); } })));
      m.classList.add('combo-menu');
      m.style.minWidth = `${cb.closest('.rb-combo').offsetWidth}px`;
      cb.classList.add('open');
      m.querySelector('.menu-item:has(svg)')?.scrollIntoView({ block: 'center' });
      return;
    }
    const sw = e.target.closest('[data-swatch]');
    if (sw) {
      const it = items.get(sw.dataset.swatch);
      if (sw.classList.contains('open')) { closeOpenMenu(); return; }
      openPalette(sw.closest('.rb-swatch'), it, (c) => { if (c !== 'none') setSwatch(it.id, c); it.action?.(c); });
      sw.classList.add('open');
      return;
    }
    const b = e.target.closest('button[data-id]');
    if (!b) return;
    const it = items.get(b.dataset.id);
    if (!it) return;
    if (it.kind === 'menu') {
      if (b.classList.contains('open')) { closeOpenMenu(); return; }
      openMenu(b, typeof it.items === 'function' ? it.items() : it.items);
      return;
    }
    if (it.kind === 'swatch') { it.action?.(it.value || '#000000'); return; }
    it.action?.();
  });
  // combos: Enter applies what was typed; Escape puts the value back
  root.addEventListener('keydown', (e) => {
    const inp = e.target.closest('.rb-combo input');
    if (!inp) return;
    const it = items.get(inp.dataset.id);
    if (e.key === 'Enter') { e.preventDefault(); inp.blur(); it?.action?.(inp.value); }
    if (e.key === 'Escape') { inp.value = inp.dataset.prev || ''; inp.blur(); }
  });
  root.addEventListener('focusin', (e) => { const inp = e.target.closest('.rb-combo input'); if (inp) { inp.dataset.prev = inp.value; inp.select(); } });
  function setValue(id, v) { root.querySelectorAll(`input[data-id="${id}"]`).forEach((n) => { if (document.activeElement !== n) n.value = v ?? ''; }); }
  function setSwatch(id, c) { const it = items.get(id); if (it) it.value = c; root.querySelectorAll(`[data-id="${id}"] .rb-bar`).forEach((b) => { b.style.background = c; }); }
  root.addEventListener('change', (e) => {
    if (e.target.closest('.rb-combo')) return; // combos apply on Enter or a pick
    const id = e.target.dataset?.id;
    if (id && e.target.type !== 'search') items.get(id)?.action?.(e.target.type === 'number' ? (e.target.value === '' ? null : Number(e.target.value)) : e.target.value);
  });
  root.addEventListener('input', (e) => {
    const id = e.target.dataset?.id;
    if (id && e.target.type === 'search') items.get(id)?.action?.(e.target.value);
  });
  root.addEventListener('keydown', (e) => { if (e.target.matches('input')) e.stopPropagation(); });

  // File dropdown
  const fileBtn = root.querySelector('[data-file]');
  fileBtn.addEventListener('click', () => {
    if (fileBtn.classList.contains('open')) { closeOpenMenu(); return; }
    openMenu(fileBtn, fileMenu.map((m) => (m === 'sep' ? 'sep' : { label: m.label, icon: m.icon, shortcut: m.shortcut, action: m.action })));
  });

  return {
    el: root,
    activateTab,
    /** Marks a toggle-style button on or off */
    setOn(id, on) { root.querySelectorAll(`button[data-id="${id}"]`).forEach((n) => n.classList.toggle('on', !!on)); },
    setLabel(id, text) { const n = root.querySelector(`[data-id="${id}"]`); if (n) { const s = n.querySelector('.rb-lb') || n.querySelector('span') || n; s.firstChild && s.firstChild.nodeType === 3 ? (s.firstChild.textContent = text) : (s.textContent = text); } },
    setDisabled(id, off) { root.querySelectorAll(`[data-id="${id}"]`).forEach((n) => { n.disabled = !!off; }); },
    get(id) { return root.querySelector(`[data-id="${id}"]`); },
    value(id) { return root.querySelector(`[data-id="${id}"]`)?.value; },
    /** Shows or hides a contextual tab (like Word's Table tab); hiding the open one returns to Home */
    setTabVisible(id, on) {
      const tab = root.querySelector(`.rb-tab[data-tab="${id}"]`);
      if (!tab || tab.hidden === !on) return;
      tab.hidden = !on;
      if (!on && tab.classList.contains('active')) { activateTab(tabs[0]?.id); lostTab = { id, at: Date.now() }; }
      // shown again straight after (an edit briefly moved the selection): go back to it
      if (on && lostTab?.id === id && Date.now() - lostTab.at < 2500) { activateTab(id); lostTab = null; }
    },
    setValue,
    setSwatch,
    dispose() { closeOpenMenu(); },
  };
}
