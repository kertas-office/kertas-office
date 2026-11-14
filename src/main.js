// Development builds freeze the editor's shared document JSON, so any code that edits it fails loudly
if (import.meta.env?.DEV) globalThis.__kertasFreezeJson = true;
// Kertas Office - app shell: title bar tabs, home, commands, open/save/export routing
import { userName, setUserName } from './lib/user.js';
import './styles.css';
import {
  openFiles, readPath, saveFile, recentFiles, pinRecent, removeRecent, confirmDialog, onNative, confirmClose,
  kindForName, extOf, isDesktop, platform, previewPdf, printToPdf, setNativeTheme,
  recoveryWrite, recoveryDelete, recoveryList, recoveryRead, convertLegacyNative, updates,
} from './lib/platform.js';
import { openDialog as formDialog } from './lib/dialog.js';
import { detectFeatures, atRisk } from './lib/fidelity.js';
import { isLegacy, openLegacy } from './lib/legacy.js';
import { templateById, templatesFor, templateBytes } from './lib/templates.js';
import { icon } from './lib/icons.js';
import { installFontLoader, installFontFaces } from './lib/font-assets.js';

installFontLoader();
installFontFaces();

const MODULES = {
  writer: { label: 'W', loader: () => import('./modules/writer.js'), newName: 'Document', ext: 'docx', title: 'Writer' },
  sheets: { label: 'S', loader: () => import('./modules/sheets.js'), newName: 'Book', ext: 'xlsx', title: 'Sheets' },
  pdf: { label: 'P', loader: () => import('./modules/pdf.js'), newName: 'Document', ext: 'pdf', title: 'PDF' },
};

const tabsEl = document.getElementById('tabs');
const viewsEl = document.getElementById('views');
const statusEl = document.getElementById('status-text');
const statusRight = document.getElementById('status-right');
const toastEl = document.getElementById('toast');

/** @type {Map<string, {id, kind, name, path, dirty, el, tabEl, instance}>} */
const tabs = new Map();
let activeId = 'home';
const untitled = { writer: 0, sheets: 0, pdf: 0 };

document.body.classList.add(platform);
const hydrateIcons = (root = document) => root.querySelectorAll('[data-ic]').forEach((n) => { n.innerHTML = icon(n.dataset.ic, 17); n.style.display = 'inline-grid'; });
hydrateIcons();

// ------------------------------------------------------------ helpers
export function toast(msg, isError = false) {
  toastEl.textContent = msg;
  toastEl.className = `toast show${isError ? ' error' : ''}`;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { toastEl.className = 'toast'; }, isError ? 5000 : 2400);
}
function setStatus(text) { statusEl.textContent = text; }
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const BADGE_ICON = { writer: 'doc', sheets: 'sheet', pdf: 'pdf' };
const badge = (kind) => `<span class="badge ${kind}">${icon(BADGE_ICON[kind] || 'doc', 11)}</span>`;
const isDark = () => document.documentElement.dataset.theme === 'dark';
const baseName = (n) => n.replace(/\.[^.]+$/, '');

function timeAgo(ts) {
  const s = (Date.now() - ts) / 1000;
  if (s < 60) return 'Just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  if (s < 86400 * 7) return `${Math.floor(s / 86400)} d ago`;
  return new Date(ts).toLocaleDateString();
}

// ------------------------------------------------------------ theme
function applyTheme(dark) {
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  try { localStorage.setItem('kertas-theme', dark ? 'dark' : 'light'); } catch { /* private mode */ }
  document.querySelector('#btn-theme [data-ic]').innerHTML = icon(dark ? 'sun' : 'moon', 17);
  for (const t of tabs.values()) t.instance?.setTheme?.(dark);
  setNativeTheme(dark);
}
document.getElementById('btn-theme').addEventListener('click', () => applyTheme(!isDark()));
applyTheme(isDark());

// ------------------------------------------------------------ tabs
function activate(id) {
  activeId = id;
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === id));
  document.querySelectorAll('.view').forEach((v) => v.classList.toggle('active', v.id === `view-${id}`));
  const tab = tabs.get(id);
  document.title = tab ? `${tab.name} - Kertas Office` : 'Kertas Office';
  setStatus(tab ? (tab.path || 'Not saved yet') : 'Ready');
  statusRight.textContent = tab ? MODULES[tab.kind].title : 'All files stay on this device';
  if (id === 'home') renderHome();
  tab?.instance?.onShow?.();
}

function updateTab(tab) {
  tab.tabEl.querySelector('.title').textContent = tab.name;
  tab.tabEl.title = tab.path || tab.name;
  tab.tabEl.classList.toggle('is-dirty', tab.dirty);
  if (tab.id === activeId) document.title = `${tab.name}${tab.dirty ? ' *' : ''} - Kertas Office`;
}

async function openTab(kind, { name, path = null, data = null, html = null } = {}) {
  if (path) for (const t of tabs.values()) if (t.path === path) { activate(t.id); return t; }
  const id = `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  if (!name) name = `${MODULES[kind].newName}${++untitled[kind]}.${MODULES[kind].ext}`;

  const tabEl = document.createElement('button');
  tabEl.className = 'tab';
  tabEl.dataset.tab = id;
  tabEl.innerHTML = `${badge(kind)}<span class="title"></span><span class="dirty"></span><span class="close" title="Close (Ctrl+W)">${icon('close', 13)}</span>`;
  tabEl.addEventListener('click', (e) => {
    if (e.target.closest('.close')) { e.stopPropagation(); closeTab(id); } else activate(id);
  });
  tabEl.addEventListener('auxclick', (e) => { if (e.button === 1) closeTab(id); });
  tabsEl.appendChild(tabEl);
  tabEl.scrollIntoView({ inline: 'nearest' });

  const el = document.createElement('section');
  el.className = 'view';
  el.id = `view-${id}`;
  el.innerHTML = `<div class="loading"><div><div class="spinner"></div>Opening ${esc(name)}</div></div>`;
  viewsEl.appendChild(el);

  const tab = { id, kind, name, path, dirty: false, el, tabEl, instance: null };
  tabs.set(id, tab);
  updateTab(tab);
  activate(id);

  try {
    const mod = await MODULES[kind].loader();
    el.innerHTML = '';
    tab.instance = await mod.mount(el, {
      data, name, html,
      dark: isDark(),
      setDirty: (d) => { if (d) tab.rev = (tab.rev || 0) + 1; if (tab.dirty !== d) { tab.dirty = d; updateTab(tab); } },
      setStatus: (s) => { if (activeId === id) setStatus(s); },
      toast,
      save: (opts) => saveTab(tab, opts),
      fileMenu: fileMenuFor(tab),
      printToPdf,
      // Lets a module open its output as a new tab (PDF to Word, mail merge...)
      openDocument: (k, opts) => openTab(k, opts).then((t) => { t?.instance && setTimeout(() => t.instance && (t.dirty = true, updateTab(t)), 600); return t; }),
    });
    tab.instance.setTheme?.(isDark());
    // What in this file Kertas cannot keep, to mention now and ask about before saving over it
    if (data && (kind === 'writer' || kind === 'sheets') && /\.(docx|docm|dotx|xlsx|xlsm|xltx)$/i.test(name)) {
      detectFeatures(data, kind).then((found) => {
        tab.risk = atRisk(found, kind);
        if (tab.risk.length) toast(`${name} has ${riskList(tab.risk, 2)}, which Kertas can't keep. You'll be asked before it is saved.`);
      }).catch(() => {});
    }
  } catch (err) {
    console.error(err);
    el.innerHTML = `<div class="loading"><div style="text-align:center">Could not open ${esc(name)}<br><small>${esc(err.message || err)}</small></div></div>`;
    toast(`Could not open ${name}`, true);
  }
  return tab;
}

async function closeTab(id) {
  const tab = tabs.get(id);
  if (!tab) return true;
  if (tab.dirty) {
    activate(id);
    const choice = await confirmDialog(`Save changes to ${tab.name}?`, 'Your changes will be lost if you do not save them.', ['Save', "Don't Save", 'Cancel']);
    if (choice === 2) return false;
    if (choice === 0 && !(await saveTab(tab))) return false;
  }
  try { tab.instance?.dispose?.(); } catch (e) { console.warn(e); }
  recoveryDelete(tab.id);
  tab.tabEl.remove();
  tab.el.remove();
  tabs.delete(id);
  if (activeId === id) {
    const rest = [...tabs.keys()];
    activate(rest.length ? rest[rest.length - 1] : 'home');
  }
  return true;
}

/** "macros and SmartArt diagrams", lower-case, at most `max` named */
function riskList(risk, max = 3) {
  // first letter lower-case only, so names like Excel and VBA keep their capitals
  const names = risk.map((r) => { const t = r.label.replace(/\s*\(.*\)$/, ''); return /^[A-Z][a-z]/.test(t) ? t[0].toLowerCase() + t.slice(1) : t; });
  if (names.length > max) return `${names.slice(0, max).join(', ')} and ${names.length - max} more`;
  return names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}` : names[0];
}

/** Before the first save of a file with content Kertas cannot keep: what to do */
async function confirmRisk(tab) {
  const res = await formDialog({
    title: "Some things in this file can't be kept",
    description: `Saving ${tab.name} in Kertas removes the following. The rest of the file is kept.`,
    width: 500,
    fields: [{ id: 'list', type: 'info', html: `<ul class="risk-list">${tab.risk.map((r) => `<li>${esc(r.label)}${r.count > 1 ? ` <span>(${r.count})</span>` : ''}</li>`).join('')}</ul><p class="risk-tip">To keep the original as it is, save your changes as a copy.</p>` }],
    actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'copy', label: 'Save a copy' }, { id: 'save', label: 'Save anyway', primary: true }],
  });
  if (!res || res.action === 'cancel') return 'cancel';
  return res.action;
}

async function saveTab(tab, { saveAs = false, format } = {}) {
  if (!tab?.instance) return false;
  const ext = format || (tab.kind === 'sheets' && extOf(tab.name) === 'csv' ? 'csv' : MODULES[tab.kind].ext);
  if (tab.risk?.length && !tab.riskOk && (ext === 'docx' || ext === 'xlsx')) {
    const choice = await confirmRisk(tab);
    if (choice === 'cancel') return false;
    if (choice === 'copy') saveAs = true;
    tab.riskOk = true; // asked once per file
  }
  try {
    setStatus('Saving...');
    const data = await tab.instance.save(ext);
    const name = `${baseName(tab.name)}.${ext}`;
    const inPlace = !saveAs && tab.path && extOf(tab.path) === ext;
    const res = await saveFile({ data, path: inPlace ? tab.path : null, name, kind: tab.kind, saveAs: !inPlace });
    if (!res) { setStatus(tab.path || 'Not saved yet'); return false; }
    Object.assign(tab, { name: res.name, path: res.path, dirty: false });
    tab.snapRev = tab.rev;
    recoveryDelete(tab.id);
    tab.instance.markSaved?.();
    updateTab(tab);
    setStatus(tab.path || `Downloaded ${res.name}`);
    toast(`Saved ${res.name}`);
    return true;
  } catch (err) {
    console.error(err);
    toast(`Save failed: ${err.message || err}`, true);
    setStatus('Save failed');
    return false;
  }
}

async function exportPdf(tab, { preview = false } = {}) {
  if (!tab?.instance?.exportPdf) { toast('Export to PDF is not available here'); return; }
  try {
    setStatus(preview ? 'Preparing print preview...' : 'Exporting PDF...');
    const data = await tab.instance.exportPdf();
    if (!data) return; // browser print dialog handled it
    const name = `${baseName(tab.name)}.pdf`;
    if (preview) { await previewPdf(data, name); setStatus('Print preview opened'); return; }
    const res = await saveFile({ data, name, kind: 'pdf', saveAs: true });
    if (res) { toast(`Exported ${res.name}`); setStatus(res.path || `Downloaded ${res.name}`); }
    else setStatus(tab.path || 'Not saved yet');
  } catch (err) {
    console.error(err);
    toast(`Export failed: ${err.message || err}`, true);
  }
}

async function openFileRecords(records) {
  const opened = [];
  for (const f of records) {
    const kind = kindForName(f.name);
    if (!kind) { toast(`Unsupported file: ${f.name}`, true); continue; }
    if (isLegacy(f.name)) {
      // Older formats are converted first, then open like any other document
      try {
        setStatus(`Converting ${f.name}...`);
        const conv = await openLegacy(f, isDesktop ? convertLegacyNative : null);
        opened.push(await openTab(kind, { ...f, data: conv.data || null, html: conv.html || null }));
        if (conv.note) toast(conv.note);
      } catch (e) {
        console.error(e);
        toast(`Could not open ${f.name}: ${e.message || e}`, true);
      }
      continue;
    }
    opened.push(await openTab(kind, f));
  }
  return opened;
}
async function openPath(p) {
  try { await openFileRecords([await readPath(p)]); } catch { toast(`Could not open ${p}`, true); }
}
async function openDialog(kind = 'all') { return openFileRecords(await openFiles(kind, true)); }

async function mergePdfs() {
  const files = await openFiles('pdf', true);
  if (!files.length) return;
  const tab = await openTab('pdf', files[0]);
  if (files.length > 1) await tab.instance?.appendFiles?.(files.slice(1));
}
async function imagesToPdfCmd() {
  const images = await openFiles('image', true);
  if (!images.length) return;
  try {
    const mod = await MODULES.pdf.loader();
    const data = await mod.createFromImages(images);
    if (!data) return;
    const tab = await openTab('pdf', { name: `${baseName(images[0].name)}.pdf`, data });
    if (tab) { tab.dirty = true; updateTab(tab); }
  } catch (err) {
    toast(`Could not create the PDF: ${err.message}`, true);
  }
}
async function signPdf() {
  const [file] = await openFiles('pdf', false);
  if (!file) return;
  const tab = await openTab('pdf', file);
  tab.instance?.startSignature?.();
}

// ------------------------------------------------------------ crash recovery
// Every minute, unsaved changes are copied to a recovery store. They are removed
// when the document is saved or closed, so anything left at start-up was lost.
const RECOVERY_EVERY = 60 * 1000;
let snapshotting = false;
async function snapshotAll() {
  if (snapshotting) return;
  snapshotting = true;
  try {
    for (const tab of tabs.values()) {
      if (!tab.dirty || !tab.instance || tab.rev === tab.snapRev) continue;
      const ext = tab.kind === 'sheets' && extOf(tab.name) === 'csv' ? 'csv' : MODULES[tab.kind].ext;
      try {
        const data = await tab.instance.save(ext, { background: true });
        await recoveryWrite(tab.id, { name: `${baseName(tab.name)}.${ext}`, path: tab.path, kind: tab.kind, savedAt: Date.now() }, data);
        tab.snapRev = tab.rev;
      } catch (e) { console.warn('snapshot failed', tab.name, e); }
    }
  } finally { snapshotting = false; }
}
setInterval(snapshotAll, RECOVERY_EVERY);
const liveTabIds = () => new Set(tabs.keys());

async function openRecovered(entry) {
  try {
    const data = await recoveryRead(entry.id);
    const tab = await openTab(entry.kind, { name: entry.name, data });
    tab.path = entry.path || null;
    tab.dirty = true;
    tab.rev = 1;
    updateTab(tab);
    await recoveryDelete(entry.id);
    toast(`Recovered ${entry.name}. Save it to keep the changes.`);
  } catch (e) {
    toast(`Could not recover ${entry.name}`, true);
  }
}

/** A new, unsaved file from one of the ready-made templates */
async function newFromTemplate(id) {
  const t = templateById(id);
  if (!t) return null;
  let data;
  try { data = await templateBytes(t); } catch { toast(`Could not open the ${t.name} template`, true); return null; }
  untitled[t.kind] += 1;
  return openTab(t.kind, { data, name: `${t.name} ${untitled[t.kind]}.${t.ext}` });
}

/** Picture gallery of the templates for Writer or Sheets */
function openTemplates(kind) {
  document.querySelector('.tpl-backdrop')?.remove();
  const list = templatesFor(kind);
  const blank = kind === 'writer' ? 'Blank document' : 'Blank workbook';
  const back = document.createElement('div');
  back.className = 'modal-backdrop tpl-backdrop';
  back.innerHTML = `<div class="modal glass tpl-gallery ${kind}" role="dialog" aria-modal="true" aria-labelledby="tpl-title">
      <div class="tpl-head"><div><h3 id="tpl-title">${kind === 'writer' ? 'Writer' : 'Sheets'} templates</h3>
        <p>Each one opens as a new file with example content to replace. ${kind === 'sheets' ? 'Blue figures are the ones to type; the rest works itself out.' : 'Fonts, spacing and page numbers are already set.'}</p></div>
        <button class="icon-btn" data-close title="Close (Esc)">${icon('close', 16)}</button></div>
      <div class="tpl-grid">
        <button class="tpl-card blank" data-blank><span class="tpl-pic"><span class="tpl-plus">${icon('plus', 28)}</span></span><strong>${blank}</strong><span class="desc">Start from an empty ${kind === 'writer' ? 'page' : 'sheet'}</span></button>
        ${list.map((t) => `<button class="tpl-card" data-tpl="${t.id}"><span class="tpl-pic"><img src="${t.preview}" alt="" loading="lazy"></span><strong>${esc(t.name)}</strong><span class="desc">${esc(t.desc)}</span></button>`).join('')}
      </div></div>`;
  const close = () => { back.remove(); document.removeEventListener('keydown', onKey, true); };
  const onKey = (e) => { if (e.key === 'Escape') { e.stopPropagation(); close(); } };
  document.addEventListener('keydown', onKey, true);
  back.addEventListener('click', async (e) => {
    if (e.target === back || e.target.closest('[data-close]')) { close(); return; }
    if (e.target.closest('[data-blank]')) { close(); openTab(kind); return; }
    const card = e.target.closest('[data-tpl]');
    if (card) { close(); await newFromTemplate(card.dataset.tpl); }
  });
  document.body.appendChild(back);
  back.querySelector('.tpl-card')?.focus();
}

// ------------------------------------------------------------ default app
const DEFAULT_ASKED = 'kertas-default-asked';
const defaultAsked = () => { try { return localStorage.getItem(DEFAULT_ASKED) === '1'; } catch { return true; } };
const markDefaultAsked = () => { try { localStorage.setItem(DEFAULT_ASKED, '1'); } catch { /* private mode */ } };
async function makeDefaultApp() {
  markDefaultAsked();
  const res = await window.kertasNative?.makeDefaultApp?.() || { how: 'manual' };
  const steps = {
    settings: { title: 'Choose Kertas in Settings', text: 'Windows Settings is now open at Default apps. Set Kertas Office for .docx, .xlsx and .pdf (and any other types you like). Windows asks you to do this yourself so no app can take over your files without asking.' },
    mac: { title: 'Make Kertas the default on your Mac', text: 'In Finder, select a .docx file and press Cmd+I. Under "Open with", choose Kertas Office, then click "Change All". Do the same for a .xlsx and a .pdf file.' },
    appimage: { title: 'Make Kertas the default', text: 'This copy runs as an AppImage, which has no menu entry of its own for your desktop to use. Install the .deb package (or add the AppImage to your menu with AppImageLauncher), then use this command again.' },
    manual: { title: 'Make Kertas the default', text: 'Right-click a .docx, .xlsx or .pdf file, choose "Open with", pick Kertas Office and tick "Always use this app".' },
    dev: { title: 'Not in a development build', text: 'The installed app registers itself for documents. Try this in a build made with npm run dist.' },
  };
  if (res.how === 'done') { toast('Kertas now opens your documents, spreadsheets and PDFs'); renderHome(); return; }
  const st = steps[res.how] || steps.manual;
  await formDialog({ title: st.title, description: st.text, fields: [], actions: [{ id: 'ok', label: 'Got it', primary: true }] });
  renderHome();
}

// ------------------------------------------------------------ commands
const active = () => tabs.get(activeId);
const COMMANDS = [
  { id: 'new-doc', label: 'New document', icon: 'doc', key: 'Ctrl+N', run: () => openTab('writer') },
  { id: 'new-sheet', label: 'New spreadsheet', icon: 'sheet', key: 'Ctrl+Shift+N', run: () => openTab('sheets') },
  { id: 'templates-writer', label: 'New document from a template...', icon: 'template', run: () => openTemplates('writer') },
  { id: 'templates-sheets', label: 'New workbook from a template...', icon: 'template', run: () => openTemplates('sheets') },
  { id: 'open', label: 'Open...', icon: 'open', key: 'Ctrl+O', run: () => openDialog() },
  { id: 'save', label: 'Save', icon: 'save', key: 'Ctrl+S', run: () => active() && saveTab(active()), needsTab: true },
  { id: 'save-as', label: 'Save as...', icon: 'save-as', key: 'Ctrl+Shift+S', run: () => active() && saveTab(active(), { saveAs: true }), needsTab: true },
  { id: 'export-pdf', label: 'Export as PDF...', icon: 'export', key: 'Ctrl+Shift+E', run: () => exportPdf(active()), needsTab: true },
  { id: 'print', label: 'Print...', icon: 'print', key: 'Ctrl+P', run: () => exportPdf(active(), { preview: true }), needsTab: true },
  { id: 'close-tab', label: 'Close tab', icon: 'close', key: 'Ctrl+W', run: () => active() && closeTab(activeId), needsTab: true },
  { id: 'merge-pdfs', label: 'Merge PDFs...', icon: 'merge', run: mergePdfs },
  { id: 'sign-pdf', label: 'Sign a PDF...', icon: 'sign', run: signPdf },
  { id: 'images-to-pdf', label: 'Images to PDF...', icon: 'file-image', run: imagesToPdfCmd },
  { id: 'home', label: 'Go to Home', icon: 'home', run: () => activate('home') },
  { id: 'theme', label: 'Toggle dark mode', icon: 'moon', run: () => applyTheme(!isDark()) },
  ...(isDesktop ? [{ id: 'default-app', label: 'Make Kertas the default app for documents', icon: 'shield-check', run: () => makeDefaultApp() }] : []),
  { id: 'check-updates', label: 'Check for updates', icon: 'refresh', run: () => checkUpdates(true) },
  { id: 'about', label: 'About Kertas Office', icon: 'sparkles', run: () => showAbout() },
  ...(SUPPORT_BUILD ? [{ id: 'support', label: 'Support Kertas (buy me a coffee)', icon: 'coffee', run: () => openSupport() }] : []),
];
const runCommand = (id) => COMMANDS.find((c) => c.id === id)?.run();

function fileMenuFor() {
  const pick = (id) => { const c = COMMANDS.find((x) => x.id === id); return { id, label: c.label, icon: c.icon, shortcut: c.key, action: c.run }; };
  return [pick('new-doc'), pick('new-sheet'), pick('open'), 'sep', pick('save'), pick('save-as'), pick('export-pdf'), pick('print'), 'sep', pick('close-tab')];
}

// Command palette (Ctrl+K)
function openPalette() {
  if (document.querySelector('.palette')) return;
  const back = document.createElement('div');
  back.className = 'modal-backdrop';
  back.style.background = 'rgba(15,23,42,.18)';
  const pal = document.createElement('div');
  pal.className = 'palette glass';
  pal.innerHTML = '<input placeholder="Type a command or search recent files..." /><div class="list"></div>';
  back.appendChild(pal);
  document.body.appendChild(back);
  const input = pal.querySelector('input');
  const list = pal.querySelector('.list');
  let items = []; let sel = 0;
  const close = () => back.remove();
  const render = async () => {
    const q = input.value.toLowerCase().trim();
    const recent = (await recentFiles()).slice(0, 8).map((r) => ({ label: r.name, sub: r.path, icon: kindForName(r.name) === 'sheets' ? 'sheet' : kindForName(r.name) === 'pdf' ? 'pdf' : 'doc', run: () => openPath(r.path) }));
    items = [...COMMANDS.filter((c) => !c.needsTab || active()).map((c) => ({ label: c.label, key: c.key, icon: c.icon, run: c.run })), ...recent]
      .filter((it) => !q || it.label.toLowerCase().includes(q) || it.sub?.toLowerCase().includes(q));
    sel = Math.min(sel, Math.max(0, items.length - 1));
    list.innerHTML = items.map((it, i) => `<button class="menu-item${i === sel ? ' sel' : ''}" data-i="${i}">${icon(it.icon, 16)}<span>${esc(it.label)}${it.sub ? `<small style="display:block;color:var(--muted)">${esc(it.sub)}</small>` : ''}</span><kbd>${esc(it.key || '')}</kbd></button>`).join('') || '<div class="empty">No matches</div>';
  };
  input.addEventListener('input', () => { sel = 0; render(); });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { sel = Math.min(items.length - 1, sel + 1); render(); e.preventDefault(); }
    if (e.key === 'ArrowUp') { sel = Math.max(0, sel - 1); render(); e.preventDefault(); }
    if (e.key === 'Enter') { const it = items[sel]; close(); it?.run(); }
    if (e.key === 'Escape') close();
    e.stopPropagation();
  });
  list.addEventListener('click', (e) => { const i = e.target.closest('[data-i]')?.dataset.i; if (i !== undefined) { close(); items[+i].run(); } });
  back.addEventListener('mousedown', (e) => { if (e.target === back) close(); });
  render();
  input.focus();
}

// ------------------------------------------------------------ home

/** Where a file is, without its name: up to two folders, e.g. "Documents" or "Documents › Finance" */
function folderLabel(full) {
  let parts = String(full || '').split(/[\\/]+/).filter(Boolean);
  parts.pop(); // the file name is already the title
  // the user's own folder (C:\Users\name, /Users/name, /home/name) says nothing new
  const home = parts.findIndex((p, k) => /^(users|home)$/i.test(p) && k <= 1);
  if (home >= 0 && parts.length === home + 2) return 'Home';
  if (home >= 0 && parts.length > home + 2) parts = parts.slice(home + 2);
  return parts.slice(-2).join(' › ');
}

/** One file in Pinned or Recent files, with its pin button */
function recentRow(r) {
  const k = kindForName(r.name);
  return `<div class="recent-item${r.pinned ? ' pinned' : ''}${r.missing ? ' missing' : ''}">
    <button class="recent-row" data-path="${esc(r.path)}" title="${esc(r.path)}"${r.missing ? ' data-missing="1"' : ''}>${k ? badge(k) : ''}
      <span style="min-width:0"><span class="name">${esc(r.name)}</span><span class="path">${esc(folderLabel(r.path))}</span></span>
      <span class="when">${r.missing ? 'Not found' : timeAgo(r.openedAt)}</span></button>
    ${isDesktop ? `<button class="pin-btn" data-pin="${esc(r.path)}" data-pinned="${r.pinned ? 1 : 0}" title="${r.pinned ? 'Unpin from Home' : 'Pin to Home'}" aria-label="${r.pinned ? 'Unpin' : 'Pin'} ${esc(r.name)}">${r.pinned ? `<span class="is-pinned">${icon('pin', 15)}</span><span class="unpin">${icon('pin-off', 15)}</span>` : icon('pin', 15)}</button>
    <button class="rm-btn" data-remove="${esc(r.path)}" title="Remove from list" aria-label="Remove ${esc(r.name)} from the list">${icon('close', 15)}</button>` : ''}
  </div>`;
}

async function renderHome() {
  const home = document.getElementById('view-home');
  const all = await recentFiles();
  const pinned = all.filter((r) => r.pinned).sort((a, b) => (a.pinnedAt || 0) - (b.pinnedAt || 0));
  const recent = all.filter((r) => !r.pinned).slice(0, 10);
  const live = liveTabIds();
  const recovered = (await recoveryList()).filter((r) => !live.has(r.id));
  const hour = new Date().getHours();
  const greet = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
  home.innerHTML = `
    <div class="home">
      <div class="hero">
        <div>
          <div class="hero-greet">${greet}</div>
          <h1>Your workspace.</h1>
          <p>Documents, spreadsheets and PDFs. All offline.</p>
          <div class="local-badge">${icon('shield-check', 15)}<span><strong>Everything stays on this computer.</strong> No account and no uploads.</span></div>
        </div>
        <div class="kbd-hint"><kbd>Ctrl</kbd><kbd>K</kbd> for commands</div>
      </div>
      ${isDesktop && !defaultAsked() ? `<div class="default-offer glass">
        <div class="do-text">${icon('shield-check', 18)}<div><strong>Open your files with Kertas?</strong><span>Make Kertas the app that opens .docx, .xlsx and .pdf files when you double-click them.</span></div></div>
        <button class="chip primary" data-default="yes">Set as default</button><button class="chip" data-default="no">Not now</button>
      </div>` : ''}
      ${recovered.length ? `<div class="recovered glass">
        <div class="rec-head">${icon('restore', 18)}<div><strong>Recovered documents</strong><span>These had unsaved changes when Kertas last closed.</span></div></div>
        ${recovered.map((r) => `<div class="rec-row">${badge(r.kind)}<span class="name">${esc(r.name)}</span><span class="when">${timeAgo(r.savedAt)}</span>
          <button class="chip primary" data-recover="${esc(r.id)}">Open</button><button class="chip" data-discard="${esc(r.id)}">Discard</button></div>`).join('')}
      </div>` : ''}
      <div class="new-grid">
        <div class="app-card writer">
          <div class="app-icon">${icon('doc', 24)}</div>
          <strong>Writer</strong><span class="desc">Word-compatible documents (.docx, also opens .doc, .rtf, .odt)</span>
          <div class="actions">
            <button class="chip primary" data-cmd="new-doc">${icon('plus', 14)}Blank</button>
            <button class="chip" data-cmd="templates-writer">${icon('template', 14)}Templates</button>
            <button class="chip" data-open="writer">${icon('open', 14)}Open</button>
          </div>
        </div>
        <div class="app-card sheets">
          <div class="app-icon">${icon('sheet', 24)}</div>
          <strong>Sheets</strong><span class="desc">Excel-compatible workbooks (.xlsx, .csv, also opens .xls, .ods)</span>
          <div class="actions">
            <button class="chip primary" data-cmd="new-sheet">${icon('plus', 14)}Blank</button>
            <button class="chip" data-cmd="templates-sheets">${icon('template', 14)}Templates</button>
            <button class="chip" data-open="sheets">${icon('open', 14)}Open</button>
          </div>
        </div>
        <div class="app-card pdf">
          <div class="app-icon">${icon('pdf', 24)}</div>
          <strong>PDF</strong><span class="desc">Fill forms, comment, sign, protect, redact and organise pages</span>
          <div class="actions">
            <button class="chip primary" data-open="pdf">${icon('open', 14)}Open PDF</button>
            <button class="chip" data-cmd="merge-pdfs">${icon('merge', 14)}Merge</button>
            <button class="chip" data-cmd="sign-pdf">${icon('sign', 14)}Sign</button>
          </div>
        </div>
      </div>
      ${pinned.length ? `<div class="section-head" style="margin-top:30px"><h2>Pinned</h2></div>
      <div class="recent glass">${pinned.map(recentRow).join('')}</div>` : ''}
      <div class="section-head" style="margin-top:30px"><h2>Recent files</h2>
        <button class="chip" data-open="all">${icon('open', 14)}Open file</button></div>
      <div class="recent glass">
        ${recent.length ? recent.map(recentRow).join('') : pinned.length
    ? '<div class="empty small">Other files you open will show up here.</div>'
    : `<div class="empty"><div class="big-ic">${icon('clock', 24)}</div>${isDesktop ? 'Files you open will show up here. Pin the ones you use most to keep them at the top.' : 'Recent files appear in the desktop app.'}<br><small>Tip: drag files onto this window to open them.</small></div>`}
      </div>
      <footer class="home-foot">Free &amp; Open Source<span>·</span>No Ads<span>·</span>No Subscriptions</footer>
    </div>`;
  home.querySelectorAll('[data-default]').forEach((b) => b.addEventListener('click', () => {
    if (b.dataset.default === 'yes') makeDefaultApp(); else { markDefaultAsked(); renderHome(); }
  }));
  home.querySelectorAll('[data-cmd]').forEach((b) => b.addEventListener('click', () => runCommand(b.dataset.cmd)));
  home.querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => openDialog(b.dataset.open)));
  home.querySelectorAll('[data-path]').forEach((b) => b.addEventListener('click', () => {
    if (b.dataset.missing) { toast('This file was moved or deleted. Remove it from the list, or open it from its new place.', true); return; }
    openPath(b.dataset.path);
  }));
  home.querySelectorAll('[data-pin]').forEach((b) => b.addEventListener('click', async (e) => {
    e.stopPropagation();
    const on = b.dataset.pinned !== '1';
    await pinRecent(b.dataset.pin, on);
    renderHome();
  }));
  home.querySelectorAll('[data-remove]').forEach((b) => b.addEventListener('click', async (e) => {
    e.stopPropagation();
    await removeRecent(b.dataset.remove);
    renderHome();
  }));
  home.querySelectorAll('[data-recover]').forEach((b) => b.addEventListener('click', () => openRecovered(recovered.find((r) => r.id === b.dataset.recover))));
  home.querySelectorAll('[data-discard]').forEach((b) => b.addEventListener('click', async () => { await recoveryDelete(b.dataset.discard); renderHome(); }));
}

// ------------------------------------------------------------ global wiring
document.querySelector('[data-tab="home"]').addEventListener('click', () => activate('home'));
document.getElementById('btn-palette').addEventListener('click', openPalette);

onNative('menu', runCommand);
onNative('update', (st) => showUpdate(st));
onNative('open-path', openPath);
onNative('request-close', async () => {
  for (const id of [...tabs.keys()]) if (!(await closeTab(id))) return;
  confirmClose();
});

window.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (k === 'k') { e.preventDefault(); openPalette(); return; }
  if (e.key === 'Tab') {
    e.preventDefault();
    const ids = ['home', ...tabs.keys()];
    const i = ids.indexOf(activeId);
    activate(ids[(i + (e.shiftKey ? -1 : 1) + ids.length) % ids.length]);
    return;
  }
  const map = {
    s: e.shiftKey ? 'save-as' : 'save', o: 'open', p: 'print',
    e: e.shiftKey ? 'export-pdf' : null, n: e.shiftKey ? 'new-sheet' : null,
  };
  if (map[k]) { e.preventDefault(); runCommand(map[k]); }
}, true);

if (!isDesktop) {
  window.addEventListener('beforeunload', (e) => {
    if ([...tabs.values()].some((t) => t.dirty)) { e.preventDefault(); e.returnValue = ''; }
  });
}

// Drag and drop
let dragDepth = 0;
// Pictures dragged over a Writer document are inserted by the editor where they land,
// so the "Drop files to open" cover stays away for them
const picturesIntoWriter = (dt) => active()?.kind === 'writer' && dt?.items?.length
  && [...dt.items].every((i) => i.kind === 'file' && i.type.startsWith('image/'));
const endDrag = () => { dragDepth = 0; document.body.classList.remove('dragging', 'drop-into'); };
window.addEventListener('dragenter', (e) => {
  if (!e.dataTransfer?.types.includes('Files') || picturesIntoWriter(e.dataTransfer)) return;
  dragDepth++;
  document.body.classList.add('dragging');
});
window.addEventListener('dragend', endDrag, true);
window.addEventListener('dragleave', () => { if (--dragDepth <= 0) endDrag(); });
window.addEventListener('dragover', (e) => {
  if (!e.dataTransfer?.types.includes('Files')) return;
  e.preventDefault();
  // a panel that takes files itself (the PDF Files tab, the page thumbnails) shows its own
  // highlight, so the "Drop files to open" cover steps aside while the pointer is over it
  document.body.classList.toggle('drop-into', !!fileDropTarget(e.target));
});
/** The nearest element in the open tab that handles dropped files itself, if any */
function fileDropTarget(node) {
  const t = node?.closest?.('[data-file-drop]');
  return t && typeof t.kertasFileDrop === 'function' && t.getClientRects().length ? t : null;
}
// Handled on the way down, before the editors: they would otherwise take (and stop)
// any drop on a page. Only pictures dropped on a Writer document go on to the editor.
window.addEventListener('drop', async (e) => {
  endDrag();
  if (!e.dataTransfer?.files?.length) return;
  if (picturesIntoWriter(e.dataTransfer)) {
    // the editor puts it where it lands on the page, or at the cursor when dropped elsewhere;
    // the browser itself must never take it and leave the app to show the picture
    if (!e.target.closest?.('.superdoc-page')) e.preventDefault();
    return;
  }
  e.preventDefault();
  e.stopPropagation();
  const target = fileDropTarget(e.target);
  const records = await Promise.all([...e.dataTransfer.files].map(async (f) => ({
    name: f.name, path: window.kertasNative?.pathForFile?.(f) || null, data: new Uint8Array(await f.arrayBuffer()),
  })));
  if (target) { await target.kertasFileDrop(records, e); return; }
  await openFileRecords(records);
}, true);

// Test hook so automated checks can drive the app
window.__kertas = { openTab, saveTab, tabs, closeTab, activate, exportPdf, runCommand, applyTheme, snapshotAll, newFromTemplate, openTemplates };

renderHome();


// ---------------------------------------------------------------- support link
// Kertas is free with no ads; a quiet coffee cup is the only ask. It opens in the
// system browser, nothing is tracked, and it can be hidden (or left out of a build
// with VITE_NO_SUPPORT_LINK=1, e.g. for schools and companies).
const SUPPORT_URL = 'https://buymeacoffee.com/hasheee';
const SUPPORT_BUILD = import.meta.env?.VITE_NO_SUPPORT_LINK !== '1';
const coffeeBtn = document.getElementById('btn-coffee');
const coffeeHidden = () => { try { return localStorage.getItem('kertas-hide-coffee') === '1'; } catch { return false; } };
function showCoffee() { coffeeBtn.hidden = !SUPPORT_BUILD || coffeeHidden(); }
function openSupport() { window.open(SUPPORT_URL, '_blank', 'noopener'); }
coffeeBtn.addEventListener('click', openSupport);
showCoffee();

// ---------------------------------------------------------------- updates and about
const APP_VERSION = typeof __APP_VERSION__ !== 'undefined' ? __APP_VERSION__ : 'dev';
const pill = document.getElementById('btn-update');
pill.querySelector('[data-ic]').innerHTML = icon('refresh', 15);
let lastUpdate = { state: 'idle' };
let announced = null;

function showUpdate(st) {
  lastUpdate = st || lastUpdate;
  const s = lastUpdate.state;
  pill.hidden = !(s === 'downloading' || s === 'ready');
  pill.classList.toggle('ready', s === 'ready');
  pill.querySelector('.label').textContent = s === 'ready' ? `Restart to update` : `Updating ${lastUpdate.percent || 0}%`;
  pill.title = s === 'ready' ? `Kertas Office ${lastUpdate.available} is ready. Restart to finish updating.` : `Downloading Kertas Office ${lastUpdate.available || ''}`;
  if (s === 'ready' && announced !== lastUpdate.available) {
    announced = lastUpdate.available;
    toast(`Kertas Office ${lastUpdate.available} is ready. It installs when you restart.`);
  }
}

async function installUpdate() {
  // Same as closing the window: unsaved work is offered for saving first
  for (const id of [...tabs.keys()]) if (!(await closeTab(id))) return;
  await updates('install');
}
pill.addEventListener('click', () => { if (lastUpdate.state === 'ready') installUpdate(); });

async function checkUpdates(manual = false) {
  if (!isDesktop) { if (manual) toast('The web version is always up to date'); return; }
  if (manual) toast('Checking for updates...');
  const st = await updates('check');
  showUpdate(st);
  if (!manual) return;
  if (st.state === 'current') toast(`You have the latest version (${APP_VERSION})`);
  else if (st.state === 'dev') toast('Updates are checked in the installed app, not in development builds');
  else if (st.state === 'unconfigured') toast('This build was made without an update source');
  else if (st.state === 'error') toast(`Could not check for updates: ${st.message || 'no connection'}`, true);
}

async function showAbout() {
  const st = isDesktop ? await updates('state') : { state: 'web' };
  const lo = isDesktop && window.kertasNative?.hasLibreOffice ? await window.kertasNative.hasLibreOffice() : false;
  const status = { idle: 'Not checked yet', checking: 'Checking...', current: 'Up to date', downloading: `Downloading ${st.available || ''}`, ready: `${st.available} ready, restart to install`, error: `Last check failed: ${st.message || ''}`, dev: 'Development build (no updates)', web: 'Web version', unconfigured: 'This build has no update source' }[st.state] || st.state;
  const res = await formDialog({
    title: 'Kertas Office',
    description: `Version ${APP_VERSION} · free and open source (AGPL-3.0-or-later)`,
    width: 460,
    fields: [
      { id: 'info', type: 'info', html: `<table class="stat-table"><tr><td>Updates</td><td>${status}</td></tr><tr><td>Older formats</td><td>${isDesktop ? (lo ? 'LibreOffice found: .doc, .xls, .rtf, .odt and .ods keep their formatting' : 'Built in: .xls and .ods, and the text of .doc. Install LibreOffice for full formatting and .rtf / .odt') : '.xls and .ods'}</td></tr><tr><td>Your files</td><td>Stay on this device</td></tr></table>` },
      { id: 'name', label: 'Your name (on tracked changes and comments)', value: userName(), placeholder: 'e.g. Aisha Rahman' },
      ...(isDesktop && st.packaged ? [{ id: 'auto', type: 'checkbox', label: 'Download updates automatically', value: st.auto !== false }] : []),
      ...(SUPPORT_BUILD ? [
        { id: 'support', type: 'info', html: `<div class="support-line">${icon('coffee', 16)}<span>Kertas is free and open source. If it saves you time, you can <a href="${SUPPORT_URL}" target="_blank" rel="noopener">buy me a coffee</a>.</span></div>` },
        { id: 'cup', type: 'checkbox', label: 'Show the coffee cup in the title bar', value: !coffeeHidden() },
      ] : []),
    ],
    actions: [
      ...(isDesktop ? [{ id: 'check', label: 'Check for updates' }] : []),
      { id: 'ok', label: 'Close', primary: true },
    ],
  });
  if (!res) return;
  if (res.values.name !== undefined && res.values.name.trim() !== userName()) { setUserName(res.values.name); toast(res.values.name.trim() ? `Your changes and comments now show as ${res.values.name.trim()}` : 'Name cleared'); }
  if (res.values.auto !== undefined && res.values.auto !== (st.auto !== false)) await updates('set-auto', res.values.auto);
  if (res.values.cup !== undefined) { try { localStorage.setItem('kertas-hide-coffee', res.values.cup ? '0' : '1'); } catch { /* private mode */ } showCoffee(); }
  if (res.action === 'check') checkUpdates(true);
}

if (isDesktop) updates('state').then(showUpdate).catch(() => {});

// Warm start: once the app is idle, load the Writer, Sheets and PDF code in the background, so the
// first document opens as fast as the second (loading and compiling them took most of a second)
{
  const idle = (fn) => (typeof requestIdleCallback === 'function' ? requestIdleCallback(fn, { timeout: 4000 }) : setTimeout(fn, 1500));
  const queue = ['writer', 'sheets', 'pdf'];
  const next = () => { const k = queue.shift(); if (!k) return; MODULES[k].loader().catch(() => {}).finally(() => idle(next)); };
  window.addEventListener('load', () => setTimeout(() => idle(next), 600), { once: true });
}
