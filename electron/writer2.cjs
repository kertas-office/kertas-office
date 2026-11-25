// The new Writer (trial). Word documents open in an editor built on the GenOffice
// document engine (Apache-2.0, see NOTICE), hosted as tabs of the Kertas window. The current
// Writer stays the default; this one is switched on in Settings or with
// KERTAS_NEW_WRITER=1.
const { app, ipcMain } = require('electron');
const os = require('os');
const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, 'writer2');
const MAIN = path.join(ROOT, 'main', 'index.cjs');
const available = fs.existsSync(MAIN);

let host = null;
const windows = new Set(); // webContents ids of the editor's windows

// Custom schemes must be declared once, before the app is ready, together with Kertas's own
function schemes() {
  if (!available) return [];
  host = require(MAIN);
  return [host.RENDERER_SCHEME_PRIVILEGE, host.DOCX_MEDIA_SCHEME_PRIVILEGE];
}

// What the editor asks its host about. Kertas has no AI assistant, so those
// questions get "nothing here" and the editor's AI features stay switched off.
function answerHostQuestions() {
  const inch = /-(US|LR|MM)$/i.test(app.getLocale()) || app.getLocale() === 'en';
  const answers = {
    'app:get-theme': () => (dark ? 'dark' : 'light'),
    'app:get-document-theme': () => 'light',
    'app:get-measurement-unit': () => ({ unit: inch ? 'in' : 'cm', explicit: false }),
    'app:set-measurement-unit': () => ({ unit: inch ? 'in' : 'cm', explicit: false }),
    'app:get-author-name': () => { if (author) return author; try { return os.userInfo().username; } catch { return 'Author'; } },
    'app:get-auto-save-default': () => ({ on: false, updatedAt: 0 }),
    'app:get-ai-panel-prefs': () => AI_OFF,
    'app:set-ai-panel-prefs': () => AI_OFF,
    'kb:list': () => [],
    'kb:search': () => [],
  };
  for (const [channel, fn] of Object.entries(answers)) ipcMain.handle(channel, fn);
}
// The name set in Kertas's About, used on tracked changes and comments
let author = '';
const setAuthor = (name) => { author = String(name || '').trim().slice(0, 120); return true; };
const stash = (data) => keepResult(new Uint8Array(data));

// Kertas's light or dark look; the page itself stays as the document has it
let dark = false;
function setTheme(isDark) {
  dark = !!isDark;
  for (const t of tabs.values()) if (!t.view.webContents.isDestroyed()) t.view.webContents.send('app:theme-changed', dark ? 'dark' : 'light');
}
const AI_OFF = { side: 'left', fontSize: 'default', customFontSize: 14, spellcheck: true, openInNewDocs: false, showAiSidebar: false, showAiAskTrigger: false };

// After the app is ready
function start() {
  if (!host) return;
  answerHostQuestions();
  // its own folder for its recent list, settings and templates, apart from Kertas's files
  host.setDataDir(path.join(app.getPath('userData'), 'writer'));
  // Kertas owns the application menu
  host.setDocsMenuGate(() => false);
  host.configureDocsRuntime({
    preloadPath: path.join(ROOT, 'preload', 'index.js'),
    rendererFile: path.join(ROOT, 'renderer', 'index.html'),
  });
  host.installRendererProtocol({ docs: path.join(ROOT, 'renderer') });
  host.setDocsLanguage(app.getLocale());
  host.installDocsContextMenu((contents) => windows.has(contents.id));
  host.registerDocsIpc();
  host.setDocsShellHooks({
    openTab: (openPath, options) => { if (openPath) tell('open-path', null, { path: openPath }); else tell('new', null, {}); },
    openResultDocTab: (doc) => tell('result', null, { token: keepResult(doc.bytes) }),
    listTabs: () => [],
    focusTab: () => {},
    closeActiveTab: () => { const id = activeTab(); if (id) tell('close', id, {}); },
  });
  host.setActiveDocsResolver(() => { const t = tabs.get(activeTab()); return t ? t.view.webContents : null; });
  host.setDocsFileSavedHook((wc, filePath) => { const id = tabOf(wc.id); if (id) tell('path', id, { path: filePath }); });
  host.setDocsFileOpenedHook((wcId, filePath) => { const id = tabOf(wcId); if (id) tell('path', id, { path: filePath }); });
}

// ---- as tabs of the Kertas window --------------------------------------------
// Each new-Writer tab is a view of its own, laid over the tab's pane in the Kertas
// window. Kertas's page says where the pane is and when it shows; the view follows.
let mainWin = null;
const tabs = new Map(); // tab id -> { view, wcId, rect }
const results = new Map(); // token -> bytes of a document made here, waiting for its tab
let resultSeq = 0;

function attach(win) {
  mainWin = win;
  if (host) host.setDocsShellWindow(win);
}
const tell = (type, id, more) => { if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('kertas:new-writer-event', { type, id, ...more }); };
const tabOf = (wcId) => { for (const [id, t] of tabs) if (t.wcId === wcId) return id; return null; };
const activeTab = () => { for (const [id, t] of tabs) if (t.shown) return id; return null; };
const keepResult = (bytes) => { const token = `r${++resultSeq}`; results.set(token, Buffer.from(bytes)); return token; };

function openTab(id, { filePath = null, result = null } = {}) {
  if (!host || !mainWin || tabs.has(id)) return false;
  const view = host.createDocsView(filePath || undefined);
  const wc = view.webContents;
  const wcId = wc.id;
  windows.add(wcId);
  const bytes = result && results.get(result);
  if (bytes) { results.delete(result); host.queueDocsResultDoc(wcId, { bytes }); }
  else if (!filePath) host.markDocsNewBlank(wcId);
  view.setVisible(false);
  mainWin.contentView.addChildView(view);
  tabs.set(id, { view, wcId, shown: false });
  wc.on('page-title-updated', (_e, title) => tell('title', id, { title }));
  // Kertas's own tab keys still work while the editor has the keyboard
  wc.on('before-input-event', (e, input) => {
    if (input.type !== 'keyDown' || !(input.control || input.meta)) return;
    if (input.key === 'Tab') { e.preventDefault(); tell(input.shift ? 'prev-tab' : 'next-tab', id, {}); }
  });
  wc.once('destroyed', () => { windows.delete(wcId); tabs.delete(id); });
  return true;
}

// Where the tab's pane is in the window, in CSS pixels; null hides the view
function layout(id, rect) {
  const t = tabs.get(id);
  if (!t || !mainWin || mainWin.isDestroyed()) return;
  if (!rect) { if (t.shown) { t.shown = false; t.view.setVisible(false); } return; }
  const z = mainWin.webContents.getZoomFactor();
  t.view.setBounds({ x: Math.round(rect.x * z), y: Math.round(rect.y * z), width: Math.max(0, Math.round(rect.width * z)), height: Math.max(0, Math.round(rect.height * z)) });
  if (!t.shown) {
    for (const other of tabs.values()) if (other !== t && other.shown) { other.shown = false; other.view.setVisible(false); }
    t.shown = true;
    t.view.setVisible(true);
    if (rect.focus !== false) t.view.webContents.focus();
  }
}

async function isDirty(id) {
  const t = tabs.get(id);
  return t ? host.docsQueryDirty(t.view.webContents) : false;
}

// Asks about unsaved changes (Save / Don't Save / Cancel), then removes the tab's view
async function closeTab(id) {
  const t = tabs.get(id);
  if (!t) return true;
  if (!(await host.requestDocsClose(t.view.webContents, mainWin))) return false;
  tabs.delete(id);
  try { mainWin.contentView.removeChildView(t.view); } catch { /* window already gone */ }
  try { t.view.webContents.close({ waitForBeforeUnload: false }); } catch { /* already closed */ }
  return true;
}

// A menu or palette command from Kertas, run in the tab's editor
function command(id, name) {
  const t = tabs.get(id);
  if (t) t.view.webContents.send('menu:command', name);
}

// PDF to Word with the layout kept; the result waits here until its tab asks for it
async function convertPdfForTab(data) {
  if (!host) return null;
  return keepResult(await host.convertPdfBytesToDocx(new Uint8Array(data), path.join(ROOT, 'wasm', 'pdfium.wasm')));
}

module.exports = { available, schemes, start, setTheme, setAuthor, stash, attach, openTab, layout, isDirty, closeTab, command, convertPdfForTab };
