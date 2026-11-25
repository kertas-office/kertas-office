// The new Writer (trial). Word documents open in an editor built on the GenOffice
// document engine (Apache-2.0, see NOTICE), hosted as tabs of the Kertas window. The current
// Writer stays the default; this one is switched on in Settings or with
// KERTAS_NEW_WRITER=1.
const { app, ipcMain, session, utilityProcess } = require('electron');
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

// What the editor asks its host about. (The editor's AI assistant is not part of Kertas's
// build of it; the one setting about it that the shared interface code still reads says "off".)
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
  serveDictionaries();
  // its own folder for its recent list, settings and templates, apart from Kertas's files
  host.setDataDir(path.join(app.getPath('userData'), 'writer'));
  // Kertas owns the application menu
  host.setDocsMenuGate(() => false);
  host.configureDocsRuntime({
    preloadPath: path.join(ROOT, 'preload', 'index.js'),
    rendererFile: path.join(ROOT, 'renderer', 'index.html'),
  });
  host.installRendererProtocol({ docs: path.join(ROOT, 'renderer') });
  // English, like the rest of Kertas: the editor has other languages, Kertas's own parts do not yet
  host.setDocsLanguage('en');
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

// ---- spelling, without the internet --------------------------------------------
// Chromium's spell checker fetches its dictionaries from a Google server: always on Linux, and
// on Windows for a language the system has no checker for. Kertas goes online for nothing, so
// the checker is pointed at this computer instead, where a tiny server hands out the
// dictionaries shipped in writer2/dict. A language that is not there gets no dictionary, and
// no request leaves the machine.
const DICT = path.join(ROOT, 'dict');
function serveDictionaries() {
  let files = [];
  try { files = fs.readdirSync(DICT).filter((f) => f.endsWith('.bdic')); } catch { /* built without them */ }
  const http = require('http');
  const server = http.createServer((req, res) => {
    const name = decodeURIComponent(String(req.url || '').split('?')[0].replace(/^\/+/, '')).toLowerCase();
    const file = files.find((f) => f.toLowerCase() === name);
    if (!file) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': 'application/octet-stream' });
    fs.createReadStream(path.join(DICT, file)).pipe(res);
  });
  server.on('error', () => { /* no port: spelling falls back to the system's own, or stays off */ });
  server.listen(0, '127.0.0.1', () => {
    const ses = session.defaultSession;
    ses.setSpellCheckerDictionaryDownloadURL(`http://127.0.0.1:${server.address().port}/`);
    if (process.platform === 'linux' && files.length) {
      // the shipped languages, the system's own first
      const have = files.map((f) => f.replace(/-\d+-\d+\.bdic$/, ''));
      const locale = app.getLocale();
      const pick = have.find((l) => l.toLowerCase() === locale.toLowerCase()) || have.find((l) => l.slice(0, 2) === locale.slice(0, 2)) || 'en-US';
      ses.setSpellCheckerEnabled(true);
      try { ses.setSpellCheckerLanguages([pick]); } catch { /* not a language this build knows */ }
    }
  });
  app.on('will-quit', () => server.close());
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

// A picture of the tab's editor as it is now, for the page to show in its place while the
// view steps aside for one of Kertas's dialogs
async function snapshot(id) {
  const t = tabs.get(id);
  if (!t || !t.shown) return null;
  try { return (await t.view.webContents.capturePage()).toDataURL(); } catch { return null; }
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

// PDF to Word with the layout kept. The work runs in a process of its own (see
// writer2-pdf-worker.cjs), so the window stays responsive however long it takes; the result
// waits here until its tab asks for it.
function convertPdfForTab(data) {
  if (!host) return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const child = utilityProcess.fork(path.join(__dirname, 'writer2-pdf-worker.cjs'), [], { serviceName: 'Kertas PDF to Word' });
    let done = false;
    const finish = (fn, value) => { if (done) return; done = true; clearTimeout(timer); try { child.kill(); } catch { /* already gone */ } fn(value); };
    // generous: a scanned book can take minutes; a process that has stopped answering is ended
    const timer = setTimeout(() => finish(reject, new Error('The conversion took too long and was stopped')), 15 * 60 * 1000);
    child.on('message', (msg) => {
      if (msg.docx) finish(resolve, keepResult(msg.docx));
      else if (msg.error) finish(reject, new Error(msg.error));
      else if (msg.total) tell('convert-progress', null, { page: msg.page, total: msg.total });
    });
    child.on('exit', (code) => finish(reject, new Error(`The converter stopped unexpectedly (${code})`)));
    child.once('spawn', () => child.postMessage({ pdf: Buffer.from(data), wasm: path.join(ROOT, 'wasm', 'pdfium.wasm') }));
  });
}

module.exports = { available, schemes, start, setTheme, setAuthor, stash, attach, openTab, layout, snapshot, isDirty, closeTab, command, convertPdfForTab };
