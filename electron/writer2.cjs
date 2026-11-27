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
// the engine files the PDF editor loads (PDFium, and HarfBuzz for fonts in edited text)
process.env.KERTAS_WASM_DIR = path.join(ROOT, 'wasm');

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
    'app:get-language': () => 'en',
    'app:get-csv-import-limit': () => null,
  };
  for (const [channel, fn] of Object.entries(answers)) ipcMain.handle(channel, fn);
}
// The name set in Kertas's About, used on tracked changes and comments
let author = '';
const setAuthor = (name) => {
  author = String(name || '').trim().slice(0, 120);
  // the spreadsheet and PDF editors read it from their settings file, for comments and notes
  if (host) {
    const file = path.join(app.getPath('userData'), 'writer', 'app-settings.json');
    let settings = {};
    try { settings = JSON.parse(fs.readFileSync(file, 'utf8')) || {}; } catch { /* first time */ }
    if (author) settings.authorName = author; else delete settings.authorName;
    try { fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, JSON.stringify(settings, null, 2)); } catch { /* best effort */ }
    for (const t of tabs.values()) if (!t.view.webContents.isDestroyed()) t.view.webContents.send('app:author-name-changed', author || os.userInfo().username);
  }
  return true;
};
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
  startSheetsAndPdf();
  // a document dropped on one of the editors opens in Kertas, as anywhere else in the window
  ipcMain.on('app:open-dropped-files', (e, paths) => {
    if (!windows.has(e.sender.id) || !Array.isArray(paths)) return;
    for (const p of paths.slice(0, 20)) if (typeof p === 'string' && path.isAbsolute(p) && fs.existsSync(p)) tell('open-path', null, { path: p });
  });
  // a command waiting for a PDF tab (Sign a PDF), taken once the document is ready
  ipcMain.handle('kertas:pdf-pending', (e) => { const name = pdfPending.get(e.sender.id) || null; pdfPending.delete(e.sender.id); return name; });
  // the File button of the spreadsheet and PDF editors: Kertas carries out the command
  const FILE_ACTIONS = new Set(['new-doc', 'new-sheet', 'new-slides', 'open', 'save', 'save-as', 'export-pdf', 'export-csv', 'print', 'close-tab']);
  ipcMain.handle('kertas:editor-file', (e, action) => {
    const id = tabOf(e.sender.id);
    if (id && FILE_ACTIONS.has(action)) tell('file', id, { action });
  });
}

// The spreadsheet and PDF editors from the same project, hosted the same way. Each is
// there only if this build has it; the spreadsheet one also needs its helper program,
// which reads and writes the workbook files.
// The installed app carries the helper built for its own processor (resources/sidecar, see
// scripts/build-sidecars.cjs); a development copy uses the one the build script made here.
const SIDECAR_EXE = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar';
const SIDECAR = app.isPackaged ? path.join(process.resourcesPath, 'sidecar', SIDECAR_EXE) : path.join(ROOT, 'native', SIDECAR_EXE);
const has = {
  writer: true,
  sheets: fs.existsSync(path.join(ROOT, 'sheets', 'renderer', 'index.html')) && fs.existsSync(SIDECAR),
  pdf: fs.existsSync(path.join(ROOT, 'pdf', 'renderer', 'index.html')),
  slides: fs.existsSync(path.join(ROOT, 'slides', 'renderer', 'index.html')),
};
function startSheetsAndPdf() {
  const openHere = (filePath) => { tell('open-path', null, { path: filePath }); return true; };
  const roots = { docs: path.join(ROOT, 'renderer') };
  if (has.sheets) {
    roots.sheets = path.join(ROOT, 'sheets', 'renderer');
    host.configureSheetsRuntime({
      preloadPath: path.join(ROOT, 'sheets', 'preload', 'index.js'),
      rendererFile: path.join(ROOT, 'sheets', 'renderer', 'index.html'),
      sidecarPath: SIDECAR,
      openGeneratedPath: openHere,
    });
    host.registerSheetsIpc();
    host.setSheetsWorkbookOpenedHook((wc, filePath) => { const id = tabOf(wc.id); if (id) tell('path', id, { path: filePath }); });
    host.setSheetsCloseTabHook(() => { const id = activeTab(); if (id) tell('close', id, {}); });
    app.on('will-quit', () => { try { host.stopSheetsSidecar(); } catch { /* never started */ } });
  }
  if (has.pdf) {
    roots.pdf = path.join(ROOT, 'pdf', 'renderer');
    host.configurePdfRuntime({
      preloadPath: path.join(ROOT, 'pdf', 'preload', 'index.js'),
      rendererFile: path.join(ROOT, 'pdf', 'renderer', 'index.html'),
      openGeneratedPath: openHere,
    });
    host.setPdfRenamedHook((wc, _from, to) => { const id = tabOf(wc.id); const t = tabs.get(id); if (t) { t.path = to; tell('path', id, { path: to }); } });
    // The PDF editor's Convert menu. The file as it is on disk is converted, in the
    // background process; a Word result opens as a document, a workbook is saved first.
    // Recognising text on scanned pages is Kertas's own as well (tesseract, offline, on every
    // system): each page picture goes to the Kertas page and its lines come back
    const asks = new Map();
    let askSeq = 0;
    host.setKertasOcr((png, langs) => new Promise((resolve) => {
      const req = ++askSeq;
      const timer = setTimeout(() => { asks.delete(req); resolve([]); }, 3 * 60 * 1000);
      asks.set(req, (lines) => { clearTimeout(timer); asks.delete(req); resolve(Array.isArray(lines) ? lines : []); });
      tell('ocr', null, { req, png, langs });
    }));
    ocrAnswer = (req, lines) => asks.get(req)?.(lines);
    // a question for the Kertas page, answered through ask-result (false if no answer comes)
    askPage = (kind, id, payload, wait = 10 * 60 * 1000) => new Promise((resolve) => {
      const req = ++askSeq;
      const timer = setTimeout(() => { asks.delete(req); resolve(false); }, wait);
      asks.set(req, (value) => { clearTimeout(timer); asks.delete(req); resolve(value); });
      tell(kind, id, { req, ...payload });
    });
    // Compress is Kertas's own (picture levels), run by the Kertas page on the tab's file
    ipcMain.handle('pdf:kertas-compress', (e) => {
      const id = tabOf(e.sender.id);
      const t = tabs.get(id);
      if (t && t.kind === 'pdf' && t.path) tell('compress', id, { path: t.path });
    });
    let converting = false;
    ipcMain.handle('pdf:convert-office', async (e, format) => {
      const t = tabs.get(tabOf(e.sender.id));
      if (!t || t.kind !== 'pdf' || !t.path || converting || !['docx', 'xlsx'].includes(format)) return;
      converting = true;
      try {
        // A scan converts best with Kertas's own converter, which rebuilds paragraphs in one
        // font and size from the recognised text; the page answers true when it did that
        if (format === 'docx' && await askPage('convert-scan', tabOf(e.sender.id), { path: t.path }) === true) return;
        const out = await convertPdfForTab(await fs.promises.readFile(t.path), format);
        if (format === 'docx') { tell('result', null, { token: out }); return; }
        const { dialog } = require('electron');
        const res = await dialog.showSaveDialog(mainWin, { defaultPath: t.path.replace(/\.pdf$/i, '') + '.xlsx', filters: [{ name: 'Excel workbook', extensions: ['xlsx'] }] });
        if (res.canceled || !res.filePath) return;
        await fs.promises.writeFile(res.filePath, out);
        tell('open-path', null, { path: res.filePath });
      } catch (err) { tell('error', null, { message: `Could not convert this PDF: ${String((err && err.message) || err)}` }); }
      finally { converting = false; tell('convert-done', null, {}); }
    });
    host.setPdfRedactionSavedHook((wc, filePath) => { const id = tabOf(wc.id); if (id) tell('path', id, { path: filePath }); });
  }
  if (has.slides) {
    roots.slides = path.join(ROOT, 'slides', 'renderer');
    host.configureSlidesRuntime({
      preloadPath: path.join(ROOT, 'slides', 'preload', 'index.js'),
      rendererFilePath: path.join(ROOT, 'slides', 'renderer', 'index.html'),
      openGeneratedPath: openHere,
    });
    host.registerSlidesIpc();
    host.setSlidesOpenedHook((wc, filePath) => { const id = tabOf(wc.id); const t = tabs.get(id); if (t) { t.path = filePath; tell('path', id, { path: filePath }); } });
    host.setSlidesCloseTabHook(() => { const id = activeTab(); if (id) tell('close', id, {}); });
    // a new deck's first Save asks where to keep it, suggesting the tab's name
    host.setSlidesFirstSaveAsks?.((wc) => tabs.get(tabOf(wc.id))?.name || null);
    // the slide show: the editor goes full screen over the whole window, and back
    host.setSlidesHostWindowHook(() => mainWin);
    host.setSlidesShowBleed((wc, on) => { const id = tabOf(wc.id); const t = tabs.get(id); if (t) { t.bleed = !!on; place(t); } });
  }
  host.installRendererProtocol(roots);
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
let ocrAnswer = () => {};
const pdfPending = new Map(); // webContents id -> command
const tabs = new Map(); // tab id -> { view, wcId, rect }
const results = new Map(); // token -> bytes of a document made here, waiting for its tab
let resultSeq = 0;

function attach(win) {
  mainWin = win;
  if (!host) return;
  host.setDocsShellWindow(win);
  if (has.sheets) host.setSheetsShellWindow(win);
  if (has.slides) host.setSlidesShellWindow(win);
  // a slide show covering the window follows it as it goes full screen and back
  win.on('resize', () => { for (const t of tabs.values()) if (t.bleed) place(t); });
}
const tell = (type, id, more) => { if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('kertas:new-writer-event', { type, id, ...more }); };
const tabOf = (wcId) => { for (const [id, t] of tabs) if (t.wcId === wcId) return id; return null; };
const activeTab = () => { for (const [id, t] of tabs) if (t.shown) return id; return null; };
const keepResult = (bytes) => { const token = `r${++resultSeq}`; results.set(token, Buffer.from(bytes)); return token; };

// A workbook that is not a file yet (new, from a template, converted from an older format)
// opens from a copy in a temporary folder; its first Save asks where to keep it, starting
// from its name in Documents, and a crash keeps a copy that is offered again next time.
const SHEETS_TEMP = path.join(os.tmpdir(), 'kertas-sheets');
async function sheetsBacking(name, bytes, recoverAs) {
  const docs = app.getPath('documents');
  const clean = (n) => String(n || 'Book.xlsx').replace(/[\\/:*?"<>|]/g, '_').replace(/\.(xlsx|xlsm|xls|ods|csv)$/i, '') || 'Book';
  let suggested = recoverAs || path.join(docs, `${clean(name)}.xlsx`);
  for (let i = 2; !recoverAs && (fs.existsSync(suggested) || host.sheetsSuggestedPathTaken(suggested)) && i < 200; i++) suggested = path.join(docs, `${clean(name)} ${i}.xlsx`);
  const tempDir = path.join(SHEETS_TEMP, require('crypto').randomUUID());
  fs.mkdirSync(tempDir, { recursive: true });
  const backing = path.join(tempDir, path.basename(suggested));
  host.markSheetsUnsavedNew(backing, suggested, tempDir);
  try { await fs.promises.writeFile(backing, bytes || Buffer.from(await host.blankXlsxBuffer())); } catch (e) { host.unmarkSheetsUnsavedNew(backing); throw e; }
  return backing;
}
const sheetRecoveries = () => { try { return host && has.sheets ? host.pendingUnsavedNewRecoveries() : []; } catch { return []; } };

async function openTab(id, { kind = 'writer', filePath = null, result = null, name = null, recoverAs = null } = {}) {
  if (!host || !mainWin || tabs.has(id) || !has[kind]) return false;
  let view;
  if (kind === 'sheets') {
    if (!filePath) {
      const bytes = result && results.get(result);
      if (result) results.delete(result);
      if (recoverAs && !sheetRecoveries().includes(recoverAs)) recoverAs = null;
      filePath = await sheetsBacking(name, bytes, recoverAs);
    }
    view = host.createSheetsView({ includeAiHandlers: false, openingWorkbook: true });
    host.queueWorkbookForView(view.webContents, filePath);
  } else if (kind === 'pdf') view = host.createPdfView(filePath || undefined);
  else if (kind === 'slides') {
    // a template (or other deck that is not a file yet) opens from a temporary copy as a new deck
    const bytes = !filePath && result && results.get(result);
    if (bytes) {
      results.delete(result);
      const tmp = path.join(os.tmpdir(), `kertas-slides-${require('crypto').randomUUID()}.pptx`);
      await fs.promises.writeFile(tmp, bytes);
      view = host.createSlidesView(tmp, { asNew: true });
      setTimeout(() => fs.promises.unlink(tmp).catch(() => {}), 60000);
    } else view = host.createSlidesView(filePath || undefined);
  }
  else view = host.createDocsView(filePath || undefined);
  const wc = view.webContents;
  const wcId = wc.id;
  windows.add(wcId);
  if (kind === 'writer') {
    const bytes = result && results.get(result);
    if (bytes) { results.delete(result); host.queueDocsResultDoc(wcId, { bytes }); }
    else if (!filePath) host.markDocsNewBlank(wcId);
  }
  view.setVisible(false);
  mainWin.contentView.addChildView(view);
  tabs.set(id, { view, wcId, kind, path: filePath, name, shown: false });
  if (kind === 'writer') wc.on('page-title-updated', (_e, title) => tell('title', id, { title }));
  // Kertas's own tab keys still work while the editor has the keyboard
  wc.on('before-input-event', (e, input) => {
    if (input.type !== 'keyDown' || !(input.control || input.meta)) return;
    if (input.key === 'Tab') { e.preventDefault(); tell(input.shift ? 'prev-tab' : 'next-tab', id, {}); return; }
    // the spreadsheet and PDF editors leave Save, Open, Print and the rest to their host's menu
    if (kind === 'writer' || input.alt) return;
    const key = String(input.key).toLowerCase();
    if (key === 'w') { e.preventDefault(); tell('close', id, {}); }
    else if ('sopk'.includes(key) && key.length === 1 || (input.shift && (key === 'e' || key === 'n'))) { e.preventDefault(); tell('shortcut', id, { key, shift: !!input.shift }); }
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
  t.rect = { x: Math.round(rect.x * z), y: Math.round(rect.y * z), width: Math.max(0, Math.round(rect.width * z)), height: Math.max(0, Math.round(rect.height * z)) };
  place(t);
  if (!t.shown) {
    for (const other of tabs.values()) if (other !== t && other.shown) { other.shown = false; other.view.setVisible(false); }
    t.shown = true;
    t.view.setVisible(true);
    if (t.kind === 'sheets') host.setActiveSheetsWebContents(t.view.webContents);
    if (t.kind === 'slides') host.setActiveSlidesWebContents(t.view.webContents);
    if (rect.focus !== false) t.view.webContents.focus();
  }
}

// The view over its pane, or over the whole window during a slide show
function place(t) {
  if (!mainWin || mainWin.isDestroyed()) return;
  if (t.bleed) {
    const [width, height] = mainWin.getContentSize();
    t.view.setBounds({ x: 0, y: 0, width, height });
  } else if (t.rect) t.view.setBounds(t.rect);
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
  if (!t) return false;
  if (t.kind === 'sheets') return host.sheetsPendingEditCount(t.wcId) > 0;
  if (t.kind === 'pdf') return host.pdfIsDirty(t.wcId);
  if (t.kind === 'slides') return host.slidesIsDirty(t.wcId);
  return host.docsQueryDirty(t.view.webContents);
}

// Asks about unsaved changes (Save / Don't Save / Cancel), then removes the tab's view
async function closeTab(id) {
  const t = tabs.get(id);
  if (!t) return true;
  const ask = t.kind === 'sheets' ? host.requestSheetsClose : t.kind === 'pdf' ? host.requestPdfClose : t.kind === 'slides' ? host.requestSlidesClose : host.requestDocsClose;
  if (!(await ask(t.view.webContents, mainWin))) return false;
  tabs.delete(id);
  try { mainWin.contentView.removeChildView(t.view); } catch { /* window already gone */ }
  try { t.view.webContents.close({ waitForBeforeUnload: false }); } catch { /* already closed */ }
  return true;
}

// A menu or palette command from Kertas, run in the tab's editor
function command(id, name) {
  const t = tabs.get(id);
  if (!t) return;
  if (t.kind === 'sheets') { host.setActiveSheetsWebContents(t.view.webContents); host.sendSheetsMenuAction(name); }
  else if (t.kind === 'slides') { host.setActiveSlidesWebContents(t.view.webContents); t.view.webContents.send('slides:menu', name); }
  else if (t.kind === 'pdf') {
    if (name === 'save') host.flushPdfSave(t.view.webContents, true);
    else if (name === 'print') host.sendPdfPrintRequest(t.view.webContents);
    else if (name === 'save-as') savePdfAs(t);
    else if (name === 'sign') { pdfPending.set(t.wcId, 'sign'); t.view.webContents.send('kertas:pdf-command', 'sign'); }
  } else t.view.webContents.send('menu:command', name);
}
async function savePdfAs(t) {
  const { dialog } = require('electron');
  const res = await dialog.showSaveDialog(mainWin, { filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (!res.canceled && res.filePath) host.requestPdfSaveAs(t.view.webContents, res.filePath);
}

// PDF to Word with the layout kept. The work runs in a process of its own (see
// writer2-pdf-worker.cjs), so the window stays responsive however long it takes; the result
// waits here until its tab asks for it.
function convertPdfForTab(data, format = 'docx', { raw = false, onProgress = null } = {}) {
  if (!host) return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const child = utilityProcess.fork(path.join(__dirname, 'writer2-pdf-worker.cjs'), [], { serviceName: 'Kertas PDF to Word' });
    let done = false;
    const finish = (fn, value) => { if (done) return; done = true; clearTimeout(timer); try { child.kill(); } catch { /* already gone */ } fn(value); };
    // generous: a scanned book can take minutes; a process that has stopped answering is ended
    const timer = setTimeout(() => finish(reject, new Error('The conversion took too long and was stopped')), 15 * 60 * 1000);
    child.on('message', (msg) => {
      if (msg.docx) finish(resolve, format === 'docx' && !raw ? keepResult(msg.docx) : Buffer.from(msg.docx));
      else if (msg.error) finish(reject, new Error(msg.error));
      else if (msg.total) {
        if (onProgress) onProgress(msg.page, msg.total);
        else tell('convert-progress', null, { page: msg.page, total: msg.total, format });
      }
    });
    child.on('exit', (code) => finish(reject, new Error(`The converter stopped unexpectedly (${code})`)));
    child.once('spawn', () => child.postMessage({ pdf: Buffer.from(data), format, wasm: path.join(ROOT, 'wasm', 'pdfium.wasm') }));
  });
}

// A Word document to PDF for the Convert app: the new Writer opens it out of sight and
// prints it to the target path the way its own Export as PDF does, with no dialog. Pages
// come out exactly as the Writer lays them out.
function docxToPdf(filePath, outPath) {
  if (!host || !has.writer) return Promise.reject(new Error('The Writer is not available in this build'));
  return new Promise((resolve, reject) => {
    const view = host.createDocsView(filePath);
    const wc = view.webContents;
    windows.add(wc.id);
    try { wc.setBackgroundThrottling(false); } catch { /* older Electron */ }
    // the Writer lays pages out only where it is painted: an invisible window that still paints
    const { BrowserWindow } = require('electron');
    const holder = new BrowserWindow({ show: false, width: 1280, height: 900, paintWhenInitiallyHidden: true, skipTaskbar: true, webPreferences: { backgroundThrottling: false } });
    holder.contentView.addChildView(view);
    view.setBounds({ x: 0, y: 0, width: 1280, height: 900 });
    let done = false;
    let tries = 0;
    let retry = null;
    const finish = (err, value) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      clearTimeout(retry);
      windows.delete(wc.id);
      try { wc.close({ waitForBeforeUnload: false }); } catch { /* already closed */ }
      try { holder.destroy(); } catch { /* already gone */ }
      if (err) reject(err); else resolve(value);
    };
    const timer = setTimeout(() => finish(new Error('The document took too long to convert')), 3 * 60 * 1000);
    host.setKertasAutoPdf(
      wc.id,
      outPath,
      (out, error) => (out ? finish(null, out) : finish(new Error(error || 'The PDF could not be written'))),
      () => clearTimeout(retry), // the export is under way: no more asking
    );
    // the document is laid out a moment after the page loads; asking again covers a slow start
    const ask = () => {
      if (done) return;
      if (++tries > 6) { finish(new Error('The document did not finish loading')); return; }
      wc.send('menu:command', 'export-pdf');
      retry = setTimeout(ask, 15000);
    };
    wc.once('did-finish-load', () => setTimeout(ask, 2500));
    wc.once('render-process-gone', () => finish(new Error('The Writer stopped while converting')));
  });
}

// The licence, notice and third-party notice files of the editors, opened in a folder the
// system can show (they are inside the app's archive otherwise)
async function openLicences() {
  const { shell } = require('electron');
  const dir = path.join(app.getPath('userData'), 'licences');
  fs.mkdirSync(dir, { recursive: true });
  for (const f of ['LICENSE', 'NOTICE', 'LICENSE-UNICODE.txt', 'THIRD-PARTY-NOTICES.txt', 'FONTS.md', 'README.txt']) {
    try { fs.copyFileSync(path.join(ROOT, f), path.join(dir, f === 'LICENSE' ? 'LICENSE-editors.txt' : f === 'NOTICE' ? 'NOTICE-editors.txt' : f)); } catch { /* not in this build */ }
  }
  return shell.openPath(dir);
}
const ocrResult = (arg) => ocrAnswer(Number(arg && arg.req), arg && arg.lines);
let askPage = async () => false;
const askResult = (arg) => ocrAnswer(Number(arg && arg.req), arg && arg.value);
module.exports = { available, has, ocrResult, askResult, sheetRecoveries, openLicences, schemes, start, setTheme, setAuthor, stash, attach, openTab, layout, snapshot, isDirty, closeTab, command, convertPdfForTab, docxToPdf };
