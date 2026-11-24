// The new Writer (trial). Word documents open in an editor built on the GenOffice
// document engine (Apache-2.0, see NOTICE), hosted in its own window. The current
// Writer stays the default; this one is switched on in Settings or with
// KERTAS_NEW_WRITER=1.
const { app, ipcMain, nativeTheme } = require('electron');
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
    'app:get-theme': () => 'light',
    'app:get-document-theme': () => 'light',
    'app:get-measurement-unit': () => ({ unit: inch ? 'in' : 'cm', explicit: false }),
    'app:set-measurement-unit': () => ({ unit: inch ? 'in' : 'cm', explicit: false }),
    'app:get-author-name': () => { try { return os.userInfo().username; } catch { return 'Author'; } },
    'app:get-auto-save-default': () => ({ on: false, updatedAt: 0 }),
    'app:get-ai-panel-prefs': () => AI_OFF,
    'app:set-ai-panel-prefs': () => AI_OFF,
    'kb:list': () => [],
    'kb:search': () => [],
  };
  for (const [channel, fn] of Object.entries(answers)) ipcMain.handle(channel, fn);
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
    openTab: (openPath, options) => open(openPath, options),
    listTabs: () => [],
    focusTab: () => {},
    closeActiveTab: () => {},
  });
}

function open(filePath, options = {}) {
  if (!host) return false;
  const win = host.createDocsWindow(filePath || undefined);
  const id = win.webContents.id;
  windows.add(id);
  if (!filePath && options.newBlank !== false) host.markDocsNewBlank(id);
  win.on('closed', () => windows.delete(id));
  return true;
}

// A document made here (a converted PDF) opens untitled, ready to check and save
function openResult(bytes) {
  if (!host) return false;
  const win = host.createDocsWindow();
  const id = win.webContents.id;
  windows.add(id);
  host.queueDocsResultDoc(id, { bytes: Buffer.from(bytes) });
  win.on('closed', () => windows.delete(id));
  return true;
}

// PDF to Word with the layout kept (columns, tables, pictures), all on this computer
async function convertPdf(data) {
  if (!host) return false;
  const docx = await host.convertPdfBytesToDocx(new Uint8Array(data), path.join(ROOT, 'wasm', 'pdfium.wasm'));
  return openResult(docx);
}

module.exports = { available, schemes, start, open, convertPdf };
