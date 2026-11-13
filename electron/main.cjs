// Kertas Office - Electron main process
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, protocol, net, session } = require('electron');
const path = require('path');
const { pathToFileURL } = require('url');
const os = require('os');
const fs = require('fs/promises');

const DEV_URL = process.env.KERTAS_DEV_URL;
const SUPPORTED = ['.docx', '.xlsx', '.csv', '.pdf', '.doc', '.rtf', '.odt', '.xls', '.xlsm', '.xlsb', '.ods'];
const { convertLegacy, findSoffice } = require('./legacy.cjs');
const { setupUpdates } = require('./updates.cjs');
const RECENT_FILE = () => path.join(app.getPath('userData'), 'recent.json');
const MAX_RECENT = 10;

let mainWindow = null;
let rendererReady = false;
let forceClose = false; // set when quitting to install an update

// Serve the built UI from a private app:// origin instead of file://, so module
// workers (formula engine, PDF renderer) load the same way they do on the web.
const DIST = path.join(__dirname, '..', 'dist');
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, stream: true } },
]);
function registerAppProtocol() {
  protocol.handle('app', (req) => {
    const { pathname } = new URL(req.url);
    const target = path.normalize(path.join(DIST, decodeURIComponent(pathname)));
    if (target !== DIST && !target.startsWith(DIST + path.sep)) return new Response('Not found', { status: 404 });
    return net.fetch(pathToFileURL(target).toString());
  });
}
const pendingOpen = []; // files requested before the window was ready

// ---------- File access guard ----------
// The renderer shows untrusted documents, so the main process never takes its word for which
// files it may touch. It may read only document types, and write only to files the person
// opened or picked in a Save dialog in this session (or that crash recovery restored), and
// only as the formats Kertas saves. Even a compromised page could not read SSH keys or drop
// a script into a startup folder.
const granted = new Set();
const norm = (p) => path.resolve(String(p));
const grant = (p) => { if (typeof p === 'string' && p) granted.add(norm(p)); return p; };
const WRITE_TYPES = /\.(docx|xlsx|csv|pdf)$/i;
const SAVE_MANY_TYPES = /\.(docx|xlsx|pdf|png|jpe?g)$/i;
function assertReadable(p) {
  if (typeof p !== 'string' || !path.isAbsolute(p) || !RECENT_TYPES.test(p)) throw new Error('This file type cannot be opened');
}
function assertWritable(p) {
  if (typeof p !== 'string' || !path.isAbsolute(p) || !WRITE_TYPES.test(p) || !granted.has(norm(p))) throw new Error('Kertas can only save to a file you opened or chose');
}

function filesFromArgv(argv) {
  return argv.slice(app.isPackaged ? 1 : 2)
    .filter((a) => !a.startsWith('-') && SUPPORTED.includes(path.extname(a).toLowerCase()));
}

// Only documents Kertas edits go in Recent, not pictures or certificates picked along the way
const RECENT_TYPES = /\.(docx|docm|dotx|doc|rtf|odt|xlsx|xlsm|xltx|xls|xlsb|ods|csv|pdf)$/i;

async function readRecent() {
  try { return JSON.parse(await fs.readFile(RECENT_FILE(), 'utf8')).filter((r) => RECENT_TYPES.test(r.path || '')); } catch { return []; }
}

// Pinned files stay in the list however many others are opened after them
const trimRecent = (list) => {
  let unpinned = 0;
  return list.filter((r) => r.pinned || ++unpinned <= MAX_RECENT);
};
async function writeRecent(list) {
  await fs.writeFile(RECENT_FILE(), JSON.stringify(trimRecent(list), null, 2));
}

async function addRecent(filePath) {
  if (!RECENT_TYPES.test(filePath)) return;
  const all = await readRecent();
  const was = all.find((r) => r.path === filePath);
  const list = all.filter((r) => r.path !== filePath);
  list.unshift({ path: filePath, name: path.basename(filePath), openedAt: Date.now(), ...(was?.pinned ? { pinned: true, pinnedAt: was.pinnedAt } : {}) });
  await writeRecent(list);
  app.addRecentDocument(filePath);
}

async function pinRecent(filePath, pinned) {
  const list = await readRecent();
  const r = list.find((x) => x.path === filePath);
  if (!r) return false;
  if (pinned) { r.pinned = true; r.pinnedAt = Date.now(); } else { delete r.pinned; delete r.pinnedAt; }
  await writeRecent(list);
  return true;
}

async function removeRecent(filePath) {
  const list = await readRecent();
  const left = list.filter((r) => r.path !== filePath);
  if (left.length === list.length) return false;
  await writeRecent(left);
  return true;
}

async function loadFile(filePath) {
  assertReadable(filePath);
  const data = await fs.readFile(filePath);
  grant(filePath);
  await addRecent(filePath);
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(data) };
}

// A file the person picked in an Open dialog shown by this process. The dialog itself limits
// the types (pictures, certificates, documents), so any of them may be read; only documents
// are granted for saving and listed in Recent.
async function loadPicked(filePath) {
  if (typeof filePath !== 'string' || !path.isAbsolute(filePath)) throw new Error('This file cannot be opened');
  const data = await fs.readFile(filePath);
  if (RECENT_TYPES.test(filePath)) { grant(filePath); await addRecent(filePath); }
  return { path: filePath, name: path.basename(filePath), data: new Uint8Array(data) };
}

function sendOpen(filePath) {
  if (!mainWindow || !rendererReady) { pendingOpen.push(filePath); return; }
  mainWindow.webContents.send('kertas:open-path', filePath);
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.focus();
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    title: 'Kertas Office',
    backgroundColor: '#eef1f7',
    ...(process.platform === 'linux' ? { icon: path.join(__dirname, 'icon.png') } : {}),
    // Frameless glass title bar: our tab strip is the title bar
    titleBarStyle: 'hidden',
    ...(process.platform === 'darwin'
      ? { trafficLightPosition: { x: 16, y: 15 }, vibrancy: 'under-window', visualEffectState: 'active' }
      : { titleBarOverlay: { color: '#00000000', symbolColor: '#334155', height: 44 } }),
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webSecurity: true,
      spellcheck: false, // Kertas has its own offline spell checker; Chromium's can download dictionaries
    },
  });

  if (DEV_URL) mainWindow.loadURL(DEV_URL);
  else mainWindow.loadURL('app://kertas/index.html');

  rendererReady = false;
  mainWindow.webContents.on('did-finish-load', () => {
    rendererReady = true;
    pendingOpen.splice(0).forEach(sendOpen);
  });

  // Never let a dropped file or link navigate the app window away
  mainWindow.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith('app://') && !(DEV_URL && url.startsWith(DEV_URL))) e.preventDefault();
  });

  // External links open in the system browser, never inside the app
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });

  // Let the renderer veto closing when there are unsaved tabs
  forceClose = false;
  mainWindow.on('close', (e) => {
    if (forceClose) return;
    e.preventDefault();
    mainWindow.webContents.send('kertas:request-close');
  });
  ipcMain.removeAllListeners('kertas:confirm-close');
  ipcMain.on('kertas:confirm-close', () => { forceClose = true; mainWindow.close(); });
  mainWindow.on('closed', () => { mainWindow = null; });
}

function buildMenu() {
  const send = (ch) => () => mainWindow && mainWindow.webContents.send('kertas:menu', ch);
  // Shortcuts are handled in the renderer so they work the same with or without a menu bar
  const item = (label, key, ch) => ({ label, accelerator: key, registerAccelerator: false, click: send(ch) });
  const isMac = process.platform === 'darwin';
  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        item('New Document', 'CmdOrCtrl+N', 'new-doc'),
        item('New Spreadsheet', 'CmdOrCtrl+Shift+N', 'new-sheet'),
        item('New Document from Template...', null, 'templates-writer'),
        item('New Spreadsheet from Template...', null, 'templates-sheets'),
        { type: 'separator' },
        item('Open...', 'CmdOrCtrl+O', 'open'),
        item('Save', 'CmdOrCtrl+S', 'save'),
        item('Save As...', 'CmdOrCtrl+Shift+S', 'save-as'),
        item('Export as PDF...', 'CmdOrCtrl+Shift+E', 'export-pdf'),
        item('Print...', 'CmdOrCtrl+P', 'print'),
        { type: 'separator' },
        item('Close Tab', 'CmdOrCtrl+W', 'close-tab'),
        ...(isMac ? [] : [{ type: 'separator' }, { role: 'quit' }]),
      ],
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
    { label: 'Help', submenu: [{ label: 'Kertas Office Website', click: () => shell.openExternal('https://kertasoffice.com') }, { label: 'Report a Problem...', click: () => shell.openExternal('https://github.com/kertas-office/kertas-office/issues/new/choose') }, { type: 'separator' }, { label: 'Check for Updates...', click: send('check-updates') }, { label: 'About Kertas Office', click: send('about') }] },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------- IPC ----------
ipcMain.on('kertas:theme', (_e, dark) => {
  if (!mainWindow || process.platform === 'darwin') return;
  try { mainWindow.setTitleBarOverlay({ color: '#00000000', symbolColor: dark ? '#e8ecf5' : '#334155', height: 44 }); } catch { /* not supported */ }
  mainWindow.setBackgroundColor(dark ? '#0b0f1a' : '#eef1f7');
});

// Renders the window with the print stylesheet into PDF bytes (used by Writer)
ipcMain.handle('kertas:print-to-pdf', async (_e, opts = {}) => {
  const data = await mainWindow.webContents.printToPDF({
    printBackground: true,
    preferCSSPageSize: true,
    margins: { marginType: 'none' },
    ...opts,
  });
  return new Uint8Array(data);
});

// Opens a PDF in Chromium's built-in viewer, which has print, zoom and download
ipcMain.handle('kertas:preview-pdf', async (_e, { data, name }) => {
  const dir = path.join(os.tmpdir(), 'kertas-preview');
  await fs.mkdir(dir, { recursive: true });
  const safe = String(name || 'preview.pdf').replace(/[^\w.\- ]+/g, '_');
  const file = path.join(dir, `${Date.now()}-${safe}`);
  await fs.writeFile(file, Buffer.from(data));
  const win = new BrowserWindow({
    width: 1000, height: 900, title: `Print preview - ${name}`, parent: mainWindow,
    autoHideMenuBar: true, webPreferences: { plugins: true, sandbox: true, contextIsolation: true },
  });
  // The PDF may come from someone else: its links open in the browser, never in this window
  win.webContents.setWindowOpenHandler(({ url }) => { if (/^https?:/.test(url)) shell.openExternal(url); return { action: 'deny' }; });
  win.webContents.on('will-navigate', (e, url) => { if (url !== win.webContents.getURL()) e.preventDefault(); });
  win.loadURL(pathToFileURL(file).toString());
  win.on('closed', () => fs.unlink(file).catch(() => {}));
  return true;
});

const FILTERS = {
  all: [{ name: 'Office files', extensions: ['docx', 'xlsx', 'csv', 'pdf'] }],
  writer: [{ name: 'Word Document', extensions: ['docx'] }],
  sheets: [{ name: 'Excel Workbook', extensions: ['xlsx'] }, { name: 'CSV (current sheet)', extensions: ['csv'] }],
  pdf: [{ name: 'PDF Document', extensions: ['pdf'] }],
  image: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg'] }],
  cert: [{ name: 'Certificates', extensions: ['p12', 'pfx'] }],
  any: [{ name: 'All files', extensions: ['*'] }],
};
// Opening also accepts older and OpenDocument formats (they are saved as .docx / .xlsx)
const OPEN_FILTERS = {
  all: [{ name: 'All supported files', extensions: ['docx', 'doc', 'rtf', 'odt', 'xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv', 'pdf'] }, ...FILTERS.writer, ...FILTERS.sheets.slice(0, 1), ...FILTERS.pdf],
  writer: [{ name: 'Documents', extensions: ['docx', 'doc', 'rtf', 'odt'] }],
  sheets: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv'] }],
};

ipcMain.handle('kertas:open-dialog', async (_e, { kind = 'all', multi = true } = {}) => {
  const res = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', ...(multi ? ['multiSelections'] : [])],
    filters: OPEN_FILTERS[kind] || FILTERS[kind] || OPEN_FILTERS.all,
  });
  if (res.canceled) return [];
  return Promise.all(res.filePaths.map(loadPicked));
});

ipcMain.handle('kertas:read-path', async (_e, filePath) => loadFile(filePath));

// A file attached to a PDF, saved where the person chooses in a Save dialog shown here
// (any type, as in Acrobat: the dialog is the person's decision, not the page's)
ipcMain.handle('kertas:save-attachment', async (_e, { name, data }) => {
  const base = path.basename(String(name || 'attachment')).replace(/[<>:"|?*\x00-\x1f]/g, '_') || 'attachment';
  const res = await dialog.showSaveDialog(mainWindow, { defaultPath: base });
  if (res.canceled || !res.filePath) return null;
  await fs.writeFile(res.filePath, Buffer.from(data));
  return { path: res.filePath, name: path.basename(res.filePath) };
});

// .doc / .xls / .rtf / .odt / .ods -> .docx / .xlsx (LibreOffice when installed)
ipcMain.handle('kertas:convert-legacy', async (_e, { name, data }) => convertLegacy({ name, data }));
ipcMain.handle('kertas:has-libreoffice', () => !!findSoffice());

ipcMain.handle('kertas:save-dialog', async (_e, { kind, defaultName }) => {
  const res = await dialog.showSaveDialog(mainWindow, {
    defaultPath: defaultName,
    filters: FILTERS[kind] || FILTERS.all,
  });
  return res.canceled ? null : grant(res.filePath);
});

ipcMain.handle('kertas:write-file', async (_e, { path: filePath, data }) => {
  assertWritable(filePath);
  // Write to a temp file first so a crash mid-save never corrupts the original
  const tmp = `${filePath}.kertas-tmp`;
  await fs.writeFile(tmp, Buffer.from(data));
  await fs.rename(tmp, filePath);
  await addRecent(filePath);
  return { path: filePath, name: path.basename(filePath) };
});

// Writes several files into one chosen folder without overwriting existing ones
ipcMain.handle('kertas:save-many', async (_e, files) => {
  const res = await dialog.showOpenDialog(mainWindow, { title: 'Choose a folder for the files', properties: ['openDirectory', 'createDirectory'] });
  if (res.canceled || !res.filePaths[0]) return null;
  const folder = res.filePaths[0];
  for (const f of files) {
    // a bare file name of a type Kertas makes, so nothing lands outside the chosen folder
    const name = path.basename(String(f.name || '')).replace(/[<>:"|?*\x00-\x1f]/g, '_');
    if (!name || !SAVE_MANY_TYPES.test(name)) throw new Error('Unexpected file type');
    const ext = path.extname(name);
    const base = path.basename(name, ext);
    let target = path.join(folder, name);
    for (let n = 1; ; n++) {
      try { await fs.access(target); target = path.join(folder, `${base} (${n})${ext}`); } catch { break; }
    }
    await fs.writeFile(target, Buffer.from(f.data));
  }
  shell.openPath(folder);
  return { folder, count: files.length };
});

// Clipboard and selection commands from ribbon buttons
const EDIT_ACTIONS = new Set(['cut', 'copy', 'paste', 'pasteAndMatchStyle', 'selectAll', 'undo', 'redo']);
ipcMain.handle('kertas:edit', (e, action) => {
  if (!EDIT_ACTIONS.has(action)) throw new Error('Unknown edit action');
  e.sender[action]();
  return true;
});

// Crash recovery: a copy of each unsaved document, written every minute
const RECOVERY_DIR = () => path.join(app.getPath('userData'), 'recovery');
const safeId = (id) => String(id).replace(/[^\w-]/g, '').slice(0, 64);
ipcMain.handle('kertas:recovery', async (_e, op, args = {}) => {
  const dir = RECOVERY_DIR();
  await fs.mkdir(dir, { recursive: true });
  const id = args.id ? safeId(args.id) : null;
  if (op === 'write') {
    await fs.writeFile(path.join(dir, `${id}.bin.tmp`), Buffer.from(args.data));
    await fs.rename(path.join(dir, `${id}.bin.tmp`), path.join(dir, `${id}.bin`));
    await fs.writeFile(path.join(dir, `${id}.json`), JSON.stringify({ ...args.meta, id }));
    return true;
  }
  if (op === 'delete') {
    await fs.rm(path.join(dir, `${id}.bin`), { force: true });
    await fs.rm(path.join(dir, `${id}.json`), { force: true });
    return true;
  }
  if (op === 'read') return new Uint8Array(await fs.readFile(path.join(dir, `${id}.bin`)));
  if (op === 'list') {
    const out = [];
    for (const f of await fs.readdir(dir)) {
      if (!f.endsWith('.json')) continue;
      try {
        const meta = JSON.parse(await fs.readFile(path.join(dir, f), 'utf8'));
        if (meta.path) grant(meta.path); // a recovered document can be saved back where it was
        out.push(meta);
      } catch { /* skip broken entries */ }
    }
    return out.sort((a, b) => b.savedAt - a.savedAt);
  }
  throw new Error('Unknown recovery operation');
});

// ------------------------------------------------------------ default app
// Windows and macOS keep the choice of default app with the user, so Kertas takes them to the
// place where it is made; on Linux the desktop's own tool sets it directly.
const DEFAULT_MIMES = [
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/msword', 'application/rtf',
  'application/vnd.oasis.opendocument.text', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel', 'application/vnd.oasis.opendocument.spreadsheet', 'text/csv', 'application/pdf',
];
async function makeDefaultApp() {
  if (!app.isPackaged) return { how: 'dev' };
  if (process.platform === 'win32') {
    // the app's own page in Settings > Default apps (Windows 11), or the Default apps page
    try { await shell.openExternal('ms-settings:defaultapps?registeredAppUser=Kertas%20Office'); return { how: 'settings' }; } catch { /* older Windows */ }
    try { await shell.openExternal('ms-settings:defaultapps'); return { how: 'settings' }; } catch { return { how: 'manual' }; }
  }
  if (process.platform === 'darwin') return { how: 'mac' };
  // Linux: the .deb installs kertas-office.desktop; an AppImage has no desktop entry of its own
  const { execFile } = require('child_process');
  const desktop = 'kertas-office.desktop';
  const dirs = ['/usr/share/applications', path.join(os.homedir(), '.local/share/applications')];
  let found = false;
  for (const d of dirs) { try { await fs.access(path.join(d, desktop)); found = true; break; } catch { /* not here */ } }
  if (!found) return { how: 'appimage' };
  const ok = await new Promise((resolve) => execFile('xdg-mime', ['default', desktop, ...DEFAULT_MIMES], { timeout: 15000 }, (err) => resolve(!err)));
  return { how: ok ? 'done' : 'manual' };
}
ipcMain.handle('kertas:default-app', () => makeDefaultApp());

ipcMain.handle('kertas:pin-recent', (_e, { path: p, pinned }) => pinRecent(p, !!pinned));
ipcMain.handle('kertas:remove-recent', (_e, p) => removeRecent(String(p || '')));
ipcMain.handle('kertas:recent', async () => {
  const list = await readRecent();
  const checked = await Promise.all(list.map(async (r) => {
    try { await fs.access(r.path); return r; } catch { return r.pinned ? { ...r, missing: true } : null; }
  }));
  return checked.filter(Boolean);
});

ipcMain.handle('kertas:confirm', async (_e, { message, detail, buttons }) => {
  const res = await dialog.showMessageBox(mainWindow, {
    type: 'question', message, detail, buttons, defaultId: 0, cancelId: buttons.length - 1,
  });
  return res.response;
});

ipcMain.handle('kertas:show-in-folder', (_e, filePath) => { if (granted.has(norm(filePath))) shell.showItemInFolder(filePath); });

// ---------- Lifecycle ----------
if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => filesFromArgv(argv).forEach(sendOpen));
  app.on('open-file', (e, filePath) => { e.preventDefault(); sendOpen(filePath); }); // macOS
  // No page gets camera, microphone, location, notifications or similar. The app's own
  // pages may use the clipboard (Paste buttons); nothing else is ever granted.
  const trusted = (url) => String(url || '').startsWith('app://kertas/') || (DEV_URL && String(url).startsWith(DEV_URL));
  const CLIPBOARD = new Set(['clipboard-read', 'clipboard-sanitized-write']);
  app.on('web-contents-created', (_e, contents) => {
    contents.on('will-attach-webview', (ev) => ev.preventDefault());
  });
  app.whenReady().then(() => {
    session.defaultSession.setPermissionRequestHandler((wc, permission, cb) => cb(CLIPBOARD.has(permission) && trusted(wc.getURL())));
    session.defaultSession.setPermissionCheckHandler((wc, permission) => CLIPBOARD.has(permission) && trusted(wc?.getURL()));
    session.defaultSession.setDevicePermissionHandler(() => false);
    registerAppProtocol();
    buildMenu();
    createWindow();
    setupUpdates({
      send: (st) => { if (mainWindow && rendererReady) mainWindow.webContents.send('kertas:update-state', st); },
      beforeInstall: () => { forceClose = true; },
    });
    filesFromArgv(process.argv).forEach(sendOpen);
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
  });
  app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
}
