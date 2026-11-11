// File access layer. Uses the Electron bridge when present, falls back to
// browser APIs so the same build also runs as a plain web app for testing.
const native = typeof window !== 'undefined' ? window.kertasNative : undefined;
export const isDesktop = !!native;
export const platform = native?.platform === 'darwin' ? 'mac' : native?.platform === 'win32' ? 'win' : native ? 'linux' : 'web';

const ACCEPT = {
  all: '.docx,.doc,.rtf,.odt,.xlsx,.xlsm,.xlsb,.xls,.ods,.csv,.pdf',
  writer: '.docx,.doc,.rtf,.odt',
  sheets: '.xlsx,.xlsm,.xlsb,.xls,.ods,.csv',
  pdf: '.pdf',
  image: '.png,.jpg,.jpeg',
  cert: '.p12,.pfx',
};

export function kindForName(name = '') {
  const ext = name.toLowerCase().split('.').pop();
  if (['docx', 'doc', 'rtf', 'odt'].includes(ext)) return 'writer';
  if (['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv'].includes(ext)) return 'sheets';
  if (ext === 'pdf') return 'pdf';
  return null;
}

export function extOf(name = '') {
  return name.toLowerCase().split('.').pop();
}

/** Desktop only: converts an older format with LibreOffice (or reads a .doc's text). See electron/legacy.cjs */
export async function convertLegacyNative(file) {
  if (!native?.convertLegacy) return null;
  return native.convertLegacy({ name: file.name, data: file.data });
}

/** Returns [{name, path|null, data: Uint8Array}] */
export async function openFiles(kind = 'all', multi = true) {
  if (native) return native.openDialog({ kind, multi });
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = ACCEPT[kind] || ACCEPT.all;
    input.multiple = multi;
    input.onchange = async () => {
      const files = await Promise.all([...input.files].map(async (f) => ({
        name: f.name, path: null, data: new Uint8Array(await f.arrayBuffer()),
      })));
      resolve(files);
    };
    input.oncancel = () => resolve([]);
    input.click();
  });
}

export async function readPath(p) {
  if (!native) throw new Error('Opening by path is only available in the desktop app');
  return native.readPath(p);
}

/**
 * Saves bytes. If `path` is given and saveAs is false, writes in place.
 * Returns {name, path} or null if cancelled.
 */
export async function saveFile({ data, path, name, kind, saveAs = false }) {
  if (native) {
    let target = path;
    if (!target || saveAs) target = await native.saveDialog({ kind, defaultName: name });
    if (!target) return null;
    return native.writeFile(target, data);
  }
  // Browser: trigger a download
  const blob = new Blob([data]);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  return { name, path: null };
}

export async function recentFiles() {
  if (native) return native.recent();
  return [];
}

/** Pins a recent file to Home, or unpins it */
export async function pinRecent(path, pinned) {
  if (native?.pinRecent) return native.pinRecent(path, pinned);
  return false;
}

/** Returns index of the chosen button */
export async function confirmDialog(message, detail, buttons) {
  if (native) return native.confirm({ message, detail, buttons });
  // Browser fallback: two-way only
  return window.confirm(`${message}\n\n${detail || ''}`) ? 0 : buttons.length - 1;
}

export function onNative(event, cb) {
  if (!native) return;
  if (event === 'open-path') native.onOpenPath(cb);
  if (event === 'menu') native.onMenu(cb);
  if (event === 'request-close') native.onRequestClose(cb);
  if (event === 'update' && native.onUpdate) native.onUpdate(cb);
}

/** Desktop updates: op 'state' | 'check' | 'install' | 'set-auto' */
export async function updates(op, arg) {
  if (!native?.update) return { state: 'web' };
  return native.update(op, arg);
}

export function confirmClose() {
  if (native) native.confirmClose();
}

/** Shows a PDF in a print-preview window (desktop) or a new browser tab */
export async function previewPdf(data, name) {
  if (native) return native.previewPdf(data, name);
  const url = URL.createObjectURL(new Blob([data], { type: 'application/pdf' }));
  window.open(url, '_blank');
  setTimeout(() => URL.revokeObjectURL(url), 60000);
  return null;
}

/**
 * Renders the current window to PDF using the print stylesheet.
 * Desktop only; returns null in a browser after opening the print dialog.
 */
export async function printToPdf(options = {}) {
  if (native) return native.printToPdf(options);
  window.print();
  return null;
}

export function setNativeTheme(dark) {
  native?.setTheme?.(dark);
}

/** Saves several files at once: one folder prompt on desktop, one download each in a browser */
export async function saveMany(files) {
  if (native) return native.saveMany(files.map((f) => ({ name: f.name, data: f.data })));
  for (const f of files) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([f.data]));
    a.download = f.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
    await new Promise((r) => setTimeout(r, 250)); // browsers drop rapid-fire downloads
  }
  return { folder: null, count: files.length };
}

/** Cut, copy, paste, paste as plain text, select all. Runs against the focused element. */
export async function editCommand(action) {
  if (native?.edit) return native.edit(action);
  if (action === 'pasteAndMatchStyle') {
    const text = await navigator.clipboard.readText();
    document.execCommand('insertText', false, text);
    return true;
  }
  if (action === 'paste') {
    // Browsers only allow programmatic paste of plain text
    const text = await navigator.clipboard.readText();
    document.execCommand('insertText', false, text);
    return true;
  }
  if (!document.execCommand(action)) throw new Error(`${action} is blocked by the browser`);
  return true;
}

// ---------------------------------------------------------------- crash recovery
// Desktop: files in the app's data folder. Browser: IndexedDB.
const IDB = 'kertas-recovery';
function idb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore('docs');
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function idbDo(mode, fn) {
  const db = await idb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('docs', mode);
    const out = fn(tx.objectStore('docs'));
    tx.oncomplete = () => resolve(out?.result ?? out);
    tx.onerror = () => reject(tx.error);
  });
}

export async function recoveryWrite(id, meta, data) {
  try {
    if (native?.recovery) return await native.recovery('write', { id, meta, data });
    return await idbDo('readwrite', (st) => st.put({ meta: { ...meta, id }, data }, id));
  } catch (e) { console.warn('recovery write failed', e); return null; }
}
export async function recoveryDelete(id) {
  try {
    if (native?.recovery) return await native.recovery('delete', { id });
    return await idbDo('readwrite', (st) => st.delete(id));
  } catch { return null; }
}
/** Returns [{ id, name, path, kind, savedAt }] */
export async function recoveryList() {
  try {
    if (native?.recovery) return await native.recovery('list', {});
    const rows = await idbDo('readonly', (st) => st.getAll());
    return (rows || []).map((r) => r.meta);
  } catch { return []; }
}
export async function recoveryRead(id) {
  if (native?.recovery) return native.recovery('read', { id });
  const row = await idbDo('readonly', (st) => st.get(id));
  return row ? new Uint8Array(row.data) : null;
}
