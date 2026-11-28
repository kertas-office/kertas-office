// Kertas Office - the Convert app's side in the main process: the files it may read (picked
// or dropped), the folders it may write into, and the conversions that need this process
// (PDF to Word or Excel, Word to PDF through the Writer, web pages to PDF, older formats
// through LibreOffice). Everything else converts in the page itself.
const { ipcMain, dialog, shell, BrowserWindow } = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs/promises');
const fss = require('fs');
const { execFile } = require('child_process');
const { pathToFileURL } = require('url');
const ffmpeg = require('./ffmpeg.cjs');

// What the converter reads, and what it writes
const MEDIA = 'mp4|mov|mkv|webm|avi|wmv|flv|mpe?g|m4v|3gp|ts|mts|m2ts|ogv|mp3|m4a|aac|wav|flac|ogg|oga|opus|wma|aiff?|amr';
const INPUT = new RegExp(`\\.(jpe?g|png|webp|avif|gif|bmp|ico|svg|tiff?|heic|heif|pdf|docx|docm|dotx|doc|rtf|odt|md|markdown|html?|txt|xlsx|xlsm|xls|xlsb|ods|csv|tsv|json|${MEDIA})$`, 'i');
const OUTPUT = /\.(jpe?g|png|webp|avif|gif|bmp|ico|tiff?|pdf|docx|odt|rtf|txt|html|md|xlsx|ods|csv|tsv|json|mp4|mov|mkv|webm|avi|mp3|m4a|wav|flac|ogg|opus)$/i;
const MAX_READ = 512 * 1024 * 1024;

const readable = new Set();
const folders = new Set();
const norm = (p) => path.resolve(p).toLowerCase();

function remember(file) {
  readable.add(norm(file));
  folders.add(norm(path.dirname(file)));
}
function assertInput(p) {
  if (typeof p !== 'string' || !path.isAbsolute(p) || !INPUT.test(p)) throw new Error('This kind of file cannot be converted');
}
function assertRead(p) {
  assertInput(p);
  if (!readable.has(norm(p))) throw new Error('Add the file to the list first');
}

async function describe(file) {
  const st = await fs.stat(file);
  if (!st.isFile()) throw new Error('Not a file');
  remember(file);
  return { path: file, name: path.basename(file), size: st.size };
}

/** A name in the folder that does not overwrite anything: "report.pdf", "report (1).pdf"... */
async function freeName(dir, name) {
  const ext = path.extname(name);
  const base = path.basename(name, ext);
  let target = path.join(dir, name);
  for (let n = 1; ; n++) {
    try { await fs.access(target); target = path.join(dir, `${base} (${n})${ext}`); } catch { return target; }
  }
}
function cleanName(name) {
  const n = path.basename(String(name || '')).replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').trim();
  if (!n || !OUTPUT.test(n)) throw new Error('Unexpected output type');
  return n;
}
function assertFolder(dir) {
  if (typeof dir !== 'string' || !path.isAbsolute(dir) || !folders.has(norm(dir))) throw new Error('Choose the folder first');
}
async function target(dir, name) {
  assertFolder(dir);
  return freeName(dir, cleanName(name));
}

/** An HTML page printed to PDF by Chromium, out of sight */
async function htmlToPdf(html, { landscape = false } = {}) {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'kertas-html-'));
  const file = path.join(dir, 'page.html');
  await fs.writeFile(file, String(html));
  // a session of its own, so blocking the internet here touches nothing else in the app
  const win = new BrowserWindow({ show: false, width: 1000, height: 1400, webPreferences: { sandbox: true, contextIsolation: true, javascript: false, partition: 'kertas-convert-html' } });
  try {
    // nothing in the page may reach the internet
    win.webContents.session.webRequest.onBeforeRequest({ urls: ['http://*/*', 'https://*/*'] }, (_d, cb) => cb({ cancel: true }));
    await win.loadURL(pathToFileURL(file).toString());
    const data = await win.webContents.printToPDF({ printBackground: true, landscape, pageSize: 'A4', margins: { marginType: 'default' } });
    return new Uint8Array(data);
  } finally {
    win.destroy();
    fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

/** LibreOffice's own conversion, for the formats only it reads well */
function viaLibreOffice(soffice, input, to) {
  return new Promise((resolve, reject) => {
    fs.mkdtemp(path.join(os.tmpdir(), 'kertas-lo-')).then((dir) => {
      const ext = path.extname(input).slice(1).toLowerCase();
      const copy = path.join(dir, `input.${ext}`);
      const profile = `file://${path.join(dir, 'profile').replace(/\\/g, '/').replace(/^([A-Za-z]):/, '/$1:')}`;
      fs.copyFile(input, copy).then(() => {
        execFile(soffice, [`-env:UserInstallation=${profile}`, '--headless', '--norestore', '--nolockcheck', '--convert-to', to, '--outdir', dir, copy], { timeout: 180000, windowsHide: true }, async (err, _out, stderr) => {
          try {
            if (err) throw new Error(stderr || err.message);
            resolve(new Uint8Array(await fs.readFile(path.join(dir, `input.${to.split(':')[0]}`))));
          } catch (e) { reject(e); } finally { fs.rm(dir, { recursive: true, force: true }).catch(() => {}); }
        });
      }, reject);
    }, reject);
  });
}

function register({ getWindow, writer2, findSoffice }) {
  const send = (msg) => { const w = getWindow(); if (w && !w.isDestroyed()) w.webContents.send('kertas:convert-event', msg); };

  ipcMain.handle('kertas:convert', async (_e, op, arg = {}) => {
    const win = getWindow();
    switch (op) {
      case 'pick': {
        const res = await dialog.showOpenDialog(win, {
          title: 'Choose files to convert',
          properties: ['openFile', 'multiSelections'],
          filters: [
            { name: 'Files Kertas can convert', extensions: ['jpg', 'jpeg', 'png', 'webp', 'avif', 'gif', 'bmp', 'ico', 'svg', 'tif', 'tiff', 'heic', 'heif', 'pdf', 'docx', 'doc', 'rtf', 'odt', 'md', 'html', 'htm', 'txt', 'xlsx', 'xlsm', 'xls', 'xlsb', 'ods', 'csv', 'tsv', 'json', ...MEDIA.replace(/\?/g, '').replace('mpeg', 'mpeg|mpg').replace('aiff', 'aiff|aif').split('|')] },
            { name: 'All files', extensions: ['*'] },
          ],
        });
        if (res.canceled) return [];
        const out = [];
        for (const p of res.filePaths) if (INPUT.test(p)) out.push(await describe(p));
        return out;
      }
      // files dropped on the app: the page passes the paths the drop gave it
      case 'add': {
        const out = [];
        for (const p of Array.isArray(arg.paths) ? arg.paths : []) {
          try { assertInput(p); out.push(await describe(p)); } catch { /* not a file Kertas converts */ }
        }
        return out;
      }
      case 'read': {
        assertRead(arg.path);
        const st = await fs.stat(arg.path);
        if (st.size > MAX_READ) throw new Error('This file is too large to convert here');
        return new Uint8Array(await fs.readFile(arg.path));
      }
      case 'folder': {
        const res = await dialog.showOpenDialog(win, { title: 'Choose where to save the converted files', properties: ['openDirectory', 'createDirectory'] });
        if (res.canceled || !res.filePaths[0]) return null;
        folders.add(norm(res.filePaths[0]));
        return res.filePaths[0];
      }
      case 'write': {
        const file = await target(arg.dir, arg.name);
        const tmp = `${file}.kertas-tmp`;
        await fs.writeFile(tmp, Buffer.from(arg.data));
        await fs.rename(tmp, file);
        return { path: file, name: path.basename(file) };
      }
      case 'open': {
        if (typeof arg.path !== 'string' || !folders.has(norm(path.dirname(arg.path)))) return false;
        const err = await shell.openPath(arg.path);
        return !err;
      }
      case 'reveal': {
        if (typeof arg.path === 'string' && folders.has(norm(path.dirname(arg.path)))) shell.showItemInFolder(arg.path);
        return true;
      }
      case 'has-libreoffice': return !!findSoffice();
      case 'pdf-to-office': {
        assertRead(arg.path);
        const format = arg.format === 'xlsx' ? 'xlsx' : 'docx';
        const bytes = await fs.readFile(arg.path);
        const out = await writer2.convertPdfForTab(bytes, format, { raw: true, onProgress: (page, total) => send({ job: arg.job, page, total }) });
        if (!out) throw new Error('The PDF converter is not available in this build');
        return new Uint8Array(out);
      }
      case 'docx-to-pdf': {
        const file = await target(arg.dir, arg.name);
        // a document converted from an older format arrives as bytes, and waits in a temp file
        let src = arg.path;
        let tmp = null;
        if (arg.data) {
          tmp = await fs.mkdtemp(path.join(os.tmpdir(), 'kertas-docx-'));
          src = path.join(tmp, `${path.basename(file, '.pdf')}.docx`);
          await fs.writeFile(src, Buffer.from(arg.data));
        } else assertRead(src);
        try {
          const out = await writer2.docxToPdf(src, file);
          return { path: out, name: path.basename(out) };
        } finally { if (tmp) fs.rm(tmp, { recursive: true, force: true }).catch(() => {}); }
      }
      case 'html-to-pdf': return htmlToPdf(arg.html, arg);
      case 'libreoffice': {
        assertRead(arg.path);
        const soffice = findSoffice();
        if (!soffice) throw new Error('This needs LibreOffice, which is not installed');
        const to = String(arg.to || '');
        if (!/^(pdf|docx|odt|rtf|xlsx|ods|csv|html|txt)$/.test(to)) throw new Error('Unexpected output type');
        return viaLibreOffice(soffice, arg.path, to === 'txt' ? 'txt:Text' : to);
      }
      case 'media-status': return ffmpeg.status();
      case 'media-install': return ffmpeg.install((p) => send({ engine: p }));
      case 'media': {
        assertRead(arg.path);
        const file = await target(arg.dir, arg.name);
        const out = await ffmpeg.convert(arg.path, file, String(arg.to), { height: Number(arg.height) || 0 }, (p) => send({ job: arg.job, progress: p }), arg.job);
        return { path: out, name: path.basename(out) };
      }
      case 'cancel': ffmpeg.cancel(arg.job); return true;
      default: throw new Error(`Unknown convert operation: ${op}`);
    }
  });
}

module.exports = { register, INPUT, OUTPUT };
// exposed for tests
module.exports._freeName = freeName;
module.exports._exists = (p) => fss.existsSync(p);
