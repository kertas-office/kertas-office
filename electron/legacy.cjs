// Opening older and non-Microsoft formats (.doc, .xls, .rtf, .odt, .ods).
// With LibreOffice installed we let it convert to .docx / .xlsx, which keeps
// the formatting. Without it, .doc files fall back to their text (word-extractor)
// and spreadsheets are read in the renderer (SheetJS).
const { execFile } = require('child_process');
const fs = require('fs/promises');
const fss = require('fs');
const os = require('os');
const path = require('path');

const TARGET = { doc: 'docx', rtf: 'docx', odt: 'docx', xls: 'xlsx', ods: 'xlsx', xlsb: 'xlsx' };

let cachedSoffice;
/** Finds LibreOffice's soffice binary, or null */
function findSoffice() {
  if (cachedSoffice !== undefined) return cachedSoffice;
  if (process.env.KERTAS_NO_LIBREOFFICE) return (cachedSoffice = null);
  const candidates = [];
  if (process.env.KERTAS_SOFFICE) candidates.push(process.env.KERTAS_SOFFICE);
  if (process.platform === 'win32') {
    for (const base of [process.env.ProgramFiles, process.env['ProgramFiles(x86)'], 'C:\\Program Files', 'C:\\Program Files (x86)']) {
      if (base) candidates.push(path.join(base, 'LibreOffice', 'program', 'soffice.exe'));
    }
  } else if (process.platform === 'darwin') {
    candidates.push('/Applications/LibreOffice.app/Contents/MacOS/soffice', path.join(os.homedir(), 'Applications/LibreOffice.app/Contents/MacOS/soffice'));
  } else {
    candidates.push('/usr/bin/soffice', '/usr/bin/libreoffice', '/usr/local/bin/soffice', '/opt/libreoffice/program/soffice', '/snap/bin/libreoffice');
    for (const dir of String(process.env.PATH || '').split(path.delimiter)) if (dir) candidates.push(path.join(dir, 'soffice'));
  }
  cachedSoffice = candidates.find((c) => { try { return fss.statSync(c).isFile(); } catch { return false; } }) || null;
  return cachedSoffice;
}

function run(bin, args, timeout) {
  return new Promise((resolve, reject) => {
    execFile(bin, args, { timeout, windowsHide: true }, (err, stdout, stderr) => (err ? reject(new Error(stderr || err.message)) : resolve(stdout)));
  });
}

async function viaLibreOffice(name, data) {
  const soffice = findSoffice();
  const ext = path.extname(name).slice(1).toLowerCase();
  const to = TARGET[ext];
  if (!soffice || !to) return null;
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'kertas-convert-'));
  try {
    const input = path.join(dir, `input.${ext}`);
    await fs.writeFile(input, Buffer.from(data));
    // A private profile, so a running LibreOffice window does not block us
    const profile = `file://${path.join(dir, 'profile').replace(/\\/g, '/').replace(/^([A-Za-z]):/, '/$1:')}`;
    await run(soffice, [`-env:UserInstallation=${profile}`, '--headless', '--norestore', '--nolockcheck', '--convert-to', to, '--outdir', dir, input], 90000);
    const out = await fs.readFile(path.join(dir, `input.${to}`));
    return { data: new Uint8Array(out), ext: to, via: 'libreoffice' };
  } finally {
    fs.rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

/** The text parts of a Word 97-2003 file */
async function docText(data) {
  const WordExtractor = require('word-extractor');
  const doc = await new WordExtractor().extract(Buffer.from(data));
  return {
    body: doc.getBody(),
    headers: doc.getHeaders({ includeFooters: false }),
    footers: doc.getFooters(),
    footnotes: doc.getFootnotes(),
    endnotes: doc.getEndnotes(),
    comments: doc.getAnnotations(),
  };
}

/**
 * { name, data } -> { data, ext, via } (converted file) | { text, via: 'text' } (.doc without LibreOffice)
 * | { via: 'none', reason } (the renderer handles it or says what is needed)
 */
async function convertLegacy({ name, data }) {
  const ext = path.extname(name).slice(1).toLowerCase();
  try {
    const converted = await viaLibreOffice(name, data);
    if (converted) return converted;
  } catch (e) {
    console.warn('LibreOffice conversion failed', e.message);
  }
  if (ext === 'doc') return { text: await docText(data), via: 'text' };
  return { via: 'none', reason: findSoffice() ? 'failed' : 'no-libreoffice' };
}

module.exports = { convertLegacy, findSoffice, docText };
