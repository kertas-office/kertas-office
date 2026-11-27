// One file, one conversion: reads it, turns it into the format asked for with the engine that
// fits, and writes the result (or results: a page per picture, a sheet per CSV) next to the
// original or into the folder chosen. Returns [{ path, name }].
import { GROUPS, groupOf, extOf, baseOf, canonical } from './formats.js';

const native = () => window.kertasNative;
export const io = {
  read: (path) => native().convert('read', { path }),
  write: (dir, name, data) => native().convert('write', { dir, name, data }),
  pdfToOffice: (path, format, job) => native().convert('pdf-to-office', { path, format, job }),
  docxToPdf: (arg) => native().convert('docx-to-pdf', arg),
  htmlToPdf: (html, opts = {}) => native().convert('html-to-pdf', { html, ...opts }),
  libreoffice: (path, to) => native().convert('libreoffice', { path, to }),
  hasLibreOffice: () => native().convert('has-libreoffice'),
  convertLegacy: (name, data) => native().convertLegacy({ name, data }),
  media: (arg) => native().convert('media', arg),
};

export const folderOf = (p) => p.replace(/[\\/][^\\/]*$/, '');

/**
 * item: { path, name }; to: an output id; opts: { dir, quality (0..1), maxSide, dpi, level, onProgress }
 */
export async function convertFile(item, to, opts = {}) {
  const dir = opts.dir || folderOf(item.path);
  const base = baseOf(item.name);
  const ext = canonical(extOf(item.name));
  const group = groupOf(item.name);
  const progress = opts.onProgress || (() => {});
  const written = [];
  const put = async (name, data) => { written.push(await io.write(dir, name, data)); };
  if (!GROUPS[group]) throw new Error('Kertas cannot convert this kind of file');

  if (group === 'image') {
    const [{ decodeImage, encodeImage }, bytes] = await Promise.all([import('./image.js'), io.read(item.path)]);
    const canvas = await decodeImage(bytes, item.name, { maxSide: opts.maxSide || 0 });
    if (to === 'pdf') {
      const { imagesToPdf } = await import('./pdf.js');
      // photos stay JPG inside the PDF; pictures with transparency stay PNG
      const keepJpg = ext === 'jpg' && !opts.maxSide;
      const fmt = keepJpg || ext === 'heic' ? 'jpg' : 'png';
      const img = keepJpg ? bytes : await encodeImage(canvas, fmt, { quality: opts.quality ?? 0.9 });
      await put(`${base}.pdf`, await imagesToPdf([{ bytes: img, fmt }], { pageSize: 'fit' }));
    } else {
      await put(`${base}.${to}`, await encodeImage(canvas, to, { quality: opts.quality ?? 0.9 }));
    }
    return written;
  }

  if (group === 'pdf') {
    if (to === 'docx' || to === 'xlsx') {
      await put(`${base}.${to}`, await io.pdfToOffice(item.path, to, opts.job));
      return written;
    }
    const [pdf, bytes] = await Promise.all([import('./pdf.js'), io.read(item.path)]);
    if (to === 'png' || to === 'jpg') {
      const pages = await pdf.pdfToImages(bytes, to, { dpi: opts.dpi || 150, quality: opts.quality ?? 0.9, onPage: (n, total) => progress(n / total) });
      for (const [i, data] of pages.entries()) await put(pages.length > 1 ? `${base} - page ${i + 1}.${to}` : `${base}.${to}`, data);
    } else if (to === 'txt') {
      await put(`${base}.txt`, new TextEncoder().encode(await pdf.pdfToText(bytes)));
    } else if (to === 'pages') {
      const parts = await pdf.splitPdf(bytes);
      for (const [i, data] of parts.entries()) await put(`${base} - page ${i + 1}.pdf`, data);
    } else if (to === 'compressed') {
      await put(`${base} (compressed).pdf`, await pdf.compressPdf(bytes, opts.level || 'balanced'));
    } else throw new Error(`PDFs cannot be converted to ${to}`);
    return written;
  }

  if (GROUPS[group].media) {
    written.push(await io.media({ path: item.path, to, dir, name: `${base}.${to}`, height: opts.height || 0, job: opts.job }));
    return written;
  }
  if (group === 'doc') return convertDocument(item, to, { dir, base, ext, put, written });
  if (group === 'sheet') return convertSheet(item, to, { dir, base, ext, put, written });
  return written;
}

async function convertDocument(item, to, { dir, base, ext, put, written }) {
  const doc = await import('./doc.js');
  const enc = new TextEncoder();
  const hasLo = ['doc', 'rtf', 'odt'].includes(ext) || to === 'odt' || to === 'rtf' ? await io.hasLibreOffice() : false;

  // older formats: LibreOffice converts them straight to what is asked, when it is installed
  if (['doc', 'rtf', 'odt'].includes(ext) || to === 'odt' || to === 'rtf') {
    if (hasLo && to !== 'md') {
      await put(`${base}.${to}`, await io.libreoffice(item.path, to === 'html' ? 'html' : to));
      return written;
    }
    if (to === 'odt' || to === 'rtf') throw new Error(`Saving as .${to} needs LibreOffice, which is not installed`);
  }

  // the source as Word bytes (docx, or an older format turned into one), or as HTML
  let docx = null; let html = null;
  const bytes = await io.read(item.path);
  if (ext === 'docx') docx = bytes;
  else if (['doc', 'rtf', 'odt'].includes(ext)) {
    if (hasLo) docx = await io.libreoffice(item.path, 'docx');
    else {
      const res = await io.convertLegacy(item.name, bytes);
      if (res?.data) docx = new Uint8Array(res.data);
      else if (res?.text) {
        const { docTextToHtml } = await import('../legacy.js');
        html = docTextToHtml(res.text);
      } else throw new Error(`Opening .${ext} files needs LibreOffice, which is not installed`);
    }
  } else if (ext === 'md') html = await doc.markdownToHtml(doc.textOf(bytes));
  else if (ext === 'html') html = doc.textOf(bytes);
  else if (ext === 'txt') html = doc.textToHtml(doc.textOf(bytes));

  if (docx) {
    if (to === 'pdf') { written.push(await io.docxToPdf(ext === 'docx' ? { path: item.path, dir, name: `${base}.pdf` } : { data: docx, dir, name: `${base}.pdf` })); return written; }
    if (to === 'docx') { await put(`${base}.docx`, docx); return written; }
    if (to === 'txt') { await put(`${base}.txt`, enc.encode(await doc.docxToText(docx))); return written; }
    html = await doc.docxToHtml(docx);
  }
  const page = /<html[\s>]/i.test(html) ? html : doc.htmlPage(html, base);
  if (to === 'pdf') await put(`${base}.pdf`, await io.htmlToPdf(page));
  else if (to === 'docx') await put(`${base}.docx`, await doc.htmlToDocx(page, base));
  else if (to === 'html') await put(`${base}.html`, enc.encode(page));
  else if (to === 'md') await put(`${base}.md`, enc.encode(await doc.htmlToMarkdown(page)));
  else if (to === 'txt') await put(`${base}.txt`, enc.encode(doc.htmlToText(page)));
  else throw new Error(`Documents cannot be converted to ${to}`);
  return written;
}

async function convertSheet(item, to, { base, ext, put, written }) {
  const sheet = await import('./sheet.js');
  // LibreOffice keeps the formatting of older workbooks; without it they are read as tables
  if (['xls', 'xlsb', 'ods'].includes(ext) && (to === 'xlsx' || to === 'ods' || to === 'pdf') && await io.hasLibreOffice()) {
    await put(`${base}.${to}`, await io.libreoffice(item.path, to));
    return written;
  }
  const wb = await sheet.readWorkbook(await io.read(item.path), item.name);
  if (to === 'pdf') {
    await put(`${base}.pdf`, await io.htmlToPdf(await sheet.workbookHtml(wb, base), { landscape: true }));
    return written;
  }
  for (const { suffix, data } of await sheet.writeWorkbook(wb, to)) await put(`${base}${suffix}.${to}`, data);
  return written;
}
