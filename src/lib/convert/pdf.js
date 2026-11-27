// PDFs in the page: pages to pictures and text with pdf.js, and merge, split, compress and
// pictures-to-PDF with Kertas's own PDF tools.
import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

/** The document and a way to let it go (the loading task owns the worker side) */
async function load(bytes) {
  const task = pdfjsLib.getDocument({ data: bytes.slice(), isEvalSupported: false });
  const doc = await task.promise;
  doc.release = () => task.destroy();
  return doc;
}

/** Each page as a picture; dpi 72..300. onPage(n, total) after each one */
export async function pdfToImages(bytes, to, { dpi = 150, quality = 0.9, onPage } = {}) {
  const doc = await load(bytes);
  const out = [];
  try {
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const vp = page.getViewport({ scale: dpi / 72 });
      const canvas = new OffscreenCanvas(Math.ceil(vp.width), Math.ceil(vp.height));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: vp }).promise;
      const blob = await canvas.convertToBlob({ type: to === 'png' ? 'image/png' : 'image/jpeg', quality });
      out.push(new Uint8Array(await blob.arrayBuffer()));
      page.cleanup();
      onPage?.(n, doc.numPages);
    }
  } finally { await doc.release(); }
  return out;
}

/** The text of every page, in reading order, a blank line between pages */
export async function pdfToText(bytes) {
  const doc = await load(bytes);
  const pages = [];
  try {
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const { items } = await page.getTextContent();
      // lines by their baseline, left to right within a line
      const lines = [];
      for (const it of items) {
        if (!('str' in it)) continue;
        const y = Math.round(it.transform[5]);
        let line = lines.find((l) => Math.abs(l.y - y) <= Math.max(2, (it.height || 10) * 0.4));
        if (!line) { line = { y, parts: [] }; lines.push(line); }
        line.parts.push({ x: it.transform[4], s: it.str, end: it.transform[4] + (it.width || 0) });
        if (it.hasEOL) line.eol = true;
      }
      lines.sort((a, b) => b.y - a.y);
      pages.push(lines.map((l) => {
        l.parts.sort((a, b) => a.x - b.x);
        let s = '';
        let prevEnd = null;
        for (const p of l.parts) {
          if (prevEnd !== null && p.x - prevEnd > 1.5 && !s.endsWith(' ') && !p.s.startsWith(' ')) s += ' ';
          s += p.s;
          prevEnd = p.end;
        }
        return s.trimEnd();
      }).join('\n'));
      page.cleanup();
    }
  } finally { await doc.release(); }
  return `${pages.join('\n\n').trim()}\n`;
}

async function pageCount(bytes) {
  const { PDFDocument } = await import('@cantoo/pdf-lib');
  return (await PDFDocument.load(bytes, { updateMetadata: false, ignoreEncryption: true })).getPageCount();
}

/** All the PDFs as one, in the order given */
export async function mergePdfs(list) {
  const { buildPdf } = await import('../pdf-ops.js');
  const sources = {}; const pages = [];
  for (const [i, bytes] of list.entries()) {
    const id = `s${i}`;
    sources[id] = bytes;
    const count = await pageCount(bytes);
    for (let index = 0; index < count; index++) pages.push({ src: id, index });
  }
  return buildPdf(sources, pages);
}

/** One PDF per page */
export async function splitPdf(bytes) {
  const { buildPdf } = await import('../pdf-ops.js');
  const count = await pageCount(bytes);
  const out = [];
  for (let index = 0; index < count; index++) out.push(await buildPdf({ a: bytes }, [{ src: 'a', index }]));
  return out;
}

/** A smaller copy: pictures resized and re-saved at the level chosen (strong, balanced, light) */
export async function compressPdf(bytes, level = 'balanced') {
  const [{ buildPdf }, { recodeImage }] = await Promise.all([import('../pdf-ops.js'), import('../pdf-recode.js')]);
  const count = await pageCount(bytes);
  const pages = Array.from({ length: count }, (_, index) => ({ src: 'a', index }));
  return buildPdf({ a: bytes }, pages, { compress: { images: level, recodeImage } });
}

/** Pictures (PNG or JPG bytes) as pages of one PDF */
export async function imagesToPdf(images, opts) {
  const { imagesToPdf: build } = await import('../pdf-ops.js');
  return build(images, opts);
}
