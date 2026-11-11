// Makes PDFs smaller without changing how they look on screen:
//   - photos are re-saved as JPEG at a sensible size for the chosen level
//   - streams saved without compression are compressed
//   - objects nothing refers to any more (left over from earlier edits) are dropped
//   - page thumbnails and editing leftovers from other apps (PieceInfo) are removed
// Objects are packed into object streams when the document is saved.
//
// compressDocument(doc, { images, recodeImage })
//   images       'strong' | 'balanced' | 'light' | null (null leaves pictures alone)
//   recodeImage  async ({ data, kind: 'jpeg'|'raw', width, height, channels, maxSide, quality })
//                -> { bytes, width, height } as JPEG, or null to keep the picture.
//                The app passes a canvas-based one; pure Node callers can pass their own.
import { PDFName, PDFDict, PDFArray, PDFRef, PDFRawStream, PDFStream, PDFNumber, decodePDFRawStream } from '@cantoo/pdf-lib';

export const LEVELS = {
  strong: { maxSide: 1300, quality: 0.6, label: 'Smallest file' },
  balanced: { maxSide: 2200, quality: 0.75, label: 'Balanced' },
  light: { maxSide: 3400, quality: 0.86, label: 'Best quality' },
};
const MIN_IMAGE_BYTES = 24 * 1024; // smaller pictures are not worth the work

const N = (n) => PDFName.of(n);
const nameOf = (v) => (v instanceof PDFName ? v.decodeText() : null);

function filters(dict) {
  const f = dict.lookup(N('Filter'));
  if (!f) return [];
  if (f instanceof PDFArray) return f.asArray().map((x) => nameOf(x));
  return [nameOf(f)];
}

/** Colour channels for a colour space we can safely turn into JPEG, or 0 */
function channelsOf(doc, cs) {
  const v = cs instanceof PDFRef ? doc.context.lookup(cs) : cs;
  const n = nameOf(v);
  if (n === 'DeviceRGB' || n === 'CalRGB') return 3;
  if (n === 'DeviceGray' || n === 'CalGray') return 1;
  if (v instanceof PDFArray && nameOf(v.get(0)) === 'ICCBased') {
    const st = doc.context.lookup(v.get(1));
    const c = st?.dict?.lookup(N('N'));
    const k = c instanceof PDFNumber ? c.asNumber() : 0;
    return k === 3 || k === 1 ? k : 0;
  }
  return 0;
}

async function recompressImages(doc, level, recodeImage) {
  const { maxSide, quality } = LEVELS[level] || LEVELS.balanced;
  const objects = doc.context.enumerateIndirectObjects();
  // pictures used as another picture's transparency stay exact
  const masks = new Set();
  for (const [, obj] of objects) {
    const d = obj instanceof PDFStream ? obj.dict : null;
    const m = d?.get(N('SMask'));
    if (m instanceof PDFRef) masks.add(m.toString());
  }
  let saved = 0; let count = 0;
  for (const [ref, obj] of objects) {
    if (!(obj instanceof PDFRawStream)) continue;
    const d = obj.dict;
    if (nameOf(d.get(N('Subtype'))) !== 'Image' || masks.has(ref.toString())) continue;
    if (d.lookup(N('ImageMask'))?.asBoolean?.() || d.get(N('Decode'))) continue;
    if (obj.contents.length < MIN_IMAGE_BYTES) continue;
    const bpc = d.lookup(N('BitsPerComponent'))?.asNumber?.();
    const channels = channelsOf(doc, d.get(N('ColorSpace')));
    const fl = filters(d);
    const width = d.lookup(N('Width'))?.asNumber?.() || 0;
    const height = d.lookup(N('Height'))?.asNumber?.() || 0;
    if (!channels || bpc !== 8 || !width || !height) continue;
    let input;
    if (fl.length === 1 && fl[0] === 'DCTDecode') input = { data: obj.contents, kind: 'jpeg' };
    else if (fl.length === 1 && fl[0] === 'FlateDecode') {
      const parms = d.lookup(N('DecodeParms'));
      if (parms instanceof PDFDict && (parms.lookup(N('Predictor'))?.asNumber?.() || 1) > 1) continue;
      let raw;
      try { raw = decodePDFRawStream(obj).decode(); } catch { continue; }
      if (raw.length < width * height * channels) continue;
      input = { data: raw, kind: 'raw' };
    } else continue;
    // A JPEG that is already small for its size is left alone
    const already = input.kind === 'jpeg' && Math.max(width, height) <= maxSide && obj.contents.length / (width * height) < 0.18 * quality;
    if (already) continue;
    let res = null;
    try { res = await recodeImage({ ...input, width, height, channels, maxSide, quality }); } catch { res = null; }
    if (!res?.bytes || res.bytes.length >= obj.contents.length * 0.92) continue;
    const dict = doc.context.obj({
      Type: 'XObject', Subtype: 'Image', Width: res.width, Height: res.height,
      ColorSpace: 'DeviceRGB', BitsPerComponent: 8, Filter: 'DCTDecode', Length: res.bytes.length,
    });
    for (const keep of ['SMask', 'Interpolate', 'Intent', 'OC', 'Alternates', 'Metadata', 'Name', 'StructParent']) {
      const v = d.get(N(keep));
      if (v) dict.set(N(keep), v);
    }
    doc.context.assign(ref, PDFRawStream.of(dict, res.bytes));
    saved += obj.contents.length - res.bytes.length; count++;
  }
  return { saved, count };
}

/** Compresses content and other streams that were saved without compression */
function flateLoose(doc) {
  let n = 0;
  for (const [ref, obj] of doc.context.enumerateIndirectObjects()) {
    if (!(obj instanceof PDFRawStream) || obj.dict.get(N('Filter'))) continue;
    if (nameOf(obj.dict.get(N('Type'))) === 'XRef' || nameOf(obj.dict.get(N('Type'))) === 'Metadata') continue;
    if (obj.contents.length < 256) continue;
    const dict = PDFDict.withContext(doc.context);
    for (const [k, v] of obj.dict.entries()) if (k.decodeText() !== 'Length' && k.decodeText() !== 'DecodeParms') dict.set(k, v);
    const st = doc.context.flateStream(obj.contents, {});
    for (const [k, v] of dict.entries()) st.dict.set(k, v);
    doc.context.assign(ref, st);
    n++;
  }
  return n;
}

/** Thumbnails and other apps' private editing data are not needed to show the document */
function dropExtras(doc) {
  for (const page of doc.getPages()) {
    page.node.delete(N('Thumb'));
    page.node.delete(N('PieceInfo'));
  }
  doc.catalog.delete(N('PieceInfo'));
}

/** Removes objects that nothing in the document refers to */
export function dropUnused(doc) {
  const ctx = doc.context;
  const seen = new Set();
  const stack = [];
  const push = (v) => { if (v) stack.push(v); };
  const t = ctx.trailerInfo;
  push(t.Root); push(t.Info); push(t.Encrypt);
  while (stack.length) {
    const v = stack.pop();
    if (v instanceof PDFRef) {
      const key = v.toString();
      if (seen.has(key)) continue;
      seen.add(key);
      push(ctx.lookup(v));
    } else if (v instanceof PDFDict) {
      for (const [, x] of v.entries()) push(x);
    } else if (v instanceof PDFArray) {
      for (const x of v.asArray()) push(x);
    } else if (v instanceof PDFStream) {
      push(v.dict);
    }
  }
  let dropped = 0;
  for (const [ref] of ctx.enumerateIndirectObjects()) {
    if (!seen.has(ref.toString())) { ctx.delete(ref); dropped++; }
  }
  return dropped;
}

export async function compressDocument(doc, { images = 'balanced', recodeImage } = {}) {
  dropExtras(doc);
  const pics = images && recodeImage ? await recompressImages(doc, images, recodeImage) : { saved: 0, count: 0 };
  const streams = flateLoose(doc);
  const dropped = dropUnused(doc);
  return { images: pics.count, imageBytesSaved: pics.saved, streams, dropped };
}
