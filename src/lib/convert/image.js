// Pictures: read with the browser's own decoders where it has them (JPG, PNG, WEBP, AVIF, GIF,
// BMP, ICO, SVG), with small libraries for TIFF and HEIC; written by the canvas (PNG, JPG,
// WEBP) or by our own writers (BMP, ICO, TIFF, GIF, AVIF).
import { canonical, extOf } from './formats.js';

const MIME = { jpg: 'image/jpeg', png: 'image/png', webp: 'image/webp', avif: 'image/avif', gif: 'image/gif', bmp: 'image/bmp', ico: 'image/x-icon', svg: 'image/svg+xml' };

/** A canvas with the picture, at most maxSide on its longest side (0: as it is) */
export async function decodeImage(bytes, name, { maxSide = 0 } = {}) {
  const ext = canonical(extOf(name));
  let source; let w; let h;
  if (ext === 'tiff') {
    const UTIF = (await import('utif')).default;
    const ifds = UTIF.decode(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
    const page = ifds[0];
    if (!page) throw new Error('This TIFF has no picture in it');
    UTIF.decodeImage(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength), page);
    const rgba = new Uint8ClampedArray(UTIF.toRGBA8(page));
    w = page.width; h = page.height;
    source = new ImageData(rgba, w, h);
  } else if (ext === 'heic') {
    const libheif = (await import('libheif-js/wasm-bundle')).default;
    const decoder = new libheif.HeifDecoder();
    const [image] = decoder.decode(bytes);
    if (!image) throw new Error('This HEIC file has no picture in it');
    w = image.get_width(); h = image.get_height();
    source = new ImageData(w, h);
    await new Promise((resolve, reject) => image.display(source, (out) => (out ? resolve() : reject(new Error('Could not read this HEIC picture')))));
  } else if (ext === 'svg') {
    const url = URL.createObjectURL(new Blob([bytes], { type: MIME.svg }));
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      w = img.naturalWidth || 1024; h = img.naturalHeight || 1024;
      // a drawing with no size of its own comes out at a useful size
      if (!img.naturalWidth) { w = 1024; h = 1024; }
      source = img;
    } finally { setTimeout(() => URL.revokeObjectURL(url), 1000); }
  } else {
    source = await createImageBitmap(new Blob([bytes], { type: MIME[ext] || '' }));
    w = source.width; h = source.height;
  }
  const scale = maxSide && Math.max(w, h) > maxSide ? maxSide / Math.max(w, h) : 1;
  const cw = Math.max(1, Math.round(w * scale)); const ch = Math.max(1, Math.round(h * scale));
  const canvas = new OffscreenCanvas(cw, ch);
  const ctx = canvas.getContext('2d');
  if (source instanceof ImageData) {
    if (scale === 1) ctx.putImageData(source, 0, 0);
    else ctx.drawImage(await createImageBitmap(source), 0, 0, cw, ch);
  } else {
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(source, 0, 0, cw, ch);
  }
  source.close?.();
  return canvas;
}

const blobBytes = async (blob) => new Uint8Array(await blob.arrayBuffer());
function pixels(canvas) {
  return canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
}
/** The same picture on white, for formats with no transparency */
function onWhite(canvas) {
  const out = new OffscreenCanvas(canvas.width, canvas.height);
  const ctx = out.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, out.width, out.height);
  ctx.drawImage(canvas, 0, 0);
  return out;
}
function resized(canvas, side) {
  const s = side / Math.max(canvas.width, canvas.height);
  const out = new OffscreenCanvas(Math.max(1, Math.round(canvas.width * s)), Math.max(1, Math.round(canvas.height * s)));
  const ctx = out.getContext('2d');
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(canvas, 0, 0, out.width, out.height);
  return out;
}

/** 24-bit BMP, bottom-up rows padded to 4 bytes */
function bmp(canvas) {
  const { data, width: w, height: h } = pixels(onWhite(canvas));
  const row = Math.ceil((w * 3) / 4) * 4;
  const size = 54 + row * h;
  const out = new Uint8Array(size);
  const v = new DataView(out.buffer);
  out[0] = 0x42; out[1] = 0x4d;
  v.setUint32(2, size, true); v.setUint32(10, 54, true); v.setUint32(14, 40, true);
  v.setInt32(18, w, true); v.setInt32(22, h, true); v.setUint16(26, 1, true); v.setUint16(28, 24, true);
  v.setUint32(34, row * h, true); v.setInt32(38, 2835, true); v.setInt32(42, 2835, true);
  for (let y = 0; y < h; y++) {
    const src = (h - 1 - y) * w * 4;
    let o = 54 + y * row;
    for (let x = 0; x < w; x++) {
      const i = src + x * 4;
      out[o++] = data[i + 2]; out[o++] = data[i + 1]; out[o++] = data[i];
    }
  }
  return out;
}

/** An icon with PNG pictures at the usual sizes, up to the picture's own size */
async function ico(canvas) {
  const sizes = [16, 24, 32, 48, 64, 128, 256].filter((s) => s <= Math.max(canvas.width, canvas.height, 16));
  const images = [];
  for (const s of sizes) {
    // square, the picture centred in it
    const sq = new OffscreenCanvas(s, s);
    const k = s / Math.max(canvas.width, canvas.height);
    const dw = Math.round(canvas.width * k); const dh = Math.round(canvas.height * k);
    const ctx = sq.getContext('2d');
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(canvas, Math.round((s - dw) / 2), Math.round((s - dh) / 2), dw, dh);
    images.push({ s, data: await blobBytes(await sq.convertToBlob({ type: 'image/png' })) });
  }
  const head = 6 + images.length * 16;
  const total = head + images.reduce((n, i) => n + i.data.length, 0);
  const out = new Uint8Array(total);
  const v = new DataView(out.buffer);
  v.setUint16(2, 1, true); v.setUint16(4, images.length, true);
  let offset = head;
  images.forEach((img, n) => {
    const e = 6 + n * 16;
    out[e] = img.s >= 256 ? 0 : img.s; out[e + 1] = img.s >= 256 ? 0 : img.s;
    v.setUint16(e + 4, 1, true); v.setUint16(e + 6, 32, true);
    v.setUint32(e + 8, img.data.length, true); v.setUint32(e + 12, offset, true);
    out.set(img.data, offset);
    offset += img.data.length;
  });
  return out;
}

async function gif(canvas) {
  const { GIFEncoder, quantize, applyPalette } = await import('gifenc');
  const { data, width, height } = pixels(canvas);
  const palette = quantize(data, 256, { format: 'rgba4444', oneBitAlpha: true });
  const index = applyPalette(data, palette, 'rgba4444');
  const enc = GIFEncoder();
  const transparent = palette.findIndex((c) => c[3] === 0);
  enc.writeFrame(index, width, height, { palette, ...(transparent >= 0 ? { transparent: true, transparentIndex: transparent } : {}) });
  enc.finish();
  return enc.bytes();
}

async function tiff(canvas) {
  const UTIF = (await import('utif')).default;
  const { data, width, height } = pixels(canvas);
  return new Uint8Array(UTIF.encodeImage(data.buffer, width, height));
}

async function avif(canvas, quality) {
  const { default: encode } = await import('@jsquash/avif/encode.js');
  const buf = await encode(pixels(canvas), { quality: Math.round(quality * 100), speed: 6 });
  return new Uint8Array(buf);
}

/**
 * The picture in canvas as `to` bytes. quality 0..1 for JPG, WEBP and AVIF.
 */
export async function encodeImage(canvas, to, { quality = 0.9 } = {}) {
  switch (to) {
    case 'png': return blobBytes(await canvas.convertToBlob({ type: 'image/png' }));
    case 'jpg': return blobBytes(await onWhite(canvas).convertToBlob({ type: 'image/jpeg', quality }));
    case 'webp': return blobBytes(await canvas.convertToBlob({ type: 'image/webp', quality }));
    case 'avif': return avif(canvas, quality);
    case 'bmp': return bmp(canvas);
    case 'ico': return ico(Math.max(canvas.width, canvas.height) > 256 ? resized(canvas, 256) : canvas);
    case 'gif': return gif(canvas);
    case 'tiff': return tiff(canvas);
    default: throw new Error(`Pictures cannot be saved as ${to}`);
  }
}
