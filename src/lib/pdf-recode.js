// Re-saving a PDF's pictures smaller (Compress): a picture is scaled to fit maxSide and
// saved as JPEG at the given quality. Used by pdf-compress.js through buildPdf.
export async function recodeImage({ data, kind, width, height, channels, maxSide, quality }) {
  const s = Math.min(1, maxSide / Math.max(width, height));
  const w = Math.max(1, Math.round(width * s)); const h = Math.max(1, Math.round(height * s));
  let source;
  if (kind === 'jpeg') source = await createImageBitmap(new Blob([data], { type: 'image/jpeg' }));
  else {
    const rgba = new Uint8ClampedArray(width * height * 4);
    for (let i = 0, j = 0; i < width * height; i++, j += channels) {
      const r = data[j]; const g = channels === 3 ? data[j + 1] : r; const b = channels === 3 ? data[j + 2] : r;
      rgba[i * 4] = r; rgba[i * 4 + 1] = g; rgba[i * 4 + 2] = b; rgba[i * 4 + 3] = 255;
    }
    source = document.createElement('canvas');
    source.width = width; source.height = height;
    source.getContext('2d').putImageData(new ImageData(rgba, width, height), 0, 0);
  }
  const c = document.createElement('canvas');
  c.width = w; c.height = h;
  const ctx = c.getContext('2d');
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(source, 0, 0, w, h);
  source.close?.();
  const blob = await new Promise((r) => c.toBlob(r, 'image/jpeg', quality));
  return blob ? { bytes: new Uint8Array(await blob.arrayBuffer()), width: w, height: h } : null;
}
