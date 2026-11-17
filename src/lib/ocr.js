// OCR for scanned pages: tesseract.js (Apache-2.0) with its engine and language
// files shipped in the app (tess/), so nothing is sent anywhere.
import { createWorker, OEM } from 'tesseract.js';

export const OCR_LANGUAGES = [
  { id: 'eng', label: 'English' },
  { id: 'msa', label: 'Malay' },
  { id: 'mal', label: 'Malayalam' },
  { id: 'ara', label: 'Arabic' },
  { id: 'hin', label: 'Hindi' },
  { id: 'tam', label: 'Tamil' },
  { id: 'chi_sim', label: 'Chinese (simplified)' },
];

// Several workers read pages side by side: one per spare processor core, up to four
// (each holds its own copy of the language data, around 100 MB).
let workers = [];
let workerKey = '';
let progressFns = [];
const abs = (p) => new URL(p, document.baseURI).href;

export function ocrWorkerCount(pages = Infinity) {
  const cores = Number(globalThis.navigator?.hardwareConcurrency) || 2;
  return Math.max(1, Math.min(4, cores - 1, pages));
}

async function makeWorker(langs, slot) {
  return createWorker(langs, OEM.LSTM_ONLY, {
    workerPath: abs('tess/worker.min.js'),
    corePath: abs('tess/core/'),
    langPath: abs('tess/lang'),
    gzip: true,
    workerBlobURL: false,
    logger: (m) => { if (m.status === 'recognizing text') progressFns[slot]?.(m.progress); },
  });
}

/** Ready workers for these languages (at least count of them) */
export async function ocrWorkers(langs, count = 1) {
  const key = langs.join('+');
  if (workerKey !== key) await stopOcr();
  workerKey = key;
  while (workers.length < count) {
    const slot = workers.length;
    workers.push(await makeWorker(langs, slot));
  }
  return workers.slice(0, count);
}

export async function stopOcr() {
  const old = workers;
  workers = []; workerKey = ''; progressFns = [];
  await Promise.all(old.map((w) => w.terminate().catch(() => {})));
}

function wordsOf(data) {
  const words = [];
  for (const block of data.blocks || []) {
    for (const para of block.paragraphs || []) {
      for (const line of para.lines || []) {
        const base = line.baseline ? (line.baseline.y0 + line.baseline.y1) / 2 : null;
        for (const word of line.words || []) {
          const text = (word.text || '').trim();
          if (!text || word.confidence < 25) continue;
          const b = word.bbox;
          words.push({ str: text, x0: b.x0, y0: b.y0, x1: b.x1, y1: b.y1, base: base ?? b.y1, conf: word.confidence });
        }
      }
    }
  }
  return words;
}

/**
 * Words from a rendered page image, in image pixels:
 * [{ str, x0, y0, x1, y1, base, conf }]  (base = baseline y)
 */
export async function recognise(image, langs, onProgress) {
  const [w] = await ocrWorkers(langs, 1);
  progressFns[0] = onProgress;
  const { data } = await w.recognize(image, {}, { blocks: true, text: false });
  return wordsOf(data);
}

/**
 * Reads many pages, several at a time. render(i) returns a canvas for item i (called
 * just before it is read, so only a few page images exist at once); onPage(i, words)
 * gets each result as it comes. Stops early when shouldStop() turns true.
 */
export async function recogniseMany(count, render, langs, { onPage, shouldStop } = {}) {
  const n = ocrWorkerCount(count);
  const pool = await ocrWorkers(langs, n);
  let next = 0;
  await Promise.all(pool.map(async (w) => {
    while (next < count && !shouldStop?.()) {
      const i = next++;
      let canvas = await render(i);
      try {
        const { data } = await w.recognize(canvas, {}, { blocks: true, text: false });
        onPage?.(i, wordsOf(data));
      } finally {
        canvas.width = 0; canvas.height = 0; canvas = null; // let the page image go
      }
    }
  }));
}

// Rendering scale for OCR: 300 DPI for ordinary pages, but never a page image bigger than
// about 3600 pixels on its long side. Scanner and phone apps often save a page at the
// photo's pixel size (a "page" of 2480 x 3508 points), and 300 DPI of that is a 150
// megapixel image that takes minutes to read.
export const OCR_MAX_SIDE = 3600;
export function ocrScale(width, height) {
  return Math.min(300 / 72, OCR_MAX_SIDE / Math.max(width, height, 1));
}
