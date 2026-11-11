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

let worker = null;
let workerKey = '';
const abs = (p) => new URL(p, document.baseURI).href;

export async function ocrWorker(langs, onProgress) {
  const key = langs.join('+');
  if (worker && workerKey === key) return worker;
  if (worker) { await worker.terminate().catch(() => {}); worker = null; }
  worker = await createWorker(langs, OEM.LSTM_ONLY, {
    workerPath: abs('tess/worker.min.js'),
    corePath: abs('tess/core/'),
    langPath: abs('tess/lang'),
    gzip: true,
    workerBlobURL: false,
    logger: (m) => { if (m.status === 'recognizing text') onProgress?.(m.progress); },
  });
  workerKey = key;
  return worker;
}

export async function stopOcr() { if (worker) { await worker.terminate().catch(() => {}); worker = null; workerKey = ''; } }

/**
 * Words from a rendered page image, in image pixels:
 * [{ str, x0, y0, x1, y1, base, conf }]  (base = baseline y)
 */
export async function recognise(image, langs, onProgress) {
  const w = await ocrWorker(langs, onProgress);
  const { data } = await w.recognize(image, {}, { blocks: true, text: false });
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
