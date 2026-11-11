import { defineConfig } from 'vite';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';

const pkg = JSON.parse(readFileSync(new URL('./package.json', import.meta.url), 'utf8'));

// OCR (tesseract.js) loads its worker, engine and language files by URL at run
// time, so they are shipped under fixed paths in tess/ rather than bundled.
export const OCR_LANGS = ['eng', 'msa', 'mal', 'ara', 'hin', 'tam', 'chi_sim'];
const require = createRequire(import.meta.url);
// The engine version tesseract.js asks for (npm may nest it under tesseract.js)
const CORE_DIR = dirname(require.resolve('tesseract.js-core/package.json', { paths: [dirname(require.resolve('tesseract.js/package.json'))] }));
const nm = (p) => (p.startsWith('core:') ? join(CORE_DIR, p.slice(5)) : fileURLToPath(new URL(`./node_modules/${p}`, import.meta.url)));
const CORE = 'core:';
const OCR_FILES = {
  'tess/worker.min.js': 'tesseract.js/dist/worker.min.js',
  ...Object.fromEntries(['lstm', 'simd-lstm', 'relaxedsimd-lstm'].map((v) => [`tess/core/tesseract-core-${v}.wasm.js`, `${CORE}tesseract-core-${v}.wasm.js`])),
  ...Object.fromEntries(OCR_LANGS.map((l) => [`tess/lang/${l}.traineddata.gz`, `@tesseract.js-data/${l}/4.0.0_best_int/${l}.traineddata.gz`])),
};
function ocrAssets() {
  return {
    name: 'kertas-ocr-assets',
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        const key = decodeURIComponent((req.url || '').split('?')[0]).replace(/^\//, '');
        if (!OCR_FILES[key]) return next();
        res.setHeader('Content-Type', key.endsWith('.js') ? 'text/javascript' : 'application/octet-stream');
        res.end(readFileSync(nm(OCR_FILES[key])));
      });
    },
    generateBundle() {
      for (const [fileName, src] of Object.entries(OCR_FILES)) this.emitFile({ type: 'asset', fileName, source: readFileSync(nm(src)) });
    },
  };
}

export default defineConfig({
  plugins: [ocrAssets()],
  root: 'src',
  base: './', // relative paths so Electron can load dist/index.html from disk
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    target: 'es2022',
    chunkSizeWarningLimit: 8000,
    rollupOptions: {
      // Third-party React components ship "use client" banners that are harmless here
      onwarn(w, warn) { if (w.code === 'MODULE_LEVEL_DIRECTIVE') return; warn(w); },
    },
  },
  define: { __APP_VERSION__: JSON.stringify(pkg.version) },
  worker: { format: 'es' },
  server: { port: 5173, strictPort: true },
});
