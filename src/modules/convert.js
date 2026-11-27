// Kertas Office - Convert: a list of files, each with the format it should become, converted
// one after another on this computer. Pictures, PDFs, documents and spreadsheets; the work
// itself is in ../lib/convert.
import { icon } from '../lib/icons.js';
import { openDialog } from '../lib/dialog.js';
import { GROUPS, LABELS, groupOf, outputsFor as allOutputs, defaultOutput, extOf, needsLibreOffice } from '../lib/convert/formats.js';
import { convertFile, folderOf, io } from '../lib/convert/run.js';

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const size = (n) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)} MB` : n >= 1024 ? `${Math.round(n / 1024)} KB` : `${n} B`);
const GROUP_ICON = { image: 'image', pdf: 'pdf', doc: 'doc', sheet: 'sheet', video: 'video', audio: 'audio' };
const KEY = 'kertas-convert-options';
const DEFAULTS = { quality: 90, maxSide: 0, dpi: 150, level: 'balanced', height: 0, where: 'same', folder: null };

function loadOptions() {
  try { return { ...DEFAULTS, ...JSON.parse(localStorage.getItem(KEY) || '{}'), folder: null, where: 'same' }; } catch { return { ...DEFAULTS }; }
}
function saveOptions(o) {
  try { localStorage.setItem(KEY, JSON.stringify({ quality: o.quality, maxSide: o.maxSide, dpi: o.dpi, level: o.level, height: o.height })); } catch { /* private window */ }
}

export async function mount(el, ctx) {
  const items = []; // { id, path, name, size, group, to, state: 'ready'|'working'|'done'|'error', progress, out: [], error }
  const opts = loadOptions();
  let running = false;
  let seq = 0;
  // what this computer can do: LibreOffice for .odt and .rtf, the FFmpeg engine for audio and video
  const [hasLo, engineState] = await Promise.all([io.hasLibreOffice().catch(() => false), window.kertasNative.convert('media-status').catch(() => ({ supported: false }))]);
  const engine = { ...engineState, progress: null, error: null };
  const outputsFor = (name) => allOutputs(name).filter((o) => hasLo || !needsLibreOffice(o));
  let current = null; // the item being converted, so Stop can end it

  el.classList.add('cv-view');
  el.innerHTML = `
    <div class="cv">
      <header class="cv-head">
        <span class="cv-mark">${icon('convert', 22)}</span>
        <div><h1>Convert</h1><p>Pictures, PDFs, documents, spreadsheets, audio and video, converted on this computer. Nothing is uploaded.</p></div>
      </header>
      <div class="cv-drop">
        <div class="cv-drop-ic">${icon('files', 26)}</div>
        <strong>Drop files here</strong>
        <span>or</span>
        <button class="chip primary" data-act="pick">${icon('open', 14)}Choose files</button>
        <small>Pictures (JPG, PNG, WEBP, AVIF, HEIC, TIFF, GIF, SVG) · PDF · Documents (DOCX, DOC, RTF, ODT, Markdown, HTML, TXT) · Spreadsheets (XLSX, XLS, ODS, CSV, JSON) · Video (MP4, MOV, MKV, WEBM, AVI) · Audio (MP3, M4A, WAV, FLAC, OGG)</small>
      </div>
      <div class="cv-engine" hidden></div>
      <div class="cv-list" hidden></div>
      <div class="cv-tools" hidden></div>
      <footer class="cv-bar" hidden>
        <label class="cv-all">Convert all to <select class="rb-select" data-act="all"></select></label>
        <button class="chip" data-act="where">${icon('open', 14)}<span class="cv-where"></span></button>
        <button class="chip" data-act="options">${icon('settings', 14)}Options</button>
        <span class="cv-gap"></span>
        <button class="chip" data-act="clear">Clear list</button>
        <button class="chip primary cv-go" data-act="go">${icon('convert', 14)}<span>Convert</span></button>
      </footer>
    </div>`;
  const $ = (s) => el.querySelector(s);
  const list = $('.cv-list');
  const drop = $('.cv-drop');

  async function add(found) {
    let added = 0;
    for (const f of found) {
      const group = groupOf(f.name);
      if (!group || items.some((i) => i.path === f.path)) continue;
      const to = outputsFor(f.name).includes(defaultOutput(f.name)) ? defaultOutput(f.name) : outputsFor(f.name)[0];
      items.push({ id: `c${++seq}`, path: f.path, name: f.name, size: f.size, group, to, state: 'ready', out: [] });
      added++;
    }
    if (!added && found.length) ctx.toast('Those files are already in the list, or Kertas cannot convert them');
    render();
  }

  // files dropped anywhere on the tab: their paths, without reading them
  el.dataset.fileDrop = 'paths';
  el.kertasFileDrop = async (records) => {
    const paths = records.map((r) => r.path).filter(Boolean);
    if (!paths.length) { ctx.toast('Drop files from a folder on this computer', true); return; }
    await add(await window.kertasNative.convert('add', { paths }));
  };

  const rowHtml = (it) => {
    const outs = outputsFor(it.name);
    let status = '';
    if (it.state === 'working') status = `<span class="cv-bar-track"><span style="width:${Math.round((it.progress || 0) * 100)}%"></span></span><span class="cv-state">Converting${it.progress ? ` ${Math.round(it.progress * 100)}%` : '...'}</span>`;
    else if (it.state === 'done') status = `<span class="cv-state ok">${icon('check', 13)}${it.out.length > 1 ? `${it.out.length} files` : esc(it.out[0]?.name || 'Done')}</span><button class="cv-link" data-act="open" data-id="${it.id}">Open</button><button class="cv-link" data-act="reveal" data-id="${it.id}">Show in folder</button>`;
    else if (it.state === 'error') status = `<span class="cv-state err" title="${esc(it.error)}">${esc(it.error)}</span>`;
    return `<div class="cv-row ${it.state}" data-id="${it.id}">
      <span class="cv-kind ${it.group}">${icon(GROUP_ICON[it.group], 15)}</span>
      <span class="cv-name"><strong title="${esc(it.path)}">${esc(it.name)}</strong><small>${esc(GROUPS[it.group].label)} · ${size(it.size)}</small></span>
      <span class="cv-arrow">${icon('right', 14)}</span>
      <select class="rb-select cv-to" data-act="to" data-id="${it.id}" ${it.state === 'working' || running ? 'disabled' : ''}>
        ${outs.map((o) => `<option value="${o}" ${o === it.to ? 'selected' : ''}>${esc(LABELS[o] || o)}</option>`).join('')}
      </select>
      <span class="cv-status">${status}</span>
      <button class="cv-x" data-act="remove" data-id="${it.id}" title="Remove from the list" ${it.state === 'working' ? 'disabled' : ''}>${icon('close', 13)}</button>
    </div>`;
  };

  function render() {
    const any = items.length > 0;
    list.hidden = !any; $('.cv-bar').hidden = !any; $('.cv-tools').hidden = !any;
    drop.classList.toggle('compact', any);
    list.innerHTML = items.map(rowHtml).join('');
    // "Convert all to": the outputs every file in the list can become
    const common = items.length ? items.map((i) => outputsFor(i.name)).reduce((a, b) => a.filter((o) => b.includes(o))) : [];
    const allSel = $('[data-act="all"]');
    const same = items.length && items.every((i) => i.to === items[0].to) ? items[0].to : '';
    allSel.innerHTML = `<option value="">${same ? '' : 'Mixed'}</option>${common.map((o) => `<option value="${o}" ${o === same ? 'selected' : ''}>${esc(LABELS[o] || o)}</option>`).join('')}`;
    if (same) allSel.querySelector('option[value=""]').remove();
    allSel.disabled = running || !common.length;
    $('.cv-all').title = common.length ? '' : 'These files have no format in common';
    const media = items.some((i) => GROUPS[i.group].media);
    const banner = $('.cv-engine');
    banner.hidden = !media || engine.installed;
    if (!banner.hidden) {
      const mb = Math.round((engine.size || 0) / 1048576);
      banner.innerHTML = !engine.supported
        ? `<span>${icon('lock', 15)}Audio and video conversion is not available on this kind of computer yet.</span>`
        : engine.progress !== null
          ? `<span>Downloading the audio and video engine... ${Math.round(engine.progress * 100)}%</span><span class="cv-bar-track"><span style="width:${Math.round(engine.progress * 100)}%"></span></span>`
          : `<span>Audio and video use FFmpeg, a free engine Kertas downloads once${mb ? ` (about ${mb} MB)` : ''}. It then works offline like everything else.${engine.error ? ` <b>${esc(engine.error)}</b>` : ''}</span><button class="chip primary" data-act="engine">${icon('export', 14)}Download</button>`;
    }
    $('.cv-where').textContent = opts.where === 'folder' && opts.folder ? `Save to ${opts.folder.split(/[\\/]/).pop()}` : 'Save next to each file';
    const todo = items.filter((i) => i.state !== 'done').length;
    const go = $('.cv-go');
    go.disabled = !todo && !running;
    go.querySelector('span').textContent = running ? 'Stop' : todo ? `Convert ${todo} file${todo === 1 ? '' : 's'}` : 'All converted';
    // tools that make one file out of several
    const pdfs = items.filter((i) => i.group === 'pdf');
    const pics = items.filter((i) => i.group === 'image');
    $('.cv-tools').innerHTML = [
      pdfs.length > 1 ? `<button class="chip" data-act="merge">${icon('merge', 14)}Merge ${pdfs.length} PDFs into one</button>` : '',
      pics.length > 1 ? `<button class="chip" data-act="combine">${icon('file-image', 14)}Put ${pics.length} pictures in one PDF</button>` : '',
    ].join('');
    $('.cv-tools').hidden = !pdfs.length && !pics.length || (pdfs.length < 2 && pics.length < 2);
  }

  const outDir = (it) => (opts.where === 'folder' && opts.folder ? opts.folder : folderOf(it.path));

  async function convertItem(it) {
    if (GROUPS[it.group].media && !engine.installed) {
      it.state = 'error'; it.error = 'Download the audio and video engine first';
      render();
      return false;
    }
    it.state = 'working'; it.progress = 0; it.error = null;
    current = it;
    render();
    try {
      it.out = await convertFile(it, it.to, {
        dir: outDir(it), quality: opts.quality / 100, maxSide: opts.maxSide, dpi: opts.dpi, level: opts.level, height: opts.height, job: it.id,
        onProgress: (p) => { it.progress = p; updateRow(it); },
      });
      it.state = 'done';
    } catch (e) {
      console.error(e);
      it.state = 'error'; it.error = String(e?.message || e).replace(/^Error invoking remote method '[^']+': (Error: )?/, '');
    }
    current = null;
    render();
    return it.state === 'done';
  }

  async function runAll() {
    running = true;
    stop = false;
    render();
    let ok = 0; let failed = 0;
    for (const it of items) {
      if (stop) break;
      if (it.state === 'done') continue;
      if (await convertItem(it)) ok++; else failed++;
    }
    running = false;
    render();
    if (ok || failed) ctx.toast(failed ? `Converted ${ok} file${ok === 1 ? '' : 's'}, ${failed} could not be converted` : `Converted ${ok} file${ok === 1 ? '' : 's'}`, failed > 0 && !ok);
  }
  let stop = false;
  function updateRow(it) {
    const row = list.querySelector(`.cv-row[data-id="${it.id}"]`);
    if (row) row.outerHTML = rowHtml(it);
  }
  // pages done in PDF to Word or Excel, reported by the converter as it goes
  window.kertasNative?.onConvert?.((ev) => {
    if (typeof ev.engine === 'number') { engine.progress = ev.engine; render(); return; }
    const it = items.find((i) => i.id === ev.job);
    if (!it || it.state !== 'working') return;
    if (ev.total) it.progress = ev.page / ev.total;
    else if (typeof ev.progress === 'number') it.progress = ev.progress;
    updateRow(it);
  });

  async function downloadEngine() {
    engine.progress = 0; engine.error = null;
    render();
    try {
      Object.assign(engine, await window.kertasNative.convert('media-install'));
      ctx.toast('The audio and video engine is ready');
      for (const it of items) if (GROUPS[it.group].media && it.state === 'error') { it.state = 'ready'; it.error = null; }
    } catch (e) {
      engine.error = String(e?.message || e).replace(/^Error invoking remote method '[^']+': (Error: )?/, '');
    }
    engine.progress = null;
    render();
  }

  async function oneOfMany(kind) {
    const group = kind === 'merge' ? 'pdf' : 'image';
    const picked = items.filter((i) => i.group === group);
    if (picked.length < 2) return;
    const first = picked[0];
    const name = `${first.name.replace(/\.[^.]+$/, '')} (${kind === 'merge' ? 'merged' : 'pictures'}).pdf`;
    ctx.setStatus(kind === 'merge' ? 'Merging PDFs...' : 'Making the PDF...');
    try {
      const pdf = await import('../lib/convert/pdf.js');
      let data;
      if (kind === 'merge') data = await pdf.mergePdfs(await Promise.all(picked.map((i) => io.read(i.path))));
      else {
        const { decodeImage, encodeImage } = await import('../lib/convert/image.js');
        const images = [];
        for (const i of picked) {
          const bytes = await io.read(i.path);
          const jpg = extOf(i.name) === 'jpg' || extOf(i.name) === 'jpeg';
          images.push(jpg && !opts.maxSide ? { bytes, fmt: 'jpg' } : { bytes: await encodeImage(await decodeImage(bytes, i.name, { maxSide: opts.maxSide }), 'jpg', { quality: opts.quality / 100 }), fmt: 'jpg' });
        }
        data = await pdf.imagesToPdf(images, { pageSize: 'a4' });
      }
      const out = await io.write(outDir(first), name, data);
      ctx.toast(`Saved ${out.name}`);
      await window.kertasNative.convert('reveal', { path: out.path });
    } catch (e) {
      ctx.toast(`Could not make the PDF: ${e.message || e}`, true);
    }
    ctx.setStatus('');
  }

  async function options() {
    const res = await openDialog({
      title: 'Convert options',
      fields: [
        { id: 'quality', type: 'range', label: 'Picture quality (JPG, WEBP, AVIF)', value: opts.quality, min: 40, max: 100, step: 5 },
        { id: 'maxSide', type: 'select', label: 'Picture size', value: opts.maxSide, options: [
          { value: 0, label: 'Keep the original size' }, { value: 3840, label: 'At most 3840 pixels' }, { value: 2560, label: 'At most 2560 pixels' },
          { value: 1920, label: 'At most 1920 pixels' }, { value: 1280, label: 'At most 1280 pixels' }, { value: 800, label: 'At most 800 pixels' },
        ] },
        { id: 'dpi', type: 'segmented', label: 'PDF pages as pictures', value: String(opts.dpi), options: [{ value: '72', label: 'Screen (72 dpi)' }, { value: '150', label: 'Standard (150 dpi)' }, { value: '300', label: 'Print (300 dpi)' }] },
        { id: 'height', type: 'select', label: 'Video size', value: opts.height, options: [
          { value: 0, label: 'Keep the original size' }, { value: 1080, label: 'At most 1080p' }, { value: 720, label: 'At most 720p' }, { value: 480, label: 'At most 480p' },
        ] },
        { id: 'level', type: 'segmented', label: 'Smaller PDF', value: opts.level, options: [{ value: 'strong', label: 'Smallest file' }, { value: 'balanced', label: 'Balanced' }, { value: 'light', label: 'Best quality' }] },
      ],
      actions: [{ id: 'cancel', label: 'Cancel' }, { id: 'ok', label: 'Done', primary: true }],
    });
    if (!res || res.action !== 'ok') return;
    opts.quality = Number(res.values.quality) || 90;
    opts.maxSide = Number(res.values.maxSide) || 0;
    opts.dpi = Number(res.values.dpi) || 150;
    opts.level = res.values.level || 'balanced';
    opts.height = Number(res.values.height) || 0;
    saveOptions(opts);
  }

  async function where() {
    if (opts.where === 'folder') { opts.where = 'same'; opts.folder = null; render(); return; }
    const folder = await window.kertasNative.convert('folder');
    if (folder) { opts.where = 'folder'; opts.folder = folder; }
    render();
  }

  el.addEventListener('click', async (e) => {
    const b = e.target.closest('[data-act]');
    if (!b || b.tagName === 'SELECT') return;
    const it = items.find((i) => i.id === b.dataset.id);
    switch (b.dataset.act) {
      case 'pick': await add(await window.kertasNative.convert('pick')); break;
      case 'go':
        if (running) { stop = true; if (current && GROUPS[current.group].media) window.kertasNative.convert('cancel', { job: current.id }); } else runAll();
        break;
      case 'engine': await downloadEngine(); break;
      case 'clear': if (!running) { items.length = 0; render(); } break;
      case 'remove': if (it && it.state !== 'working') { items.splice(items.indexOf(it), 1); render(); } break;
      case 'open': if (it?.out[0]) await window.kertasNative.convert('open', { path: it.out[0].path }); break;
      case 'reveal': if (it?.out[0]) await window.kertasNative.convert('reveal', { path: it.out[0].path }); break;
      case 'options': await options(); break;
      case 'where': await where(); break;
      case 'merge': await oneOfMany('merge'); break;
      case 'combine': await oneOfMany('combine'); break;
      default:
    }
  });
  el.addEventListener('change', (e) => {
    const s = e.target.closest('select[data-act]');
    if (!s) return;
    if (s.dataset.act === 'to') {
      const it = items.find((i) => i.id === s.dataset.id);
      if (it) { it.to = s.value; if (it.state !== 'working') { it.state = 'ready'; it.out = []; } render(); }
    } else if (s.dataset.act === 'all' && s.value) {
      for (const it of items) if (it.state !== 'working' && outputsFor(it.name).includes(s.value)) { it.to = s.value; it.state = 'ready'; it.out = []; }
      render();
    }
  });

  render();
  ctx.setStatus('Drop files or choose them to convert');
  return {
    add: (found) => add(found),
    convertOne: (id) => { const it = items.find((i) => i.id === id); return it ? convertItem(it) : false; },
    isBusy: () => running,
    setTheme() {},
    destroy() { stop = true; },
  };
}
