// Slash commands for Writer, in the style of Notion: type "/" at the start of a
// line (or after a space) and a menu of blocks and actions opens at the cursor.
// Keep typing to filter, arrows and Enter to choose, Esc to leave the text as typed.
// Ctrl+/ opens the same menu with its own search box, anywhere.
import { icon } from '../lib/icons.js';

const RECENT_KEY = 'kertas-slash-recent';
const HINT_KEY = 'kertas-slash-hint';
const store = {
  get(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* private mode */ } },
};
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Score of a command for a query: 0 = no match. Prefix matches on the label win. */
export function matchScore(cmd, query) {
  const q = String(query || '').trim().toLowerCase();
  if (!q) return 1;
  const label = cmd.label.toLowerCase();
  const words = [label, ...(cmd.keys || [])].map((w) => w.toLowerCase());
  if (label.startsWith(q)) return 100 - label.length / 100;
  if (words.some((w) => w === q)) return 90;
  if (words.some((w) => w.startsWith(q))) return 70;
  if (label.split(/\s+/).some((w) => w.startsWith(q))) return 60;
  if (words.some((w) => w.includes(q))) return 30;
  // Letters in order ("tbl" finds Table)
  let k = 0;
  for (const ch of label) if (ch === q[k]) k++;
  return k === q.length ? 10 : 0;
}

/** Commands for a query, best first; with no query, recent ones lead */
export function filterCommands(commands, query, recent = []) {
  const q = String(query || '').trim();
  if (!q) {
    const rec = recent.map((id) => commands.find((c) => c.id === id)).filter(Boolean).slice(0, 3);
    return [...rec.map((c) => ({ ...c, recent: true })), ...commands.filter((c) => !rec.includes(c))];
  }
  return commands.map((c) => ({ c, s: matchScore(c, q) })).filter((x) => x.s > 0).sort((a, b) => b.s - a.s).map((x) => x.c);
}

export function createSlashMenu({ root, api, superdoc, commands, toast, onSlash = true, caretTo }) {
  let menu = null; // { el, anchor: { blockId, start } | null, query, items, sel, input }
  let busy = false;

  // --------------------------------------------------------------- reading the document
  async function caret() {
    const s = await Promise.resolve().then(() => api().selection.current()).catch(() => null);
    const seg = s?.target?.segments?.[0];
    return seg ? { blockId: seg.blockId, offset: seg.range.end ?? seg.range.start, collapsed: seg.range.start === seg.range.end } : null;
  }
  async function blockText(blockId) {
    const r = await api().blocks.list({ nodeIds: [blockId], includeText: true });
    const b = (r?.blocks || r?.items || []).find((x) => x.nodeId === blockId);
    return b ? String(b.text ?? b.textPreview ?? '') : null;
  }

  /** Screen position of a text offset, from the painted page (spans carry editor positions) */
  function pointAt(blockId, offset) {
    const id = CSS.escape(String(blockId));
    const frags = [...root.querySelectorAll(`.superdoc-fragment[data-block-id="${id}"], .superdoc-fragment[data-source-node-id="${id}"]`)];
    if (!frags.length) return null;
    const base = Number(frags[0].dataset.pmStart);
    const target = base + offset;
    for (const f of frags) {
      for (const run of f.querySelectorAll('.superdoc-text-run')) {
        const a = Number(run.dataset.pmStart); const b = Number(run.dataset.pmEnd);
        if (!(target >= a && target <= b)) continue;
        const text = run.firstChild?.nodeType === 3 ? run.firstChild : null;
        if (!text) { const r = run.getBoundingClientRect(); return { x: r.left, top: r.top, bottom: r.bottom }; }
        const i = Math.max(0, Math.min(text.length, target - a));
        const range = document.createRange();
        range.setStart(text, i); range.setEnd(text, i);
        const r = range.getClientRects()[0] || run.getBoundingClientRect();
        return { x: r.left, top: r.top, bottom: r.bottom };
      }
    }
    // Empty paragraph or the end of the block: its last line
    const line = frags.at(-1).querySelector('.superdoc-line') || frags.at(-1);
    const r = line.getBoundingClientRect();
    return { x: r.left, top: r.top, bottom: r.bottom };
  }

  // --------------------------------------------------------------- the menu
  function render() {
    if (!menu) return;
    const recent = store.get(RECENT_KEY, []);
    menu.items = filterCommands(commands, menu.query, recent);
    menu.sel = Math.max(0, Math.min(menu.sel, menu.items.length - 1));
    const list = menu.el.querySelector('.slash-list');
    if (!menu.items.length) {
      list.innerHTML = `<div class="slash-empty">No command matches "${esc(menu.query)}"</div>`;
      return;
    }
    let group = null;
    list.innerHTML = menu.items.map((c, i) => {
      const g = c.recent && !menu.query ? 'Recent' : c.group;
      const head = !menu.query && g !== group ? `<div class="slash-group">${esc((group = g))}</div>` : '';
      return `${head}<button class="slash-item${i === menu.sel ? ' on' : ''}" data-i="${i}" type="button"><span class="slash-ic">${icon(c.icon || 'plus', 17)}</span><span class="slash-text"><b>${esc(c.label)}</b><small>${esc(c.hint || '')}</small></span>${c.shortcut ? `<kbd>${esc(c.shortcut)}</kbd>` : ''}</button>`;
    }).join('');
    list.querySelector('.slash-item.on')?.scrollIntoView({ block: 'nearest' });
  }

  function place(pt) {
    const el = menu.el;
    const w = el.offsetWidth || 320; const h = el.offsetHeight || 360;
    let x = pt ? pt.x - 6 : window.innerWidth / 2 - w / 2;
    let y = pt ? pt.bottom + 6 : window.innerHeight / 3;
    if (pt && y + h > window.innerHeight - 12) y = Math.max(12, pt.top - h - 6); // flip above near the bottom
    x = Math.max(12, Math.min(x, window.innerWidth - w - 12));
    el.style.left = `${x}px`; el.style.top = `${y}px`;
  }

  function open(anchor, pt) {
    close();
    hideHint();
    const el = document.createElement('div');
    el.className = 'slash-menu glass';
    el.setAttribute('role', 'listbox');
    el.innerHTML = `${anchor ? '' : `<div class="slash-search">${icon('search', 15)}<input placeholder="Search commands..." aria-label="Search commands"/></div>`}<div class="slash-list"></div><div class="slash-foot"><span>↑↓ to choose</span><span>Enter to insert</span><span>Esc to close</span></div>`;
    document.body.appendChild(el);
    menu = { el, anchor, query: '', items: [], sel: 0, input: el.querySelector('input') };
    render();
    place(pt);
    el.addEventListener('mousedown', (e) => e.preventDefault()); // keep the editor's cursor
    el.addEventListener('click', (e) => {
      const b = e.target.closest('.slash-item');
      if (b) run(menu.items[Number(b.dataset.i)]);
    });
    el.addEventListener('mousemove', (e) => {
      const b = e.target.closest('.slash-item');
      if (b && Number(b.dataset.i) !== menu.sel) { menu.sel = Number(b.dataset.i); for (const n of el.querySelectorAll('.slash-item')) n.classList.toggle('on', Number(n.dataset.i) === menu.sel); }
    });
    if (menu.input) {
      menu.input.focus();
      menu.input.addEventListener('input', () => { menu.query = menu.input.value; menu.sel = 0; render(); });
      menu.input.addEventListener('keydown', (e) => { if (navKey(e)) { e.preventDefault(); e.stopPropagation(); } });
      menu.input.addEventListener('blur', () => setTimeout(() => { if (menu && !menu.anchor && document.activeElement !== menu.input) close(); }, 120));
    }
  }

  function close() {
    if (!menu) return;
    menu.el.remove();
    menu = null;
  }

  /** Arrow keys, Enter, Tab and Esc while the menu is open; true when handled */
  function navKey(e) {
    if (!menu) return false;
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      if (!menu.items.length) return true;
      menu.sel = (menu.sel + (e.key === 'ArrowDown' ? 1 : -1) + menu.items.length) % menu.items.length;
      render();
      return true;
    }
    if (e.key === 'Enter' || e.key === 'Tab') {
      if (!menu.items.length) { close(); return false; }
      run(menu.items[menu.sel]);
      return true;
    }
    if (e.key === 'Escape') {
      close();
      superdoc()?.focus?.();
      return true;
    }
    return false;
  }

  async function run(cmd) {
    if (!cmd || busy) return;
    busy = true;
    const anchor = menu?.anchor;
    const query = menu?.query || '';
    close();
    try {
      // Take the "/query" text out first, so the command acts where it was typed
      if (anchor) {
        const end = await settledEnd(anchor, query);
        if (end !== null) {
          const P = (o) => ({ kind: 'text', blockId: anchor.blockId, offset: o });
          await api().replace({ target: { kind: 'selection', start: P(anchor.start), end: P(end) }, text: '' });
        }
      }
      superdoc()?.focus?.();
      store.set(RECENT_KEY, [cmd.id, ...store.get(RECENT_KEY, []).filter((x) => x !== cmd.id)].slice(0, 6));
      await cmd.run();
    } catch (e) {
      console.error(e);
      toast?.(`${cmd.label}: ${e.message || 'could not be done here'}`, true);
    } finally {
      busy = false;
    }
  }

  // The query is tracked from keystrokes: the document model catches up a moment
  // after typing, so reading it back per key would lag behind fast typists.
  function setQuery(q) {
    if (!menu) return;
    if (q.length > 32 || /^\s|\s{2}/.test(q)) { close(); return; }
    menu.query = q;
    menu.sel = 0;
    render();
    if (/\s$/.test(q) && !menu.items.length) close(); // a space after no match: the "/" was just text
  }
  let pending = null; // keys typed while the "/" is being checked
  function trackKey(e, apply) {
    if (e.key === 'Backspace') { if (!apply('')) close(); return; }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { apply(e.key); return; }
    if (['ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown', 'Delete'].includes(e.key)) close();
  }
  const applyToMenu = (ch) => {
    if (!menu?.anchor) return false;
    if (!ch) { if (!menu.query) return false; setQuery(menu.query.slice(0, -1)); return true; }
    setQuery(menu.query + ch);
    return true;
  };

  /** Waits until the model shows the "/" (it may trail the screen), then the end of the query */
  async function settledEnd(anchor, query) {
    for (let k = 0; k < 20; k++) {
      const text = await blockText(anchor.blockId);
      if (text && text[anchor.start] === '/' && text.slice(anchor.start + 1, anchor.start + 1 + query.length) === query) return anchor.start + 1 + query.length;
      await new Promise((r) => setTimeout(r, 50));
    }
    return null;
  }

  // --------------------------------------------------------------- keyboard
  function onKeyDown(e) {
    if (e.isComposing) return;
    if ((e.ctrlKey || e.metaKey) && e.key === '/') {
      e.preventDefault(); e.stopPropagation();
      if (menu) close(); else open(null, null);
      return;
    }
    if (menu?.anchor) {
      if (navKey(e)) { e.preventDefault(); e.stopPropagation(); return; }
      trackKey(e, applyToMenu);
      return;
    }
    const slashKey = e.key === '/' && !e.ctrlKey && !e.metaKey && !e.altKey;
    if (pending && !slashKey) {
      // Keys typed while an earlier "/" is still being checked belong to its query;
      // anything that is not typing (Enter, arrows...) means it was not a command
      // Fast "/h2" + Enter: hold the Enter and run the best match once the menu is up
      if (e.key === 'Enter' && pending.keys && filterCommands(commands, pending.keys).length) { e.preventDefault(); e.stopPropagation(); pending.enter = true; return; }
      if (e.key !== 'Backspace' && (e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey)) { pending = null; return; }
      if (e.key === 'Backspace') { if (!pending.keys) pending = null; else pending.keys = pending.keys.slice(0, -1); return; }
      pending.keys += e.key;
      return;
    }
    if (slashKey && onSlash) {
      // Type the "/" ourselves, so the editor's own "/" handling (which swallows it at the start
      // of a line) never runs, whether or not its patch in scripts/patch-deps.cjs is in place
      const view = superdoc()?.activeEditor?.view;
      if (view?.dom?.contains?.(e.target) && view.editable) {
        e.preventDefault(); e.stopPropagation();
        view.dispatch(view.state.tr.insertText('/').scrollIntoView());
      }
      checkSlash();
    }
  }
  async function checkSlash() {
    // The model may trail the screen: wait until it shows the "/" (and possibly some
    // of what was typed after it), then judge from the text before the "/"
    const p = (pending = { keys: '', enter: false });
    try {
      for (let k = 0; k < 12 && pending === p; k++) {
        if (k) await new Promise((r) => setTimeout(r, 30));
        const c = await caret();
        if (!c || !c.collapsed) return;
        const text = await blockText(c.blockId);
        if (text === null || pending !== p) return;
        const typed = `/${p.keys}`;
        let m = Math.min(typed.length, c.offset);
        while (m > 0 && text.slice(c.offset - m, c.offset) !== typed.slice(0, m)) m--;
        if (m === 0) continue; // the "/" has not reached the model yet
        pending = null;
        const start = c.offset - m;
        const before = text.slice(0, start);
        if (before && !/\s$/.test(before)) { if (p.enter) await newLineAfter(c.blockId, text, c.offset); return; } // "and/or", dates, links
        open({ blockId: c.blockId, start }, pointAt(c.blockId, start));
        if (p.keys) setQuery(p.keys);
        if (p.enter && menu?.items.length) run(menu.items[0]);
        return;
      }
    } finally {
      if (pending === p) pending = null;
    }
  }
  /** An Enter held back while checking turned out to be ordinary typing: make the new line */
  async function newLineAfter(blockId, text, offset) {
    if (offset < text.length) return;
    try {
      const r = await api().create.paragraph({ at: { kind: 'after', target: { kind: 'block', nodeType: 'paragraph', nodeId: blockId } }, text: '' });
      const id = r?.paragraph?.nodeId;
      if (id) caretTo?.(id, 0);
      superdoc()?.focus?.();
    } catch { /* leave it */ }
  }

  // --------------------------------------------------------------- "Type / for commands" hint
  let hintEl = null; let hintTimer = 0;
  const hintOn = () => store.get(HINT_KEY, true) !== false;
  function hideHint() { hintEl?.remove(); hintEl = null; }
  function scheduleHint() {
    clearTimeout(hintTimer);
    hideHint();
    if (!hintOn() || menu) return;
    hintTimer = setTimeout(async () => {
      if (menu || !root.isConnected || root.offsetParent === null) return;
      const c = await caret();
      if (!c || !c.collapsed) return;
      const text = await blockText(c.blockId);
      if (text !== '' || menu) return;
      const pt = pointAt(c.blockId, 0);
      if (!pt) return;
      hintEl = document.createElement('div');
      hintEl.className = 'slash-hint';
      hintEl.textContent = 'Type / for commands';
      hintEl.style.left = `${pt.x + 2}px`;
      hintEl.style.top = `${pt.top}px`;
      hintEl.style.height = `${Math.max(14, pt.bottom - pt.top)}px`;
      document.body.appendChild(hintEl);
    }, 700);
  }

  // The editor's typing surface may live outside this view (SuperDoc mounts it separately)
  const ours = (t) => root.contains(t) || (!!root.closest?.('.view.active') && !!superdoc()?.activeEditor?.view?.dom?.contains?.(t));
  const onDocKey = (e) => { if (ours(e.target)) onKeyDown(e); };
  const onDocKeyUp = (e) => { if (ours(e.target)) scheduleHint(); };
  const onPointer = (e) => {
    if (menu && !menu.el.contains(e.target)) close();
    if (root.contains(e.target)) scheduleHint(); else hideHint();
  };
  const onScroll = () => { if (menu?.anchor) close(); hideHint(); };
  document.addEventListener('keydown', onDocKey, true);
  document.addEventListener('keyup', onDocKeyUp, true);
  document.addEventListener('pointerdown', onPointer, true);
  root.addEventListener('scroll', onScroll, true);
  window.addEventListener('resize', onScroll);

  return {
    open: () => open(null, null),
    close,
    get isOpen() { return !!menu; },
    hintEnabled: hintOn,
    setHint(on) { store.set(HINT_KEY, !!on); if (!on) hideHint(); else scheduleHint(); },
    dispose() {
      close(); hideHint();
      clearTimeout(hintTimer);
      document.removeEventListener('keydown', onDocKey, true);
      document.removeEventListener('keyup', onDocKeyUp, true);
      document.removeEventListener('pointerdown', onPointer, true);
      root.removeEventListener('scroll', onScroll, true);
      window.removeEventListener('resize', onScroll);
    },
  };
}
