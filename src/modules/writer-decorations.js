// Page borders and line numbers, drawn on top of the painted pages.
// Both are kept in the document by the editor (and saved to Word); this makes them visible
// while editing, as Word and LibreOffice show them.
const PX_PER_PT = 96 / 72;
const CSS_STYLE = { single: 'solid', thick: 'solid', double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed', dotDash: 'dashed', dotDotDash: 'dotted', triple: 'double', thinThickSmallGap: 'double', thickThinSmallGap: 'double', wave: 'solid', inset: 'inset', outset: 'outset' };

function borderCss(b) {
  if (!b || !b.style || b.style === 'none' || b.style === 'nil') return 'none';
  const w = Math.max(1, ((b.size || 4) / 8) * PX_PER_PT); // size is in eighths of a point
  const css = CSS_STYLE[b.style] || 'solid';
  const color = !b.color || b.color === 'auto' ? '#000' : `#${b.color}`;
  return `${css === 'double' ? Math.max(3, w) : w}px ${css} ${color}`;
}

export function createDecorations(host, getEditor) {
  let timer = null;
  let busy = false;
  const OWN = 'data-k-deco';

  let secCache = null; let secDoc = null;
  function sectionsInfo(ed) {
    if (secDoc === ed.state.doc && secCache) return secCache;
    try { secCache = ed.doc.sections.list().items || []; } catch { secCache = []; }
    secDoc = ed.state.doc;
    return secCache;
  }
  function pageSections(ed) {
    try { return (ed.presentationEditor?.getLayoutSnapshot?.()?.layout?.pages || []).map((p) => p.sectionIndex ?? 0); } catch { return []; }
  }

  function paint() {
    const ed = getEditor();
    if (!ed?.doc || busy) return;
    busy = true;
    try {
      const secs = sectionsInfo(ed);
      const needed = secs.some((x) => x.lineNumbering?.enabled || (x.pageBorders && ['top', 'right', 'bottom', 'left'].some((k) => x.pageBorders[k]?.style && x.pageBorders[k].style !== 'none')));
      if (!needed && !host.querySelector(`[${OWN}]`)) return;
      const map = pageSections(ed);
      const pages = [...host.querySelectorAll('.superdoc-page')];
      let running = 0; let lastSection = -1;
      pages.forEach((page, i) => {
        page.querySelectorAll(`[${OWN}]`).forEach((n) => n.remove());
        const idx = map[Number(page.dataset.pageIndex ?? i)] ?? 0;
        const sec = secs[idx] || {};
        // ---- page border
        const pb = sec.pageBorders;
        if (pb && ['top', 'right', 'bottom', 'left'].some((s) => pb[s]?.style && pb[s].style !== 'none')) {
          const fromText = pb.offsetFrom === 'text';
          const m = sec.margins || {};
          const pt = (side) => (pb[side]?.space || 0) * PX_PER_PT;
          const box = document.createElement('div');
          box.setAttribute(OWN, 'border');
          box.className = 'k-page-border';
          const inset = fromText
            ? { top: (m.top || 1) * 96 - pt('top'), right: (m.right || 1) * 96 - pt('right'), bottom: (m.bottom || 1) * 96 - pt('bottom'), left: (m.left || 1) * 96 - pt('left') }
            : { top: pt('top'), right: pt('right'), bottom: pt('bottom'), left: pt('left') };
          box.style.cssText = `position:absolute;pointer-events:none;z-index:0;top:${inset.top}px;right:${inset.right}px;bottom:${inset.bottom}px;left:${inset.left}px;border-top:${borderCss(pb.top)};border-right:${borderCss(pb.right)};border-bottom:${borderCss(pb.bottom)};border-left:${borderCss(pb.left)}`;
          page.appendChild(box);
        }
        // ---- line numbers
        const ln = sec.lineNumbering;
        if (idx !== lastSection) { if (ln?.restart !== 'continuous' || lastSection === -1) running = (ln?.start ?? 1) - 1; lastSection = idx; }
        if (!ln?.enabled) return;
        if (ln.restart === 'newPage') running = (ln.start ?? 1) - 1;
        const every = Math.max(1, ln.countBy || 1);
        const pageRect = page.getBoundingClientRect();
        const scale = pageRect.width / (page.offsetWidth || pageRect.width) || 1;
        const marginLeft = (sec.margins?.left ?? 1) * 96;
        const distance = ln.distance != null ? ln.distance * 96 : 0.25 * 96;
        const layer = document.createElement('div');
        layer.setAttribute(OWN, 'lines');
        layer.className = 'k-line-numbers';
        layer.style.cssText = 'position:absolute;inset:0;pointer-events:none;z-index:2';
        for (const line of page.querySelectorAll('.superdoc-line')) {
          // Word does not number header, footer, footnote or table lines
          if (line.closest('.superdoc-page-header, .superdoc-page-footer, [class*=footnote], .superdoc-table-fragment')) continue;
          if (!line.textContent.trim()) continue;
          running += 1;
          if (running % every) continue;
          const r = line.getBoundingClientRect();
          const n = document.createElement('span');
          n.textContent = String(running);
          n.style.cssText = `position:absolute;top:${(r.top - pageRect.top) / scale}px;height:${r.height / scale}px;line-height:${r.height / scale}px;right:${(page.offsetWidth - marginLeft + distance)}px`;
          layer.appendChild(n);
        }
        page.appendChild(layer);
      });
    } finally {
      busy = false;
    }
  }
  // Runs once typing pauses and the browser is idle; documents without page borders or line
  // numbers cost almost nothing (the section check is cached between document changes)
  const idle = (fn) => (typeof requestIdleCallback === 'function' ? requestIdleCallback(fn, { timeout: 800 }) : fn());
  const schedule = () => { clearTimeout(timer); timer = setTimeout(() => idle(paint), 350); };
  // The pages are repainted as the document changes; follow them (ignoring our own layers)
  const obs = new MutationObserver((list) => {
    if (busy) return;
    if (list.every((m) => [...m.addedNodes, ...m.removedNodes].every((n) => n.nodeType === 1 && n.hasAttribute?.(OWN)))) return;
    schedule();
  });
  obs.observe(host, { childList: true, subtree: true });
  schedule();
  return { refresh: schedule, destroy: () => { obs.disconnect(); clearTimeout(timer); } };
}
