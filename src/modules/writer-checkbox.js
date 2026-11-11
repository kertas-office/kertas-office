// Clickable tick boxes: clicking a ☐ in the page ticks it (☑), clicking again clears it.
// The boxes are ordinary characters, so Word, LibreOffice and PDF show them as they are.
const BOXES = { '☐': '☑', '☑': '☐', '☒': '☐' };

/** The tick-box character under the pointer, with its document position, or null */
function boxAt(x, y) {
  // The painted text can sit under the editor's own overlays, so look through every layer
  const span = document.elementsFromPoint(x, y).find((el) => el.matches?.('span[data-pm-start]'));
  if (!span || !/[☐☑☒]/.test(span.textContent)) return null;
  let before = 0;
  const walker = document.createTreeWalker(span, NodeFilter.SHOW_TEXT);
  const glyph = document.createRange();
  for (let t = walker.nextNode(); t; t = walker.nextNode()) {
    for (let i = 0; i < t.data.length; i++) {
      const ch = t.data[i];
      if (!BOXES[ch]) continue;
      glyph.setStart(t, i); glyph.setEnd(t, i + 1);
      const b = glyph.getBoundingClientRect();
      if (x >= b.left - 1 && x <= b.right + 1 && y >= b.top - 2 && y <= b.bottom + 2) return { ch, pos: Number(span.dataset.pmStart) + before + i };
    }
    before += t.data.length;
  }
  return null;
}

export function enableClickableBoxes(host, getEditor, onChange) {
  const down = (e) => {
    if (e.button !== 0 || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey) return;
    const hit = boxAt(e.clientX, e.clientY);
    if (!hit) return;
    const ed = getEditor();
    if (!ed?.state || ed.options?.documentMode === 'viewing') return;
    const now = ed.state.doc.textBetween(hit.pos, hit.pos + 1);
    if (now !== hit.ch) return; // the painted page is behind the document; leave it
    e.preventDefault(); e.stopPropagation();
    const tr = ed.state.tr.insertText(BOXES[hit.ch], hit.pos, hit.pos + 1);
    ed.view.dispatch(tr);
    onChange?.();
  };
  const move = (e) => {
    const over = !!boxAt(e.clientX, e.clientY);
    if (over !== host.classList.contains('k-over-box')) host.classList.toggle('k-over-box', over);
  };
  // The editor acts on pointerdown, so the click is claimed there; mousedown is swallowed too
  let claimed = false;
  const pdown = (e) => { claimed = false; if (e.isPrimary !== false) { down(e); claimed = e.defaultPrevented; } };
  const mdown = (e) => { if (claimed) { e.preventDefault(); e.stopPropagation(); } };
  host.addEventListener('pointerdown', pdown, true);
  host.addEventListener('mousedown', mdown, true);
  host.addEventListener('mousemove', move, { passive: true });
  return () => { host.removeEventListener('pointerdown', pdown, true); host.removeEventListener('mousedown', mdown, true); host.removeEventListener('mousemove', move); };
}
