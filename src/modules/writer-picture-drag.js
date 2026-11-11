// Drag a floating picture to move it, as in Word. SuperDoc draws resize handles on a selected
// picture but has no way to move one with the mouse; this adds that. The picture keeps moving
// with its paragraph: the drop is saved as an offset from the page's left edge and from the top
// of the paragraph it is anchored to (Word's "Move with text").
export function enablePictureDrag(host, { editor, api, toast, onMoved }) {
  let drag = null;

  const zoomOf = (page) => (page.getBoundingClientRect().width / (page.offsetWidth || 1)) || 1;

  function down(e) {
    if (e.button !== 0) return;
    const frag = e.target.closest?.('.superdoc-image-fragment');
    if (!frag || e.target.closest('.resize-handle, .sd-editor-resize-handle')) return;
    drag = { x: e.clientX, y: e.clientY, frag, started: false, id: null };
  }

  async function begin() {
    // the painted picture names its paragraph and its place in it: "<blockId>-image-<n>"
    const m = /^(.*)-image-(\d+)$/.exec(drag.frag.dataset.blockId || '');
    let info = null;
    try {
      if (m) {
        const items = ((await api().images.list({}))?.items || []).filter((x) => x.anchor?.blockId === m[1]);
        info = items[Number(m[2])] || items[0] || null;
      }
      if (!info) {
        const node = editor()?.state?.selection?.node;
        const sel = node?.type?.name === 'image' ? node.attrs.sdImageId : null;
        if (sel) info = await api().images.get({ imageId: sel });
      }
    } catch { return false; }
    const id = info?.sdImageId || info?.address?.nodeId;
    if (!id) return false;
    if (info?.properties?.placement !== 'floating') {
      toast('Set Wrap text to Square, Tight or another wrap first, then drag the picture');
      return false;
    }
    drag.id = id; drag.info = info;
    const r = drag.frag.getBoundingClientRect();
    const ghost = drag.frag.cloneNode(true);
    ghost.className = 'k-pic-ghost';
    ghost.style.cssText = `position:fixed;left:${r.left}px;top:${r.top}px;width:${r.width}px;height:${r.height}px;pointer-events:none;z-index:2000;opacity:.7;outline:2px dashed var(--accent);`;
    document.body.appendChild(ghost);
    drag.ghost = ghost; drag.rect = r;
    drag.frag.style.opacity = '.35';
    return true;
  }

  async function move(e) {
    if (!drag) return;
    const dx = e.clientX - drag.x; const dy = e.clientY - drag.y;
    if (!drag.started) {
      if (Math.hypot(dx, dy) < 5) return;
      drag.started = true;
      drag.pending = begin().then((ok) => { if (!ok) end(); return ok; });
    }
    e.preventDefault(); e.stopPropagation();
    if (drag?.ghost) drag.ghost.style.transform = `translate(${dx}px, ${dy}px)`;
  }

  function end() {
    if (!drag) return;
    drag.ghost?.remove();
    if (drag.frag) drag.frag.style.opacity = '';
    drag = null;
  }

  async function up(e) {
    if (!drag) return;
    const d = drag;
    if (!d.started) { drag = null; return; }
    e.preventDefault(); e.stopPropagation();
    const ok = await d.pending;
    if (!ok || !d.id) { end(); return; }
    const dx = e.clientX - d.x; const dy = e.clientY - d.y;
    const page = d.frag.closest('.superdoc-page');
    const z = zoomOf(page);
    const pr = page.getBoundingClientRect();
    const left = (d.rect.left + dx - pr.left) / z;
    // the anchor paragraph on this page (falls back to the page top)
    const anchorId = d.info.anchor?.blockId;
    const para = anchorId ? page.querySelector(`.superdoc-fragment[data-block-id="${CSS.escape(anchorId)}"]`) : null;
    const top = para ? (d.rect.top + dy - para.getBoundingClientRect().top) / z : (d.rect.top + dy - pr.top) / z;
    end();
    try {
      const res = await api().images.setPosition({
        imageId: d.id,
        position: { hRelativeFrom: 'page', vRelativeFrom: para ? 'paragraph' : 'page', alignH: null, alignV: null, marginOffset: { horizontal: Math.round(left), top: Math.round(top) } },
      });
      if (res && res.success === false) throw new Error(res.failure?.message || 'The picture could not be moved');
      onMoved?.(d.id);
    } catch (err) { toast(String(err?.message || err), true); }
  }

  host.addEventListener('pointerdown', down, true);
  // the browser's own picture drag-and-drop would swallow the pointer release
  const noNative = (e) => { if (e.target.closest?.('.superdoc-image-fragment')) e.preventDefault(); };
  host.addEventListener('dragstart', noNative, true);
  // on the window, ahead of the editor's own pointer handling (it keeps the pointer while selecting)
  window.addEventListener('pointermove', move, true);
  window.addEventListener('pointerup', up, true);
  window.addEventListener('mouseup', up, true);
  return () => {
    host.removeEventListener('pointerdown', down, true);
    host.removeEventListener('dragstart', noNative, true);
    window.removeEventListener('pointermove', move, true);
    window.removeEventListener('pointerup', up, true);
    window.removeEventListener('mouseup', up, true);
  };
}
