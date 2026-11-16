/* kertas: diagonal cell borders */
// Inserted into SuperDoc 1.47.2's table painter by scripts/patch-deps.cjs.
//
// Word table cells can have diagonal borders (w:tl2br from top left to bottom right, w:tr2bl from
// top right to bottom left), often used in the corner cell of a table with row and column headings.
// SuperDoc 1.x read them but never drew them. This draws each one as a line in an SVG layer behind
// the cell's text, with the border's colour, width and line style. The approach follows SuperDoc
// 2.x's AGPL table painter (packages/layout-engine/painters/dom/src/table/border-utils.ts).

function kertasCellDiagonals(doc, cellEl, cell, width, height) {
	const borders = cell?.attrs?.tableCellProperties?.borders;
	if (!borders || !(width > 0) || !(height > 0)) return;
	const sides = ['tl2br', 'tr2bl'].filter((side) => {
		const b = borders[side];
		return b && b.val && b.val !== 'nil' && b.val !== 'none';
	});
	if (!sides.length) return;
	const NS = 'http://www.w3.org/2000/svg';
	const svg = doc.createElementNS(NS, 'svg');
	svg.setAttribute('width', String(width));
	svg.setAttribute('height', String(height));
	svg.setAttribute('aria-hidden', 'true');
	svg.style.cssText = 'position:absolute;left:0;top:0;overflow:visible;pointer-events:none';
	svg.dataset.kertasDiagonal = sides.join(' ');
	for (const side of sides) {
		const b = borders[side];
		// size is in eighths of a point; at least a hairline
		const px = Math.max(0.75, ((Number(b.size) || 4) / 8) * (96 / 72));
		const color = !b.color || b.color === 'auto' ? '#000000' : (b.color.startsWith('#') ? b.color : `#${b.color}`);
		const [x1, x2] = side === 'tl2br' ? [0, width] : [width, 0];
		const len = Math.hypot(width, height);
		// a line moved sideways by offset px, square to its direction
		const line = (offset, strokeWidth, dash) => {
			const nx = (-height / len) * offset * (side === 'tl2br' ? 1 : -1);
			const ny = (width / len) * offset;
			const el = doc.createElementNS(NS, 'line');
			el.setAttribute('x1', String(x1 + nx));
			el.setAttribute('y1', String(0 + ny));
			el.setAttribute('x2', String(x2 + nx));
			el.setAttribute('y2', String(height + ny));
			el.setAttribute('stroke', color);
			el.setAttribute('stroke-width', String(strokeWidth));
			if (dash) el.setAttribute('stroke-dasharray', dash);
			svg.appendChild(el);
		};
		const style = String(b.val);
		if (style === 'double') { line(-px, Math.max(0.6, px / 2)); line(px, Math.max(0.6, px / 2)); }
		else if (style === 'dashed' || style === 'dashSmallGap') line(0, px, `${px * 4} ${px * 2}`);
		else if (style === 'dotted') line(0, px, `${px} ${px * 1.5}`);
		else if (style === 'dotDash') line(0, px, `${px * 4} ${px * 2} ${px} ${px * 2}`);
		else if (style === 'dotDotDash') line(0, px, `${px * 4} ${px * 2} ${px} ${px * 2} ${px} ${px * 2}`);
		else line(0, px);
	}
	// behind the text: first in the cell
	cellEl.insertBefore(svg, cellEl.firstChild);
}
