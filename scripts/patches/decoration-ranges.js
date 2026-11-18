// kertas: fast decoration ranges
// After a document change SuperDoc tries to put each decoration (search highlights among
// them) back by finding its text in the document. The original read the whole document
// again for every candidate of every decoration: with a few hundred highlights a single
// keystroke took seconds, and a long document froze. Here the document's text, its
// offset-to-position table and the matches of each text are worked out once per document.
var kertasRangeCache = /* @__PURE__ */ new WeakMap();
function kertasRangeIndex(doc$1, blockSep, leafSep) {
	let c = kertasRangeCache.get(doc$1);
	if (c && c.blockSep === blockSep && c.leafSep === leafSep) return c;
	const size = doc$1.content.size;
	// lens[p] = doc.textBetween(0, p, blockSep, leafSep).length, for every position p
	const lens = new Uint32Array(size + 2);
	const bl = blockSep ? blockSep.length : 0;
	const ll = typeof leafSep === "string" ? leafSep.length : 0;
	let first = true;
	doc$1.nodesBetween(0, size, (node, pos) => {
		const at = Math.min(pos + 1, size + 1); // counts once the range reaches into the node
		const leafText = node.isText ? 1 : node.isLeaf ? ll : 0;
		if (node.isBlock && (node.isLeaf && leafText || node.isTextblock) && bl) {
			if (first) first = false;
			else lens[at] += bl;
		}
		if (node.isText) { const n = node.text.length; for (let i = 1; i <= n && pos + i <= size + 1; i++) lens[pos + i] += 1; }
		else if (node.isLeaf && ll) lens[at] += ll;
	}, 0);
	for (let p = 1; p < lens.length; p++) lens[p] += lens[p - 1];
	c = { blockSep, leafSep, size, lens, full: null, matches: /* @__PURE__ */ new Map() };
	kertasRangeCache.set(doc$1, c);
	return c;
}
function charOffsetToPosition(doc$1, charOffset, blockSep, leafSep) {
	const c = kertasRangeIndex(doc$1, blockSep, leafSep);
	const docSize = c.size;
	if (charOffset <= 0) return 0;
	if (charOffset >= c.lens[docSize]) return docSize;
	let low = 0;
	let high = docSize;
	while (low < high) {
		const mid = Math.floor((low + high) / 2);
		if (c.lens[mid] < charOffset) low = mid + 1;
		else high = mid;
	}
	return low;
}
function findRangeByText(doc$1, text, hintFrom) {
	if (!text) return null;
	const c = kertasRangeIndex(doc$1, TEXT_RANGE_BLOCK_SEP, TEXT_RANGE_LEAF_SEP);
	if (c.full == null) c.full = doc$1.textBetween(0, c.size, TEXT_RANGE_BLOCK_SEP, TEXT_RANGE_LEAF_SEP);
	let matches = c.matches.get(text);
	if (!matches) {
		matches = [];
		for (let i = 0; ;) {
			const idx = c.full.indexOf(text, i);
			if (idx === -1) break;
			matches.push(idx);
			i = idx + 1;
		}
		c.matches.set(text, matches);
	}
	if (matches.length === 0) return null;
	const toPos = (charOffset) => charOffsetToPosition(doc$1, charOffset, TEXT_RANGE_BLOCK_SEP, TEXT_RANGE_LEAF_SEP);
	let charOffsetFrom = matches[0];
	if (hintFrom != null) {
		// positions grow with offsets, so the match nearest the hint is next to where the hint
		// would sit in the sorted list; ties go to the earlier match, as before
		const hintChar = c.lens[Math.max(0, Math.min(c.size, Math.round(hintFrom)))];
		let lo = 0;
		let hi = matches.length;
		while (lo < hi) { const mid = (lo + hi) >> 1; if (matches[mid] < hintChar) lo = mid + 1; else hi = mid; }
		let best = null;
		let bestDist = Infinity;
		for (let k = Math.max(0, lo - 2); k <= Math.min(matches.length - 1, lo + 1); k++) {
			const d = Math.abs(toPos(matches[k]) - hintFrom);
			if (d < bestDist) { best = matches[k]; bestDist = d; }
		}
		charOffsetFrom = best;
	}
	const from$1 = toPos(charOffsetFrom);
	const to = toPos(charOffsetFrom + text.length);
	return from$1 < to ? {
		from: from$1,
		to
	} : null;
}
