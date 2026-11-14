/* kertas: symbol fonts */
// Inserted into SuperDoc 1.47.2's text layout by scripts/patch-deps.cjs.
//
// Word documents draw ticks, check boxes, arrows, Greek letters and the like with the Symbol,
// Wingdings and Wingdings 2 fonts, either as a <w:sym> element or as ordinary text in that font.
// Most computers outside Windows do not have those fonts, and SuperDoc 1.x did not draw <w:sym>
// at all. Here, at the moment text is laid out, each such character is swapped for the standard
// Unicode character that looks the same, in a font every system has. Only what is drawn changes:
// the document keeps the original character and font, so saving writes them back unchanged.
// Every swap is one character for one character, so the caret and selection stay in step.

const KERTAS_SYMBOL_STACK = '"Segoe UI Symbol", "Apple Symbols", "Noto Sans Symbols 2", "Noto Sans Symbols", "DejaVu Sans", sans-serif';

const kertasSymbolTables = (() => {
	const table = (pairs) => { const m = new Map(); for (const [code, ch] of pairs) m.set(code, ch); return m; };
	const range = (from, chars) => [...chars].map((ch, i) => [from + i, ch]);
	// Symbol: the standard Symbol encoding (Greek letters, maths and arrows)
	const symbol = table([
		[0x22, '∀'], [0x24, '∃'], [0x27, '∋'], [0x2a, '∗'], [0x2d, '−'], [0x40, '≅'], [0x5c, '∴'], [0x5e, '⊥'], [0x7e, '∼'],
		...range(0x41, 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ'),
		...range(0x61, 'αβχδεφγηιϕκλμνοπθρστυϖωξψζ'),
		[0xa1, 'ϒ'], [0xa2, '′'], [0xa3, '≤'], [0xa4, '⁄'], [0xa5, '∞'], [0xa6, 'ƒ'], [0xa7, '♣'], [0xa8, '♦'], [0xa9, '♥'], [0xaa, '♠'],
		[0xab, '↔'], [0xac, '←'], [0xad, '↑'], [0xae, '→'], [0xaf, '↓'], [0xb0, '°'], [0xb1, '±'], [0xb2, '″'], [0xb3, '≥'], [0xb4, '×'],
		[0xb5, '∝'], [0xb6, '∂'], [0xb7, '•'], [0xb8, '÷'], [0xb9, '≠'], [0xba, '≡'], [0xbb, '≈'], [0xbc, '…'], [0xc0, 'ℵ'], [0xc1, 'ℑ'],
		[0xc2, 'ℜ'], [0xc3, '℘'], [0xc4, '⊗'], [0xc5, '⊕'], [0xc6, '∅'], [0xc7, '∩'], [0xc8, '∪'], [0xc9, '⊃'], [0xca, '⊇'], [0xcb, '⊄'],
		[0xcc, '⊂'], [0xcd, '⊆'], [0xce, '∈'], [0xcf, '∉'], [0xd0, '∠'], [0xd1, '∇'], [0xd2, '®'], [0xd3, '©'], [0xd4, '™'], [0xd5, '∏'],
		[0xd6, '√'], [0xd7, '⋅'], [0xd8, '¬'], [0xd9, '∧'], [0xda, '∨'], [0xdb, '⇔'], [0xdc, '⇐'], [0xdd, '⇑'], [0xde, '⇒'], [0xdf, '⇓'],
		[0xe0, '◊'], [0xe1, '〈'], [0xe5, '∑'], [0xf1, '〉'], [0xf2, '∫'],
	]);
	// digits and the other ASCII characters Symbol draws as themselves (all but the radical extender at 60)
	for (let c = 0x21; c <= 0x7e; c++) if (!symbol.has(c) && c !== 0x60) symbol.set(c, String.fromCharCode(c));
	// Wingdings: the characters documents use most (ticks, crosses, boxes, bullets, arrows, numbers)
	const wingdings = table([
		[0x21, '✏'], [0x22, '✂'], [0x23, '✁'], [0x28, '☎'], [0x2a, '✉'], [0x36, '⌛'], [0x37, '⌨'], [0x3f, '✍'], [0x41, '✌'],
		[0x45, '☜'], [0x46, '☞'], [0x47, '☝'], [0x48, '☟'], [0x4a, '☺'], [0x4c, '☹'], [0x4e, '☠'], [0x4f, '⚐'], [0x51, '✈'],
		[0x52, '☼'], [0x54, '❄'], [0x56, '✞'], [0x58, '✠'], [0x59, '✡'], [0x5a, '☪'], [0x5b, '☯'], [0x5d, '☸'],
		...range(0x5e, '♈♉♊♋♌♍♎♏♐♑♒♓'),
		[0x6c, '●'], [0x6d, '❍'], [0x6e, '■'], [0x6f, '□'], [0x71, '❑'], [0x72, '❒'], [0x74, '⧫'], [0x75, '◆'], [0x76, '❖'],
		[0x78, '⌧'], [0x7a, '⌘'], [0x7b, '❀'], [0x7c, '✿'], [0x7d, '❝'], [0x7e, '❞'],
		[0x80, '⓪'], ...range(0x81, '①②③④⑤⑥⑦⑧⑨⑩'), [0x8b, '⓿'], ...range(0x8c, '❶❷❸❹❺❻❼❽❾❿'),
		[0x9e, '·'], [0x9f, '•'], [0xa0, '▪'], [0xa1, '○'], [0xa4, '◉'], [0xa5, '◎'], [0xa7, '▪'], [0xa8, '◻'], [0xaa, '✦'],
		[0xab, '★'], [0xac, '✶'], [0xad, '✴'], [0xae, '✹'], [0xaf, '✵'], [0xb1, '⌖'], [0xd5, '⌫'], [0xd6, '⌦'], [0xd8, '➢'],
		[0xdf, '←'], [0xe0, '→'], [0xe1, '↑'], [0xe2, '↓'], [0xe3, '↖'], [0xe4, '↗'], [0xe5, '↙'], [0xe6, '↘'], [0xe8, '➔'],
		[0xef, '⇦'], [0xf0, '⇨'], [0xf1, '⇧'], [0xf2, '⇩'], [0xf3, '⬄'], [0xf4, '⇳'],
		[0xfb, '✘'], [0xfc, '✔'], [0xfd, '☒'], [0xfe, '☑'],
	]);
	// Wingdings 2: the tick and check box characters forms use
	const wingdings2 = table([[0x4f, '✗'], [0x50, '✓'], [0x52, '☑'], [0x54, '☒'], [0xa3, '□']]);
	return new Map([['symbol', symbol], ['wingdings', wingdings], ['wingdings 2', wingdings2]]);
})();

/** The first family in a CSS font-family list, lower case and without quotes */
function kertasFirstFamily(family) {
	if (typeof family !== 'string') return '';
	return family.split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
}

/** Swaps symbol-font characters in a text run for the Unicode characters that look the same */
function kertasSymbolRun(run) {
	const table = kertasSymbolTables.get(kertasFirstFamily(run?.fontFamily));
	if (!table || typeof run.text !== 'string' || !run.text) return run;
	let changed = false;
	let missing = false;
	let out = '';
	for (const ch of run.text) {
		const code = ch.charCodeAt(0);
		// Word stores these characters either as plain codes or moved into the private-use area at F000
		const mapped = table.get(code >= 0xf000 && code <= 0xf0ff ? code - 0xf000 : code);
		if (mapped) { out += mapped; changed = true; }
		else { out += ch; if (ch.trim()) missing = true; }
	}
	// a run with a character we cannot swap stays as it is, so that character still draws if the font is there
	if (!changed || missing) return run;
	return { ...run, text: out, fontFamily: KERTAS_SYMBOL_STACK };
}

/** Draws a <w:sym> element, which SuperDoc 1.x keeps but does not lay out */
function kertasSymNodeToRun(params) {
	const node = params.node;
	if (node?.attrs?.originalName !== 'w:sym') return null;
	const a = node.attrs.originalXml?.attributes || {};
	const code = parseInt(a['w:char'], 16);
	if (!Number.isFinite(code)) return null;
	const run = textNodeToRun({ ...params, node: { ...node, type: 'text', text: String.fromCharCode(code), marks: node.marks ?? [] } });
	const pos = params.positions?.get(node);
	if (pos) { run.pmStart = pos.start; run.pmEnd = pos.end; }
	if (a['w:font']) run.fontFamily = a['w:font'];
	return kertasSymbolRun(run);
}
