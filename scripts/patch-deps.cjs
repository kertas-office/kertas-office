// Small fixes to third-party packages, applied after `npm install` (idempotent).
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..', 'node_modules');
const patches = [
  {
    // superdoc 1.47.2: right-clicking inside selected table cells moved the cursor and dropped the
    // cell selection, because "inside the selection" only looked at the first and last positions.
    // Word keeps the selected cells, so the menu's actions apply to them.
    dir: 'superdoc/dist/chunks',
    find: /isClickInsideSelection = hit\.pos >= from\$1 && hit\.pos <= to;/,
    replace: 'isClickInsideSelection = hit.pos >= from$$1 && hit.pos <= to; if (!isClickInsideSelection && typeof editorState.selection.forEachCell === "function") editorState.selection.forEachCell((n, p) => { if (hit.pos >= p && hit.pos <= p + n.nodeSize) isClickInsideSelection = true; }); // kertas: right-click in selected cells',
    marker: 'kertas: right-click in selected cells',
  },
  {
    // superdoc 1.47.2: once a table had any border set directly (one edge, say), the borders from
    // its table style were dropped, so setting the top border erased all the others. Word keeps the
    // style's borders and overrides only the edges that were set. The converter already merges the
    // two; this uses that merged set whenever the table has direct borders.
    dir: 'superdoc/dist/chunks',
    find: /const getBorderSource = \(\) => \{(\s*)if \(node\.attrs\?\.borders && /,
    replace: 'const getBorderSource = () => {$1if (node.attrs?.tableProperties?.borders && hydratedTableStyle?.borders && typeof hydratedTableStyle.borders === "object") return { borders: hydratedTableStyle.borders, unit: "eighthPoints" }; // kertas: style and direct table borders$1if (node.attrs?.borders && ',
    marker: 'kertas: style and direct table borders',
  },
  {
    // superdoc 1.47.2: "single" line spacing was always 1.15 times the font size. Word uses
    // each font's own line height instead: 1.22 for Calibri, 1.17 for Cambria, 1.15 for Arial
    // and Times New Roman. Calibri text (most Word files) came out about 6% tighter than in
    // Word and WPS, so pages also broke in different places.
    dir: 'superdoc/dist/chunks',
    find: /const lineHeight = resolveLineHeight\(spacing, fontSize, ascent \+ descent\);/,
    replace: 'const lineHeight = resolveLineHeight(spacing, fontSize, ascent + descent) * ((spacing?.lineRule ?? "auto") === "auto" ? (({ calibri: 1.2207, "calibri light": 1.2207, carlito: 1.2207, cambria: 1.1724, caladea: 1.1724, verdana: 1.2153, tahoma: 1.2075, "segoe ui": 1.3301, georgia: 1.1362, consolas: 1.1709, "courier new": 1.1328, cousine: 1.1328 })[String(fontInfo?.fontFamily || "").split(",")[0].replace(/[^a-z ]/gi, "").trim().toLowerCase()] || 1.15) / 1.15 : 1); // kertas: word line height',
    marker: 'kertas: word line height',
  },
  {
    // the same for empty paragraphs
    dir: 'superdoc/dist/chunks',
    find: /const lineHeight = roundValue\(resolveLineHeight\(spacing, resolvedFontSize, Math\.max\(resolvedFontSize, ascent \+ descent\)\)\);/,
    replace: 'const lineHeight = roundValue(resolveLineHeight(spacing, resolvedFontSize, Math.max(resolvedFontSize, ascent + descent)) * ((spacing?.lineRule ?? "auto") === "auto" ? (({ calibri: 1.2207, "calibri light": 1.2207, carlito: 1.2207, cambria: 1.1724, caladea: 1.1724, verdana: 1.2153, tahoma: 1.2075, "segoe ui": 1.3301, georgia: 1.1362, consolas: 1.1709, "courier new": 1.1328, cousine: 1.1328 })[String(fontInfo?.fontFamily || "").split(",")[0].replace(/[^a-z ]/gi, "").trim().toLowerCase()] || 1.15) / 1.15 : 1)); // kertas: word line height (empty)',
    marker: 'kertas: word line height (empty)',
  },
  {
    // fontkit 2.0.4: OpenType allows a NULL anchor in mark attachment tables
    // (Noto Sans Malayalam and Devanagari have them); fontkit crashes on it.
    files: ['fontkit/dist/module.mjs', 'fontkit/dist/main.cjs', 'fontkit/dist/browser-module.mjs', 'fontkit/dist/browser.cjs'],
    find: /getAnchor\(anchor\) \{\n(\s*)\/\/ TODO: contour point, device tables\n/,
    replace: 'getAnchor(anchor) {\n$1if (!anchor) return { x: 0, y: 0 }; // kertas: NULL anchor\n$1// TODO: contour point, device tables\n',
    marker: 'kertas: NULL anchor',
  },
  {
    // superdoc 1.47.2 (AGPL): text beside a right-floating box or picture measured its width
    // from the page edge instead of the margin, so lines ran underneath the box.
    dir: 'superdoc/dist/chunks',
    find: /(const columnOrigin = getColumnX\(getColumnGeometry\(currentColumns\), columnIndex, marginLeft\);[\s\S]{0,900}?)let leftBoundary = 0;/,
    replace: '$1let leftBoundary = columnOrigin; // kertas: right-float wrap',
    marker: 'kertas: right-float wrap',
  },
  {
    // superdoc 1.47.2: a display equation was saved as a bare <m:oMathPara> in the body.
    // Word and LibreOffice expect it inside a paragraph, and LibreOffice dropped it.
    dir: 'superdoc/dist/chunks',
    find: /mathBlock: translatePassthroughNode,/,
    replace: 'mathBlock: (params) => { const x = translatePassthroughNode(params); return x && x.name === "m:oMathPara" ? { type: "element", name: "w:p", elements: [x] } : x; }, // kertas: math in paragraph',
    marker: 'kertas: math in paragraph',
  },
  {
    // superdoc 1.47.2: pictures made floating in the editor were saved with simplePos="1",
    // which pins them to the top-left corner of the page in Word and LibreOffice.
    dir: 'superdoc/dist/chunks',
    find: /if \(attrs\.originalAttributes\?\.simplePos === void 0 && hasSimplePos\) inlineAttrs\.simplePos = "1";/,
    replace: 'if (attrs.originalAttributes?.simplePos === void 0 && hasSimplePos) inlineAttrs.simplePos = attrs.simplePos?.x || attrs.simplePos?.y ? "1" : "0"; // kertas: simplePos',
    marker: 'kertas: simplePos',
  },
  {
    // superdoc 1.47.2: typing "/" always opened SuperDoc's menu. Kertas has its own "/" menu,
    // so contextMenu: { noSlash: true } leaves "/" to it (right-click still opens SuperDoc's).
    dir: 'superdoc/dist/chunks',
    find: /if \(event\.key === "\/" && !pluginState\.open\) \{/,
    replace: 'if (event.key === "/" && !pluginState.open && !editor.options?.contextMenuConfig?.noSlash) { // kertas: own slash menu',
    marker: 'kertas: own slash menu',
  },
  {
    // superdoc 1.47.2: the "/" menu only matched labels; Kertas items also carry keywords
    // ("check" finds Tick box, "formula" finds Equation).
    dir: 'superdoc/dist/chunks',
    find: /return flattenedItems\.value\.filter\(\(item\) => item\.label\?\.toLowerCase\(\)\.includes\(searchQuery\.value\.toLowerCase\(\)\)\);/,
    replace: 'const q = searchQuery.value.toLowerCase(); return flattenedItems.value.filter((item) => item.label?.toLowerCase().includes(q) || (Array.isArray(item.keywords) && item.keywords.some((k) => String(k).toLowerCase().includes(q)))); // kertas: menu keywords',
    marker: 'kertas: menu keywords',
  },
  {
    // superdoc 1.47.2: after typing to filter the "/" menu, Enter did nothing, because the
    // highlighted item still pointed at one that had been filtered out.
    dir: 'superdoc/dist/chunks',
    find: /const selectedItem = currentItems\.find\(\(item\) => item\.id === selectedId\.value\);/,
    replace: 'const selectedItem = currentItems.find((item) => item.id === selectedId.value) || currentItems[0]; // kertas: enter picks first match',
    marker: 'kertas: enter picks first match',
  },
  {
    // ...and the highlight follows the first match while typing, so what is shown is what Enter picks
    dir: 'superdoc/dist/chunks',
    find: /watch\(flattenedItems, \(newItems\) => \{\n(\t+)if \(newItems\.length > 0\) selectedId\.value = newItems\[0\]\.id;/,
    replace: 'watch(filteredItems, (shown) => { if (shown.length && !shown.some((item) => item.id === selectedId.value)) selectedId.value = shown[0].id; }); // kertas: highlight first match\n\t\twatch(flattenedItems, (newItems) => {\n$1if (newItems.length > 0) selectedId.value = newItems[0].id;',
    marker: 'kertas: highlight first match',
  },
  {
    // superdoc 1.47.2: structural inserts never let fields through (the raw-mode opt-in is
    // dropped on the way in), so page numbers could not be added to headers and footers.
    // Kertas is the only caller, and it asks for fields on purpose.
    dir: 'superdoc/dist/chunks',
    find: /const pmFragment = materializeFragment\(schema, content, collectExistingBlockIds\(editor\.state\.doc\), "insert"\);/,
    replace: 'const pmFragment = materializeFragment(schema, content, collectExistingBlockIds(editor.state.doc), "insert", { rawMode: true }); // kertas: raw insert',
    marker: 'kertas: raw insert',
  },
  {
    // ...and PAGE / NUMPAGES become the editor's live page-number nodes (saved as Word fields)
    // instead of fixed text
    dir: 'superdoc/dist/chunks',
    find: /function materializeInlineField\(schema, node\) \{\n(\t)const payload = resolvePayload\(node, "field"\);/,
    replace: 'function materializeInlineField(schema, node) {\n$1const payload = resolvePayload(node, "field");\n$1{ const t = String(payload.fieldType || payload.instruction || "").trim().split(/\\s+/)[0].toUpperCase(); const live = t === "PAGE" ? "page-number" : t === "NUMPAGES" || t === "SECTIONPAGES" ? "total-page-number" : null; if (live && schema.nodes[live]) return schema.nodes[live].create(); } // kertas: live page numbers',
    marker: 'kertas: live page numbers',
  },
  {
    // superdoc 1.47.2: images inserted structurally were always inline. Kertas passes the
    // floating layout (anchor, wrap, behind text) so watermarks can sit in the header.
    dir: 'superdoc/dist/chunks',
    find: /(function materializeInlineImage\(schema, node\) \{\n\tconst payload = resolvePayload\(node, "image"\);\n\tconst attrs = \{ src: payload\.src \};)/,
    replace: '$1\n\tif (payload.pmAttrs && typeof payload.pmAttrs === "object") Object.assign(attrs, payload.pmAttrs); // kertas: image layout',
    marker: 'kertas: image layout',
  },
  {
    // superdoc 1.47.2: a new header or footer declared only the w and r namespaces, so a
    // picture inside it (a watermark) made the saved file unreadable in Word and LibreOffice
    dir: 'superdoc/dist/chunks',
    find: /(name: kind === "header" \? "w:hdr" : "w:ftr",\n\t\tattributes: \{\n\t\t\t"xmlns:w": WORDPROCESSINGML_XMLNS,\n\t\t\t"xmlns:r": OFFICE_DOCUMENT_RELS_XMLNS)/,
    replace: '$1, "xmlns:wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main", "xmlns:pic": "http://schemas.openxmlformats.org/drawingml/2006/picture", "xmlns:wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape", "xmlns:wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing", "xmlns:w14": "http://schemas.microsoft.com/office/word/2010/wordml", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:m": "http://schemas.openxmlformats.org/officeDocument/2006/math", "xmlns:v": "urn:schemas-microsoft-com:vml", "xmlns:o": "urn:schemas-microsoft-com:office:office" /* kertas: header namespaces */',
    marker: 'kertas: header namespaces',
  },
  {
    // ...and it got no relationships part, so pictures in it lost their link to the media file
    dir: 'superdoc/dist/chunks',
    find: /if \(sourceSnapshot\.relsPart\) operations\.push\(\{\n(\t+)editor,\n\t+partId: toRelsPathForPart\(newPartPath\),\n\t+operation: "create",\n\t+source: "createHeaderFooterPart",\n\t+initial: sourceSnapshot\.relsPart\n\t+\}\);/,
    replace: 'operations.push({ editor, partId: toRelsPathForPart(newPartPath), operation: "create", source: "createHeaderFooterPart", initial: sourceSnapshot.relsPart ?? { elements: [{ type: "element", name: "Relationships", attributes: { xmlns: RELS_XMLNS }, elements: [] }] } }); // kertas: header rels',
    marker: 'kertas: header rels',
  },
  {
    // ...and when the document API wrote a header, the relationships its export created (the
    // picture's link) were thrown away
    dir: 'superdoc/dist/chunks',
    find: /const \{ result \} = converter\.exportToXmlJson\(\{\n(\t+)data: pmJson,([\s\S]{0,400}?)bodyContent = result\?\.elements\?\.\[0\]\?\.elements \?\? \[\];/,
    replace: 'const { result, params: kertasParams } = converter.exportToXmlJson({\n$1data: pmJson,$2bodyContent = result?.elements?.[0]?.elements ?? [];\n$1if (kertasParams?.relationships?.length) converter.convertedXml[relsPartId] = { declaration: converter.initialJSON?.declaration, elements: [{ type: "element", name: "Relationships", attributes: { xmlns: "http://schemas.openxmlformats.org/package/2006/relationships" }, elements: mergeRelationshipElements(existingRelationships, kertasParams.relationships) }] }; // kertas: header picture rels',
    marker: 'kertas: header picture rels',
  },
  {
    // superdoc 1.47.2: every tracked change was saved with w:authorEmail, which is not part of
    // Word's file format; it is now written only when there is an address (Kertas never sets one)
    dir: 'superdoc/dist/chunks',
    find: /"w:authorEmail": trackedMark\.attrs\.authorEmail,/g,
    replace: '...(trackedMark.attrs.authorEmail ? { "w:authorEmail": trackedMark.attrs.authorEmail } : {}), /* kertas: no empty authorEmail */',
    marker: 'kertas: no empty authorEmail',
  },
  {
    // superdoc 1.47.2: the ruler only knew inches. Kertas sets globalThis.__kertasRulerUnit to
    // "cm" for a centimetre ruler (labels every centimetre, half and quarter ticks).
    dir: 'superdoc/dist/chunks',
    find: /const ticks = \[\];\n(\t)let currentX = 0;\n\tfor \(let inch = 0; inch < (pageSize\.width|pageWidthInches); inch\+\+\) \{/g,
    replace: 'const ticks = [];\n$1let currentX = 0;\n$1if (globalThis.__kertasRulerUnit === "cm") { const pxPerCm = ppi / 2.54; const totalCm = ($2) * 2.54; for (let q = 0; q * 0.25 <= totalCm + 1e-9; q++) { const x = q * 0.25 * pxPerCm; ticks.push(q % 4 === 0 ? { size: "main", height: "20%", label: q / 4, x } : q % 2 === 0 ? { size: "half", height: "40%", x } : { size: "eighth", height: "10%", x }); } } else /* kertas: cm ruler */\n$1for (let inch = 0; inch < $2; inch++) {',
    marker: 'kertas: cm ruler',
  },
  {
    // superdoc 1.47.2: paragraph and run properties were saved in the order they were added
    // (an applied style could land after the paragraph mark's run properties). Word requires the
    // schema order and may refuse such files, so they are sorted on save.
    dir: 'superdoc/dist/chunks',
    find: /(const elements = decodeProperties\(params, propertyTranslatorsBySdName\$\d+, currentValue\);)\n(\t+)if \(elements\.length === 0\) return;/,
    replace: '$1\n$2{ const order = ({ "w:pPr": ["w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange"], "w:rPr": ["w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange"] })[xmlName]; if (order) elements.sort((a, b) => { const ia = order.indexOf(a.name); const ib = order.indexOf(b.name); return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib); }); } // kertas: schema order\n$2if (elements.length === 0) return;',
    marker: 'kertas: schema order',
  },
  {
    // Speed: every keystroke rebuilt the layout of every table in the document. An unchanged
    // table is now reused: it keeps ids of its own (so the painter knows it is the same table and
    // does not redraw it), and only its positions move when text before it changed.
    dir: 'superdoc/dist/chunks',
    find: /function handleTableNode\(node, context\) \{\n\tconst \{ blocks, recordBlockKind, nextBlockId, positions, trackedChangesConfig, bookmarks, hyperlinkConfig, converters: converters\$1, converterContext, enableComments \} = context;\n\tconst tableBlock = tableNodeToBlock\(node, \{\n\t\tnextBlockId,/,
    replace: `function kertasShiftBlock(value, delta) {
\tif (Array.isArray(value)) return value.map((v) => kertasShiftBlock(v, delta));
\tif (value && typeof value === "object") {
\t\tconst out = {};
\t\tfor (const k in value) {
\t\t\tconst v = value[k];
\t\t\tif ((k === "pmStart" || k === "pmEnd") && typeof v === "number") out[k] = v + delta;
\t\t\telse out[k] = typeof v === "object" && v !== null ? kertasShiftBlock(v, delta) : v;
\t\t}
\t\treturn out;
\t}
\treturn value;
}
function handleTableNode(node, context) {
\tconst { blocks, recordBlockKind, nextBlockId, positions, trackedChangesConfig, bookmarks, hyperlinkConfig, converters: converters$1, converterContext, enableComments } = context;
\tconst kertasCache = context.flowBlockCache ? (context.flowBlockCache.kertasTables ??= new Map()) : null; // kertas: table cache
\tconst kertasSeen = context.flowBlockCache ? (context.flowBlockCache.kertasSeen ??= new Set()) : null;
\tlet kertasKey = node.attrs?.sdBlockId ?? node.attrs?.paraId ?? null;
\tif (kertasKey != null && kertasSeen) { if (kertasSeen.has(kertasKey)) kertasKey = null; else kertasSeen.add(kertasKey); }
\tconst kertasStart = positions.get(node)?.start ?? 0;
\tif (kertasCache && kertasKey != null) {
\t\tconst hit = kertasCache.get(kertasKey);
\t\tif (hit && (hit.node === node || hit.json === JSON.stringify(node))) {
\t\t\tconst delta = kertasStart - hit.pmStart;
\t\t\tconst block = delta === 0 ? hit.block : kertasShiftBlock(hit.block, delta);
\t\t\tblocks.push(block);
\t\t\trecordBlockKind?.(block.kind);
\t\t\tkertasCache.set(kertasKey, { node, json: hit.json, block, pmStart: kertasStart });
\t\t\treturn;
\t\t}
\t}
\tlet kertasN = 0;
\tconst kertasPrefix = \`\${context.blockIdPrefix ?? ""}t\${kertasKey}.\`;
\tconst kertasNextId = (kind) => \`\${kertasPrefix}\${kertasN++}-\${kind}\`;
\tconst tableBlock = tableNodeToBlock(node, {
\t\tnextBlockId: kertasCache && kertasKey != null ? kertasNextId : nextBlockId,`,
    marker: 'kertas: table cache',
  },
  {
    // ...store what was built
    dir: 'superdoc/dist/chunks',
    find: /(\t\tenableComments\n\t\}\);\n\tif \(tableBlock\) \{\n\t\tblocks\.push\(tableBlock\);\n\t\trecordBlockKind\?\.\(tableBlock\.kind\);)/,
    replace: '$1\n\t\tif (kertasCache && kertasKey != null) kertasCache.set(kertasKey, { node, json: JSON.stringify(node), block: tableBlock, pmStart: kertasStart }); // kertas: table cache store',
    marker: 'kertas: table cache store',
  },
  {
    // ...each layout pass starts with no table seen yet
    dir: 'superdoc/dist/chunks',
    find: /(\tbegin\(\) \{\n\t\tthis\.#next\.clear\(\);)/,
    replace: '$1\n\t\tthis.kertasSeen?.clear(); // kertas: table cache pass',
    marker: 'kertas: table cache pass',
  },
  {
    // Speed: spell check re-read every paragraph of the document on every keystroke. Squiggles
    // now move with the text at once, and the document is re-read once typing pauses.
    dir: 'superdoc/dist/chunks',
    find: /\tonDocumentChanged\(doc\$1, changedRanges, mapping\) \{\n\t\tif \(!this\.isEnabled\) return;\n\t\tthis\.#documentEpoch\+\+;/,
    replace: '\tonDocumentChanged(doc$1, changedRanges, mapping) {\n\t\tif (!this.isEnabled) return;\n\t\tthis.#documentEpoch++;\n\t\tthis.#store.kertasMapAll?.(mapping); // kertas: spell batch\n\t\tconst q = this.kertasQueue ??= { ranges: [] };\n\t\tq.ranges = q.ranges.map((r) => ({ from: mapping.map(r.from, -1), to: mapping.map(r.to, 1) })).concat(changedRanges);\n\t\tq.doc = doc$1;\n\t\tq.identity = new mapping.constructor();\n\t\tclearTimeout(q.timer);\n\t\tq.timer = setTimeout(() => { const go = () => { this.kertasQueue = null; this.kertasOnDocumentChangedNow(q.doc, q.ranges, q.identity); }; typeof requestIdleCallback === "function" ? requestIdleCallback(go, { timeout: 600 }) : go(); }, 400);\n\t}\n\tkertasOnDocumentChangedNow(doc$1, changedRanges, mapping) {\n\t\tif (!this.isEnabled) return;\n\t\tthis.#documentEpoch++;',
    marker: 'kertas: spell batch',
  },
  {
    // ...all existing issues follow edits (upstream only moved those in the edited paragraph)
    dir: 'superdoc/dist/chunks',
    find: /(\tremapIssues\(segmentIds, mapping, recheckId\) \{)/,
    replace: '\tkertasMapAll(mapping) { // kertas: spell map all\n\t\tfor (const [id, list] of this.#issuesBySegment) {\n\t\t\tconst keep = [];\n\t\t\tfor (const issue of list) {\n\t\t\t\tconst a = mapping.mapResult(issue.pmFrom, -1);\n\t\t\t\tconst b = mapping.mapResult(issue.pmTo, 1);\n\t\t\t\tif (a.deleted || b.deleted || a.pos >= b.pos) continue;\n\t\t\t\tissue.pmFrom = a.pos;\n\t\t\t\tissue.pmTo = b.pos;\n\t\t\t\tkeep.push(issue);\n\t\t\t}\n\t\t\tif (keep.length) this.#issuesBySegment.set(id, keep);\n\t\t\telse this.#issuesBySegment.delete(id);\n\t\t}\n\t}\n$1',
    marker: 'kertas: spell map all',
  },
  {
    // Speed: a table's paint version included the document position of every run, so typing
    // anywhere above a table redrew the whole table. Positions are patched in place instead,
    // exactly as for paragraphs (whose version never included them).
    dir: 'superdoc/dist/chunks',
    find: /hash\$3 = hashNumber\(hash\$3, run\.pmStart \?\? -1\);\n(\t+)hash\$3 = hashNumber\(hash\$3, run\.pmEnd \?\? -1\);/g,
    replace: '/* kertas: table version without positions */',
    marker: 'kertas: table version without positions',
  },
  {
    // Speed: after each keystroke the page update waited for the next animation frame (up to
    // 16 ms of nothing). It now starts right after the keystroke, so the result is ready for the
    // very next frame; several changes in one task are still combined into one update.
    dir: 'superdoc/dist/chunks',
    find: /(\t#scheduleRerender\(\) \{\n\t\tif \(this\.#renderScheduled\) return;\n\t\tthis\.#renderScheduled = true;\n)\t\tthis\.#rafHandle = \(this\.#visibleHost\.ownerDocument\?\.defaultView \?\? window\)\.requestAnimationFrame\(\(\) => \{\n\t\t\tthis\.#renderScheduled = false;/,
    replace: '$1\t\tqueueMicrotask(() => { // kertas: render without waiting a frame\n\t\t\tif (!this.#renderScheduled || !this.#visibleHost) return;\n\t\t\tthis.#renderScheduled = false;',
    marker: 'kertas: render without waiting a frame',
  },
  {
    // Speed: the editing-permission plugin re-scanned the whole document after every edit, even
    // in documents without permission ranges. It now skips the scan when there are none and the
    // edit did not add any.
    dir: 'superdoc/dist/chunks',
    find: /(\t\t\t\tapply\(tr, value, _oldState, newState\) \{\n\t\t\t\t\tif \(!tr\.docChanged\) return value;\n)(\t\t\t\t\tconst permissionTypeInfo = getPermissionTypeInfo\(newState\.schema\);)/,
    replace: '$1\t\t\t\t\tif (value && Array.isArray(value.allRanges) && value.allRanges.length === 0 && !globalThis.__kertasStepsAddPerm?.(tr)) return value; // kertas: permission scan skip\n$2',
    marker: 'kertas: permission scan skip',
  },
  {
    dir: 'superdoc/dist/chunks',
    find: /(\t\t\tappendTransaction\(transactions, oldState, newState\) \{\n\t\t\t\tif \(!transactions\.some\(\(tr\$1\) => tr\$1\.docChanged\)\) return null;\n)/,
    replace: '$1\t\t\t\t{ const kv = PERMISSION_PLUGIN_KEY.getState(oldState); if (kv && Array.isArray(kv.allRanges) && kv.allRanges.length === 0 && !transactions.some((t) => globalThis.__kertasStepsAddPerm?.(t))) return null; } // kertas: permission append skip\n',
    marker: 'kertas: permission append skip',
  },
  {
    // the check both of the above use: did this edit insert a permission marker?
    dir: 'superdoc/dist/chunks',
    find: /(var buildPermissionState = \(doc\$1, allowedIdentifiers)/,
    replace: 'globalThis.__kertasStepsAddPerm ??= (tr) => tr.steps.some((st) => { let found = false; st.slice?.content?.descendants?.((n) => { if (/^perm(Start|End)/.test(n.type.name)) found = true; return !found; }); return found; }); // kertas: permission step check\n$1',
    marker: 'kertas: permission step check',
  },
  {
    // Speed: after every paint SuperDoc re-read the style of every line on every mounted page for
    // a debugging snapshot nothing in the app reads (the click and caret index uses only the
    // snapshot's fields, content controls and pictures, which are still collected).
    // Set globalThis.__kertasFullPaintSnapshot = true to get the full snapshot back.
    dir: 'superdoc/dist/chunks',
    find: /const lineElements = Array\.from\(pageEl\.querySelectorAll\("\.superdoc-line"\)\);/,
    replace: 'const lineElements = globalThis.__kertasFullPaintSnapshot ? Array.from(pageEl.querySelectorAll(".superdoc-line")) : []; // kertas: lean paint snapshot',
    marker: 'kertas: lean paint snapshot',
  },
  {
    // Speed: a plugin turned every step of every transaction into JSON just to find
    // "remove formatting" steps. Only steps that carry a mark can be one.
    dir: 'superdoc/dist/chunks',
    find: /(\t\tt\.steps\.forEach\(\(step\) => \{\n)(\t\t\tconst jsonStep = step\.toJSON\?\.\(\);\n\t\t\tif \(jsonStep\?\.stepType !== "removeMark")/,
    replace: '$1\t\t\tif (!step.mark) return; // kertas: mark steps only\n$2',
    marker: 'kertas: mark steps only',
  },
  {
    // Speed: when a transaction had many steps (opening a long document stamps an id on every
    // paragraph), the table-style plugin mapped every step through every later one (quadratic).
    // It now simply checks the whole document in that case.
    dir: 'superdoc/dist/chunks',
    find: /(\t+)(for \(let i = 0; i < allMaps\.length; i\+\+\) allMaps\[i\]\.forEach\(\(_oldFrom, _oldTo, newFrom, newTo\) => \{\n\t+for \(let j = i \+ 1; j < allMaps\.length; j\+\+\) \{)/,
    replace: '$1if (allMaps.length > 24) ranges.push(0, newState.doc.content.size); else /* kertas: no quadratic mapping */\n$1$2',
    marker: 'kertas: no quadratic mapping',
  },
  {
    // Speed: changed ranges were found by mapping every step through every other step of the
    // transaction (quadratic; slow when opening a long document). With many steps the whole
    // document counts as changed, which is a safe superset.
    dir: 'superdoc/dist/chunks',
    find: /(\ttransactions\.forEach\(\(tr, index\) => \{\n\t\tif \(!tr\.docChanged\) return;\n)(\t\tconst perTransactionRanges = \[\];)/,
    replace: '$1\t\tif (tr.steps.length > 24) { ranges.push({ from: 0, to: docSize }); return; } // kertas: linear changed ranges\n$2',
    marker: 'kertas: linear changed ranges',
  },
  {
    // ...same for the tab plugin's list of paragraphs to re-measure
    dir: 'superdoc/dist/chunks',
    find: /(function getAffectedParagraphStarts\(tr, newState\) \{\n\tconst affected = \/\* @__PURE__ \*\/ new Set\(\);\n)/,
    replace: '$1\tif (tr.steps.length > 24) { newState.doc.descendants((node, pos) => { if (node.type.name === "paragraph") { affected.add(pos); return false; } return true; }); return affected; } // kertas: linear tab paragraphs\n',
    marker: 'kertas: linear tab paragraphs',
  },
  {
    // Safety for the JSON memo: SuperDoc's Word export edits the document JSON it is given
    // (for example it strips tracked-change marks as it writes them). It now always works on
    // its own copy, so saving can never alter the shared JSON the editor reuses.
    dir: 'superdoc/dist/chunks',
    find: /(\texportToXmlJson\(\{ data, editorSchema[^\n]*\) \{\n)/,
    replace: '$1\t\tif (data && typeof data === "object") data = typeof structuredClone === "function" ? structuredClone(data) : JSON.parse(JSON.stringify(data)); // kertas: export on a copy\n',
    marker: 'kertas: export on a copy',
  },
  {
    // ...and forget them whenever SuperDoc forgets its own cache (styles, theme or fonts changed)
    dir: 'superdoc/dist/chunks',
    find: /setHasExternalChanges\(value\) \{\n(\t+)this\.#hasExternalChanges = value;\n\t+\}\n\t+get\(id, node\) \{/,
    replace: 'setHasExternalChanges(value) {\n$1this.#hasExternalChanges = value;\n$1if (value) this.kertasTables?.clear(); // kertas: table cache reset\n\t}\n\tget(id, node) {',
    marker: 'kertas: table cache reset',
  },
  {
    dir: 'superdoc/dist/chunks',
    find: /(\tclear\(\) \{\n\t\tthis\.#previous\.clear\(\);\n\t\tthis\.#next\.clear\(\);)/,
    replace: '$1\n\t\tthis.kertasTables?.clear(); // kertas: table cache clear',
    marker: 'kertas: table cache clear',
  },
  {
    // superdoc 1.47.2: the toolbar recomputed every command's state after every keystroke
    // (about 13 ms each on a long document). It now waits for a short pause in typing;
    // clicking a toolbar button still refreshes it at once.
    dir: 'superdoc/dist/chunks',
    find: /const handleChange = \(\) => \{\n\t\tif \(destroyed \|\| refreshQueued\) return;\n\t\trefreshQueued = true;\n\t\tqueueMicrotask\(\(\) => \{\n\t\t\trefreshQueued = false;\n\t\t\tif \(destroyed\) return;\n\t\t\trefreshControllerState\(\);\n\t\t\}\);\n\t\};/,
    replace: 'let kertasTimer = null;\n\tconst handleChange = () => {\n\t\tif (destroyed) return;\n\t\tclearTimeout(kertasTimer);\n\t\tkertasTimer = setTimeout(() => { const run = () => { if (!destroyed) refreshControllerState(); }; typeof requestIdleCallback === "function" ? requestIdleCallback(run, { timeout: 250 }) : run(); }, 260); // kertas: toolbar after typing pause\n\t};',
    marker: 'kertas: toolbar after typing pause',
  },
  {
    // Speed: document nodes that did not change between keystrokes are the same objects, so
    // their JSON is remembered instead of rebuilt; only the edited paragraph is serialised again
    // (SuperDoc turned the whole document into JSON twice per keystroke).
    dir: 'superdoc/dist/chunks',
    find: /\ttoJSON\(\) \{\n\t\tlet obj = \{ type: this\.type\.name \};\n\t\tfor \(let _ in this\.attrs\) \{\n\t\t\tobj\.attrs = this\.attrs;\n\t\t\tbreak;\n\t\t\}\n\t\tif \(this\.content\.size\) obj\.content = this\.content\.toJSON\(\);\n\t\tif \(this\.marks\.length\) obj\.marks = this\.marks\.map\(\(n\) => n\.toJSON\(\)\);\n\t\treturn obj;\n\t\}/,
    replace: '\ttoJSON() {\n\t\tconst kertasMemo = globalThis.__kertasJsonMemo ??= new WeakMap(); // kertas: json memo\n\t\tconst kertasHit = kertasMemo.get(this);\n\t\tif (kertasHit) return kertasHit;\n\t\tlet obj = { type: this.type.name };\n\t\tfor (let _ in this.attrs) {\n\t\t\tobj.attrs = this.attrs;\n\t\t\tbreak;\n\t\t}\n\t\tif (this.content.size) obj.content = this.content.toJSON();\n\t\tif (this.marks.length) obj.marks = this.marks.map((n) => n.toJSON());\n\t\tif (this.isText) obj.text = this.text;\n\t\tif (globalThis.__kertasFreezeJson) Object.freeze(obj);\n\t\tkertasMemo.set(this, obj);\n\t\treturn obj;\n\t}',
    marker: 'kertas: json memo',
  },
  {
    // ...the text node's own toJSON adds its text to the (now shared) object; it already has it
    dir: 'superdoc/dist/chunks',
    find: /\ttoJSON\(\) \{\n\t\tlet base\$1 = super\.toJSON\(\);\n\t\tbase\$1\.text = this\.text;\n\t\treturn base\$1;\n\t\}/,
    replace: '\ttoJSON() {\n\t\treturn super.toJSON(); // kertas: json memo text\n\t}',
    marker: 'kertas: json memo text',
  },
  {
    // Speed: "is this command available?" rebuilt the editor's whole command list on every
    // check, dozens of times per keystroke. An editor's commands are fixed once it starts.
    dir: 'superdoc/dist/chunks',
    find: /function hasCommand\(editor, command\$1\) \{\n\treturn typeof editor\.commands\?\.\[command\$1\] === "function";\n\}/,
    replace: 'function hasCommand(editor, command$1) {\n\tif (!editor) return false;\n\tconst kertasMemo = globalThis.__kertasCommandNames ??= new WeakMap(); // kertas: command names\n\tlet names = kertasMemo.get(editor);\n\tif (!names) { const all = editor.commands; if (!all) return false; names = new Set(Object.keys(all).filter((k) => typeof all[k] === "function")); kertasMemo.set(editor, names); }\n\treturn names.has(command$1);\n}',
    marker: 'kertas: command names',
  },
  {
    // Speed: the old page-break adjustment forced a browser layout on every render even though
    // the current layout engine has no such breaks
    dir: 'superdoc/dist/chunks',
    find: /function adjustPaginationBreaks\(editorElem, editor\) \{\n\tif \(!editorElem\.value \|\| !editor\?\.value\?\.options\?\.scale \|\| isHeadless\(editor\)\) return;/,
    replace: 'function adjustPaginationBreaks(editorElem, editor) {\n\tif (!editorElem.value || !editor?.value?.options?.scale || isHeadless(editor)) return;\n\tif (!editorElem.value.querySelector(".pagination-break-wrapper")) return; // kertas: no forced layout',
    marker: 'kertas: no forced layout',
  },
  {
    // superdoc 1.47.2: spell checking splits a run of text into pieces so it can underline a
    // misspelt word. Text after a tab is placed at a fixed position, and every piece kept that
    // same position, so all the words after a tab were drawn on top of each other. The pieces
    // now share one positioned wrapper and flow inside it.
    dir: 'superdoc/dist/chunks',
    find: /(\toriginalSpan\.removeAttribute\("data-pm-end"\);\n)\tfor \(const sibling of siblings\) parent\.insertBefore\(sibling, originalSpan\);\n\treturn siblings;/,
    replace: '$1\tif (originalSpan.style.position === "absolute") { // kertas: proofing keeps tab positions\n\t\tconst wrap = originalSpan.ownerDocument.createElement("span");\n\t\twrap.setAttribute("data-sd-proofing-wrap", "");\n\t\twrap.style.position = "absolute"; wrap.style.left = originalSpan.style.left; wrap.style.top = originalSpan.style.top; wrap.style.whiteSpace = "pre";\n\t\tfor (const sibling of siblings) { sibling.style.position = "static"; sibling.style.left = ""; sibling.style.top = ""; wrap.appendChild(sibling); }\n\t\tparent.insertBefore(wrap, originalSpan);\n\t} else for (const sibling of siblings) parent.insertBefore(sibling, originalSpan);\n\treturn siblings;',
    marker: 'kertas: proofing keeps tab positions',
  },
  {
    // ...and taking the spell-check pieces away again removes their wrapper too
    dir: 'superdoc/dist/chunks',
    find: /\t\tunhideOriginalSpan\(original\);\n\t\tfor \(const sibling of siblings\) parent\.removeChild\(sibling\);\n\t\}\n([\s\S]{0,400}?for \(const element of hiddenOriginals\) unhideOriginalSpan\(element\);\n)\treturn true;/,
    replace: '\t\tunhideOriginalSpan(original);\n\t\tfor (const sibling of siblings) sibling.remove(); // kertas: proofing wrap removal\n\t}\n$1\tfor (const wrap of Array.from(container.querySelectorAll("[data-sd-proofing-wrap]"))) wrap.remove();\n\treturn true;',
    marker: 'kertas: proofing wrap removal',
  },
  {
    // superdoc 1.47.2: resolving a comment through the document API changed the document but not
    // the comment list that saving reads, so a resolved comment was saved as still open. The list
    // now takes the "resolved" update too (reopening already sent an ordinary update).
    dir: 'superdoc/dist/chunks',
    find: /if \(COMMENT_EVENTS\?\.UPDATE && type === COMMENT_EVENTS\.UPDATE && commentPayload\) \{\n(\t+)const \{ id, existingComment \} = resolveDocumentScopedCommentMatch\(commentPayload\);\n(\t+)if \(id\) \{\n(\t+)const resolvedText = commentPayload\.commentText \|\| commentPayload\.text;\n(\t+)if \(existingComment\) applyReplayUpdateToComment\(existingComment, commentPayload, resolvedText\);\n(\t+)else \{/,
    replace: 'if (COMMENT_EVENTS?.UPDATE && (type === COMMENT_EVENTS.UPDATE || type === COMMENT_EVENTS.RESOLVED) && commentPayload) { // kertas: resolved comments saved\n$1const { id, existingComment } = resolveDocumentScopedCommentMatch(commentPayload);\n$2if (id) {\n$3const resolvedText = commentPayload.commentText || commentPayload.text;\n$4if (existingComment) applyReplayUpdateToComment(existingComment, commentPayload, resolvedText);\n$5else if (type === COMMENT_EVENTS.UPDATE) {',
    marker: 'kertas: resolved comments saved',
  },
  // @univerjs/engine-formula 1.4.0: a formula inside a conditional format (or another
  // non-cell formula) was given a made-up cell position near the top-left of the sheet. Real
  // cells there then looked like they depended on it, the engine reported a circular reference,
  // and those cells were overwritten with 0 (in a Gantt chart, the dates and blank cells the
  // colour rule looks at turned into 0). Such formulas now sit in no cell.
  ...['lib', 'lib/es', 'lib/cjs'].map((sub) => ({
    dir: `@univerjs/engine-formula/${sub}`,
    find: /\ttoRTreeItem\(\) \{\n/g,
    replace: '\ttoRTreeItem() {\n\t\tif (this.type === 1) return []; // kertas: other formulas have no cell\n',
    marker: 'kertas: other formulas have no cell',
  })),
];

for (const p of patches) {
  const files = p.files || (fs.existsSync(path.join(root, p.dir))
    ? fs.readdirSync(path.join(root, p.dir)).filter((f) => /\.(c|m)?js$/.test(f)).map((f) => `${p.dir}/${f}`)
      .filter((rel) => { const t = fs.readFileSync(path.join(root, rel), 'utf8'); p.find.lastIndex = 0; return p.find.test(t) || t.includes(p.marker); })
    : []);
  if (p.dir && !files.length) console.warn(`patch-deps: nothing to patch in ${p.dir} (package changed?)`);
  for (const rel of files) {
    const file = path.join(root, rel);
    if (!fs.existsSync(file)) continue;
    const src = fs.readFileSync(file, 'utf8');
    if (src.includes(p.marker)) continue;
    p.find.lastIndex = 0;
    const out = src.replace(p.find, p.replace);
    if (out === src) { console.warn(`patch-deps: pattern not found in ${rel}`); continue; }
    fs.writeFileSync(file, out);
    console.log(`patch-deps: patched ${rel}`);
  }
}
