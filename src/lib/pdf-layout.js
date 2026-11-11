// Reading order and structure from a PDF's positioned text: lines, paragraphs,
// headings, lists and tables. Feeds PDF to Word and PDF to Excel. Pure module.
//
// A page is { width, height, items: [{ str, x, y, w, size, bold, italic, font }] }
// with x, y in points from the top-left corner (y is the baseline).

const median = (a) => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[Math.floor(s.length / 2)]; };
const BULLET = /^\s*(?:[•▪◦●○■□·]\s*|[–\-*]\s+)(?=\S)/;
const NUMBERED = /^\s*((?:\d{1,3}|[a-zA-Z]|[ivxIVX]{1,4})[.)])\s+/;
const NUMBER = /^\(?-?[$€£¥₹]?\s?(?:RM|USD|QAR|SGD|MYR|INR)?\s?-?\d[\d,]*(?:\.\d+)?\)?\s?%?$/;

/** Groups items into lines and splits each line into segments at wide gaps */
export function buildLines(page) {
  const items = page.items.filter((it) => it.str && it.str.trim() !== '' && it.size > 0)
    .sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [];
  for (const it of items) {
    // A line holds items whose baselines are close and that overlap vertically
    let line = null;
    for (let i = lines.length - 1; i >= 0 && i >= lines.length - 4; i--) {
      const l = lines[i];
      if (Math.abs(l.y - it.y) <= Math.max(l.size, it.size) * 0.4) { line = l; break; }
    }
    if (!line) { line = { y: it.y, size: it.size, items: [] }; lines.push(line); }
    line.items.push(it);
    line.size = Math.max(line.size, it.size);
  }
  for (const l of lines) {
    l.items.sort((a, b) => a.x - b.x);
    l.segments = [];
    let seg = null;
    for (const it of l.items) {
      const gap = seg ? it.x - seg.x1 : 0;
      if (!seg || gap > Math.max(it.size * 1.6, 14)) {
        seg = { x0: it.x, x1: it.x + it.w, runs: [], size: it.size };
        l.segments.push(seg);
      } else if (gap > it.size * 0.12 && !/\s$/.test(seg.runs.at(-1)?.text || '') && !/^\s/.test(it.str)) {
        pushRun(seg.runs, ' ', it);
      }
      pushRun(seg.runs, it.str, it);
      seg.x1 = Math.max(seg.x1, it.x + it.w);
      seg.size = Math.max(seg.size, it.size);
    }
    for (const s of l.segments) s.text = s.runs.map((r) => r.text).join('').replace(/\s+/g, ' ').trim();
    l.x0 = l.segments[0].x0;
    l.x1 = Math.max(...l.segments.map((s) => s.x1));
    l.text = l.segments.map((s) => s.text).join(' ');
    const chars = l.items.reduce((n, it) => n + it.str.length, 0) || 1;
    l.bold = l.items.reduce((n, it) => n + (it.bold ? it.str.length : 0), 0) / chars > 0.6;
  }
  return lines.filter((l) => l.text);
}

function pushRun(runs, text, it) {
  const last = runs.at(-1);
  const style = { bold: !!it.bold, italic: !!it.italic, size: Math.round(it.size * 2) / 2, font: it.font || null };
  if (last && last.bold === style.bold && last.italic === style.italic && last.size === style.size && last.font === style.font) last.text += text;
  else runs.push({ text, ...style });
}

/** Font sizes used for headings, largest first, given the body size */
function headingLevels(allLines, body) {
  const sizes = [...new Set(allLines.filter((l) => l.size >= body * 1.18 && l.text.length < 140).map((l) => Math.round(l.size)))].sort((a, b) => b - a);
  return (size) => {
    const i = sizes.indexOf(Math.round(size));
    return i < 0 ? 0 : Math.min(3, i + 1);
  };
}

/** Column bands for a block of table-like lines: x ranges separated by whitespace shared by every line */
export function columnBands(lines) {
  const iv = lines.flatMap((l) => l.segments.map((s) => [s.x0, s.x1])).sort((a, b) => a[0] - b[0]);
  const bands = [];
  for (const [a, b] of iv) {
    const last = bands.at(-1);
    if (last && a <= last[1] + 4) last[1] = Math.max(last[1], b);
    else bands.push([a, b]);
  }
  return bands;
}
const bandOf = (bands, s) => {
  const c = (s.x0 + s.x1) / 2;
  let best = 0;
  bands.forEach((b, i) => { if (c >= b[0] - 2) best = i; });
  return best;
};

/**
 * Structure of a page: [{ type: 'heading', level, runs } | { type: 'para', runs, align, indent }
 *   | { type: 'list', ordered, items: [runs] } | { type: 'table', rows: [[text]] }]
 * bodySize: the document's usual text size (computed over all pages by analyse()).
 */
export function pageBlocks(page, { bodySize, level }) {
  const lines = buildLines(page);
  const blocks = [];
  const left = Math.min(...lines.map((l) => l.x0));
  const right = Math.max(...lines.map((l) => l.x1));
  const centre = page.width / 2;
  let i = 0;
  while (i < lines.length) {
    const l = lines[i];
    // Tables: two or more consecutive lines that each have several separated segments
    if (l.segments.length >= 2) {
      let j = i;
      while (j < lines.length && lines[j].segments.length >= 2 && (j === i || lines[j].y - lines[j - 1].y < lines[j].size * 3.2)) j++;
      if (j - i >= 2) {
        const rowsLines = lines.slice(i, j);
        const bands = columnBands(rowsLines);
        if (bands.length >= 2) {
          const rows = rowsLines.map((rl) => {
            const row = Array(bands.length).fill('');
            for (const s of rl.segments) { const k = bandOf(bands, s); row[k] = row[k] ? `${row[k]} ${s.text}` : s.text; }
            return row;
          });
          blocks.push({ type: 'table', rows, header: rowsLines[0].bold });
          i = j;
          continue;
        }
      }
    }
    const lv = level(l.size);
    if (lv || (l.bold && l.text.length < 90 && !/[.,;:]$/.test(l.text) && l.size >= bodySize * 0.95 && (i + 1 >= lines.length || lines[i + 1].y - l.y > l.size * 1.5))) {
      blocks.push({ type: 'heading', level: lv || 3, runs: lineRuns(l), align: alignOf(l, left, right, centre) });
      i++;
      continue;
    }
    const bullet = l.text.match(BULLET);
    const numbered = !bullet && l.text.match(NUMBERED);
    if (bullet || numbered) {
      const ordered = !!numbered;
      const items = [];
      while (i < lines.length) {
        const cur = lines[i];
        const m = ordered ? cur.text.match(NUMBERED) : cur.text.match(BULLET);
        if (!m) break;
        const runs = stripPrefix(lineRuns(cur), m[0]);
        i++;
        // Wrapped continuation lines of this item (indented, no marker)
        while (i < lines.length && !lines[i].text.match(BULLET) && !lines[i].text.match(NUMBERED) && lines[i].x0 > cur.x0 + 2 && lines[i].y - lines[i - 1].y < lines[i].size * 1.8 && lines[i].segments.length === 1) {
          appendRuns(runs, lineRuns(lines[i]));
          i++;
        }
        items.push(runs);
      }
      blocks.push({ type: 'list', ordered, items });
      continue;
    }
    // Ordinary paragraph: join following lines that continue it
    const runs = lineRuns(l);
    const align = alignOf(l, left, right, centre);
    let prev = l;
    i++;
    while (i < lines.length) {
      const n = lines[i];
      const gap = n.y - prev.y;
      const sameSize = Math.abs(n.size - prev.size) <= prev.size * 0.12;
      const continues = sameSize && n.segments.length === 1 && gap < prev.size * 1.75 && !level(n.size)
        && Math.abs(n.x0 - left) < prev.size * 3 + (align === 'center' ? 1e3 : 0)
        && !n.text.match(BULLET) && !n.text.match(NUMBERED)
        // A short line that ends a sentence usually ends the paragraph
        && !(prev.x1 < right - prev.size * 4 && /[.!?:]["')\]]?$/.test(prev.text));
      if (!continues) break;
      appendRuns(runs, lineRuns(n));
      prev = n;
      i++;
    }
    const indent = align === 'left' && l.x0 - left > bodySize * 1.5 ? Math.round(l.x0 - left) : 0;
    blocks.push({ type: 'para', runs, align, indent });
  }
  return blocks;
}

function lineRuns(l) {
  const runs = [];
  l.segments.forEach((s, k) => {
    if (k) runs.push({ ...s.runs[0], text: ' ' });
    for (const r of s.runs) runs.push({ ...r });
  });
  return runs;
}
function appendRuns(runs, more) {
  const last = runs.at(-1);
  // Words broken with a hyphen at the end of a line are joined again
  if (last && /[A-Za-z]-$/.test(last.text) && /^[a-z]/.test(more[0]?.text || '')) last.text = last.text.slice(0, -1);
  else if (last && !/\s$/.test(last.text)) last.text += ' ';
  for (const r of more) {
    const prev = runs.at(-1);
    if (prev && prev.bold === r.bold && prev.italic === r.italic && prev.size === r.size && prev.font === r.font) prev.text += r.text;
    else runs.push({ ...r });
  }
}
function stripPrefix(runs, prefix) {
  let n = prefix.length;
  const out = runs.map((r) => ({ ...r }));
  while (n > 0 && out.length) {
    const r = out[0];
    if (r.text.length <= n) { n -= r.text.length; out.shift(); } else { r.text = r.text.slice(n); n = 0; }
  }
  return out;
}
function alignOf(l, left, right, centre) {
  const mid = (l.x0 + l.x1) / 2;
  if (l.x0 - left > 30 && Math.abs(mid - centre) < l.size * 1.5 && l.x1 < right - 20) return 'center';
  if (l.x0 - left > (right - left) * 0.45 && Math.abs(l.x1 - right) < l.size) return 'right';
  return 'left';
}

/** Whole-document analysis: body size, heading levels, blocks per page */
export function analyse(pages) {
  const all = pages.flatMap((p) => buildLines(p));
  const weighted = all.flatMap((l) => Array(Math.min(50, Math.ceil(l.text.length / 10))).fill(l.size));
  const bodySize = median(weighted) || 11;
  const level = headingLevels(all, bodySize);
  return { bodySize, pages: pages.map((p) => ({ width: p.width, height: p.height, blocks: pageBlocks(p, { bodySize, level }), empty: !p.items.some((it) => it.str.trim()) })) };
}

// ---------------------------------------------------------------- spreadsheet

/** "1,234.50" "(12)" "RM 9" "45%" -> { value, format } or null */
export function parseNumber(text) {
  const t = String(text).trim();
  if (!NUMBER.test(t)) return null;
  const neg = /^\(.*\)$/.test(t) || /^-|\(-|\s-/.test(t);
  const pct = /%$/.test(t);
  const digits = t.replace(/[^\d.]/g, '');
  if (!digits || digits === '.') return null;
  let value = Number(digits);
  if (!Number.isFinite(value)) return null;
  if (neg) value = -value;
  if (pct) value /= 100;
  const decimals = (digits.split('.')[1] || '').length;
  const comma = /,/.test(t);
  const format = pct ? (decimals ? `0.${'0'.repeat(decimals)}%` : '0%') : comma || decimals ? `${comma ? '#,##0' : '0'}${decimals ? `.${'0'.repeat(decimals)}` : ''}` : null;
  return { value, format };
}

/**
 * Rows of cells for a page: tables keep their columns, other lines go in column A.
 * Returns [[string]] (empty strings for blank cells).
 */
export function pageRows(page) {
  const lines = buildLines(page);
  const rows = [];
  let i = 0;
  while (i < lines.length) {
    if (lines[i].segments.length >= 2) {
      let j = i;
      while (j < lines.length && lines[j].segments.length >= 2) j++;
      const block = lines.slice(i, j);
      const bands = columnBands(block);
      for (const l of block) {
        const row = Array(bands.length).fill('');
        for (const s of l.segments) { const k = bandOf(bands, s); row[k] = row[k] ? `${row[k]} ${s.text}` : s.text; }
        rows.push(row);
      }
      i = j;
    } else {
      rows.push([lines[i].text]);
      i++;
    }
  }
  return rows;
}
