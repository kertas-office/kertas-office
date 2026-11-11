// Equations for Word: LaTeX -> MathML (temml) -> Office Math (OMML).
// The MathML -> OMML step is our own, covering what temml produces for
// everyday maths: fractions, scripts, roots, big operators, brackets,
// matrices, accents and bars. Pure module, works in Node and the browser.
import temml from 'temml';

export const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// ------------------------------------------------------------------ tiny XML reader
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' };
const decode = (s) => s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => (e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : Number(e.slice(1))) : ENT[e] ?? m));
export function parseXml(src) {
  const root = { name: '#root', attrs: {}, children: [] };
  const stack = [root];
  const re = /<!--[\s\S]*?-->|<\/([\w:.-]+)\s*>|<([\w:.-]+)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(src))) {
    const top = stack.at(-1);
    if (m[1]) { if (stack.length > 1) stack.pop(); continue; }
    if (m[2]) {
      const attrs = {};
      for (const a of m[3].matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) attrs[a[1]] = decode(a[2] ?? a[3]);
      const node = { name: m[2].replace(/^\w+:/, ''), attrs, children: [] };
      top.children.push(node);
      if (!m[4]) stack.push(node);
      continue;
    }
    if (m[5] !== undefined && m[5].trim() !== '') top.children.push({ name: '#text', text: decode(m[5]) });
  }
  return root;
}
const textOf = (n) => (n.name === '#text' ? n.text : n.children.map(textOf).join(''));
const kids = (n) => n.children.filter((c) => c.name !== '#text');

// ------------------------------------------------------------------ OMML writer
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';
const run = (text, plain = false) => `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}${FONT}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
const NARY = new Set(['∑', '∏', '∐', '∫', '∬', '∭', '∮', '∯', '∰', '⋃', '⋂', '⋁', '⋀', '⨁', '⨂', '⨀']);
const INTEGRALS = new Set(['∫', '∬', '∭', '∮', '∯', '∰']);
// Spacing accents (what temml emits) and the combining marks Word expects
const ACCENTS = { 'ˆ': '̂', '^': '̂', '˜': '̃', '~': '̃', '¯': '̄', '˙': '̇', '¨': '̈', 'ˇ': '̌', '´': '́', '`': '̀', '˘': '̆', '→': '⃗', '←': '⃖', '⃗': '⃗', '˚': '̊' };
const BARS = new Set(['‾', '¯', '_', '―', '─', '̲', '̅']);
const STOP = /^[=+\-−±∓<>≤≥≠≈≡∼∝,;]$/;

const e = (inner) => `<m:e>${inner}</m:e>`;
const isOp = (n, set) => n && n.name === 'mo' && set.has(textOf(n).trim());
const unwrap = (n) => { while (n && n.name === 'mrow' && kids(n).length === 1) n = kids(n)[0]; return n; };
const baseOp = (n) => { n = unwrap(n); return ['msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover'].includes(n?.name) ? kids(n)[0] : n; };

function nary(node, operand) {
  const [b, a1, a2] = kids(node);
  const chr = textOf(b).trim();
  let sub = ''; let sup = '';
  if (node.name === 'msub' || node.name === 'munder') sub = conv(a1);
  else if (node.name === 'msup' || node.name === 'mover') sup = conv(a1);
  else if (node.name === 'msubsup' || node.name === 'munderover') { sub = conv(a1); sup = conv(a2); }
  const under = node.name.startsWith('mu') || (!INTEGRALS.has(chr) && b.attrs.movablelimits !== 'false' && node.name !== 'msubsup');
  const loc = INTEGRALS.has(chr) ? 'subSup' : under || !INTEGRALS.has(chr) ? 'undOvr' : 'subSup';
  return `<m:nary><m:naryPr><m:chr m:val="${esc(chr)}"/><m:limLoc m:val="${loc}"/>${sub ? '' : '<m:subHide m:val="1"/>'}${sup ? '' : '<m:supHide m:val="1"/>'}</m:naryPr><m:sub>${sub}</m:sub><m:sup>${sup}</m:sup>${e(operand)}</m:nary>`;
}

/** Converts a list of sibling nodes, folding big operators over what follows them */
function seq(list) {
  let out = '';
  for (let i = 0; i < list.length; i++) {
    const n = unwrap(list[i]);
    const b = baseOp(n);
    if (isOp(b, NARY)) {
      // The operand runs to the next relation or plus/minus at this level
      let j = i + 1;
      const operand = [];
      while (j < list.length && !(list[j].name === 'mo' && STOP.test(textOf(list[j]).trim()))) operand.push(list[j++]);
      out += nary(n === b ? { name: 'msub', children: [b, { name: 'mrow', attrs: {}, children: [] }], attrs: {} } : n, seq(operand));
      i = j - 1;
      continue;
    }
    out += conv(n);
  }
  return out;
}

function fenced(n) {
  const c = kids(n);
  if (c.length < 2) return null;
  const [first, last] = [c[0], c.at(-1)];
  const isFence = (x, form) => x.name === 'mo' && (x.attrs.fence === 'true' || x.attrs.form === form) && /^[([{|‖⟨⌈⌊)\]}⟩⌉⌋.]?$/.test(textOf(x).trim());
  if (!isFence(first, 'prefix') || !isFence(last, 'postfix')) return null;
  const open = textOf(first).trim(); const close = textOf(last).trim();
  if (open === '' && close === '') return null;
  // Separators inside (\middle| or commas in mfenced) stay as text
  return `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/></m:dPr>${e(seq(c.slice(1, -1)))}</m:d>`;
}

function conv(n) {
  if (!n) return '';
  const c = kids(n);
  switch (n.name) {
    case '#text': return run(n.text);
    case 'math': case 'mrow': case 'mstyle': case 'mpadded': case 'semantics': case 'mtd': {
      if (n.name === 'semantics') return conv(c[0]);
      if (n.name === 'mrow') { const d = fenced(n); if (d) return d; }
      return seq(c);
    }
    case 'annotation': case 'annotation-xml': case 'mphantom': return '';
    case 'mi': {
      const t = textOf(n);
      return run(t, n.attrs.mathvariant === 'normal' || [...t].length > 1);
    }
    case 'mn': return run(textOf(n), true);
    case 'mo': return run(textOf(n).replace(/⁡/g, ''), true);
    case 'mtext': case 'ms': return run(textOf(n), true);
    case 'mspace': return Number.parseFloat(n.attrs.width) >= 0.1 ? run(' ', true) : '';
    case 'mfrac': {
      const bar = n.attrs.linethickness && Number.parseFloat(n.attrs.linethickness) === 0 ? '<m:fPr><m:type m:val="noBar"/></m:fPr>' : '';
      return `<m:f>${bar}<m:num>${conv(c[0])}</m:num><m:den>${conv(c[1])}</m:den></m:f>`;
    }
    case 'msup': return isOp(c[0], NARY) ? nary(n, '') : `<m:sSup>${e(conv(c[0]))}<m:sup>${conv(c[1])}</m:sup></m:sSup>`;
    case 'msub': return isOp(c[0], NARY) ? nary(n, '') : `<m:sSub>${e(conv(c[0]))}<m:sub>${conv(c[1])}</m:sub></m:sSub>`;
    case 'msubsup': return isOp(c[0], NARY) ? nary(n, '') : `<m:sSubSup>${e(conv(c[0]))}<m:sub>${conv(c[1])}</m:sub><m:sup>${conv(c[2])}</m:sup></m:sSubSup>`;
    case 'msqrt': return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(seq(c))}</m:rad>`;
    case 'mroot': return `<m:rad><m:deg>${conv(c[1])}</m:deg>${e(conv(c[0]))}</m:rad>`;
    case 'mover': case 'munder': {
      if (isOp(c[0], NARY)) return nary(n, '');
      const mark = textOf(c[1]).trim();
      const top = n.name === 'mover';
      if (top && ACCENTS[mark] && c[1].name === 'mo') return `<m:acc><m:accPr><m:chr m:val="${ACCENTS[mark]}"/></m:accPr>${e(conv(c[0]))}</m:acc>`;
      if (BARS.has(mark)) return `<m:bar><m:barPr><m:pos m:val="${top ? 'top' : 'bot'}"/></m:barPr>${e(conv(c[0]))}</m:bar>`;
      if (/^[⏞⏟︷︸]$/.test(mark)) return `<m:groupChr><m:groupChrPr><m:chr m:val="${mark}"/><m:pos m:val="${top ? 'top' : 'bot'}"/></m:groupChrPr>${e(conv(c[0]))}</m:groupChr>`;
      return top ? `<m:limUpp>${e(conv(c[0]))}<m:lim>${conv(c[1])}</m:lim></m:limUpp>` : `<m:limLow>${e(conv(c[0]))}<m:lim>${conv(c[1])}</m:lim></m:limLow>`;
    }
    case 'munderover':
      if (isOp(c[0], NARY)) return nary(n, '');
      return `<m:limUpp>${e(`<m:limLow>${e(conv(c[0]))}<m:lim>${conv(c[1])}</m:lim></m:limLow>`)}<m:lim>${conv(c[2])}</m:lim></m:limUpp>`;
    case 'menclose': {
      const note = n.attrs.notation || '';
      if (/top|overline/.test(note)) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${e(seq(c))}</m:bar>`;
      if (/bottom|underline/.test(note)) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr>${e(seq(c))}</m:bar>`;
      if (/box|roundedbox/.test(note)) return `<m:borderBox>${e(seq(c))}</m:borderBox>`;
      if (/strike|horizontalstrike|updiagonal/.test(note)) return `<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/><m:strikeH m:val="1"/></m:borderBoxPr>${e(seq(c))}</m:borderBox>`;
      return seq(c);
    }
    case 'mtable': {
      const rows = c.filter((r) => r.name === 'mtr' || r.name === 'mlabeledtr');
      const cols = Math.max(1, ...rows.map((r) => kids(r).length));
      const mcs = `<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cols}"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr>`;
      return `<m:m>${mcs}${rows.map((r) => `<m:mr>${Array.from({ length: cols }, (_, k) => e(conv(kids(r)[k]))).join('')}</m:mr>`).join('')}</m:m>`;
    }
    case 'mfenced': {
      const open = n.attrs.open ?? '('; const close = n.attrs.close ?? ')';
      return `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/></m:dPr>${c.map((x) => e(conv(x))).join('')}</m:d>`;
    }
    default: return seq(c);
  }
}

/** MathML string -> <m:oMath> */
export function mathmlToOmml(mathml) {
  const root = parseXml(mathml);
  const math = kids(root).find((n) => n.name === 'math') || root;
  return `<m:oMath>${conv(math)}</m:oMath>`;
}

/** LaTeX -> { omml (paragraph-level <m:oMathPara>), mathml } ; throws with a readable message */
export function latexToOmml(latex, { display = true } = {}) {
  let mathml;
  try {
    mathml = temml.renderToString(latex, { displayMode: display, throwOnError: true, annotate: false });
  } catch (err) {
    throw new Error(String(err.message || err).replace(/^Temml parse error:\s*/i, ''));
  }
  const omath = mathmlToOmml(mathml);
  return { mathml, omath, paragraph: `<w:p>${display ? `<m:oMathPara>${omath}</m:oMathPara>` : omath}</w:p>` };
}

export const EQUATION_PRESETS = [
  { label: 'Quadratic formula', tex: 'x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}' },
  { label: 'Pythagoras', tex: 'a^2+b^2=c^2' },
  { label: 'Area of a circle', tex: 'A=\\pi r^2' },
  { label: 'Sum', tex: '\\sum_{i=1}^{n} i=\\frac{n(n+1)}{2}' },
  { label: 'Integral', tex: '\\int_{0}^{1} x^2\\,dx=\\frac{1}{3}' },
  { label: 'Binomial theorem', tex: '(x+a)^n=\\sum_{k=0}^{n}\\binom{n}{k}x^k a^{n-k}' },
  { label: 'Matrix', tex: 'A=\\begin{pmatrix}a & b\\\\ c & d\\end{pmatrix}' },
  { label: 'Limit', tex: '\\lim_{x\\to 0}\\frac{\\sin x}{x}=1' },
  { label: 'Compound interest', tex: 'A=P\\left(1+\\frac{r}{n}\\right)^{nt}' },
  { label: 'Standard deviation', tex: '\\sigma=\\sqrt{\\frac{1}{N}\\sum_{i=1}^{N}(x_i-\\mu)^2}' },
];
