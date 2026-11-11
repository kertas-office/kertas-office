// Citation and bibliography formatting: APA 7, MLA 9, Chicago (author-date),
// Harvard and IEEE. Sources use the fields of Word's Source Manager.
// Pure module.

export const CITATION_STYLES = [
  { id: 'apa', label: 'APA (7th)', heading: 'References' },
  { id: 'mla', label: 'MLA (9th)', heading: 'Works Cited' },
  { id: 'chicago', label: 'Chicago (author-date)', heading: 'References' },
  { id: 'harvard', label: 'Harvard', heading: 'Reference list' },
  { id: 'ieee', label: 'IEEE', heading: 'References' },
];
export const SOURCE_TYPES = [
  { value: 'book', label: 'Book' },
  { value: 'journalArticle', label: 'Journal article' },
  { value: 'website', label: 'Web page' },
  { value: 'report', label: 'Report' },
  { value: 'conferenceProceedings', label: 'Conference paper' },
  { value: 'thesis', label: 'Thesis' },
  { value: 'misc', label: 'Other' },
];

/** "Newport, Cal; Kaplan, Rachel" or "Cal Newport and Rachel Kaplan" -> [{ first, middle, last }] */
export function parseAuthors(text) {
  const parts = String(text || '').split(/\s*;\s*|\s+and\s+|\s*&\s*/i).map((s) => s.trim()).filter(Boolean);
  return parts.map((p) => {
    if (p.includes(',')) {
      const [last, rest = ''] = p.split(/\s*,\s*/);
      const [first = '', ...mid] = rest.split(/\s+/).filter(Boolean);
      return { last, first, middle: mid.join(' ') };
    }
    const words = p.split(/\s+/);
    if (words.length === 1) return { last: words[0] };
    const last = words.pop();
    const [first, ...mid] = words;
    return { last, first, middle: mid.join(' ') };
  });
}
export const formatAuthors = (list = []) => list.map((a) => [a.last, [a.first, a.middle].filter(Boolean).join(' ')].filter(Boolean).join(', ')).join('; ');

const initials = (a) => [a.first, a.middle].filter(Boolean).join(' ').split(/[\s-]+/).filter(Boolean).map((n) => `${n[0].toUpperCase()}.`).join(' ');
const full = (a) => [a.first, a.middle, a.last].filter(Boolean).join(' ');
const lastFirst = (a) => [a.last, [a.first, a.middle].filter(Boolean).join(' ')].filter(Boolean).join(', ');
const end = (s) => (/[.?!]$/.test(s) ? s : `${s}.`);
// APA and Harvard titles are kept as typed (forcing sentence case would lower-case proper nouns);
// only the first word and the word after a colon get a capital.
const sentenceCase = (t) => String(t || '').replace(/^(\p{Ll})|([:?!]\s+)(\p{Ll})/gu, (m, a, p, c) => (a ? a.toUpperCase() : p + c.toUpperCase()));
const titleCase = (t) => String(t || '').replace(/\b(\p{Ll})(\p{L}*)/gu, (m, a, b, off, s) => (off > 0 && /^(a|an|the|and|but|or|nor|for|of|on|in|at|to|by|with|as|vs)$/.test(m) && !/[:.]\s*$/.test(s.slice(0, off)) ? m : a.toUpperCase() + b));
const pages = (p) => String(p || '').replace(/\s*[\u2013\u2014]\s*/g, '-');

function joinNames(names, { and = '&', serial = true, two = null } = {}) {
  if (names.length <= 1) return names[0] || '';
  if (names.length === 2) return `${names[0]}${two ?? ` ${and} `}${names[1]}`;
  return `${names.slice(0, -1).join(', ')}${serial ? ',' : ''} ${and} ${names.at(-1)}`;
}

// ------------------------------------------------------------------ in-text
function shortAuthor(src, style) {
  const a = src.fields?.authors || [];
  const title = src.fields?.shortTitle || src.fields?.title || 'Untitled';
  if (!a.length) return style === 'mla' ? `"${title}"` : title;
  if (a.length === 1) return a[0].last;
  if (a.length === 2) return `${a[0].last} ${style === 'apa' ? '&' : 'and'} ${a[1].last}`;
  return `${a[0].last} et al.`;
}

/**
 * In-text citation for one or more sources. numbers: Map sourceId -> n (IEEE).
 * page: an optional locator ("12" or "12-14").
 */
export function inText(sources, style, { numbers, page } = {}) {
  if (!sources.length) return '';
  if (style === 'ieee') {
    const ns = sources.map((s) => numbers?.get(s.sourceId) ?? '?');
    return ns.map((n) => `[${n}${page && ns.length === 1 ? `, p. ${page}` : ''}]`).join(', ');
  }
  const loc = page ? (/[-,]/.test(page) ? `pp. ${page}` : `p. ${page}`) : '';
  const parts = sources.map((s) => {
    const who = shortAuthor(s, style);
    const year = s.fields?.year || 'n.d.';
    if (style === 'mla') return [who, page].filter(Boolean).join(' ');
    if (style === 'chicago') return `${who} ${year}${page ? `, ${page}` : ''}`;
    return `${who}, ${year}${loc ? `, ${loc}` : ''}`;
  });
  return `(${parts.join('; ')})`;
}

// ------------------------------------------------------------------ reference list
/** One reference as runs: [{ text, italic }] */
export function entry(src, style, n) {
  const f = src.fields || {};
  const type = src.type || 'book';
  const a = f.authors || [];
  const year = f.year || 'n.d.';
  const runs = [];
  const t = (text) => text && runs.push({ text });
  const i = (text) => text && runs.push({ text, italic: true });
  const journal = type === 'journalArticle';
  const web = type === 'website';
  const link = f.doi ? `https://doi.org/${f.doi.replace(/^https?:\/\/doi\.org\//, '')}` : f.url || '';
  const vol = (prefix = '') => [f.volume && `${prefix}${f.volume}`, f.issue && `(${f.issue})`].filter(Boolean).join('');

  if (style === 'apa') {
    const names = a.map((p) => `${p.last}${initials(p) ? `, ${initials(p)}` : ''}`);
    const who = a.length > 20 ? `${names.slice(0, 19).join(', ')}, ... ${names.at(-1)}` : joinNames(names, { and: '&', two: ', & ' });
    const title = sentenceCase(f.title || 'Untitled');
    if (who) t(`${end(who)} (${year}). `); else { i(end(title)); t(` (${year}).`); }
    if (journal) {
      if (who) t(`${end(title)} `);
      i(f.journalName || ''); if (f.volume) { t(', '); i(f.volume); } if (f.issue) t(`(${f.issue})`); if (f.pages) t(`, ${pages(f.pages)}`); t('.');
    } else {
      if (who) i(f.edition ? title : end(title));
      if (f.edition) t(` (${f.edition} ed.).`);
      if (f.publisher && f.publisher !== (a.length ? '' : null)) t(`${who ? ' ' : ' '}${end(f.publisher)}`);
    }
    if (link) t(` ${link}`);
    return runs;
  }
  if (style === 'mla') {
    const who = a.length === 0 ? '' : a.length === 1 ? lastFirst(a[0]) : a.length === 2 ? `${lastFirst(a[0])}, and ${full(a[1])}` : `${lastFirst(a[0])}, et al`;
    const title = titleCase(f.title || 'Untitled');
    if (who) t(`${end(who)} `);
    if (journal || web) {
      t(`"${end(title)}" `);
      const bits = [];
      if (journal) {
        i(f.journalName || '');
        if (f.volume) bits.push(`vol. ${f.volume}`);
        if (f.issue) bits.push(`no. ${f.issue}`);
        bits.push(year === 'n.d.' ? '' : year);
        if (f.pages) bits.push(`pp. ${pages(f.pages)}`);
        t(`${bits.filter(Boolean).length ? `, ${bits.filter(Boolean).join(', ')}` : ''}`);
      } else {
        i(f.publisher || ''); if (year !== 'n.d.') t(`, ${year}`);
      }
      if (link) t(`, ${link.replace(/^https?:\/\//, '')}`);
      t('.');
    } else {
      i(end(title));
      t(` ${[f.publisher, year !== 'n.d.' ? year : ''].filter(Boolean).join(', ')}.`);
    }
    return runs;
  }
  if (style === 'chicago') {
    const names = a.map((p, k) => (k === 0 ? lastFirst(p) : full(p)));
    const who = a.length > 10 ? `${names.slice(0, 7).join(', ')}, et al` : joinNames(names, { and: 'and', two: ', and ' });
    const title = titleCase(f.title || 'Untitled');
    if (who) t(`${end(who)} ${year}. `);
    if (journal) {
      t(`"${end(title)}" `); i(f.journalName || ''); if (f.volume) t(` ${f.volume}`); if (f.issue) t(` (${f.issue})`); if (f.pages) t(`: ${pages(f.pages)}`); t('.');
    } else {
      i(end(title)); if (!who) t(` ${year}.`);
      const pub = [f.city, f.publisher].filter(Boolean).join(': ');
      if (pub) t(` ${end(pub)}`);
    }
    if (link) t(` ${link}.`);
    return runs;
  }
  if (style === 'harvard') {
    const names = a.map((p) => `${p.last}${initials(p) ? `, ${initials(p).replace(/\s/g, '')}` : ''}`);
    const who = a.length >= 4 ? `${names[0]} et al.` : joinNames(names, { and: 'and', serial: false });
    const title = sentenceCase(f.title || 'Untitled');
    if (!who && !journal) {
      i(title); t(` (${year})`);
      if (f.publisher) t(` ${f.publisher}`);
      t('.');
      if (link) t(` Available at: ${link}.`);
      return runs;
    }
    t(`${who || ''}${who ? ' ' : ''}(${year}) `);
    if (journal) {
      t(`'${title}', `); i(f.journalName || ''); const v = vol(); if (v) t(`, ${v}`); if (f.pages) t(`, pp. ${pages(f.pages)}`); t('.');
    } else {
      i(title); t('.');
      const pub = [f.city, f.publisher].filter(Boolean).join(': ');
      if (pub) t(` ${end(pub)}`);
    }
    if (link) t(` Available at: ${link}.`);
    return runs;
  }
  // IEEE
  const names = a.map((p) => [initials(p), p.last].filter(Boolean).join(' '));
  const who = a.length > 6 ? `${names[0]} et al.` : joinNames(names, { and: 'and', serial: true });
  const title = f.title || 'Untitled';
  t(`[${n ?? 1}]\t`);
  if (who) t(`${who}, `);
  if (journal) {
    t(`"${title}," `); i(f.journalName || ''); const bits = [f.volume && `vol. ${f.volume}`, f.issue && `no. ${f.issue}`, f.pages && `pp. ${pages(f.pages)}`, year !== 'n.d.' && year].filter(Boolean); t(`${bits.length ? `, ${bits.join(', ')}` : ''}.`);
  } else if (web) {
    t(`"${title}." `); if (f.publisher) i(`${f.publisher}. `); if (link) t(`[Online]. Available: ${link}`);
    return runs;
  } else {
    i(title); t('. ');
    const pub = [f.city, f.publisher].filter(Boolean).join(': ');
    t(`${[pub, year !== 'n.d.' ? year : ''].filter(Boolean).join(', ')}.`);
  }
  if (link) t(` ${f.doi ? `doi: ${f.doi.replace(/^https?:\/\/doi\.org\//, '')}` : `[Online]. Available: ${link}`}`);
  return runs;
}

const sortKey = (s) => (s.fields?.authors?.[0]?.last || s.fields?.title || '').toLowerCase() + (s.fields?.year || '');

/** Sources in reference-list order. cited: sourceIds in order of first citation (for IEEE) */
export function ordered(sources, style, cited = []) {
  if (style === 'ieee') {
    const pos = (s) => { const k = cited.indexOf(s.sourceId); return k < 0 ? 1e6 + sources.indexOf(s) : k; };
    return [...sources].sort((x, y) => pos(x) - pos(y));
  }
  return [...sources].sort((x, y) => sortKey(x).localeCompare(sortKey(y)));
}

/** Numbering for IEEE: Map sourceId -> n */
export const numbering = (sources, cited) => new Map(ordered(sources, 'ieee', cited).map((s, k) => [s.sourceId, k + 1]));

const xmlEsc = (s) => String(s).replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
const runXml = (r) => `<w:r>${r.italic ? '<w:rPr><w:i/></w:rPr>' : ''}${r.text.split('\t').map((part, k) => `${k ? '<w:tab/>' : ''}${part ? `<w:t xml:space="preserve">${xmlEsc(part)}</w:t>` : ''}`).join('')}</w:r>`;

/** Heading and entries as Word paragraphs (OOXML) */
export function bibliographyXml(sources, style, cited = []) {
  const meta = CITATION_STYLES.find((s) => s.id === style) || CITATION_STYLES[0];
  const list = ordered(sources, style, cited);
  const pPr = style === 'ieee'
    ? '<w:pPr><w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs><w:spacing w:after="120"/><w:ind w:left="567" w:hanging="567"/></w:pPr>'
    : `<w:pPr><w:spacing w:after="120"${style === 'apa' || style === 'mla' ? ' w:line="360" w:lineRule="auto"' : ''}/><w:ind w:left="720" w:hanging="720"/></w:pPr>`;
  const heading = `<w:p><w:pPr><w:pStyle w:val="Heading1"/>${style === 'mla' || style === 'apa' ? '<w:jc w:val="center"/>' : ''}</w:pPr><w:r><w:t>${meta.heading}</w:t></w:r></w:p>`;
  return heading + list.map((s, k) => `<w:p>${pPr}${entry(s, style, k + 1).map(runXml).join('')}</w:p>`).join('');
}

/** Plain text of a reference (for previews and tests) */
export const entryText = (src, style, n) => entry(src, style, n).map((r) => r.text).join('');
