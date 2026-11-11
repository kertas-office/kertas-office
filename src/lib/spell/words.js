// Word finding for spell checking. Pure, shared by the worker and tests.

// Letters (any script), joined by apostrophes or hyphens inside a word
const WORD = /[\p{L}\p{M}]+(?:['’][\p{L}\p{M}]+)*/gu;

/** Words worth checking in a piece of text: [{ word, start, end }] */
export function wordsIn(text) {
  const out = [];
  for (const m of String(text).matchAll(WORD)) {
    const word = m[0];
    const start = m.index;
    const end = start + word.length;
    const before = text[start - 1] || '';
    const after = text[end] || '';
    // Skip parts of web addresses, emails, file names and code-like tokens
    if (/[@/\\_:\d]/.test(before)) continue;
    if (before === '.' && /[\p{L}\d]/u.test(text[start - 2] || '')) continue; // mail.com, file.txt
    if (/[@/\\_\d]/.test(after) || (after === ':' && text[end + 1] === '/') || (after === '.' && /[\p{L}\d]/u.test(text[end + 1] || ''))) continue;
    if (word.length < 2) continue;
    if (/^[\p{Lu}]+$/u.test(word) && word.length <= 6) continue; // acronyms (NASA, UPSC)
    if (/[\p{Ll}][\p{Lu}]/u.test(word)) continue; // camelCase / brand names (iPhone)
    out.push({ word, start, end });
  }
  return out;
}

/** Language codes we have dictionaries for, from a document or user language tag */
export function dictionaryFor(lang) {
  const l = String(lang || '').toLowerCase();
  if (l.startsWith('en-gb') || l.startsWith('en-au') || l.startsWith('en-nz') || l.startsWith('en-in') || l.startsWith('en-my') || l.startsWith('en-ie') || l.startsWith('en-za')) return 'en-GB';
  if (l.startsWith('en')) return 'en-US';
  if (l.startsWith('fr')) return 'fr';
  if (l.startsWith('es')) return 'es';
  if (l.startsWith('de')) return 'de';
  return null;
}

export const LANGUAGES = [
  { id: 'en-US', label: 'English (US)' },
  { id: 'en-GB', label: 'English (UK, Malaysia, India)' },
  { id: 'fr', label: 'French' },
  { id: 'es', label: 'Spanish' },
  { id: 'de', label: 'German' },
];
