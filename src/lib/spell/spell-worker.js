// Runs Hunspell dictionaries (nspell) off the main thread.
// Messages: { id, type: 'load', lang, aff, dic } | { id, type: 'check', lang, words } | { id, type: 'suggest', lang, word }
//           | { id, type: 'personal', words }
import nspell from 'nspell';

const spells = new Map(); // lang -> nspell
let personal = new Set();

self.onmessage = async (e) => {
  const { id, type } = e.data;
  try {
    if (type === 'load') {
      if (!spells.has(e.data.lang)) {
        const [aff, dic] = await Promise.all([fetch(e.data.aff).then((r) => r.text()), fetch(e.data.dic).then((r) => r.text())]);
        spells.set(e.data.lang, nspell(aff, dic));
      }
      self.postMessage({ id, ok: true });
    } else if (type === 'personal') {
      personal = new Set(e.data.words.map((w) => w.toLowerCase()));
      self.postMessage({ id, ok: true });
    } else if (type === 'check') {
      const sp = spells.get(e.data.lang);
      const bad = sp ? e.data.words.filter((w) => !personal.has(w.toLowerCase()) && !sp.correct(w) && !sp.correct(w.replace(/’/g, "'"))) : [];
      self.postMessage({ id, ok: true, bad });
    } else if (type === 'suggest') {
      const sp = spells.get(e.data.lang);
      self.postMessage({ id, ok: true, suggestions: sp ? suggest(sp, e.data.word) : [] });
    }
  } catch (err) {
    self.postMessage({ id, ok: false, error: String(err?.message || err) });
  }
};

// Swapped letters ("teh", "recieve") are the most common typo, and Hunspell
// often ranks the fix low, so those candidates go first.
function suggest(sp, word) {
  const swaps = [];
  for (let i = 0; i < word.length - 1; i++) {
    const w = word.slice(0, i) + word[i + 1] + word[i] + word.slice(i + 2);
    if (w !== word && sp.correct(w)) swaps.push(w);
  }
  const out = [];
  for (const w of [...swaps, ...sp.suggest(word)]) if (!out.includes(w)) out.push(w);
  return out.slice(0, 6);
}
