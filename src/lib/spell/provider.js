// Spell checking for the Writer: a SuperDoc proofing provider backed by
// Hunspell dictionaries in a worker. SuperDoc draws the red underlines and the
// right-click suggestions; we only say which words are wrong.
import { wordsIn, dictionaryFor } from './words.js';
import enUSaff from '../../../node_modules/dictionary-en/index.aff?url';
import enUSdic from '../../../node_modules/dictionary-en/index.dic?url';
import enGBaff from '../../../node_modules/dictionary-en-gb/index.aff?url';
import enGBdic from '../../../node_modules/dictionary-en-gb/index.dic?url';
import frAff from '../../../node_modules/dictionary-fr/index.aff?url';
import frDic from '../../../node_modules/dictionary-fr/index.dic?url';
import esAff from '../../../node_modules/dictionary-es/index.aff?url';
import esDic from '../../../node_modules/dictionary-es/index.dic?url';
import deAff from '../../../node_modules/dictionary-de/index.aff?url';
import deDic from '../../../node_modules/dictionary-de/index.dic?url';

const DICTS = {
  'en-US': [enUSaff, enUSdic], 'en-GB': [enGBaff, enGBdic], fr: [frAff, frDic], es: [esAff, esDic], de: [deAff, deDic],
};
const abs = (u) => new URL(u, document.baseURI).href;

const PERSONAL_KEY = 'kertas-personal-dictionary';
export function personalWords() { try { return JSON.parse(localStorage.getItem(PERSONAL_KEY) || '[]'); } catch { return []; } }
function savePersonal(list) { try { localStorage.setItem(PERSONAL_KEY, JSON.stringify([...new Set(list)].sort())); } catch { /* storage off */ } }

let worker = null;
let seq = 0;
const pending = new Map();
function call(msg) {
  if (!worker) {
    worker = new Worker(new URL('./spell-worker.js', import.meta.url), { type: 'module' });
    worker.onmessage = (e) => { const p = pending.get(e.data.id); if (!p) return; pending.delete(e.data.id); e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error)); };
    worker.postMessage({ id: 0, type: 'personal', words: personalWords() });
  }
  const id = ++seq;
  return new Promise((resolve, reject) => { pending.set(id, { resolve, reject }); worker.postMessage({ ...msg, id }); });
}
const loaded = new Map();
function ensure(lang) {
  if (!loaded.has(lang)) loaded.set(lang, call({ type: 'load', lang, aff: abs(DICTS[lang][0]), dic: abs(DICTS[lang][1]) }).catch((e) => { loaded.delete(lang); throw e; }));
  return loaded.get(lang);
}

export function addToPersonal(word) {
  const list = [...personalWords(), word];
  savePersonal(list);
  if (worker) call({ type: 'personal', words: list });
}

/** The provider SuperDoc calls. getLanguage() returns the user's chosen dictionary, or null for "follow the document". */
export function createSpellProvider(getLanguage) {
  const langOf = (segLang, def) => getLanguage() || dictionaryFor(segLang) || dictionaryFor(def) || 'en-US';
  return {
    id: 'kertas-hunspell',
    getCapabilities: () => ({ issueKinds: ['spelling'], supportsSuggestions: true, supportsMultipleLanguages: true, supportsBatching: true, requiresNetwork: false }),
    async check({ segments, defaultLanguage }) {
      const issues = [];
      const byLang = new Map();
      for (const seg of segments) {
        const lang = langOf(seg.language, defaultLanguage);
        if (!DICTS[lang]) continue;
        if (!byLang.has(lang)) byLang.set(lang, []);
        byLang.get(lang).push(seg);
      }
      for (const [lang, segs] of byLang) {
        await ensure(lang);
        const found = segs.map((s) => ({ s, words: wordsIn(s.text) }));
        const unique = [...new Set(found.flatMap((f) => f.words.map((w) => w.word)))];
        if (!unique.length) continue;
        const { bad } = await call({ type: 'check', lang, words: unique });
        const badSet = new Set(bad);
        for (const { s, words } of found) {
          for (const w of words) if (badSet.has(w.word)) issues.push({ segmentId: s.id, start: w.start, end: w.end, kind: 'spelling', message: 'Possible spelling mistake' });
        }
      }
      return { issues };
    },
    async suggest({ segment, issue, defaultLanguage }) {
      const lang = langOf(segment.language, defaultLanguage);
      if (!DICTS[lang]) return { replacements: [] };
      await ensure(lang);
      const word = segment.text.slice(issue.start, issue.end);
      const { suggestions } = await call({ type: 'suggest', lang, word });
      return { replacements: suggestions };
    },
    dispose() {},
  };
}
