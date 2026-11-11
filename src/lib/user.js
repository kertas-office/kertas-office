// The name written on tracked changes and comments (Word: File > Options > User name).
const KEY = 'kertas-user-name';

export function userName() {
  try { return (localStorage.getItem(KEY) || '').trim(); } catch { return ''; }
}
export function setUserName(name) {
  const n = String(name || '').trim().slice(0, 80);
  try { if (n) localStorage.setItem(KEY, n); else localStorage.removeItem(KEY); } catch { /* private mode */ }
  window.dispatchEvent(new CustomEvent('kertas-user', { detail: n }));
  return n;
}
export const userInitials = (name) => String(name || '').split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0].toUpperCase()).join('') || 'ME';
/** The user object for the editor; no e-mail, so none is written into documents */
export const editorUser = () => ({ name: userName() || 'Me', email: '' });
