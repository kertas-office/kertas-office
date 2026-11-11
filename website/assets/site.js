// Kertas Office website: small enhancements on top of pages that work without JavaScript.
// Live data (latest release, changelog, roadmap, known issues) comes from /api/*, served by
// the Cloudflare Pages Functions in website/functions. If those fail, the static text stays.

// Set to true once Windows installers are signed through SignPath. It switches the install
// steps, FAQ and policy text from the "unsigned" wording to the "signed" wording.
const SIGNED = false;
document.documentElement.dataset.signed = String(SIGNED);

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const mb = (bytes) => `${Math.round(bytes / 1048576)} MB`;
const fmtDate = (iso) => new Date(iso).toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

async function api(path) {
  try {
    const res = await fetch(path, { headers: { accept: 'application/json' } });
    if (!res.ok) return null;
    return await res.json();
  } catch { return null; }
}

function detectOS() {
  const ua = navigator.userAgent;
  const plat = navigator.userAgentData?.platform || navigator.platform || '';
  if (/android|iphone|ipad|ipod/i.test(ua)) return 'mobile';
  if (/win/i.test(plat) || /windows/i.test(ua)) return 'windows';
  if (/mac/i.test(plat) || /mac os/i.test(ua)) return 'mac';
  if (/linux|x11|cros/i.test(plat + ua)) return 'linux';
  return 'other';
}

const REPO = 'https://github.com/kertas-office/kertas-office';
const LATEST = `${REPO}/releases/latest/download/`;
const FILES = {
  windows: { file: 'Kertas-Office-Setup.exe', label: 'Download for Windows' },
  mac: { file: 'Kertas-Office-mac-arm64.dmg', label: 'Download for Mac' },
  linux: { file: 'Kertas-Office-linux-x86_64.AppImage', label: 'Download for Linux' },
};

function header() {
  const h = $('.site-header');
  const onScroll = () => h.classList.toggle('scrolled', window.scrollY > 4);
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });
  const btn = $('.menu-btn');
  const nav = $('#mobile-nav');
  btn?.addEventListener('click', () => {
    const open = btn.getAttribute('aria-expanded') !== 'true';
    btn.setAttribute('aria-expanded', String(open));
    nav.classList.toggle('open', open);
  });
  nav?.addEventListener('click', (e) => { if (e.target.closest('a')) { btn.setAttribute('aria-expanded', 'false'); nav.classList.remove('open'); } });
}

function platformCards(os) {
  // Home: highlight the visitor's system and make its button the primary one
  const card = $(`[data-platform="${os}"]`);
  if (card) {
    card.classList.add('on');
    const tag = $('[data-detected]', card);
    if (tag) tag.hidden = false;
    const b = $('.btn', card);
    if (b) { b.classList.remove('btn-outline'); b.classList.add('btn-grad'); }
  }
  // Home hero button: go straight to the right file
  const label = $('[data-os-label]');
  if (label && FILES[os]) label.textContent = FILES[os].label;
}

function downloadPage(os) {
  const primary = $('[data-primary-dl]');
  if (!primary) return;
  const pick = FILES[os] ? os : 'windows';
  primary.href = LATEST + FILES[pick].file;
  $('[data-primary-label]').textContent = FILES[pick].label;
  $('[data-primary-file]').innerHTML = `${FILES[pick].file}<span data-size="${FILES[pick].file}"></span>`;
  $$('[data-install-steps]').forEach((el) => { el.hidden = el.dataset.installSteps !== pick; });
  if (os === 'mac') $('[data-mac-note]').hidden = false;
}

async function release() {
  if (!$('[data-release-version], [data-size], [data-size-cell]')) return;
  const r = await api('/api/release');
  if (!r?.version) return;
  $$('[data-release-version]').forEach((el) => { el.textContent = /^v/.test(el.textContent.trim()) ? `v${r.version}` : r.version; });
  $$('[data-release-date]').forEach((el) => { el.textContent = r.date ? ` · ${fmtDate(r.date)}` : ''; });
  const size = (name) => r.assets?.find((a) => a.name === name)?.size;
  $$('[data-size]').forEach((el) => { const s = size(el.dataset.size); if (s) el.textContent = ` · ${mb(s)}`; });
  $$('[data-size-cell]').forEach((el) => { const s = size(el.dataset.sizeCell); if (s) el.textContent = mb(s); });
}

// A small, safe Markdown renderer for release notes: headings, lists, bold, code and links
function md(text) {
  const inline = (s) => esc(s)
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g, '<a href="$2">$1</a>');
  const out = [];
  let list = false;
  for (const raw of String(text || '').replace(/\r/g, '').split('\n')) {
    const line = raw.trimEnd();
    const li = /^\s*[-*]\s+(.*)/.exec(line);
    if (li) { if (!list) { out.push('<ul>'); list = true; } out.push(`<li>${inline(li[1])}</li>`); continue; }
    if (list) { out.push('</ul>'); list = false; }
    const h = /^(#{1,4})\s+(.*)/.exec(line);
    if (h) {
      const label = h[2].trim();
      const app = { pdf: ['P', 'var(--pdf)'], writer: ['W', 'var(--writer)'], sheets: ['S', 'var(--sheets)'] }[label.toLowerCase()];
      out.push(`<h3>${app ? `<span class="group-ic" style="background:${app[1]}">${app[0]}</span>` : ''}${inline(label)}</h3>`);
    } else if (line.trim()) out.push(`<p>${inline(line)}</p>`);
  }
  if (list) out.push('</ul>');
  return out.join('');
}

async function changelog() {
  const box = $('[data-releases]');
  if (!box) return;
  const list = await api('/api/releases');
  if (!Array.isArray(list) || !list.length) return;
  box.innerHTML = list.map((r, i) => `
    <article class="release">
      <div class="when">
        <span class="vbadge">v${esc(r.version)}</span>
        <div class="small" style="margin-top:10px">${r.date ? esc(fmtDate(r.date)) : ''}</div>
        ${i === 0 ? '<div class="small" style="margin-top:4px">Latest</div>' : ''}
      </div>
      <div class="notes">
        <div class="title">${esc(r.name || `Kertas Office ${r.version}`)}</div>
        ${md(r.notes)}
        <p><a href="${esc(r.url)}">Files and full notes on GitHub</a></p>
      </div>
    </article>`).join('');
  $('[data-releases-live]')?.classList.add('on');
}

async function roadmap() {
  const lanes = $$('[data-lane]');
  const known = $('[data-known-list]');
  if (!lanes.length && !known) return;
  const data = await api('/api/roadmap');
  if (!data) return;
  const item = (i) => `<li><a href="${esc(i.url)}">${esc(i.title)}</a></li>`;
  if (lanes.length && ['now', 'next', 'later'].some((k) => data[k]?.length)) {
    lanes.forEach((ul) => {
      const items = data[ul.dataset.lane] || [];
      ul.innerHTML = items.length ? items.map(item).join('') : '<li class="empty">Nothing here yet</li>';
    });
    $('[data-roadmap-live]')?.classList.add('on');
  }
  if (known && data.known?.length) {
    known.innerHTML = data.known.map((i) => `<a href="${esc(i.url)}"><span class="num">#${esc(i.number)}</span><span>${esc(i.title)}</span></a>`).join('');
    $('[data-known-section]').hidden = false;
    $('[data-live]')?.classList.add('on');
  }
}

function tour() {
  const tabs = $$('#tour [role="tab"]');
  if (!tabs.length) return;
  const glow = $('[data-tour-glow]');
  const select = (tab, focus) => {
    tabs.forEach((t) => {
      const on = t === tab;
      t.setAttribute('aria-selected', String(on));
      t.tabIndex = on ? 0 : -1;
      $(`#${t.getAttribute('aria-controls')}`).hidden = !on;
    });
    const c = tab.dataset.glow;
    if (glow) glow.style.background = `radial-gradient(circle at 18% 20%, ${c}30, transparent 50%), radial-gradient(circle at 82% 92%, ${c}22, transparent 50%), #f8f8fb`;
    if (focus) tab.focus();
  };
  tabs.forEach((t, i) => {
    t.addEventListener('click', () => select(t));
    t.addEventListener('keydown', (e) => {
      const d = { ArrowRight: 1, ArrowLeft: -1 }[e.key];
      if (d) { e.preventDefault(); select(tabs[(i + d + tabs.length) % tabs.length], true); }
      if (e.key === 'Home') { e.preventDefault(); select(tabs[0], true); }
      if (e.key === 'End') { e.preventDefault(); select(tabs[tabs.length - 1], true); }
    });
  });
  select(tabs[0]);
}

function helpSearch() {
  const input = $('#help-search');
  if (!input) return;
  const articles = $$('[data-articles] .article');
  const empty = $('[data-search-empty]');
  const topics = $('[data-help-topics]');
  const text = new Map(articles.map((a) => [a, a.textContent.toLowerCase()]));
  input.addEventListener('input', () => {
    const words = input.value.toLowerCase().split(/\s+/).filter(Boolean);
    let shown = 0;
    articles.forEach((a) => {
      const hit = words.every((w) => text.get(a).includes(w));
      a.hidden = !hit;
      if (hit) shown++;
    });
    empty.hidden = shown > 0;
    topics.hidden = words.length > 0;
  });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { const first = articles.find((a) => !a.hidden); first?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
  });
}

function policyNav() {
  const links = $$('[data-policy-nav] a');
  if (!links.length) return;
  const sections = links.map((a) => $(a.getAttribute('href')));
  const io = new IntersectionObserver((entries) => {
    entries.forEach((e) => {
      if (!e.isIntersecting) return;
      links.forEach((a) => a.classList.toggle('active', a.getAttribute('href') === `#${e.target.id}`));
    });
  }, { rootMargin: '-30% 0px -60% 0px' });
  sections.forEach((s) => s && io.observe(s));
}

document.addEventListener('DOMContentLoaded', () => {
  const os = detectOS();
  header();
  platformCards(os);
  downloadPage(os);
  tour();
  helpSearch();
  policyNav();
  release();
  changelog();
  roadmap();
});
