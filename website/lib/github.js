// Shared by the /api/* Pages Functions: reads public data from the GitHub API and caches it
// at Cloudflare's edge for a few minutes, so visitors never hit GitHub's rate limit.
// Set GITHUB_TOKEN in the Pages project (Settings > Variables and Secrets) to a fine-grained
// token with read-only access to public repositories; it raises the limit and is never sent
// to browsers. Without it, the functions still work at a lower rate limit.

export const REPO = 'kertas-office/kertas-office';
const TTL = 600; // seconds

export async function gh(env, path) {
  const headers = { accept: 'application/vnd.github+json', 'user-agent': 'kertasoffice.com', 'x-github-api-version': '2022-11-28' };
  if (env.GITHUB_TOKEN) headers.authorization = `Bearer ${env.GITHUB_TOKEN}`;
  const res = await fetch(`https://api.github.com/repos/${REPO}${path}`, { headers, cf: { cacheTtl: TTL, cacheEverything: true } });
  if (!res.ok) throw new Error(`GitHub ${res.status}`);
  return res.json();
}

/** Answers from the edge cache when it can, otherwise builds the JSON and caches it */
export async function cached(context, build) {
  const cache = caches.default;
  const key = new Request(new URL(context.request.url).toString(), { method: 'GET' });
  const hit = await cache.match(key);
  if (hit) return hit;
  let body;
  let status = 200;
  try { body = await build(); } catch (e) { body = { error: String(e.message || e) }; status = 502; }
  const res = new Response(JSON.stringify(body), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': status === 200 ? `public, max-age=${TTL}` : 'no-store',
      'access-control-allow-origin': '*',
    },
  });
  if (status === 200) context.waitUntil(cache.put(key, res.clone()));
  return res;
}

export const releaseInfo = (r) => ({
  version: String(r.tag_name || '').replace(/^v/, ''),
  name: r.name || null,
  date: r.published_at,
  url: r.html_url,
  notes: r.body || '',
  assets: (r.assets || []).map((a) => ({ name: a.name, size: a.size, downloads: a.download_count, url: a.browser_download_url })),
});
