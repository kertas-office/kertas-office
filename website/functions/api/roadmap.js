// GET /api/roadmap: open issues labelled roadmap:now, roadmap:next, roadmap:later and known-issue
import { gh, cached } from '../../lib/github.js';

const LABELS = { now: 'roadmap:now', next: 'roadmap:next', later: 'roadmap:later', known: 'known-issue' };

export const onRequestGet = (context) => cached(context, async () => {
  const out = {};
  await Promise.all(Object.entries(LABELS).map(async ([key, label]) => {
    const issues = await gh(context.env, `/issues?state=open&per_page=30&sort=updated&labels=${encodeURIComponent(label)}`);
    out[key] = issues
      .filter((i) => !i.pull_request)
      .map((i) => ({ number: i.number, title: i.title, url: i.html_url }));
  }));
  return out;
});
