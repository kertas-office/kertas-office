// GET /api/releases: the last 15 published releases, newest first, for the changelog
import { gh, cached, releaseInfo } from '../../lib/github.js';

export const onRequestGet = (context) => cached(context, async () => {
  const list = await gh(context.env, '/releases?per_page=20');
  return list.filter((r) => !r.draft && !r.prerelease).slice(0, 15).map((r) => {
    const { assets, ...rest } = releaseInfo(r);
    return rest;
  });
});
