// GET /api/release: the latest published release (version, date, files and sizes)
import { gh, cached, releaseInfo } from '../../lib/github.js';

export const onRequestGet = (context) => cached(context, async () => releaseInfo(await gh(context.env, '/releases/latest')));
