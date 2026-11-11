# kertasoffice.com

The Kertas Office website. Plain HTML and CSS with a little JavaScript, built by a script with no dependencies and hosted on Cloudflare Pages.

## Editing

| What | Where |
|---|---|
| Page content | `src/pages/*.html` (the comment lines at the top set the title, description and URL) |
| Header and footer (all pages) | `src/layout.html` |
| Styles | `assets/site.css` |
| Behaviour (download picker, tabs, search, live data) | `assets/site.js` |
| Screenshots and icons | `img/` |
| Headers, redirects, robots.txt | `static/` |

Build and preview:

```bash
node website/build.mjs        # writes website/dist
npx serve website/dist        # or any static server
```

The build fails if a page contains an em dash or is missing its title, description, path or nav line.

## Live data from GitHub

These parts update themselves, with no site rebuild:

- **Download sizes and latest version:** from the latest GitHub Release.
- **Changelog:** from GitHub Releases. Write the release notes on GitHub using `## PDF`, `## Writer` and `## Sheets` headings with bullet lists, and they get the coloured app badges on the site.
- **Roadmap:** open issues labelled `roadmap:now`, `roadmap:next` and `roadmap:later`.
- **Known issues (Compatibility page):** open issues labelled `known-issue`.

They are fetched through the Pages Functions in `functions/api/`, which cache GitHub's answers for 10 minutes. If GitHub cannot be reached (or the repo is private), the pages keep the text written in the HTML.

Download buttons link to `releases/latest/download/<file>`, which works because installer file names carry no version number.

When Windows signing through SignPath is live, set `SIGNED = true` at the top of `assets/site.js`. The install steps, FAQ, verify box, policy text and footer credit switch to the signed wording.
