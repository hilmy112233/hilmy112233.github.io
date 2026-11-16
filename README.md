# Hilmy Abdurrahman Darmawan — portfolio

Static site: the files here are served exactly as they are, no build step.

## Publish

The repo is `hilmy112233/hilmy112233.github.io`. Push to `main`; GitHub Pages serves the
root of that branch directly, so the files here are the live site — no build, no Actions
workflow. Settings → Pages → Source is **Deploy from a branch: main / (root)**.

Two consequences worth knowing:

- Anything committed to the root ships. `_preview/`, `_shot.mjs` and the `_*` scratch files
  are gitignored so they stay local; if you add a new working folder, gitignore it too.
- `cv/` is committed and therefore public at `/cv/cv.html`. That is deliberate — the CV
  source is part of the repo — but it is not linked from any page.

## Structure

- `index.html` — about, skills, journey, featured projects, contact
- `projects.html` — all 5 builds: IoT & Hardware (2) and Software, Mobile & Systems (3)
- `projects/<slug>/index.html` — one write-up per project, 4 of the 5 have one
- `assets/CV_Hilmy_Abdurrahman_Darmawan.pdf` — English CV, linked from the topbar
- `assets/CV_Hilmy_Abdurrahman_Darmawan_ID.pdf` — Indonesian CV
- `assets/og-card.png` — 1200×630 link preview, referenced by every page
- `assets/favicon.svg` — tab icon
- `cv/` — CV source and build script (not published; see `cv/README.md`)

## Adding a project

1. Create `projects/<slug>/` with an `index.html` and its images.
2. Add an `<article>` card in `projects.html` under the right section, with a link to
   `projects/<slug>/`.
3. Update the counts in the hero buttons and the topbar nav (`IoT & Hardware · 2`,
   `Software & Systems · 3`) and the "Here are five engineering builds" sentence in the
   hero — they are all written by hand.
4. Add the new URL to `sitemap.xml`.

## Diagrams

The two SVG diagrams in `projects/motorcycle-fuel-telemetry/` and the two in
`projects/arjuna-mobile/` are hand-generated vector art using the site tokens. They
scale to any width and stay sharp on retina. Every `<img>` carries `width`/`height` so
the page does not reflow while loading — if you swap an image, update those numbers
(`figure img` sets `height:auto` so the ratio is preserved).

## Photos

Card images fill their frame. Photos use `object-fit: cover`; diagrams and screenshots
use `contain` on white so nothing important is cut off. To swap one, replace the file
and keep the filename.

## Not part of the site

- `_preview/` — render screenshots kept for review only (gitignored)
- `_shot.mjs` — throwaway headless-Edge previewer that fills `_preview/` (gitignored)
