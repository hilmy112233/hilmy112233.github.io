# CV source

One HTML file, two PDFs. Edit `cv.html`, run the build, done.

```bash
node cv/build.mjs
```

Writes:

- `assets/CV_Hilmy_Abdurrahman_Darmawan.pdf` — English (this is what the site links to)
- `assets/CV_Hilmy_Abdurrahman_Darmawan_ID.pdf` — Indonesian

## Editing

Every translatable string is a `<span>` carrying both versions:

```html
<span data-id="Teks Indonesia" data-en="English text">English text</span>
```

Change both attributes. The visible text is only a fallback for the no-JS case.

## The build checks

`build.mjs` fails loudly instead of shipping a broken file:

- **one page** — counts PDF page objects, so an overflowing CV cannot pass silently
- **no fallback font** — any face other than Fraunces / Work Sans in the PDF means a glyph
  was silently substituted. Work Sans has no arrows (`→ ←`), no Greek (`α`), and no
  `•`-like marks, so use `>` for arrows and spell out `alpha`
- **layout height** — measures the rendered DOM against the 1047 px A4 content box, per
  language, since Indonesian runs longer than English

Fonts are inlined as data URIs during the build for exactly that reason: Chromium's print
path drops `file://` font requests even with `--allow-file-access-from-files`.

## Adding an entry

Copy a `.job` block. Keep dates in `<span class="d">` — the layout right-aligns them and
uses tabular numerals so the column stays straight.

## Room left

English sits at 1014 px of 1047; Indonesian at 1028. That is one bullet line of headroom,
and the Indonesian is the tighter of the two. Add a line here and you will need to cut one
somewhere else.
Past that, cut something first, or accept two pages deliberately and drop the page check
in `build.mjs`.
