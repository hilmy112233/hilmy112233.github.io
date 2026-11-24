# CV source

One HTML file, two PDFs, two Word files. Edit `cv.html`, run both builds, done.

```bash
node cv/build.mjs
python cv/build-docx.py
```

Writes:

- `assets/CV_Hilmy_Abdurrahman_Darmawan.pdf` — English (this is what the site links to)
- `assets/CV_Hilmy_Abdurrahman_Darmawan_ID.pdf` — Indonesian
- `assets/CV_Hilmy_Abdurrahman_Darmawan.docx` — English, editable
- `assets/CV_Hilmy_Abdurrahman_Darmawan_ID.docx` — Indonesian, editable

The `.docx` files are generated from `build-docx.py`, which holds the same content as a
Python list. **If you edit `cv.html`, edit `build-docx.py` too**, or the Word files drift
out of sync with the PDFs. There is no automatic sync between them.

## Editing

Every translatable string in `cv.html` is a `<span>` carrying both versions:

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

## ATS compatibility

The PDFs and the Word files are both checked, because an ATS reads them differently.

```bash
python _atscheck.py assets/CV_Hilmy_Abdurrahman_Darmawan.pdf assets/CV_Hilmy_Abdurrahman_Darmawan_ID.pdf
python _docxcheck.py assets/CV_Hilmy_Abdurrahman_Darmawan.docx assets/CV_Hilmy_Abdurrahman_Darmawan_ID.docx
```

`_atscheck.py` runs **pypdf** and **pdfminer** — two independent extractors — and checks
that the text comes back in reading order, that the four section headings are present, that
the contact fields scrape cleanly, and that there are no replacement glyphs.

`_docxcheck.py` checks the Word files for the structures that break Word parsing: text
boxes, layout tables, and content parked in headers or footers. It also confirms the skills
block keeps each `Label: values` pair on a single paragraph.

Both must print `PASS`. The thing that actually fails ATS is not the PDF format — it is a
two-column skills grid, where the parser reads every label first and every value second, so
`C/C++` stops being attached to `Programming`. That is why `.skills` is a single column of
`<p><b>Label:</b> value</p>` and not a `<dl>`.

## Page count

Both PDFs and both Word files must be **one A4 page**. Word's own count is the authority
for the `.docx`:

```bash
python cv/_build_and_count.py
```

That rebuilds the Word files and reports the page count through Word itself, failing if
either comes out at two pages.

## Adding an entry

Copy a `.job` block. Keep dates in `<span class="d">` — the layout right-aligns them and
uses tabular numerals so the column stays straight. Then mirror it in `build-docx.py`.

## Room left

English sits at 1023 px of 1047; Indonesian at 1023. Roughly one bullet line of headroom
each. Add a line here and you will need to cut one somewhere else.

Past that, cut something first, or accept two pages deliberately and drop the page check
in `build.mjs` and `_build_and_count.py`.
