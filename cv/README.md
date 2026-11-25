# CV source

Two HTML pages, four PDFs, four Word files. Edit a page, run both builds, done.

| Page | Purpose | Output stem |
|---|---|---|
| `cv.html` | generalist CV — panel technician first, software second | `assets/CV_Hilmy_Abdurrahman_Darmawan` |
| `cv-itsupport.html` | targeted at IT Support Engineer roles | `assets/CV_Hilmy_Abdurrahman_Darmawan_ITSupport` |

Each stem produces `_ID.pdf` / `_ID.docx` for Indonesian and the bare name for English.

```bash
node cv/build.mjs          # both pages -> four PDFs
python cv/_build_and_count.py   # both DOCX builders -> four .docx + Word page count
```

The `.docx` files come from `build-docx.py` and `build-docx-itsupport.py`, which hold the
same content as Python lists. **If you edit an HTML page, edit its builder too**, or the Word
files drift from the PDFs. There is no automatic sync; `_drift.py` catches it afterwards.

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

## ATS compatibility

The PDFs and the Word files are both checked, because an ATS reads them differently.

```bash
python _atscheck.py assets/CV_Hilmy_Abdurrahman_Darmawan.pdf assets/CV_Hilmy_Abdurrahman_Darmawan_ITSupport.pdf
python _docxcheck.py assets/CV_Hilmy_Abdurrahman_Darmawan.docx assets/CV_Hilmy_Abdurrahman_Darmawan_ITSupport.docx
python _drift.py        # PDF vs DOCX text, all four pairs
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

One more trap: a hyphenated compound at a line break extracts as two tokens (`per-` +
`second`), which breaks the PDF↔DOCX drift check. Spell them open instead —
"telemetry every second", "a team of six".

## Page count

Every PDF and every Word file must be **one A4 page**. Word's own count is the authority
for the `.docx`:

```bash
python cv/_build_and_count.py
```

That rebuilds all four Word files and reports the page count through Word itself, failing if
any comes out at two pages. If Word is already open with one of these files, COM will refuse
to open it — close Word and re-run.

## Adding an entry

Copy a `.job` block. Keep dates in `<span class="d">` — the layout right-aligns them and
uses tabular numerals so the column stays straight. Then mirror it in the matching builder.

## Room left

| File | Height |
|---|---|
| generalist EN / ID | 1043 / 1047 px |
| IT Support EN | 1013 / 1047 px |
| IT Support ID | 1044 / 1047 px |

Roughly one bullet line of headroom each. Add a line here and you will need to cut one
somewhere else. Past that, cut something first, or accept two pages deliberately and drop
the page check in `build.mjs` and `_build_and_count.py`.

## Backups

`assets/backup-2026-11-25/` holds the four CV files as they stood before the IT Support
variant was added.
