---
name: docx
description: "Create, read, edit, validate, comment on, and rewrite Word documents (.docx/.dotx) with a set of small Python CLIs. Use whenever the task involves a Word file: generating a report/memo/letter from a spec, extracting or inspecting existing content, find-and-replace across a document, editing tables, managing comments and tracked changes, filling {{token}} templates, extracting/applying formatting templates, fixing table column widths, or converting linear formulas into native Office Math (OMML). Also use when a deliverable should be a .docx. Do NOT use for PDF, spreadsheets (XLSX/CSV), slide decks, or plain one-shot text extraction (use markitdown)."
license: MIT
---

# Word documents (.docx)

Scripts live flat in `scripts/` — **27 `.py` + 1 `.ps1` = 28 files**, no
sub-package, no `__init__.py`. **26 are CLIs** (each carries an
`if __name__ == "__main__"` guard); **`_bootstrap.py` is the one pure library**.
Of the 26, **19 expose an argparse surface** with `--help` and the other 7 print a
one-line self-description or run a self-check. Run them with any Python 3.10+ that
has `python-docx` and `lxml` installed.

Output convention differs by script and is deliberate:

- **JSON on stdout** — `docx_create`, `docx_read`, `docx_edit`, `docx_validate`,
  `docx_comments`, `docx_revisions`, `docx_section`, `docx_numbering`,
  `docx_crossref`, `docx_toc`. Parse these.
- **Human-readable text** — `docx_mathify`, `table_width`, `insert_formula`,
  `docx_template` (except `fill`, which emits JSON on stdout), `fix_fonts`,
  `fix_zoom`, `merge_docs`, `pdf_export`, `docx_pack` (a bare run prints how to
  import it), and `docx_common --api` (a name-plus-signature listing). These
  report progress and are meant to be read.
- **Either, on request** — `docx_preview` prints human-readable lines by default
  and one JSON object with `--json`. It never mixes the two on one run.

Exit codes are reliable: `0` on success, `1` on an expected failure, `2` on a
usage error. See [Exit codes](#exit-codes) for the full table and for `3`, the
one partial-success code.

## Setup

```bash
python -m pip install -r requirements.txt
```

`latex2mathml` is only needed for the formula features. A local Microsoft Office
install supplies `MML2OMML.XSL`, which the LaTeX→OMML converter reuses — point
`MML2OMML_XSL` at it if it lives somewhere unusual.

**Two capabilities are Windows-only, and they fail at call time rather than at
install time.** The **formula chain** needs `MML2OMML.XSL`, and
`math_omml.find_xslt()` probes only Windows Office locations
(`%ProgramFiles%`, `%ProgramFiles(x86)%`, `%LOCALAPPDATA%`, `C:\Program Files…`),
so on Linux or macOS it raises `FileNotFoundError` unless you supply a copy and
set `MML2OMML_XSL`. **PDF export** drives Word through COM. Everything else —
create / read / edit / validate / comments / revisions / numbering /
cross-references / static TOC / sections / table geometry / fonts / merge /
templates / `docx_pack` — is pure `python-docx` + `lxml` and needs no Office.

## Working on an unpacked document

Two commands (`table_width.py`, `insert_formula.py`) operate on an **unpacked**
document — a `.docx` extracted to a directory so its XML parts can be edited
directly. `docx_template.py apply --to`, `fix_fonts.py`, and `fix_zoom.py` instead
take a packaged `.docx` and rewrite it in place through a temporary file, so pass
the file itself rather than an extracted directory.
Produce one with:

```bash
python -c "import sys; sys.path.insert(0,'scripts'); from docx_pack import unpack; unpack('in.docx','unpacked/')"
# ...edit files under unpacked/ ...
python -c "import sys; sys.path.insert(0,'scripts'); from docx_pack import pack; pack('unpacked/','out.docx',original='in.docx')"
```

Or import it where you already have Python:

```python
from docx_pack import unpack, pack
unpack("in.docx", "unpacked/")
pack("unpacked/", "out.docx", original="in.docx")
```

`pack` reuses the original's entry order and compression method, so a
round-trip that changes nothing produces an equivalent archive. Pass
`original=` and a missing, empty, or non-zip original raises rather than quietly
packing a document that lost its carry-over.

`unpack` treats the archive as untrusted input and refuses every entry that
would land outside the target directory, raising `UnsafeArchiveError` before
anything is written. The refused shapes, each one exercised by the test suite:

- **`..` traversal** — `../escape.txt`, and the backslash spelling
  `..\..\evil.txt` (the `zipfile` layer normalises it to `../../evil.txt` on
  Windows, and the `normpath` check catches it there).
- **POSIX absolute paths** — `/abs/evil.txt`.
- **Windows drive letters** — `D:/x.txt` and `D:x.txt`. Neither is POSIX
  absolute, yet `pathlib` on Windows treats both as rooted, so `dst / name`
  would drop `dst` entirely.
- **UNC shares** — `\\server\share\x.txt`, reported as a drive prefix.
- **Reserved device names** — `CON`, `NUL`, `AUX`, `PRN`, `CONIN$`, `CONOUT$`,
  `COM0`–`COM9`, `LPT0`–`LPT9`, matched case-insensitively on the last component
  with any extension ignored, so `nul`, `CON.txt` and `word/AUX.xml` are all
  rejected. Win32 resolves these to a device in any directory, and `NUL`
  swallows the write while reporting success.
- **A trailing dot or space** in the last component — `evil.txt ` and
  `evil3. .`. Win32 strips it, so two archive entries would collide on one file
  name and the extracted tree would stop matching the archive.

## Quick reference

| Task | Command |
|------|---------|
| Create from a JSON spec | `python scripts/docx_create.py spec.json out.docx` |
| Read text / outline / styles | `python scripts/docx_read.py in.docx --text` |
| Extract images | `python scripts/docx_read.py in.docx --images DIR` |
| Inspect tracked changes | `python scripts/docx_read.py in.docx --revisions` |
| Replace text | `python scripts/docx_edit.py replace in.docx --find A --replace B -o out.docx` |
| Set a table cell | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X -o out.docx` |
| Insert a paragraph | `python scripts/docx_edit.py insert in.docx --index 3 --text "..." -o out.docx` |
| Delete a paragraph | `python scripts/docx_edit.py delete in.docx --index 3 -o out.docx` |
| Apply a paragraph style | `python scripts/docx_edit.py style in.docx --index 0 --style "Heading 1" -o out.docx` |
| Merge adjacent identical runs | `python scripts/docx_edit.py normalize in.docx -o out.docx` |
| Insert a TOC *field* | `python scripts/docx_edit.py toc in.docx --index 1 -o out.docx` |
| Add "Page X of Y" to the footer | `python scripts/docx_edit.py page-numbers in.docx -o out.docx` |
| Health-check a document | `python scripts/docx_validate.py in.docx` |
| List comments | `python scripts/docx_comments.py list in.docx` |
| Add a comment | `python scripts/docx_comments.py add in.docx --target "some text" --text "note"` |
| Delete a comment | `python scripts/docx_comments.py delete in.docx --id 0` (ids start at `0`; `list` prints them) |
| List / accept / reject revisions | `python scripts/docx_revisions.py list\|accept-all\|reject-all in.docx` |
| Fill `{{token}}` placeholders | `python scripts/docx_template.py tpl.docx values.json out.docx` |
| Extract a formatting template | `python scripts/docx_template.py extract tpl.docx -o tpl.json` |
| Apply a formatting template | `python scripts/docx_template.py apply tpl.json --to in.docx -o out.docx` |
| Diagnose table widths | `python scripts/table_width.py report unpacked/` |
| Force table widths | `python scripts/table_width.py set unpacked/ --widths 2.4,4.2,4.2` |
| Copy `tcW` into the grid | `python scripts/table_width.py sync unpacked/` |
| Spread columns over the text width | `python scripts/table_width.py auto unpacked/ [--content-width 14.6]` |
| Batch linear formulas → Office Math | `python scripts/docx_mathify.py in.docx -o out.docx --report conv.md` |
| Insert one LaTeX formula | `python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display` |
| Inspect numbering definitions | `python scripts/docx_numbering.py list in.docx` |
| Add a numbering definition | `python scripts/docx_numbering.py define in.docx --text "[%1]" -o out.docx` |
| Number paragraphs | `python scripts/docx_numbering.py apply in.docx --index 155:171 --num-id 1` |
| Restart a list mid-document | `python scripts/docx_numbering.py restart in.docx --index 7 --num-id 1` |
| Strip numbering | `python scripts/docx_numbering.py detach in.docx --index 3:9` |
| Make a numbered caption | `python scripts/docx_crossref.py caption in.docx --index 42 --label 图 --text "架构"` |
| Insert a cross-reference | `python scripts/docx_crossref.py ref in.docx --index 70 --target 图_1 --prefix "如图" --suffix 所示` |
| Put a bookmark on a paragraph | `python scripts/docx_crossref.py bookmark in.docx --index 12 --name 图_1 -o out.docx` |
| Refresh all field results | `python scripts/docx_crossref.py renumber in.docx -o out.docx` |
| List bookmarks / fields | `python scripts/docx_crossref.py list in.docx` |
| Write a static table of contents | `python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7` (one `--pages` value per entry — the three values here mean the source has exactly three entries; a different count is exit 2) |
| Inspect / edit sections | `python scripts/docx_section.py list\|break\|pagenum\|pagenums\|header\|footer\|cols in.docx` |
| Measure the displayed page of paragraphs | `python tools/word_accept.py in.docx --page-probe 1,9,43 --json out.json` |
| Unify fonts (CJK body / Latin Times New Roman) | `python scripts/fix_fonts.py in.docx [more.docx ...]` |
| Merge two documents | `python scripts/merge_docs.py base.docx src.docx out.docx` |
| Add the missing `w:percent` to `<w:zoom>` | `python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]` |
| Export to PDF with the text layer intact | `pwsh scripts/pdf_export.ps1 -DocxPath in.docx [-PdfPath out.pdf]` |
| Render + lint the layout (PDF, page PNGs, page-fill report) | `python scripts/docx_preview.py in.docx --pdf out.pdf --pages png/ --lint --json` |

Common commands (each one checked against `--help` on this machine, exit `0`):

```bash
python scripts/docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex
python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display --after "marker"
python scripts/table_width.py set unpacked/ --table 1 --widths 2.4,4.2,2.4,5.6
python scripts/merge_docs.py base.docx src.docx merged.docx
python scripts/fix_fonts.py a.docx b.docx --heading-threshold 14
python scripts/docx_template.py fill tpl.docx values.json out.docx --strict
```

`docx_edit page-numbers` is **idempotent**: a footer that already carries a
`PAGE` or `NUMPAGES` field — complex form or `<w:fldSimple>` — is left
byte-for-byte alone, an in-place re-run writes nothing at all, and stdout grows
two extra keys, `changed` and `already_present`, so a script can tell "added" from
"was already there".

**`--index` is always 0-based**, and what it counts depends on the script:

| Script | What `--index` counts | Out of range |
|---|---|---|
| `docx_edit insert` | body paragraphs; the index is the paragraph the new one goes **before**, so the legal range is `[0, N-1]` | exit `2`, message names the legal range |
| `docx_edit delete` / `style` | body paragraphs | exit `2`, message names the legal range |
| `docx_edit toc` | body paragraphs | exit `2`, message names the legal range |
| `docx_crossref` | body paragraphs | exit `2`, message names the legal range |
| `docx_numbering` | body **and table** paragraphs, in document order; `lo:hi` is half-open and `--index` is repeatable | exit `1`, "selection matched no paragraphs" |
| `docx_toc --insert-at` | body paragraphs | exit `2`, message names the legal range |

Every one of those refusals is decided **before the document is touched**: no
file is written, the input is byte-for-byte as it was, and stdout stays empty —
the message goes to stderr as a single `error: …` line naming the option, the
value you passed, and the range that would have worked
(`error: delete --index 99 is out of range: legal range is [0, 3], current N=4`).
A negative index is refused on the same footing rather than read through Python's
negative indexing: `-1` is not "the last paragraph" in any durable sense — add or
delete one paragraph and the same argument means a different one.

`docx_toc`'s `inserted_at` in its report is a different thing: it is where the
TOC was actually written, measured after insertion, not an input.

**`docx_section.py` owns section-level page setup.** `list` prints every
section's geometry as JSON; `break` appends a section; `pagenum` writes
`<w:pgNumType>` (format and start value); `pagenums` writes the PAGE/NUMPAGES
fields into a section's footer; `header` / `footer` write that section's own
running text; `cols` sets the column count. Every subcommand except `list` takes
`--section N` (default `-1`, the last section).

```bash
python scripts/docx_section.py break in.docx --start-type new_page -o out.docx
python scripts/docx_section.py pagenum out.docx --section 1 --fmt lowerRoman --start 1 -o out.docx
python scripts/docx_section.py pagenums out.docx --template page_of -o out.docx
```

The `lowerRoman start=1` in the middle line is exactly the case `docx_toc` warns
about: that section prints `i` on whatever physical page it starts at, so the
displayed page number and the physical page index are two different numbers.

**`pagenums --template` offers five shapes**, and `zh` is the one for Chinese
delivery: `第 <PAGE> 页　共 <NUMPAGES> 页`. The separator between `页` and `共` is
a full-width space (U+3000), not an ordinary space, because a half-width gap
reads as a typo at print size. `--total sectionpages` swaps the second field for
`SECTIONPAGES`, so "共 Y 页" counts the section rather than the document. `zh`
ignores `--prefix` / `--suffix` — it builds its own text — and says so in the
report (`"ignored": ["prefix"]`) instead of dropping them in silence.

`--template custom` takes `--prefix` / `--suffix` and expands **`{page}`** and
**`{total}`** inside them into the PAGE and total fields, in the order they
appear, so `--prefix "{page}/" --suffix " of {total}"` gives `1/ of 12`. With no
placeholder in either string, `custom` behaves exactly as it always did. The
same keys work in a spec: a top-level `footer_page_numbers`, or one inside a
`section_break`, accepts `"template": "zh"` and `custom` with placeholders.

**`docx_template.py apply` carries two independent formula channels, on two
separate defaults, and they are mutually exclusive on the command line:**

| Flag | Channel | Default |
|---|---|---|
| `--no-fix-formulas` | marker channel — paired `_xx_` / `^xx^` → `vertAlign` runs | **on** |
| `--fix-formulas` | OMML channel — LaTeX / linear formulas → native Office Math (runs `docx_mathify`) | off |

They read like a `+`/`-` pair but write different destinations
(`fix_formulas` / `omml_formulas`), so passing both used to switch the marker
pass off *and* hand raw `_xx_` markers to the OMML converter — which then failed
while the CLI still exited `0`. It is now an argparse conflict, exit `2`. Without
`-o`, `apply --to` rewrites the target in place and writes a `<file>.bak` first;
with `-o`, the `--to` file is left byte-for-byte alone.

**`--fix-formulas` passes `docx_mathify`'s exit code back up, including `3`.** The
OMML channel runs mathify in place, and `3` there means "the document was written,
but at least one formula is still linear text" — the one exit that combines an
artifact with a non-zero code. Reporting `0` for that would hand a half-converted
file to the caller as finished, so `apply` now returns exactly what mathify
returned: `1` and `2` unchanged, and `3` as `3`. Only `0` collapses back to `0`.
`--no-fix-formulas` never calls mathify, so it behaves as before.

## Creating a document

`docx_create.py` takes a JSON spec. Top-level keys: `page`, `styles`, `header`,
`footer`, `footer_page_numbers`, and `blocks`. Each block has a `type`:

`footer_page_numbers` (top level or inside a `section_break`) accepts `fmt`,
`start`, `align`, `template`, `total`, `prefix`, `suffix`, and `page_type` — and
`font`, `size_pt`, `bold`, which style the runs that carry the page number:
`font` and `size_pt` reach both the literal runs and the field runs, while
`bold` reaches the field runs holding the digits only when it is truthy — the
table below is the precise statement. The last three used to raise `ValueError` the
moment you set them; they now apply, and on a mixed-script footer `font` binds all
four `w:rFonts` slots (`ascii` / `hAnsi` / `eastAsia` / `cs`) so a Chinese
`第 X 页　共 Y 页` prefix does not fall back to the theme's East Asian face while the
digits use the font you asked for. (The separator here is the same full-width
space, U+3000, that the `zh` template uses — see the templates section above.)

**`bold` styles two different kinds of run by two different rules, so `"bold":
false` and leaving `bold` out are not the same request.** A `page_of` footer is a
mix of literal runs (`"Page "`, `" of "`, a Chinese prefix/suffix) and field runs
(the `PAGE` / `NUMPAGES` fields — the digits themselves plus their begin /
separate / end plumbing). The literal runs go through `run.font.bold`, which is
three-state; the field runs go through `_rpr()`, which tests `if bold:` and is
therefore two-state. Measure it on one footer with `font` and `size_pt` held
constant and the difference is visible in the XML:

| `bold` | literal runs | field runs (the digits) | `word/footer1.xml` | `<w:b` |
|---|---|---|---|---|
| `true` | `run.font.bold = True` → `<w:b/>` | `<w:b w:val="1"/>` **and** `<w:bCs w:val="1"/>` on all ten | 4232 bytes | 22 — 12 `w:b` (one per run) plus 10 `w:bCs` |
| `false` | `run.font.bold = False` → `<w:b w:val="0"/>`, an explicit *off* | **nothing written**; `False` and "unset" take the same branch | 3912 bytes | 2 — the literal runs only |
| omitted | nothing written, the run keeps inheriting | nothing written | 3880 bytes | 0 |

Those three numbers are the whole point: `false` and omitted differ by 32 bytes
(16 per literal run) and are **not equivalent**, and neither one un-bolds the page
number **digit**. The digits sit in field runs, and a field run only ever receives
`w:b` when `bold` is truthy — so on a footer whose `Footer` style is bold,
`"bold": false` suppresses bold on `"Page "` and `" of "` while the number stays
bold, which looks like a bug and is the documented behaviour. Either turn bold off
at the style, or leave `bold` truthy and let both kinds be bold. (The
`section_break` keys `header_bold` / `footer_bold` are a different story: those
write plain text runs with no fields in them, so `false` there is a real explicit
off — measured, one `<w:b w:val="0"/>` on the only run — and does mean "not
bold".)

| `type` | Fields |
|---|---|
| `heading` | `text`, `level` (0–9) |
| `paragraph` | `text`, or `runs`, plus optional `style` |
| `bullet_list` | `items` |
| `numbered_list` | `items`; optional `num_id` (reuse an existing counter), `num_text` (default `%1.`), `num_fmt`, `start`, `ilvl`, `indent_left`, `hanging` |
| `caption` | `label` (default `图`), `text`, optional `bookmark` — inserts `图 N  题注` with a `SEQ` counter + bookmark |
| `table_caption` | the same fields as `caption`, for a table |
| `crossref` | `target` (a bookmark), optional `prefix`, `suffix`, `format` (`default`/`full`/`above-below`/`page`/`number`), `cached`, `own_paragraph` |
| `table` | `header`, `rows`, optional `style` (default `Table Grid`), `header_bold`, `keep_together` |
| `image` | `path`, and `width_mm` **or** `width_frac` (mutually exclusive) |
| `page_break` | — |
| `section_break` | `start_type`, optional `page`, `header` / `footer` (+ their `_align` / `_font` / `_size_pt` / `_bold`), `footer_page_numbers` (`fmt` / `start` / `align` / `template` / `total` / `prefix` / `suffix` / `page_type`, plus `font` / `size_pt` / `bold` for the page-number runs — see the `bold` table above, `false` reaches the literal runs but not the field runs), `different_first_page`, `cols`, `no_footer`, `reuse_last_paragraph`, `apply_to` |
| `toc` | optional `instr` — default `TOC \o "1-3" \h \z \u`; pass `TOC \h \c "图"` for a table of figures. **A field**: Word must update it (F9) before any entry exists |

**The spec is validated before a single paragraph is built, and an unknown key
or an unknown block type is an error, not a no-op.** `{"blockz": [...]}` used to
build a valid, empty document, print `"blocks": 0` and exit `0`, so the author
had no way to see that the content never went in. It now exits `2` with one line
naming the offending key and, when it is close enough, a "did you mean" hint:

```
spec: unknown key "blockz" (did you mean "blocks"?); valid keys here: blocks, footer, footer_page_numbers, header, page, styles
blocks[0]: unknown block type 'paragraf' (did you mean "paragraph"?); choose one of: bullet_list, caption, crossref, heading, image, numbered_list, page_break, paragraph, section_break, table, table_caption, toc
```

Exit `2` promises that **nothing was written**, so you can fix the spec and
retry without wondering whether half a document exists. Values are not clamped
either: an out-of-range value is reported with the range it would have been
clamped into.

`styles[]` entries take `font`, `font_east_asia`, `size_pt`, `bold`, `italic`,
`color`, `base`, and `outline_level` — the last one writes
`<w:pPr><w:outlineLvl>` so Word treats the style as a heading. Without it a
template's own `卷标题` / `ChapterTitle` is *body text* to Word no matter how it
looks: the navigation pane is empty and `TOC \o` collects nothing.

**A `styles[]` entry that names an existing style modifies it in place — it does
not fail, and it touches only the keys you gave.** `{"name": "Heading 1",
"size_pt": 16}` changes the size of Word's built-in `Heading 1` and leaves its
font, its weight, its colour and its `outlineLvl` exactly as they were; a second
apply of the same entry still leaves one `<w:style>` behind. Keys you omit are
never reset, so nothing silently reverts. To change a colour back, say so
explicitly with `color` — and note that `Heading 1`–`Heading 9`, `Title` and
`Subtitle` are **black by default** in a document this toolkit builds (see
[Headings are black by default](#headings-are-black-by-default)), so a style you
never coloured is already `000000` rather than the theme's blue.

```json
{"styles": [
  {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
   "bold": true, "outline_level": 1}
]}
```

`outline_level` is 1-based (`1` → `w:outlineLvl val="0"`). `0` means "explicitly
body text" (`val="9"`), for a style that must never be collected.

`heading.level` is `0`–`9`, not `1`–`9`: python-docx's `add_heading` treats `0`
as Word's **Title** style and `1`–`9` as Heading 1–9. `10` is a spec error, not
a clamp.

### Headings are black by default

Every document this toolkit builds gets a pass over `Heading 1`–`Heading 9`,
`Title` and `Subtitle` **before** any `styles[]` entry is applied: the theme
colour binding (`w:themeColor` / `w:themeShade` / `w:themeTint`) is removed and
`w:val="000000"` written in its place, all four character variants
(`Heading1Char` … `TitleChar`, `SubtitleChar`) get the same treatment, `Title`'s
paragraph bottom border switches to black (`w:color="000000"` on `w:pBdr`, with
`w:sz` and `w:space` unchanged), and `Subtitle` loses its italic (`w:i` and
`w:iCs`). The point is that a heading looks the same on a machine whose theme
resolves to a different accent, and that a generated document does not arrive in
theme blue when nobody asked for blue.

This is a **default, not a lock**. A `styles[]` entry that gives `color` wins,
because the neutralising pass runs first — the same door the in-place modify
above opened. Heading 9 keeps its italic; only `Subtitle` was in scope for that
change.

### Making a table stay on one page

`{"type": "table", …, "keep_together": true}` asks Word to hold the whole table
together. Word has no table-level "keep together" attribute, so the request is
written in two layers: `w:trPr/w:cantSplit` on **every** row (a row may not be
split across a page break) and `w:keepNext` on every paragraph of every row
except the last (each row pulls the next one with it). The default is `false`,
and when the key is absent not one of those elements is written.

**`cantSplit` only forbids splitting a row — a table taller than one page will
still break across pages.** `keep_together` is a request, not a guarantee; it
helps the 8-row table that would otherwise lose its last two rows, and does
nothing for a 60-row table that cannot fit on a page in the first place.

`table_caption` followed immediately by a `table` is now **linked by default**,
with no key to set: the caption paragraph gets `keep_with_next` and the table's
first row gets `cantSplit`, so a one-line caption stops being orphaned at the
bottom of the previous page. Anything between them — a paragraph, a page break,
another table — breaks the linkage, because "right after" is measured on the
block list, not guessed. This linkage changes the layout of spec files written
before it existed, which is the point; if you need the old behaviour, put a
paragraph between the caption and the table.

### Sizing an image relative to the text block

`image` blocks take `path` plus **either** `width_mm` (an absolute width) **or**
`width_frac` (a share of the current section's text width, `0 < f ≤ 1`), never
both — passing both is exit `2` with nothing written. `width_frac` is resolved
against `content_width_mm(section)`, so it follows the page size and the margins
instead of hard-coding a number: A4 with 26 mm left and right margins gives a
158 mm text block, so `1.0` is 158.0 mm and `0.5` is 79.0 mm. Out-of-range
values are refused with the range rather than clamped, and `0.0` is refused
rather than read as "no width" — a number the caller wrote is never quietly
turned into something else.

### Forcing a heading onto a new page

Any block's `page_break_before: true` sets `<w:pageBreakBefore/>` in its
`<w:pPr>`, which is what Word itself writes for "page break before" in
*Paragraph → Line and Page Breaks*. Use that key on the heading.

**Do not reach for a `{"type": "page_break"}` block instead.** A page break is an
empty paragraph carrying a break run; it sits wherever the block order puts it,
so one stray break before a heading, or two in a row, produces a blank page or
pushes content a page too far — and a blank page in the middle of a document is
exactly the artefact that `docx_preview.py --lint` reports as `blank_pages`.
`page_break_before` attaches the break to the paragraph it belongs to, so moving
that paragraph moves the break with it.

## A static table of contents

`toc` above is a **field**: the entries do not exist until someone presses F9,
and a hand-written cache inside the field is discarded the moment they do. When
the file must be right on open — a deliverable, a PDF conversion, a diff — write
the entries as real paragraphs with `scripts/docx_toc.py` instead.

```
# one --pages value per TOC entry: this line assumes the document has exactly
# three entries. A mismatched count is a usage error (exit 2), not a silent pad.
python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7
```

```python
import docx_toc as toc
report = toc.build_static_toc(doc, levels=(1, 2, 3), pages=["1", "3", "7"],
                              title="目　录")
# report: entries / levels / styles / tab_pos_mm / tab_pos_source
#         dropped_entries / dropped_titles / adopted / section_index / inserted_at
```

Each entry is `title<TAB>page` in style `toc 1/2/3`, with a right tab stop at the
text-block width and a dot leader. The width comes from
`docx_create.content_width_mm(section)` — never hard-code it, because the six
test volumes genuinely differ (170 mm vs 160 mm).

The flags worth knowing:

| Flag | Effect |
|---|---|
| `--levels` | outline levels to include, `1,2` or `1-3` (default `1,2,3`) |
| `--pages` | one page number per entry; **a different count is a usage error (exit 2)** |
| `--title` | the heading text of the TOC itself |
| `--insert-at` | body paragraph index to insert before (default: where the previous TOC was, else the end); out of range is exit `2` |
| `--section-index` | which section's text width sets the right tab stop |
| `--indent-chars` | per-level left indent in characters (default `0,2,4`) |
| `--adopt-unmarked` | also replace an unmarked TOC this module did not write — off by default, because that may be the user's own |
| `--dry-run` | report what would be written, save nothing |

**Page numbers have to be measured, not guessed.** `pages=None` leaves the slot
empty; fill it from `tools/word_accept.py --page-probe`, which reports
`paragraph_pages[].displayed_text` — the number the reader sees, taken from
`Information(WD_ACTIVE_END_ADJUSTED_PAGE)`. `docx_toc.py` has no probing flag of
its own: you measure with `word_accept.py`, then pass the numbers back through
`--pages`. That is *not* the physical page
order: a `lowerRoman start=1` section prints `i` on what is physically page 8.
See `references/toc.md` for the two-pass recipe and the full trap list.

Rebuilding is safe and idempotent: every entry carries a `_DocxTocEntry…`
bookmark, and `drop_existing_toc` removes only what *this module* wrote — a
table of contents the user brought along with the built-in `toc 1` style is left
alone unless `adopt_unmarked=True` is passed.


`numbered_list` really numbers now. The previous version applied Word's
`List Number` *style*, which carries no `numPr`, so the paragraphs rendered with
no numbers at all; it now defines an `abstractNum` + `num` and attaches a real
`numId`. Pass the same `num_id` to two blocks to continue one sequence, or call
`docx_numbering.py restart` to break it.

```json
{
  "blocks": [
    {"type": "heading",   "text": "Report title", "level": 1},
    {"type": "paragraph", "text": "Plain paragraph."},
    {"type": "paragraph", "runs": [
        {"text": "bold ", "bold": true},
        {"text": "italic", "italic": true},
        {"text": ", underlined", "underline": true}
    ]},
    {"type": "bullet_list", "items": ["first", "second"]},
    {"type": "numbered_list", "items": ["甲", "乙"], "num_text": "[%1]"},
    {"type": "caption", "label": "图", "text": "系统架构示意图"},
    {"type": "paragraph", "text": "系统架构如"},
    {"type": "crossref", "target": "图_1", "suffix": "所示。"},
    {"type": "table",  "header": ["Item", "Value"],
     "rows": [["Voltage", "220 V"]]}
  ]
}
```

`crossref` appends its field to the *previous* paragraph, so a sentence can be
split across the two blocks above and read as `系统架构如1所示。`. Set
`"own_paragraph": true` to give it a fresh paragraph instead.

## Numbering and cross-references

Two scripts own these features:

- `docx_numbering.py` — `list` / `define` / `apply` / `detach` / `restart`
- `docx_crossref.py` — `caption` / `ref` / `bookmark` / `renumber` / `list`
- `docx_toc.py` — static TOC: `build_static_toc` / `collect_outline` / `drop_existing_toc`

Word numbering is a two-level indirection that no other script here exposed:
`<w:abstractNum>` defines the format (``lvlText`` like `[%1]`), `<w:num>` is a
counter instance pointing at it, and a paragraph's `<w:numPr>` names the
`numId`. Two paragraphs sharing a `numId` share one running counter.

A cross-reference is a `REF <bookmark>` field, and a caption is a
`SEQ <label> \* ARABIC` field wrapped in a bookmark. Every field this toolkit
writes carries a **cached result** (`begin / instrText / separate / result /
end`), so the number is visible before Word recalculates anything and stays
correct after F9. `renumber` recomputes `SEQ` counters and rewrites every
cached `REF` value, which is what makes a generated document correct for
readers that never update fields.

`docx_crossref bookmark` is the piece the `crossref` block type needs first: it
puts a named bookmark on one paragraph so a later `ref` — or a `crossref` block
in a spec — has something to point at. The name has to be unique; a second run
with the same name exits `1` with `bookmark already exists`.

## Formula conversion (LaTeX → native Office Math)

Two entry points share one engine (`math_omml.py`):

- **`docx_mathify.py` — batch.** Scans a whole document for formulas written as
  Unicode characters plus real sub/superscript runs (`U`+sub`oc`, `√2`,
  `1kΩ=10³Ω`, `∠53.1°`), converts each into native OMML, and leaves the Chinese
  prose, table structure, and font sizes untouched. `--dry-run` reports without
  writing; `--report md` records every before/after pair, plus a count and a
  per-paragraph table for spans that took the no-layout tier. `--inplace` writes
  a `<name>.bak.docx` first; `--inplace` together with `-o` is a usage error
  (exit `2`), because the write target would be ambiguous. So is `-o` together
  with `--dry-run` (exit `2`): dry-run writes nothing, so the path `-o` names can
  never be produced — accepting it silently meant echoing a destination that was
  never created.
- **`insert_formula.py` — single.** Takes LaTeX and drops one equation into an
  unpacked document at an anchor.

`insert_formula.py` has three rules, and each one exists to stop a silent
misfire rather than to be polite:

- **An anchored insert reports where it landed.** After writing, the script
  looks the new paragraph up in the tree and appends the measured coordinate to
  stdout — `at body[3]` for a body-level anchor, `at tc[1]` for one that went
  into a table cell. The index is the container's own child index (counted with
  `w:sectPr` and `w:tbl` included), which is what makes it checkable against
  `<w:body>` directly; it is not the argument you passed.
- **`--dump` together with `--after` / `--before` is a usage error, exit `2`.**
  `--dump` only prints the OMML and never writes the document, so the anchor
  would be swallowed without a word. `--after` and `--before` are mutually
  exclusive for the same reason.
- **An empty marker is a usage error, exit `2`.** `--after ""` would append the
  formula at the end of the document while looking anchored; omit the flag when
  appending at the end is what you actually want.

A marker that matches nothing exits `1` and names the range it searched — the
0-based body paragraph range plus the table-cell count.

`linear_to_latex.py` decides *what counts as a formula*. It deliberately stays
conservative so ordinary prose survives: `[0, 1]`, `f(x, y)`, and CJK headings
are left alone; `U=IR` and `ΣI=0` become equations. Recognising a span is not
the same as needing an equation *object* — see the tiers below.

The chain has three steps, and skipping one leaves half a formula:

```
linear text --(linear_to_latex)--> LaTeX --(latex2mathml)--> MathML --(MML2OMML.XSL via lxml)--> OMML
```

`latex_to_omml_lxml(latex, display=False, raw_spacing=False)` returns an element
carrying `xmlns:m`, ready to append into a `w:p`. `assert_legal()` rejects the
two structures that make Word say "the file may be corrupted": a nested
`<m:oMath>` inside `m:num` / `m:den` / `m:e` / `m:sub` / `m:sup` / `m:deg` /
`m:fName` / `m:lim`, and an `m:rad` missing its `m:e`.

### Three output tiers

`docx_mathify` places every recognised span in one of three tiers. The order is
fixed: master switch, then the no-layout tier, then the script tier, then OMML.

| Tier | Test | Output | Example |
|---|---|---|---|
| Structural formula | contains a command from `_STRUCTURAL_CMDS` (43 entries: `\frac`, `\sqrt`, `\int`, `\sum`, `\begin{`, `\left`, `\overline`, …) | OMML equation object | `x^2/16+y^2/9=1`, `v=s/t` |
| Plain script | no structural command, and at least one `^` or `_` | ordinary run + `w:vertAlign` — a real Word superscript, restored to full size by `Ctrl+Shift+=` | `a=9.8m/s2`, `10^-7` |
| No layout needed | no structural command, no `^`/`_`, exactly one `=`, a single ASCII letter on the left, a number plus an optional unit on the right | ordinary text run | `s=10km`, `t=2h`, `m=5kg`, `v=3m/s` |

The third tier is `no_layout_reason()` in `docx_mathify.py`. Its unit tail is a
whitelist, not "any letters": each factor must be a member of
`linear_to_latex.UNITS`. That is why `y=2x` and `n=2t` stay equations — they are
function relations, not "2 metres" and "2 tonnes" — and why `v=s/t` stays a real
fraction (`t` is deliberately absent from `UNITS`). `x=3` lands here too, as
`plain_equation` rather than `plain_quantity`. Current conservatism: `T=25℃`
(no `℃` in `UNITS`) and Greek quantity symbols (`η=80%`) stay equations.

### What the plain tier costs

A span that drops to the plain tier loses its math italic — `s` becomes upright.
That is an accepted trade decided on 2026-11-07, not an outstanding defect:
where nothing needs two-dimensional layout, something expressible as text should
not become an equation object. An equation object means opening the formula
editor to change one character, losing shape when pasted elsewhere, and a line
height the formula engine stretches.

When the italic is wanted, supply it from the shape of the input rather than from
the tier: write the form with a script (`a=9.8m/s2`, which lands in the
plain-script tier), or set `vert_align` explicitly on the runs of a `docx_create`
spec — `superscript` / `subscript` booleans are accepted as aliases, and an
explicit `vert_align` wins over them.

### Switching the tiers off

Two module-level constants in `docx_mathify.py`:

| Constant | Default | Set to `False` |
|---|---|---|
| `MIXED_SUPERSCRIPT` | `True` | every span goes to OMML — the master switch |
| `NO_LAYOUT_NEEDED` | `True` | only the no-layout tier stops; the plain-script tier still applies |

`NO_LAYOUT_NEEDED` sits *inside* `MIXED_SUPERSCRIPT`: turning the master switch
off stops both, which is what keeps "`MIXED_SUPERSCRIPT = False` equals the old
behaviour" true.

Two defects in source-text handling were fixed alongside these tiers (consecutive
Unicode superscripts, and a unit after an exponent read as a fraction); both are
written up in [formula-guide](references/formula-guide.md#two-defects-fixed-in-this-round).

Equation font defaults to **Times New Roman**. Override with the `MATH_FONT`
environment variable (`MATH_FONT="Cambria Math"` restores Word's default look;
`MATH_FONT=""` disables injection entirely).

### The `$...$` marker layer

The regex must not be the thing that decides what is a formula. Let the author
mark it instead: wrap LaTeX in `$...$` in the source text, and the parser only
splits — it never guesses semantics.

```python
def _add_mixed(p, text):
    """Parse $...$ markers into OMML, everything else into runs.

    The author knows where the formula is; do not make the regex guess.
    """
    pos = 0
    for m in re.finditer(r"\$([^$]+)\$", text):
        if m.start() > pos:
            _run(p, text[pos:m.start()])
        tex = m.group(1)
        try:
            el = latex_to_omml_lxml(tex)
            fix_phasor_dots(el)      # must run, or the dots float large
            p._p.append(el)
        except Exception as exc:
            FAILS.append((tex, str(exc)[:150]))   # log it, fall back to text
            _run(p, tex)                          # never drop a formula silently
        pos = m.end()
    if pos < len(text):
        _run(p, text[pos:])
```

`FAILS` has to be printed or written to a file. A formula that vanishes without
a trace is the hardest failure to find.

### Phasor dots need a correction pass

`latex2mathml` turns `\dot{I}` into `m:limUpp`, which Word renders as "a large
glyph with a floating dot". `fix_phasor_dot.fix_phasor_dots(omml) -> int`
rewrites it in place as `m:acc` (the `\u0307` combining dot) and returns how many
sites it changed.

Three layers, so nothing is misread:

1. **`latex_to_omml_lxml()` does not fix it.** It is a pure conversion function;
   `\dot{I}` comes out as `limUpp=1 / acc=0`.
2. **Both CLIs already fix it, so the normal path never shows the bad form.**
   `docx_mathify.py`'s `_to_omml()` is the single exit for every conversion, and
   `insert_formula.py` fixes inside `_omml_stdlib()`. Both count their work and
   report `修正相量点 N 处` on the summary line.
3. **Only hand-written callers** — code that imports `math_omml` and converts
   directly — have to add the pass themselves. The upside of doing it yourself is
   that you can then assert `limUpp == 0` as a self-check.

Measured baselines:

```
İ=U/Z 成立                   → 修正相量点 1 / oMath=1 limUpp=0 acc=2 TNR=16
电流 İ_{L}+İ_{C} = İ 成立     → 修正相量点 3 / oMath=1 limUpp=0 acc=6
İ_{L}+İ_{C}（无运算符，不转） → oMath=0   ← 正常：find_math_spans 要求有运算符/等式
```

Self-check: `python scripts/fix_phasor_dot.py`.

## Table geometry

Word resolves column widths in a way that surprises people:

- `tblLayout` absent (autofit) → each cell's `<w:tcW>` acts as a *hint*; columns
  come out proportional to it.
- `tblLayout type="fixed"` → the renderer lays out purely from
  `<w:tblGrid><w:gridCol>` and **ignores `tcW`**. A default grid makes every
  column equal no matter what you set.

`table.autofit = False` writes exactly that `fixed` layout, which is why setting
only `cell.width` silently yields equal columns. All three places have to agree:
`tblW` + `gridCol` + `tcW`, in dxa (1 cm = 566.93 dxa).

`table_width.py` keeps the three consistent. Use `report` first to see the
current state, then `sync` to copy `tcW` values into the grid, `set --widths` to
pin them explicitly, or `auto` to spread the columns over the usable text width
(`--content-width`, default 14.6 cm = A4 minus 3.17 cm margins).

`sync` without reading `report` first is the trap: it faithfully copies whatever
`tcW` currently holds into the grid, wrong values included. And `set --widths`
with more values than the table has columns is a usage error, exit `2` — it
refuses to drop the extras:

```
error: table 1: --widths has 5 values but the table has 3 columns; refusing to drop the extra 2
```

## Fonts: two rules that Word enforces

`fix_fonts.py` unifies fonts across three layers at once — styles, runs, and
manual headings — and binds the CJK face to the `eastAsia` slot while
`ascii` / `hAnsi` / `cs` stay `Times New Roman`. The reason it has to touch the
XML by hand is the two rules below. Break either one and Word opens the file
with "the file may be corrupted", not with a warning and not with a silent
reformat.

1. **Duplicate children inside `w:rPr` are illegal.** `w:rPr` is an
   `xsd:sequence` and most of its children declare `maxOccurs="1"`, so a second
   `w:rFonts` is a schema violation even though parsers happily read it. Real
   Word and WPS both show the repair prompt on a document that carries one.
   `dedupe_rpr(rpr)` collapses each name to its **last** occurrence before
   anything else runs, so later attributes win — the same precedence a lenient
   reader would have applied.
2. **`w:rFonts` must be the first child of `w:rPr`.** It opens the `CT_RPr`
   sequence. After deleting the old `rFonts`, the replacement has to go back in
   with `rpr.insert(0, rf)`; `append` parks it at the tail, which is a legal XML
   tree and an illegal `w:rPr`, and Word rejects it. That line is
   `set_fonts(rpr, ea)`.

```
python scripts/fix_fonts.py in.docx [more.docx ...] [--heading-threshold 14]
```

Runs are dispatched by whether their paragraph is a heading, and a run with no
`w:pStyle` whose font size reaches `--heading-threshold` (default 14 pt) is
treated as a *manual* heading — that is the only way to catch a document whose
author formatted headings by hand instead of applying `Heading 1`. Formula runs
inside `m:oMath` are the exception: only `ascii` / `hAnsi` / `cs` are touched,
`eastAsia` is left alone, because OMML manages its own glyph selection and an
overwritten `eastAsia` visibly degrades the equation.

Files are rewritten in place through a temporary file and `shutil.move`, so an
interrupted run leaves the original intact. **It keeps no extra copy**: unlike
`docx_mathify --inplace` and `docx_template apply --to`, `fix_fonts` leaves no
`.bak` behind, so copy the file yourself if the original matters.

## `<w:zoom>` and its required attribute

python-docx emits `<w:zoom w:val="bestFit"/>` on save, without the
`w:percent` that `CT_Zoom` declares `use="required"`. It is well-formed XML and
invalid OOXML, so a strict validator stops with

```
Element '{...}zoom': The attribute '{...}percent' is required but missing.
```

Word tolerates it; the XSD does not, and any pipeline that validates against the
schema fails on a document this toolkit just produced.

`docx_create.py` patches this itself, right after `doc.save()`: it calls
`fix_zoom.fix_docx()`, so a document it produces already carries
`w:percent="100"` and no manual second pass is needed. Use the CLI directly for
documents that came from somewhere else — python-docx in your own script, an
older file, or a template written by another tool:

```
python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]
```

Three things the script gets right that a naive patch does not:

- A `<w:zoom>` that already carries `w:percent` is returned untouched. Adding a
  second `percent` would be a duplicate attribute — malformed XML, worse than
  missing. It is idempotent: a second run reports `zoom element already valid`.
- Rewriting the zip reuses the original `ZipInfo` entries, so part order and
  compression survive the round trip.
- A package with no `word/settings.xml` is legal, not broken: the script reports
  `nothing to do` and exits `0` rather than inventing the part.

The pure entry point is `fix_settings_xml(xml, percent) -> tuple[str, int]`: XML
in, `(patched_xml, sites_changed)` out, no file I/O, which is what the tests
import directly instead of shelling out.

## Merging two documents

`merge_docs.py base.docx src.docx out.docx` appends `src` to `base` and keeps
styles, numbering, and media. Three decisions are deliberate and worth knowing
before you inspect the result:

- **Page setup comes from `base` only; `src`'s `sectPr` is dropped.** Keeping
  both would not error — Word takes the last section's geometry, so a merge of a
  document with 170 mm text width into one with 160 mm silently reflows the
  whole base document to the wrong measure.
- **Numbering is deduplicated by `w:numId` as a set union**, together with the
  `w:abstractNum` those ids depend on. If `base` has no `numbering.xml` at all,
  the part is copied wholesale plus a `[Content_Types]` override and a
  relationship entry, since those are what actually make the part reachable.
- **Styles are deduplicated by `w:styleId`**; a style that already exists in
  `base` is never overwritten, so the base document's look wins.

Self-checks after a merge: paragraph count equals `base × 2 + 1` (the extra
paragraph carries the page break), the `word/media/*` count is conserved, and
exactly one `sectPr` survives with every attribute equal to `base`'s. If a
merged file looks wrong, those three numbers localise the problem faster than
reading the XML.

## Exporting to PDF (additional capability)

This is not what the toolkit is for, and the frontmatter says so — it is a
delivery path for documents this toolkit *produced*, not a general converter.

It exists for one specific case: **the document contains OMML equations.** WPS
exports those with a broken text layer — the formula area becomes an image or
disappears, and the text around an equation can no longer be searched in the PDF.
Word's own exporter keeps the layer intact, so the second prerequisite is
**Windows with Microsoft Word installed**. Both conditions together, or the
script is not the right tool.

```powershell
pwsh scripts/pdf_export.ps1 -DocxPath in.docx [-PdfPath out.pdf]
```

It drives Word COM: `Documents.Open($path, $false, $true)` — the third argument
is `ReadOnly`, so the source file is not held open and picks up no
"recently opened" state — then `ExportAsFixedFormat($pdf, 17)`, where `17` is
`wdExportFormatPDF`. Everything sits inside `try` / `finally` with
`Close($false)` (do not save) and `Quit()`, because a COM exception without the
`Quit()` leaves a zombie `WINWORD` process holding the file handle, and a locked
directory is much harder to diagnose than an error message. Pass absolute paths;
COM resolves relative ones against the process working directory, which is not
the shell's.

Do not route a document with OMML through WPS.

## Rendering and linting the layout

`docx_preview.py` answers the question the rest of the toolkit cannot: **does
page 4 look right?** It renders the document through a headless LibreOffice,
writes a PDF and one PNG per page, and reports what it sees about the page fill.
No Word, no COM, no Office licence:

```bash
# all three capabilities in one conversion
python scripts/docx_preview.py in.docx --pdf out.pdf --pages png/ --lint --json
# a quick look: page images only, any DPI
python scripts/docx_preview.py in.docx --pages png/ --dpi 150
# machine-readable layout report, no files kept
python scripts/docx_preview.py in.docx --lint --json
```

| Flag | Effect |
|---|---|
| `--pdf OUT.pdf` | write the converted PDF here (default: a temporary file, deleted afterwards) |
| `--pages DIR` | one `page-%03d.png` per page into `DIR`, at `--dpi` (default 110) |
| `--lint` | report the layout: fill ratio, half-empty pages, blank pages, oversized images, split tables, orphaned last lines |
| `--min-fill RATIO` | pages below this share of ink are "sparse" (default `0.40`) |
| `--max-orphan-chars N` | a page's last line of at most N characters is an orphan (default `2`) |
| `--timeout SECONDS` | soffice conversion timeout (default `180`) |
| `--json` / `--text` | one JSON object, or human-readable lines (the default) |

The three capabilities compose: pass all of them and one conversion produces
everything. At least one is required. The `--lint` report names six classes of
problem — `fill_ratio` per page, `sparse_pages`, `blank_pages`,
`image_over_width` (an image wider than the text block), `table_split` (a table
that continues onto the next page) and `orphan_last` (a single character alone on
the last line). `table_split` is a **heuristic and may over-report**: it matches
table geometry across a page boundary, so a table that merely ends near a page
break can be flagged. Read it as "look at this page", not "this is broken".

Requirements differ from the rest of the toolkit: `docx_preview.py` needs
`soffice` on `PATH` (or in a standard LibreOffice location) and PyMuPDF for the
PNG and lint halves. Either one missing is exit `1` with one `error:` line — an
environment that cannot do the job, not a bad argument. `fitz` is imported inside
a function, never at module level, so importing the module never needs PyMuPDF.

> **Manual `soffice` calls need `-Wait`.** If you drive the converter yourself
> instead of through this script, `Start-Process soffice --headless --convert-to
> pdf …` **without `-Wait`** returns immediately: the exit code is empty, the PDF
> has not been written yet, and the next line of your script reads a file that
> does not exist. Add `-Wait`, or use `& soffice …` and let the shell wait.
> `docx_preview.py` does not have this problem because it calls `subprocess.run`,
> which blocks until the child exits.

## Library modules

**`_bootstrap.py` is the one pure library**: no `__main__`, no argparse, imported
by every other script before argument parsing. The seven modules below also exist
to be imported — six ship a one-line `__main__` self-description and
`fix_phasor_dot` ships a self-check, so all seven answer a bare run without doing
any work. Scripts reach them all by bare name, which works because every script
bootstraps its own directory onto `sys.path` (see [Conventions](#conventions)).

| Module | What it provides |
|---|---|
| `linear_to_latex.py` | `find_math_spans` / `atoms_to_linear` / `linear_to_latex` / `repair_merged_scripts`, plus the `UNITS` whitelist |
| `math_omml.py` | `latex_to_omml_lxml` / `latex_to_omml_string` / `assert_legal` / `apply_math_font` / `find_xslt` |
| `fix_phasor_dot.py` | `fix_phasor_dots(omml) -> int`; `python scripts/fix_phasor_dot.py` runs its self-check |
| `docx_pack.py` | `unpack` / `pack` — the round-trip described under [Working on an unpacked document](#working-on-an-unpacked-document) |
| `docx_utils.py` | `wtag` / `register_all` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`, `FONT_HEITI`, `FONT_SONG` |
| `encoding.py` | `_fix_stdio()` — switches stdout and stderr to UTF-8 on Windows |
| `minidom_compat.py` | `defusedxml.minidom` with a stdlib fallback |

`docx_common.py` deliberately is **not** in that table: it has a real argparse
surface (`--api` prints the helpers it exports), so it counts as a CLI.

`minidom_compat` matters more than it looks. `defusedxml` hardens XML parsing
against entity expansion, but a bundled interpreter — the DSH runtime, a portable
Python, a slim container — often has no third-party packages at all, and then
every script that touches docx XML dies at *import* time. The module installs a
stand-in built on `xml.dom.minidom` when the real package is missing, so call
sites stay identical; parsing is simply unhardened, which is an acceptable
trade for documents this toolkit wrote itself.

`_bootstrap.py` is the single entry that wires all of this up, and **the first
line of every new script is `import _bootstrap`** — one import covers `sys.path`
self-location, the minidom fallback, and the UTF-8 stdio switch. `encoding.py`
is the concrete implementation of that last part; import it directly only when
you genuinely want the stdio fix without the rest.

## Reference notes

The body of this file covers the decisions. Measured detail — recipes, trap
lists, schema quotes — lives in `references/`, loaded only when you need it.

| File | Read it when |
|---|---|
| `references/formula-guide.md` | writing formula-chain production code, calling `_add_mixed`, debugging Word's "file may be corrupted", handling `%` / `℃` / `～` / absolute-value / compound-unit source traps |
| `references/table-geometry.md` | column widths not taking effect, deciding between `set` and `sync`, `<w:zoom>` missing `w:percent`, keeping a table on one page (`keep_together`), sizing an image by `width_frac` |
| `references/font-scheme.md` | mixed CJK/Latin font conventions, `w:rPr` ordering and dedupe, what `--heading-threshold` means, why Heading/Title/Subtitle come out black and how to colour one deliberately |
| `references/merge-and-export.md` | after merging: page area / numbering / images wrong; Word COM export details; multi-page PDF via PIL; driving `soffice` by hand (`-Wait`, `docx_preview.py`) |
| `references/revisions-and-comments.md` | OOXML representation of comments and tracked changes, and the CLI semantics |
| `references/toc.md` | building a static TOC: the two-pass recipe (`--page-probe` first, then write entries), tab stops and dot leaders, the full trap list, the five page-number footer templates (`zh` / `custom` placeholders) |

## Layout

```
scripts/     flat, imported by bare module name — 27 .py + 1 .ps1 = 28 files (26 CLIs, 1 pure library)
references/  longer how-to notes, loaded on demand — 6 files
tools/       Word COM acceptance helper; --page-probe lives here, in word_accept.py
tests/       pytest suite — 58 .py (57 test_*.py + _omml.py helper)
```

## Exit codes

| Code | Meaning |
|---|---|
| `0` | success |
| `1` | expected failure — a marker that did not match, an unknown or duplicate bookmark name, "the document does not contain what you asked for", a runtime exception |
| `2` | usage error — argparse, an unknown key or unknown block type in a spec, more `--widths` values than columns, a `--pages` count that does not match the entry count, mutually exclusive flags passed together (including `width_frac` with `width_mm`), an `--index` / `--insert-at` outside the legal range (or negative), an input path that is missing, a directory, or not a Word package |
| `3` | **`docx_mathify` only** (and `docx_template apply --fix-formulas`, which passes it through): the document was written, but at least one formula failed to convert |

A usage error is a **refusal, not a repair**: it is decided before anything is read
or written, the input file is byte-for-byte untouched, the message is one
`error: …` line on stderr, and stdout stays empty.

The input-path clause in row `2` is scoped to the CLIs that open a document, and it
is not a claim that the whole tree answers the same way: `docx_validate` is the
deliberate exception, because "this package is broken" is its *answer* rather than
a misuse of the tool — it prints `{"ok": false, …}` on stdout and exits `1`, so
read its JSON, not its exit code. Every other failure that a document-reading CLI
reports as `1` (a marker that did not match, a bookmark that is not there) happens
*after* the path was accepted, which is exactly the line row `2` draws.

`docx_mathify` returning `3` is deliberate: the output file is valid and mostly
converted, but at least one equation is still linear text. Check the exit code if
you are delivering the result. "Missing XSL" does not always mean `3` — a
document whose spans are all degradable (text runs, text scripts) never touches
the XSL and still exits `0`.

`docx_read` is the sharpest case of the stdout rule, because *every* one of its
success modes writes JSON: on a missing file, a directory, a non-ZIP, or a ZIP
that is not a Word package it exits `2` with one line on stderr and **nothing at
all on stdout**. Relaying the failure through stdout as JSON would make every
`json.loads(stdout)` caller read a failure as a success. "The document is real but
has a problem" is `docx_validate`'s job and exits `1` instead.

## Conventions

- Every script bootstraps its own directory onto `sys.path`, so it runs from any
  working directory — `python scripts/x.py` and `python x.py` both work.
- Machine-readable scripts write JSON to stdout; human-facing ones write prose.
  Diagnostics and errors always go to stderr.
- Exit code reflects the outcome; check it before parsing.
- Files are read and written as UTF-8 explicitly; stdio is reconfigured to UTF-8
  on Windows (`_bootstrap.py`).
- `requirements.txt` is pure ASCII on purpose — pip reads it with the system
  locale encoding, and a stray em dash breaks the install on a CJK Windows.

## Hard rules

- Formulas in reports and formal documents go through LaTeX→OMML. Only
  `H₂O`-level sub/superscripts use a `vertAlign` run instead.
- Every script runs under `python -X utf8`. On Windows `_bootstrap` already
  switches stdio to UTF-8 — do not add a second `reconfigure()`.
- Back up before an in-place rewrite. `docx_mathify --inplace` writes
  `<name>.bak.docx` and `docx_template.py apply --to` without `-o` writes
  `<file>.bak`; `fix_fonts.py`, `fix_zoom.py`, and `merge_docs.py` keep no copy,
  so copy the file yourself when the original matters.
- Only `docx_template.py apply -o` needs node (docx-js skeleton). When the chain
  is broken, use `apply --to` instead.