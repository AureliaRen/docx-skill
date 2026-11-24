---
name: docx
description: "Create, read, edit, validate, comment on, and rewrite Word documents (.docx/.dotx) with a set of small Python CLIs. Use whenever the task involves a Word file: generating a report/memo/letter from a spec, extracting or inspecting existing content, find-and-replace across a document, editing tables, managing comments and tracked changes, filling {{token}} templates, extracting/applying formatting templates, fixing table column widths, converting linear formulas into native Office Math (OMML), or writing a document whose equations start as LaTeX. Also use when a deliverable should be a .docx. Do NOT use for PDF, spreadsheets (XLSX/CSV), slide decks, or plain one-shot text extraction (use markitdown)."
license: MIT
---

# Word documents (.docx)

Scripts live flat in `scripts/` — **28 `.py` + 1 `.ps1` = 29 files**, no
sub-package, no `__init__.py`. **27 are CLIs** (each carries an
`if __name__ == "__main__"` guard); **`_bootstrap.py` is the one pure library**.
Of the 27, **20 expose an argparse surface** with `--help` and the other 7 print a
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

### Importing the scripts from your own code

`scripts/` is flat on purpose: no package, no `__init__.py`. A sibling import
(`from math_omml import …`) therefore resolves **only** when that directory is
itself on `sys.path`. Every script in the tree bootstraps its own directory
before importing anything, which is why `python scripts/docx_read.py` runs from
any working directory; a script of your own, or a `python -c` line at the
repository root, has no such bootstrap and stops at the first sibling import:

```
ModuleNotFoundError: No module named 'math_omml'
```

Two copyable forms. Put the directory on the path yourself:

```python
import sys
sys.path.insert(0, r"<path-to-skill>/scripts")   # the flat directory itself
import _bootstrap                                # noqa: F401
```

or, when `scripts/` is already reachable (a `PYTHONPATH` entry, or a script that
lives there), one line — `_bootstrap` re-inserts its own directory idempotently,
installs the `minidom` fallback and switches stdio to UTF-8:

```python
import _bootstrap  # noqa: F401 — sys.path + minidom fallback + UTF-8 stdio
```

Either way, sibling modules then import by bare name. A minimal runnable pair,
measured on this tree (both imports resolve, and `\dot{I}` comes out with
`limUpp=0`):

```python
from math_omml import latex_to_omml_safe    # convert + fix phasor dots + self-check
from fix_phasor_dot import fix_phasor_dots  # only for a hand-rolled pass, or to assert limUpp == 0

el = latex_to_omml_safe(r"\dot{I} = \frac{U}{Z}")
paragraph._p.append(el)     # paragraph is a python-docx Paragraph
```

Without the `sys.path.insert` line the same import raises `ModuleNotFoundError`
— measured from the repository root with no `scripts/` on the path.

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

Or import it where you already have Python — with the same one line the command
above has, because `scripts/` is flat and nothing puts it on `sys.path` for you:

```python
import sys; sys.path.insert(0, r"<skill>/scripts")   # or: import _bootstrap
from docx_pack import unpack, pack
unpack("in.docx", "unpacked/")
pack("unpacked/", "out.docx", original="in.docx")
```

That caveat is general, not specific to `docx_pack`; it has its own section —
[Importing the scripts from your own code](#importing-the-scripts-from-your-own-code).

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
| Write a document with equations | a `{"type": "latex", "latex": "\\frac{a}{b}=c", "display": true}` block — see [Writing a document that contains formulas](#writing-a-document-that-contains-formulas) |
| Read text / outline / styles | `python scripts/docx_read.py in.docx --text` |
| Extract images | `python scripts/docx_read.py in.docx --images DIR` |
| Inspect tracked changes | `python scripts/docx_read.py in.docx --revisions` |
| Inspect run formatting (before writing into a cell) | `python scripts/docx_read.py in.docx --runs` |
| Replace text | `python scripts/docx_edit.py replace in.docx --find A --replace B -o out.docx` |
| Set a table cell | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X -o out.docx` |
| Set a cell, dropping all its formatting | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X --replace-cell -o out.docx` |
| Insert a paragraph | `python scripts/docx_edit.py insert in.docx --index 3 --text "..." -o out.docx` |
| Delete a paragraph | `python scripts/docx_edit.py delete in.docx --index 3 -o out.docx` |
| Apply a paragraph style | `python scripts/docx_edit.py style in.docx --index 0 --style "Heading 1" -o out.docx` |
| Merge adjacent identical runs | `python scripts/docx_edit.py normalize in.docx -o out.docx` |
| Highlight / colour / bold / italic an anchored string | `python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta --bold -o out.docx` |
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

The total is a **field**, not literal text, and that changes how you verify it.
`第 1 页　共 <NUMPAGES> 页` reaches `footer*.xml` with a cached total of `1`,
because how many pages a document has is not known while the package is still
being written; Word and LibreOffice recalculate the field when the document
opens. On a five-page document the PDF says `共 5 页` and the XML still says `1`.
Check a footer in the rendered page, never in the raw XML — an XML check that
expects the real total fails on a document that is perfectly correct.

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
| `crossref` | `target` (a bookmark), optional `prefix`, `suffix`, `format` (`default`/`full`/`above-below`/`page`/`number`/`paragraph-number`/`relative-context`/`full-context`), `cached`, `own_paragraph`, `superscript` / `vert_align`, `hyperlink`, `static`, `no_mergeformat` |
| `table` | `header`, `rows`, optional `style` (default `Table Grid`), `header_bold`, `keep_together` |
| `image` | `path`, and `width_mm` **or** `width_frac` (mutually exclusive) |
| `page_break` | — |
| `latex` | `latex` (required, a non-empty string), `display` (bool, default `false`) — one native Office Math equation per block; see [Writing a document that contains formulas](#writing-a-document-that-contains-formulas) |
| `section_break` | `start_type`, optional `page`, `header` / `footer` (+ their `_align` / `_font` / `_size_pt` / `_bold`), `footer_page_numbers` (`fmt` / `start` / `align` / `template` / `total` / `prefix` / `suffix` / `page_type`, plus `font` / `size_pt` / `bold` for the page-number runs — see the `bold` table above, `false` reaches the literal runs but not the field runs), `different_first_page`, `cols`, `no_footer`, `reuse_last_paragraph`, `apply_to` |
| `toc` | optional `instr` — default `TOC \o "1-3" \h \z \u`; pass `TOC \h \c "图"` for a table of figures. **A field**: Word must update it (F9) before any entry exists |

**The spec is validated before a single paragraph is built, and an unknown key
or an unknown block type is an error, not a no-op.** `{"blockz": [...]}` used to
build a valid, empty document, print `"blocks": 0` and exit `0`, so the author
had no way to see that the content never went in. It now exits `2` with one line
naming the offending key and, when it is close enough, a "did you mean" hint:

```
spec: unknown key "blockz" (did you mean "blocks"?); valid keys here: blocks, footer, footer_page_numbers, header, page, styles
blocks[0]: unknown block type 'paragraf' (did you mean "paragraph"?); choose one of: bullet_list, caption, crossref, heading, image, latex, numbered_list, page_break, paragraph, section_break, table, table_caption, toc
```

Exit `2` promises that **nothing was written**, so you can fix the spec and
retry without wondering whether half a document exists. Values are not clamped
either: an out-of-range value is reported with the range it would have been
clamped into.

**A spec key must be the name the target command actually reads, and that is not
free.** The validator only proves a key is spelled as *this* table spells it; the
builder then hands the value to `cmd_ref` / `cmd_caption` through a namespace
object, and those read it with `getattr(args, "name", default)` — an idiom that
cannot distinguish "absent" from "false". A mismatched name is therefore a
**silent no-op**: measured before it was fixed, `{"type": "crossref", "target":
"g1", "hyperlink": false}` still produced a document containing
`<w:hyperlink>`, with `exit 0` and `{"ok": true}`. The spec says `hyperlink`
(positive, default true); `cmd_ref` reads `no_hyperlink` (negative), so the
builder translates rather than forwards. `tests/test_create_crossref_args.py`
guards the whole set by parsing the `getattr` names out of `docx_crossref.py`
and failing if the builder stops supplying one.

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

### Writing a document that contains formulas

Two routes. The first is a spec key; the second is a few lines of `python-docx`
when you already own the paragraph handling.

**Route 1 — a `latex` block.** One block, one equation, native Office Math:

```json
{"blocks": [
  {"type": "heading", "text": "Ohm's law", "level": 1},
  {"type": "latex", "latex": "E = mc^2"},
  {"type": "latex", "latex": "\\frac{a}{b} = c", "display": true}
]}
```

`latex` is required and must be a non-empty string; `display` is optional and
must be a real JSON boolean — `"display": "yes"` is a spec error (exit `2`, one
`error:` line, nothing written), because it is truthy and would otherwise turn an
inline equation into a display one without a word. The block also accepts the
paragraph format keys listed above (`alignment`, `space_before_pt`,
`keep_with_next`, …), and it converts through `latex_to_omml_safe()`, so
`\dot{I}` arrives as `m:acc` instead of the `m:limUpp` Word draws as a large
glyph with a floating dot.

**Route 2 — hand-written `python-docx`.** The same conversion, with your own
paragraph handling. This skeleton runs as written (the imports come from
[Importing the scripts from your own code](#importing-the-scripts-from-your-own-code)):

```python
import sys
sys.path.insert(0, r"<path-to-skill>/scripts")
import _bootstrap  # noqa: F401 — sys.path + minidom fallback + UTF-8 stdio

from docx import Document
from math_omml import latex_to_omml_safe

FAILS = []
doc = Document()
for tex, display in [(r"E = mc^2", False), (r"\frac{a}{b} = c", True)]:
    p = doc.add_paragraph()
    try:
        p._p.append(latex_to_omml_safe(tex, display=display))
        if not display:
            p.add_run("\u200b")      # see (a) below
    except Exception as exc:          # log it, then fall back to text
        FAILS.append((tex, str(exc)[:150]))
        p.add_run(tex)
doc.save("out.docx")
print("FAILS:", FAILS)                # never let a formula vanish silently
```

Three behaviours belong to both routes and cost an hour each when they surprise
you:

**(a) `display: false` is an inline equation *plus* a zero-width run.** The
paragraph that carries an inline equation holds `m:oMath` **and** one U+200B run.
Without it, an equation that is the only content of its paragraph is *promoted*
by Word to a display equation — centred, `\sum` stacked — and the layout moves
under the author. `docx_read.py --blocks` then reports a paragraph whose `text`
is a single invisible character, which is the expected reading, not damage. Same
trick and same reason as `insert_formula` rule 2 and the `docx_mathify` cell
branch.

**(b) One equation that fails does not abort the document.** Measured on a spec
whose third formula is `\frac`:

```json
{"ok": false,
 "output": "…/out.docx",
 "blocks": 4,
 "formula_failures": [{"block": 2, "latex": "\\frac",
                       "error": "RuntimeError: latex2mathml 转换失败: '\\\\frac' -> "}]}
```

(Shown decoded: in the bytes on the wire every non-ASCII character inside a JSON
string is `\u`-escaped, which changes nothing about the shape.)

One line goes to stderr —

```
warning: 1 latex block(s) did not convert; the document was written with the LaTeX source left as plain text
```

— the exit code is **3**, and the paragraph that failed holds the LaTeX source
as ordinary text rather than an empty gap. `ok` is `false` in that payload while
`output` still names the file that exists: **branching on `ok` alone will
conclude nothing was written, so read `output`.**

**(c) Conversion needs `latex2mathml` *and* `MML2OMML.XSL`.** `math_omml` is
imported inside the block handler on purpose, so `docx_create` still imports and
every formula-free spec still works on a machine that has neither. Measured with
`MML2OMML_XSL` pointing at a nonexistent file: three `latex` blocks, three
entries in `formula_failures`, the document written with all three equations left
as LaTeX text, exit `3` — the same outcome `docx_mathify` gives when the XSL is
missing.

One calibration, because it costs an hour otherwise: **`latex2mathml` tolerates a
missing closing brace.** `\frac{a` converts into a perfectly normal `m:oMath`;
what actually raises is `\frac`, `\sqrt`, `\left(`, `x^{`, `\mathrm`,
`\overline{`, and the empty string. Do not use `\frac{a` as the broken sample in
a test or a validation step — it passes, and the test that uses it is green for
the wrong reason.

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

`path` resolves against the **current working directory**, not against the
spec file's directory, and a path that does not exist is refused as an input
error — exit `2`, the block index in the message, nothing written — instead of
surfacing later as a `FileNotFoundError` traceback from inside the build.

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

### `ref --format`: the eight values, and the field each one writes

`--format` picks which Word field the reference becomes. There are eight legal
values; the first five cover the ordinary cases and the last three are the ones
that only appear when a document needs them.

| `--format` | Writes | Reads as |
|---|---|---|
| `number` | a bare `REF <bookmark>` number | `3` |
| `default` | `\h` — a hyperlink to the bookmark | `3`, clickable |
| `full` | the whole caption text (needs the bookmark to span the paragraph) | `图 3 一张…` |
| `above-below` | `\p` — the relative-position word | `上图` / `下图` |
| `page` | `PAGEREF` — the page the bookmark falls on | `7` |
| `paragraph-number` | `\n` — the paragraph number | `[1]`, the one the reference list shows |
| `relative-context` | `\r` — the number relative to the nearest numbered item | `relative to the enclosing heading` |
| `full-context` | `\w` — the full context of the bookmark | the number plus its enclosing context |

The four switches are the difference between "a number" and "a number a reader
can act on": `\p` answers *where*, `PAGEREF` answers *on which page*, `\n`
answers *which numbered item*, `\r` answers *relative to what*. Pick the field by
the question the sentence asks, not by which output looks tidiest.

`--format full` needs the bookmark span. `--format page` cannot be combined with
`--static` (see below): the page number does not exist until pagination does.

### Superscript references: field or static, and what each one costs

A superscript citation like 图¹ is the case where the two designs visibly
differ, and the choice is not cosmetic. `--superscript` writes
`<w:vertAlign w:val="superscript">` on the runs of the reference. Whether that
survives to the PDF depends on which of two shapes you asked for.

| | `ref` (default, a field) | `ref --static` |
|---|---|---|
| Markup | `<w:hyperlink>` around a 5-run `REF` field | `<w:hyperlink>` around **one** `<w:r>` holding the number as `<w:t>` |
| Number tracks its caption | yes — F9 / print recomputes it | **no** — frozen at write time |
| Survives `Fields.Update()` | yes, **but only with `\* MERGEFORMAT`** | yes, unconditionally: there is no field to re-render |
| Needs F9 before it reads right | no (cached result written) | never |
| Add a figure above it | renumbering is automatic | re-run the command |

The middle row is the one that bites. Measured on Word 16.0, full
`Fields.Update()` then PDF export, spans read back with PyMuPDF:

```
field + \* MERGEFORMAT     '1'  sz 7.0  y 197.18     raised 2.52 pt, half size
field, --no-mergeformat    '1'  sz 11.0 y 199.70     the superscript is gone
--static                   '1'  sz 7.0  y 197.18     identical to row 1
body text                  sz 11.0 y 199.70          the baseline, for reference
```

So `MERGEFORMAT` is not decoration: without it Word re-renders the field result
from the bookmark target and the raised look is lost on the next update — which
is exactly what a PDF export performs. It is written by default; `--no-mergeformat`
turns it off and is only for reproducing that failure.

`--static` is for when you want the look to be **structural** rather than
dependent on a switch — for example when something else in the pipeline rebuilds
fields, or when the document must read correctly in a viewer that never updates
fields. Take it knowing the number stops tracking: adding or removing a figure
means re-running the build. For a document still being edited, the field form is
the right default; `--static` is for the last pass.

`--static` refuses `--format page` and `--format above-below` (exit `2`): a page
number depends on pagination that does not exist at build time, and 上方/下方 is
decided by Word per occurrence. Freezing either would write a plausible-looking
wrong number.

Both modes report which one they used — `"mode": "field"` with
`"needs_field_update": true`, or `"mode": "static"` with
`"needs_field_update": false` — so a build script can assert it got the shape it
asked for instead of inferring it from the markup.

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

### What counts as evidence (and the fraction characters)

A span has to earn its way in. The default demand is a relation operator
(`=`, `≈`, `≤`, `±`) or a strong operator (`√`, `∑`, `∫`, `∂`, `∇`). Two things
are exempt because they are the author's own declaration rather than our guess:

- an explicit sub/superscript group — `S_{甲}`, `m^{2}`;
- **a vulgar fraction character** — `½ ¼ ⅓ ⅕ ⅛ ⅐ ↉` and the other eleven.

The second one is worth spelling out, because the character is a hazard rather
than an error. U+00BD is a single code point that fonts render as a small `1`,
a slash and a lowered `2`: it looks *nearly* right in one font and wrong in
every other, so nothing gets reported and the problem surfaces at print. All
nineteen Unicode fraction code points translate to `\frac{n}{d}`, and one of
them alone is enough to open a span — `½+1` converts even though it contains no
`=`, and so does a lone `½` between Chinese characters (`第½章`, with the
framing left as text).

A fraction character is also a **value anchor** for the unit grammar, which is
what keeps the two meanings apart:

```
½ m/s     → \frac{1}{2}\,\text{m/s}     measurement: "half a metre per second"
0.5 m/s   → 0.5\,\text{m/s}             the same quantity, same layout
v=m/s     → v=\frac{m}{s}               a formula: m divided by s
m/s       → \,\text{m/s}                a bare unit string
```

The rule behind that column is not "both sides look like units" — `m/s` and
`v=m/s` are identical on that test. It is an anchor check: a unit string is a
*measurement* when a number (`0.5`, `10⁻⁷`, `½`, a closing paren or a power)
sits to its left. So `½ m/s` and `0.5 m/s` take the unit path, while `v=m/s`
does not and stays a fraction.

One consequence worth knowing when you read the intermediate output: a spacing
command (`\,` `\;` `\:` `\!` `\ `) is **one token**, not a backslash plus
punctuation. It has to be, or the operand scan of a fraction pulls the backslash
into the denominator and the brace closes early — `1/2 m/s` would come out as
`\frac{1}{2\},\text{m/s}`, whose denominator renders as the literal `2},`. That
was a live bug until this round; the shape that triggers it is "a fraction
followed by a unit string", which is most of a physics hand-out.

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
3. **Hand-written callers use `latex_to_omml_safe()` and add nothing.** That
   wrapper is conversion + `fix_phasor_dots` + a second `assert_legal`, so
   `\dot{I}` arrives as `m:acc` and `limUpp == 0` holds; `docx_mathify`,
   `insert_formula` and the `docx_create` `latex` block all go through it.
   Reaching for the bare `latex_to_omml_lxml` is now a deliberate choice, and its
   docstring opens with a ⚠️ saying so: pure conversion, no dot repair, no
   post-processing. Take that path only when you want the untouched OMML for a
   pass of your own — and then run `fix_phasor_dots(el)` yourself, which is also
   where you can assert `limUpp == 0` as a self-check.

Measured baselines:

```
İ=U/Z 成立                   → 修正相量点 1 / oMath=1 limUpp=0 acc=2 TNR=16
电流 İ_{L}+İ_{C} = İ 成立     → 修正相量点 3 / oMath=1 limUpp=0 acc=6
İ_{L}+İ_{C}（无运算符，不转） → oMath=0   ← 正常：find_math_spans 要求有运算符/等式
```

Self-check: `python scripts/fix_phasor_dot.py`.

## Editing a table cell without destroying its formatting

`set-cell` writes a *flat string* into a cell. Templates are rarely flat: one
cell often holds several runs, and only one of them carries the formatting that
makes the line readable. A template row like

```
v₀ = 12.5 m/s      ← three runs: "v" plain, "0" subscript, "= 12.5 m/s" bold
```

must not be overwritten with `cell.text = ...`. That setter is documented as
"Any existing content or revisions are replaced": it clears the cell, adds one
paragraph, adds one run, and drops every run-level property — paragraph
alignment, size, font, `vertAlign`, everything. Worse, it does so silently,
with `exit 0` and no warning, and the loss only shows up when a human opens the
file.

So `set-cell` defaults to `mode: "preserve"` and reports what it had to do:

| Situation | What happens | `mode` |
|---|---|---|
| Cell holds exactly one writable run | The run's text is replaced; its `rPr` is untouched | `preserve` |
| Cell holds several runs | The first writable run receives the text; the rest are **removed**; `vertAlign` is stripped from that run when it carries one, so the new text does not inherit a stray subscript | `preserve` + `warnings` |
| `--replace-cell` | Old behaviour, exactly: clear, one paragraph, one run, no formatting | `replace-cell` |

A run is *not writable* when it is a field (`w:fldChar` / `w:instrText`), a
drawing, or a picture — those are left alone rather than clobbered, because
replacing a field run breaks the field result it produced.

Read the cell before you write into it:

```bash
python scripts/docx_read.py in.docx --runs
```

which prints, per paragraph, a `mixed_format` flag and one row per run with
`text` / `vert_align` / `bold` / `italic` / `size_pt` / `font` / `is_field` /
`is_drawing`, plus a `path` such as `table[0].r0.c0` so you can find the cell in
the output. `mixed_format: true` means "a flat string cannot express this cell"
— either pass `--replace-cell` on purpose, or rewrite the cell contents through
`python-docx` with a deep copy of the run you want to keep (see
`references/editing-traps.md`).

Anything the call could not keep is reported on stderr as a `warning:` line and
in the JSON `warnings` array; `exit 0` alone does not mean the cell came out
identical.

## Character formatting: highlight / colour / bold / italic on an anchor

`set-cell` and `replace` write text. `char-format` changes only the *appearance*
of text that is already there, and it addresses the text by content, not by
index:

```bash
python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta
python scripts/docx_edit.py char-format in.docx --find "重点" --color 00a0ff --bold -o out.docx
python scripts/docx_edit.py char-format in.docx --find "关键词" --bold --italic -o out.docx
```

**Scope is an anchor string, and every occurrence is formatted.** `--find`
matches against the text of each paragraph (body, tables, nested tables, headers
and footers) and every occurrence counts — measured on two paragraphs each
holding `重点`: `{"occurrences": 2, "runs_changed": 2,
"paragraphs_touched": 2, "paragraphs_searched": 8}`. There is no `--index`:
`docx_read --runs` numbers paragraphs across body *and* cells *and* boxes while
`doc.paragraphs` numbers body paragraphs only, so "index 7" would mean two
different paragraphs in the same tool pair. Text is also the durable selector —
an edit shifts every later position, and a shifted position still "works" while
formatting the wrong line.

**Runs are split when the anchor is only part of one**, so the characters
around it keep the formatting they had. Measured on `本段包含重点两个字，其余不变。`
after `--find 重点`: `--runs` reports three runs, `本段包含` / `重点` /
`两个字，其余不变。`, and only the middle one carries the new properties. A run
that cannot be split (it also holds a tab, a break, a field or a drawing) is
refused rather than formatted whole, because "highlight a bit more text than you
asked for" is indistinguishable from a bug at your end.

### `--highlight` values, and the ones python-docx cannot name

Legal `w:highlight` values (all 17, from `ST_HighlightColor`):

```
black blue cyan darkBlue darkCyan darkGray darkGreen darkMagenta darkRed
darkYellow green lightGray magenta none red white yellow
```

**Spelling is free, the value is not.** `darkMagenta`, `DARK_MAGENTA`,
`dark_magenta` and the python-docx member name `VIOLET` all normalize to
`darkMagenta` — measured: `--highlight violet` exits `0` and stores
`w:highlight w:val="darkMagenta"`. The command takes the alias route on purpose:
`WD_COLOR_INDEX` has 18 members of which only 16 are legal here (`AUTO` writes
`default`, `INHERITED` writes nothing) **and it is missing `darkMagenta`**, which
is the colour this feature was asked for; `VIOLET` writes it under a name nobody
uses. So the enum is read for its `xml_value` as an alias source, never used as
the gate — the 17-value list is.

The report echoes the **normalized** value both inside `applied` and at top level,
so confirming what actually landed does not require opening the XML:

```json
{"find": "重点", "occurrences": 2, "runs_changed": 2, "paragraphs_touched": 2,
 "paragraphs_searched": 8, "applied": {"highlight": "darkMagenta", "color": "00A0FF", "bold": true},
 "highlight": "darkMagenta", "color": "00A0FF", "bold": true, "ok": true, "output": "out.docx"}
```

`--color` takes six hex digits with or without `#` (`FF0000`, `#00a0ff`) or
`auto`; it is upper-cased on write (`#00a0ff` lands as `00A0FF`) and replaces a
theme colour on that run, because in `CT_Color` a theme attribute wins over the
RGB twin at render time.

### `--highlight none` and `--highlight inherit` are not the same thing

Both leave the page looking unhighlighted on a plain run, which is exactly why
the difference has to be stated. Measured on a run that already carries
`w:highlight w:val="yellow"`:

| Flag | `w:rPr` after the call | Meaning |
|---|---|---|
| `--highlight none` | `<w:rPr/>` — the element is **removed** | explicitly no highlight; a style-level highlight is overridden |
| `--highlight inherit` | `<w:rPr><w:highlight w:val="yellow"/></w:rPr>` — untouched | nothing is written; whatever the style says applies |

`inherit` is the "don't touch this flag" option: it exists so one call can set
`--color` (or `--bold`) without silently clearing a highlight somebody else put
there. `none` is a deliberate removal. Both are legal *values* of `--highlight`;
neither writes a fake `w:val`.

For bold and italic the same three-state logic applies, with `w:val="0"` as the
explicit off: `--bold` writes `<w:b/>`, `--no-bold` writes `<w:b w:val="0"/>`
(measured: `<w:rPr><w:b w:val="0"/><w:highlight w:val="yellow"/></w:rPr>`), and
omitting both writes nothing. Only the Latin twin `w:b` / `w:i` is written, never
`w:bCs` — a `w:bCs` that disagrees with `w:b` is its own kind of surprise.

### Refusals: `exit 2`, one line on stderr, nothing written

Every rejection below happens **before** `doc.save()`, so the input file is
byte-for-byte unchanged and stdout stays empty:

| Case | stderr (measured) |
|---|---|
| a value outside the 17 | `error: --highlight 'purple' is not a legal w:highlight value; legal: black, blue, cyan, … (any spelling, e.g. darkMagenta / DARK_MAGENTA / dark_magenta)` |
| an anchor no run carries | `error: char-format --find '不存在的字': no run in the document carries this text. Searched 8 paragraph(s) across body, tables, headers and footers (text boxes and tracked-change content are out of scope, as for every other run-level writer).` |
| no format flag at all | `error: char-format needs at least one format to set: --highlight, --color, --bold/--no-bold, --italic/--no-italic` |
| `--color` that is not hex | `error: --color 'red' is not a colour: expected 6 hex digits (FF0000 or #00a0ff) or 'auto'` |

A misspelt colour used to be written verbatim, which produced a document that
validates, opens, and shows **no highlight at all** — nothing to explain why. An
anchor that matches nothing is refused for the same reason `replace` is
(see [Never rewrite text through a shell string
replacement](#4-never-rewrite-text-through-a-shell-string-replacement)): a silent
zero looks like success.

### Verify by reading it back

Setting is not evidence; reading it back is. `docx_read --runs` reports
`highlight` and `color` per run, which is the pair this command writes:

```bash
python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta -o out.docx
python scripts/docx_read.py out.docx --runs
```

Measured read-back after the split: the run that held `重点` reports
`{"text": "重点", "highlight": "darkMagenta", "color": "00A0FF", "bold": true}`,
and its two neighbours report `null` for all three. That null-in-the-neighbours
part is the real check — it is what proves the split did not bleed formatting
into the surrounding text.

## Floating text boxes

A text box's paragraphs are not body paragraphs. They live in
`w:txbxContent`, whose only parents are `wps:txbx` (inside a DrawingML
`w:drawing`) or `v:textbox` (inside a VML `w:pict`) — never a direct child of
`w:body`. So `doc.paragraphs` does not list them, `body.iterchildren()` does not
reach them, and `Paragraph.text` stops at its own `w:r` children: **every mode
that walks body paragraphs reads such a document as empty**. Measured on a real
annotation file: 43 comments, all in floating boxes, and `--text` / `--blocks` /
`--runs` each reported zero of them — the reason a `python-docx` sweep once
concluded "43 coloured boxes, no annotation text".

All read modes now cover them:

| Mode | Where the box text appears |
|---|---|
| `--text` | a separate `textboxes` key (one string per box); `body` keeps its old meaning |
| `--blocks` | a `textbox` block right after the paragraph that anchors it, with `index` / `flavor` / `text` / `paragraphs` |
| `--runs` | the box's own paragraphs, each carrying `textbox: <n>` and a `path` of `body textbox[<n>]` |
| `--md` | the box text as a quoted (`>`) block |

Body paragraphs and table cells are untouched by this: on a document with no
boxes, the block list is what it was, `body` keeps its meaning, and every
per-paragraph run field is unchanged — the only additions are an empty
`textboxes` list and a `textbox: null` field on each `--runs` record. That last
part is why "unchanged" must not be read as "byte-identical": `--runs` gained a
key, so its output is a **superset**, not the same bytes.

Indices move in a document that *does* have boxes, and that is unavoidable:
a `textbox` block is inserted after its host paragraph, so every entry after the
insertion point shifts down (measured: six plain paragraphs numbered 0-5; four
boxes anchored inside paragraph 1 moved paragraph 5 from index `4` to `12`,
because each box contributes its own paragraph record plus a `textbox` block). It
is worth knowing before you cache an index across an edit — the block's `type` /
`text` / `style` keys are the stable part, not its position.

Two things to know when you read the output:

- **One box, one entry.** Word and WPS write a box as an
  `mc:AlternateContent` holding an `mc:Choice` (DrawingML) **and** an
  `mc:Fallback` (VML) that both carry the same text. The reader takes the
  Choice when it has text and falls back only when it does not, so a 43-box
  document reports **43**, not 86. If you write your own traversal, collapse
  that pair or you will silently double every count.
- **`index` is the box ordinal**, shared by `--blocks` and `--runs`, so the two
  modes name the same box with the same number.

`docx_footer.py` reads boxes in header/footer parts and measures their geometry;
`docx_read.py` reads box **text** in the body. Neither duplicates the other.

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

## Workflow: the checks after every save

Four rules, each earned by a real delivery that went wrong. A reading-annotation
assignment was laid out in Word and had to be handed over. The author had this
toolkit in hand, used none of its checkers, and instead looped *PDF → PNG → look
at the picture → measure a y coordinate by hand* for **a dozen rounds of
rework**. One `--lint` run on the finished file then named the same defects he
had been chasing by eye:

```
lint: 2 issue(s), 5 page(s)
lint: sparse_pages=[3, 5] blank_pages=[]
  warning sparse-page page 3: fill_ratio 0.213 < 0.40 (content covers 21.3% of the page height; …)
  warning sparse-page page 5: fill_ratio 0.117 < 0.40 (content covers 11.7% of the page height; …)
```

Twelve rounds of hand-measuring; one command said the same thing. Current builds
also print `content_width_mm` on that first `lint:` line. The order below is by
return per second spent.

### 1. Lint after every save — highest priority

```bash
python scripts/docx_preview.py FILE.docx --lint --text
```

`--lint` alone converts to a temporary PDF (deleted afterwards), measures every
page and reports **eight** issue codes — seven of them always, and `tail-gap`
only once `--gaps` is added. (`fill_ratio` is a per-page number, not a code.)
Read the second
output line first —
`sparse_pages` non-empty means half-empty pages, and that is work left undone,
not a warning to accept:

| JSON key | Issue code (severity) | What it means, and what to do |
|---|---|---|
| `sparse_pages` | `sparse-page` (warning) | pages whose `fill_ratio` is under `--min-fill` (default `0.40`). **Non-empty = must be dealt with**: move a block, shrink a figure, remove a needless page break |
| `blank_pages` | `blank-page` (error) | a page with neither text nor an image |
| `table_split` | `table-split` (warning) | a table continues onto the next page. Heuristic, may over-report — read it as "look at this page" |
| `orphan_last` | `orphan-last` (warning) | the page's last body line holds 1–2 characters (`--max-orphan-chars`) |
| `image_over_width` | `image-over-width` (error) | an image wider than the text block (`content_width_mm`) |
| `issues` | `content-width-unknown` (warning) | the document reports no section page width, so image widths **were not checked at all**. A degradation signal, not a defect: `image_over_width` being empty here means "not measured", not "clean" |
| `issues` | `section-margins-differ` (warning) | sections disagree on margins while the geometry is taken from section 1 only, so `head_gap_mm` / `tail_gap_mm` on later sections are measured against the wrong text-block edges |
| `gap_pages` (needs `--gaps`) | `tail-gap` (warning) | `tail_gap_mm` over `--max-tail-gap` (default `60`) |

**`exit 0` does not mean the layout is fine.** Lint findings are *output*, not
failure: `main()` prints the report and returns `0` whatever the severities were.
Measured on a one-page file whose last line held a lone `脚`: the run printed
`warning orphan-last page 1 ends with a lone '脚' on its last line` and still
exited `0`. So grep the report — `sparse_pages`, `blank_pages`, `issues` — and
never gate on the exit code. This is the same trap `docx_validate` documents from
the other side: there, `1` *is* the answer.

`fill_ratio` is a **ratio**; `--gaps` adds **absolute lengths** in millimetres.
The two answer different questions — a page can score `0.213` and still have a
small tail gap if its text starts low on the page. When the report is for a
human, quote the millimetres: `tail_gap_mm` and `head_gap_mm` are what a ruler
would say ("8.3 cm of white"), and `top_margin_mm + head_gap_mm` is checkable
against the rendered first line.

### 2. Before saying "the document has no X", cross-check with a second reader

`python-docx` walks one tree, and it silently skips the others. `paragraph.runs`
covers **direct runs only**: `w:txbxContent` (text-box content) and
`mc:AlternateContent` (compatibility wrapper) are not on that path and are
dropped without an error. Measured on the same annotation file: **43 comments,
every one inside a floating box**, and a `paragraph.runs` sweep read **0** of
them — which was then reported as "43 coloured boxes, no annotation text". The
core deliverable got diagnosed as the main defect and the fix advice was built on
that. It was wrong advice, produced by a correct-looking loop.

Cross-read the three modes instead:

```bash
python scripts/docx_read.py FILE.docx --blocks   # ordered body stream + textbox blocks
python scripts/docx_read.py FILE.docx --text     # body text + a separate textboxes list
python scripts/docx_read.py FILE.docx --runs     # per-paragraph runs, textbox: <n>
```

All three now reach floating boxes — `--text` carries a `textboxes` key,
`--blocks` emits a `textbox` block, `--runs` tags each box paragraph with
`textbox: <n>`. The mechanics, the `mc:Choice` / `mc:Fallback` double-count trap
and the shared box ordinal are in [Floating text boxes](#floating-text-boxes);
read that before writing your own traversal.

The rule is one line: **"I swept it and saw nothing" is not "the document does
not contain it" — sweep it with a different tool.** A count of zero from one
reader is a measurement of that reader, not of the file.

### 3. Validate after every save

```bash
python scripts/docx_validate.py FILE.docx
```

`docx_validate` answers the structural question lint cannot: is this package
sound? It catches an **empty `<w:drawing/>`** (Word refuses to open the file),
**invalid nesting** of the content parts, dangling relationships
(`dangling-rel`), an `r:embed` / `r:id` with no relationship
(`unresolved-reference`), a style reference with no style definition
(`missing-style`), a run property carrying a value Word will not accept
(`invalid-run-property` / `nonstandard-run-property`, below), and the footer
checks: a floating footer box too narrow to hold its line
(`footer-textbox-overflow` — measured on a 4.56 pt box asked to carry 121 pt of
text), a footer with no PAGE field (`footer-no-page-field`), an unreadable footer
part.

### `invalid-run-property` and `nonstandard-run-property`

`<w:rPr>` children whose `w:val` comes from a closed value list (`w:highlight`,
`w:vertAlign`, `w:effect`, `w:u`, `w:color`) are checked against what **Word**
does with the value, not against what the schema allows. Two codes come out of
that split, and the difference is whether the file opens:

| Code | Severity | Meaning |
|---|---|---|
| `invalid-run-property` | **error** (exit `1`) | the value makes Word **refuse to open the whole package** — `文件可能已经损坏`, COM `Documents.Open` → `com_error` `0x800A1529`. Every relationship in such a file still resolves and python-docx still opens it, so no other check in the tool can see it |
| `nonstandard-run-property` | **warning** (exit `0`) | the value is outside the schema but Word **opens the file and re-interprets** it. Real, worth naming, not a rejection |

Measured counter-examples, one per direction:

- `w:highlight="purple"` → **the file does not open**. `invalid-run-property`.
  The same applies to `DARK_MAGENTA`, `Magenta`, `FF0000`, `darkMagento`,
  `vertAlign="SUPERSCRIPT"`, `vertAlign="Baseline"`, `effect="blink"`, and to
  `<w:highlight/>` / `<w:vertAlign/>` / `<w:effect/>` with no `w:val` at all.
- `w:u="SINGLE"` → **the file opens**, Word just re-reads it as `single`.
  `nonstandard-run-property`. `w:color="#FF0000"`, `w:color="F00"`,
  `w:color="red"`, `w:color="ZZZZZZ"` and `w:color=""` behave the same way.

**Why two tiers and not one.** This report is consumed downstream by
`docx_edit`, `docx_template` and `docx_mathify`, and `ok: false` makes them
refuse the document. Calling `w:color="red"` an error therefore rejects a file
that Word opens fine — the validator becomes the defect. The split is the whole
design point: **severity is decided by Word's measured behaviour, never by the
schema**, because schema and Word disagree in *both* directions:

| Value | In the schema? | Word opens? | Verdict |
|---|---|---|---|
| `vertAlign="normal"` | **no** | **yes** (baseline, `Superscript=0 Subscript=0`, dropped on save) | accepted — no issue at all |
| `effect="ants"` | **yes** | **no** | `invalid-run-property` |

`normal` is also `docx_crossref.py`'s default spelling of "no offset", i.e. the
default value of `--vert-align`: rejecting it would reject the repository's own
output. `ants` is the opposite trap — a schema-legal value listed as legal while
Word refuses the package, which is the one outcome a validator must never
produce. Both were measured with Word 16.0 COM, one freshly built document per
value.

`missing-style` stays an **error** even though Word opens the file, and the
distinction is worth knowing: Word does not merely re-read an unknown style id,
it *substitutes* a different style on save (measured: `NoSuchStyleXyz` came back
as `aa`), so the page is silently re-formatted. That is content loss, not a
spelling difference.

It is a **structure** check, not a layout check. A file can come back
`{"ok": true, "issues": []}` and still have three half-empty pages and an image
hanging past the right margin — that verdict belongs to `--lint`. Run both; they
do not overlap.

Read the JSON, not the exit code: `ok: false` with `issues` is this tool's normal
answer and exits `1`, while a bad input path exits `2` with one `error:` line on
stderr and empty stdout. That asymmetry is deliberate and is spelled out under
[Exit codes](#exit-codes).

### 4. Never rewrite text through a shell string replacement

`-replace` in PowerShell, and `.Replace()` on any .NET string, **return the input
unchanged when nothing matches** and report nothing. A script that logs
`saved: …` after the call proves only that the call returned. On one file this
failed silently **at least six times**; it was noticed because the log line read
wrong, not because anything raised. Multi-line text, CJK, and indented blocks
break the match most easily — an invisible difference of whitespace or a soft
line break is enough.

Use a precise single-point replacement:

```bash
python scripts/docx_edit.py replace FILE.docx --find "OLD EXACT TEXT" --replace "NEW" -o OUT.docx
```

Then read the count. `docx_edit.py` reports `{"replacements": n, …}` and,
measured, a `--find` that matches nothing still exits `0`, still writes the
output file, and still claims `"ok": true` — with `"replacements": 0`. The count
is the only signal the replacement happened, so **read the text back**: `--text`
or `--runs` on the output, and confirm the old string is gone and the new one is
present. An editor's exact-match edit or an agent's `edit` tool is equally fine;
a regex over the whole file is not.

### The sequence, in one block

| After you… | Run |
|---|---|
| saved the document (any writer) | `docx_preview.py FILE.docx --lint --text` — `sparse_pages` empty, `issues` reviewed |
| suspect something is missing | `docx_read.py` with all three of `--blocks` / `--text` / `--runs` |
| changed structure (tables, images, footers, cross-references) | `docx_validate.py FILE.docx` — `ok: true` |
| replaced or inserted text | `docx_read.py --text` and read the passage back |
| are about to hand the file over | all of the above, plus one look at the page PNGs (`--pages DIR --dpi 150`) |

## Rendering and linting the layout

`docx_preview.py` answers the question the rest of the toolkit cannot: **does
page 4 look right?** It renders the document through a headless LibreOffice,
writes a PDF and one PNG per page, and reports what it sees about the page fill.
No Word, no COM, no Office licence:

> **LibreOffice is not Word at delivery time.** The same `.docx` breaks lines at
> different words, fills a page differently, and can even change page count
> between the two renderers. Measured when this gap was reported: one delivered
> document came out **6 pages** through `docx_preview.py` and **5 pages** through
> `pdf_export.ps1`. Measured again while writing this line on a fresh sample:
> page count agreed at 3, but page 1 held **36 text lines under LibreOffice
> against 31 under Word** — the breaks moved. **Judge the delivered layout from
> the `pdf_export.ps1` PDF** — page breaks, page count, fill — and use preview to
> look at layout fast and to run the structural lint, which measures geometry
> instead of trusting a renderer.

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
| `--gaps` | also report absolute white space in mm: `text_bottom_mm`, `tail_gap_mm`, `head_gap_mm`, `top_margin_mm`, plus `next_page_lead` and `gap_pages` |
| `--max-tail-gap MM` | a page whose body ink stops more than this above the text block's bottom is a `gap_page` (default `60`) |
| `--skip-last-page-gap` | do **not** report the final page's tail white. Off by default: whether an ending page counts as "half empty" is a judgement, not a measurement |
| `--no-header-band` | count running headers as body ink (the pre-P18 reading). Only for reproducing an old report |
| `--timeout SECONDS` | soffice conversion timeout (default `180`) |
| `--json` / `--text` | one JSON object, or human-readable lines (the default) |

The three capabilities compose: pass all of them and one conversion produces
everything. At least one is required. The `--lint` report carries **eight issue
codes** — seven always, plus `tail-gap` once `--gaps` is added — alongside a
per-page `fill_ratio` number. The codes: `sparse_pages`, `blank_pages`,
`image_over_width` (an image wider than the text block), `table_split` (a table
that continues onto the next page), `orphan_last` (a single character alone on
the last line), `content-width-unknown` (no page width in the document, so image
widths were not checked at all), `section-margins-differ` (sections disagree
on margins, so geometry from section 1 stands in for the rest), and `tail-gap`.
`table_split` is
a **heuristic and may over-report**: it matches
table geometry across a page boundary, so a table that merely ends near a page
break can be flagged. Read it as "look at this page", not "this is broken".
`content-width-unknown` and `section-margins-differ` are the two **degradation**
codes: they do not describe the layout, they say a measurement was skipped or
approximated. A clean report from a document that emitted either one is a weaker
statement than the same report without them.

### What counts as body ink (and why the header is not)

`--gaps` answers "how much blank paper is at the bottom of this page" — the number
a person measures with a ruler when they say *8.3 cm of white*. Two bands are
removed first, or the answer is meaningless:

- **footer**: a block starting below `0.88 × page height` is a page number, not body;
- **header**: a block ending above the **text-block top** is a running header.

The header cut is anchored on the document's own top margin, not on a fixed
fraction of the page height. That distinction is measured, not stylistic: on A4
a 12 % band ends at 101 pt, while a legitimate first body line above a 36 pt top
margin ends at 92 pt — a fraction-based cut silently deletes the first line of
every tightly-margined document. Pinned by
`tests/test_preview_header_band.py::test_a_body_line_near_the_top_is_not_mistaken_for_a_header`.

The numbers reconcile, and that is how you check them rather than trust them:

```
top_margin_mm + head_gap_mm == y of the first body line, in mm
```

On `source-snapshot/test_unpack/test_input.docx`: `24.99 + 9.56 = 34.55 mm`,
which is exactly where the body's first line starts (97.9 pt). Before the header
cut, that same page reported `head_gap_mm = −9.76` — the running header
(`TEST HEADER`, ending 12 pt above the text-block top) was being counted as body
ink.

When no top margin is available the header cut is skipped rather than guessed.
The asymmetry is deliberate: an escaped header shows up as a *negative*
`head_gap_mm` a reader will notice, while a body line deleted by a guessed cut is
invisible in the report.


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
**Your own script needs that bootstrap once** — a bare `from math_omml import …`
outside `scripts/` is a `ModuleNotFoundError`, not a typo; see
[Importing the scripts from your own code](#importing-the-scripts-from-your-own-code).

| Module | What it provides |
|---|---|
| `linear_to_latex.py` | `find_math_spans` / `atoms_to_linear` / `linear_to_latex` / `repair_merged_scripts`, plus the `UNITS` whitelist |
| `math_omml.py` | `latex_to_omml_safe` (convert + phasor-dot repair + self-check — the entry point to use) / `latex_to_omml_lxml` / `latex_to_omml_string` / `assert_legal` / `apply_math_font` / `find_xslt` |
| `fix_phasor_dot.py` | `fix_phasor_dots(omml) -> int`; `python scripts/fix_phasor_dot.py` runs its self-check |
| `docx_pack.py` | `unpack` / `pack` — the round-trip described under [Working on an unpacked document](#working-on-an-unpacked-document) |
| `docx_utils.py` | `wtag` / `register_all` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`, `FONT_HEITI`, `FONT_SONG` |
| `encoding.py` | `_fix_stdio()` / `harden_stream()` — UTF-8 stdio on Windows, with a verification step and a `TextIOWrapper` fallback; `STDIO_MODE` / `STDIO_MODES` report the branch that ran |
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
| `references/editing-traps.md` | editing an existing document and the result came out silently wrong: `cell.text =` flattening a multi-run cell, a stray `vertAlign` subscripts a whole line, `inline_shapes` index drift, `empty-drawing` that Word refuses to open, `PermissionError` from an open WPS/Word file, cp936 mojibake in console output |
| `references/toc.md` | building a static TOC: the two-pass recipe (`--page-probe` first, then write entries), tab stops and dot leaders, the full trap list, the five page-number footer templates (`zh` / `custom` placeholders) |

## Layout

```
scripts/     flat, imported by bare module name — 28 .py + 1 .ps1 = 29 files (27 CLIs, 1 pure library)
references/  longer how-to notes, loaded on demand — 7 files
tools/       Word COM acceptance helper; --page-probe lives here, in word_accept.py
tests/       pytest suite — 81 .py (80 test_*.py + _omml.py helper)
```

## Where a write lands, and what protects it

Three different things are true of different scripts, and the difference matters
because it decides whether a mistake costs you the original.

**`-o` defaults to overwriting the input.** Every editing CLI takes `-o` /
`--output`, and leaving it off means "write back over the file you gave me":
`docx_edit` (all nine subcommands), `docx_section`, `docx_crossref`,
`docx_toc`, `docx_comments`, `docx_revisions`. The help text says so — `output
path (default: overwrite input)` or `(default: overwrite)` — and it is the
reason to pass `-o` explicitly whenever a human may have the file open in Word
or WPS (see [`references/editing-traps.md`](references/editing-traps.md)).

Two writers take the output as a **positional** argument with no default —
`docx_create` (`spec.json out.docx`) and `merge_docs` (`out.docx`) — because a
document that overwrites something by omission is never what you wanted. Both
serialise straight to that path; neither keeps a copy of anything it replaced.

**Rewriting is atomic or backed up — and the two are not the same guarantee.**

| Script | Mechanism | Leaves a `.bak`? |
|---|---|---|
| `docx_edit`, `docx_section`, `docx_crossref`, `docx_toc`, `docx_comments`, `docx_revisions` | `doc.save(out)` directly — python-docx serialises over the target | no |
| `docx_template apply --to`, `fix_fonts`, `fix_zoom` | temp file → `shutil.move` / `Path.replace`, so an interrupted run leaves the original byte-intact | `apply --to` writes `<file>.bak`; `fix_fonts` / `fix_zoom` **do not** |
| `docx_mathify --inplace` | `doc.save()` then the original is renamed to `<name>.bak.docx` | yes, `.bak.docx` |
| `merge_docs` | builds the merged package in memory and writes it straight to the positional `out` | no |

`fix_zoom` is the odd one among the atomic writers: it takes `-o` (default
"patch in place"), so you can send the result somewhere else — but the swap is
still temp-file-then-replace, and it still keeps no copy of the input.
`fix_fonts` takes **no** output argument at all: it rewrites each file you name.
Both are crash-safe (the swap is atomic) but not mistake-safe — a wrong flag
still ends with a rewritten file and no second copy, which is why the docs say
to copy the file yourself when the original matters.

The `.bak` from `apply --to` and `--inplace` is a one-deep backup — running
twice does not produce `.bak.1`, it overwrites the earlier backup.

**An unpacked directory is a different target.** `table_width.py` and
`insert_formula.py` take an extracted directory, not a `.docx`, and edit the XML
parts in place; you repack afterwards. Passing one of them a `.docx` is a usage
error, not a silent conversion. See [Working on an unpacked
document](#working-on-an-unpacked-document).

## What `{"ok": true}` covers — and what it does not

`docx_validate` answers one question: **is this a well-formed, internally
consistent Word package?** A `{"ok": true, "issues": []}` means the parts are
present, the relationships resolve, the XML matches the schema's element order,
no `<w:drawing/>` is empty, and python-docx can open it. That is worth knowing,
and it is not the same as "this document looks right".

**It is legal XML, not healthy layout.** Two defects measured on this toolkit's
own output, both of which came back `{"ok": true, "issues": []}`:

- a table **19.05 cm wide on a 16 cm text block** — overflows the margins, Word
  draws it off the page, and nothing in the package is malformed;
- a **floating-textbox footer** whose box is 4.56 pt wide — the page number is
  clipped the moment it reaches two digits, and again every part is valid.

So a green `docx_validate` is a necessary condition, not a delivery check. Ask
the layout-aware commands separately:

| Question | Command |
|---|---|
| Do the columns add up to something the page can hold? | `table_width.py check --page-width 16` (`exit 1` on overflow) |
| Is the footer clip-safe, and does it carry a PAGE field? | `docx_footer.py check` |
| Where is the ink on each page, and how much of the page is empty? | `docx_preview.py --lint --gaps` |
| Do cross-reference anchors all resolve? | `docx_crossref.py list` / `check_crossref_anchors` |

The rule of thumb: **`docx_validate` reads the package, these read the page.**
Run `docx_validate` before handing a file over; run the layout checks before
believing it is finished.

### `docx_footer check`: the width is an estimate, in both directions

`page_number_width_pt` is not a measurement. No `hmtx` is read and no layout
engine is invoked, so the width is computed from a declared character model —
full-width 1.0 em, ASCII digits/letters 0.5 em, space 0.25 em, other printable
half-width 0.3 em — at the footer's own font size. The `estimate_model` field in
the report states which model ran, so a number can be explained rather than
trusted.

**The error goes both ways**, which matters because a reader will treat
`overflow: false` as "measured, and it fits":

| case | model says | real | effect |
|---|---|---|---|
| CJK at the box's limit | ~11 % **low** | wider | 85 pt box, `第 12 页　共 12 页` needs 90 pt → reports `ok`, clips on paper |
| Roman numerals | ~80 % **high** | narrower | 10 pt box, `iii` is 7.48 pt but is reported as an error |

So treat the verdict as **reliable at the extremes and unreliable at the
boundary**: a 4.56 pt box holding an 88 pt line is a real finding, while a box
within a few points of the estimate should be confirmed by rendering one page
(`docx_preview.py --pages`). The `usable_width_pt` (box minus insets) is
included for exactly that comparison.

One more thing the numbers do not say: whether the text box is *floating*. A
floating box does not re-flow and does not push its neighbours, so clipping there
is content loss; an inline box would wrap instead. `shape: "floating-textbox"` in
the report is what makes the difference explicit.

## Updating fields: `Fields.Update()` renumbers SEQ too

When a document comes back from Word with stale cross-references, the reflex is
`doc.Fields.Update()`. That is a **document-wide, ordered** recalculation, not a
refresh of one thing: it walks every field in body order and rewrites each
result — `SEQ` counters included. On a document whose captions were fine, a
global update can therefore *change* the numbering (with two caption families
whose counters were interleaved by hand, a full update is what swaps 表 3 and
表 4), and it also drops any field result that was hand-formatted where the
instruction does not pin the format (`\* MERGEFORMAT`; see the superscript table
above).

What to do instead, by intent:

| Goal | Use |
|---|---|
| Refresh the table of contents only | `doc.TablesOfContents(i).Update()` |
| Refresh one field | `doc.Fields(i).Update()` |
| Refresh a page-number or TOC field after a repagination | update that field, not the collection |
| Inspect without recalculating anything | `tools/word_accept.py --no-update-fields` |
| Nothing at all (fields written with cached results) | do nothing — open the document |

`word_accept.py` exposes `--no-update-fields` for exactly the last two rows: a
probe that must observe what the file says, not what Word would make of it.

## Exit codes

| Code | Meaning |
|---|---|
| `0` | success |
| `1` | expected failure — a marker that did not match, an unknown or duplicate bookmark name, "the document does not contain what you asked for", a runtime exception |
| `2` | usage error — argparse, an unknown key or unknown block type in a spec, more `--widths` values than columns, a `--pages` count that does not match the entry count, mutually exclusive flags passed together (including `width_frac` with `width_mm`), an `--index` / `--insert-at` outside the legal range (or negative), an input path that is missing, a directory, or not a Word package, an `image` block whose `path` points at nothing on disk or is not a string |
| `3` | the document was written, but at least one formula did not convert — `docx_create` (a `latex` block, listed in `formula_failures`), `docx_mathify`, and `docx_template apply --fix-formulas`, which passes mathify's code through |

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

`docx_create` reuses `3` for a `latex` block that did not convert, and it is the
one payload in the toolkit where `"ok": false` still means a file exists:
`output` is non-null there and names the document that was written. Judge it by
`output`, not by `ok`.

`docx_read` is the sharpest case of the stdout rule, because *every* one of its
success modes writes JSON: on a missing file, a directory, a non-ZIP, or a ZIP
that is not a Word package it exits `2` with one line on stderr and **nothing at
all on stdout**. Relaying the failure through stdout as JSON would make every
`json.loads(stdout)` caller read a failure as a success. "The document is real but
has a problem" is `docx_validate`'s job and exits `1` instead.

## Conventions

- Every script bootstraps its own directory onto `sys.path`, so it runs from any
  working directory — `python scripts/x.py` and `python x.py` both work. Your own
  script needs the same line once; see
  [Importing the scripts from your own code](#importing-the-scripts-from-your-own-code).
- Machine-readable scripts write JSON to stdout; human-facing ones write prose.
  Diagnostics and errors always go to stderr.
- Exit code reflects the outcome; check it before parsing.
- Files are read and written as UTF-8 explicitly. On Windows `_bootstrap` →
  `encoding._fix_stdio()` puts stdout and stderr on UTF-8 **and verifies it**:
  `reconfigure(encoding="utf-8", errors="backslashreplace")`, then a
  codec-normalised re-read of `stream.encoding` and of `errors` (a stream still
  on `replace` / `ignore` counts as not done), and only when that check fails an
  `io.TextIOWrapper(stream.buffer, …)` rebind. Non-Windows is left exactly as it
  was — PEP 538/540 already lift the C locale there. `encoding.STDIO_MODE` and
  `STDIO_MODES` say which branch ran; read them first when someone reports
  mojibake, because they separate "the switch worked" from "it never ran".
- **The bytes are UTF-8; the display is not this toolkit's call.** A reader that
  decodes those bytes as GBK/cp936 — a Windows PowerShell 5.1 pipeline, a
  `chcp 936` console, a parent `subprocess(text=True)` with no `encoding=` —
  shows mojibake from a perfectly correct stream, and the fix belongs on that
  side (`[Console]::OutputEncoding`, `encoding="utf-8"`). This skill guarantees
  the bytes, not every terminal.
- `requirements.txt` is pure ASCII on purpose — pip reads it with the system
  locale encoding, and a stray em dash breaks the install on a CJK Windows.

## Hard rules

- Formulas in reports and formal documents go through LaTeX→OMML. Only
  `H₂O`-level sub/superscripts use a `vertAlign` run instead.
- Every script runs under `python -X utf8`. On Windows `_bootstrap` already
  hardens stdio (see [Conventions](#conventions)) — do not add a second
  `reconfigure()`. If `encoding.STDIO_MODE` still reads `untouched` on Windows,
  the stream had no `.buffer` to rebind; that is a caller's stream, not this
  toolkit's.
- Back up before an in-place rewrite. `docx_mathify --inplace` writes
  `<name>.bak.docx` and `docx_template.py apply --to` without `-o` writes
  `<file>.bak`; `fix_fonts.py`, `fix_zoom.py`, and `merge_docs.py` keep no copy,
  so copy the file yourself when the original matters.
- Only `docx_template.py apply -o` needs node (docx-js skeleton). When the chain
  is broken, use `apply --to` instead.
- **Never write into a template cell through `cell.text =`.** It clears the cell
  and drops every run-level property. Use `docx_edit.py set-cell` (default
  `preserve`), and read `docx_read.py --runs` first when the cell may hold more
  than one run.
- **`exit 0` is not evidence the document is right.** Read the output back —
  `docx_read.py --runs` for a filled template, `docx_validate.py` before
  handing anything over — and open it in a real renderer at least once. Four
  separate silent-damage modes got through a green run in P13; see
  `references/editing-traps.md`.