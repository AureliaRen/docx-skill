---
name: docx
description: "Create, read, edit, validate, comment on, and rewrite Word documents (.docx/.dotx) with a set of small Python CLIs. Use whenever the task involves a Word file: generating a report/memo/letter from a spec, extracting or inspecting existing content, find-and-replace across a document, editing tables, managing comments and tracked changes, filling {{token}} templates, extracting/applying formatting templates, fixing table column widths, or converting linear formulas into native Office Math (OMML). Also use when a deliverable should be a .docx. Do NOT use for PDF, spreadsheets (XLSX/CSV), slide decks, or plain one-shot text extraction (use markitdown)."
license: MIT
---

# Word documents (.docx)

All scripts live flat in `scripts/` and are argparse CLIs with `--help`. Run them
with any Python 3.10+ that has `python-docx` and `lxml` installed.

Output convention differs by script and is deliberate:

- **JSON on stdout** — `docx_create`, `docx_read`, `docx_edit`, `docx_validate`,
  `docx_comments`, `docx_revisions`, `docx_delete`-style machine-readable tools.
  Parse these.
- **Human-readable text** — `docx_mathify`, `table_width`, `insert_formula`,
  `docx_template` (except `fill`, which emits JSON on stdout). These report
  progress and are meant to be read.

Exit codes are reliable: `0` on success, non-zero on failure. See
[Exit codes](#exit-codes) for the one historical exception.

## Setup

```bash
python -m pip install -r requirements.txt
```

`latex2mathml` is only needed for the formula features. A local Microsoft Office
install supplies `MML2OMML.XSL`, which the LaTeX→OMML converter reuses — point
`MML2OMML_XSL` at it if it lives somewhere unusual.

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
round-trip that changes nothing produces an equivalent archive. `unpack` treats
the archive as untrusted input and refuses entries that would escape the target
directory — `../`, absolute paths, backslashes — raising `UnsafeArchiveError`.

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
| Health-check a document | `python scripts/docx_validate.py in.docx` |
| List comments | `python scripts/docx_comments.py list in.docx` |
| Add a comment | `python scripts/docx_comments.py add in.docx --target "some text" --text "note"` |
| Delete a comment | `python scripts/docx_comments.py delete in.docx --id 1` |
| List / accept / reject revisions | `python scripts/docx_revisions.py list\|accept-all\|reject-all in.docx` |
| Fill `{{token}}` placeholders | `python scripts/docx_template.py tpl.docx values.json out.docx` |
| Extract a formatting template | `python scripts/docx_template.py extract tpl.docx -o tpl.json` |
| Apply a formatting template | `python scripts/docx_template.py apply tpl.json --to in.docx -o out.docx` |
| Diagnose table widths | `python scripts/table_width.py report unpacked/` |
| Force table widths | `python scripts/table_width.py set unpacked/ --widths 2.4,4.2,4.2` |
| Batch linear formulas → Office Math | `python scripts/docx_mathify.py in.docx -o out.docx --report conv.md` |
| Insert one LaTeX formula | `python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display` |
| Inspect numbering definitions | `python scripts/docx_numbering.py list in.docx` |
| Add a numbering definition | `python scripts/docx_numbering.py define in.docx --text "[%1]" -o out.docx` |
| Number paragraphs | `python scripts/docx_numbering.py apply in.docx --index 155:171 --num-id 1` |
| Restart a list mid-document | `python scripts/docx_numbering.py restart in.docx --index 7 --num-id 1` |
| Strip numbering | `python scripts/docx_numbering.py detach in.docx --index 3:9` |
| Make a numbered caption | `python scripts/docx_crossref.py caption in.docx --index 42 --label 图 --text "架构"` |
| Insert a cross-reference | `python scripts/docx_crossref.py ref in.docx --index 70 --target 图_1 --prefix "如图" --suffix 所示` |
| Refresh all field results | `python scripts/docx_crossref.py renumber in.docx -o out.docx` |
| List bookmarks / fields | `python scripts/docx_crossref.py list in.docx` |
| Write a static table of contents | `python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7` |
| Measure the displayed page of paragraphs | `python tools/word_accept.py in.docx --page-probe 1,9,43 --json out.json` |
| Unify fonts (CJK body / Latin Times New Roman) | `python scripts/fix_fonts.py in.docx [more.docx ...]` |
| Merge two documents | `python scripts/merge_docs.py base.docx src.docx out.docx` |
| Add the missing `w:percent` to `<w:zoom>` | `python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]` |
| Export to PDF with the text layer intact | `pwsh scripts/pdf_export.ps1 -DocxPath in.docx [-PdfPath out.pdf]` |

## Creating a document

`docx_create.py` takes a JSON spec. Top-level keys: `page`, `styles`, `header`,
`footer`, `footer_page_numbers`, and `blocks`. Each block has a `type`:

| `type` | Fields |
|---|---|
| `heading` | `text`, `level` (1–9) |
| `paragraph` | `text`, or `runs`, plus optional `style` |
| `bullet_list` | `items` |
| `numbered_list` | `items`; optional `num_id` (reuse an existing counter), `num_text` (default `%1.`), `num_fmt`, `start`, `ilvl`, `indent_left`, `hanging` |
| `caption` | `label` (default `图`), `text`, optional `bookmark` — inserts `图 N  题注` with a `SEQ` counter + bookmark |
| `crossref` | `target` (a bookmark), optional `prefix`, `suffix`, `format` (`default`/`full`/`above-below`/`page`/`number`), `cached`, `own_paragraph` |
| `table` | `header`, `rows`, optional `style` (default `Table Grid`), `header_bold` |
| `image` | `path`, optional `width_mm` |
| `page_break` | — |
| `toc` | optional `instr` — default `TOC \o "1-3" \h \z \u`; pass `TOC \h \c "图"` for a table of figures. **A field**: Word must update it (F9) before any entry exists |

`styles[]` entries take `font`, `font_east_asia`, `size_pt`, `bold`, `italic`,
`color`, `base`, and `outline_level` — the last one writes
`<w:pPr><w:outlineLvl>` so Word treats the style as a heading. Without it a
template's own `卷标题` / `ChapterTitle` is *body text* to Word no matter how it
looks: the navigation pane is empty and `TOC \o` collects nothing.

```json
{"styles": [
  {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
   "bold": true, "outline_level": 1}
]}
```

`outline_level` is 1-based (`1` → `w:outlineLvl val="0"`). `0` means "explicitly
body text" (`val="9"`), for a style that must never be collected.

## A static table of contents

`toc` above is a **field**: the entries do not exist until someone presses F9,
and a hand-written cache inside the field is discarded the moment they do. When
the file must be right on open — a deliverable, a PDF conversion, a diff — write
the entries as real paragraphs with `scripts/docx_toc.py` instead.

```
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

**Page numbers have to be measured, not guessed.** `pages=None` leaves the slot
empty; fill it from `word_accept.py --page-probe`, which reports
`paragraph_pages[].displayed_text` — the number the reader sees, taken from
`Information(WD_ACTIVE_END_ADJUSTED_PAGE)`. That is *not* the physical page
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

## Formula conversion (LaTeX → native Office Math)

Two entry points share one engine (`math_omml.py`):

- **`docx_mathify.py` — batch.** Scans a whole document for formulas written as
  Unicode characters plus real sub/superscript runs (`U`+sub`oc`, `√2`,
  `1kΩ=10³Ω`, `∠53.1°`), converts each into native OMML, and leaves the Chinese
  prose, table structure, and font sizes untouched. `--dry-run` reports without
  writing; `--report md` records every before/after pair, plus a count and a
  per-paragraph table for spans that took the no-layout tier.
- **`insert_formula.py` — single.** Takes LaTeX and drops one equation into an
  unpacked document at an anchor.

`linear_to_latex.py` decides *what counts as a formula*. It deliberately stays
conservative so ordinary prose survives: `[0, 1]`, `f(x, y)`, and CJK headings
are left alone; `U=IR` and `ΣI=0` become equations. Recognising a span is not
the same as needing an equation *object* — see the tiers below.

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

## Table geometry

Word resolves column widths in a way that surprises people:

- `tblLayout` absent (autofit) → each cell's `<w:tcW>` acts as a *hint*; columns
  come out proportional to it.
- `tblLayout type="fixed"` → the renderer lays out purely from
  `<w:tblGrid><w:gridCol>` and **ignores `tcW`**. A default grid makes every
  column equal no matter what you set.

`table_width.py` keeps `tblW`, `gridCol`, and `tcW` consistent. Use `report`
first to see the current state, then `sync` to copy `tcW` values into the grid,
or `set --widths` to pin them explicitly.

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
interrupted run leaves the original intact.

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
  `nothing to do` and exits `0`.

The pure entry point is `fix_settings_xml(xml, percent) -> tuple[str, int]`: XML
in, `(patched_xml, sites_changed)` out. The tests import it directly rather than
shelling out.

```
python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]
```

Three details make it safe to run anywhere:

- A `<w:zoom>` that already carries `w:percent` is returned untouched. Adding a
  second `percent` would be a duplicate attribute — malformed XML, worse than
  the original defect.
- The archive is rebuilt by reusing each original `ZipInfo`, so entry order and
  timestamps survive and the byte diff stays limited to `settings.xml`.
- A package with no `word/settings.xml` is legal. The script reports
  `nothing to do` and exits `0` rather than inventing the part.

The pure entry point is `fix_settings_xml(xml, percent) -> tuple[str, int]`: XML
in, `(xml, count)` out, no file I/O, which is what the tests import. Re-running
is a no-op.

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

## Library modules

Nine modules have no CLI and exist to be imported. Scripts reach them by bare
name, which works because every script bootstraps its own directory onto
`sys.path` (see [Conventions](#conventions)).

| Module | What it provides |
|---|---|
| `linear_to_latex.py` | `find_math_spans` / `atoms_to_linear` / `linear_to_latex` / `repair_merged_scripts`, plus the `UNITS` whitelist |
| `math_omml.py` | `latex_to_omml_lxml` / `latex_to_omml_string` / `assert_legal` / `apply_math_font` / `find_xslt` |
| `fix_phasor_dot.py` | `fix_phasor_dots(omml) -> int`; `python scripts/fix_phasor_dot.py` runs its self-check |
| `docx_pack.py` | `unpack` / `pack` — the round-trip described under [Working on an unpacked document](#working-on-an-unpacked-document) |
| `docx_utils.py` | `wtag` / `register_all` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`, `FONT_HEITI`, `FONT_SONG` |
| `docx_common.py` | `iter_all_paragraphs` / `iter_part_roots` / `replace_in_paragraph` (`--api` prints the surface) |
| `encoding.py` | `_fix_stdio()` — switches stdout and stderr to UTF-8 on Windows |
| `minidom_compat.py` | `defusedxml.minidom` with a stdlib fallback |
| `_bootstrap.py` | flat-layout bootstrap: `sys.path` plus the minidom / UTF-8 fallbacks |

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
| `references/table-geometry.md` | column widths not taking effect, deciding between `set` and `sync`, `<w:zoom>` missing `w:percent` |
| `references/font-scheme.md` | mixed CJK/Latin font conventions, `w:rPr` ordering and dedupe, what `--heading-threshold` means |
| `references/merge-and-export.md` | after merging: page area / numbering / images wrong; Word COM export details; multi-page PDF via PIL |
| `references/revisions-and-comments.md` | OOXML representation of comments and tracked changes, and the CLI semantics |
| `references/toc.md` | building a static TOC: the two-pass recipe (`--page-probe` first, then write entries), tab stops and dot leaders, the full trap list |

## Layout

```
scripts/     every CLI, flat, imported by bare module name — 26 .py + 1 .ps1
references/  longer how-to notes, loaded on demand — 6 files
tools/       Word COM helper used by docx_toc's --page-probe
tests/       pytest suite
```

## Exit codes

| Code | Meaning |
|---|---|
| `0` | success |
| `1` | expected failure (missing file, marker not found, bad spec) |
| `2` | usage error (argparse, or a required mode flag missing) |
| `3` | partial success — `docx_mathify` wrote the document but some formulas failed to convert |

`docx_mathify` returning `3` is deliberate: the output file is valid and mostly
converted, but at least one equation is still linear text. Check the exit code if
you are delivering the result.

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
