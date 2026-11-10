# docx

Create, read, edit, validate, and rewrite Word documents from the command line.

A skill for agents and humans alike: small, single-purpose Python CLIs that each
do one thing to a `.docx` and print JSON. No Office installation required for the
core operations — only `python-docx` and `lxml`.

> Chinese version: [README.zh-CN.md](README.zh-CN.md) — a mirror kept at the same
> level of detail. English is the published edition; script names, flags, and XML
> element names stay untranslated in both.

## Why another docx toolkit

Most document automation either hides the XML completely (and then can't express
what you need) or drops you into raw OOXML with no guard rails. This one sits in
between: each script is a narrow, testable operation with an argparse surface
you can call from a shell, a Makefile, or an agent.

Two things it does that are hard to find elsewhere:

- **Native Office Math from LaTeX.** Converts equations into real OMML —
  editable in Word, not an image and not a text approximation. Works in bulk
  across a whole document, or one equation at a time.
- **Table widths that actually stick.** Explains and fixes the
  `tblLayout` / `gridCol` / `tcW` interaction that makes column widths appear to
  be ignored.

And two more, added after an audit against a real 20-page thesis:

- **Real Word auto-numbering.** Writes `<w:abstractNum>` / `<w:num>` / `<w:numPr>`
  instead of applying the numeral-less `List Number` style, so `[1] … [14]`
  actually renders.
- **Cross-references that survive.** `SEQ` caption counters plus `REF` pointers
  with cached results, so `如图3所示` is correct the moment the file opens — and
  `renumber` puts it right again after an edit shifts the numbering.

### Beyond the basics

The hard cases above are not the only ones that usually end with "so I opened Word
and clicked". Six more capabilities cover the rest — five repairs, plus one export
path:

- **`fix_fonts`** — font normalisation that survives CJK mixing. Two rules on
  `w:rPr` are non-negotiable: each of its child elements may appear at most once,
  and `w:rFonts` must be the first child. East Asian faces bind to the `eastAsia`
  slot, not `ascii`, or Word substitutes without telling you.
- **`merge_docs`** — appending one document to another, with three rulings worth
  knowing before you inspect the result: page setup comes from `base` and `src`'s
  `sectPr` is dropped; `numbering` is unioned by `w:numId`; `styles` are unioned
  by `w:styleId` and a style `base` already has is never overwritten.
- **`fix_zoom`** — python-docx writes `<w:zoom w:val="bestFit"/>` with no
  `w:percent`; a strict XSD validator rejects that and Word's zoom control opens
  blank. `docx_create` already patches its own output, so the CLI is for
  documents from elsewhere. One pass rewrites it.
- **`minidom_compat` + `_bootstrap`** — `defusedxml` hardens XML parsing but is
  absent from most bundled interpreters, and then every script dies at *import*
  time. `minidom_compat` supplies a stdlib stand-in; `_bootstrap` is the single
  line that wires `sys.path`, that fallback, and UTF-8 stdio together.
- **`encoding`** — switches stdout and stderr to UTF-8 on Windows, where the
  default console codec mangles CJK output into unreadable bytes.

One more is an **additional capability** rather than a core operation.
`pdf_export.ps1` gives a document this toolkit produced a PDF export path that
does not lose its formulas: Word's own exporter keeps the OMML text layer intact,
so the equations stay searchable and nothing degrades into an image. It is not a
general docx→PDF converter. Two prerequisites have to hold — **the document
contains OMML equations**, and **it runs on Windows with Microsoft Word
installed** — because it drives Word COM. See
[Exporting to PDF](SKILL.md#exporting-to-pdf-additional-capability).

## A table of contents that reads correctly without Word

A `TOC` field is empty until someone presses F9, and a cache pre-filled into it
is thrown away the moment they do. For a deliverable — anything that gets PDF'd,
diffed, or opened by a tool that has no Office engine — the entries have to be
ordinary paragraphs. `scripts/docx_toc.py` writes them that way:

```bash
python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7
```

Each line is `title<TAB>page` in Word's own `toc 1/2/3` styles, with a
right-aligned dot-leader tab stop at the text-block width. Page numbers are
*measured*, never guessed: `tools/word_accept.py --page-probe` reports the glyph
the reader sees (`Information(WD_ACTIVE_END_ADJUSTED_PAGE)`), which is not the
physical page order — a `lowerRoman` section prints `i` on page eight.

Two prerequisites it will not invent for you: the heading styles need
`outline_level` in the spec (a style that merely *looks* like a heading is body
text to Word), and the numbers come from a second pass over the finished file.
`references/toc.md` has the recipe, the trap list, and the measured limits.

## Formulas are not all the same thing

A converter that turns everything it recognises into an equation object gets the
easy cases wrong. This one sorts each recognised span into three tiers:

| What the span is | What you get | Example |
|---|---|---|
| Structural — fractions, roots, integrals, sums, matrices, delimiters | an OMML equation object | `x^2/16+y^2/9=1`, `v=s/t` |
| Scripts only, no structure | a plain run carrying `w:vertAlign`, i.e. the superscript Word itself makes with `Ctrl+Shift+=` | `a=9.8m/s2`, `10^-7` |
| Neither — one `=`, one letter, a number and maybe a unit | plain text | `s=10km`, `t=2h`, `m=5kg` |

The third tier exists because an equation object is a cost when nothing needs
two-dimensional layout: one character takes a formula editor to change, a copy
into another program reshapes, and the line box grows. The trade is that a
quantity symbol there is no longer math italic — `s` in `s=10km` comes out
upright. That is deliberate. When you want the italic, write the span with a
script or set `vert_align` in a `docx_create` spec.

Both lower tiers sit behind switches in `docx_mathify.py` — `NO_LAYOUT_NEEDED`
for the third, `MIXED_SUPERSCRIPT` as the master switch covering the second and
third. See [SKILL.md](SKILL.md#three-output-tiers).

## Requirements

| Item | Requirement |
|---|---|
| Python | **3.10 or newer** — the CI floor; the matrix runs 3.10 and 3.12 on both Linux and Windows |
| Dependencies | `python-docx`, `lxml` — `python -m pip install -r requirements.txt` |
| Tests | additionally `pypdfium2` and `pytest`; no script imports either one |
| Formula chain | **Windows** + Microsoft Office: needs `MML2OMML.XSL` (ships with Office; `math_omml.find_xslt()` probes the Windows install paths, `MML2OMML_XSL` overrides) |
| PDF export | **Windows** + Microsoft Word (COM). **Not WPS**: WPS drops the text layer of a document that carries OMML |
| Template `apply` (building a new document from a template) | Node.js with the `docx` package |

**The platform boundary, stated plainly: two capabilities do not leave Windows.**

- **The formula chain** — `docx_mathify.py`, `insert_formula.py`, and anything
  that reaches `math_omml`. The LaTeX→OMML step reuses Microsoft's
  `MML2OMML.XSL`, and `math_omml.find_xslt()` looks for it under
  `%ProgramFiles%`, `%ProgramFiles(x86)%`, `%LOCALAPPDATA%` and the hard-coded
  `C:\Program Files…` roots — Windows paths only — then raises
  `FileNotFoundError` with the two ways to fix it. On Linux or macOS there is
  nothing for it to find unless you supply a copy yourself and point
  `MML2OMML_XSL` at it; the chain has only been measured on Windows.
- **PDF export** — `pdf_export.ps1` drives Word through COM, so it is Windows
  plus a local Word install by construction.

Everything else — create, read, edit, validate, comments, revisions, numbering,
cross-references, static TOC, sections, table geometry, fonts, merge, template
`extract` / `fill`, and `docx_pack` — is pure `python-docx` + `lxml` and runs
with no Office installed at all.

Legacy `.doc` (binary) files are not supported — save as `.docx` first, or use
LibreOffice externally.

## Install

**For agents** (Claude Code, Cursor, Codex, Copilot, Gemini CLI and 20+ others):

```bash
npx skills add AureliaRen/docx-skill
```

**For humans:**

```bash
git clone https://github.com/AureliaRen/docx-skill.git
cd docx-skill
python -m pip install -r requirements.txt
```

Python **3.10** or newer. `requirements.txt` carries exactly three hard
requirements — `python-docx`, `lxml`, `latex2mathml`; it also lists `pypdfium2`
and `pytest` for the test suite, but no script imports either one.

Optional extras:

| Feature | Extra requirement |
|---|---|
| Formula conversion | `latex2mathml` + a local **Windows** Microsoft Office (for `MML2OMML.XSL`) |
| Template apply (creating a doc from a template) | Node.js with the `docx` package |

The formula step reuses the stylesheet the local Office already ships instead of
re-deriving a MathML→OMML mapping. When Office lives somewhere unusual, point the
environment variable at it:

```powershell
$env:MML2OMML_XSL = 'D:\Office\root\Office16\MML2OMML.XSL'
```

`docx_mathify.py` and `insert_formula.py` share that lookup, and only the formula
features need it — every other script runs on a machine with no Office at all.

## Use

Every script runs from a shell on its own, and **18 of the 25 CLIs expose an
argparse surface** — start with `--help`:

```bash
python scripts/docx_create.py --help
python scripts/docx_read.py --help
```

A minimal round trip:

```bash
# create
cat > spec.json <<'JSON'
{"blocks": [
  {"type": "heading",   "text": "Report", "level": 1},
  {"type": "paragraph", "text": "Ohm's law: U=IR"}
]}
JSON
python scripts/docx_create.py spec.json report.docx

# inspect
python scripts/docx_read.py report.docx --structure

# convert the formula into native Office Math
python scripts/docx_mathify.py report.docx -o report-math.docx --report conv.md
```

Full command table in [SKILL.md](SKILL.md).

One rule holds for an agent consumer: the scripts that report structure write it
to **stdout as JSON** — `docx_create`, `docx_read`, `docx_edit`, `docx_validate`,
`docx_comments`, `docx_revisions`, `docx_section`, `docx_numbering`,
`docx_crossref`, `docx_toc`. The rest write human-readable text that is meant to be
read: `docx_mathify`, `table_width`, `insert_formula`, `docx_template` (`fill`
excepted — it emits JSON). Diagnostics always go to stderr, and exit codes are
reliable: `0` success, non-zero failure, and `docx_mathify` returning `3` means
the document was written but at least one formula did not convert.

## Scripts

`scripts/` holds **26 `.py` + 1 `.ps1` = 27 files**, all flat, no sub-package and
no `__init__.py`. The split below follows one definition, and `ci.yml` derives it
mechanically rather than trusting this prose — two unioned rules,
`_*`-prefixed **or** no `if __name__ == "__main__"` guard, mean *pure library*:

| Bucket | Count | Test |
|---|---|---|
| **CLI** | **25** | carries a `__main__` guard, so it can be run from a shell |
| **pure library** | **1** | `_bootstrap.py` — private, imported before argument parsing, never run on its own |

Of those 25 CLIs, **18 expose an argparse surface** (`--help` prints `usage: …`)
and are tabulated first. The other 7 have a `__main__` but no argparse: six print
a one-line `… is a library module, import it:` self-description (`docx_pack`,
`docx_utils`, `encoding`, `linear_to_latex`, `math_omml`, `minidom_compat`) and
`fix_phasor_dot` runs its self-check. They are imported by bare name, exactly
like `_bootstrap`. Full flags per command are in [SKILL.md](SKILL.md).

**CLI tools (18 `.py` with argparse)**

| File | What it does |
|---|---|
| `docx_create.py` | Builds a `.docx` from a JSON spec (`page` / `styles` / `header` / `footer` / `blocks`) |
| `docx_read.py` | Reads a document back: `--text` / `--structure` / `--styles` / `--images DIR` / `--revisions` / `--blocks` / `--md` |
| `docx_edit.py` | `replace` / `set-cell` / `insert` / `delete` / `style` / `normalize` / `toc` / `page-numbers` |
| `docx_validate.py` | Package-level health check; exits 1 when `ok:false` |
| `docx_comments.py` | Comments: `list` / `add --target` / `delete --id` (**ids start at `0`**) |
| `docx_revisions.py` | Tracked changes: `list` / `accept-all` / `reject-all` / `accept` / `reject` |
| `docx_common.py` | Shared paragraph walkers: `iter_all_paragraphs` / `iter_part_roots` / `replace_in_paragraph`; `--api` prints the exported surface |
| `docx_template.py` | `extract` / `apply` (`-o` needs node + docx-js, `--to` writes into an existing file) / `fill` (`{{token}}`) |
| `docx_numbering.py` | Real auto-numbering: `list` / `define` / `apply` / `detach` / `restart` |
| `docx_crossref.py` | Captions and cross-references: `caption` / `ref` / `bookmark` / `renumber` / `list` |
| `docx_toc.py` | Static TOC as ordinary paragraphs: `build_static_toc` / `collect_outline` / `drop_existing_toc` |
| `docx_section.py` | Sections: `list` / `break` / `pagenum` / `pagenums` / `header` / `footer` / `cols` |
| `docx_mathify.py` | Bulk-converts linear formulas to OMML (`--report` / `--tex` / `--dry-run` / `--no-merged-script-fix`) |
| `insert_formula.py` | Inserts one LaTeX formula into an unpacked directory (`--after` / `--before` / `--display` / `--dump`) |
| `table_width.py` | Table geometry: `report` / `set --widths` / `sync` / `auto --content-width` |
| `fix_fonts.py` | Font normalisation across style layer, run layer, and over-sized headings (`--heading-threshold`) |
| `fix_zoom.py` | Adds the `w:percent` that `<w:zoom>` needs to pass a strict XSD (`--percent`) |
| `merge_docs.py` | Appends one document to another, keeping styles / numbering / media (`base src out`) |

The 27th file is not a `.py`: `pdf_export.ps1` is the PowerShell export path —
Word COM, `ExportAsFixedFormat(17)`, opened `ReadOnly` (`-DocxPath` /
`-PdfPath`). See [Exporting to PDF](SKILL.md#exporting-to-pdf-additional-capability).

**Library and self-describing modules (7)** — imported by bare name; the first
six ship a one-line `__main__` self-description, `fix_phasor_dot` ships a
self-check. `_bootstrap.py` is the one pure library and is documented last.

| Module | What it provides |
|---|---|
| `docx_pack.py` | `unpack` / `pack` — round-trips a `.docx` through a directory, restoring the original entry order and per-entry compression. `table_width.py` and `insert_formula.py` operate on a directory produced by it, though neither imports it. |
| `docx_utils.py` | `wtag` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`, `FONT_HEITI`, `FONT_SONG` |
| `linear_to_latex.py` | `find_math_spans` / `linear_to_latex` / `repair_merged_scripts`, plus the `UNITS` whitelist |
| `math_omml.py` | LaTeX → MathML → OMML: `latex_to_omml_lxml` / `assert_legal` / `apply_math_font` |
| `fix_phasor_dot.py` | `m:limUpp` → `m:acc`, so a phasor dot stops rendering oversized; `python scripts/fix_phasor_dot.py` runs its self-check |
| `encoding.py` | `_fix_stdio()` — stdout and stderr to UTF-8 on Windows |
| `minidom_compat.py` | `defusedxml.minidom` with a stdlib fallback when the package is absent |

And the one file that is a **pure library**:

| Module | What it provides |
|---|---|
| `_bootstrap.py` | Flat-layout bootstrap: `sys.path` self-location plus the minidom / UTF-8 fallbacks. No `__main__`, no argparse — every other script imports it before parsing arguments. |

`docx_pack` is worth knowing by name even if you never call the CLI:

```python
from docx_pack import unpack, pack

unpack("in.docx", "unpacked/")
# ... edit the XML under unpacked/ ...
pack("unpacked/", "out.docx", original="in.docx")
```

Passing `original` is what lets `pack` reuse the source archive's entry order and
compression, so a no-op round trip produces an equivalent archive rather than a
re-zipped approximation. `unpack` refuses archive entries that would escape the
target directory (`../`, absolute paths, backslashes, and Windows drive-letter
forms such as `D:/x.txt`) and raises `UnsafeArchiveError` — a `.docx` is untrusted
input.

## Layout

```
scripts/     flat CLI modules, imported by bare name — 27 files (26 .py + 1 .ps1)
tools/       Word COM acceptance helper — word_accept.py, which owns --page-probe
references/  longer notes loaded on demand
tests/       pytest suite
examples/    sample specs
```

`tools/` holds one large helper, `word_accept.py`, which drives Word COM to accept
pending revisions and to report the page each paragraph actually lands on. **The
`--page-probe` flag lives there and only there** — `docx_toc.py` has no probing
flag of its own; you run `word_accept.py --page-probe` first and hand the measured
numbers to `docx_toc.py --pages`. Page numbers are measured, never computed. That
matters because the probe returns the glyph the reader sees, which is not the
physical page order once a section restarts its numbering.

All scripts share one import discipline: each puts its own directory on
`sys.path` before importing a sibling. That means `python scripts/x.py`,
`python x.py` from inside `scripts/`, and an absolute path from a different
directory all work — no `__init__.py` chain to get wrong.

### Reference notes

The measured detail — recipes, trap lists, schema quotes — lives in `references/`
and is loaded only when you need it:

| File | Read it when |
|---|---|
| [references/toc.md](references/toc.md) | building a static TOC: the two-pass recipe, measured page numbers, the three hard requirements on bookmarks |
| [references/formula-guide.md](references/formula-guide.md) | the three output tiers in detail, source-text traps, `MATH_FONT`, the two defects already fixed |
| [references/table-geometry.md](references/table-geometry.md) | `tblLayout` / `gridCol` / `tcW` agreeing, with measured before-and-after numbers |
| [references/revisions-and-comments.md](references/revisions-and-comments.md) | comment anchoring, and the boundary of accept/reject on tracked changes |
| [references/font-scheme.md](references/font-scheme.md) | style-layer plus run-layer font normalisation, and how the heading-size threshold is decided |
| [references/merge-and-export.md](references/merge-and-export.md) | what to watch when merging two documents and exporting to PDF |

## Testing

```bash
python -m pytest tests/ -v
```

The suite drives the argparse CLIs through `subprocess`, generates its fixtures on
the spot, and writes artifacts to a temp directory. Tests that need
`latex2mathml` **and** `MML2OMML.XSL` (`test_formula`, `test_units`,
`test_mixed_superscript`) skip themselves with an explicit reason —
`needs latex2mathml + Microsoft MML2OMML.XSL (local Office install)` — so a machine
without Office gets a **green run with a larger skip count, not a red one**.
Measured on the whole suite with `MML2OMML_XSL` pointed at a nonexistent file, i.e.
no Office anywhere: **1357 passed / 370 skipped / 0 failed**. Everything else
follows `requirements.txt`: `pypdfium2` backs the PDF text-layer assertions and
`pytest` is the runner; neither is a runtime dependency. Word COM tests skip exactly
the same way when Word is absent.

## License

MIT — see [LICENSE](LICENSE).

The formula engine reuses Microsoft's `MML2OMML.XSL` from a local Office
installation at runtime; that file is not redistributed here.

## Acknowledgments

The base CRUD, comment, revision, and template CLIs are adapted from the
clean-room MIT implementations in
[NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent)
(`skills/productivity/microsoft-office`), originally written from functional
specs against `python-docx` and friends.

Per-file provenance and the list of local changes are in
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md). That file also records what this
repository does **not** redistribute: `MML2OMML.XSL` is read from a local Microsoft
Office at runtime, `pdf_export.ps1` drives Word through COM without containing any
Microsoft code, and `latex2mathml` / `python-docx` / `lxml` are installed by the
user.

