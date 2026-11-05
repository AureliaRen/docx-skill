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
  `template`. These report progress and are meant to be read, with JSON coming
  later if it proves necessary.

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

A few commands (`table_width.py`, `insert_formula.py`, `template.py apply --to`)
operate on an **unpacked** document — a `.docx` extracted to a directory so its
XML parts can be edited directly. Produce one with:

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
round-trip that changes nothing produces an equivalent archive.

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
| Extract a formatting template | `python scripts/template.py extract tpl.docx -o tpl.json` |
| Apply a formatting template | `python scripts/template.py apply tpl.json --to in.docx -o out.docx` |
| Diagnose table widths | `python scripts/table_width.py report unpacked/` |
| Force table widths | `python scripts/table_width.py set unpacked/ --widths 2.4,4.2,4.2` |
| Batch linear formulas → Office Math | `python scripts/docx_mathify.py in.docx -o out.docx --report conv.md` |
| Insert one LaTeX formula | `python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display` |

## Creating a document

`docx_create.py` takes a JSON spec. Top-level keys: `page`, `styles`, `header`,
`footer`, `footer_page_numbers`, and `blocks`. Each block has a `type`:

| `type` | Fields |
|---|---|
| `heading` | `text`, `level` (1–9) |
| `paragraph` | `text`, or `runs`, plus optional `style` |
| `bullet_list` | `items` |
| `numbered_list` | `items` |
| `table` | `header`, `rows`, optional `style` (default `Table Grid`), `header_bold` |
| `image` | `path`, optional `width_mm` |
| `page_break` | — |
| `toc` | — (inserts a TOC field; update fields in Word to populate) |

To control column widths, run `table_width.py` on the unpacked document — a
`table` block's own width is not configurable at creation time.

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
    {"type": "table",  "header": ["Item", "Value"],
     "rows": [["Voltage", "220 V"]]}
  ]
}
```

## Formula conversion (LaTeX → native Office Math)

Two entry points share one engine (`math_omml.py`):

- **`docx_mathify.py` — batch.** Scans a whole document for formulas written as
  Unicode characters plus real sub/superscript runs (`U`+sub`oc`, `√2`,
  `1kΩ=10³Ω`, `∠53.1°`), converts each into native OMML, and leaves the Chinese
  prose, table structure, and font sizes untouched. `--dry-run` reports without
  writing; `--report md` records every before/after pair.
- **`insert_formula.py` — single.** Takes LaTeX and drops one equation into an
  unpacked document at an anchor.

`linear_to_latex.py` decides *what counts as a formula*. It deliberately stays
conservative so ordinary prose survives: `[0, 1]`, `f(x, y)`, and CJK headings
are left alone; `U=IR`, `ΣI=0`, `f=50Hz` are converted.

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

## Layout

```
scripts/     every CLI, flat, imported by bare module name
references/  longer how-to notes, loaded on demand
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
  on Windows (`encoding_fix.py`).
- `requirements.txt` is pure ASCII on purpose — pip reads it with the system
  locale encoding, and a stray em dash breaks the install on a CJK Windows.
