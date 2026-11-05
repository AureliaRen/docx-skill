---
name: docx
description: "Create, read, edit, validate, comment on, and rewrite Word documents (.docx/.dotx) with a set of small Python CLIs. Use whenever the task involves a Word file: generating a report/memo/letter from a spec, extracting or inspecting existing content, find-and-replace across a document, editing tables, managing comments and tracked changes, filling {{token}} templates, extracting/applying formatting templates, fixing table column widths, or converting linear formulas into native Office Math (OMML). Also use when a deliverable should be a .docx. Do NOT use for PDF, spreadsheets (XLSX/CSV), slide decks, or plain one-shot text extraction (use markitdown)."
license: MIT
---

# Word documents (.docx)

All scripts live flat in `scripts/`, are argparse CLIs with `--help`, print JSON
to stdout, and exit non-zero on failure. Run them with any Python 3.10+ that has
`python-docx` and `lxml` installed.

## Setup

```bash
python -m pip install -r requirements.txt
```

`latex2mathml` is only needed for the formula features. A local Microsoft Office
install supplies `MML2OMML.XSL`, which the LaTeX→OMML converter reuses — point
`MML2OMML_XSL` at it if it lives somewhere unusual.

## Quick reference

| Task | Command |
|------|---------|
| Create from a JSON spec | `python scripts/docx_create.py spec.json out.docx` |
| Read text / outline / styles | `python scripts/docx_read.py in.docx --text` |
| Extract images | `python scripts/docx_read.py in.docx --images DIR` |
| Inspect tracked changes | `python scripts/docx_read.py in.docx --revisions` |
| Replace text | `python scripts/docx_edit.py in.docx replace --find A --replace B -o out.docx` |
| Set a table cell | `python scripts/docx_edit.py in.docx set-cell --table 0 --row 1 --col 2 --text X -o out.docx` |
| Insert a paragraph | `python scripts/docx_edit.py in.docx insert --index 3 --text "..." -o out.docx` |
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
| `table` | `header`, `rows`, optional `style` (default `Table Grid`), `header_bold`, `width_mm` |
| `image` | `path`, optional `width_mm` |
| `page_break` | — |
| `toc` | — (inserts a TOC field; update fields in Word to populate) |

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

## Conventions

- Every script bootstraps its own directory onto `sys.path`, so it runs from any
  working directory — `python scripts/x.py` and `python x.py` both work.
- Scripts write JSON to stdout and human-readable diagnostics to stderr.
- Exit code is non-zero when the operation failed; check it before parsing.
- Files are read and written as UTF-8 explicitly; stdio is reconfigured to UTF-8
  on Windows (`encoding_fix.py`).
