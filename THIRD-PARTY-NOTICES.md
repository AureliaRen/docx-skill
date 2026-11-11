# Third-party notices

This project is released under the MIT License; the full license text is in
[`LICENSE`](LICENSE) (`Copyright (c) 2026 AureliaRen`). This file records the
provenance of everything in the repository that is not original work, so that
upstream attribution travels with the source.

Scope of this file: **every file in the repository** is accounted for below —
`scripts/` (27 files), `tools/` (2), `tests/` (47), `references/` (6),
`examples/` (2), `.github/` (1), plus the top-level documents. Nothing is
redistributed from a third party: there is no vendored library, no `.xsd` /
`.xsl` copy, and no binary blob in the tree.

## NousResearch/hermes-agent (MIT)

Files derived from [NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent),
path `skills/productivity/microsoft-office/`, released under the MIT License,
Copyright (c) 2026 Nous Research.

**Twelve files under `scripts/`:**

| File | Evidence of lineage in the file itself |
|---|---|
| `docx_create.py` | header `# MIT License. Part of the Hermes docx skill.` |
| `docx_edit.py` | same header |
| `docx_read.py` | same header |
| `docx_validate.py` | same header |
| `docx_comments.py` | same header |
| `docx_revisions.py` | same header |
| `docx_crossref.py` | same header |
| `docx_numbering.py` | same header |
| `docx_section.py` | same header |
| `docx_toc.py` | same header |
| `docx_template.py` | docstring documents the hermes CLI's positional-argument form (`fill` as `_with_default_subcommand()`), and records that upstream's `--fix-formulas` imported an `office/` module from the hermes tree that is **not** shipped here |
| `docx_common.py` | shared helper surface of the CLIs above; header reads `# MIT License. Shared helpers for the docx skill scripts.` |

Local modifications to those files are limited to: an added `sys.path`
bootstrap header (`_bootstrap`), a UTF-8 stdio import, the portability and bug
fixes recorded in the commit history, the merged behaviour taken from the other
lineage of this toolkit, and — in `docx_common.py` — an added self-describing
`--api` entry point plus a replacement-iteration budget that prevents an
infinite loop when a replacement value contains the search token. The original
MIT grant above applies to all of them unchanged.

> Two notes for a future auditor, stated rather than hidden:
>
> 1. The ten `scripts/` headers above are the author's own statement of
>    lineage and are the evidence used here. `README.zh-CN.md`'s「自研增量」
>    paragraph still lists `docx_numbering` / `docx_crossref` / `docx_toc` /
>    `docx_section` as this project's own work; that summary predates this
>    alignment and needs the same correction. Until it is fixed, treat this
>    file as authoritative for provenance and the README line as stale.
> 2. The upstream directory listing could not be fetched while this list was
>    built (the GitHub contents API answered `429 Too Many Requests`), so the
>    set below is derived from in-repo evidence rather than from the upstream
>    tree. To confirm coverage, diff `git ls-tree` of
>    `skills/productivity/microsoft-office/docx/scripts` upstream against
>    `scripts/` here.

## Original work

Everything else is original work for this project:

- **Formula chain** — `linear_to_latex.py`, `math_omml.py`, `docx_mathify.py`,
  `insert_formula.py`, `fix_phasor_dot.py`
- **Local capabilities** — `table_width.py`, `fix_zoom.py`, `fix_fonts.py`,
  `merge_docs.py`, `pdf_export.ps1`
- **Runtime plumbing** — `_bootstrap.py`, `encoding.py`, `minidom_compat.py`,
  `docx_utils.py`, `docx_pack.py`
- **Acceptance tooling** — `tools/word_accept.py`, `tools/README.md`
- **Tests** — `tests/` (47 files: test modules plus the `_omml.py` helper and
  fixture generators)
- **Reference documentation** — `references/` (6 documents: `font-scheme`,
  `formula-guide`, `merge-and-export`, `revisions-and-comments`,
  `table-geometry`, `toc`)
- **Examples** — `examples/basic.json`, `examples/formulas.json`
- **CI** — `.github/workflows/ci.yml`
- **Top-level documents** — `README.md`, `README.zh-CN.md`, `SKILL.md`,
  `SKILL.zh-CN.md`, `CHANGELOG.md`, `THIRD-PARTY-NOTICES.md`, `requirements.txt`,
  `LICENSE`. No fixture corpus is checked in: most tests build their documents
  from specs at run time, and the one external sample they look for
  (`../_smoke/`, a smoke baseline kept above the repository root) is not part of
  this tree — the affected tests skip when it is absent.

`tools/word_accept.py` is this project's own file (header:
`# MIT License. Part of the docx skill (patched): unified acceptance probe.`); it
contains no upstream code.

## Runtime components that are not redistributed

`MML2OMML.XSL` is a Microsoft Office component. It is used at runtime from a
local Office installation (or from a copy the user points `MML2OMML_XSL` at) and
is **not** redistributed here. The same applies to Microsoft Word itself, which
`pdf_export.ps1` drives through COM: the script ships no Microsoft code.

Python dependencies are installed by the user from `requirements.txt` and are
**not** vendored. Licenses as reported by the installed distributions on the
reference install (2026-11-11); the version ranges are what `requirements.txt`
declares:

| Package | Requirement | License |
|---|---|---|
| `python-docx` | `>=1.1.0` | MIT |
| `lxml` | `>=5.0.0` | BSD-3-Clause |
| `latex2mathml` | `>=3.77.0` | MIT |
| `pypdfium2` | `>=4.0` (test-only: PDF text-layer assertions) | BSD-3-Clause, Apache-2.0 |
| `defusedxml` | optional, commented out on purpose | PSF-2.0 |
| `pytest` | dev-only, commented out on purpose | MIT |

`minidom_compat.py` prefers `defusedxml.minidom` when it is installed and falls
back to the standard library `xml.dom.minidom` otherwise; no third-party code is
copied into the shim.

## Normative references (no text redistributed)

- **ISO/IEC 29500-4** (`wml.xsd`) — `docx_utils.py` records the element order of
  `CT_RPr`, `CT_SectPr` and friends from the published schema. No schema file is
  copied into this repository; the ordering is expressed as Python tuples.
- **MML2OMML.XSL** — see above: read at runtime, never shipped.
- Pandoc issue #11674 is cited in `references/formula-guide.md` as the record of
  the "single inline equation is promoted to display" behaviour. Citation only.

## Verification

Re-check that no proprietary text rode along, from the repository root:

```
git grep -nE 'Anthropic|anthropic|All rights reserved|Proprietary'
git grep -n 'Part of the Hermes docx skill'       # 10 scripts + the line in this file
python -c "import importlib.metadata as m;print(m.metadata('latex2mathml')['License'])"
```

At the time of writing the first command matches **nothing but the pattern line
inside this file** — no `Anthropic`, no `Proprietary` text anywhere in the tree
or in history — and the second matches exactly the ten `scripts/` files listed in
the table (plus this file's own mention), with the other two documented from
their docstrings.