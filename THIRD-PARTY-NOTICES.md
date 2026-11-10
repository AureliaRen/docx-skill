# Third-party notices

This project is released under the MIT License; the full license text is in
[`LICENSE`](LICENSE). This file records the provenance of code that is not
original work, so the upstream attribution travels with the source.

## NousResearch/hermes-agent (MIT)

The following files under `scripts/` are derived from
[NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent),
path `skills/productivity/microsoft-office/`, which is released under the MIT
License, Copyright (c) 2026 Nous Research:

- `docx_create.py`
- `docx_edit.py`
- `docx_read.py`
- `docx_validate.py`
- `docx_comments.py`
- `docx_revisions.py`
- `docx_template.py`
- `docx_common.py`

Local modifications to those files are limited to an added `sys.path`
bootstrap header, a UTF-8 stdio import, and the portability/bug fixes noted in
the commit history — including, in `docx_common.py`, an added self-describing
`--api` entry point and a replacement-iteration budget that prevents an
infinite loop when a replacement value contains the search token.
The original MIT grant above applies to them unchanged.

## Original work

The remaining modules are original work for this project:

- `_bootstrap.py`, `encoding.py`, `minidom_compat.py`, `docx_utils.py`, `docx_pack.py`
- `linear_to_latex.py`, `math_omml.py`, `docx_mathify.py`, `insert_formula.py`, `fix_phasor_dot.py`
- `table_width.py`, `fix_zoom.py`, `fix_fonts.py`, `merge_docs.py`
- `docx_numbering.py`, `docx_crossref.py`, `docx_toc.py`, `docx_section.py`
- `pdf_export.ps1`

## Runtime components that are not redistributed

`MML2OMML.XSL` is a Microsoft Office component. It is used at runtime from a
local Office installation and is **not** redistributed here. The same applies
to Microsoft Word itself, which `pdf_export.ps1` drives through COM: the
script ships no Microsoft code.

`latex2mathml`, `python-docx`, and `lxml` are runtime dependencies installed
separately by the user; see `requirements.txt`. They are not vendored here.
