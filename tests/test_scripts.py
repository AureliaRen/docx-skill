# -*- coding: utf-8 -*-
"""End-to-end tests for the non-formula scripts.

The formula chain has its own suite; these cover the document lifecycle, which
previously had no automated coverage -- which is how a `template.py` off-by-one
and a Python 3.11-only API survived to release.

Everything here runs on a bare runner: no Office, no Node, no fixtures on disk.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

from docx_common import replace_in_paragraph  # noqa: E402
from docx_pack import pack, unpack  # noqa: E402


@pytest.fixture()
def simple_docx(tmp_path: Path) -> Path:
    d = Document()
    d.add_heading("Title", level=1)
    d.add_paragraph("hello world")
    d.add_paragraph("a cat sat")
    t = d.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = f"r{r}c{c}"
    p = tmp_path / "simple.docx"
    d.save(p)
    return p


# ---------------------------------------------------------------------------
# replace_in_paragraph — the infinite-loop regression
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text,old,new,expected", [
    ("a cat sat", "cat", "catcat", "a catcat sat"),
    ("x x x", "x", "xx", "xx xx xx"),
    ("abc", "b", "", "ac"),
    ("hello world", "world", "there", "hello there"),
    ("no match", "zzz", "yyy", "no match"),
])
def test_replace_terminates_and_is_correct(text, old, new, expected):
    """Replacement must terminate even when `new` contains `old`.

    The original implementation looped on `while old in para.text`, so
    "cat" -> "catcat" never finished: no error, no exit, just a hung process.
    """
    d = Document()
    para = d.add_paragraph(text)
    replace_in_paragraph(para, old, new)
    assert para.text == expected


def test_replace_reports_original_occurrence_count():
    d = Document()
    para = d.add_paragraph("x x x")
    assert replace_in_paragraph(para, "x", "xx") == 3


# ---------------------------------------------------------------------------
# docx_pack round-trip
# ---------------------------------------------------------------------------
def test_unpack_pack_roundtrip(simple_docx: Path, tmp_path: Path):
    up = tmp_path / "unpacked"
    unpack(simple_docx, up)
    assert (up / "word" / "document.xml").is_file()
    assert (up / "[Content_Types].xml").is_file()

    out = tmp_path / "repacked.docx"
    pack(up, out, original=simple_docx)
    assert out.is_file()

    # Reopening proves the archive is a structurally valid .docx.
    d = Document(out)
    assert any(par.text == "hello world" for par in d.paragraphs)


def test_pack_preserves_edited_part(simple_docx: Path, tmp_path: Path):
    up = tmp_path / "unpacked"
    unpack(simple_docx, up)
    doc_xml = up / "word" / "document.xml"
    doc_xml.write_text(
        doc_xml.read_text(encoding="utf-8").replace("hello world", "edited text"),
        encoding="utf-8",
    )
    out = tmp_path / "edited.docx"
    pack(up, out, original=simple_docx)
    d = Document(out)
    assert any(par.text == "edited text" for par in d.paragraphs)


def test_unpack_rejects_path_traversal(tmp_path: Path):
    import zipfile

    from docx_pack import UnsafeArchiveError

    evil = tmp_path / "evil.docx"
    with zipfile.ZipFile(evil, "w") as zf:
        zf.writestr("../../escaped.txt", "pwned")
    with pytest.raises(UnsafeArchiveError):
        unpack(evil, tmp_path / "out")
    assert not (tmp_path.parent / "escaped.txt").exists()


# ---------------------------------------------------------------------------
# create -> read -> validate
# ---------------------------------------------------------------------------
def test_create_read_validate_cycle(tmp_path: Path):
    from docx_create import main as create_main
    from docx_read import main as read_main
    from docx_validate import main as validate_main

    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps({"blocks": [
        {"type": "heading", "text": "Report", "level": 1},
        {"type": "paragraph", "text": "plain text"},
        {"type": "bullet_list", "items": ["one", "two"]},
        {"type": "table", "header": ["A", "B"], "rows": [["1", "2"]]},
    ]}), encoding="utf-8")
    out = tmp_path / "out.docx"

    argv = sys.argv
    try:
        sys.argv = ["docx_create.py", str(spec), str(out)]
        assert create_main() == 0
    finally:
        sys.argv = argv
    assert out.is_file()

    d = Document(out)
    assert any(p.text == "plain text" for p in d.paragraphs)
    assert len(d.tables) == 1

    try:
        sys.argv = ["docx_validate.py", str(out)]
        assert validate_main() == 0
    finally:
        sys.argv = argv


# ---------------------------------------------------------------------------
# template extract / apply --to — the off-by-one regression
# ---------------------------------------------------------------------------
def test_template_extract_writes_json(simple_docx: Path, tmp_path: Path):
    from docx_template import _extract

    tpl = _extract(simple_docx)
    assert isinstance(tpl, dict)
    assert "styles" in tpl or "page" in tpl


def test_template_apply_to_existing_document(simple_docx: Path, tmp_path: Path):
    from docx_template import _apply_existing, _extract

    tpl = _extract(simple_docx)
    # Must not raise: the old code computed the repo root one level too high and
    # called a script that did not exist, so this path crashed 100% of the time.
    _apply_existing(tpl, simple_docx)
    d = Document(simple_docx)
    assert any(p.text == "hello world" for p in d.paragraphs)


# ---------------------------------------------------------------------------
# table_width report on merged cells — the false-positive regression
# ---------------------------------------------------------------------------
def test_table_width_report_handles_merged_cells(tmp_path: Path, capsys):
    from table_width import main as tw_main

    d = Document()
    t = d.add_table(rows=2, cols=3)
    for r in range(2):
        for c in range(3):
            t.cell(r, c).text = f"r{r}c{c}"
    t.cell(0, 0).merge(t.cell(0, 1))     # gridSpan=2 on the first row
    src = tmp_path / "merged.docx"
    d.save(src)

    up = tmp_path / "merged_up"
    unpack(src, up)

    argv = sys.argv
    try:
        sys.argv = ["table_width.py", "report", str(up)]
        tw_main()
    finally:
        sys.argv = argv

    out = capsys.readouterr().out
    # A merged cell's tcW is the sum of the columns it spans. Comparing it to a
    # single grid column produced a bogus "! col N" mismatch, and following that
    # advice corrupted the table.
    assert "!" not in out, f"false mismatch reported:\n{out}"
    assert "merged cell" in out
