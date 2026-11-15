#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""`set-cell` on cells whose text is not a single run.

These are the cases a flat ``--text`` cannot express. The rule the code keeps:
preserve every layer a flat string allows, refuse to invent the rest, and never
do it quietly.

The trap this pins down is real: a template cell holding ``d`` + ``0``
(``0`` subscript) came back from a fill looking fine, with the subscript gone.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))
EDIT = SCRIPTS / "docx_edit.py"


def run_edit(*args):
    p = subprocess.run([sys.executable, "-X", "utf8", str(EDIT), "set-cell",
                        *args], capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    return p.returncode, p.stdout, p.stderr


def _dimension_cell(path: Path):
    """`d` + subscript `0` + a second paragraph — the reported shape."""
    d = Document()
    t = d.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("d")
    r1.font.size = Pt(14)
    r1.font.name = "Times New Roman"
    r2 = p.add_run("0")
    r2.font.size = Pt(14)
    r2.font.name = "Times New Roman"
    r2.font.subscript = True
    c.add_paragraph("= 5 mm")
    d.save(str(path))
    return path


def test_two_run_cell_keeps_paragraph_layer(tmp_path):
    src = _dimension_cell(tmp_path / "dim.docx")
    out = tmp_path / "out.docx"
    rc, _, err = run_edit(str(src), "--table", "0", "--row", "0", "--col",
                          "0", "--text", "d0=5", "-o", str(out))
    assert rc == 0, err
    para = Document(str(out)).tables[0].cell(0, 0).paragraphs[0]
    assert para.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert para.runs[0].font.size == Pt(14)
    assert para.runs[0].font.name == "Times New Roman"


def test_two_run_cell_drops_subscript_rather_than_spreading_it(tmp_path):
    """The failure mode: the carrier run was the subscript one.

    Reusing its rPr wholesale turns every new character into a subscript. A flat
    string cannot say which characters were meant to be, so the answer is to
    drop the attribute and say so — not to guess.
    """
    src = _dimension_cell(tmp_path / "dim.docx")
    out = tmp_path / "out.docx"
    rc, _, err = run_edit(str(src), "--table", "0", "--row", "0", "--col",
                          "0", "--text", "d0=5", "-o", str(out))
    assert rc == 0
    new_run = Document(str(out)).tables[0].cell(0, 0).paragraphs[0].runs[0]
    assert new_run.text == "d0=5"
    assert new_run.font.subscript is not True
    assert "warning:" in err


def test_source_document_is_untouched_when_output_is_separate(tmp_path):
    src = _dimension_cell(tmp_path / "dim.docx")
    before = src.read_bytes()
    out = tmp_path / "out.docx"
    run_edit(str(src), "--table", "0", "--row", "0", "--col", "0",
             "--text", "x", "-o", str(out))
    assert src.read_bytes() == before


def test_replace_cell_is_the_escape_hatch(tmp_path):
    src = _dimension_cell(tmp_path / "dim.docx")
    out = tmp_path / "out.docx"
    rc, so, _ = run_edit(str(src), "--table", "0", "--row", "0", "--col", "0",
                         "--text", "flat", "--replace-cell", "-o", str(out))
    assert rc == 0
    assert json.loads(so)["mode"] == "replace-cell"
    cell = Document(str(out)).tables[0].cell(0, 0)
    assert len(cell.paragraphs) == 1
    assert cell.paragraphs[0].text == "flat"


def test_field_run_is_not_overwritten(tmp_path):
    """A field is not text; replacing it silently deletes a computed value.

    This cell holds one plain text run plus a PAGE field. Only one run carries
    text, so the lossless path applies: that run's text is replaced and the
    field runs are left alone. The point of the assertion is that the field is
    still there afterwards — a whole-cell rewrite would have taken it with it.
    """
    d = Document()
    t = d.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    p = c.paragraphs[0]
    p.add_run("total: ")
    # a complex PAGE field, same shape docx_edit._add_field builds
    from docx.oxml.ns import qn
    for ftype, payload in (("begin", None), (None, " PAGE "),
                           ("separate", None), (None, "1"), ("end", None)):
        r = p._p.makeelement(qn("w:r"), {})
        p._p.append(r)
        if ftype is not None:
            r.append(r.makeelement(qn("w:fldChar"),
                                   {qn("w:fldCharType"): ftype}))
        elif payload == " PAGE ":
            it = r.makeelement(qn("w:instrText"), {})
            it.text = payload
            r.append(it)
        else:
            tx = r.makeelement(qn("w:t"), {})
            tx.text = payload
            r.append(tx)
    src = tmp_path / "field.docx"
    d.save(str(src))

    out = tmp_path / "out.docx"
    rc, so, _ = run_edit(str(src), "--table", "0", "--row", "0", "--col",
                         "0", "--text", "X", "-o", str(out))
    assert rc == 0
    from docx.oxml.ns import qn as _qn
    cell_out = Document(str(out)).tables[0].cell(0, 0)
    instr = "".join(it.text or "" for it in cell_out._tc.iter(_qn("w:instrText")))
    assert "PAGE" in instr, "the field must survive a text write"
    assert json.loads(so)["mode"] == "preserve"
