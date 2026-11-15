#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""`docx_read --runs` and `docx_edit set-cell` preserve-format behaviour.

Two user-visible promises are pinned here, both from a real editing session
where the document came back "looking fine" but silently mangled:

1. **A flat string cannot say which character is a subscript.** So the default
   ``set-cell`` keeps everything a flat string *can* keep (paragraph formatting,
   a single run's own ``w:rPr``) and refuses to invent the rest: on a cell whose
   runs disagree it strips ``vertAlign`` from the carrier run and says so on
   stderr. The old whole-cell rewrite is still reachable via ``--replace-cell``.

2. **The trap has to be visible before the write.** ``docx_read --runs`` reports
   per-run ``vert_align`` and a paragraph-level ``mixed_format`` flag, so a
   caller can see "this cell is two runs with different formatting" without
   opening the XML.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))
READ = SCRIPTS / "docx_read.py"
EDIT = SCRIPTS / "docx_edit.py"


def run_cli(script: Path, *args):
    """Run a CLI and return (rc, stdout, stderr)."""
    p = subprocess.run([sys.executable, "-X", "utf8", str(script), *args],
                       capture_output=True, text=True, encoding="utf-8",
                       errors="replace")
    return p.returncode, p.stdout, p.stderr


def _make_trap(path: Path):
    """A cell whose first run is subscript and second is not.

    This is the exact shape that turned a whole ``FS`` into subscript when the
    fill helper kept "the first run's format" — the first run *was* the
    subscript one.
    """
    d = Document()
    t = d.add_table(rows=1, cols=2)
    c = t.cell(0, 0)
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("F")
    r1.font.subscript = True
    r1.font.size = Pt(14)
    r2 = p.add_run("S")
    r2.font.size = Pt(14)
    d.save(str(path))
    return path


def _make_plain(path: Path):
    """A cell with a single, richly formatted run and one paragraph."""
    d = Document()
    t = d.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    p = c.paragraphs[0]
    r = p.add_run("plain")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.name = "KaiTi"
    d.save(str(path))
    return path


def _cell(path: Path, row=0, col=0):
    return Document(str(path)).tables[0].cell(row, col)


# --------------------------------------------------------------------------
# docx_read --runs
# --------------------------------------------------------------------------

def test_runs_reports_vert_align_per_run(tmp_path):
    src = _make_trap(tmp_path / "trap.docx")
    rc, out, err = run_cli(READ, str(src), "--runs")
    assert rc == 0, err
    data = json.loads(out)
    runs = data["paragraphs"][0]["runs"]
    assert [r["text"] for r in runs] == ["F", "S"]
    assert runs[0]["vert_align"] == "subscript"
    assert runs[1]["vert_align"] is None


def test_runs_flags_mixed_format(tmp_path):
    """The paragraph that bit the user must be flagged, or the flag is useless."""
    src = _make_trap(tmp_path / "trap.docx")
    rc, out, _ = run_cli(READ, str(src), "--runs")
    assert rc == 0
    assert json.loads(out)["paragraphs"][0]["mixed_format"] is True


def test_runs_does_not_flag_uniform_paragraph(tmp_path):
    """A single run is not a trap; crying wolf here makes the flag noise."""
    src = _make_plain(tmp_path / "plain.docx")
    rc, out, _ = run_cli(READ, str(src), "--runs")
    assert rc == 0
    para = json.loads(out)["paragraphs"][0]
    assert not para.get("mixed_format")


def test_runs_reports_size_and_font(tmp_path):
    src = _make_plain(tmp_path / "plain.docx")
    rc, out, _ = run_cli(READ, str(src), "--runs")
    assert rc == 0
    run = json.loads(out)["paragraphs"][0]["runs"][0]
    assert run["size_pt"] == 16.0
    assert run["font"] == "KaiTi"
    assert run["bold"] is True


def test_runs_does_not_change_blocks_contract(tmp_path):
    """--blocks keeps its old keys; only --runs adds the run-level ones."""
    src = _make_plain(tmp_path / "plain.docx")
    rc, out, _ = run_cli(READ, str(src), "--blocks")
    assert rc == 0
    data = json.loads(out)
    flat = json.dumps(data)
    assert "mixed_format" not in flat
    assert "vert_align" not in flat


# --------------------------------------------------------------------------
# docx_edit set-cell — preserve path
# --------------------------------------------------------------------------

def test_single_run_keeps_its_own_format(tmp_path):
    """One run is the lossless case: rPr survives, so bold/size/font stay."""
    src = _make_plain(tmp_path / "plain.docx")
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "set-cell", str(src), "--table", "0",
                          "--row", "0", "--col", "0", "--text", "hello",
                          "-o", str(out))
    assert rc == 0, err
    assert err.strip() == "", f"single-run cell must not warn: {err!r}"
    assert json.loads(so)["mode"] == "preserve"

    run = _cell(out).paragraphs[0].runs[0]
    assert run.text == "hello"
    assert run.font.bold is True
    assert run.font.size == Pt(16)
    assert run.font.name == "KaiTi"


def test_paragraph_alignment_survives(tmp_path):
    src = _make_trap(tmp_path / "trap.docx")
    out = tmp_path / "out.docx"
    rc, _, _ = run_cli(EDIT, "set-cell", str(src), "--table", "0", "--row",
                       "0", "--col", "0", "--text", "FS9", "-o", str(out))
    assert rc == 0
    assert _cell(out).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_mixed_cell_strips_vert_align_and_warns(tmp_path):
    """The regression this whole task exists for.

    Keeping the first run's format used to make the entire new string a
    subscript, because the first run *was* the subscript one.
    """
    src = _make_trap(tmp_path / "trap.docx")
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "set-cell", str(src), "--table", "0", "--row",
                          "0", "--col", "0", "--text", "FS9", "-o", str(out))
    assert rc == 0, err
    assert "warning:" in err

    run = _cell(out).paragraphs[0].runs[0]
    assert run.text == "FS9"
    assert run.font.subscript is not True, "the new text must not be subscript"
    assert run.font.size == Pt(14), "size is still carryable, so keep it"


def test_mixed_cell_warning_also_reaches_stdout(tmp_path):
    """stderr alone is easy to lose when a caller swallows it; JSON must agree."""
    src = _make_trap(tmp_path / "trap.docx")
    out = tmp_path / "out.docx"
    rc, so, _ = run_cli(EDIT, "set-cell", str(src), "--table", "0", "--row",
                        "0", "--col", "0", "--text", "FS9", "-o", str(out))
    assert rc == 0
    payload = json.loads(so)
    assert payload["warnings"]
    assert any("vertAlign" in w for w in payload["warnings"])


def test_replace_cell_keeps_the_old_whole_cell_rewrite(tmp_path):
    """--replace-cell is the documented escape hatch; it must still clear."""
    src = _make_trap(tmp_path / "trap.docx")
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "set-cell", str(src), "--table", "0",
                          "--row", "0", "--col", "0", "--text", "FS9",
                          "--replace-cell", "-o", str(out))
    assert rc == 0, err
    assert json.loads(so)["mode"] == "replace-cell"

    cell = _cell(out)
    assert len(cell.paragraphs) == 1
    run = cell.paragraphs[0].runs[0]
    assert run.text == "FS9"
    assert run.font.subscript is None
    assert run.font.size is None, "old path drops the size; that is its point"


def test_extra_paragraphs_are_reported_not_silently_dropped(tmp_path):
    """A cell that loses paragraphs must say so — silent is the whole complaint."""
    d = Document()
    t = d.add_table(rows=1, cols=1)
    c = t.cell(0, 0)
    c.paragraphs[0].add_run("first")
    c.add_paragraph("second")
    src = tmp_path / "two.docx"
    d.save(str(src))

    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "set-cell", str(src), "--table", "0", "--row",
                          "0", "--col", "0", "--text", "only", "-o", str(out))
    assert rc == 0
    assert "paragraph" in err.lower()
    assert json.loads(so)["warnings"]


# --------------------------------------------------------------------------
# refusal paths must stay refusals
# --------------------------------------------------------------------------

@pytest.mark.parametrize("args", [
    ["--table", "9", "--row", "0", "--col", "0"],
    ["--table", "0", "--row", "9", "--col", "0"],
    ["--table", "0", "--row", "0", "--col", "9"],
])
def test_out_of_range_still_exits_2_with_one_error_line(tmp_path, args):
    src = _make_plain(tmp_path / "plain.docx")
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "set-cell", str(src), *args,
                          "--text", "x", "-o", str(out))
    assert rc == 2
    lines = [ln for ln in err.strip().splitlines() if ln.strip()]
    assert len(lines) == 1 and lines[0].startswith("error:")
    assert so.strip() == ""
    assert not out.exists(), "a refused call must not write"
