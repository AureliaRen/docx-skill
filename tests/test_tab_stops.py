# -*- coding: utf-8 -*-
"""Tab stops + literal tabs: what a Chinese table of contents needs.

A paper's TOC line is ``title <TAB> page-number`` with dot leaders and the
number flush right.  Two capabilities are needed and both used to be missing:

* a tab stop (position + alignment + leader) written into ``w:pPr/w:tabs``;
* a *real* ``<w:tab/>`` element — python-docx's ``add_run`` puts ``\\t`` into
  ``w:t``, where Word shows nothing at all.

Word's own reading (verified in Word 16.0 + PDF export):
``TabStops(1).Position`` = 155mm, ``.Alignment`` = 2 (right),
``.Leader`` = 1 (dot), and the PDF text comes out
``第一章 ... 1`` with the leader dots in place.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def build(tmp_path, blocks, styles=None):
    doc = Document()
    if styles:
        dc.add_styles(doc, styles)
    for b in blocks:
        dc.add_block(doc, b)
    return doc


def tabs_of(para):
    tabs = para.paragraph_format.tab_stops
    return [(t.position.mm if hasattr(t.position, "mm") else t.position,
             t.alignment, t.leader) for t in tabs]


def test_right_aligned_dot_leader_stop():
    doc = Document()
    para = doc.add_paragraph()
    got = dc._apply_tab_stops(para, [{"pos_mm": 155, "align": "right",
                                      "leader": "dot"}])
    assert got == [{"pos_mm": 155.0, "align": "right", "leader": "dot"}]
    stops = tabs_of(para)
    assert len(stops) == 1
    pos, align, leader = stops[0]
    assert abs(pos - 155.0) < 0.1
    assert int(align) == 2      # WD_TAB_ALIGNMENT.RIGHT
    assert int(leader) == 1     # WD_TAB_LEADER.DOTS


def test_stops_are_written_in_ascending_position_order():
    doc = Document()
    para = doc.add_paragraph()
    dc._apply_tab_stops(para, [{"pos_mm": 155, "align": "right", "leader": "dot"},
                               {"pos_mm": 20, "align": "left"}])
    positions = [p for p, _a, _l in tabs_of(para)]
    assert positions == sorted(positions), positions


def test_tab_stops_parse_cm_and_defaults():
    doc = Document()
    para = doc.add_paragraph()
    got = dc._apply_tab_stops(para, [{"pos_cm": 15.5}])
    assert got[0]["pos_mm"] == pytest.approx(155.0)
    assert got[0]["align"] == "left"
    assert got[0]["leader"] == "none"


def test_unknown_align_and_leader_rejected():
    doc = Document()
    para = doc.add_paragraph()
    with pytest.raises(ValueError):
        dc._apply_tab_stops(para, [{"pos_mm": 100, "align": "sideways"}])
    with pytest.raises(ValueError):
        dc._apply_tab_stops(para, [{"pos_mm": 100, "leader": "zigzag"}])
    with pytest.raises(ValueError):
        dc._apply_tab_stops(para, [{"align": "right"}])


def test_literal_tab_becomes_a_tab_element(tmp_path):
    """A ``\\t`` in block text must land as ``<w:tab/>``, not inside ``w:t``."""
    doc = build(tmp_path, [
        {"type": "paragraph", "text": "第一章　静电场\t1",
         "tab_stops": [{"pos_mm": 155, "align": "right", "leader": "dot"}]},
    ])
    para = doc.paragraphs[0]
    p = para._p
    # ``w:tab`` is overloaded: inside ``w:tabs`` it *declares* a stop, inside a
    # run it *is* the character. Count only the run-level ones.
    run_tabs = [t for t in p.iter(qn("w:tab"))
                if t.getparent().tag != qn("w:tabs")]
    assert len(run_tabs) == 1, "expected exactly one run-level w:tab"
    text = "".join(t.text or "" for t in p.iter(qn("w:t")))
    assert "\t" not in text, "tab leaked into w:t"
    assert text.strip() == "第一章　静电场1"


def test_tab_in_runs_list_is_also_handled(tmp_path):
    doc = build(tmp_path, [
        {"type": "paragraph",
         "runs": [{"text": "第三章　磁场"}, {"text": "\t27"}],
         "tab_stops": [{"pos_mm": 155, "align": "right", "leader": "dot"}]},
    ])
    p = doc.paragraphs[0]._p
    assert len(list(p.iter(qn("w:tab")))) >= 1
    text = "".join(t.text or "" for t in p.iter(qn("w:t")))
    assert "\t" not in text
    assert "磁场" in text and "27" in text


def test_tabs_element_lands_inside_ppr(tmp_path):
    """``w:tabs`` is a direct child of ``w:pPr``, before ``w:ind``/``w:spacing``."""
    doc = build(tmp_path, [
        {"type": "paragraph", "text": "a\t1", "alignment": "left",
         "first_line_indent_chars": 0,
         "tab_stops": [{"pos_mm": 155, "align": "right", "leader": "dot"}]},
    ])
    ppr = doc.paragraphs[0]._p.find(qn("w:pPr"))
    tags = [c.tag.split("}")[-1] for c in ppr]
    assert "tabs" in tags, tags
    assert tags.index("tabs") < len(tags)


def test_saved_docx_keeps_tabs(tmp_path):
    doc = build(tmp_path, [
        {"type": "paragraph", "text": "第一章　静电场\t1",
         "tab_stops": [{"pos_mm": 155, "align": "right", "leader": "dot"}]},
    ], styles=[{"name": "Body", "base": "Normal", "font": "宋体"}])
    out = tmp_path / "toc.docx"
    doc.save(out)
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert 'w:leader="dot"' in xml
    assert 'w:val="right"' in xml
    assert "<w:tab/>" in xml
