# -*- coding: utf-8 -*-
"""Answer lines are paragraph bottom borders, not blank paragraphs.

Reported by the physics agent: its "answer rules" were eight empty paragraphs
styled ``SRule``.  ``SRule`` carries only a font — **no border whatsoever** — so
the student saw blank white space where the rule should be.  There was no way to
express a rule at all: ``docx_create`` had no border vocabulary.

Word stores a paragraph rule in ``w:pPr/w:pBdr/w:bottom``; Word 16.0 reads it
back as ``Borders(-3).LineStyle == 1`` (single) with ``LineWidth == w:sz``.
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


def build(blocks, styles=None):
    doc = Document()
    if styles:
        dc.add_styles(doc, styles)
    for b in blocks:
        dc.add_block(doc, b)
    return doc


def ppr_of(para):
    return para._p.find(qn("w:pPr"))


def pbdr_of(para):
    ppr = ppr_of(para)
    return None if ppr is None else ppr.find(qn("w:pBdr"))


def test_bottom_border_true_gives_solid_rule():
    doc = build([{"type": "paragraph", "text": "", "bottom_border": True}])
    pBdr = pbdr_of(doc.paragraphs[0])
    assert pBdr is not None, "no w:pBdr written"
    bottom = pBdr.find(qn("w:bottom"))
    assert bottom is not None
    assert bottom.get(qn("w:val")) == "single"
    assert bottom.get(qn("w:sz")) == "8"        # 1pt
    assert bottom.get(qn("w:color")) == "000000"


def test_bottom_border_dict_is_honoured():
    doc = build([{"type": "paragraph", "text": "",
                  "bottom_border": {"style": "dotted", "size": 6,
                                    "space": 2, "color": "FF0000"}}])
    bottom = pbdr_of(doc.paragraphs[0]).find(qn("w:bottom"))
    assert bottom.get(qn("w:val")) == "dotted"
    assert bottom.get(qn("w:sz")) == "6"
    assert bottom.get(qn("w:space")) == "2"
    assert bottom.get(qn("w:color")) == "FF0000"


def test_bottom_border_false_removes_inherited_rule():
    doc = build([
        {"type": "paragraph", "text": "a", "bottom_border": True},
        {"type": "paragraph", "text": "b", "bottom_border": False},
    ])
    assert pbdr_of(doc.paragraphs[0]) is not None
    assert pbdr_of(doc.paragraphs[1]) is None


def test_unknown_border_style_rejected():
    doc = Document()
    para = doc.add_paragraph()
    with pytest.raises(ValueError):
        dc._apply_bottom_border(para, {"style": "squiggly"})


def test_border_sits_in_schema_order():
    """``w:pBdr`` must precede ``w:spacing``/``w:ind``/``w:jc`` in ``w:pPr``."""
    doc = build([{"type": "paragraph", "text": "",
                  "bottom_border": True, "line_spacing": 2.0,
                  "indent_left_mm": 7.5, "alignment": "left"}])
    tags = [c.tag.split("}")[-1] for c in ppr_of(doc.paragraphs[0])]
    assert "pBdr" in tags, tags
    idx = tags.index("pBdr")
    for later in ("shd", "tabs", "spacing", "ind", "jc"):
        if later in tags:
            assert tags.index(later) > idx, (tags, later)


def test_border_survives_save(tmp_path):
    doc = build([{"type": "paragraph", "text": "",
                  "bottom_border": {"style": "single", "size": 8}}],
                styles=[{"name": "Body", "base": "Normal", "font": "宋体"}])
    out = tmp_path / "b.docx"
    doc.save(out)
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert "<w:pBdr>" in xml
    assert '<w:bottom w:val="single" w:sz="8"' in xml


def test_many_rule_paragraphs_are_all_distinct():
    """An answer area is several separate rules, not one paragraph repeated."""
    doc = build([{"type": "paragraph", "text": "",
                  "bottom_border": True} for _ in range(4)])
    assert len(doc.paragraphs) == 4
    assert sum(1 for p in doc.paragraphs if pbdr_of(p) is not None) == 4


def test_style_without_border_renders_blank():
    """Guard against the original bug: a plain style has no visible rule."""
    doc = build([{"type": "paragraph", "style": "Body", "text": ""}],
                styles=[{"name": "Body", "base": "Normal", "font": "宋体"}])
    assert pbdr_of(doc.paragraphs[0]) is None
