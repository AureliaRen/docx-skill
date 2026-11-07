# -*- coding: utf-8 -*-
"""A spec style that asks for 黑体 must bind the *East Asian* font slot too.

``style.font.name`` only writes ``w:rFonts/@ascii`` and ``@hAnsi``.  CJK runs
in such a style then fall back to the theme's East Asian face, and Word reports
``Font.NameFarEast == '+中文正文'`` — the heading silently renders 宋体 instead of
黑体 while the XML looks perfectly reasonable.  These tests pin the fix:
``add_styles`` writes ``ascii`` / ``hAnsi`` / ``eastAsia`` / ``cs`` in one
``w:rFonts``, and keeps that element in its schema slot inside ``w:rPr``.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# CT_RPr is an xsd:sequence; these all have to come *after* w:rFonts.
AFTER_RFONTS = ("b", "bCs", "i", "iCs", "color", "sz", "szCs", "u",
                "vertAlign", "lang")


def style_el(doc, name):
    return doc.styles[name].element


def rfonts(doc, name):
    return style_el(doc, name).find(f"{W}rPr/{W}rFonts")


def test_style_binds_east_asia():
    doc = Document()
    dc.add_styles(doc, [{"name": "Head", "base": "Normal",
                         "font": "黑体", "size_pt": 16, "bold": True}])
    rf = rfonts(doc, "Head")
    assert rf is not None, "no w:rFonts written"
    assert rf.get(qn("w:eastAsia")) == "黑体"
    assert rf.get(qn("w:ascii")) == "黑体"
    assert rf.get(qn("w:hAnsi")) == "黑体"
    assert rf.get(qn("w:cs")) == "黑体"


def test_style_font_east_asia_override():
    """A latin face plus a separate CJK face is expressible."""
    doc = Document()
    dc.add_styles(doc, [{"name": "Bil", "base": "Normal",
                         "font": "Times New Roman",
                         "font_east_asia": "宋体", "size_pt": 12}])
    rf = rfonts(doc, "Bil")
    assert rf.get(qn("w:ascii")) == "Times New Roman"
    assert rf.get(qn("w:eastAsia")) == "宋体"


def test_rfonts_stays_before_later_rpr_children():
    """Schema order: rFonts must not land after b / sz / color."""
    doc = Document()
    dc.add_styles(doc, [{"name": "S", "base": "Normal",
                         "font": "宋体", "size_pt": 10.5, "bold": True,
                         "color": "1F4E79"}])
    rpr = style_el(doc, "S").find(f"{W}rPr")
    tags = [c.tag.split("}")[-1] for c in rpr]
    assert "rFonts" in tags, tags
    idx = tags.index("rFonts")
    for later in AFTER_RFONTS:
        if later in tags:
            assert tags.index(later) > idx, (tags, later)
    assert tags[0] in ("rStyle", "rFonts"), tags


def test_no_rfonts_when_no_font_given():
    """A colour-only style must not grow a bogus empty rFonts."""
    doc = Document()
    dc.add_styles(doc, [{"name": "C", "base": "Normal", "color": "FF0000"}])
    assert rfonts(doc, "C") is None


def test_existing_rfonts_is_extended_not_duplicated():
    """add_styles on a style that already has rFonts updates it in place."""
    doc = Document()
    dc.add_styles(doc, [{"name": "Twice", "base": "Normal", "font": "楷体"}])
    dc.add_styles(doc, [{"name": "Other", "base": "Normal", "font": "黑体"}])
    rpr = style_el(doc, "Twice").find(f"{W}rPr")
    found = [c for c in rpr if c.tag.split("}")[-1] == "rFonts"]
    assert len(found) == 1, "rFonts duplicated"
