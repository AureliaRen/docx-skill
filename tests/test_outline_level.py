# -*- coding: utf-8 -*-
"""Outline levels on *styles*, and the numbers that follow from them.

A template's own heading style (``卷标题``, ``节标题``, ``中文标题``) is plain
text to Word unless it carries ``w:outlineLvl``: the page looks right, the
navigation pane is empty, and "Insert Table of Contents" produces nothing — with
no visible symptom anywhere except a missing TOC.  So the spec vocabulary has to
be able to say "this style is level N", and Word's TOC family (``toc 1/2/3``)
has to exist with a real definition for those entries to inherit.

Three separate contracts are pinned here, all on the XML, none needing Office:

* ``styles[].outline_level`` → ``w:pPr/w:outlineLvl`` at the position
  ``CT_PPr``'s ``xsd:sequence`` requires.  Written at the wrong position the
  element is **silently discarded by Word** while the file still parses — the
  worst kind of failure, so the order is asserted, not assumed.
* ``docx_read.heading_level`` must read back exactly the level that was
  written.  Two modules, one number: a mismatch means the TOC collects a
  different outline than the one the author described.
* ``content_width_mm`` is what a right-aligned TOC tab stop has to sit at, and
  the six volumes genuinely differ (170 mm maths, 160 mm the rest) — a single
  hard-coded number misplaces one family of documents.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.enum.style import WD_STYLE_TYPE  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Mm  # noqa: E402

import docx_create as dc  # noqa: E402
import docx_read  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# The ``CT_PPr`` child order, transcribed from the schema (and identical to
# ``docx_utils.PPR_ORDER``).  ``outlineLvl`` is what these tests are about.
PPR_SEQ = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]


# ---- helpers -----------------------------------------------------------------

def style_pPr(style):
    return style.element.find(qn("w:pPr"))


def style_outline_el(style):
    pPr = style_pPr(style)
    return None if pPr is None else pPr.find(qn("w:outlineLvl"))


def style_outline_val(style):
    el = style_outline_el(style)
    return None if el is None else el.get(qn("w:val"))


def ppr_tags(style):
    pPr = style_pPr(style)
    return [] if pPr is None else [c.tag.split("}")[-1] for c in pPr]


def build_style(spec, doc=None):
    doc = doc or Document()
    dc.add_styles(doc, [dict({"name": "卷标题", "base": "Normal"}, **spec)])
    return doc, doc.styles["卷标题"]


# ---- 1. the value written ----------------------------------------------------

@pytest.mark.parametrize("level, val", [(1, "0"), (2, "1"), (3, "2"),
                                        (4, "3"), (9, "8")])
def test_outline_level_is_written_one_based_to_zero_based(level, val):
    """Spec counts from 1 (what an author says); Word counts from 0."""
    _doc, style = build_style({"outline_level": level})
    assert style_outline_val(style) == val


def test_outline_level_zero_is_explicit_body_text():
    """``0`` is *not* "no opinion": it writes Word's body-text marker ``9``."""
    _doc, style = build_style({"outline_level": 0})
    assert style_outline_val(style) == dc.OUTLINE_BODY_VAL == "9"


def test_missing_outline_level_writes_nothing():
    """Omitting the key must leave the style exactly as ``base`` defines it."""
    _doc, style = build_style({})
    assert style_outline_el(style) is None
    assert "outlineLvl" not in ppr_tags(style)


def test_outline_lvl_precedes_a_paragraph_mark_rpr():
    """``w:pPr/w:rPr`` is the last thing ``outlineLvl`` must come before.

    (A style's *font* lives in ``w:style/w:rPr``, a sibling of ``w:pPr``, so it
    never competes with ``outlineLvl``; the paragraph-mark ``rPr`` inside
    ``w:pPr`` does, and python-docx appends it at the end.)
    """
    doc = Document()
    style = doc.styles.add_style("带层级", WD_STYLE_TYPE.PARAGRAPH)
    pPr = style.element.get_or_add_pPr()
    pPr.append(OxmlElement("w:rPr"))
    dc._apply_style_outline_level(style, 2)
    tags = ppr_tags(style)
    assert tags.index("outlineLvl") < tags.index("rPr"), tags


def test_a_styled_heading_style_does_not_double_the_level():
    """A style-level outlineLvl and a Heading-N paragraph yield one level."""
    doc = Document()
    dc.add_styles(doc, [{"name": "中文标题", "base": "Normal",
                         "outline_level": 3}])
    para = doc.add_paragraph("章", style="中文标题")
    assert docx_read.heading_level(para) == 3
    heading = doc.add_heading("另一章", level=3)      # built-in path, untouched
    assert docx_read.heading_level(heading) == 3


def test_outline_level_rejects_nonsense():
    for bad in (-1, 10, "标题", 1.5, object()):
        doc = Document()
        with pytest.raises(ValueError):
            dc.add_styles(doc, [{"name": "Bad", "outline_level": bad}])


# ---- 2. the position in CT_PPr ----------------------------------------------

def test_outline_lvl_sits_after_the_paragraph_geometry():
    """``outlineLvl`` comes after ``spacing``/``ind``/``jc``, before ``rPr``.

    Written at the end of ``w:pPr`` Word drops it without a warning, so this
    is the assertion that makes the feature real rather than cosmetic.
    """
    doc, style = build_style({"outline_level": 1, "font": "黑体"})
    # Force every neighbour that lives before outlineLvl to exist as well.
    pf = style.paragraph_format
    pf.line_spacing = 1.5
    pf.left_indent = Mm(10)
    pf.alignment = 1                              # centre -> w:jc
    pPr = style.element.get_or_add_pPr()
    tabs = OxmlElement("w:tabs")
    pPr.insert_element_before(tabs, *["w:" + t for t in PPR_SEQ[
        PPR_SEQ.index("tabs") + 1:]])

    tags = ppr_tags(style)
    idx = tags.index("outlineLvl")
    before = [t for t in ("tabs", "spacing", "ind", "jc") if t in tags]
    after = [t for t in ("divId", "cnfStyle", "rPr", "sectPr") if t in tags]
    assert before, tags
    for tag in before:
        assert tags.index(tag) < idx, (tags, tag)
    for tag in after:
        assert tags.index(tag) > idx, (tags, tag)


def test_outline_lvl_position_matches_the_schema_order():
    """Every child present must be in ``PPR_SEQ`` order, end to end."""
    doc, style = build_style({"outline_level": 3, "font": "宋体",
                              "size_pt": 12})
    pf = style.paragraph_format
    pf.line_spacing = 1.5
    pf.alignment = 0
    pf.left_indent = Mm(5)
    style.element.get_or_add_pPr().get_or_add_ind()
    tags = [t for t in ppr_tags(style) if t in PPR_SEQ]
    positions = [PPR_SEQ.index(t) for t in tags]
    assert positions == sorted(positions), tags


def test_outline_lvl_stays_put_when_the_style_gains_a_font():
    """Writing fonts afterwards must not disturb the paragraph properties."""
    doc = Document()
    dc.add_styles(doc, [{"name": "带层级", "base": "Normal",
                         "outline_level": 2}])
    style = doc.styles["带层级"]
    before = ppr_tags(style)
    style.font.name = "黑体"
    assert ppr_tags(style) == before
    assert style_outline_val(style) == "1"


# ---- 3. the heading block is untouched ---------------------------------------

def test_heading_block_writes_no_outline_level_of_its_own():
    """``doc.add_heading`` already carries the level on its style.

    Stacking a second, direct ``w:outlineLvl`` on the paragraph would shadow the
    style for every consumer *and* survive a style change, so the heading path
    must stay exactly as python-docx wrote it.
    """
    doc = Document()
    made = dc.add_block(doc, {"type": "heading", "text": "第一章", "level": 2})
    para = made[0]
    pPr = para._p.find(qn("w:pPr"))
    tags = [c.tag.split("}")[-1] for c in pPr]
    assert "outlineLvl" not in tags, tags
    assert tags == ["pStyle"], tags


def test_heading_block_keeps_its_style_and_level():
    doc = Document()
    dc.add_block(doc, {"type": "heading", "text": "第一章", "level": 1})
    dc.add_block(doc, {"type": "heading", "text": "1.1 节", "level": 3})
    assert [p.style.name for p in doc.paragraphs] == ["Heading 1",
                                                     "Heading 3"]
    assert [docx_read.heading_level(p) for p in doc.paragraphs] == [1, 3]


def test_a_heading_style_that_already_exists_raises():
    """``add_style`` crashing on a name clash is the reason ``ensure_toc_styles``
    looks first; this pins that the crash is real rather than a superstition."""
    doc = Document()
    with pytest.raises(ValueError):
        dc.add_styles(doc, [{"name": "Heading 2", "base": "Normal",
                             "outline_level": 2}])


# ---- 4. symmetric read-back --------------------------------------------------

@pytest.mark.parametrize("level", [1, 2, 3])
def test_heading_level_reads_back_the_same_level(level):
    """The single judge the TOC uses must see what ``add_styles`` wrote."""
    doc, style = build_style({"outline_level": level})
    para = doc.add_paragraph("标题文字", style="卷标题")
    assert docx_read.heading_level(para) == level


def test_heading_level_returns_none_for_explicit_body_text():
    doc, _style = build_style({"outline_level": 0})
    para = doc.add_paragraph("正文", style="卷标题")
    assert docx_read.heading_level(para) is None


def test_heading_level_returns_none_without_outline_level():
    doc, _style = build_style({"font": "宋体"})
    para = doc.add_paragraph("正文", style="卷标题")
    assert docx_read.heading_level(para) is None


def test_read_back_survives_a_save_and_reopen(tmp_path):
    doc, _style = build_style({"outline_level": 2})
    doc.add_paragraph("节标题", style="卷标题")
    out = tmp_path / "outline.docx"
    doc.save(out)
    again = Document(str(out))
    assert docx_read.heading_level(again.paragraphs[0]) == 2
    with zipfile.ZipFile(out) as archive:
        styles = archive.read("word/styles.xml").decode("utf-8")
    assert '<w:outlineLvl w:val="1"/>' in styles


def test_outline_level_is_written_into_the_style_not_the_paragraph(tmp_path):
    """A paragraph inherits the level; it must not carry a copy of it."""
    doc, _style = build_style({"outline_level": 1})
    para = doc.add_paragraph("章标题", style="卷标题")
    pPr = para._p.find(qn("w:pPr"))
    tags = [] if pPr is None else [c.tag.split("}")[-1] for c in pPr]
    assert "outlineLvl" not in tags, tags


# ---- 5. toc 1/2/3 ------------------------------------------------------------

def test_ensure_toc_styles_creates_three_styles():
    doc = Document()
    report = dc.ensure_toc_styles(doc)
    assert report["styles"] == ["toc 1", "toc 2", "toc 3"]
    assert report["levels"] == [1, 2, 3]
    assert report["style_ids"] == ["toc1", "toc2", "toc3"]
    for name in report["styles"]:
        assert doc.styles[name].name == name


def test_ensure_toc_styles_uses_words_look():
    """basedOn / next / uiPriority / unhideWhenUsed / qFormat, all present."""
    doc = Document()
    dc.ensure_toc_styles(doc, levels=(1,))
    element = doc.styles["toc 1"].element
    assert element.get(qn("w:styleId")) == "toc1"
    assert element.get(qn("w:customStyle")) is None
    assert element.find(qn("w:basedOn")).get(qn("w:val")) == "Normal"
    assert element.find(qn("w:next")).get(qn("w:val")) == "Normal"
    assert element.find(qn("w:uiPriority")).get(qn("w:val")) == "39"
    assert element.find(qn("w:unhideWhenUsed")) is not None
    assert element.find(qn("w:qFormat")) is not None


def test_ensure_toc_styles_writes_the_reference_line_spacing():
    """The reference volume uses 1.5 lines: ``w:line="360" w:lineRule="auto"``."""
    doc = Document()
    dc.ensure_toc_styles(doc)
    for level in (1, 2, 3):
        pPr = style_pPr(doc.styles["toc %d" % level])
        spacing = pPr.find(qn("w:spacing"))
        assert spacing is not None
        assert spacing.get(qn("w:line")) == "360"
        assert spacing.get(qn("w:lineRule")) == "auto"


@pytest.mark.parametrize("level, chars", [(1, 0), (2, 2), (3, 4)])
def test_ensure_toc_styles_indents_per_level(level, chars):
    """Level 2 is ``leftChars=200``, level 3 ``leftChars=400`` (the truth)."""
    doc = Document()
    dc.ensure_toc_styles(doc)
    pPr = style_pPr(doc.styles["toc %d" % level])
    ind = pPr.find(qn("w:ind"))
    assert ind is not None
    assert ind.get(qn("w:leftChars")) == str(chars * 100)
    expected_twips = int(round(chars * dc.DEFAULT_CHAR_PT * 20))
    assert int(ind.get(qn("w:left"))) == expected_twips


def test_ensure_toc_styles_honours_a_custom_indent_tuple():
    doc = Document()
    dc.ensure_toc_styles(doc, indent_chars=(0, 3, 6))
    pPr2 = style_pPr(doc.styles["toc 2"])
    pPr3 = style_pPr(doc.styles["toc 3"])
    assert pPr2.find(qn("w:ind")).get(qn("w:leftChars")) == "300"
    assert pPr3.find(qn("w:ind")).get(qn("w:leftChars")) == "600"


def test_ensure_toc_styles_is_idempotent():
    """Second call must not raise (``add_style`` does) and must not duplicate."""
    doc = Document()
    first = dc.ensure_toc_styles(doc)
    second = dc.ensure_toc_styles(doc)
    assert first["styles"] == second["styles"] == ["toc 1", "toc 2", "toc 3"]
    assert second["created"] == []
    assert second["updated"] == ["toc 1", "toc 2", "toc 3"]
    matching = [s for s in doc.styles
                if str(s.name).strip().lower() in ("toc 1", "toc 2", "toc 3")]
    assert len(matching) == 3


def test_ensure_toc_styles_updates_an_existing_style_in_place():
    """A template's own ``toc 1`` is completed, not replaced."""
    doc = Document()
    existing = doc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
    existing.element.set(qn("w:styleId"), "TOC1")
    report = dc.ensure_toc_styles(doc, levels=(1,))
    assert report["created"] == []
    assert report["style_ids"] == ["toc1"]
    style = doc.styles["toc 1"]
    assert style.element is existing.element
    assert style.element.get(qn("w:styleId")) == "toc1"
    assert style.element.find(qn("w:uiPriority")).get(qn("w:val")) == "39"


def test_ensure_toc_styles_is_idempotent_with_a_pre_existing_style():
    doc = Document()
    doc.styles.add_style("toc 2", WD_STYLE_TYPE.PARAGRAPH)
    dc.ensure_toc_styles(doc)
    dc.ensure_toc_styles(doc)
    matching = [s for s in doc.styles
                if str(s.name).strip().lower() == "toc 2"]
    assert len(matching) == 1
    ind = style_pPr(matching[0]).find(qn("w:ind"))
    assert ind.get(qn("w:leftChars")) == "200"


def test_ensure_toc_styles_survives_a_save(tmp_path):
    doc = Document()
    dc.ensure_toc_styles(doc)
    doc.add_paragraph("第一章　静电场", style="toc 1")
    out = tmp_path / "tocstyles.docx"
    doc.save(out)
    with zipfile.ZipFile(out) as archive:
        styles = archive.read("word/styles.xml").decode("utf-8")
    for style_id in ("toc1", "toc2", "toc3"):
        assert 'w:styleId="%s"' % style_id in styles
    assert styles.count("<w:uiPriority w:val=\"39\"/>") >= 3
    for chars in ("0", "200", "400"):
        assert 'w:leftChars="%s"' % chars in styles
    again = Document(str(out))
    assert again.styles["toc 3"].name == "toc 3"


def test_ensure_toc_styles_rejects_bad_input():
    doc = Document()
    with pytest.raises(ValueError):
        dc.ensure_toc_styles(doc, levels=(0,))


def test_toc_styles_close_the_toc_loop():
    """Everything ``docx_toc`` needs from T1: id recognised + level readable."""
    doc = Document()
    dc.ensure_toc_styles(doc)
    para = doc.add_paragraph("第一章　静电场", style="toc 1")
    # a TOC line is *not* a heading of the document it sits in
    assert docx_read.heading_level(para) is None
    assert doc.styles["toc 1"].element.get(qn("w:styleId")).startswith("toc")


# ---- 6. content width --------------------------------------------------------

def section_with(page_mm, left_mm, right_mm):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(page_mm)
    section.left_margin = Mm(left_mm)
    section.right_margin = Mm(right_mm)
    return section


def test_content_width_mm_maths_volume_is_170():
    """math: A4 210 mm, 20/20 margins -> 170 mm text block."""
    assert dc.content_width_mm(section_with(210, 20, 20)) == pytest.approx(
        170.0, abs=0.01)


def test_content_width_mm_other_volumes_are_160():
    """physics/chemistry/biology/comprehensive/demo: 30/20 -> 160 mm."""
    for _ in range(5):
        assert dc.content_width_mm(section_with(210, 30, 20)) == pytest.approx(
            160.0, abs=0.01)


def test_content_width_mm_is_not_a_constant():
    """A narrower page must move the tab stop, or the leaders run off the page."""
    a4 = dc.content_width_mm(section_with(210, 30, 20))
    a5 = dc.content_width_mm(section_with(148, 20, 20))
    assert a5 < a4
    assert a5 == pytest.approx(108.0, abs=0.01)


def test_content_width_mm_returns_a_float():
    value = dc.content_width_mm(section_with(210, 20, 20))
    assert isinstance(value, float)


def test_content_width_mm_handles_a_missing_margin():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.left_margin = Mm(20)
    section.right_margin = None
    assert dc.content_width_mm(section) == pytest.approx(190.0, abs=0.01)


def test_content_width_mm_pairs_with_the_toc_tab_stop():
    """The 170/160 split is what the TOC's right tab stop is set from."""
    import docx_toc as toc
    doc = Document()
    doc.add_heading("第一章", level=1)
    doc.sections[0].page_width = Mm(210)
    doc.sections[0].left_margin = Mm(20)
    doc.sections[0].right_margin = Mm(20)
    report = toc.build_static_toc(doc)
    assert report["tab_pos_mm"] == pytest.approx(170.0, abs=0.01)
    assert report["tab_pos_source"] == "content_width_mm"


# ---- the position contract against real neighbours ---------------------------
#
# ``_apply_style_outline_level`` inserts with an explicit "insert before these"
# list, so its correctness depends on which children already exist.  One
# end-to-end test (above) exercises the common case -- an essentially empty
# ``w:pPr`` -- and that is not enough: a mutation that appends instead of
# inserting still passes it, because there is nothing to be out of order with.
# The cases below each pre-populate a different neighbourhood, so "before" and
# "after" are both observable (work/REVIEW-ADVERSARIAL-R6.md 第三轮 M2).

def _style_with_ppr(children):
    """A fresh document plus a style whose ``w:pPr`` holds *children* in order.

    *children* are ``CT_PPr`` child tag names (without the ``w:`` prefix),
    appended in the order given, so a test controls exactly the neighbourhood
    ``_apply_style_outline_level`` will insert into.
    """
    doc = Document()
    style = doc.styles.add_style("OutlineProbe", WD_STYLE_TYPE.PARAGRAPH)
    pPr = style.element.get_or_add_pPr()
    for tag in children:
        pPr.append(OxmlElement("w:" + tag))
    return style


def _child_tags(style):
    pPr = style.element.find(qn("w:pPr"))
    return [c.tag.split("}")[-1] for c in pPr]


def test_outline_goes_after_ind_and_jc():
    """``ind``/``jc`` precede ``outlineLvl`` in the sequence, so it goes last."""
    style = _style_with_ppr(["ind", "jc"])
    dc._apply_style_outline_level(style, 1)
    assert _child_tags(style) == ["ind", "jc", "outlineLvl"]


def test_outline_goes_after_pstyle_numpr_and_shd():
    """The run of early children does not push ``outlineLvl`` ahead of them."""
    style = _style_with_ppr(["pStyle", "numPr", "shd"])
    dc._apply_style_outline_level(style, 2)
    assert _child_tags(style) == ["pStyle", "numPr", "shd", "outlineLvl"]


def test_outline_goes_before_div_id():
    """``divId`` follows ``outlineLvl``; inserting must not append past it.

    This is the case that separates ``insert_element_before`` from ``append``:
    with ``divId`` already present, appending would put ``outlineLvl`` after it.
    """
    style = _style_with_ppr(["divId"])
    dc._apply_style_outline_level(style, 3)
    assert _child_tags(style) == ["outlineLvl", "divId"]


def test_outline_goes_before_rpr_and_sectpr():
    style = _style_with_ppr(["rPr", "sectPr"])
    dc._apply_style_outline_level(style, 1)
    assert _child_tags(style) == ["outlineLvl", "rPr", "sectPr"]


def test_appending_after_outline_breaks_the_sequence():
    """The *reverse* of the contract, pinned so nobody adds a child later.

    ``add_styles`` sets ``outline_level`` while ``w:pPr`` is still effectively
    empty, which is why the production path is safe.  A caller that appends a
    lower-order child *afterwards* silently breaks ``CT_PPr`` order and Word
    then discards the element.  This test does not assert that it is fine; it
    records the trap so the next person adding an indent in ``add_styles``
    sees why the order matters.
    """
    style = _style_with_ppr([])
    dc._apply_style_outline_level(style, 1)
    assert _child_tags(style) == ["outlineLvl"]

    # Deliberately wrong: ``ind`` belongs *before* ``outlineLvl``.
    pPr = style.element.find(qn("w:pPr"))
    pPr.append(OxmlElement("w:ind"))
    assert _child_tags(style) == ["outlineLvl", "ind"], (
        "appending a lower-order child after outlineLvl produces an order "
        "Word rejects; use insert_element_before with the same successor list")
