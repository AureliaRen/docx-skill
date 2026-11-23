# -*- coding: utf-8 -*-
"""One authority, one derivation: ``CT_RPr`` / ``CT_PPr`` successor lists.

``w:rPr`` and ``w:pPr`` are ``xsd:sequence``.  A child at the wrong index is
well-formed XML, passes ``docx_validate`` (well-formedness plus a few structure
rules), and is then **silently dropped by Word** — the page simply does not show
the property and nothing reports an error.  So "where does this element go" is
one of the few positions in this repo that must have exactly one answer.

It used to have several.  ``docx_create`` carried three hand-typed successor
lists:

===================================  ========  ==============================
site                                 anchor    what it actually served
===================================  ========  ==============================
``_set_style_fonts``                 rFonts    ``w:rPr`` — 27 of the 37 items
``_apply_style_outline_level``       outlineLvl ``w:pPr`` — exact tail, 5 items
``_apply_bottom_border``             pBdr      ``w:pPr`` — exact tail, 27 items
===================================  ========  ==============================

The first one was already *wrong by construction* (missing ``caps``,
``smallCaps``, ``strike``, ``dstrike``, ``outline``, ``shadow``, ``emboss``,
``imprint``, ``vanish``, ``webHidden``) and only agreed with the authority by
luck: all ten missing names happen to sort after ``rFonts``, so the two lists
picked the same insertion point *as long as* no document ever contained one of
them without also containing an item from the 27.  The moment ``RPR_ORDER``
gains or loses a name, none of the three lists would follow.

These tests pin the derivation, not a transcribed result:

* :func:`docx_crossref.successors` equals ``order[order.index(tag)+1:]`` with
  the ``w:`` prefix, for every name in both tables — no hard-coded清单;
* the positions the refactor produces are **byte-identical** to the positions
  the deleted inline lists produced (the refactor is pure), proved by
  regenerating the same document with the old literals monkeypatched back in;
* the one place where the old list was not merely shorter but *wrong* is
  recorded explicitly, and shown to be unreachable from ``add_styles``.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402
import docx_edit  # noqa: E402
from docx_crossref import RPR_ORDER, successors  # noqa: E402
from docx_utils import PPR_ORDER  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# The three literals that were deleted from ``docx_create``.  They are kept here
# **as the old-code oracle**: ``test_pbdr_position_matches_old_inline_list`` and
# its outline twin regenerate the same document with these substituted back in
# and demand byte-identical XML.  They are not the source of truth for anything.
OLD_RFONTS_TAIL = [
    "b", "bCs", "i", "iCs", "noProof", "snapToGrid", "color", "spacing", "w",
    "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]
OLD_OUTLINELVL_TAIL = ["divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]
OLD_PBDR_TAIL = [
    "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi",
    "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
    "rPr", "sectPr", "pPrChange",
]
# The ten names the rFonts literal was missing.  All of them sort after
# ``rFonts`` in RPR_ORDER, which is why the shorter list still worked.
RFONTS_GAP = ["caps", "smallCaps", "strike", "dstrike", "outline", "shadow",
              "emboss", "imprint", "vanish", "webHidden"]


# ---------------------------------------------------------------------------
# 1. the derivation itself — equivalence, not a transcribed checklist
# ---------------------------------------------------------------------------

def test_successors_is_the_tail_of_the_authority():
    """``successors(tag)`` == ``["w:"+n for n in RPR_ORDER[i+1:]]``, every tag."""
    assert len(RPR_ORDER) == 39
    for i, tag in enumerate(RPR_ORDER):
        assert successors(tag) == ["w:" + n for n in RPR_ORDER[i + 1:]], tag
    # The ends are the interesting ones: no successors, one successor.
    assert successors(RPR_ORDER[-1]) == []
    assert successors(RPR_ORDER[-2]) == ["w:" + RPR_ORDER[-1]]


def test_successors_agrees_with_the_ppr_authority():
    """The pPr table works the same way, including the ``prefix`` argument."""
    assert len(PPR_ORDER) == 36
    for i, tag in enumerate(PPR_ORDER):
        assert successors(tag, PPR_ORDER) == ["w:" + n
                                              for n in PPR_ORDER[i + 1:]], tag
    assert successors("pPrChange", PPR_ORDER) == []
    # Prefix is a parameter, not a hard-coded "w:".
    assert successors("outlineLvl", PPR_ORDER, prefix="")[0] == "divId"


def test_rfonts_has_37_successors_not_27():
    """The gap that motivated the whole package, pinned as a number."""
    derived = successors("rFonts")
    assert len(derived) == 37
    assert len(OLD_RFONTS_TAIL) == 27
    for name in RFONTS_GAP:
        assert "w:" + name in derived
        assert name not in OLD_RFONTS_TAIL
    # And the old list was a *subsequence* of the derived one, in order — which
    # is exactly why it agreed with the authority on every reachable document.
    positions = [derived.index("w:" + n) for n in OLD_RFONTS_TAIL]
    assert positions == sorted(positions)


def test_unknown_tag_raises_instead_of_appending():
    """A typo must not degrade into "put it at the end"."""
    with pytest.raises(ValueError):
        successors("rFontz")
    with pytest.raises(ValueError):
        successors("pBdr", RPR_ORDER)          # right name, wrong table


def test_docx_edit_helper_is_the_same_function():
    """``docx_edit._rpr_successors`` is an alias, not a second copy."""
    for tag in RPR_ORDER:
        assert docx_edit._rpr_successors(tag) == successors(tag)


# ---------------------------------------------------------------------------
# 2. behaviour unchanged — regenerate with the old literals monkeypatched back
# ---------------------------------------------------------------------------

def _old_successors(tail_map):
    """A stand-in for ``successors`` that returns the pre-refactor literals."""
    def fake(tag, order=RPR_ORDER, prefix="w:"):
        if tag in tail_map:
            return [prefix + n for n in tail_map[tag]]
        return [prefix + n for n in order[order.index(tag) + 1:]]
    return fake


def _ppr_doc_xml() -> bytes:
    """A document whose ``w:pPr`` carries the neighbours these paths care about.

    Built by hand rather than from a spec so that every child either side of the
    insertion point is present: ``pBdr`` must land after ``jc``, ``ind`` and
    ``shd``, and ``outlineLvl`` must land before ``divId``/``cnfStyle``/``rPr``/
    ``sectPr``.  Those neighbours are what make the successor list do work; a
    bare ``w:pPr`` would give the same answer for any list, including an empty
    one.
    """
    doc = Document()
    para = doc.add_paragraph("x")
    pPr = para._p.get_or_add_pPr()
    for name in ("jc", "ind", "shd", "tabs", "spacing"):
        el = OxmlElement("w:" + name)
        pPr.append(el)
    # An outline level on a *style*, the path _apply_style_outline_level serves.
    style = doc.styles.add_style("Probe", 1)
    s_pPr = style.element.get_or_add_pPr()
    for name in ("divId", "cnfStyle", "sectPr"):
        s_pPr.append(OxmlElement("w:" + name))
    dc._apply_style_outline_level(style, 2)
    # And a bottom border on the paragraph.
    dc._apply_bottom_border(para, {"sides": ["top", "bottom"], "size": 8})
    return doc.element.xml.encode("utf-8")


def test_ppr_positions_match_old_inline_lists(monkeypatch):
    """Same document, old literals substituted in: XML must be byte-identical.

    This is the pure-refactor claim for the two ``w:pPr`` sites, and unlike the
    ``rFonts`` branch it runs on genuinely reachable code.
    """
    baseline = _ppr_doc_xml()

    monkeypatch.setattr(dc, "successors",
                        _old_successors({"outlineLvl": OLD_OUTLINELVL_TAIL,
                                         "pBdr": OLD_PBDR_TAIL}))
    replayed = _ppr_doc_xml()
    assert replayed == baseline


def test_pbdr_lands_before_its_successors_not_at_the_end():
    """The property the successor list exists for, asserted on the tree."""
    doc = Document()
    para = doc.add_paragraph("x")
    pPr = para._p.get_or_add_pPr()
    pPr.append(OxmlElement("w:jc"))
    pPr.append(OxmlElement("w:rPr"))
    dc._apply_bottom_border(para, True)
    tags = [c.tag.split("}")[-1] for c in pPr]
    assert "pBdr" in tags
    assert tags.index("pBdr") < tags.index("jc")
    assert tags.index("pBdr") < tags.index("rPr")


def test_outline_lvl_lands_before_divid_not_at_the_end():
    doc = Document()
    style = doc.styles.add_style("Probe", 1)
    pPr = style.element.get_or_add_pPr()
    pPr.append(OxmlElement("w:divId"))
    pPr.append(OxmlElement("w:rPr"))
    dc._apply_style_outline_level(style, 1)
    tags = [c.tag.split("}")[-1] for c in pPr]
    assert tags.index("outlineLvl") < tags.index("divId")
    assert tags.index("outlineLvl") < tags.index("rPr")


# ---------------------------------------------------------------------------
# 3. the rFonts branch: reachability, and the one real semantic change
# ---------------------------------------------------------------------------

def _bare_style_with_rpr_children(names):
    """A style whose ``w:rPr`` starts empty, then gets *names* appended."""
    doc = Document()
    style = doc.styles.add_style("Probe", 1)
    rpr = style.element.get_or_add_rPr()
    for name in names:
        rpr.append(OxmlElement("w:" + name))
    return style, rpr


def test_set_style_fonts_inserts_before_every_old_list_member():
    """For any neighbour the *old* list knew about, the position is unchanged."""
    for neighbour in OLD_RFONTS_TAIL:
        style, rpr = _bare_style_with_rpr_children([neighbour])
        dc._set_style_fonts(style, "黑体", "宋体")
        tags = [c.tag.split("}")[-1] for c in rpr]
        assert tags[0] == "rFonts", (neighbour, tags)
        assert tags.index("rFonts") < tags.index(neighbour)


def test_set_style_fonts_inserts_before_every_successor_in_the_table():
    """Exhaustive: **each** of the 37 successors, one at a time.

    This is the test that catches drift.  Deleting a single name from the
    authority — or from the derived list — makes exactly one of these 37
    iterations insert ``rFonts`` after its neighbour instead of before it, so
    the mutation cannot pass.
    """
    derived = successors("rFonts")
    assert len(derived) == 37
    for prefixed in derived:
        name = prefixed.split(":", 1)[1]
        style, rpr = _bare_style_with_rpr_children([name])
        dc._set_style_fonts(style, "黑体", "宋体")
        tags = [c.tag.split("}")[-1] for c in rpr]
        assert tags == ["rFonts", name], (name, tags)


def test_pbdr_inserts_before_every_ppr_successor():
    """Exhaustive over ``PPR_ORDER[pBdr+1:]`` — catches a dropped pPr successor."""
    derived = successors("pBdr", PPR_ORDER)
    assert len(derived) == 27
    for prefixed in derived:
        name = prefixed.split(":", 1)[1]
        doc = Document()
        para = doc.add_paragraph("x")
        pPr = para._p.get_or_add_pPr()
        pPr.append(OxmlElement("w:" + name))
        dc._apply_bottom_border(para, True)
        tags = [c.tag.split("}")[-1] for c in pPr]
        assert tags == ["pBdr", name], (name, tags)


def test_outline_lvl_inserts_before_every_ppr_successor():
    """Exhaustive over ``PPR_ORDER[outlineLvl+1:]``."""
    derived = successors("outlineLvl", PPR_ORDER)
    assert len(derived) == 5
    for prefixed in derived:
        name = prefixed.split(":", 1)[1]
        doc = Document()
        style = doc.styles.add_style("Probe", 1)
        pPr = style.element.get_or_add_pPr()
        pPr.append(OxmlElement("w:" + name))
        dc._apply_style_outline_level(style, 1)
        tags = [c.tag.split("}")[-1] for c in pPr]
        assert tags == ["outlineLvl", name], (name, tags)


def test_set_style_fonts_handles_the_ten_names_the_old_list_dropped():
    """Where the old literal was not shorter but **wrong**, and why it never bit.

    With only ``w:caps`` in ``w:rPr`` the old 27-name list matched nothing and
    ``insert_element_before`` fell through to ``append`` — producing
    ``[caps, rFonts]``, i.e. ``rFonts`` **after** ``caps``, which Word drops.
    The derived list finds ``w:caps`` and inserts before it.

    This is a behaviour change, and it is unreachable from ``add_styles``:
    ``font.name`` runs first and python-docx writes ``w:rFonts`` there, so the
    ``rf is None`` branch only fires for a caller that hands
    ``_set_style_fonts`` an rPr python-docx has not touched.  Recorded here so
    the difference is a documented fix rather than a surprise.
    """
    for name in RFONTS_GAP:
        style, rpr = _bare_style_with_rpr_children([name])
        dc._set_style_fonts(style, "黑体", "宋体")
        tags = [c.tag.split("}")[-1] for c in rpr]
        assert tags.index("rFonts") < tags.index(name), (name, tags)


def test_the_rfonts_branch_is_unreachable_from_add_styles():
    """``font.name`` always creates ``w:rFonts`` first, so the branch is dead.

    Asserted rather than assumed: the whole ``rf is None`` path exists for a
    caller that has not run python-docx's font setter, and ``add_styles`` always
    has.  If a future edit reorders those two lines this test goes red, which is
    the point.
    """
    doc = Document()
    style = doc.styles.add_style("Probe", 1)
    rpr = style.element.get_or_add_rPr()
    assert rpr.find(qn("w:rFonts")) is None
    style.font.name = "黑体"                      # what add_styles does first
    assert rpr.find(qn("w:rFonts")) is not None
    # ... and a second _set_style_fonts call then takes the find-existing path,
    # which is what makes the font write idempotent instead of duplicating.
    dc._set_style_fonts(style, "宋体", None)
    fonts = [c for c in rpr if c.tag == qn("w:rFonts")]
    assert len(fonts) == 1
    assert fonts[0].get(qn("w:ascii")) == "宋体"


def test_add_styles_with_font_keeps_rfonts_first_in_rpr():
    """End-to-end: the spec path that actually runs still produces legal order."""
    doc = Document()
    dc.add_styles(doc, [{"name": "Head", "base": "Normal", "font": "黑体",
                         "size_pt": 16, "bold": True, "color": "1F4E79",
                         "outline_level": 1}])
    rpr = doc.styles["Head"].element.find(qn("w:rPr"))
    tags = [c.tag.split("}")[-1] for c in rpr]
    assert tags[0] == "rFonts"
    order = {name: i for i, name in enumerate(RPR_ORDER)}
    present = [order[t] for t in tags if t in order]
    assert present == sorted(present), tags
