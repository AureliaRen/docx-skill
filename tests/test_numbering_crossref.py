# -*- coding: utf-8 -*-
"""Tests for auto-numbering and cross-references.

Both features are pure OOXML: the failure modes are structural (wrong child
order, a ``numId`` that names no ``<w:num>``, a field with no cached result),
so every test inspects the XML rather than trusting a round trip.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from docx_crossref import (  # noqa: E402
    all_bookmarks, bm_name, cmd_caption, cmd_ref, cmd_renumber, field_elements,
    seq_counters,
)
from docx_numbering import (  # noqa: E402
    cmd_apply, cmd_define, cmd_detach, cmd_list, cmd_restart, current_numpr,
    set_numpr,
)


@pytest.fixture()
def doc():
    d = Document()
    d.add_heading("Title", level=1)
    for i in range(6):
        d.add_paragraph(f"条目 {i}")
    for i in range(3):
        d.add_table(rows=1, cols=1).cell(0, 0).text = f"表内 {i}"
    return d


class Args:
    """Minimal argparse stand-in."""

    def __init__(self, **kw):
        self.__dict__.update(kw)


def P(d, i):
    return d.paragraphs[i]


# ---------------------------------------------------------------------------
# numbering
# ---------------------------------------------------------------------------
def test_define_creates_definition_and_counter(doc):
    r = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                             indent_left=425, hanging=425))
    assert r["ok"] and r["numId"] >= 1
    listing = cmd_list(doc)
    assert any(n["numId"] == str(r["numId"]) for n in listing["nums"])
    assert any(a["lvlText"] == "[%1]" for a in listing["abstract_nums"])


def test_define_rejects_unknown_format(doc):
    with pytest.raises(SystemExit) as e:
        cmd_define(doc, Args(text="%1", fmt="nonsense", start=1, jc="left",
                             indent_left=0, hanging=0))
    assert "unknown --fmt" in str(e.value)


def test_define_appends_counter_after_all_abstract_nums(doc):
    """CT_Numbering is xsd:sequence: every abstractNum precedes every num."""
    num_id = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                                  indent_left=0, hanging=0))["numId"]
    root = doc.part.numbering_part.element
    tags = [c.tag.split("}")[-1] for c in root]
    first_num = tags.index("num")
    assert "abstractNum" not in tags[first_num:], tags
    assert num_id == int(root.findall(qn("w:num"))[-1].get(qn("w:numId")))


def test_apply_sets_numpr_and_reports_count(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    r = cmd_apply(doc, Args(index=["1:4"], style=None, match=None,
                            num_id=nid, ilvl=0))
    assert r["applied"] == 3
    assert [current_numpr(P(doc, i)) for i in (1, 2, 3)] == [(nid, 0)] * 3
    assert current_numpr(P(doc, 4)) is None


def test_apply_unknown_numid_fails(doc):
    with pytest.raises(SystemExit) as e:
        cmd_apply(doc, Args(index=["1:2"], style=None, match=None,
                            num_id=9999, ilvl=0))
    assert "no such numId" in str(e.value)


def test_apply_by_style_and_by_match(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    r = cmd_apply(doc, Args(index=None, style="Heading 1", match=None,
                            num_id=nid, ilvl=0))
    assert r["applied"] == 1
    r = cmd_apply(doc, Args(index=None, style=None, match="条目",
                            num_id=nid, ilvl=0))
    assert r["applied"] == 6


def test_selectors_combine_with_and(doc):
    """Style + index must intersect, not union.

    This is the case that matters in practice: one body style can carry both
    prose and the reference list, so "style alone" is too broad.
    """
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    r = cmd_apply(doc, Args(index=["2:5"], style="Normal", match=None,
                            num_id=nid, ilvl=0))
    assert r["applied"] == 3, r          # paragraphs 2,3,4 — not 1 and not 5+
    r = cmd_apply(doc, Args(index=None, style="Normal", match="条目",
                            num_id=nid, ilvl=0))
    assert r["applied"] == 6             # every 条目 paragraph
    r = cmd_apply(doc, Args(index=["1:3"], style="Normal", match="条目",
                            num_id=nid, ilvl=0))
    assert r["applied"] == 2             # 条目 AND inside 1:3


def test_apply_skips_paragraphs_failing_any_selector(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    # the heading has the style but not the text; the 条目s have the text but
    # not the style; the intersection is empty, which is an error rather than
    # a silent no-op.
    with pytest.raises(SystemExit) as e:
        cmd_apply(doc, Args(index=None, style="Heading 1", match="条目",
                            num_id=nid, ilvl=0))
    assert "matched no paragraphs" in str(e.value)
    assert current_numpr(P(doc, 0)) is None
    assert current_numpr(P(doc, 2)) is None


def test_numpr_child_order_matches_schema(doc):
    """CT_NumPr is a sequence: ilvl before numId, or Word drops the list."""
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    set_numpr(P(doc, 2), nid, 0)
    numPr = P(doc, 2)._p.pPr.find(qn("w:numPr"))
    assert [c.tag.split("}")[-1] for c in numPr] == ["ilvl", "numId"]


def test_numpr_sits_after_pstyle_but_before_spacing(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    p = P(doc, 1)
    from docx.oxml import OxmlElement
    sp = OxmlElement("w:spacing")
    p._p.get_or_add_pPr().append(sp)
    set_numpr(p, nid, 0)
    tags = [c.tag.split("}")[-1] for c in p._p.pPr]
    assert tags.index("numPr") < tags.index("spacing")


def test_restart_uses_a_distinct_counter(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    cmd_apply(doc, Args(index=["1:6"], style=None, match=None,
                        num_id=nid, ilvl=0))
    r = cmd_restart(doc, Args(index=["4:5"], style=None, match=None,
                              num_id=nid, ilvl=0, start_value=1))
    assert r["newNumId"] != nid
    assert r["applied"] == 1
    assert current_numpr(P(doc, 4))[0] == r["newNumId"]
    assert current_numpr(P(doc, 3))[0] == nid


def test_detach_removes_numpr(doc):
    nid = cmd_define(doc, Args(text="%1", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    cmd_apply(doc, Args(index=["1:6"], style=None, match=None,
                        num_id=nid, ilvl=0))
    r = cmd_detach(doc, Args(index=["1:3"], style=None, match=None))
    assert r["detached"] == 2
    assert current_numpr(P(doc, 2)) is None
    assert current_numpr(P(doc, 3)) == (nid, 0)


def test_numbering_survives_save_and_reload(tmp_path, doc):
    nid = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                               indent_left=0, hanging=0))["numId"]
    cmd_apply(doc, Args(index=["1:4"], style=None, match=None,
                        num_id=nid, ilvl=0))
    out = tmp_path / "n.docx"
    doc.save(out)
    again = Document(out)
    assert current_numpr(again.paragraphs[1]) == (nid, 0)
    assert cmd_list(again)["usage"] == {str(nid): 3}


# ---------------------------------------------------------------------------
# cross references
# ---------------------------------------------------------------------------
def test_caption_creates_bookmark_and_seq_field(doc):
    r = cmd_caption(doc, Args(index=1, label="图", text="架构", bookmark=None))
    assert r["number"] == 1
    assert r["bookmark"] in all_bookmarks(doc)
    assert seq_counters(doc) == {"图": 1}
    xml = P(doc, 1)._p.xml
    assert "SEQ 图" in xml and "fldCharType=\"separate\"" in xml
    assert "架构" in P(doc, 1).text


def test_caption_counter_increments_per_label(doc):
    cmd_caption(doc, Args(index=1, label="图", text="", bookmark=None))
    cmd_caption(doc, Args(index=2, label="图", text="", bookmark=None))
    cmd_caption(doc, Args(index=3, label="表", text="", bookmark=None))
    assert seq_counters(doc) == {"图": 2, "表": 1}
    assert bm_name("图", 2) in all_bookmarks(doc)


def test_ref_needs_an_existing_bookmark(doc):
    with pytest.raises(SystemExit) as e:
        cmd_ref(doc, Args(index=5, target="nope_1", prefix="", suffix="",
                          format="default", cached=None))
    assert "no such bookmark" in str(e.value)


def test_ref_builds_a_field_with_cached_result(doc):
    r = cmd_caption(doc, Args(index=1, label="图", text="架构", bookmark=None))
    before = P(doc, 5).text
    ref = cmd_ref(doc, Args(index=5, target=r["bookmark"], prefix="如图",
                            suffix="所示", format="default", cached=None))
    assert ref["cached"] == "1"
    text = P(doc, 5).text
    # ref appends to the paragraph rather than replacing it
    assert text == before + "如图1所示", text
    assert f"REF {r['bookmark']}" in P(doc, 5)._p.xml


def test_ref_cached_value_is_what_renders(doc):
    """A field with no cached result shows nothing until Word updates it."""
    r = cmd_caption(doc, Args(index=1, label="图", text="", bookmark=None))
    before = P(doc, 5).text
    cmd_ref(doc, Args(index=5, target=r["bookmark"], prefix="", suffix="",
                      format="default", cached="7"))
    assert P(doc, 5).text == before + "7"


def test_renumber_fixes_stale_cached_results(doc):
    """Simulate an edit that shifted the numbers, then repair them."""
    r1 = cmd_caption(doc, Args(index=1, label="图", text="", bookmark=None))
    r2 = cmd_caption(doc, Args(index=2, label="图", text="", bookmark=None))
    cmd_ref(doc, Args(index=5, target=r2["bookmark"], prefix="如图",
                      suffix="", format="default", cached="9"))
    # corrupt both cached results the way a bad edit would
    for i in (1, 2):
        for t in P(doc, i)._p.iter(qn("w:t")):
            if (t.text or "").isdigit():
                t.text = "99"
    for t in P(doc, 5)._p.iter(qn("w:t")):
        if (t.text or "").isdigit():
            t.text = "99"
    fixed = cmd_renumber(doc, Args())
    assert fixed["seq_updated"] == 2
    assert fixed["ref_updated"] == 1
    assert [t.text for t in P(doc, 1)._p.iter(qn("w:t"))
            if (t.text or "").isdigit()] == ["1"]
    assert [t.text for t in P(doc, 2)._p.iter(qn("w:t"))
            if (t.text or "").isdigit()] == ["2"]
    assert [t.text for t in P(doc, 5)._p.iter(qn("w:t"))
            if (t.text or "").isdigit()] == ["2"]


def test_renumber_is_idempotent(doc):
    cmd_caption(doc, Args(index=1, label="图", text="", bookmark=None))
    cmd_caption(doc, Args(index=2, label="图", text="", bookmark=None))
    first = cmd_renumber(doc, Args())
    second = cmd_renumber(doc, Args())
    assert first["seq_updated"] == 0
    assert second["seq_updated"] == 0
    assert second["counters"] == {"图": 2}


def test_field_elements_shape():
    els = field_elements("REF fig_1 \\h", "3")
    kinds = []
    for e in els:
        f = e.find(qn("w:fldChar"))
        if f is not None:
            kinds.append(f.get(qn("w:fldCharType")))
    assert kinds == ["begin", "separate", "end"]
    instr = [e.find(qn("w:instrText")) for e in els]
    assert any(i is not None and "REF fig_1" in i.text for i in instr)


def test_bookmark_names_are_word_legal():
    assert bm_name("图", 3) == "图_3"
    assert bm_name("Table", 12) == "Table_12"
    assert len(bm_name("非常长的标签名称" * 5, 1)) <= 40
    assert bm_name("a b/c", 1).isidentifier() or "_" in bm_name("a b/c", 1)
