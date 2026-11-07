# -*- coding: utf-8 -*-
"""Cross-references v2: real hyperlinks, vertAlign, recursive renumber.

The v1 tests (``test_numbering_crossref.py``) cover numbering and the basic
field shape.  These cover what the owner actually asked for:

* the reference must be **clickable** — ``w:hyperlink w:anchor`` around the
  field, not just the ``\\h`` switch;
* it must render **superscript or normal size**;
* ``renumber`` must still see a field that now lives inside a hyperlink;
* the cached result must match what Word itself computes for each flavour;
* ``caption`` must stop throwing the paragraph's own content away.

Everything is XML-level and ``tmp_path`` based, so it runs without Word.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document                                   # noqa: E402
from docx.oxml import OxmlElement                           # noqa: E402
from docx.oxml.ns import qn                                 # noqa: E402

from docx_crossref import (                                 # noqa: E402
    CACHED_KIND, MERGE_FORMAT, REF_FORMATS, RPR_ORDER, VERT_ALIGN, all_bookmarks,
    bookmark_span_text, bookmarked_text, cached_result, cmd_bookmark,
    cmd_caption, cmd_list, cmd_ref, cmd_renumber, field_elements,
    hyperlink_wrap, list_number_for, paragraph_text, with_vert_align,
)
from docx_numbering import cmd_define, set_numpr              # noqa: E402


class Args:
    """Minimal argparse stand-in."""

    def __init__(self, **kw):
        self.__dict__.update(kw)


@pytest.fixture()
def doc():
    d = Document()
    for i in range(5):
        d.add_paragraph(f"正文 {i}")
    return d


def P(d, i):
    return d.paragraphs[i]


def hyperlinks(p) -> list:
    return list(p._p.iter(qn("w:hyperlink")))


def wrapped_texts(p) -> list:
    """Text carried by the reference itself (inside any hyperlink wrapper)."""
    out = []
    for h in hyperlinks(p):
        out.extend(t.text for t in h.iter(qn("w:t")) if (t.text or "").strip())
    return out


def rpr_of(el):
    return el.find(qn("w:rPr"))


def caption(d, index, label="图", text="", **kw):
    kw.setdefault("bookmark", None)
    return cmd_caption(d, Args(index=index, label=label, text=text, **kw))


def ref(d, index, target, **kw):
    kw.setdefault("prefix", "")
    kw.setdefault("suffix", "")
    kw.setdefault("format", "default")
    kw.setdefault("cached", None)
    return cmd_ref(d, Args(index=index, target=target, **kw))


# ---------------------------------------------------------------------------
# 1. real hyperlink wrapper
# ---------------------------------------------------------------------------
def test_ref_wraps_field_in_real_hyperlink(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"])
    hs = hyperlinks(P(doc, 3))
    assert len(hs) == 1, P(doc, 3)._p.xml
    assert out["hyperlink"] is True and out["anchor"] == r["bookmark"]
    # the whole field lives inside the wrapper
    inner = list(hs[0])
    kinds = [e.find(qn("w:fldChar")).get(qn("w:fldCharType"))
             for e in inner if e.find(qn("w:fldChar")) is not None]
    assert kinds == ["begin", "separate", "end"]


def test_ref_anchor_is_the_bookmark_and_history_is_on(doc):
    r = caption(doc, 1, text="架构")
    cmd_ref(doc, Args(index=3, target=r["bookmark"], prefix="", suffix="",
                      format="default", cached=None))
    h = hyperlinks(P(doc, 3))[0]
    assert h.get(qn("w:anchor")) == r["bookmark"]
    assert h.get(qn("w:history")) == "1"
    # internal anchor => no relationship id needed
    assert h.get(qn("r:id")) is None


def test_hyperlink_wrapper_keeps_the_inner_ref_h_switch(doc):
    r = caption(doc, 1, text="架构")
    cmd_ref(doc, Args(index=3, target=r["bookmark"], prefix="", suffix="",
                      format="default", cached=None))
    instr = "".join(i.text or ""
                    for i in P(doc, 3)._p.iter(qn("w:instrText")))
    assert "REF" in instr and r["bookmark"] in instr and "\\h" in instr


def test_no_hyperlink_flag_writes_a_bare_field(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"], no_hyperlink=True)
    assert out["hyperlink"] is False and out["anchor"] is None
    assert hyperlinks(P(doc, 3)) == []
    assert "REF" in P(doc, 3)._p.xml


def test_hyperlink_survives_save_and_reload(tmp_path, doc):
    r = caption(doc, 1, text="架构")
    before = P(doc, 3).text
    cmd_ref(doc, Args(index=3, target=r["bookmark"], prefix="如图", suffix="所示",
                      format="default", cached=None))
    out = tmp_path / "x.docx"
    doc.save(out)
    again = Document(out)
    assert "如图1所示" in again.paragraphs[3].text
    assert len(hyperlinks(again.paragraphs[3])) == 1
    assert again.paragraphs[3].text.startswith(before)


# ---------------------------------------------------------------------------
# 2. superscript / normal / subscript
# ---------------------------------------------------------------------------
def superscripts(runs) -> list:
    out = []
    for e in runs:
        rpr = rpr_of(e)
        va = None if rpr is None else rpr.find(qn("w:vertAlign"))
        out.append(None if va is None else va.get(qn("w:val")))
    return out


def test_ref_superscript_marks_every_field_run(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"], vert_align="superscript")
    assert out["vert_align"] == "superscript"
    inner = list(hyperlinks(P(doc, 3))[0])
    assert len(inner) == 5
    assert superscripts(inner) == ["superscript"] * 5


def test_superscript_shorthand_matches_vert_align_flag(doc):
    r = caption(doc, 1, text="架构")
    ref(doc, 3, r["bookmark"], superscript=True)
    ref(doc, 4, r["bookmark"], vert_align="superscript")
    for i in (3, 4):
        inner = list(hyperlinks(P(doc, i))[0])
        assert superscripts(inner) == ["superscript"] * 5


def test_ref_subscript_and_baseline(doc):
    r = caption(doc, 1, text="架构")
    ref(doc, 3, r["bookmark"], vert_align="subscript")
    ref(doc, 4, r["bookmark"], vert_align="baseline")
    assert superscripts(list(hyperlinks(P(doc, 3))[0])) == ["subscript"] * 5
    assert superscripts(list(hyperlinks(P(doc, 4))[0])) == ["baseline"] * 5


def test_normal_vert_align_adds_no_element(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"])
    assert out["vert_align"] == "normal"
    assert superscripts(list(hyperlinks(P(doc, 3))[0])) == [None] * 5


def test_normal_ref_after_superscript_ref_is_not_superscript(doc):
    """Inherited run properties must not leak an earlier superscript."""
    r = caption(doc, 1, text="架构")
    ref(doc, 3, r["bookmark"], superscript=True)
    ref(doc, 3, r["bookmark"])
    assert superscripts(list(hyperlinks(P(doc, 3))[1])) == [None] * 5


def test_vert_align_sits_in_the_ct_rpr_sequence_slot():
    rpr = OxmlElement("w:rPr")
    for tag in ("w:rFonts", "w:sz", "w:u", "w:lang"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "x")
        rpr.append(el)
    out = with_vert_align(rpr, "superscript")
    tags = [c.tag.split("}")[-1] for c in out]
    assert "vertAlign" in RPR_ORDER
    # CT_RPr is a sequence: after sz/u (and fitText), before rtl/lang
    assert tags.index("sz") < tags.index("vertAlign") < tags.index("lang")
    assert tags.index("u") < tags.index("vertAlign")
    assert RPR_ORDER.index("sz") < RPR_ORDER.index("vertAlign") \
        < RPR_ORDER.index("lang")


def test_with_vert_align_is_idempotent_and_strips_inherited():
    rpr = OxmlElement("w:rPr")
    first = with_vert_align(rpr, "superscript")
    second = with_vert_align(first, "superscript")
    assert len(second.findall(qn("w:vertAlign"))) == 1
    stripped = with_vert_align(first, None)
    assert stripped is None or stripped.find(qn("w:vertAlign")) is None


def test_with_vert_align_keeps_other_run_properties():
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "32")
    rpr.append(sz)
    out = with_vert_align(rpr, "superscript")
    assert out.find(qn("w:sz")).get(qn("w:val")) == "32"


# ---------------------------------------------------------------------------
# 3. renumber sees through the hyperlink wrapper
# ---------------------------------------------------------------------------
def corrupt_numbers(d, indexes) -> None:
    for i in indexes:
        for t in P(d, i)._p.iter(qn("w:t")):
            if (t.text or "").isdigit():
                t.text = "99"


def ref_texts(d, i) -> list:
    return [t.text for t in P(d, i)._p.iter(qn("w:t")) if (t.text or "").strip()]


def test_renumber_rewrites_refs_inside_a_hyperlink(doc):
    """Regression lock: a flat child walk used to miss wrapped fields."""
    r1 = caption(doc, 1)
    r2 = caption(doc, 2)
    cmd_ref(doc, Args(index=4, target=r2["bookmark"], prefix="如图", suffix="",
                      format="default", cached="9"))
    assert hyperlinks(P(doc, 4)), "reference must be wrapped"
    corrupt_numbers(doc, [1, 2, 4])
    fixed = cmd_renumber(doc, Args())
    assert fixed["seq_updated"] == 2
    assert fixed["ref_updated"] == 1
    assert wrapped_texts(P(doc, 4)) == ["2"]
    assert P(doc, 4).text == "正文 4如图2"


def test_renumber_is_still_idempotent_with_wrappers(doc):
    r = caption(doc, 1)
    cmd_ref(doc, Args(index=3, target=r["bookmark"], prefix="", suffix="",
                      format="default", cached=None))
    first = cmd_renumber(doc, Args())
    second = cmd_renumber(doc, Args())
    assert first["seq_updated"] == 0 and first["ref_updated"] == 0
    assert second["seq_updated"] == 0 and second["ref_updated"] == 0


def test_renumber_ignores_pageref_and_above_below(doc):
    r = caption(doc, 1)
    ref(doc, 3, r["bookmark"], format="page", cached="7")
    ref(doc, 4, r["bookmark"], format="above-below", cached="above")
    fixed = cmd_renumber(doc, Args())
    assert fixed["ref_updated"] == 0
    assert fixed["ref_skipped"] == 2
    assert wrapped_texts(P(doc, 3)) == ["7"]
    assert wrapped_texts(P(doc, 4)) == ["above"]


def test_renumber_refreshes_a_paragraph_number_from_the_list(doc):
    nid = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                               indent_left=425, hanging=425))["numId"]
    for i in (1, 2, 3):
        set_numpr(P(doc, i), nid, 0)
    cmd_bookmark(doc, Args(index=2, name="ref_2"))
    ref(doc, 4, "ref_2", format="paragraph-number", cached="9")
    fixed = cmd_renumber(doc, Args())
    assert fixed["ref_updated"] == 1
    assert wrapped_texts(P(doc, 4)) == ["[2]"]


def test_renumber_refreshes_a_wide_caption_cache(doc):
    r = caption(doc, 1, text="架构", bookmark_span="paragraph")
    ref(doc, 3, r["bookmark"], format="full", cached="99")
    fixed = cmd_renumber(doc, Args())
    assert fixed["ref_updated"] == 1
    assert wrapped_texts(P(doc, 3)) == ["图 1  架构"]


def test_renumber_of_a_ref_k_bookmark_keeps_the_declared_number(doc):
    """ref_3 names reference 3; renumber must not paste the item's prose."""
    P(doc, 1).add_run("盛万兴. 电力系统分析.")
    cmd_bookmark(doc, Args(index=1, name="ref_3"))
    ref(doc, 4, "ref_3", cached="9")
    fixed = cmd_renumber(doc, Args())
    assert fixed["ref_updated"] == 1
    assert wrapped_texts(P(doc, 4)) == ["3"]


# ---------------------------------------------------------------------------
# 4. Word-faithful flavours and cached results
# ---------------------------------------------------------------------------
def test_cached_kind_covers_every_format():
    assert set(CACHED_KIND) == set(REF_FORMATS)


def test_ref_page_format_is_a_pageref_field(doc):
    r = caption(doc, 1)
    out = ref(doc, 3, r["bookmark"], format="page")
    assert out["field"] == "PAGEREF"
    assert "PAGEREF" in P(doc, 3)._p.xml
    assert "\\p" not in P(doc, 3)._p.xml


def test_ref_above_below_uses_the_p_switch(doc):
    r = caption(doc, 1)
    out = ref(doc, 3, r["bookmark"], format="above-below")
    assert out["field"] == "REF" and out["switch"] == r" \p" + MERGE_FORMAT
    assert "\\p" in P(doc, 3)._p.xml


def test_content_text_flavours_cache_the_bookmarked_text(doc):
    r = caption(doc, 1, text="架构")
    assert cached_result(doc, r["bookmark"], "default") == "1"
    assert cached_result(doc, r["bookmark"], "number") == "1"
    wide = caption(doc, 2, text="架构", bookmark_span="paragraph")
    assert cached_result(doc, wide["bookmark"], "full") == "图 2  架构"
    assert cached_result(doc, wide["bookmark"], "default") == "图 2  架构"


def test_ref_cached_inferred_from_a_ref_k_suffix(doc):
    for i in (1, 2, 3):
        cmd_bookmark(doc, Args(index=i, name="ref_" + str(i)))
    out = ref(doc, 4, "ref_3")
    assert out["cached"] == "3"


def test_ref_cached_is_overridable(doc):
    cmd_bookmark(doc, Args(index=1, name="ref_1"))
    out = ref(doc, 4, "ref_1", cached="[1]", superscript=True)
    assert out["cached"] == "[1]"
    assert wrapped_texts(P(doc, 4)) == ["[1]"]


def test_above_below_and_page_cache_nothing_by_default(doc):
    r = caption(doc, 1, text="架构")
    assert cached_result(doc, r["bookmark"], "above-below") == ""
    assert cached_result(doc, r["bookmark"], "page") == ""
    ref(doc, 3, r["bookmark"], format="page")
    assert wrapped_texts(P(doc, 3)) == []


def test_list_number_for_matches_word_liststring(doc):
    nid = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                               indent_left=425, hanging=425))["numId"]
    for i in (1, 2, 3):
        set_numpr(P(doc, i), nid, 0)
    assert [list_number_for(doc, P(doc, i)._p) for i in (1, 2, 3)] == \
        ["[1]", "[2]", "[3]"]
    assert list_number_for(doc, P(doc, 0)._p) == ""


def test_bookmarked_text_helpers(doc):
    r = caption(doc, 1, text="架构")
    p = P(doc, 1)._p
    assert bookmarked_text(doc, r["bookmark"]) == "1"
    assert bookmark_span_text(p, r["bookmark"]) == "1"
    assert bookmark_span_text(p, "nope") is None
    assert bookmarked_text(doc, "nope") is None
    assert paragraph_text(p) == "图 1  架构"


# ---------------------------------------------------------------------------
# 5. non-destructive captions
# ---------------------------------------------------------------------------
def test_caption_preserves_paragraph_properties(doc):
    p = P(doc, 1)
    p.style = doc.styles["Normal"]
    p.alignment = 1                                     # center
    caption(doc, 1, text="架构")
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    assert P(doc, 1).alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert P(doc, 1)._p.find(qn("w:pPr")) is not None


def test_caption_keeps_existing_text_when_no_text_given(doc):
    out = caption(doc, 2)
    assert out["preserved_text"] == "正文 2"
    assert P(doc, 2).text == "图 1  正文 2"


def test_caption_drop_text_discards_it(doc):
    out = caption(doc, 2, drop_text=True)
    assert out["preserved_text"] is None
    assert P(doc, 2).text == "图 1"


def test_caption_text_argument_wins_over_existing_text(doc):
    caption(doc, 2, text="架构")
    assert P(doc, 2).text == "图 1  架构"


def test_caption_inherits_first_run_properties(doc):
    P(doc, 2).runs[0].bold = True
    caption(doc, 2, text="架构")
    first = P(doc, 2)._p.find(qn("w:r"))
    rpr = rpr_of(first)
    assert rpr is not None and rpr.find(qn("w:b")) is not None


def test_caption_inherits_rpr_from_a_run_inside_a_hyperlink(doc):
    r = caption(doc, 1, text="架构")
    ref(doc, 2, r["bookmark"])                  # wrapper run, no rPr
    cmd_caption(doc, Args(index=2, label="表", text="表题", bookmark=None))
    xml = P(doc, 2)._p.xml
    assert "SEQ 表" in xml and "fldCharType=\"separate\"" in xml


def test_recaption_updates_the_number_in_place_and_keeps_rich_runs(doc):
    r1 = caption(doc, 1, text="架构")
    caption(doc, 2, text="流程")
    p = P(doc, 1)._p
    rich = OxmlElement("w:r")
    rich.append(OxmlElement("w:rPr"))
    t = OxmlElement("w:t")
    t.text = "【手工加粗】"
    rich.append(t)
    p.append(rich)

    out = caption(doc, 1)
    assert out["recaption"] is True and out["number"] == 1
    assert out["bookmark"] == r1["bookmark"]
    assert "手工加粗" in P(doc, 1).text, "rich run must survive a re-caption"
    assert P(doc, 1).text == "图 1  架构【手工加粗】"


def test_recaption_with_text_replaces_only_the_trailing_text(doc):
    caption(doc, 1, text="旧题注")
    out = caption(doc, 1, text="新题注")
    assert out["recaption"] is True
    assert P(doc, 1).text == "图 1  新题注"


def test_recaption_does_not_duplicate_the_seq_counter(doc):
    caption(doc, 1, text="架构")
    caption(doc, 2, text="流程")
    caption(doc, 1)
    cmd_renumber(doc, Args())
    assert cmd_list(doc)["counters"] == {"图": 2}
    assert P(doc, 1).text.startswith("图 1")


def test_recaption_creates_a_missing_bookmark(doc):
    p = P(doc, 2)._p
    for el in field_elements("SEQ 图 \\* ARABIC", "1"):
        p.append(el)
    out = caption(doc, 2)
    assert out["recaption"] is True
    assert out["bookmark"] == "图_1" and out["bookmark_created"] is True
    assert "图_1" in all_bookmarks(doc)
    assert bookmark_span_text(p, "图_1") == "1"


def test_caption_bookmark_span_paragraph_covers_the_caption(doc):
    out = caption(doc, 1, text="架构", bookmark_span="paragraph")
    assert out["bookmark_span"] == "paragraph"
    assert bookmarked_text(doc, out["bookmark"]) == "图 1  架构"
    narrow = caption(doc, 2, text="流程")
    assert bookmarked_text(doc, narrow["bookmark"]) == "2"


def test_caption_rejects_an_unknown_bookmark_span(doc):
    with pytest.raises(SystemExit) as e:
        caption(doc, 1, text="架构", bookmark_span="nonsense")
    assert "bookmark-span" in str(e.value)


# ---------------------------------------------------------------------------
# 6. diagnostics + reference-list scenario end to end
# ---------------------------------------------------------------------------
def test_list_reports_hyperlink_anchors(doc):
    r = caption(doc, 1, text="架构")
    ref(doc, 3, r["bookmark"])
    ref(doc, 4, r["bookmark"], no_hyperlink=True)
    listing = cmd_list(doc)
    assert listing["hyperlinks"] == 1
    assert listing["hyperlink_anchors"] == [r["bookmark"]]


def test_reference_list_citation_is_superscript_and_clickable(doc):
    """The 主人's scenario: superscript [1] that jumps to reference 1."""
    nid = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                               indent_left=425, hanging=425))["numId"]
    set_numpr(P(doc, 1), nid, 0)
    set_numpr(P(doc, 2), nid, 0)
    P(doc, 1).add_run("盛万兴. 电力系统分析.")
    P(doc, 2).add_run("Krausman A. Radar.")
    cmd_bookmark(doc, Args(index=1, name="ref_1"))
    out = ref(doc, 3, "ref_1", format="paragraph-number", superscript=True)
    assert out["cached"] == "[1]"
    h = hyperlinks(P(doc, 3))[0]
    assert h.get(qn("w:anchor")) == "ref_1"
    assert superscripts(list(h)) == ["superscript"] * 5
    assert wrapped_texts(P(doc, 3)) == ["[1]"]


def test_hyperlink_wrap_is_importable_and_builds_a_link():
    h = hyperlink_wrap("图_1", field_elements("REF 图_1 \\h", "1"))
    assert h.get(qn("w:anchor")) == "图_1"
    assert len(list(h)) == 5


def test_vert_align_choices_are_exposed():
    assert set(VERT_ALIGN) == {"normal", "baseline", "superscript", "subscript"}
    assert VERT_ALIGN["normal"] is None


# ---------------------------------------------------------------------------
# 7. \* MERGEFORMAT — the switch that keeps the formatting across an update
# ---------------------------------------------------------------------------
def test_ref_carries_mergeformat_by_default(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"], superscript=True)
    assert out["merge_format"] is True
    instr = "".join(i.text or ""
                    for i in P(doc, 3)._p.iter(qn("w:instrText")))
    assert MERGE_FORMAT.strip() in instr


def test_no_mergeformat_flag_omits_it(doc):
    r = caption(doc, 1, text="架构")
    out = ref(doc, 3, r["bookmark"], no_mergeformat=True)
    assert out["merge_format"] is False
    instr = "".join(i.text or ""
                    for i in P(doc, 3)._p.iter(qn("w:instrText")))
    assert "MERGEFORMAT" not in instr


def test_mergeformat_is_added_to_pageref_and_above_below(doc):
    r = caption(doc, 1, text="架构")
    out_page = ref(doc, 3, r["bookmark"], format="page")
    out_ab = ref(doc, 4, r["bookmark"], format="above-below")
    assert out_page["switch"].endswith(MERGE_FORMAT)
    assert out_page["field"] == "PAGEREF"
    assert out_ab["switch"].endswith(MERGE_FORMAT)


def test_renumber_parses_codes_that_carry_mergeformat(doc):
    r1 = caption(doc, 1)
    r2 = caption(doc, 2)
    ref(doc, 4, r2["bookmark"])
    assert "MERGEFORMAT" in P(doc, 4)._p.xml
    corrupt_numbers(doc, [1, 2, 4])
    fixed = cmd_renumber(doc, Args())
    assert fixed["seq_updated"] == 2 and fixed["ref_updated"] == 1
    assert wrapped_texts(P(doc, 4)) == ["2"]


def test_renumber_skips_above_below_that_carries_mergeformat(doc):
    r = caption(doc, 1)
    ref(doc, 3, r["bookmark"], format="above-below", cached="above")
    fixed = cmd_renumber(doc, Args())
    assert fixed["ref_updated"] == 0 and fixed["ref_skipped"] == 1
    assert wrapped_texts(P(doc, 3)) == ["above"]
