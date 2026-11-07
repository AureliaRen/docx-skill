# -*- coding: utf-8 -*-
"""Caption ``_Ref`` bookmarks: the anchor a table of figures links through.

Why this file exists
--------------------
Word's "Insert Table of Figures" (``TOC \\c "图"``) does not copy caption text.
It writes one ``PAGEREF`` field per entry, and every ``PAGEREF`` points at a
**bookmark**.  Word's own caption machinery names that family ``_Ref<number>``.
A caption that only carries a hand-chosen name (``fig_1``) inserts into the
table of figures, but every entry is dead on click, because the anchor the
generated ``PAGEREF`` needs does not exist.  That is the operator's report.

``cmd_caption --ref-bookmark`` therefore *adds* a ``_Ref<number>`` bookmark
around the whole caption while keeping the caller's own bookmark, so one
caption can be referenced both ways at once.

Fact base (``work/src.docx``, a Word-saved reference document)
--------------------------------------------------------------
The reference document does **not** put ``_Ref…`` on captions: its 17
``_Ref…`` bookmarks (ids 133..149, names ``_Ref3925`` … ``_Ref8179``, four-digit
non-consecutive Word counters) all sit on ``pStyle=27`` reference-list
paragraphs and are the targets of ``REF _Ref… \\r \\h`` citations.  Its 12
captions (``pStyle=28``, name 图表名) carry ``_Toc…`` bookmarks that start
*after* the label and SEQ field.  So the shape asserted here is the one Word
produces for a caption a table of figures can link to, not a shape copied out
of the reference document — and the assertions below pin the properties the
command depends on (whole-caption coverage, name shape, uniqueness, id
uniqueness, coexistence with the caller's bookmark), not one exact byte order.

Everything is XML level on ``tmp_path`` documents, so it runs without Word.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document                                   # noqa: E402
from docx.oxml.ns import qn                                 # noqa: E402

from docx_crossref import (                                 # noqa: E402
    all_bookmarks, bm_name, bookmark_span_text, cmd_caption, next_ref_name,
    paragraph_text, seq_counters,
)


class Args:
    """Minimal argparse stand-in."""

    def __init__(self, **kw):
        self.__dict__.update(kw)


@pytest.fixture()
def doc():
    d = Document()
    for i in range(6):
        d.add_paragraph(f"正文 {i}")
    return d


def P(d, i):
    return d.paragraphs[i]


def caption(d, index, label="图", text="", **kw):
    kw.setdefault("bookmark", None)
    kw.setdefault("ref_bookmark", True)
    return cmd_caption(d, Args(index=index, label=label, text=text, **kw))


def ref_bms(p) -> list:
    """The ``_Ref…`` bookmark names on paragraph *p*, in document order."""
    return [b.get(qn("w:name")) for b in p._p.iter(qn("w:bookmarkStart"))
            if (b.get(qn("w:name")) or "").startswith("_Ref")]


def all_bm(p) -> list:
    return [b.get(qn("w:name")) for b in p._p.iter(qn("w:bookmarkStart"))
            if b.get(qn("w:name"))]


def bm_ids(p) -> list:
    return [b.get(qn("w:id")) for b in p._p.iter(qn("w:bookmarkStart"))]


def instr_texts(d) -> list:
    return [(x.text or "").strip()
            for x in d.element.body.iter(qn("w:instrText"))]


# ---- 1. the bookmark exists, is named like Word's, and is unique -----------

def test_ref_bookmark_is_created_and_named_like_words(doc):
    out = caption(doc, 1, text="架构")
    assert out["ref_bookmark"], "cmd_caption reports the _Ref name it wrote"
    names = ref_bms(P(doc, 1))
    assert names == [out["ref_bookmark"]]
    assert names[0].startswith("_Ref")
    assert names[0][4:].isdigit(), (
        "Word numbers this family _Ref<digits>; anything else will not be "
        "picked up as a caption anchor")


def test_ref_bookmark_defaults_to_off(doc):
    """The lock that keeps every existing caller byte-identical."""
    out = cmd_caption(doc, Args(index=1, label="图", text="架构",
                                bookmark=None))
    assert out["ref_bookmark"] is None
    assert ref_bms(P(doc, 1)) == []
    assert all_bm(P(doc, 1)) == ["图_1"], (
        "without --ref-bookmark only the caller's own bookmark is written")


def test_ref_names_are_unique_across_captions(doc):
    names = []
    for i in (1, 2, 3, 4):
        names.append(caption(doc, i, text=f"题注 {i}")["ref_bookmark"])
    assert len(set(names)) == len(names), f"duplicate _Ref names: {names}"
    assert all(n.startswith("_Ref") and n[4:].isdigit() for n in names)


def test_ref_name_continues_past_existing_ref_bookmarks(doc):
    """A document that already carries Word-made _Ref names is not collided."""
    from docx_crossref import _wrap_whole_paragraph
    from docx_utils import next_bm_id
    _wrap_whole_paragraph(P(doc, 5)._p, "_Ref4211", next_bm_id(doc.element.body))
    out = caption(doc, 1, text="架构")
    assert out["ref_bookmark"] != "_Ref4211"
    assert int(out["ref_bookmark"][4:]) > 4211, (
        "the counter must march past what the document already holds")


def test_ref_bookmark_ids_do_not_collide(doc):
    caption(doc, 1, text="架构")
    caption(doc, 2, text="流程")
    for i in (1, 2):
        ids = bm_ids(P(doc, i))
        assert len(set(ids)) == len(ids), f"duplicate w:id in para {i}: {ids}"
    all_ids = [b.get(qn("w:id"))
               for b in doc.element.body.iter(qn("w:bookmarkStart"))]
    assert len(set(all_ids)) == len(all_ids), (
        "w:id is document-wide; duplicates make Word repair the file")


# ---- 2. the bookmark covers the WHOLE caption ------------------------------

def test_ref_bookmark_spans_the_whole_caption(doc):
    caption(doc, 1, text="系统架构示意图")
    p = P(doc, 1)
    span = bookmark_span_text(p._p, ref_bms(p)[0])
    full = paragraph_text(p._p)
    assert span == full, (
        f"the _Ref anchor must cover the whole caption so PAGEREF lands on the "
        f"caption and REF returns its text; got {span!r} of {full!r}")


def test_ref_bookmark_spans_label_number_and_text(doc):
    """Coverage is checked piecewise: a span that merely looks long is not it."""
    caption(doc, 1, text="系统架构示意图")
    p = P(doc, 1)
    span = bookmark_span_text(p._p, ref_bms(p)[0])
    assert span.startswith("图"), f"label outside the anchor: {span!r}"
    assert "1" in span, f"the SEQ number is outside the anchor: {span!r}"
    assert "系统架构示意图" in span, f"caption text outside the anchor: {span!r}"


def test_ref_bookmark_starts_right_after_ppr(doc):
    """Word reads the anchor from the body run order; it must precede the label."""
    caption(doc, 1, text="架构")
    p = P(doc, 1)._p
    kids = [c.tag.split("}")[1] for c in p]
    if kids[0] == "pPr":
        assert kids[1] == "bookmarkStart", (
            f"the _Ref opener must sit right after pPr, found order {kids[:4]}")
    else:
        assert kids[0] == "bookmarkStart", (
            f"a caption paragraph without pPr starts with its anchor, "
            f"found {kids[:4]}")
    assert kids.index("bookmarkStart") < kids.index("r"), (
        f"the anchor must precede the label run, order {kids[:6]}")
    assert kids[-1] == "bookmarkEnd", (
        f"the _Ref closer must be the paragraph's last child, got {kids[-1]}")


def test_ref_bookmark_follows_ppr_when_a_style_is_set(doc):
    """With --pstyle the paragraph does have a pPr, and the order still holds."""
    caption(doc, 1, text="架构", pstyle="28")
    kids = [c.tag.split("}")[1] for c in P(doc, 1)._p]
    assert kids[0] == "pPr"
    assert kids[1] == "bookmarkStart"
    assert kids[-1] == "bookmarkEnd"


def test_ref_bookmark_survives_a_recaption_with_new_text(doc):
    """Re-captioning must not push the caption text outside the anchor."""
    caption(doc, 1, text="旧题注")
    out = caption(doc, 1, text="新题注")
    assert out["recaption"] is True
    p = P(doc, 1)
    span = bookmark_span_text(p._p, ref_bms(p)[0])
    assert "新题注" in span, (
        f"rewritten caption text fell outside the _Ref span: {span!r}")
    assert "旧题注" not in span


# ---- 3. the caller's bookmark coexists -------------------------------------

def test_custom_bookmark_and_ref_bookmark_coexist(doc):
    out = caption(doc, 1, label="图", text="架构", bookmark="fig_vt")
    p = P(doc, 1)
    names = all_bm(p)
    assert "fig_vt" in names, "the caller's own bookmark must not be replaced"
    assert out["ref_bookmark"] in names
    assert names.count("fig_vt") == 1
    assert len(names) == 2, f"expected exactly two bookmarks, got {names}"
    assert len(set(bm_ids(p))) == 2, "the two bookmarks need distinct w:id"


def test_custom_bookmark_still_returns_the_number(doc):
    """The narrow anchor keeps its old meaning: REF fig_vt gives the counter."""
    caption(doc, 1, text="架构", bookmark="fig_vt")
    span = bookmark_span_text(P(doc, 1)._p, "fig_vt")
    assert span == "1", f"the custom narrow bookmark must still be the number: {span!r}"


def test_custom_bookmark_span_paragraph_plus_ref_bookmark(doc):
    caption(doc, 2, text="架构", bookmark="fig_wide", bookmark_span="paragraph")
    p = P(doc, 2)
    wide = bookmark_span_text(p._p, "fig_wide")
    ref = bookmark_span_text(p._p, ref_bms(p)[0])
    assert wide == ref == paragraph_text(p._p), (
        "both anchors may cover the caption at once; neither may shadow the other")


def test_custom_name_that_looks_like_a_ref_is_not_mistaken(doc):
    """_recaption must report the caller's name, not the _Ref anchor."""
    caption(doc, 1, text="架构", bookmark="cap_1")
    out = caption(doc, 1, text="改后题注")
    assert out["bookmark"] == "cap_1"
    assert out["ref_bookmark"] and out["ref_bookmark"] != "cap_1"


# ---- 4. SEQ 图 / SEQ 表 count independently --------------------------------

def test_figure_and_table_labels_count_independently(doc):
    f1 = caption(doc, 1, label="图", text="图一")
    t1 = caption(doc, 2, label="表", text="表一")
    f2 = caption(doc, 3, label="图", text="图二")
    t2 = caption(doc, 4, label="表", text="表二")
    assert [f1["number"], t1["number"], f2["number"], t2["number"]] == [1, 1, 2, 2]
    assert seq_counters(doc) == {"图": 2, "表": 2}


def test_table_label_writes_a_real_seq_table_field(doc):
    caption(doc, 1, label="表", text="系统各部件重量分布测试表")
    instrs = instr_texts(doc)
    assert r"SEQ 表 \* ARABIC" in instrs, f"missing the table counter: {instrs}"


def test_table_caption_carries_ref_bookmark_and_number(doc):
    out = caption(doc, 1, label="表", text="重量分布")
    p = P(doc, 1)
    assert out["number"] == 1
    assert "表" in paragraph_text(p._p)
    span = bookmark_span_text(p._p, ref_bms(p)[0])
    assert span == paragraph_text(p._p)


def test_each_label_gets_its_own_ref_bookmark(doc):
    a = caption(doc, 1, label="图", text="图一")["ref_bookmark"]
    b = caption(doc, 2, label="表", text="表一")["ref_bookmark"]
    c = caption(doc, 3, label="图", text="图二")["ref_bookmark"]
    assert len({a, b, c}) == 3


# ---- 5. guidance helper ----------------------------------------------------

def test_next_ref_name_ignores_non_ref_bookmarks(doc):
    from docx_crossref import _wrap_whole_paragraph
    from docx_utils import next_bm_id
    caption(doc, 1, text="架构")
    before = next_ref_name(doc)
    _wrap_whole_paragraph(P(doc, 5)._p, "chart_9999", next_bm_id(doc.element.body))
    assert next_ref_name(doc) == before, (
        "a bookmark called chart_9999 is not _Ref<digits> and must not move the "
        "counter")


def test_bm_name_is_still_the_plain_caption_name(doc):
    """The _Ref family is additive; it does not rename the caption bookmark."""
    out = caption(doc, 1, label="图", text="架构")
    assert out["bookmark"] == bm_name("图", 1)
