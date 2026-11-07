# -*- coding: utf-8 -*-
"""Caption paragraph style: the filter Word's table of figures sorts by.

Why this file exists
--------------------
"Insert Table of Figures" collects paragraphs **by style**, not by the presence
of a SEQ field.  A caption that carries a counter but no style is invisible to
that command, so the entry never appears — the second half of the operator's
report.

``cmd_caption --pstyle <styleId>`` writes ``<w:pStyle w:val="..."/>`` into the
caption's ``w:pPr``.  Two details decide whether Word honours it:

1. ``CT_PPr`` is an ``xsd:sequence`` whose **first** member is ``pStyle``.  A
   ``pStyle`` written anywhere else is dropped by Word without a warning, and
   the paragraph silently keeps its previous style.
2. ``cmd_caption`` calls ``Paragraph.clear()``, documented as "keeps
   ``<w:pPr>``, drops content".  A style written *before* that call has to
   survive it; the implementation writes it after, and this file pins that.

The style id is a parameter, never a constant: the reference document
(``work/src.docx``) calls its caption style ``28`` while the *name* behind that
id is 图表名, and numbering differs per document and per Word locale.
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document                                   # noqa: E402
from docx.oxml.ns import qn                                 # noqa: E402

from docx_crossref import (                                 # noqa: E402
    bookmark_span_text, cmd_caption, paragraph_text, set_paragraph_style,
)
from docx_utils import PPR_ORDER                            # noqa: E402


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
    return cmd_caption(d, Args(index=index, label=label, text=text, **kw))


def pstyle_val(p):
    """First ``w:pStyle/@w:val`` of paragraph *p*, or None."""
    for el in p._p.iter(qn("w:pStyle")):
        return el.get(qn("w:val"))
    return None


def child_tags(el) -> list:
    return [c.tag.split("}")[1] for c in el]


# ---- 1. the style is written, and first -------------------------------

def test_pstyle_is_written_on_the_caption(doc):
    out = caption(doc, 1, text="架构", pstyle="28")
    assert out["pstyle"] == "28"
    assert pstyle_val(P(doc, 1)) == "28"


def test_pstyle_is_the_first_child_of_ppr(doc):
    """CT_PPr is a sequence; pStyle anywhere but first is dropped by Word."""
    caption(doc, 1, text="架构", pstyle="28")
    pPr = P(doc, 1)._p.find(qn("w:pPr"))
    assert pPr is not None
    tags = child_tags(pPr)
    assert tags[0] == "pStyle", f"pStyle must lead the pPr sequence, got {tags}"


def test_pstyle_lands_before_other_paragraph_properties(doc):
    """Written next to real pPr content, it still sorts to the front."""
    from docx.oxml import OxmlElement
    para = P(doc, 2)
    pPr = para._p.get_or_add_pPr()
    for name in ("spacing", "jc", "keepNext"):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:val"), "0" if name != "jc" else "center")
        pPr.append(el)
    caption(doc, 2, text="架构", pstyle="Caption")
    tags = child_tags(para._p.find(qn("w:pPr")))
    assert tags[0] == "pStyle", f"pStyle out of order: {tags}"
    assert tags.index("pStyle") < tags.index("spacing")
    assert tags.index("pStyle") < tags.index("jc")
    assert tags.index("pStyle") < tags.index("keepNext")


def test_pstyle_follows_the_schema_order_constant(doc):
    """The insertion point comes from PPR_ORDER, not from a lucky append."""
    assert PPR_ORDER[0] == "pStyle"
    caption(doc, 1, text="架构", pstyle="图表名")
    pPr = P(doc, 1)._p.find(qn("w:pPr"))
    assert child_tags(pPr)[0] == "pStyle"


# ---- 2. it survives cmd_caption's clear() ------------------------------

def test_pstyle_survives_the_paragraph_clear(doc):
    """clear() keeps pPr but drops content; the style must still be there."""
    assert paragraph_text(P(doc, 1)._p).strip(), "fixture paragraph has text"
    caption(doc, 1, text="架构", pstyle="28")
    assert pstyle_val(P(doc, 1)) == "28", (
        "the caption style was lost across Paragraph.clear()")
    assert "架构" in paragraph_text(P(doc, 1)._p)


def test_pstyle_replaces_an_earlier_style_instead_of_duplicating(doc):
    """xsd:sequence allows one pStyle; a duplicate breaks the schema."""
    caption(doc, 1, text="架构", pstyle="Caption")
    caption(doc, 2, text="流程", pstyle="28")
    set_paragraph_style(P(doc, 2)._p, "图表名")
    pPr = P(doc, 2)._p.find(qn("w:pPr"))
    assert child_tags(pPr).count("pStyle") == 1
    assert pstyle_val(P(doc, 2)) == "图表名"


def test_pstyle_is_applied_on_a_recaption_too(doc):
    """Re-running caption over an existing one must still fix the style."""
    caption(doc, 1, text="旧题注")
    out = caption(doc, 1, text="新题注", pstyle="28")
    assert out["recaption"] is True
    assert pstyle_val(P(doc, 1)) == "28"


def test_recaption_rewrites_the_style_when_it_changes(doc):
    """The second pass must win, not keep whatever the first pass set."""
    caption(doc, 1, text="一次", pstyle="Caption")
    assert pstyle_val(P(doc, 1)) == "Caption"
    caption(doc, 1, text="二次", pstyle="28")
    assert pstyle_val(P(doc, 1)) == "28"
    pPr = P(doc, 1)._p.find(qn("w:pPr"))
    assert child_tags(pPr).count("pStyle") == 1


def test_recaption_keeps_the_ref_span_around_the_new_text(doc):
    """Re-captioning must not push the rewritten text outside the _Ref span."""
    caption(doc, 1, text="旧文本", pstyle="Caption", ref_bookmark=True)
    caption(doc, 1, text="新文本", pstyle="Caption", ref_bookmark=True)
    p = P(doc, 1)
    ref = [b.get(qn("w:name")) for b in p._p.iter(qn("w:bookmarkStart"))
           if (b.get(qn("w:name")) or "").startswith("_Ref")]
    assert len(ref) == 1, f"a recaption must not add a second _Ref: {ref}"
    span = bookmark_span_text(p._p, ref[0])
    assert "新文本" in span, f"new caption text fell outside the span: {span!r}"
    assert "旧文本" not in span
    tags = child_tags(p._p)
    assert tags[-1] == "bookmarkEnd", f"span must still close last: {tags}"


# ---- 3. default off: existing callers are untouched -------------------

def test_no_pstyle_by_default(doc):
    out = caption(doc, 1, text="架构")
    assert out["pstyle"] is None
    assert pstyle_val(P(doc, 1)) is None, (
        "a caption without --pstyle must not gain a paragraph style")


def test_empty_pstyle_means_no_style(doc):
    out = caption(doc, 1, text="架构", pstyle="")
    assert out["pstyle"] in (None, "")
    assert pstyle_val(P(doc, 1)) is None


def test_existing_style_is_left_alone_without_the_flag(doc):
    from docx.oxml import OxmlElement
    para = P(doc, 3)
    pPr = para._p.get_or_add_pPr()
    el = OxmlElement("w:pStyle")
    el.set(qn("w:val"), "BodyText")
    pPr.insert(0, el)
    caption(doc, 3, text="架构")
    assert pstyle_val(P(doc, 3)) == "BodyText", (
        "without --pstyle the paragraph keeps whatever style it had")


# ---- 4. the style id is a parameter, not the reference document's 28 ---

def test_style_id_is_not_hardcoded_to_the_reference_document(doc):
    caption(doc, 1, text="架构", pstyle="Caption")
    assert pstyle_val(P(doc, 1)) == "Caption"
    assert pstyle_val(P(doc, 1)) != "28", (
        "28 is only the reference document's id for 图表名; other documents "
        "number their caption styles differently")


def test_style_id_survives_a_save_and_reload(doc, tmp_path):
    """The style must be in the part Word reads, not just in memory."""
    caption(doc, 1, text="架构", pstyle="28")
    out = tmp_path / "styled.docx"
    doc.save(out)
    xml = zipfile.ZipFile(out).read("word/document.xml").decode("utf-8")
    assert '<w:pStyle w:val="28"/>' in xml
    reloaded = Document(out)
    assert pstyle_val(reloaded.paragraphs[1]) == "28"


# ---- 5. the full caption shape, style + anchors together -------------

def test_caption_carries_style_anchor_and_counter_together(doc):
    caption(doc, 1, label="表", text="系统各部件重量分布测试表",
            pstyle="28", ref_bookmark=True)
    p = P(doc, 1)
    assert pstyle_val(p) == "28"
    names = [b.get(qn("w:name")) for b in p._p.iter(qn("w:bookmarkStart"))]
    assert any(n.startswith("_Ref") for n in names), f"no anchor: {names}"
    instrs = [(x.text or "").strip()
              for x in p._p.iter(qn("w:instrText"))]
    assert r"SEQ 表 \* ARABIC" in instrs, f"no table counter: {instrs}"
    assert "系统各部件重量分布测试表" in paragraph_text(p._p)


def test_child_order_is_ppr_then_anchor_then_runs(doc):
    """The whole paragraph layout the table of figures walks, in one assert."""
    caption(doc, 1, text="架构", pstyle="28", ref_bookmark=True)
    tags = child_tags(P(doc, 1)._p)
    assert tags[0] == "pPr"
    assert tags[1] == "bookmarkStart"
    assert tags.index("bookmarkStart") < tags.index("r")
    assert tags[-1] == "bookmarkEnd"
