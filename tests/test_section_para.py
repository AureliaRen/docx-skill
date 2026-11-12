# -*- coding: utf-8 -*-
"""Tests for the paragraph-format vocabulary and multi-section page numbering.

Both features are structural and both have a silent failure mode, so the
assertions here look at the XML and at the *parts* rather than at
``Paragraph.text``:

* a paragraph-format key that never reaches ``w:pPr`` simply does nothing —
  the document still opens;
* a ``w:pgNumType`` written in the wrong ``CT_SectPr`` slot makes Word ignore
  the whole section setting;
* a PAGE field written into an inherited footer part edits a different
  section's footer, and looks perfect from the writer's side.

Every test is ``tmp_path`` based: no Office, no Node, no disk fixture, so the
suite still runs on CI.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Mm, Pt, Twips  # noqa: E402

import docx_create as dc  # noqa: E402
import docx_section as ds  # noqa: E402
from docx_crossref import all_bookmarks  # noqa: E402
from docx_numbering import current_numpr  # noqa: E402


def P(d, i):
    return d.paragraphs[i]


def tags(el):
    return [c.tag.split("}")[-1] for c in el]


def instr(part):
    return [t.text.strip() for t in part.element.iter(qn("w:instrText"))]


def runs_text(part):
    return "".join(t.text or "" for t in part.element.iter(qn("w:t")))


def footer_id(section, page_type="default"):
    return ds._ref_id(section._sectPr, "footer", page_type)


def footer_part(d, section, page_type="default"):
    return d.part.related_parts[footer_id(section, page_type)]


def create_cli(tmp_path, spec, name="spec.json"):
    """Run ``docx_create.py`` as a CLI; return (process, output path).

    A real subprocess rather than ``dc.main()``: the defect this covers
    surfaced as a bare traceback and exit ``2`` from *the command line*, and
    an in-process call would not show the exit code a user actually sees.
    """
    spec_path = tmp_path / name
    spec_path.write_text(json.dumps(spec, ensure_ascii=False),
                         encoding="utf-8")
    out = tmp_path / (name + ".docx")
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_create.py"),
         str(spec_path), str(out)],
        capture_output=True, check=False,
        env=dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8"))
    return proc, out


def _classify_runs(part):
    """Split a header/footer's runs into ``(field_runs, text_runs)``。

    按文档顺序走一遍，用 ``w:fldChar`` 的状态机判定归属——不能只看「有没有
    ``fldChar``/``instrText`` 子元素」：一个**域结果 run** 只带 ``<w:t>``，
    看上去和 ``Page `` 这类文字 run 一模一样，但它位于 ``separate`` 与 ``end``
    之间，属于域。漏判它会让「文字 run 有几个」直接翻倍（``Page ``/`` of ``
    两个真文字 run + 两个 ``1`` 结果 run = 4）。
    """
    field, text = [], []
    depth = 0
    for run in part.element.iter(qn("w:r")):
        fld = run.find(qn("w:fldChar"))
        if fld is not None:
            kind = fld.get(qn("w:fldCharType"))
            field.append(run)
            if kind == "begin":
                depth += 1
            elif kind == "end":
                depth = max(0, depth - 1)
            continue
        if run.find(qn("w:instrText")) is not None:
            field.append(run)
            continue
        # `w:fldSimple` 的结果 run 是它的子元素，不在段落层级上
        in_simple = any(a.tag == qn("w:fldSimple")
                        for a in run.iterancestors())
        (field if depth > 0 or in_simple else text).append(run)
    return field, text


def field_rprs(part):
    """Every ``<w:rPr>`` that decorates a *field* run in a header/footer.

    The plain-text runs around the fields get their formatting from
    ``python-docx``; only the field runs take the ``_rpr()`` path under test,
    so the assertion has to isolate them: a run that carries
    ``<w:fldChar>``/``<w:instrText>``, or a field *result* run sitting between
    ``separate`` and ``end``.
    """
    runs, _ = _classify_runs(part)
    return [r.find(qn("w:rPr")) for r in runs if r.find(qn("w:rPr")) is not None]


def text_rprs(part):
    """Every ``<w:rPr>`` that decorates a *plain-text* run in a header/footer.

    The mirror image of :func:`field_rprs`: a run outside every field carries
    one of the literal strings around them (``Page `` / `` of ``, or a
    caller's own ``prefix``/``suffix``), so it is the one written through
    ``python-docx``'s ``run.font.name``.  Runs with no ``w:rPr`` at all cannot
    carry a font and are left out.
    """
    _, runs = _classify_runs(part)
    return [r.find(qn("w:rPr")) for r in runs if r.find(qn("w:rPr")) is not None]


def three_sections():
    """cover (no number) / toc (i, ii) / body (restart at 1)."""
    d = Document()
    d.add_paragraph("封面")
    dc.add_block(d, {"type": "section_break", "start_type": "new_page",
                     "footer_page_numbers": {"fmt": "lowerRoman", "start": 1,
                                             "template": "page",
                                             "align": "center"}})
    d.add_paragraph("目录")
    dc.add_block(d, {"type": "section_break", "start_type": "new_page",
                     "footer_page_numbers": {"fmt": "decimal", "start": 1,
                                             "template": "page"}})
    d.add_paragraph("正文")
    return d


# ---------------------------------------------------------------------------
# paragraph-level format keys
# ---------------------------------------------------------------------------
def test_indent_spacing_and_alignment_reach_ppr():
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "x",
                     "indent_left_mm": 10, "indent_right_mm": 5,
                     "space_before_pt": 6, "space_after_pt": 12,
                     "alignment": "justify"})
    pPr = P(d, 0)._p.pPr
    ind = pPr.find(qn("w:ind"))
    assert int(ind.get(qn("w:left"))) == Mm(10).twips
    assert int(ind.get(qn("w:right"))) == Mm(5).twips
    sp = pPr.find(qn("w:spacing"))
    assert int(sp.get(qn("w:before"))) == Pt(6).twips
    assert int(sp.get(qn("w:after"))) == Pt(12).twips
    assert pPr.find(qn("w:jc")).get(qn("w:val")) == "both"


def test_line_spacing_uses_the_multiple_rule():
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "x", "line_spacing": 1.5})
    sp = P(d, 0)._p.pPr.find(qn("w:spacing"))
    assert sp.get(qn("w:line")) == "360"
    assert sp.get(qn("w:lineRule")) == "auto"


def test_first_line_indent_chars_writes_both_units():
    """Word's first-line indent is a *character* count with a twips twin."""
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "x",
                     "first_line_indent_chars": 2, "char_width_pt": 12})
    ind = P(d, 0)._p.pPr.find(qn("w:ind"))
    assert ind.get(qn("w:firstLineChars")) == "200"
    assert ind.get(qn("w:firstLine")) == "480"       # 2 chars x 12pt x 20
    assert P(d, 0).paragraph_format.first_line_indent == Twips(480)


def test_first_line_indent_falls_back_to_the_run_size():
    d = Document()
    para = d.add_paragraph("x")
    para.add_run("y").font.size = Pt(14)
    assert dc._effective_pt(para, {}) == 14.0
    assert dc._effective_pt(para, {"char_width_pt": 10}) == 10.0


def test_keep_with_next_and_page_break_before():
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "x",
                     "keep_with_next": True, "page_break_before": True})
    pPr = P(d, 0)._p.pPr
    assert pPr.find(qn("w:keepNext")) is not None
    assert pPr.find(qn("w:pageBreakBefore")) is not None


def test_bookmark_wraps_the_whole_paragraph():
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "正文",
                     "alignment": "left", "bookmark": "sec_intro"})
    p = P(d, 0)._p
    kids = tags(p)
    assert kids[0] == "pPr"
    assert kids[1] == "bookmarkStart"
    assert kids[-1] == "bookmarkEnd"
    start, end = p[1], p[-1]
    assert start.get(qn("w:name")) == "sec_intro"
    assert start.get(qn("w:id")) == end.get(qn("w:id"))
    assert "sec_intro" in all_bookmarks(d)


def test_bookmark_leads_when_there_is_no_ppr():
    """A paragraph with no properties still gets a well-formed bookmark."""
    d = Document()
    dc.add_block(d, {"type": "paragraph", "text": "正文", "bookmark": "bare"})
    p = P(d, 0)._p
    assert tags(p) == ["bookmarkStart", "r", "bookmarkEnd"]
    assert p[0].get(qn("w:id")) == p[-1].get(qn("w:id"))


def test_bookmark_on_a_multi_paragraph_block_covers_only_the_first():
    d = Document()
    dc.add_block(d, {"type": "bullet_list", "items": ["a", "b", "c"],
                     "bookmark": "list_start", "alignment": "left"})
    assert "list_start" in all_bookmarks(d)
    assert [P(d, i)._p.find(qn("w:bookmarkStart")) is not None
            for i in range(3)] == [True, False, False]


def test_every_paragraph_block_shares_the_format_vocabulary():
    d = Document()
    keys = {"alignment": "center"}
    made = []
    made += dc.add_block(d, dict({"type": "heading", "text": "H"}, **keys))
    made += dc.add_block(d, dict({"type": "paragraph", "text": "p"}, **keys))
    made += dc.add_block(d, dict({"type": "bullet_list",
                                  "items": ["b1"]}, **keys))
    made += dc.add_block(d, dict({"type": "numbered_list",
                                  "items": ["n1"]}, **keys))
    made += dc.add_block(d, dict({"type": "page_break"}, **keys))
    made += dc.add_block(d, dict({"type": "toc"}, **keys))
    made += dc.add_block(d, dict({"type": "caption", "label": "图",
                                  "text": "c"}, **keys))
    assert len(made) == 7
    for para in made:
        jc = para._p.pPr.find(qn("w:jc"))
        assert jc is not None and jc.get(qn("w:val")) == "center", para.text


def test_unknown_alignment_is_rejected():
    d = Document()
    with pytest.raises(ValueError) as e:
        dc.add_block(d, {"type": "paragraph", "text": "x",
                         "alignment": "middle"})
    assert "unknown alignment" in str(e.value)


def test_para_format_keys_are_all_honoured():
    """The declared vocabulary and the implemented one must not drift."""
    d = Document()
    block = {
        "type": "paragraph", "text": "x",
        "indent_left_mm": 1, "indent_right_mm": 2,
        "first_line_indent_chars": 2, "line_spacing": 1.5,
        "space_before_pt": 3, "space_after_pt": 4,
        "alignment": "center", "keep_with_next": True,
        "page_break_before": False, "bookmark": "k",
        # tab_stops belongs to the same vocabulary: it is written into
        # w:pPr/w:tabs for a paragraph's own tab stops (TOC dot leaders).
        "tab_stops": [{"pos_mm": 155, "align": "right", "leader": "dot"}],
        # bottom_border is the exam answer rule (w:pPr/w:pBdr/w:bottom).
        "bottom_border": True,
    }
    para = dc.add_block(d, block)[0]
    applied = dc._apply_para_format(para, block)
    assert set(applied) == set(dc.PARA_FORMAT_KEYS)
    assert applied["line_spacing"] == 1.5
    assert applied["bookmark"] == "k"
    assert applied["tab_stops"][0]["leader"] == "dot"
    assert applied["bottom_border"]["bottom"]["style"] == "single"


# ---------------------------------------------------------------------------
# sections and dual page numbering
# ---------------------------------------------------------------------------
def test_three_sections_with_dual_page_numbering():
    d = three_sections()
    assert len(d.sections) == 3
    assert ds.get_pg_num_type(d.sections[0]) is None
    assert ds.get_pg_num_type(d.sections[1]) == {"fmt": "lowerRoman",
                                                 "start": "1"}
    assert ds.get_pg_num_type(d.sections[2]) == {"fmt": "decimal", "start": "1"}


def test_pgnumtype_is_ordered_between_pgmar_and_cols():
    """CT_SectPr is xsd:sequence; a misplaced pgNumType is silently ignored."""
    d = Document()
    dc.add_block(d, {"type": "section_break", "cols": 2,
                     "footer_page_numbers": {"fmt": "lowerRoman"}})
    order = tags(d.sections[-1]._sectPr)
    assert order.index("pgMar") < order.index("pgNumType") < order.index("cols")


def test_first_sections_own_the_paragraph_sectpr_last_one_owns_the_body():
    d = three_sections()
    assert d.sections[0]._sectPr.getparent().tag == qn("w:pPr")
    assert d.sections[1]._sectPr.getparent().tag == qn("w:pPr")
    assert d.sections[2]._sectPr.getparent().tag == qn("w:body")
    info = ds.cmd_list(d)["section_detail"]
    assert [i["sectPr_location"] for i in info] == \
        ["paragraph", "paragraph", "body"]
    assert [i["page_number"] for i in info] == [
        None, {"fmt": "lowerRoman", "start": "1"},
        {"fmt": "decimal", "start": "1"}]


def test_section_break_reuses_the_last_real_paragraph():
    """No stray empty carrier paragraph, so the cover cannot overflow."""
    d = Document()
    d.add_paragraph("封面一")
    d.add_paragraph("封面二")
    moved = dc.add_block(d, {"type": "section_break"})
    assert [p.text for p in d.paragraphs] == ["封面一", "封面二"]
    assert d.sections[0]._sectPr.getparent().tag == qn("w:pPr")
    pPr = d.paragraphs[1]._p.pPr
    assert pPr.find(qn("w:sectPr")) is not None
    assert ds._relocate_sectpr(d, d.sections[-1]) is False   # already moved


def test_section_break_keeps_the_carrier_when_a_table_ends_the_section():
    """A table at the end of a section must not be pushed into the next one."""
    d = Document()
    d.add_paragraph("引言")
    d.add_table(rows=1, cols=1).cell(0, 0).text = "表内"
    dc.add_block(d, {"type": "section_break"})
    assert tags(d.element.body) == ["p", "tbl", "p", "sectPr"]
    assert tags(d.element.body)[-2] == "p"          # empty carrier retained
    assert d.sections[0]._sectPr.getparent().tag == qn("w:pPr")


def test_each_section_gets_its_own_footer_part():
    d = three_sections()
    ids = [footer_id(s) for s in d.sections]
    assert ids[0] is None                     # cover: no footer at all
    assert ids[1] and ids[2] and ids[1] != ids[2]
    cover_refs = ds._hf_refs(d.sections[0]._sectPr)
    assert cover_refs == []


def test_page_field_lands_in_that_sections_own_footer():
    d = three_sections()
    toc = footer_part(d, d.sections[1])
    body = footer_part(d, d.sections[2])
    assert toc is not body
    assert instr(toc) == ["PAGE"]
    assert runs_text(toc) == "i"              # cached roman result
    assert runs_text(body) == "1"             # cached decimal result
    assert "第" not in runs_text(toc)


def test_footer_text_and_page_field_accumulate_not_overwrite():
    """own_hdrftr must keep a definition this section already owns."""
    d = three_sections()
    sec = d.sections[1]
    before = footer_id(sec)
    ds.set_hdrftr_text(sec, "第 1 章 绪论")
    assert footer_id(sec) == before           # same private part reused
    ds.add_page_number_field(sec, template="page")
    part = footer_part(d, sec)
    assert "第 1 章 绪论" in runs_text(part)
    assert instr(part) == ["PAGE"]


def test_new_section_does_not_inherit_the_previous_footer_by_accident():
    """add_section() clones the running sectPr, so the rId is shared at first."""
    d = Document()
    ds.set_hdrftr_text(d.sections[0], "封面页脚", kind="footer")
    dc.add_block(d, {"type": "section_break",
                     "header": "第 1 章", "footer_page_numbers":
                     {"fmt": "decimal", "start": 1, "template": "page"}})
    assert len(d.sections) == 2
    assert footer_id(d.sections[0]) != footer_id(d.sections[1])
    assert "封面页脚" in runs_text(footer_part(d, d.sections[0]))
    assert runs_text(footer_part(d, d.sections[1])) == "1"


def test_no_footer_blanks_an_inherited_footer():
    d = Document()
    ds.set_hdrftr_text(d.sections[0], "封面页脚", kind="footer")
    dc.add_block(d, {"type": "section_break", "no_footer": True})
    new_id = footer_id(d.sections[-1])
    assert new_id and new_id != footer_id(d.sections[0])
    assert runs_text(footer_part(d, d.sections[-1])) == ""


def test_apply_to_previous_configures_the_closed_section():
    d = Document()
    d.add_paragraph("封面")
    dc.add_block(d, {"type": "section_break", "apply_to": "previous",
                     "footer_page_numbers": {"fmt": "upperRoman", "start": 1,
                                             "template": "page"}})
    assert ds.get_pg_num_type(d.sections[0]) == {"fmt": "upperRoman",
                                                 "start": "1"}
    assert ds.get_pg_num_type(d.sections[1]) is None
    assert runs_text(footer_part(d, d.sections[0])) == "I"


def test_apply_to_rejects_an_unknown_side():
    d = Document()
    d.add_paragraph("x")
    with pytest.raises(ValueError) as e:
        dc.add_block(d, {"type": "section_break", "apply_to": "sideways"})
    assert "unknown apply_to" in str(e.value)


def test_header_is_per_section_and_cover_stays_bare():
    d = Document()
    d.add_paragraph("封面")
    dc.add_block(d, {"type": "section_break", "header": "参赛作品",
                     "header_align": "center"})
    assert ds._ref_id(d.sections[0]._sectPr, "header", "default") is None
    hdr_id = ds._ref_id(d.sections[1]._sectPr, "header", "default")
    assert hdr_id
    assert runs_text(d.part.related_parts[hdr_id]) == "参赛作品"


def test_first_page_footer_is_a_separate_slot():
    d = Document()
    d.add_paragraph("封面")
    dc.add_block(d, {"type": "section_break", "different_first_page": True,
                     "footer_page_numbers": {"fmt": "lowerRoman", "start": 1,
                                             "template": "page",
                                             "page_type": "first"}})
    sec = d.sections[-1]
    assert sec.different_first_page_header_footer is True
    assert sec._sectPr.find(qn("w:titlePg")) is not None
    assert footer_id(sec, "first") != footer_id(sec, "default")
    assert runs_text(footer_part(d, sec, "first")) == "i"


def test_multiline_footer_and_shrinking_back_to_one_paragraph():
    d = Document()
    d.add_paragraph("x")
    dc.add_block(d, {"type": "section_break"})
    sec = d.sections[-1]
    ds.set_hdrftr_text(sec, "第一行\n第二行\n第三行", kind="footer")
    assert [p.text for p in sec.footer.paragraphs] == \
        ["第一行", "第二行", "第三行"]
    ds.set_hdrftr_text(sec, "只有一行", kind="footer")
    assert [p.text for p in sec.footer.paragraphs] == ["只有一行"]


def test_columns_are_written_and_ordered():
    d = Document()
    dc.add_block(d, {"type": "section_break", "cols": 2})
    cols = d.sections[-1]._sectPr.find(qn("w:cols"))
    assert cols.get(qn("w:num")) == "2"
    assert cols.get(qn("w:space")) == "425"


def test_unknown_start_type_is_rejected():
    d = Document()
    with pytest.raises(SystemExit) as e:
        dc.add_block(d, {"type": "section_break", "start_type": "sideways"})
    assert "unknown start_type" in str(e.value)


def test_unknown_page_number_format_is_rejected():
    d = Document()
    with pytest.raises(SystemExit) as e:
        dc.add_block(d, {"type": "section_break",
                         "footer_page_numbers": {"fmt": "klingon"}})
    assert "unknown page number fmt" in str(e.value)


def test_cached_page_renders_every_supported_format():
    assert ds.cached_page(None, None) == "1"
    assert ds.cached_page("decimal", 4) == "4"
    assert ds.cached_page("lowerRoman", 1) == "i"
    assert ds.cached_page("lowerRoman", 4) == "iv"
    assert ds.cached_page("upperRoman", 9) == "IX"
    assert ds.cached_page("lowerLetter", 1) == "a"
    assert ds.cached_page("lowerLetter", 28) == "ab"
    assert ds.cached_page("upperLetter", 3) == "C"
    assert ds.cached_page("decimalZero", 3) == "03"
    assert ds.roman(0) == "0"
    assert ds.letters(0) == "0"


def test_start_type_round_trips_into_the_section_report():
    d = Document()
    d.add_paragraph("a")
    dc.add_block(d, {"type": "section_break", "start_type": "continuous"})
    d.add_paragraph("b")
    dc.add_block(d, {"type": "section_break", "start_type": "odd_page"})
    assert [i["start_type"] for i in ds.cmd_list(d)["section_detail"]] == \
        ["new_page", "continuous", "odd_page"]


def test_sections_survive_save_and_reload(tmp_path):
    d = three_sections()
    out = tmp_path / "sections.docx"
    d.save(out)
    again = Document(out)
    assert len(again.sections) == 3
    assert ds.get_pg_num_type(again.sections[1]) == {"fmt": "lowerRoman",
                                                     "start": "1"}
    assert ds.get_pg_num_type(again.sections[2]) == {"fmt": "decimal",
                                                     "start": "1"}
    ids = [footer_id(s) for s in again.sections]
    assert ids[0] is None and ids[1] != ids[2]
    assert runs_text(footer_part(again, again.sections[1])) == "i"
    assert again.sections[0]._sectPr.getparent().tag == qn("w:pPr")


# ---------------------------------------------------------------------------
# integration: the numbered block and the cross-reference block
# ---------------------------------------------------------------------------
def test_numbered_list_bookmark_prefix_names_each_item():
    d = Document()
    dc.add_block(d, {"type": "numbered_list",
                     "items": ["盛万兴", "Krausman", "钱学森"],
                     "num_text": "[%1]", "bookmark_prefix": "ref_"})
    assert {"ref_1", "ref_2", "ref_3"} <= all_bookmarks(d)
    assert current_numpr(P(d, 0)) is not None
    pPr = P(d, 0)._p.pPr
    numPr = pPr.find(qn("w:numPr"))
    assert tags(numPr) == ["ilvl", "numId"]


def test_numbered_list_without_prefix_has_no_bookmarks():
    d = Document()
    dc.add_block(d, {"type": "numbered_list", "items": ["a", "b"]})
    assert all_bookmarks(d) == set()
    assert current_numpr(P(d, 1)) is not None


def test_crossref_display_switches_are_passed_through():
    plain = dc._crossref_field_args({"target": "图_1"}, 3, "图")
    assert plain.superscript is False
    assert plain.vert_align == "normal"
    assert plain.hyperlink is True
    sup = dc._crossref_field_args({"target": "ref_1", "superscript": True},
                                  3, "")
    assert sup.superscript is True and sup.vert_align == "superscript"
    sub = dc._crossref_field_args({"target": "ref_1",
                                   "vert_align": "subscript"}, 3, "")
    assert sub.vert_align == "subscript" and sub.superscript is False
    off = dc._crossref_field_args({"target": "ref_1", "hyperlink": False,
                                   "cached": "9"}, 3, "")
    assert off.hyperlink is False and off.cached == "9"
    assert off.index == 3


def test_crossref_block_still_writes_a_ref_field():
    """The block must reach ``cmd_ref``; the switch list is cmd_ref's business."""
    d = Document()
    dc.add_block(d, {"type": "numbered_list", "items": ["A"],
                     "bookmark_prefix": "ref_"})
    dc.add_block(d, {"type": "paragraph", "text": "见"})
    dc.add_block(d, {"type": "crossref", "target": "ref_1",
                     "superscript": True, "cached": "1"})
    refs = [t.text.strip() for t in d.element.body.iter(qn("w:instrText"))
            if t.text and t.text.strip().startswith("REF ")]
    assert len(refs) == 1, refs
    assert refs[0].split()[1] == "ref_1"
    assert "1" in P(d, 1).text           # cached result is visible


# ---------------------------------------------------------------------------
# CLI end to end
# ---------------------------------------------------------------------------
def test_create_cli_writes_a_three_section_document(tmp_path):
    spec = {
        "page": {"width_mm": 210, "height_mm": 297,
                 "margins_mm": {"top": 25, "bottom": 25, "left": 30,
                                "right": 20}},
        "blocks": [
            {"type": "paragraph", "text": "封面", "alignment": "center"},
            {"type": "section_break",
             "footer_page_numbers": {"fmt": "lowerRoman", "start": 1,
                                     "template": "page"}},
            {"type": "paragraph", "text": "目录"},
            {"type": "section_break",
             "footer_page_numbers": {"fmt": "decimal", "start": 1,
                                     "template": "page_of"}},
            {"type": "paragraph", "text": "正文",
             "first_line_indent_chars": 2, "line_spacing": 1.5},
        ],
    }
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False),
                         encoding="utf-8")
    out = tmp_path / "out.docx"
    argv = sys.argv
    try:
        sys.argv = ["docx_create.py", str(spec_path), str(out)]
        assert dc.main() == 0
    finally:
        sys.argv = argv

    again = Document(out)
    assert len(again.sections) == 3
    assert ds.get_pg_num_type(again.sections[1]) == {"fmt": "lowerRoman",
                                                     "start": "1"}
    assert ds.get_pg_num_type(again.sections[2]) == {"fmt": "decimal",
                                                     "start": "1"}
    assert footer_id(again.sections[0]) is None
    body_footer = footer_part(again, again.sections[2])
    assert "".join(instr(body_footer)) == "PAGENUMPAGES"
    assert again.paragraphs[-1].paragraph_format.line_spacing == 1.5
    ind = again.paragraphs[-1]._p.pPr.find(qn("w:ind"))
    assert ind.get(qn("w:firstLineChars")) == "200"


def test_section_cli_pagenums_and_footer(tmp_path):
    d = three_sections()
    path = tmp_path / "cli.docx"
    d.save(path)

    for argv, expect in (
        (["docx_section.py", "pagenums", str(path), "--section", "1",
          "--fmt", "upperRoman", "--start", "3", "--template", "dash"],
         "upperRoman"),
        (["docx_section.py", "footer", str(path), "--section", "2",
          "--text", "第 2 章"], None),
        (["docx_section.py", "cols", str(path), "--section", "0",
          "--num", "2"], None),
        (["docx_section.py", "break", str(path), "--start-type", "continuous"],
         None),
        (["docx_section.py", "pagenum", str(path), "--section", "1",
          "--clear"], None),
    ):
        old = sys.argv
        try:
            sys.argv = argv
            assert ds.main() == 0, argv
        finally:
            sys.argv = old
        again = Document(path)
        if expect:
            assert ds.get_pg_num_type(again.sections[1])["fmt"] == expect

    again = Document(path)
    assert len(again.sections) == 4
    assert "第 2 章" in runs_text(footer_part(again, again.sections[2]))
    assert again.sections[0]._sectPr.find(qn("w:cols")).get(qn("w:num")) == "2"
    assert ds.get_pg_num_type(again.sections[1]) is None


def test_section_cli_rejects_a_bad_section_index(tmp_path):
    d = three_sections()
    path = tmp_path / "bad.docx"
    d.save(path)
    old = sys.argv
    try:
        sys.argv = ["docx_section.py", "footer", str(path),
                    "--section", "9", "--text", "x"]
        with pytest.raises(SystemExit) as e:
            ds.main()
        assert "no such section" in str(e.value)
    finally:
        sys.argv = old



# ---------------------------------------------------------------------------
# the page-number field's run properties
# ---------------------------------------------------------------------------
# ``footer_page_numbers`` accepts ``font`` / ``size_pt`` / ``bold`` (they are
# in ``docx_create.PAGE_NUMBER_KEYS``), and the field runs read them through
# ``docx_section._rpr``.  Writing an attribute needs a *qualified* name: a bare
# ``qn("ascii")`` raises ``ValueError: not enough values to unpack``, which the
# builder reports as exit ``2`` -- no document at all.  These tests pin the
# whole path: the CLI exit code, the four ``w:rFonts`` attributes, the
# half-point size, and the absence of properties nobody asked for.
#
# Pure python-docx: the assertions read the saved footer XML, so they hold on a
# runner with no Office and no ``MML2OMML.XSL``.  Nothing here touches OMML, so
# no ``@needs_omml`` gate is warranted -- and adding one would hide the test on
# exactly the machines where it can run.
def fpn_spec(**fpn):
    """A minimal spec whose page-number field carries *fpn*."""
    settings = {"template": "page_of", "fmt": "decimal", "start": 1,
                "align": "center"}
    settings.update(fpn)
    return {"page": {"width_mm": 210, "height_mm": 297,
                     "margins_mm": {"top": 25, "bottom": 25, "left": 20,
                                    "right": 20}},
            "footer_page_numbers": settings,
            "blocks": [{"type": "paragraph", "text": "Body."}]}


def test_page_number_field_font_writes_all_four_rfonts_attributes(tmp_path):
    """``font`` must reach ``w:ascii``/``hAnsi``/``eastAsia``/``cs``."""
    proc, out = create_cli(tmp_path, fpn_spec(font="Times New Roman"),
                           name="font.json")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 0, (
        "docx_create must accept a page-number font, not blow up in _rpr\n"
        "exit: %d\nstdout: %s\nstderr: %s"
        % (proc.returncode, proc.stdout.decode("utf-8", "replace"), stderr))
    assert "Traceback" not in stderr
    assert out.exists(), "exit 0 promises a document"

    d = Document(out)
    rprs = field_rprs(footer_part(d, d.sections[0]))
    assert rprs, "the page-number field runs carry no w:rPr at all"
    for rpr in rprs:
        rfonts = rpr.find(qn("w:rFonts"))
        assert rfonts is not None, "w:rPr is missing w:rFonts"
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            assert rfonts.get(qn(attr)) == "Times New Roman", (
                "w:rFonts/@%s is %r, expected the font the spec asked for"
                % (attr, rfonts.get(qn(attr))))


def test_page_number_field_size_is_halved_into_w_sz(tmp_path):
    """``size_pt`` is points; OOXML stores half-points."""
    proc, out = create_cli(tmp_path, fpn_spec(size_pt=10, bold=True),
                           name="size.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    assert out.exists()

    d = Document(out)
    rprs = field_rprs(footer_part(d, d.sections[0]))
    assert rprs
    for rpr in rprs:
        assert rpr.find(qn("w:sz")).get(qn("w:val")) == "20"
        assert rpr.find(qn("w:szCs")).get(qn("w:val")) == "20"
        assert rpr.find(qn("w:b")) is not None
        assert rpr.find(qn("w:bCs")) is not None

    # A fractional point value must round, not truncate.
    proc, out = create_cli(tmp_path, fpn_spec(size_pt=10.5), name="halfpt.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    d = Document(out)
    for rpr in field_rprs(footer_part(d, d.sections[0])):
        assert rpr.find(qn("w:sz")).get(qn("w:val")) == "21"


def test_page_number_field_does_not_invent_a_size_or_bold(tmp_path):
    """Only ``font`` was asked for, so ``w:sz``/``w:b`` must stay absent.

    ``_rpr`` builds one ``<w:rPr>`` per field run, so a default that leaked in
    here would silently restyle every page number in the document.
    """
    proc, out = create_cli(tmp_path, fpn_spec(font="Arial"), name="bare.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    d = Document(out)
    rprs = field_rprs(footer_part(d, d.sections[0]))
    assert rprs
    for rpr in rprs:
        assert tags(rpr) == ["rFonts"], (
            "asked for font only, got %r" % (tags(rpr),))
        assert rpr.find(qn("w:sz")) is None
        assert rpr.find(qn("w:b")) is None


def test_page_number_field_without_run_properties_has_no_rpr(tmp_path):
    """No ``font``/``size_pt``/``bold`` means no ``<w:rPr>`` on field runs."""
    proc, out = create_cli(tmp_path, fpn_spec(), name="plain.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    d = Document(out)
    part = footer_part(d, d.sections[0])
    assert instr(part) == ["PAGE", "NUMPAGES"]
    assert field_rprs(part) == []
    # 文字 run 也该干净：没要求字体就不许凭空长出一个 w:rFonts
    assert text_rprs(part) == []


# ---------------------------------------------------------------------------
# The *text* runs of the footer (``Page `` / `` of ``), i.e. the prefix/suffix
# around the fields, must carry the same four ``w:rFonts`` attributes as the
# field runs.  ``run.font.name = font`` only writes ``@ascii`` and ``@hAnsi``
# (``docx/text/font.py``: ``rPr.rFonts_ascii = value`` / ``rFonts_hAnsi``), so
# the ``eastAsia`` slot stayed on the theme font.  With the default English
# ``"Page "`` prefix that is invisible; a Chinese prefix (``第 X 页 共 Y 页``)
# then renders its CJK glyphs in the theme's East Asian font while the digits
# come from the requested font -- the exact mixed-script jump ``fix_fonts.py``
# exists to prevent.  It also made one footer self-inconsistent: four
# attributes on the field runs, two on the text runs.
# ---------------------------------------------------------------------------
def test_page_number_text_runs_carry_all_four_rfonts_attributes(tmp_path):
    """The ``Page `` / `` of `` text runs get ``ascii``/``hAnsi``/``eastAsia``/``cs``."""
    proc, out = create_cli(tmp_path, fpn_spec(font="Times New Roman"),
                           name="textfont.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    d = Document(out)
    part = footer_part(d, d.sections[0])
    rprs = text_rprs(part)
    assert len(rprs) == 2, (
        "expected one w:rPr per text run (prefix + suffix), got %d" % len(rprs))
    for rpr in rprs:
        rfonts = rpr.find(qn("w:rFonts"))
        assert rfonts is not None, "text run's w:rPr is missing w:rFonts"
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            assert rfonts.get(qn(attr)) == "Times New Roman", (
                "text run w:rFonts/@%s is %r, expected the font the spec "
                "asked for" % (attr, rfonts.get(qn(attr))))


def test_page_number_text_and_field_runs_agree_on_the_font(tmp_path):
    """Field run and text run in *one* footer resolve to the same four values.

    This is the self-consistency check: two writers (``_rpr`` vs
    ``run.font.name``) used to disagree inside a single paragraph.
    """
    proc, out = create_cli(tmp_path, fpn_spec(font="SimSun"), name="agree.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    d = Document(out)
    part = footer_part(d, d.sections[0])
    field_vals = [{a: rf.get(qn(a)) for a in
                   ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")}
                  for rf in (r.find(qn("w:rFonts")) for r in field_rprs(part))]
    text_vals = [{a: rf.get(qn(a)) for a in
                  ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")}
                 for rf in (r.find(qn("w:rFonts")) for r in text_rprs(part))]
    assert field_vals and text_vals
    assert set(map(tuple, (v.items() for v in field_vals))) == \
        set(map(tuple, (v.items() for v in text_vals)))


def test_chinese_page_prefix_gets_east_asia_font_on_the_text_run(tmp_path):
    """中文前缀 ``第 X 页 共 Y 页``：文字 run 的 ``eastAsia`` 必须是请求的字体。

    这是本单最有说服力的一条——默认 ``"Page "`` 前缀是纯西文，``eastAsia``
    缺省也看不出来；换成中文前缀才暴露「中英混排字体跳变」。
    """
    spec = fpn_spec(font="SimSun", template="page_of",
                    prefix="第 ", suffix=" 页 共 ")
    proc, out = create_cli(tmp_path, spec, name="cjk.json")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 0, stderr
    assert "Traceback" not in stderr

    d = Document(out)
    part = footer_part(d, d.sections[0])
    rprs = text_rprs(part)
    assert rprs, "no text run carries a w:rPr"
    for rpr in rprs:
        rfonts = rpr.find(qn("w:rFonts"))
        assert rfonts is not None
        assert rfonts.get(qn("w:eastAsia")) == "SimSun", (
            "中文前缀的文字 run 没写 eastAsia，会退回主题东亚字体")
        assert rfonts.get(qn("w:cs")) == "SimSun"


def test_page_number_text_run_size_and_bold_keep_their_own_paths(tmp_path):
    """``size_pt`` / ``bold`` on text runs stay as they were (not part of this change).

    Guards the boundary: only the *font* handling was unified with ``_rpr``.
    A text run asked for a size still stores half-points, and ``bold=True``
    still writes ``<w:b/>``.
    """
    proc, out = create_cli(tmp_path, fpn_spec(font="Arial", size_pt=9,
                                              bold=True), name="textsize.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    d = Document(out)
    for rpr in text_rprs(footer_part(d, d.sections[0])):
        assert rpr.find(qn("w:sz")).get(qn("w:val")) == "18"
        assert rpr.find(qn("w:b")) is not None


# ===========================================================================
# 输入错误契约（P9-C）
#
# 本文件此前的每一条用例都假设「输入是一个能打开的 .docx」。这一节补的是
# **打不开的那种输入**：实测旧行为是 `Document(args.path)` 直接抛
# `PackageNotFoundError` / `IsADirectoryError`，于是命令行上看到十九行
# traceback + exit 1 —— 用户无法分辨「路径写错了」和「文档坏了」。
#
# 全仓口径（P8 定案）：`2` = 用法/输入错误 → stderr 单行 `error: ...`、
# stdout 空、零写盘；`1` 留给「输入合法但处理不了」。
# ===========================================================================
def run_section(*args):
    """以真子进程跑 ``docx_section.py``。

    必须走子进程：本单修的是**命令行上看到的**退出码与 stderr 行数，
    in-process 调 ``main()`` 拿不到解释器的退出码，也看不到 traceback。
    """
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_section.py"), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", check=False,
        env=dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1"))


def one_error_line(stderr):
    """stderr 必须**恰好一行**（``error: ...``），没有 traceback 前导。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


# (子命令, 追加参数) —— `list` 只读、`pagenum` / `cols` 会 `doc.save()`，
# 后者是「零写盘」的关键（校验若落到 save 之后，就会留下半截输出文件）。
_SECTION_MODES = [
    ("list", ()),
    ("pagenum", ("--fmt", "decimal")),
    ("cols", ("--num", "2")),
]


@pytest.mark.parametrize("cmd,extra", _SECTION_MODES,
                         ids=[m[0] for m in _SECTION_MODES])
def test_missing_path_is_a_usage_error(tmp_path, cmd, extra):
    """路径不存在 → exit 2 + 单行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such.docx"
    assert not missing.exists()

    cp = run_section(cmd, missing, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line.startswith("error: ")
    assert "no such file" in line
    assert str(missing) in line
    assert cp.stdout == ""


@pytest.mark.parametrize("cmd,extra", _SECTION_MODES,
                         ids=[m[0] for m in _SECTION_MODES])
def test_directory_path_is_a_usage_error(tmp_path, cmd, extra):
    """路径是目录 → exit 2 + 单行 ``not a file``。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()

    cp = run_section(cmd, adir, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == f"error: not a file: {adir}"
    assert cp.stdout == ""


@pytest.mark.parametrize("cmd,extra", _SECTION_MODES,
                         ids=[m[0] for m in _SECTION_MODES])
def test_non_zip_file_is_a_usage_error(tmp_path, cmd, extra):
    """存在的普通文件但不是 zip → exit 2 + 单行 ``not a ZIP package``。"""
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("i am not a zip package", encoding="utf-8")

    cp = run_section(cmd, bogus, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == f"error: not a ZIP package: {bogus}"
    assert cp.stdout == ""


def test_non_zip_input_leaves_the_output_path_untouched(tmp_path):
    """``-o out.docx`` + 坏输入 → **不生成** out.docx（零写盘）。"""
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("i am not a zip package", encoding="utf-8")
    out = tmp_path / "out.docx"

    cp = run_section("pagenum", bogus, "-o", out, "--fmt", "decimal")
    assert cp.returncode == 2, cp.stderr
    assert cp.stdout == ""
    assert not out.exists()


def test_valid_document_list_is_unchanged(tmp_path):
    """合法文档 ``list`` 仍是 exit 0 + 纯 JSON（加校验没把成功路径改坏）。"""
    src = tmp_path / "good.docx"
    d = Document()
    d.add_paragraph("hello")
    d.save(str(src))

    cp = run_section("list", src)
    assert cp.returncode == 0, cp.stderr
    assert cp.stderr == ""
    payload = json.loads(cp.stdout)
    assert payload["ok"] is True
    assert payload["sections"] == 1
    assert len(payload["section_detail"]) == 1
