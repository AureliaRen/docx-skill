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
    dc.add_block(d, {"type": "section_break", "header": "冯如杯参赛作品",
                     "header_align": "center"})
    assert ds._ref_id(d.sections[0]._sectPr, "header", "default") is None
    hdr_id = ds._ref_id(d.sections[1]._sectPr, "header", "default")
    assert hdr_id
    assert runs_text(d.part.related_parts[hdr_id]) == "冯如杯参赛作品"


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
