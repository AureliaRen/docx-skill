# -*- coding: utf-8 -*-
"""Static table of contents: real paragraphs, real ``<w:tab/>``, dot leaders.

A TOC *field* is only text after Word recalculates it, so a file that has to
read correctly without an Office engine needs the lines written out.  These
tests pin the three things that make such a TOC usable:

* the entry line is ``title <TAB> page`` with an *actual* ``<w:tab/>`` element
  (python-docx's ``add_run`` puts ``\\t`` into ``w:t``, where Word draws
  nothing);
* the tab stop is right-aligned with a dot leader at the text-block width, so
  the page number is flush right and the dots fill the gap;
* rebuilding is idempotent — "edit the body, regenerate the TOC" must not
  accumulate a second copy.

No Office engine is involved anywhere: everything is asserted on the XML.
"""
from __future__ import annotations

import json
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

import docx_create as dc  # noqa: E402
import docx_toc as toc  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---- helpers -----------------------------------------------------------------

def toc_doc(tmp_path=None):
    """A document with one heading of each level and two body paragraphs."""
    doc = Document()
    doc.add_heading("第一章　静电场", level=1)
    doc.add_paragraph("电场是电荷周围存在的一种特殊物质。")
    doc.add_heading("1.1　库仑定律", level=2)
    doc.add_paragraph("真空中两个静止点电荷之间的作用力。")
    doc.add_heading("1.1.1　点电荷模型", level=3)
    doc.add_paragraph("把带电体抽象为一个几何点。")
    return doc


def entries_of(doc):
    """Paragraphs carrying one of our ``toc N`` styles."""
    return [p for p in doc.paragraphs if toc.is_toc_paragraph(p)]


def run_tabs(para):
    """Run-level ``<w:tab/>`` elements.

    ``w:tab`` is overloaded: inside ``w:tabs`` it *declares* a stop, inside a
    run it *is* the character.  Only the latter is a tab a reader sees.
    """
    return [t for t in para._p.iter(qn("w:tab"))
            if t.getparent().tag != qn("w:tabs")]


def all_text(para):
    return "".join(t.text or "" for t in para._p.iter(qn("w:t")))


def stops_of(para):
    return [(t.position.mm, int(t.alignment), int(t.leader))
            for t in para.paragraph_format.tab_stops]


# ---- style names -------------------------------------------------------------

def test_style_name_is_words_toc_number():
    assert toc.TOC_STYLE_NAME(1) == "toc 1"
    assert toc.TOC_STYLE_NAME(2) == "toc 2"
    assert toc.TOC_STYLE_NAME(3) == "toc 3"
    assert toc.TOC_STYLE_NAME("2") == "toc 2"


def test_style_name_rejects_nonsense():
    for bad in (0, 10, -1, "toc", None, "1.5"):
        with pytest.raises(ValueError):
            toc.TOC_STYLE_NAME(bad)


def test_created_style_uses_words_builtin_id():
    """Word only treats ``toc 1`` as its built-in if ``styleId`` is ``TOC1``.

    python-docx would otherwise mint ``toc1`` plus ``w:customStyle="1"``, and a
    round-trip through Word turns the lines into a custom style — after which
    "Update field" no longer owns them.
    """
    doc = Document()
    name = toc.ensure_toc_style(doc, 2)
    style = doc.styles[name]
    assert style.element.get(qn("w:styleId")) == "TOC2"
    assert style.element.get(qn("w:customStyle")) is None
    assert style.name == "toc 2"


def test_ensure_toc_style_is_idempotent():
    doc = Document()
    first = toc.ensure_toc_style(doc, 1)
    second = toc.ensure_toc_style(doc, 1)
    assert first == second == "toc 1"
    matching = [s for s in doc.styles
                if str(s.name).strip().lower() == "toc 1"]
    assert len(matching) == 1


# ---- outline -----------------------------------------------------------------

def test_collect_outline_levels_and_indices():
    doc = toc_doc()
    outline = toc.collect_outline(doc)
    assert [e["level"] for e in outline] == [1, 2, 3]
    assert [e["text"] for e in outline] == [
        "第一章　静电场", "1.1　库仑定律", "1.1.1　点电荷模型"]
    for entry in outline:
        assert doc.paragraphs[entry["para_index"]].text == entry["text"]


def test_collect_outline_only_returns_headings():
    doc = Document()
    doc.add_paragraph("正文一段，不是标题。")
    doc.add_heading("标题", level=1)
    doc.add_paragraph("正文两段。")
    outline = toc.collect_outline(doc)
    assert len(outline) == 1
    assert outline[0]["para_index"] == 1


def test_collect_outline_accepts_custom_cjk_style_with_outline_level():
    """A template's own ``中文标题`` style is a heading if it carries outlineLvl.

    This is the T1 contract leaking through: ``docx_read.heading_level`` is the
    only judge, so whatever it learns about a style is what the TOC sees.
    """
    doc = Document()
    style = doc.styles.add_style("中文标题", WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = doc.styles["Normal"]
    pPr = style.element.get_or_add_pPr()
    lvl = OxmlElement("w:outlineLvl")
    lvl.set(qn("w:val"), "1")          # Word stores 0-based; 1 -> level 2
    pPr.append(lvl)
    doc.add_paragraph("中文标题段落", style="中文标题")
    outline = toc.collect_outline(doc)
    assert len(outline) == 1
    assert outline[0]["level"] == 2


def test_collect_outline_accepts_direct_outline_level():
    doc = Document()
    para = doc.add_paragraph("直接标注的大纲级别")
    pPr = para._p.get_or_add_pPr()
    lvl = OxmlElement("w:outlineLvl")
    lvl.set(qn("w:val"), "0")
    pPr.append(lvl)
    outline = toc.collect_outline(doc)
    assert [e["level"] for e in outline] == [1]


def test_collect_outline_skips_a_previous_toc():
    """Rebuilding must not list the old TOC's own lines as headings."""
    doc = toc_doc()
    toc.build_static_toc(doc, title="目录")
    outline = toc.collect_outline(doc)
    assert [e["text"] for e in outline] == [
        "第一章　静电场", "1.1　库仑定律", "1.1.1　点电荷模型"]


# ---- entry lines -------------------------------------------------------------

def test_one_two_three_levels_are_written():
    doc = toc_doc()
    report = toc.build_static_toc(doc)
    assert report["entries"] == 3
    assert report["levels"] == [1, 2, 3]
    assert report["styles"] == ["toc 1", "toc 2", "toc 3"]
    lines = entries_of(doc)
    assert [p.style.name for p in lines] == ["toc 1", "toc 2", "toc 3"]
    assert [all_text(p) for p in lines] == [
        "第一章　静电场", "1.1　库仑定律", "1.1.1　点电荷模型"]


def test_levels_filter_selects_a_subset():
    doc = toc_doc()
    report = toc.build_static_toc(doc, levels=(1, 2))
    assert report["entries"] == 2
    assert report["levels"] == [1, 2]
    assert [p.style.name for p in entries_of(doc)] == ["toc 1", "toc 2"]


def test_tab_is_a_real_element_not_a_w_t_character():
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["1", "1", "2"])
    for line in entries_of(doc):
        assert len(run_tabs(line)) == 1, "expected one run-level <w:tab/>"
        assert "\t" not in all_text(line), "tab leaked into w:t"
    first = entries_of(doc)[0]
    assert all_text(first) == "第一章　静电场1"


def test_tab_splits_title_from_page_number():
    """The page number comes *after* the tab element, not before it."""
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["7", "8", "9"])
    para = entries_of(doc)[0]._p
    sequence = []
    for run in para.iter(qn("w:r")):
        for child in run:
            if child.tag == qn("w:tab"):
                sequence.append("TAB")
            elif child.tag == qn("w:t"):
                sequence.append(child.text or "")
    assert sequence == ["第一章　静电场", "TAB", "7"], sequence


def test_tab_stop_is_right_aligned_with_dot_leader():
    doc = toc_doc()
    toc.build_static_toc(doc)
    for line in entries_of(doc):
        stops = stops_of(line)
        assert len(stops) == 1, stops
        _pos, align, leader = stops[0]
        assert align == 2, "expected WD_TAB_ALIGNMENT.RIGHT"
        assert leader == 1, "expected WD_TAB_LEADER.DOTS"


def test_tab_stops_land_in_ppr_tabs():
    doc = toc_doc()
    toc.build_static_toc(doc, tab_pos_mm=155)
    tabs = entries_of(doc)[0]._p.find(qn("w:pPr")).find(qn("w:tabs"))
    assert tabs is not None
    assert tabs[0].get(qn("w:leader")) == "dot"
    assert tabs[0].get(qn("w:val")) == "right"


# ---- tab position = text block width ----------------------------------------

def test_tab_position_uses_content_width_mm_when_available(monkeypatch):
    """A maths paper's text block is 170 mm; the helper must drive the stop."""
    monkeypatch.setattr(dc, "content_width_mm", lambda section: 170.0,
                        raising=False)
    doc = toc_doc()
    report = toc.build_static_toc(doc)
    assert report["tab_pos_mm"] == pytest.approx(170.0)
    assert report["tab_pos_source"] == "content_width_mm"
    for line in entries_of(doc):
        assert stops_of(line)[0][0] == pytest.approx(170.0, abs=0.1)


def test_tab_position_falls_back_to_160_without_the_helper(monkeypatch):
    """Everything else in the house style is 160 mm; that is the fallback."""
    monkeypatch.delattr(dc, "content_width_mm", raising=False)
    doc = toc_doc()
    report = toc.build_static_toc(doc)
    assert report["tab_pos_mm"] == pytest.approx(160.0)
    assert report["tab_pos_source"] == "fallback"
    for line in entries_of(doc):
        assert stops_of(line)[0][0] == pytest.approx(160.0, abs=0.1)


def test_tab_position_argument_wins_over_the_helper(monkeypatch):
    monkeypatch.setattr(dc, "content_width_mm", lambda section: 170.0,
                        raising=False)
    doc = toc_doc()
    report = toc.build_static_toc(doc, tab_pos_mm=155)
    assert report["tab_pos_mm"] == pytest.approx(155.0)
    assert report["tab_pos_source"] == "tab_pos_mm"
    assert stops_of(entries_of(doc)[0])[0][0] == pytest.approx(155.0, abs=0.1)


def test_unusable_content_width_mm_is_ignored(monkeypatch):
    """A helper that returns None/0 must not put the stop at mm 0."""
    monkeypatch.setattr(dc, "content_width_mm", lambda section: None,
                        raising=False)
    doc = toc_doc()
    report = toc.build_static_toc(doc)
    assert report["tab_pos_mm"] == pytest.approx(160.0)
    assert report["tab_pos_source"] == "fallback"


def test_content_width_mm_signature_is_not_our_choice(monkeypatch):
    """T1 owns that helper's signature, so every plausible shape must work.

    Taking ``section``, taking ``doc``, or taking nothing at all are all
    reasonable reads of "content_width_mm"; a mismatch has to fall through to
    the next shape instead of aborting the lookup.
    """
    doc = toc_doc()
    for helper, expected in (
            (lambda section: 170.0, 170.0),
            (lambda document: 165.0, 165.0),
            (lambda: 158.0, 158.0),
    ):
        monkeypatch.setattr(dc, "content_width_mm", helper, raising=False)
        assert toc.build_static_toc(doc)["tab_pos_mm"] == pytest.approx(
            expected), helper


def test_stub_content_width_mm_that_raises_falls_back(monkeypatch):
    """A not-yet-implemented helper must not take TOC generation down with it."""
    def boom(*_args, **_kwargs):
        raise RuntimeError("not implemented yet")

    monkeypatch.setattr(dc, "content_width_mm", boom, raising=False)
    doc = toc_doc()
    report = toc.build_static_toc(doc)
    assert report["entries"] == 3
    assert report["tab_pos_mm"] == pytest.approx(160.0)
    assert report["tab_pos_source"] == "fallback"


def test_unresolved_pages_write_an_empty_slot():
    """``pages=None`` leaves the slot blank — the tab stop is still there, so a
    later paginator only has to append text, not rebuild the line."""
    doc = toc_doc()
    toc.build_static_toc(doc)                 # pages=None
    headings = ["第一章　静电场", "1.1　库仑定律", "1.1.1　点电荷模型"]
    for line, heading in zip(entries_of(doc), headings):
        assert all_text(line) == heading, all_text(line)
        assert len(run_tabs(line)) == 1
        assert stops_of(line)[0][1:] == (2, 1)


# ---- page number text is preserved verbatim ---------------------------------

def test_roman_page_numbers_survive():
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["i", "ii", "iii"], tab_pos_mm=160)
    assert all_text(entries_of(doc)[0]) == "第一章　静电场i"
    assert all_text(entries_of(doc)[1]) == "1.1　库仑定律ii"
    assert all_text(entries_of(doc)[2]) == "1.1.1　点电荷模型iii"


def test_letter_and_mixed_page_labels_survive():
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["A-1", "附录 2", "3.14"])
    got = [all_text(p) for p in entries_of(doc)]
    assert got == ["第一章　静电场A-1", "1.1　库仑定律附录 2",
                   "1.1.1　点电荷模型3.14"]
    # ...and the label is still on the far side of a real tab
    for line in entries_of(doc):
        assert len(run_tabs(line)) == 1


def test_short_pages_list_leaves_later_slots_empty():
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["i"])
    got = [all_text(p) for p in entries_of(doc)]
    assert got == ["第一章　静电场i", "1.1　库仑定律", "1.1.1　点电荷模型"]


# ---- indentation -------------------------------------------------------------

def test_indent_chars_per_level():
    doc = toc_doc()
    toc.build_static_toc(doc, indent_chars=(0, 2, 4))
    lines = entries_of(doc)
    assert lines[0].paragraph_format.left_indent.mm == pytest.approx(0.0,
                                                                    abs=0.1)
    assert lines[1].paragraph_format.left_indent.mm == pytest.approx(
        2 * dc.DEFAULT_CHAR_PT * 25.4 / 72, abs=0.2)
    assert lines[2].paragraph_format.left_indent.mm == pytest.approx(
        4 * dc.DEFAULT_CHAR_PT * 25.4 / 72, abs=0.2)
    # the character twin is what the ruler shows and survives a font change
    ind = lines[2]._p.find(qn("w:pPr")).find(qn("w:ind"))
    assert ind.get(qn("w:leftChars")) == "400"


# ---- title -------------------------------------------------------------------

def test_title_paragraph_is_inserted_before_the_entries():
    doc = toc_doc()
    toc.build_static_toc(doc, title="目录")
    texts = [p.text for p in doc.paragraphs]
    index = texts.index("目录")
    first_entry = next(i for i, p in enumerate(doc.paragraphs)
                       if toc.is_toc_paragraph(p))
    assert index < first_entry
    assert doc.paragraphs[index].style.name == "TOC Heading"


def test_title_can_be_custom_text_with_a_style():
    doc = Document()
    doc.add_heading("第一章", level=1)
    toc.build_static_toc(doc, title={"text": "目　录", "style": "TOC Heading"})
    assert doc.paragraphs[0].text == "目　录"
    assert toc.is_toc_title(doc.paragraphs[0])


def test_title_none_writes_no_title_line():
    doc = toc_doc()
    toc.build_static_toc(doc)
    assert not any(toc.is_toc_title(p) for p in doc.paragraphs)
    assert all_text(doc.paragraphs[0]) == "第一章　静电场"


# ---- idempotence -------------------------------------------------------------

def test_rebuild_twice_does_not_accumulate():
    doc = toc_doc()
    first = toc.build_static_toc(doc)
    count_after_first = len(doc.paragraphs)
    second = toc.build_static_toc(doc)
    assert second["entries"] == first["entries"] == 3
    assert len(doc.paragraphs) == count_after_first
    assert len(entries_of(doc)) == 3
    assert [p.style.name for p in entries_of(doc)] == ["toc 1", "toc 2", "toc 3"]


def test_rebuild_with_title_keeps_exactly_one_title():
    doc = toc_doc()
    toc.build_static_toc(doc, title="目录")
    count = len(doc.paragraphs)
    toc.build_static_toc(doc, title="目录")
    assert len(doc.paragraphs) == count
    titles = [p for p in doc.paragraphs if toc.is_toc_title(p)]
    assert len(titles) == 1


def test_rebuild_reflects_edited_body():
    """The whole point of a static TOC: edit a heading, regenerate, it follows."""
    doc = toc_doc()
    toc.build_static_toc(doc)
    for para in doc.paragraphs:
        if toc.is_toc_paragraph(para):
            continue
        if para.text == "第一章　静电场":
            para.runs[0].text = "第一章　静电场与电位"
            break
    report = toc.build_static_toc(doc)
    assert report["entries"] == 3
    assert all_text(entries_of(doc)[0]) == "第一章　静电场与电位"


def test_rebuild_stays_where_the_old_toc_was():
    """Not appending at the end: a TOC stays in front of the body."""
    doc = Document()
    doc.add_heading("第一章", level=1)
    doc.add_paragraph("正文。")
    doc.add_heading("第二章", level=1)
    toc.build_static_toc(doc, insert_at=0, title="目录")
    assert doc.paragraphs[0].text == "目录"
    toc.build_static_toc(doc, title="目录")
    assert doc.paragraphs[0].text == "目录"
    assert [p.text for p in doc.paragraphs].count("目录") == 1
    # the entries still precede the first heading
    first_entry = next(i for i, p in enumerate(doc.paragraphs)
                       if toc.is_toc_paragraph(p))
    first_heading = next(i for i, p in enumerate(doc.paragraphs)
                         if p.style.name == "Heading 1")
    assert first_entry < first_heading


def test_rebuild_after_removing_a_heading_drops_its_entry():
    doc = toc_doc()
    toc.build_static_toc(doc)
    victim = None
    for para in doc.paragraphs:
        if para.text == "1.1　库仑定律":
            victim = para
            break
    victim._p.getparent().remove(victim._p)
    report = toc.build_static_toc(doc)
    assert report["entries"] == 2
    assert [p.style.name for p in entries_of(doc)] == ["toc 1", "toc 3"]


# ---- empty document ----------------------------------------------------------

def test_empty_document_yields_zero_entries():
    doc = Document()
    report = toc.build_static_toc(doc)
    assert report["entries"] == 0
    assert report["levels"] == []
    assert report["styles"] == []
    assert report["inserted_at"] is None
    assert doc.paragraphs == [] or len(doc.paragraphs) == 0


def test_empty_document_with_title_keeps_only_the_title():
    doc = Document()
    report = toc.build_static_toc(doc, title="目录")
    assert report["entries"] == 0
    assert len(doc.paragraphs) == 1
    assert doc.paragraphs[0].text == "目录"


# ---- insert position ---------------------------------------------------------

def test_insert_at_zero_puts_the_toc_in_front():
    doc = Document()
    doc.add_heading("第一章", level=1)
    doc.add_paragraph("正文。")
    report = toc.build_static_toc(doc, insert_at=0)
    assert report["inserted_at"] == 0
    assert toc.is_toc_paragraph(doc.paragraphs[0])
    assert all_text(doc.paragraphs[0]) == "第一章"


def test_insert_at_past_the_end_appends():
    doc = Document()
    doc.add_heading("第一章", level=1)
    report = toc.build_static_toc(doc, insert_at=99)
    assert report["inserted_at"] == 1
    assert toc.is_toc_paragraph(doc.paragraphs[1])


def test_insert_at_keeps_document_order_of_entries():
    doc = Document()
    doc.add_heading("第一章", level=1)
    doc.add_heading("1.1　节", level=2)
    doc.add_heading("第二章", level=1)
    toc.build_static_toc(doc, insert_at=0)
    lines = entries_of(doc)
    assert [all_text(p) for p in lines] == ["第一章", "1.1　节", "第二章"]
    assert [p.style.name for p in lines] == ["toc 1", "toc 2", "toc 1"]


# ---- saved file --------------------------------------------------------------

def test_saved_docx_contains_the_tabs_and_leaders(tmp_path):
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["1", "1", "2"], tab_pos_mm=160,
                         title="目录")
    out = tmp_path / "static_toc.docx"
    doc.save(out)
    with zipfile.ZipFile(out) as archive:
        xml = archive.read("word/document.xml").decode("utf-8")
        styles = archive.read("word/styles.xml").decode("utf-8")
    assert "<w:tab/>" in xml
    assert 'w:leader="dot"' in xml
    assert 'w:val="right"' in xml
    assert "\t" not in xml
    for style_id in ("TOC1", "TOC2", "TOC3"):
        assert 'w:styleId="%s"' % style_id in styles


def test_saved_file_reopens_with_the_same_entries(tmp_path):
    doc = toc_doc()
    toc.build_static_toc(doc, pages=["i", "ii", "iii"], title="目录")
    out = tmp_path / "reopen.docx"
    doc.save(out)
    again = Document(str(out))
    lines = entries_of(again)
    assert [all_text(p) for p in lines] == [
        "第一章　静电场i", "1.1　库仑定律ii", "1.1.1　点电荷模型iii"]
    assert [p.style.name for p in lines] == ["toc 1", "toc 2", "toc 3"]
    for line in lines:
        assert len(run_tabs(line)) == 1
        assert stops_of(line)[0][2] == 1        # dot leader survives the save


def test_rebuild_works_on_a_reopened_file(tmp_path):
    doc = toc_doc()
    toc.build_static_toc(doc, title="目录")
    first = len(doc.paragraphs)
    out = tmp_path / "roundtrip.docx"
    doc.save(out)
    again = Document(str(out))
    report = toc.build_static_toc(again, title="目录")
    assert report["entries"] == 3
    assert len(again.paragraphs) == first


# ---- CLI: --pages must carry exactly one value per entry ---------------------
#
# ``--pages`` is mapped by position: the n-th value lands on the n-th entry.  A
# count mismatch is therefore not a matter of taste — surplus values are dropped
# and missing ones come out as blank slots, both silently, and a TOC with a blank
# page number in the middle of a delivered file reads exactly like one that was
# never paginated.  The library stays tolerant (a caller may pass a partial list
# on purpose); the *CLI* rejects the mismatch, because there it is a typo.
#
# These pin the exit code (2 = usage error), the fact that nothing is written,
# and — most importantly — that the message names both numbers, so the user can
# see whether to add or drop values without counting entries by hand.


def cli_source(tmp_path, name="pages_src.docx"):
    """``toc_doc()`` on disk: three headings (levels 1, 2, 3)."""
    path = tmp_path / name
    toc_doc().save(path)
    return path


def test_cli_rejects_fewer_pages_than_entries(tmp_path, capsys):
    src = cli_source(tmp_path)
    out = tmp_path / "few.docx"
    code = toc.main([str(src), "-o", str(out), "--pages", "1"])
    assert code == 2, "a short --pages list must be a usage error, not silence"
    assert not out.exists(), "a rejected run must not write the output file"
    err = capsys.readouterr().err
    assert "1 value" in err and "3 entries" in err, err


def test_cli_rejects_more_pages_than_entries(tmp_path, capsys):
    src = cli_source(tmp_path)
    out = tmp_path / "many.docx"
    code = toc.main([str(src), "-o", str(out), "--pages", "1,3,7,9"])
    assert code == 2, "surplus page numbers used to be dropped silently"
    assert not out.exists()
    err = capsys.readouterr().err
    assert "4 value" in err and "3 entries" in err, err


def test_cli_accepts_matching_pages_and_writes_the_same_toc(tmp_path, capsys):
    """The happy path is unchanged: exit 0, JSON report, identical lines."""
    src = cli_source(tmp_path)
    out = tmp_path / "match.docx"
    code = toc.main([str(src), "-o", str(out), "--pages", "1,3,7"])
    assert code == 0
    captured = capsys.readouterr()
    assert captured.err == ""
    report = json.loads(captured.out)
    assert report["entries"] == 3
    assert report["output"] == str(out)

    written = entries_of(Document(str(out)))
    assert [all_text(p) for p in written] == [
        "第一章　静电场1", "1.1　库仑定律3", "1.1.1　点电荷模型7"]
    # ...and the API, given the same numbers, produces the same lines
    reference = toc_doc()
    toc.build_static_toc(reference, pages=["1", "3", "7"])
    assert [all_text(p) for p in entries_of(reference)] == [
        all_text(p) for p in written]


def test_cli_page_count_follows_the_levels_filter(tmp_path, capsys):
    """The check counts the entries *this* run would write, not all headings.

    Counting headings would reject a correct ``--levels 1 --pages 1`` on a
    three-heading document — the exact bug this test exists to catch.
    """
    src = cli_source(tmp_path)
    out = tmp_path / "filtered.docx"
    code = toc.main([str(src), "-o", str(out), "--levels", "1,2",
                     "--pages", "1,3"])
    assert code == 0, capsys.readouterr().err
    assert [all_text(p) for p in entries_of(Document(str(out)))] == [
        "第一章　静电场1", "1.1　库仑定律3"]

    out2 = tmp_path / "filtered_bad.docx"
    code = toc.main([str(src), "-o", str(out2), "--levels", "1,2",
                     "--pages", "1,3,7"])
    assert code == 2
    assert not out2.exists()
    assert "3 value" in capsys.readouterr().err


def test_cli_single_entry_message_is_singular(tmp_path, capsys):
    doc = Document()
    doc.add_heading("唯一的标题", level=1)
    src = tmp_path / "one.docx"
    doc.save(src)
    code = toc.main([str(src), "-o", str(tmp_path / "one_out.docx"),
                     "--pages", "1"])
    assert code == 0
    code = toc.main([str(src), "-o", str(tmp_path / "one_bad.docx"),
                     "--pages", "1,2"])
    assert code == 2
    assert "1 entry;" in capsys.readouterr().err


def test_cli_omitting_pages_still_writes_empty_slots(tmp_path):
    """No ``--pages`` at all is the documented "paginate later" mode."""
    src = cli_source(tmp_path)
    out = tmp_path / "nopages.docx"
    assert toc.main([str(src), "-o", str(out)]) == 0
    assert [all_text(p) for p in entries_of(Document(str(out)))] == [
        "第一章　静电场", "1.1　库仑定律", "1.1.1　点电荷模型"]
