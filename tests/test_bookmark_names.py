# -*- coding: utf-8 -*-
"""Bookmark-name guards: unique, space-free, at most 40 characters.

Every rule here comes from measurement, not from the OOXML schema — see
``work/T3-REPORT.md``:

* §3.1  duplicate names are **merged** by Word (the later Start wins, so a
        table-of-contents entry jumps to the wrong chapter) and **renamed**
        ``<name>_副本_1`` by LibreOffice (so every ``PAGEREF`` dies);
* §3.2  a space in the name becomes ``_`` in both applications on save;
* §3.3  a name longer than 40 characters is truncated to 40 by both;
* §3.4  CJK names, a leading digit (``1Ref``) and an interior hyphen
        (``_Ref-a``) all survive untouched, so they are allowed here.

All three defects are silent and invisible in the XML we write, which is why
they are caught at generation time.  Nothing in this file needs Word or
LibreOffice: the checks are pure XML inspection on ``tmp_path`` documents.
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
    BOOKMARK_MAX_LEN, all_bookmarks, check_crossref_anchors,
    next_toc_name, validate_bookmark_name, validate_bookmark_uniqueness,
)
from docx_utils import next_bm_id                           # noqa: E402


def add_bookmark(doc, para, name: str) -> None:
    """Bracket *para* with a ``bookmarkStart``/``bookmarkEnd`` pair named *name*."""
    bid = next_bm_id(doc.element.body)
    start = doc.element.body.makeelement(qn("w:bookmarkStart"), {})
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = doc.element.body.makeelement(qn("w:bookmarkEnd"), {})
    end.set(qn("w:id"), str(bid))
    para._p.insert(0, start)
    para._p.append(end)


def para_text_with_bookmark(doc, text: str, name: str):
    p = doc.add_paragraph(text)
    add_bookmark(doc, p, name)
    return p


# ---- validate_bookmark_name: positives -------------------------------------

@pytest.mark.parametrize("name", [
    "_Toc1",            # the canonical TOC anchor this module mints
    "_Ref3925",         # Word's own caption family
    "_Ref图1",          # CJK inside the name (T3 §3.4: survives everywhere)
    "_Toc图一",         # CJK-only tail (T3 §3.4)
    "1Ref",             # leading digit — verified working in T3 §3.4
    "_Ref-a",           # interior hyphen — verified working in T3 §3.4
    "fig_1",            # plain ASCII the caption code already emits
])
def test_valid_names_pass(name):
    assert validate_bookmark_name(name) == []


def test_thirty_two_characters_pass():
    """T3 §5.2 recommends leaving margin under 40; 32 must be accepted."""
    name = "a" * 32
    assert len(name) == 32
    assert validate_bookmark_name(name) == []


def test_exactly_forty_characters_pass():
    """40 is the cap itself: at the limit is fine, one over is not."""
    name = "_" + "x" * 39
    assert len(name) == BOOKMARK_MAX_LEN
    assert validate_bookmark_name(name) == []


def test_cjk_name_of_forty_characters_passes():
    name = "字" * 40
    assert validate_bookmark_name(name) == []


# ---- validate_bookmark_name: negatives -------------------------------------

def test_space_is_rejected():
    """T3 §3.2: Word and LO both turn the space into '_' on save."""
    problems = validate_bookmark_name("_Ref图 1")
    assert problems
    assert any("whitespace" in p for p in problems)


def test_space_is_rejected_even_though_the_xml_would_accept_it():
    """The defect is silent, so the guard must not rely on the writer failing."""
    d = Document()
    para_text_with_bookmark(d, "正文", "_Ref图 1")     # writes fine
    assert "_Ref图 1" in all_bookmarks(d)
    assert validate_bookmark_name("_Ref图 1")          # still flagged


def test_other_whitespace_is_rejected_too():
    for name in ("_Ref\t1", "_Ref\n1", "_Ref\u00a01", "_Ref1 "):
        assert validate_bookmark_name(name), name


def test_forty_one_characters_are_rejected():
    """T3 §3.3: both applications truncate to 40, so 41 is a broken anchor."""
    name = "a" * 41
    problems = validate_bookmark_name(name)
    assert problems
    assert any(str(BOOKMARK_MAX_LEN) in p for p in problems)


def test_forty_one_cjk_characters_are_rejected():
    assert validate_bookmark_name("图" * 41)


def test_empty_name_is_rejected():
    assert validate_bookmark_name("")


def test_punctuation_is_accepted_because_word_keeps_it():
    """Measured, not assumed: Word 16.0 opens, saves and resolves all of these.

    The first version of this guard used an allow-list of ``[A-Za-z0-9_]`` plus
    CJK and rejected every name below.  Word disagrees -- each of these survived
    a save byte for byte and a ``REF`` pointing at it resolved after
    ``Fields.Update()`` (work/REVIEW-ADVERSARIAL-R6.md 阻塞-1).  Rejecting them
    blocked the most common Chinese bookmark names, so they must pass.
    """
    for name in ("_Ref图.1", "_Ref图/1", "_Ref图(1)", "_Ref图（1）",
                 "_Ref图:1", "_Ref图：1", "_Ref图#1", "_Ref图$1", "_Ref图1!",
                 "_Ref图[1]", "_Ref图【1】", "_Ref图·1", "_Ref图—1",
                 "_Ref第〇章", "_RefＡ"):
        assert validate_bookmark_name(name) == [], name


def test_whitespace_is_rejected_and_invisibles_are_not():
    """Only the whitespace class is rewritten; format characters survive.

    Measured on Word 16.0, one bookmark per character, saved and read back:

    ==================  ==================  ==========
    character           after a Word save   verdict
    ==================  ==================  ==========
    ``U+00A0`` NBSP      ``_``               reject
    ``U+3000`` 全角空格   ``_``               reject
    ``U+200B`` ZWSP      unchanged           allow
    ``U+200E`` LRM       unchanged           allow
    ``U+FEFF`` BOM       unchanged           allow
    ``U+2060`` WJ        unchanged           allow
    ==================  ==================  ==========

    ``str.isspace()`` draws exactly that line, which is why the guard uses it
    rather than a hand-written character class.
    """
    for ch in ("\u00a0", "\u3000", " ", "\t"):
        assert validate_bookmark_name("_Ref" + ch + "1"), repr(ch)
    for ch in ("\u200b", "\u200e", "\ufeff", "\u2060"):
        assert validate_bookmark_name("_Ref" + ch + "1") == [], repr(ch)


def test_fullwidth_digits_are_accepted():
    """Full width digits are not folded to ASCII, but they are also not changed."""
    assert validate_bookmark_name("_Ref１") == []


def test_control_characters_are_rejected():
    """ST_String forbids the C0 range; that is the only hard character rule."""
    for name in ("_Ref\x001", "_Ref\x1f1", "_Ref\x7f1", "_Ref\n1", "_Ref\r1"):
        assert validate_bookmark_name(name), repr(name)


def test_non_string_is_rejected():
    assert validate_bookmark_name(None)
    assert validate_bookmark_name(123)


def test_custom_max_len_is_honoured():
    assert validate_bookmark_name("abcde", max_len=5) == []
    assert validate_bookmark_name("abcdef", max_len=5)


def test_problems_are_a_list_of_strings():
    problems = validate_bookmark_name("bad name " + "x" * 50)
    assert isinstance(problems, list)
    assert all(isinstance(p, str) for p in problems)
    assert len(problems) >= 2          # whitespace + length


def test_whitespace_is_reported_once_not_twice():
    """A space is a whitespace problem; it is not also spelled out as illegal."""
    problems = validate_bookmark_name("_Ref图 1")
    assert len(problems) == 1
    assert "whitespace" in problems[0]


# ---- validate_bookmark_uniqueness ------------------------------------------

def test_clean_document_is_unique():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图1")
    para_text_with_bookmark(d, "图 2", "_Ref图2")
    report = validate_bookmark_uniqueness(d)
    assert report["ok"] is True
    assert report["duplicates"] == []
    assert report["count"] == 2
    assert report["unique"] == 2


def test_document_without_bookmarks_is_unique():
    d = Document()
    d.add_paragraph("正文")
    report = validate_bookmark_uniqueness(d)
    assert report["ok"] is True
    assert report["count"] == 0
    assert report["unique"] == 0


def test_duplicate_names_are_not_ok():
    """T3 §3.1: Word merges them, LO renames the second — both silently."""
    d = Document()
    para_text_with_bookmark(d, "图 1", "_RefDup图")
    para_text_with_bookmark(d, "图 1 又来一次", "_RefDup图")
    report = validate_bookmark_uniqueness(d)
    assert report["ok"] is False
    assert report["duplicates"] == ["_RefDup图"]
    assert report["count"] == 2
    assert report["unique"] == 1


def test_every_duplicate_name_is_listed_once():
    d = Document()
    for _ in range(3):
        para_text_with_bookmark(d, "正文", "_RefDup图")
    para_text_with_bookmark(d, "正文", "_RefDup表")
    para_text_with_bookmark(d, "正文", "_RefDup表")
    report = validate_bookmark_uniqueness(d)
    assert report["duplicates"] == ["_RefDup图", "_RefDup表"]
    assert report["count"] == 5
    assert report["unique"] == 2


def test_duplicates_are_reported_by_name_not_by_id():
    """Two names, same bookmark id, is not the failure mode; same name is."""
    d = Document()
    p1 = d.add_paragraph("图 1")
    p2 = d.add_paragraph("图 2")
    para_text_with_bookmark(d, "图 1", "_RefA")
    para_text_with_bookmark(d, "图 2", "_RefB")
    assert "_RefA" in all_bookmarks(d)
    assert validate_bookmark_uniqueness(d)["ok"] is True
    assert p1 is not p2


# ---- next_toc_name ---------------------------------------------------------

def test_next_toc_name_shape():
    d = Document()
    assert next_toc_name(d) == "_Toc0001"


def test_next_toc_name_skips_names_already_in_the_document():
    d = Document()
    para_text_with_bookmark(d, "第一章", "_Toc0001")
    para_text_with_bookmark(d, "第二章", "_Toc0002")
    assert next_toc_name(d) == "_Toc0003"


def test_next_toc_name_continues_from_an_unpadded_existing_name():
    d = Document()
    para_text_with_bookmark(d, "第一章", "_Toc7")
    name = next_toc_name(d)
    assert name == "_Toc0008"


def test_next_toc_name_hundred_calls_are_all_distinct_and_legal():
    """The batch-minting case: 100 names handed out before any is written."""
    d = Document()
    names = [next_toc_name(d) for _ in range(100)]
    assert len(set(names)) == 100
    for name in names:
        assert " " not in name
        assert len(name) <= BOOKMARK_MAX_LEN
        assert validate_bookmark_name(name) == []


def test_next_toc_name_mixed_with_real_insertion_stays_unique():
    """Interleaving insertion must not reopen a name already handed out."""
    d = Document()
    names = []
    for i in range(30):
        name = next_toc_name(d)
        names.append(name)
        para_text_with_bookmark(d, "第 %d 节" % i, name)
    assert len(set(names)) == 30
    assert set(names) == all_bookmarks(d)
    assert validate_bookmark_uniqueness(d)["ok"] is True


def test_next_toc_name_respects_a_custom_prefix():
    d = Document()
    name = next_toc_name(d, prefix="_TocSec")
    assert name.startswith("_TocSec")
    assert len(name) <= BOOKMARK_MAX_LEN
    assert validate_bookmark_name(name) == []


def test_next_toc_name_keeps_a_long_prefix_within_the_cap():
    d = Document()
    name = next_toc_name(d, prefix="_" + "T" * 60)
    assert len(name) == BOOKMARK_MAX_LEN
    assert validate_bookmark_name(name) == []


def test_next_toc_name_sanitises_an_illegal_prefix():
    d = Document()
    name = next_toc_name(d, prefix="_Toc 目录")
    assert " " not in name
    assert validate_bookmark_name(name) == []


def test_next_toc_name_handles_an_empty_or_unusable_prefix():
    assert next_toc_name(Document(), prefix="") == "_Toc0001"
    assert next_toc_name(Document(), prefix=None) == "_Toc0001"
    name = next_toc_name(Document(), prefix="图 表")
    assert " " not in name
    assert validate_bookmark_name(name) == []


def test_next_toc_name_stays_legal_past_the_four_digit_range():
    """10000 is not padded back to 4 digits, and must still be a legal name."""
    d = Document()
    para_text_with_bookmark(d, "第一万节", "_Toc9999")
    name = next_toc_name(d)
    assert name == "_Toc10000"
    assert len(name) <= BOOKMARK_MAX_LEN
    assert validate_bookmark_name(name) == []


def test_next_toc_name_is_legal_where_a_toc_field_would_reference_it():
    """The whole point: a minted name can be dropped into a dangling-free TOC."""
    d = Document()
    p = d.add_paragraph("第一章 绪论")
    name = next_toc_name(d)
    add_bookmark(d, p, name)
    assert check_crossref_anchors(d)["invalid_names"] == []
    assert check_crossref_anchors(d)["duplicate_names"] == []


# ---- check_crossref_anchors: the two new keys ------------------------------

def test_check_anchors_on_clean_document_has_empty_new_keys():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图1")
    report = check_crossref_anchors(d)
    assert report["invalid_names"] == []
    assert report["duplicate_names"] == []


def test_check_anchors_reports_duplicate_names():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_RefDup图")
    para_text_with_bookmark(d, "图 1 再来", "_RefDup图")
    report = check_crossref_anchors(d)
    assert report["duplicate_names"] == ["_RefDup图"]


def test_check_anchors_reports_invalid_names():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图 1")
    report = check_crossref_anchors(d)
    names = [item["name"] for item in report["invalid_names"]]
    assert names == ["_Ref图 1"]
    assert report["invalid_names"][0]["problems"]


def test_check_anchors_invalid_names_carry_problems_per_occurrence():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图 1")
    para_text_with_bookmark(d, "图 2", "_Ref图 1")
    report = check_crossref_anchors(d)
    # two starts with the same bad name → two entries, in document order
    assert len(report["invalid_names"]) == 2
    assert all(i["name"] == "_Ref图 1" for i in report["invalid_names"])
    assert report["duplicate_names"] == ["_Ref图 1"]


def test_check_anchors_keeps_the_existing_keys():
    """test_caption_ref / test_crossref_extra / test_crossref_forward rely on these."""
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图1")
    report = check_crossref_anchors(d)
    assert set(report) >= {"resolved", "targets", "missing", "known"}
    assert report["known"] == ["_Ref图1"]
    assert report["missing"] == []
    assert report["targets"] == 0
    assert report["resolved"] == 0


def test_check_anchors_missing_and_invalid_can_coexist():
    d = Document()
    para_text_with_bookmark(d, "图 1", "_Ref图 1")
    assert check_crossref_anchors(d)["missing"] == []
    assert check_crossref_anchors(d)["invalid_names"]


def test_check_anchors_does_not_flag_word_generated_names():
    """A real Word-saved document must stay clean, or the guard is useless."""
    d = Document()
    for name in ("_Ref3925", "_Ref8179", "_Toc图一", "_Toc1", "1Ref", "_Ref-a"):
        para_text_with_bookmark(d, "标题", name)
    report = check_crossref_anchors(d)
    assert report["invalid_names"] == []
    assert report["duplicate_names"] == []
