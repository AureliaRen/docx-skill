# -*- coding: utf-8 -*-
"""``--static`` cross-references: the number as literal text, not a field.

Why this file exists
--------------------
The operator reported, from a real Chinese typesetting job: a superscript
cross-reference could not be made to stay superscript in the exported PDF.  The
field design is where that bites.  ``REF`` keeps the number in sync with its
caption for free, but the digits a reader sees are a *field result*, and Word
re-renders a field result from its instruction on the next F9 / print / PDF
export.  ``\\* MERGEFORMAT`` is what carries the raised look across that
re-render — and a caller who wants the superscript to be structurally immune,
rather than dependent on a switch, has nothing to reach for.

``--static`` is that thing: one ``<w:hyperlink>`` holding one ``<w:r>`` with the
number as ``<w:t>``.  Clickable, styled, and nothing to recalculate.

Measured on Word 16.0 (COM, full ``Fields.Update()`` then PDF export, text spans
read back with PyMuPDF).  Four documents, same paragraph, same prefix/suffix:

    field + MERGEFORMAT   ``'1' sz 7.0 y 197.18``  raised 2.52 pt, half size
    field, no MERGEFORMAT ``'1' sz 11.0 y 199.70`` the operator's failure
    ``--static``          ``'1' sz 7.0 y 197.18``  identical to row 1
    body text             ``sz 11.0 y 199.70``     the baseline

The third row is the point: the same measurement as the first without needing
MERGEFORMAT to hold it up.  Rows 2 and 3 together are why this file asserts what
it does — a static reference must carry no instruction text at all, or it is not
static no matter what it looks like.

Everything here is XML level on ``tmp_path`` documents, so it runs on a runner
with no Word.  The one assertion that needs Word's own renderer (the PDF span
comparison) lives in ``tools/`` rather than here; see
``docs/reports/P17-C-CROSSREF.md`` for the recorded output.
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
    UsageError,
    all_bookmarks,
    cmd_bookmark,
    cmd_caption,
    cmd_ref,
)


class Args:
    """Stand-in for the argparse namespace the ``cmd_*`` functions read.

    Keys default to ``None``/"" so one class serves both ``cmd_caption`` (which
    never looks at ``target``) and ``cmd_ref``.
    """

    def __init__(self, **kw):
        self.index = kw.pop("index")
        self.target = kw.pop("target", None)
        self.prefix = kw.get("prefix", "")
        self.suffix = kw.get("suffix", "")
        self.format = kw.get("format", "default")
        self.cached = kw.get("cached")
        self.vert_align = kw.get("vert_align", "normal")
        self.superscript = kw.get("superscript", False)
        self.no_hyperlink = kw.get("no_hyperlink", False)
        self.no_mergeformat = kw.get("no_mergeformat", False)
        self.allow_missing = kw.get("allow_missing", False)
        self.static = kw.get("static", False)
        self.label = kw.get("label")
        self.text = kw.get("text")
        self.bookmark = kw.get("bookmark")
        self.bookmark_span = kw.get("bookmark_span", "number")
        self.drop_text = kw.get("drop_text", False)
        self.ref_bookmark = kw.get("ref_bookmark", False)
        self.pstyle = kw.get("pstyle")


def _doc_with_caption():
    """A document with a number-only caption bookmark; returns ``(doc, name)``.

    The name comes back from ``cmd_caption`` rather than being hard coded:
    ``图_1`` today, but the naming rule is the tool's business, not this test's.
    """
    doc = Document()
    doc.add_paragraph("图 1  系统架构示意")
    doc.add_paragraph("引用测试占位。")
    report = cmd_caption(doc, Args(index=0, label="图", text="系统架构示意",
                                   bookmark=None, bookmark_span="number",
                                   drop_text=False, ref_bookmark=False,
                                   pstyle=None))
    return doc, report["bookmark"]


def _hyperlink_xml(doc) -> str:
    from lxml import etree
    hits = list(doc.element.body.iter(qn("w:hyperlink")))
    assert len(hits) == 1, f"expected 1 hyperlink, got {len(hits)}"
    return etree.tostring(hits[0], encoding="unicode")


# ---- the property --static exists for ---------------------------------------

def test_static_writes_literal_text_and_no_field():
    """No instrText, no fldChar: nothing for Word to re-render."""
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target=BM, static=True))
    assert result["mode"] == "static"
    assert result["needs_field_update"] is False
    assert result["text"] == "1"

    body = doc.element.body
    # Scoped to the paragraph the reference went into: the caption paragraph
    # legitimately carries the SEQ field, and asserting on the whole body would
    # fail for a reason that has nothing to do with --static.
    ref_para = doc.paragraphs[1]._p
    assert not list(ref_para.iter(qn("w:instrText"))), \
        "a static reference must carry no field instruction"
    assert not list(ref_para.iter(qn("w:fldChar"))), \
        "a static reference must carry no field character"
    texts = [t.text for t in body.iter(qn("w:t"))]
    assert "1" in texts


def test_static_keeps_superscript_on_the_run_carrying_the_number():
    """The vertAlign must sit on the run that holds the digits."""
    doc, BM = _doc_with_caption()
    cmd_ref(doc, Args(index=1, target=BM, static=True,
                      superscript=True))
    xml = _hyperlink_xml(doc)
    assert 'w:val="superscript"' in xml
    # The raised run is the one holding the number, not a neighbouring run.
    from lxml import etree
    h = next(doc.element.body.iter(qn("w:hyperlink")))
    runs = list(h.iter(qn("w:r")))
    assert len(runs) == 1, f"static hyperlink should hold one run, got {len(runs)}"
    run = runs[0]
    va = run.find(qn("w:rPr") + "/" + qn("w:vertAlign"))
    assert va is not None and va.get(qn("w:val")) == "superscript"
    assert run.find(qn("w:t")).text == "1"


def test_static_number_still_clickable():
    """A static reference stays a real internal link (that half is unchanged)."""
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target=BM, static=True))
    assert result["hyperlink"] is True
    assert result["anchor"] == BM
    h = next(doc.element.body.iter(qn("w:hyperlink")))
    assert h.get(qn("w:anchor")) == BM
    assert h.get(qn("w:history")) == "1"


def test_static_no_hyperlink_writes_bare_run():
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target=BM, static=True,
                               no_hyperlink=True))
    assert result["hyperlink"] is False
    assert not list(doc.element.body.iter(qn("w:hyperlink")))


def test_static_uses_prefix_and_suffix_around_the_number():
    doc, BM = _doc_with_caption()
    cmd_ref(doc, Args(index=1, target=BM, static=True,
                      prefix="如图", suffix="所示"))
    para = doc.paragraphs[1]
    text = "".join(t.text or "" for t in para._p.iter(qn("w:t")))
    assert text == "引用测试占位。如图1所示"


# ---- the trade: static text does not track its caption ----------------------

def test_static_number_does_not_follow_the_seq_counter():
    """Documenting the cost, not asserting a bug: static means frozen.

    Two captions are added *after* a static reference has been written.  The
    field form would show the new value on F9; the static one keeps what it was
    given.  A caller who wants tracking uses the field form (the default).
    """
    doc, BM = _doc_with_caption()
    cmd_ref(doc, Args(index=1, target=BM, static=True))
    before = "".join(t.text or "" for t in doc.paragraphs[1]._p.iter(qn("w:t")))

    doc.add_paragraph("图 2  第二张")
    cmd_caption(doc, Args(index=2, label="图", text="第二张", bookmark=None,
                          bookmark_span="number", drop_text=False,
                          ref_bookmark=False, pstyle=None))
    # The static run is untouched by the SEQ machinery.
    after = "".join(t.text or "" for t in doc.paragraphs[1]._p.iter(qn("w:t")))
    assert before == after, "a static reference must not be rewritten by SEQ"
    assert not list(doc.paragraphs[1]._p.iter(qn("w:instrText")))


# ---- formats that cannot be expressed as static text ------------------------

@pytest.mark.parametrize("fmt", ["page", "above-below"])
def test_static_refuses_formats_whose_value_exist_only_after_layout(fmt):
    """Refused, not guessed: a stale page number is a wrong document."""
    doc, BM = _doc_with_caption()
    with pytest.raises(UsageError) as exc:
        cmd_ref(doc, Args(index=1, target=BM, static=True, format=fmt))
    assert "static" in str(exc.value)
    # Nothing was written before the refusal.
    assert not list(doc.paragraphs[1]._p.iter(qn("w:hyperlink")))


def test_static_unknown_bookmark_still_refused():
    doc, BM = _doc_with_caption()
    with pytest.raises(SystemExit) as exc:
        cmd_ref(doc, Args(index=1, target="nope_9", static=True))
    assert "nope_9" in str(exc.value)


def test_static_allow_missing_writes_forward_reference():
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target="ref_9", static=True,
                               allow_missing=True))
    assert result["text"] == "9"     # name-derived, the project's convention
    h = next(doc.element.body.iter(qn("w:hyperlink")))
    assert h.get(qn("w:anchor")) == "ref_9"


# ---- the field mode is unchanged -------------------------------------------

def test_field_mode_still_reports_needs_field_update():
    """Adding ``mode`` must not disturb the field form's existing contract."""
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target=BM, superscript=True))
    assert result["mode"] == "field"
    assert result["needs_field_update"] is True
    assert result["field"] == "REF"
    assert "MERGEFORMAT" in result["switch"]
    assert list(doc.element.body.iter(qn("w:instrText")))


def test_field_mode_write_is_byte_identical_to_before_static_existed():
    """The static branch must not have moved the field branch's markup.

    Pins the shape the module docstring documents: hyperlink wrapping five runs
    (begin / instr / separate / result / end), every one carrying vertAlign.
    """
    doc, BM = _doc_with_caption()
    cmd_ref(doc, Args(index=1, target=BM, superscript=True))
    h = next(doc.element.body.iter(qn("w:hyperlink")))
    runs = list(h.iter(qn("w:r")))
    assert len(runs) == 5, f"field form should keep five runs, got {len(runs)}"
    for run in runs:
        va = run.find(qn("w:rPr") + "/" + qn("w:vertAlign"))
        assert va is not None, "every field run carries vertAlign"
        assert va.get(qn("w:val")) == "superscript"


def test_no_mergeformat_still_available_and_documented_as_risky():
    """The dial that caused the operator's failure is still reachable on purpose."""
    doc, BM = _doc_with_caption()
    result = cmd_ref(doc, Args(index=1, target=BM, superscript=True,
                               no_mergeformat=True))
    assert result["merge_format"] is False
    assert "MERGEFORMAT" not in result["switch"]
