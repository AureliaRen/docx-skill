# -*- coding: utf-8 -*-
"""Tests for the page-number *glyph family* judgement inside ``_toc_audit``.

``_family`` lives **nested inside** ``word_accept._toc_audit`` (it closes over
``_ROMAN_VALUE`` and nothing else), so it cannot be imported and poked at
directly.  The only honest way to test it is through what the audit publishes:
``page_tokens`` (raw glyphs, in document order), ``monotonic`` /
``monotonic_verdict`` (what the reader is told), ``unparsed_tokens`` (how much
of the TOC the judge refused to read) and ``monotonic_partial``.

What ``_family`` decides, and why the cases below exist
------------------------------------------------------
A trailing TOC page number can be written in three alphabets: ``decimal``
(``15``), ``roman`` (``iv``) and ``letter`` (``a``, ``aa``).  The audit splits
the token stream into *runs* at every explicit family change and compares
numbers **inside** a run only, because a document that restarts its counter for
a new format (``i, ii, iii`` then ``1, 2``) is correct, not broken.  Whether
that split happens is therefore the whole correctness question.

Round 6 found two shapes with no dedicated test:

* ``['A', 'a']`` — a single letter written in both cases.  Before R7-5a
  ``_family`` lowercased before classifying, so both tokens landed in the
  *same* ``letter`` run, and that case fold was what made an ``upperLetter``
  -> ``lowerLetter`` counter change invisible to the splitter (the false
  alarm in ``work/P3-REPORT-B.md`` §1.4).  R7-5a made the case part of the
  family; this pair must still read ``1, 1`` flat and stay ``true``.
* ``['1', '01']`` — a decimal with a leading zero (Word's ``decimalZero``).
  Both tokens are ``isdigit()``, so both are ``decimal`` and both are read
  through ``int()``: ``1`` and ``1``.

The assertions record **observed behaviour**, not aspiration.  In R7-5a the
lead accepted the minimal fix from ``work/P3-REPORT-B.md`` §1.5 — the letter
and Roman families are case-aware (``letter_upper``/``letter_lower``,
``roman_upper``/``roman_lower``), a case-only family change splits runs, and a
step down that coincides with a case change is *observed* (``monotonic``
False) but not asserted (verdict ``unknown``, the same refusal a declared
``w:start`` restart gets) — and the two tests that pinned the old false alarm
were flipped in the same commit:
``test_case_change_between_letter_counters_reads_as_unknown`` and
``test_upper_then_lower_roman_reads_as_unknown``.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(TOOLS))
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import word_accept as wa  # noqa: E402

TOC_STYLE_ID = "9"
TOC_STYLE_NAME = "toc 1"


# --------------------------------------------------------------------------- helpers
def _toc_paragraphs(tokens):
    """One ``toc 1`` paragraph per token, ``title<TAB>token`` like a real TOC.

    The tab matters: ``_toc_token`` treats a tab as the authoritative
    entry/number separator, so this is the same shape Word writes.  The style is
    added to ``styles.xml`` and referenced by *id* because that is what the XML
    stores (``w:pStyle w:val="9"`` -> name ``toc 1``).
    """
    doc = Document()
    st = OxmlElement("w:style")
    st.set(qn("w:type"), "paragraph")
    st.set(qn("w:styleId"), TOC_STYLE_ID)
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), TOC_STYLE_NAME)
    st.append(nm)
    doc.styles.element.append(st)
    for i, tok in enumerate(tokens):
        p = doc.add_paragraph()
        ppr = p._p.get_or_add_pPr()
        ps = OxmlElement("w:pStyle")
        ps.set(qn("w:val"), TOC_STYLE_ID)
        ppr.insert(0, ps)
        p.add_run("entry%d" % i).add_tab()
        p.add_run(tok)
    return [p._p for p in doc.paragraphs]


def audit(tokens):
    """``_toc_audit`` over a TOC carrying ``tokens`` as its page numbers."""
    return wa._toc_audit(_toc_paragraphs(tokens), {TOC_STYLE_ID: TOC_STYLE_NAME})


# --------------------------------------------------------------------------- controls
def test_decimal_run_control():
    """``1, 2, 3`` is one decimal run and is reported monotonic."""
    ta = audit(["1", "2", "3"])
    assert ta["page_tokens"] == ["1", "2", "3"]
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0
    assert ta["monotonic_partial"] is False


def test_roman_run_control():
    """``i, ii, iii`` is read as Roman and is monotonic."""
    ta = audit(["i", "ii", "iii"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


def test_letter_run_control():
    """``a, b, c`` is read as a Latin sequence, not as Roman ``100, 1, 2``.

    This is the control for the tentatively-added ``letter`` family: reading
    those tokens as Roman would call a correct appendix a regression.
    """
    ta = audit(["a", "b", "c"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


def test_descending_letter_run_is_still_reported():
    """``c, a, b`` is a real regression and must not be excused by the splitter."""
    ta = audit(["c", "a", "b"])
    assert ta["monotonic"] is False
    assert ta["monotonic_verdict"] == "false"


# --------------------------------------------------------------------------- P1-a: ['A','a']
def test_single_letter_case_mix_reads_flat():
    """``['A', 'a']`` still reads ``1, 1`` — flat — under the case-aware family.

    Since R7-5a the case split puts each glyph in its own run
    (``letter_upper`` / ``letter_lower``), but both read as 1 and the step
    across the boundary is not a step *down*, so nothing is observed and
    nothing is downgraded: ``monotonic`` true, nothing unread, and the raw
    glyphs survive in ``page_tokens`` unchanged.
    """
    ta = audit(["A", "a"])
    assert ta["page_tokens"] == ["A", "a"]          # glyphs are not normalised
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0
    assert ta["monotonic_partial"] is False


def test_single_letter_case_mix_is_not_flagged_as_unread():
    """The same pair in reverse order is symmetric: readable, flat, ``true``."""
    ta = audit(["a", "A"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


def test_upper_case_letter_run_control():
    """``A, B, C`` is a letter run too — the fold is the reason ``A`` is not Roman."""
    ta = audit(["A", "B", "C"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


# --------------------------------------------------- P1-a, the suspicious consequence
def test_case_change_between_letter_counters_reads_as_unknown():
    """**Flipped in R7-5a** — ``upperLetter`` -> ``lowerLetter`` is no longer "false".

    ``['A', 'B', 'a', 'b']`` is what a document whose front matter is numbered
    ``upperLetter`` and whose body restarts in ``lowerLetter`` puts in its TOC.
    The letter family is case-aware now, so the stream splits into a
    ``letter_upper`` run (``1, 2``) and a ``letter_lower`` run (``1, 2``); the
    ``2 -> 1`` step across the case change is still *observed* — ``monotonic``
    is False — but the verdict is downgraded to ``unknown``: from the bytes
    alone, a restarted counter in the other case cannot be told apart from a
    mis-cased token.  Same refusal as ``page_number_restarts`` (R6).  Before
    R7-5a this reported ``false`` (work/P3-REPORT-B.md §1.4 可疑①).
    """
    ta = audit(["A", "B", "a", "b"])
    assert ta["monotonic"] is False            # the 2 -> 1 step is observed
    assert ta["monotonic_verdict"] == "unknown"
    assert ta["unparsed_tokens"] == 0          # everything was readable


def test_upper_then_lower_roman_reads_as_unknown():
    """Same shape for Roman: ``I, II, III`` then ``i, ii`` -> observed, not asserted.

    ``upperRoman`` -> ``lowerRoman`` is a case-only family change, so the runs
    split; the lone ``i`` (family ``single``) is read under the upper run's
    counter, the ``3 -> 1`` step it forms is observed, and the verdict is
    ``unknown`` for the same reason as the letter case above.  Before R7-5a
    this reported ``false`` (work/P3-REPORT-B.md §1.4 可疑②).
    """
    ta = audit(["I", "II", "III", "i", "ii"])
    assert ta["monotonic"] is False
    assert ta["monotonic_verdict"] == "unknown"


def test_explicit_family_change_is_split():
    """Control for the two tests above: ``roman`` -> ``decimal`` *does* split.

    ``i, ii, iii`` then ``1, 2, 3`` is a counter restart across a format change
    and is reported monotonic — no step down is *asserted* across an alphabet
    change, unlike a case-only change (the two tests above), where the step is
    observed but the verdict refuses to judge.
    """
    ta = audit(["i", "ii", "iii", "1", "2", "3"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


def test_case_change_without_a_step_down_stays_true():
    """``a, b, E, F`` — case flip, but the sequence rises: nothing to excuse.

    ``E, F`` continue 5, 6 above ``a, b`` = 1, 2, so the case change caused no
    step down; the audit reports ``true``, not a defensive ``unknown``.  The
    downgrade is tied to an *observed* drop, not to the mere presence of a
    case change — otherwise ``['A', 'a']`` above would lose its clean ``true``.
    """
    ta = audit(["a", "b", "E", "F"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


def test_real_regression_below_a_case_change_is_still_false():
    """``b, a, A, B`` — the drop has no case flip, so it is still asserted.

    ``b -> a`` (2 -> 1) happens inside one case; the later ``a -> A`` flip is
    flat (1 = 1).  Only a drop that *coincides* with a case change is excused,
    so a real regression keeps its ``false`` verdict.
    """
    ta = audit(["b", "a", "A", "B"])
    assert ta["monotonic"] is False
    assert ta["monotonic_verdict"] == "false"


# --------------------------------------------------------------------------- P1-b: ['1','01']
def test_leading_zero_decimal_is_the_same_family_and_value():
    """``['1', '01']``: both ``isdigit()`` -> ``decimal``; both read as 1.

    ``decimalZero`` is a distinct ``w:pgNumType/@fmt`` but the same *glyph
    family* as ``decimal``, and ``_token_number`` parses through ``int()``, so
    the leading zero changes neither the family nor the value: ``1, 1``, flat,
    monotonic, nothing unread.
    """
    ta = audit(["1", "01"])
    assert ta["page_tokens"] == ["1", "01"]         # the raw glyph is preserved
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0
    assert ta["page_parse_failures"] == []


def test_leading_zero_does_not_hide_a_real_regression():
    """``1, 2, 01`` is ``1, 2, 1`` — the padded token must still be compared."""
    ta = audit(["1", "2", "01"])
    assert ta["monotonic"] is False
    assert ta["monotonic_verdict"] == "false"


def test_zero_padded_increasing_sequence_is_monotonic():
    """``1, 02, 3`` is ``1, 2, 3``: padding is notation, not a jump."""
    ta = audit(["1", "02", "3"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"


def test_all_zero_padded_sequence_is_flat():
    """``1, 01, 001`` is three times 1 — flat, and therefore still monotonic."""
    ta = audit(["1", "01", "001"])
    assert ta["monotonic"] is True
    assert ta["monotonic_verdict"] == "true"
    assert ta["unparsed_tokens"] == 0


# ------------------------------------------------- ambiguous single letters still refuse
@pytest.mark.parametrize("tokens", [["C", "c"], ["I", "i"]])
def test_ambiguous_single_letters_are_counted_as_unread(tokens):
    """A run made only of single Roman-capable letters is judged, not asserted.

    ``c`` is 100 as a Roman numeral and 3 as a Latin sequence number;
    ``_family`` answers ``single`` for it and the run has no explicit member to
    settle the question, so the numbers are read as Latin (advisory) and the
    tokens are counted in ``unparsed_tokens`` -> verdict ``unknown``.  Refusing
    a verdict is the intended answer here, not a bug.
    """
    ta = audit(tokens)
    assert ta["unparsed_tokens"] == len(tokens)
    assert ta["monotonic_partial"] is True
    assert ta["monotonic_verdict"] == "unknown"