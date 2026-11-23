# -*- coding: utf-8 -*-
"""Body-ink measurement: the running header must not count as body text.

Why this file exists
--------------------
P17 added ``--gaps`` (absolute tail/head white in mm) and measured the corpus.
6 of 49 pages came back with a **negative** ``head_gap_mm``, all on page 1, all
between -8 and -10 mm.  The P17 report recorded that honestly but could not
explain it ("成因我没验证"), and left it as a known anomaly.

The cause was a one-sided block filter.  ``_body_blocks`` dropped blocks that
*start* inside the footer band and kept everything else, so a running header --
which sits at the very top of the page, above the text block -- was counted as
body ink.  ``head_gap_mm`` is measured as ``first_block_y - text_block_top``, so
a header at y = 43.2 pt against a 70.85 pt top margin yields a *negative* gap:

    source-snapshot/test_unpack/test_input.docx, page 1
    before:  head_gap_mm = -9.76      (the "TEST HEADER" block)
    after:   head_gap_mm = +9.56      (the "第一章 测试标题" block)

and 24.99 (top margin) + 9.56 = 34.55 mm, which is exactly the y of the real
first body line (97.9 pt / 2.835 = 34.55 mm).  The two numbers reconcile only
after the header is excluded, which is the check this file pins.

Two supporting behaviours are pinned here as well: ``top_margin_mm`` is reported
so that reconciliation is possible at all, and the last page's tail gap is a
*flag* rather than a default, because "a document ends where it ends" is a
judgement call and not something a checker should make silently.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docx_preview                                        # noqa: E402
from docx_preview import (                                 # noqa: E402
    _body_blocks,
    _head_gap_mm,
    _tail_gap_mm,
    _text_bottom_mm,
)

PT_PER_MM = 72.0 / 25.4


# ---- pure geometry: no PyMuPDF, no LibreOffice, runs everywhere -------------

def _block(y0, y1, text="x"):
    return {"x0": 0.0, "y0": y0, "x1": 100.0, "y1": y1, "text": text, "type": 0}


def test_head_gap_is_positive_when_the_body_starts_below_the_margin():
    body_top_pt = 70.85                     # 25 mm
    blocks = [_block(body_top_pt + 27.0, body_top_pt + 55.0)]
    gap = _head_gap_mm(blocks, body_top_pt)
    assert gap is not None
    assert gap > 0
    assert gap == pytest.approx(27.0 / PT_PER_MM, abs=0.01)


def test_a_header_block_would_produce_a_negative_gap_without_the_band():
    """The measured defect, reproduced as geometry.

    A header block (one that ends above the text-block top) must be removed; when
    it is not, the gap goes negative.  This asserts both halves so the fix cannot
    be "clamp the sign", which would hide the defect instead of reading it.
    """
    page_height = 841.89
    body_top_pt = 70.85
    header = _block(43.2, 58.7, "TEST HEADER")
    body = _block(97.9, 124.3, "第一章 测试标题")

    with_header = _head_gap_mm([header, body], body_top_pt)
    without = _head_gap_mm([body], body_top_pt)
    assert with_header is not None and with_header < 0, \
        "counting the header must be what produces the negative number"
    assert without is not None and without > 0
    assert with_header == pytest.approx(-9.76, abs=0.05)
    assert without == pytest.approx(9.56, abs=0.05)


def test_tail_gap_and_text_bottom_agree():
    body_bottom_pt = 841.89 - 70.85
    blocks = [_block(100.0, 300.0)]
    bottom = _text_bottom_mm(blocks)
    tail = _tail_gap_mm(blocks, body_bottom_pt)
    assert bottom is not None and tail is not None
    assert bottom + tail == pytest.approx(body_bottom_pt / PT_PER_MM, abs=0.05)


def test_empty_page_reports_none_not_zero():
    """\"nothing printed\" is not \"ink ends at 0 mm\"."""
    assert _text_bottom_mm([]) is None
    assert _tail_gap_mm([], 700.0) is None


# ---- the filter itself, with a stub page ------------------------------------

class _StubPage:
    """Minimal stand-in for a fitz Page: only ``rect`` and ``get_text``."""

    def __init__(self, height, blocks):
        self.rect = type("R", (), {"height": height, "width": 595.0})()
        self._blocks = blocks

    def get_text(self, _what):
        return self._blocks

    def get_text_dict(self, *_a, **_k):
        return {"blocks": []}


def test_body_blocks_drops_both_bands():
    height = 841.89
    body_top_pt = 70.85                     # 25 mm top margin
    # The footer cut is `FOOTER_BAND_START * height` ~= 740.9 pt, so the footer
    # block below is placed inside that band on purpose -- it is the *cut* being
    # tested, not a convenient coordinate.
    blocks = [
        (0.0, 43.2, 100.0, 58.7, "TEST HEADER", 0, 0),      # ends 12.15pt above
        (0.0, 97.9, 100.0, 124.3, "第一章", 1, 0),            # body
        (0.0, 760.0, 100.0, 780.0, "Page 1", 2, 0),          # footer band
    ]
    page = _StubPage(height, blocks)

    kept = _body_blocks(page, body_top_pt=body_top_pt)
    assert [b["text"] for b in kept] == ["第一章"]

    # The seam proves the header cut is what changed the answer.
    legacy = _body_blocks(page, header_band=False, body_top_pt=body_top_pt)
    assert [b["text"] for b in legacy] == ["TEST HEADER", "第一章"]


def test_without_a_top_margin_the_header_cut_is_skipped_not_guessed():
    """No margin -> old reading, so a header shows up as a *visible* negative gap.

    The asymmetry is deliberate: a missed header produces a negative
    ``head_gap_mm`` that a reader will question, whereas a body line deleted by
    a guessed cut is invisible in the report.
    """
    height = 841.89
    blocks = [
        (0.0, 43.2, 100.0, 58.7, "TEST HEADER", 0, 0),
        (0.0, 97.9, 100.0, 124.3, "第一章", 1, 0),
    ]
    page = _StubPage(height, blocks)
    assert [b["text"] for b in _body_blocks(page)] == ["TEST HEADER", "第一章"]


def test_a_body_line_near_the_top_is_not_mistaken_for_a_header():
    """A short top margin must not make the first body line disappear.

    A document with a 12.7 mm (36 pt) top margin puts its first line at 36 pt,
    which is *above* the 12% band (101 pt on A4).  A fraction-based cut would
    delete it; the margin-anchored cut keeps it because the line ends below the
    text-block top.  Measured on this exact geometry before the cut was
    re-anchored: ``assert [] == ['body line']``.
    """
    height = 841.89
    body_top_pt = 36.0
    blocks = [(0.0, 36.0, 100.0, 92.0, "body line", 0, 0)]
    page = _StubPage(height, blocks)
    assert [b["text"] for b in _body_blocks(page, body_top_pt=body_top_pt)] \
        == ["body line"]


def test_a_block_ending_exactly_at_the_text_top_is_body():
    """Boundary: ``y1 == body_top_pt`` is not *strictly* above it."""
    height = 841.89
    blocks = [(0.0, 40.0, 100.0, 70.85, "line at the boundary", 0, 0)]
    page = _StubPage(height, blocks)
    assert [b["text"] for b in _body_blocks(page, body_top_pt=70.85)] \
        == ["line at the boundary"]


def test_a_genuinely_tall_header_is_kept_rather_than_silently_dropped():
    """If a 'header' block reaches past the text-block top, it stays body ink.

    Pre-P18 behaviour, i.e. no regression: the cut errs towards keeping content.
    """
    height = 841.89
    blocks = [(0.0, 20.0, 100.0, 200.0, "big banner", 0, 0)]
    page = _StubPage(height, blocks)
    assert [b["text"] for b in _body_blocks(page, body_top_pt=70.85)] \
        == ["big banner"]


# ---- CLI surface -----------------------------------------------------------

def test_parser_exposes_the_two_new_switches():
    ap = docx_preview._parser()
    args = ap.parse_args(["in.docx", "--lint", "--gaps"])
    assert args.skip_last_page_gap is False     # default: report the last page
    assert args.no_header_band is False         # default: exclude headers

    args = ap.parse_args(["in.docx", "--lint", "--gaps",
                          "--skip-last-page-gap", "--no-header-band"])
    assert args.skip_last_page_gap is True
    assert args.no_header_band is True


def test_lint_pdf_accepts_the_new_keywords():
    """Signature check: the CLI passes both through by keyword."""
    import inspect
    params = inspect.signature(docx_preview.lint_pdf).parameters
    assert "tail_gap_skip_last" in params
    assert "header_band" in params
    assert params["tail_gap_skip_last"].default is False
    assert params["header_band"].default is True
