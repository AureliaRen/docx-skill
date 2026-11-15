# -*- coding: utf-8 -*-
"""Vulgar fraction characters (½ ⅓ ⅕ …) must never survive as glyphs.

``½`` is one character (U+00BD) that fonts render as a tiny raised ``1``, a
slash, and a lowered ``2``. In a physics hand-out that glyph is meant to be a
stacked fraction, and Word's own Equation editor renders it that way — so a
document that keeps the raw glyph looks *nearly* right in one font and wrong in
every other, which is the worst failure mode: nobody notices until print.

Two separate gaps used to let it through:

1. Only five of the nineteen fraction characters were in the translation table
   (``½ ¼ ¾ ⅓ ⅔``). The other fourteen — ``⅕ ⅛ ⅐ ↉`` and friends — were not even
   recognised as math characters, so they stayed inline in body text forever.

2. Even the five that *did* translate never got the chance, because a span has
   to earn its way into the fraction rewriter by containing a relation operator
   (``=``, ``≈``) or a strong operator (``√``, ``∑``). A bare ``½+1`` has
   neither, so it was skipped — while ``x=½`` worked. Same character, same
   paragraph, two different outcomes.

The fix treats a fraction character as its own sufficient evidence, on the same
footing as ``√``: it is a math glyph the author typed on purpose, not a guess.

The regression risk runs the other way, and these tests pin it down:

* a fraction character is also a *value anchor* for unit strings, so ``½ m/s``
  must read as "half a metre per second" (``\\frac{1}{2}\\,\\text{m/s}``) and not
  as a fraction divided by ``m/s``;
* dates and chapter numbers that merely sit next to a slash must stay text.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from linear_to_latex import (                                  # noqa: E402
    VULGAR_FRAC,
    find_math_spans,
    is_math_char,
    linear_to_latex,
)

# --- the whole Unicode fraction block ------------------------------------

# Every vulgar fraction that Unicode assigns a single code point, with the
# n/d it stands for. Hand-written rather than derived from VULGAR_FRAC, so that
# a typo in the table cannot make the test agree with the bug.
ALL_FRACTIONS = {
    "\u00bc": (1, 4), "\u00bd": (1, 2), "\u00be": (3, 4),
    "\u2150": (1, 7), "\u2151": (1, 9), "\u2152": (1, 10),
    "\u2153": (1, 3), "\u2154": (2, 3),
    "\u2155": (1, 5), "\u2156": (2, 5), "\u2157": (3, 5), "\u2158": (4, 5),
    "\u2159": (1, 6), "\u215a": (5, 6),
    "\u215b": (1, 8), "\u215c": (3, 8), "\u215d": (5, 8), "\u215e": (7, 8),
    "\u2189": (0, 3),
}


def _atoms(s: str) -> list[dict]:
    """A flat atom list, the shape ``find_math_spans`` consumes."""
    return [{"ch": c, "rpr": None, "script": None} for c in s]


def _spans_text(s: str) -> list[str]:
    return [s[a:b] for a, b in find_math_spans(_atoms(s))]


# --- 1. table completeness ----------------------------------------------


def test_table_covers_every_unicode_fraction() -> None:
    """The translation table has exactly the 19 fraction code points."""
    assert set(VULGAR_FRAC) == set(ALL_FRACTIONS)


@pytest.mark.parametrize("ch,n_d", sorted(ALL_FRACTIONS.items()))
def test_single_fraction_becomes_frac(ch: str, n_d: tuple[int, int]) -> None:
    """A lone fraction character is a complete formula on its own."""
    num, den = n_d
    out = linear_to_latex(ch)
    assert out == rf"\frac{{{num}}}{{{den}}}", out


@pytest.mark.parametrize("ch", sorted(ALL_FRACTIONS))
def test_fraction_char_is_a_math_char(ch: str) -> None:
    """Fourteen of these used to return False and stayed in body text."""
    assert is_math_char(ch), f"U+{ord(ch):04X} not recognised as math"


# --- 2. the span gate ---------------------------------------------------

# Inputs with no relation operator at all. Before the fix these were skipped
# wholesale, so the fraction came out as a glyph in body text.
NO_OPERATOR_CASES = [
    ("\u00bd", "\u00bd"),
    ("\u00bd+1", "\u00bd+1"),
    ("\u00bd\u00d7\u00bc", "\u00bd\u00d7\u00bc"),
    ("\u00bd\u00bd", "\u00bd\u00bd"),
    ("\u2155", "\u2155"),
    ("\u215b", "\u215b"),
]


@pytest.mark.parametrize("src,expect", NO_OPERATOR_CASES)
def test_fraction_alone_opens_a_span(src: str, expect: str) -> None:
    """The fraction is the evidence; no ``=`` needed to justify it."""
    assert _spans_text(src) == [expect]


def test_fraction_inside_prose_still_converts() -> None:
    """``第½章`` — the fraction is the only math in the run; convert just it."""
    assert _spans_text("第\u00bd章") == ["\u00bd"]


@pytest.mark.parametrize("src,needle", NO_OPERATOR_CASES)
def test_fraction_converts_end_to_end(src: str, needle: str) -> None:
    out = linear_to_latex(src)
    assert "\\frac{" in out, out
    # the raw glyph must be gone — this is the whole point
    assert not any(ch in out for ch in ALL_FRACTIONS), out


# --- 3. fraction characters are value anchors for units -----------------

# ``½ m/s`` and ``0.5 m/s`` are the same measurement. Both have to reach the
# unit-string path, or two adjacent lines of one hand-out disagree.
ANCHORED_UNIT_CASES = [
    ("\u00bd m/s", "\\text{m/s}"),
    ("\u00bd kg/m\u00b3", "\\mathrm{kg/m^{3}}"),
    ("\u00bd J/(mol\u00b7K)", "\\text{J/(mol\u00b7K)}"),
    ("\u00bd N\u00b7m", "\\text{N\u00b7m}"),
]


@pytest.mark.parametrize("src,unit_tex", ANCHORED_UNIT_CASES)
def test_fraction_anchors_a_unit_string(src: str, unit_tex: str) -> None:
    """``½ m/s`` is a measurement, not ``½`` divided by ``m/s``."""
    out = linear_to_latex(src)
    assert out == r"\frac{1}{2}\," + unit_tex, out


def test_decimal_measure_of_the_same_value_matches() -> None:
    """The 0.5 spelling is the control: both must reach the unit path."""
    assert linear_to_latex("0.5 m/s") == "0.5\\,\\text{m/s}"


def test_fraction_without_unit_stays_a_plain_fraction() -> None:
    """No slash after it, so there is no unit chain to anchor."""
    assert linear_to_latex("\u00bd kg") == "\\frac{1}{2} kg"
    assert linear_to_latex("\u00bd m") == "\\frac{1}{2} m"


# --- 4. what must NOT change -------------------------------------------

NON_MATH_CASES = [
    "2024/1/2",
    "2024\u5e741/2\u5b63\u5ea6",
    "1/2 m/s",          # already a fraction; must stay one
    "0.5 m/s",
    "",
]


@pytest.mark.parametrize("src", NON_MATH_CASES)
def test_non_math_text_is_not_turned_into_a_formula(src: str) -> None:
    """Dropping the ``len >= 2`` floor for fractions must not open the gate
    for anything else: these have no fraction character at all."""
    assert _spans_text(src) == []


def test_real_fractions_are_untouched() -> None:
    """Widening the character set must not disturb the slash rewriter."""
    assert linear_to_latex("v=s/t") == "v=\\frac{s}{t}"
    assert linear_to_latex("p=F/A") == "p=\\frac{F}{A}"
    assert linear_to_latex("1/2") == "\\frac{1}{2}"


def test_dimension_strings_are_untouched() -> None:
    """The unit grammar still owns bare unit strings."""
    assert linear_to_latex("kg/m\u00b3") == "\\,\\mathrm{kg/m^{3}}"
    assert linear_to_latex("8.31 J/(mol\u00b7K)") == "8.31\\,\\text{J/(mol\u00b7K)}"
