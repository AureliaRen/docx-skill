# -*- coding: utf-8 -*-
"""A spacing command is one token, not a backslash plus punctuation.

``\\,`` is LaTeX for a thin space. It is *two characters* and contains no
letter, so ``_CMD_RE`` (``\\\\[A-Za-z]+``) cannot match it. ``_split_atoms``
fell through to its "backslash is a single atom" branch and produced

    ``\\,``  →  atom ``\\``  +  op ``,``

which is wrong twice over: the comma became an *operator* (so a fraction's
operand scan stopped there), and the bare backslash became an *adjacent factor*
(so it got pulled into the operand).

The visible damage needed a fraction and a unit string in the same span —
exactly what a hand-out writes all the time::

    ``1/2 m/s``
      _unitify  → ``1/2\\,\\text{m/s}``
      _fracify  → ``\\frac{1}{2\\},\\text{m/s}``      ← the denominator is ``2\\}``
      OMML       → numerator ``1``, denominator ``2},``, then ``m/s``

Word shows a stacked fraction whose denominator reads ``2},``. Nobody reports
it, because it looks like a fraction from across the desk and the file opens
fine.

The fix collects the five two-character spacing commands (``\\,`` ``\\;``
``\\:`` ``\\!`` ``\\ ``) as a single ``space`` token before any operand
scanning happens.

These tests pin both halves: the token shape, and the rendered LaTeX that the
token shape is supposed to protect. They also re-check the neighbours that
must not move — ``1/2×1/2`` still refuses to nest, and ``v=m/s`` is still a
fraction rather than a unit string.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402


def conv(s: str) -> str:
    return L.linear_to_latex(s)


# --- the token shape ----------------------------------------------------


@pytest.mark.parametrize("cmd", [r"\,", r"\;", r"\:", r"\!", "\\ "])
def test_spacing_command_is_a_single_space_token(cmd: str) -> None:
    """No stray ``op`` for the punctuation half, no stray ``atom`` for the
    backslash half. One ``space`` token and nothing else."""
    assert L._split_atoms(cmd) == [("space", cmd)]


def test_spacing_command_between_factors_does_not_join_the_operand() -> None:
    """``a\\,b`` stays three tokens; ``b`` must not glue onto ``a``."""
    toks = L._split_atoms(r"a\,b")
    assert toks == [("atom", "a"), ("space", r"\,"), ("atom", "b")], toks


def test_spacing_command_inside_text_group_is_left_alone() -> None:
    """A ``\\text{...}`` group is opaque — its content never gets tokenised."""
    toks = L._split_atoms(r"x\,\text{a b}")
    assert ("space", r"\,") in toks, toks
    assert ("atom", r"\text{a b}") in toks, toks


# --- the damage it used to cause ----------------------------------------

# Each of these pairs a fraction with a unit string. Before the fix the
# denominator absorbed the backslash and the LaTeX stopped being well-formed.
# The middle dot is written as a real character, not ``\u00b7`` — inside a raw
# string the escape would stay literal and the test would compare against text
# that can never appear.
FRACTION_THEN_UNIT = [
    ("1/2 m/s", r"\frac{1}{2}\,\text{m/s}", "2"),
    ("3/4 m/s", r"\frac{3}{4}\,\text{m/s}", "4"),
    ("1/2 rad/s", r"\frac{1}{2}\,\text{rad/s}", "2"),
    ("1/2 J/(mol\u00b7K)", "\\frac{1}{2}\\,\\text{J/(mol\u00b7K)}", "2"),
    ("\u00bd m/s", r"\frac{1}{2}\,\text{m/s}", "2"),
    ("1/2 kg/m\u00b3", r"\frac{1}{2}\,\mathrm{kg/m^{3}}", "2"),
]


@pytest.mark.parametrize("src,expected,_den", FRACTION_THEN_UNIT)
def test_fraction_before_a_unit_string_stays_well_formed(
        src: str, expected: str, _den: str) -> None:
    out = conv(src)
    assert out == expected, out
    # the specific deformity: a backslash escaping the denominator's brace
    assert r"\}," not in out, out
    assert out.count("{") == out.count("}"), out


@pytest.mark.parametrize("src,_expected,den_value", FRACTION_THEN_UNIT)
def test_fraction_before_a_unit_string_renders_two_operands(
        src: str, _expected: str, den_value: str) -> None:
    """The OMML denominator must be the bare number, not ``2},``."""
    from math_omml import latex_to_omml_lxml

    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    om = latex_to_omml_lxml(conv(src))
    frac = next(om.iter(M + "f"))
    den = frac.find(M + "den")
    assert den is not None
    den_text = "".join(t.text or "" for t in den.iter(M + "t"))
    assert den_text.strip() == den_value, den_text
    assert "}" not in den_text and "\\" not in den_text, den_text


# --- neighbours that must not move --------------------------------------


def test_multiplication_still_blocks_operand_absorption() -> None:
    """``_OP_CMD_NAMES`` and the spacing fix must not fight each other."""
    assert conv("1/2\u00d71/2") == r"\frac{1}{2}\times\frac{1}{2}"
    assert conv("2/3*4/5") == r"\frac{2}{3}*\frac{4}{5}"


def test_plain_fraction_is_untouched() -> None:
    assert conv("1/2") == r"\frac{1}{2}"
    assert conv("a/b/c") == r"\frac{\frac{a}{b}}{c}"


def test_unit_anchor_rule_is_untouched() -> None:
    """``v=m/s`` is a formula; ``m/s`` is a unit string. Unchanged."""
    assert conv("v=m/s") == r"v=\frac{m}{s}"
    assert conv("m/s") == r"\,\text{m/s}"
    assert conv("0.5 m/s") == r"0.5\,\text{m/s}"
    assert conv("10 m/s") == r"10\,\text{m/s}"


def test_brace_balance_holds_for_a_batch_of_spans() -> None:
    """Cheap invariant across the whole neighbourhood — every brace closes."""
    for src in ("1/2 m/s", "3/4 rad/s", "½ kg/m³", "x=½ m/s", "1/2×1/2",
                "v=m/s", "0.5 m/s", "a/b/c", "8.31 J/(mol·K)"):
        out = conv(src)
        assert out.count("{") == out.count("}"), (src, out)
