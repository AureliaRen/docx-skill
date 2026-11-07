# -*- coding: utf-8 -*-
"""Unit strings typed the way a person actually types them.

``test_units.py`` pins the canonical forms — ``kg/m³`` with a real Unicode
superscript, ``J/(mol·K)`` with a real middle dot. Both are correct, and both
are awkward to produce from a Chinese keyboard: ``³`` has no key, and ``·``
lives behind a symbol picker. So the forms that actually reach a document are

    ``kg/m3``      a bare ASCII digit where the exponent belongs
    ``J/(mol.K)``  an ASCII full stop where the product sign belongs

Both used to fall through to the fraction rewriter and came out as
``\\frac{1.2kg}{m3}`` / ``\\frac{\\text{J}}{mol.K}`` — physically wrong, since
``/`` in a unit string means "per", not "divided by".

The regression risk runs the other way too: widening the unit grammar must not
swallow real fractions. ``v=s/t``, ``p=F/A`` and ``ρ=m/V`` have to stay
fractions, and ``kg.m`` outside parentheses must not become a dimension string
just because ``.`` is now a product sign inside ``( )``.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from linear_to_latex import (                                 # noqa: E402
    is_dimension_string, linear_to_latex, unit_slash_chains,
)

# --- bare ASCII exponents ------------------------------------------------

BARE_EXPONENT_CASES = [
    ("rho=1.2kg/m3", "kg/m"),
    ("a=9.8m/s2", "m/s"),
    ("p=1.0g/cm3", "g/cm"),
    ("P=50W/m2", "W/m"),
    ("flux=50mol/L2", "mol/L"),
]

# --- ASCII full stop as a product sign inside ( ) ------------------------

DOT_PRODUCT_CASES = [
    "k=8.31J/(mol.K)",
    "R=8.31J/(kg.K)",
    "c=4.18kJ/(kg.K)",
]


@pytest.mark.parametrize("src,_needle", BARE_EXPONENT_CASES)
def test_bare_ascii_exponent_reads_as_power(src: str, _needle: str):
    """``kg/m3`` is ``kg/m³``: a dimension string, never a fraction."""
    out = linear_to_latex(src)
    assert "\\frac" not in out, out
    assert "^{" in out, out          # the exponent became a real superscript
    assert "\\mathrm{" in out or "\\text{" in out, out


@pytest.mark.parametrize("src", DOT_PRODUCT_CASES)
def test_ascii_dot_inside_parens_is_a_product_sign(src: str):
    """``J/(mol.K)`` is ``J/(mol·K)``, not ``J`` over ``mol.K``.

    What matters is that the denominator stayed *with* the numerator inside one
    unit string: the span is the whole ``J/(...K)``, not ``\\frac{J}{...}``.
    """
    out = linear_to_latex(src)
    assert "\\frac" not in out, out
    # the denominator text survives verbatim, dot included
    rhs = src.split("=", 1)[1]
    denom = rhs[rhs.index("("):]
    assert denom.replace(" ", "") in out.replace(" ", ""), out


def test_bare_exponent_matches_unicode_superscript():
    """The ASCII and Unicode spellings must agree exactly."""
    assert linear_to_latex("rho=1.2kg/m3") == linear_to_latex("rho=1.2kg/m\u00b3")
    assert linear_to_latex("a=9.8m/s2") == linear_to_latex("a=9.8m/s\u00b2")


# --- regressions: real fractions must survive ----------------------------

REAL_FRACTIONS = ["v=s/t", "p=F/A", "rho=m/V", "1/2", "w=2pi/t", "e=Q/U"]


@pytest.mark.parametrize("src", REAL_FRACTIONS)
def test_real_fractions_still_become_frac(src: str):
    assert "\\frac" in linear_to_latex(src), linear_to_latex(src)


def test_bare_dot_outside_parens_is_not_a_dimension_string():
    """``kg.m`` has no product-sign reading at top level, so it stays alone."""
    out = linear_to_latex("kg.m")
    assert "\\mathrm{" not in out and "\\text{" not in out, out
    assert not is_dimension_string("kg.m")


def test_decimal_point_is_not_eaten_as_exponent():
    """``x=1.5kg`` must keep its decimal point and its value."""
    out = linear_to_latex("x=1.5kg")
    assert "1.5" in out, out


def test_long_digit_run_after_unit_is_rejected():
    """``m30`` is not silently mangled into ``m^{30}`` when glued to a letter."""
    out = linear_to_latex("x=2m30y")
    assert "^{30}" not in out, out


@pytest.mark.parametrize("src,needle", BARE_EXPONENT_CASES)
def test_unit_slash_chains_reports_the_span(src: str, needle: str):
    chains = unit_slash_chains(src)
    assert chains, src
    assert any(needle in c for c in chains), chains
