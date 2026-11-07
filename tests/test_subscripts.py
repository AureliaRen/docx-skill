# -*- coding: utf-8 -*-
"""`v0` → `v_{0}`：variable-plus-digit means a subscript, not a product.

Reported by the physics agent: `v0` came out as two same-height characters, so
`v₀²` rendered as `v0²` — physically wrong (it reads as v times 0).  The source
document almost never carries real subscript runs for these; the plain text is
`v0`, `R1`, `U1`, `m2`.

The rule has to live next to chemistry and function names, which travel the same
code path, so these tests pin **both** directions:

* subscripts are added for variables (`v0`, `m1v1`, `R1=10`);
* chemical formulas (`H2O`, `C6H12O6`), function calls (`\\log2x`) and
  magnitudes (`x10`, `m100`) are left exactly as they were.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402


def conv(s):
    return L.linear_to_latex(s)


@pytest.mark.parametrize("src, expected", [
    ("v0", r"v_{0}"),
    ("v0=at", r"v_{0}=at"),
    ("x0=1m", r"x_{0}=1m"),
    ("2v0", r"2v_{0}"),
    ("v0^2", r"v_{0}^{2}"),
    ("a1,a2", r"a_{1},a_{2}"),
    # 动量守恒：四个变量全要带下标
    ("m1v1+m2v2=m1u1+m2u2",
     r"m_{1}v_{1}+m_{2}v_{2}=m_{1}u_{1}+m_{2}u_{2}"),
])
def test_lowercase_variable_gets_subscript(src, expected):
    assert conv(src) == expected


@pytest.mark.parametrize("src, expected", [
    ("R1=10", r"R_{1}=10"),
    ("U1=IR1", r"U_{1}=IR1"),
    ("T1=2pi", r"T_{1}=2pi"),
    ("U2=", r"U_{2}="),
])
def test_uppercase_variable_gets_subscript_before_operator(src, expected):
    """Uppercase is only rewritten when the digit is followed by an operator."""
    assert conv(src) == expected


@pytest.mark.parametrize("src", [
    "H2O", "CO2", "H2SO4", "C6H12O6", "Na2CO3",
])
def test_chemical_formulas_are_untouched(src):
    """Guessing which digit is a subscript would corrupt chemistry."""
    assert conv(src) == src


@pytest.mark.parametrize("src", [
    r"\lg2", r"\log2x", r"\sin2x",
])
def test_function_calls_are_untouched(src):
    out = conv(src)
    assert "_{" not in out, out


@pytest.mark.parametrize("src", ["x10", "m100", "x12", "2026", "e2", "E1"])
def test_magnitudes_and_numbers_are_untouched(src):
    assert conv(src) == src


@pytest.mark.parametrize("src, must_contain", [
    ("a=9.8m/s2", r"\mathrm{m/s^{2}}"),
    ("rho=1.2kg/m3", r"\mathrm{kg/m^{3}}"),
])
def test_unit_slash_rule_still_wins(src, must_contain):
    """Unit strings must not be rewritten into subscripts."""
    out = conv(src)
    assert must_contain in out, out
    assert "_{" not in out, out


def test_does_not_double_apply_existing_subscript():
    """An already-subscripted source (`m_{1}`) must stay as-is."""
    assert conv(r"m_{1}v_{1}") == r"m_{1}v_{1}"
