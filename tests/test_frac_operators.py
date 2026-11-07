# -*- coding: utf-8 -*-
"""A fraction next to a multiplication must not absorb the next factors.

Reported by two subject agents from opposite directions:

* biology saw ``1/2×1/2`` become a nested fraction swallowing the ``×``
* math saw the same over-greedy operand scan on its formula paragraphs

Cause: the multiplication signs are rewritten to ``\\times`` / ``\\cdot`` by
``_escape_unicode`` *before* ``_fracify`` runs, and ``_split_atoms`` treated that
command as an ordinary adjacent atom — so the denominator kept absorbing factors
past the operator.

These tests pin the operator boundary and make sure the real-fraction and
dimension-string behaviour did not regress in the process.
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
    # 乘号两侧各是一整个分式 —— 修复前是嵌套分式
    ("1/2×1/2", r"\frac{1}{2}\times\frac{1}{2}"),
    ("3/4×1/4", r"\frac{3}{4}\times\frac{1}{4}"),
    ("2/3*4/5", r"\frac{2}{3}*\frac{4}{5}"),
    ("1/2*1/2", r"\frac{1}{2}*\frac{1}{2}"),
])
def test_multiplication_does_not_join_denominators(src, expected):
    assert conv(src) == expected


@pytest.mark.parametrize("src", ["2×3", "a×b"])
def test_plain_multiplication_unchanged(src):
    assert "\\times" in conv(src)
    assert "\\frac" not in conv(src)


@pytest.mark.parametrize("src, expected", [
    ("1/4", r"\frac{1}{4}"),
    ("v=s/t", r"v=\frac{s}{t}"),
    ("p=F/A", r"p=\frac{F}{A}"),
    ("rho=m/V", r"rho=\frac{m}{V}"),
    ("dp/dt", r"\frac{dp}{dt}"),
    ("1/2", r"\frac{1}{2}"),
])
def test_real_fractions_still_convert(src, expected):
    assert conv(src) == expected


@pytest.mark.parametrize("src, must_contain", [
    ("a=9.8m/s2", r"\mathrm{m/s^{2}}"),
    ("rho=1.2kg/m3", r"\mathrm{kg/m^{3}}"),
    ("P=100W/m2", r"\mathrm{W/m^{2}}"),
    ("c=0.5mol/L", r"\text{mol/L}"),
    ("k=8.31J/(mol.K)", r"\text{J/(mol.K)}"),
])
def test_unit_slashes_still_read_as_per(src, must_contain):
    out = conv(src)
    assert must_contain in out
    assert "\\frac" not in out, out


def test_no_stray_space_inside_frac_operands():
    """A command-space sentinel must not leak into the \\frac braces."""
    out = conv("1/2×1/2")
    assert "{ 1}" not in out and "{ 2}" not in out, out


def test_omega_keeps_its_command_separator():
    """An omega needs its command separator or it becomes the undefined omegaL."""
    out = conv("X=ωL/R")
    assert "\\omega L" in out or "\\omegaL" not in out, out
