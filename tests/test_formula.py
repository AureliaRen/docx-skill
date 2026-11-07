# -*- coding: utf-8 -*-
"""Regression suite for the formula chain.

Locks down three behaviours so future rule changes cannot silently break them:

  A. linear notation -> LaTeX (52 cases, including the bugs found while
     pressure-testing against physics / calculus / linear-algebra material)
  B. merged sub/superscript repair (6 cases)
  C. span detection — what must NOT be treated as a formula, and what must

The OMML tests additionally need ``latex2mathml`` plus Microsoft's
``MML2OMML.XSL`` from a local Office install. When either is missing they skip
instead of failing, so a bare CI runner still exercises everything else.

Run:  python -m pytest tests/ -v
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402
import math_omml  # noqa: E402

# ---------------------------------------------------------------------------
# A. linear -> LaTeX
# ---------------------------------------------------------------------------
LATEX_CASES = [
    # original linear form,                expected LaTeX
    (r"a/b",                            r"\frac{a}{b}"),
    (r"6/(2+2)",                        r"\frac{6}{2+2}"),
    (r"10/10k",                         r"\frac{10}{10k}"),
    (r"I = U_{oc}/(R_{0} + R_{L})",     r"I = \frac{U_{oc}}{R_{0} + R_{L}}"),
    (r"U_{A} = (6/2 + 3)/(1/2 + 1/2)",  r"U_{A} = \frac{\frac{6}{2} + 3}{\frac{1}{2} + \frac{1}{2}}"),
    (r"U=IR",                           r"U=IR"),
    (r"ΣI=0",                           r"\sum I=0"),
    (r"f=50Hz",                         r"f=50\,\text{Hz}"),
    (r"i_{L}(0^{+}) = i_{L}(0^{-})",    r"i_{L}(0^{+}) = i_{L}(0^{-})"),
    # --- appended after pressure-testing against physics / calculus / linalg ---
    (r"lim_{x→0} (sinx)/x = 1",         r"\lim_{x\to 0} \frac{\sin x}{x} = 1"),
    (r"lim_{x→∞}(1 + 1/x)^{x} = e",     r"\lim_{x\to \infty}(1 + 1/x)^{x} = e"),
    (r"s = ∫_{a}^{b} √(1 + (y′)²) dx",  r"s = \int_{a}^{b} \sqrt{1 +(y')^{2}} dx"),
    (r"∫dx/(1 + x^{2}) = arctanx + C",
     r"\int \frac{dx}{1 + x^{2}} = \arctan x + C"),
    (r"tr(A) = λ_{1} + λ_{2}",          r"\mathrm{tr}(A) = \lambda_{1} + \lambda_{2}"),
    (r"diag(A, B)",                     r"\mathrm{diag}(A, B)"),
    (r"h = 6.63×10^{−34} J·s",          r"h = 6.63\times 10^{-34}\,\text{J·s}"),
    (r"R = 8.31 J/(mol·K)",             r"R = 8.31\,\text{J/(mol·K)}"),
    (r"A^{T}A = E",                     r"A^{T}A = E"),      # units must not eat the matrix name A
    (r"x^{2} = 4",                      r"x^{2} = 4"),
    # --- matrix / cases / multi-line environments: passed through verbatim ---
    (r"A = \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}",
     r"A = \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}"),
    (r"f(x) = \begin{cases} x^{2}, & x \ge 0 \\ -x, & x < 0 \end{cases}",
     r"f(x) = \begin{cases} x^{2}, & x \ge 0 \\ -x, & x < 0 \end{cases}"),
    (r"\begin{pmatrix} \frac{a}{b} & 0 \\ 0 & 1 \end{pmatrix}",
     r"\begin{pmatrix} \frac{a}{b} & 0 \\ 0 & 1 \end{pmatrix}"),
    # aligned/align without \tag -> starred form (latex2mathml inserts (1)(2) otherwise)
    (r"\begin{aligned} f &= x^{2} \\ &= y \end{aligned}",
     r"\begin{align*} f &= x^{2} \\ &= y \end{align*}"),
    (r"\begin{align} u &= v \\ m &= n \end{align}",
     r"\begin{align*} u &= v \\ m &= n \end{align*}"),
    # with \tag: left alone (docx_mathify handles numbering itself)
    (r"\begin{align} a &= b \tag{1} \\ c &= d \tag{2} \end{align}",
     r"\begin{align} a &= b \tag{1} \\ c &= d \tag{2} \end{align}"),
    (r"T = 2\pi\begin{vmatrix} a & b \\ c & d \end{vmatrix}",
     r"T = 2\pi\begin{vmatrix} a & b \\ c & d \end{vmatrix}"),
    # --- hand-written matrix text [a, b; c, d]  (';' rows, ',' columns) ---
    (r"[a, b; c, d]",                   r"\begin{bmatrix}a & b \\ c & d\end{bmatrix}"),
    (r"(1, 2; 3, 4)",                   r"\begin{pmatrix}1 & 2 \\ 3 & 4\end{pmatrix}"),
    (r"K = [0, -jωL; jωL, 0]",          r"K = \begin{bmatrix}0 & -j\omega L \\ j\omega L & 0\end{bmatrix}"),
    (r"[R₁, 0; 0, R₂]",                 r"\begin{bmatrix}R_{1} & 0 \\ 0 & R_{2}\end{bmatrix}"),
    (r"[1, 0, 0; 0, 1, 0; 0, 0, 1]",
     r"\begin{bmatrix}1 & 0 & 0 \\ 0 & 1 & 0 \\ 0 & 0 & 1\end{bmatrix}"),
]

# ---------------------------------------------------------------------------
# A2. merged sub/superscript repair
# ---------------------------------------------------------------------------
REPAIR_CASES = [
    (r"I^{2R}",      r"I^{2}R"),
    (r"R_{1R2}",     r"R_{1}R_{2}"),
    (r"\omega_{0L}", r"\omega_{0}L"),
    (r"\omega_{0CR}", r"\omega_{0}CR"),
    (r"U^{2}",       r"U^{2}"),        # must not change
    (r"10^{-6}",     r"10^{-6}"),      # must not change
]

# ---------------------------------------------------------------------------
# C. span detection
# ---------------------------------------------------------------------------
NO_SPAN = [
    "范围 = 前 5 周",        # CJK breaks it; no right-hand operand
    "b=", "n=", "=R",        # fragments
    "6V", "2Ω",              # bare quantities
    "2026-11-08",            # date
    "电工电子技术 · 电路分析部分",
    "[0, 1]",                # interval: single line, no ';', not a matrix
    "f(x, y)",               # function call: not a matrix
]
YES_SPAN = [
    "U=IR", "ΣI=0", "f=50Hz", "i_{L}(0^{+}) = i_{L}(0^{-})",
    "A = [a, b; c, d]",
]


def atoms_of(text: str) -> list[dict]:
    """Turn plain text into the atom list find_math_spans expects."""
    return [{"ch": c, "rpr": None, "script": None} for c in text]


# OMML 可用性判定收在 `tests/_omml.py` 一处（REVIEW-CODE-R4 应修 5）。
from _omml import needs_omml  # noqa: E402


# ---------------------------------------------------------------------------
# A. string level
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("linear,expected", LATEX_CASES)
def test_linear_to_latex(linear, expected):
    assert L.linear_to_latex(linear) == expected


@pytest.mark.parametrize("source,expected", REPAIR_CASES)
def test_repair_merged_scripts(source, expected):
    got, _fixes = L.repair_merged_scripts(source)
    assert got == expected


# ---------------------------------------------------------------------------
# C. span detection
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text", NO_SPAN)
def test_text_is_not_treated_as_formula(text):
    assert not L.find_math_spans(atoms_of(text))


@pytest.mark.parametrize("text", YES_SPAN)
def test_text_is_treated_as_formula(text):
    assert L.find_math_spans(atoms_of(text))


# ---------------------------------------------------------------------------
# B. OMML legality
# ---------------------------------------------------------------------------
@needs_omml
@pytest.mark.parametrize("linear,_expected", LATEX_CASES)
def test_latex_converts_to_legal_omml(linear, _expected):
    latex = L.linear_to_latex(linear)
    # latex_to_omml_lxml runs assert_legal() internally and raises on illegal
    # structure such as a nested <m:oMath> or an <m:rad> missing <m:e>.
    omml = math_omml.latex_to_omml_lxml(latex)
    assert omml is not None
    assert omml.tag.rsplit("}", 1)[-1] in ("oMath", "oMathPara")


@needs_omml
@pytest.mark.parametrize("tex", [
    r"\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
    r"\sum_{i=1}^{n} x_i",
    r"\int_{0}^{\infty} e^{-x} dx",
    r"A = \begin{bmatrix}1 & 2\\3 & 4\end{bmatrix}",
    r"\text{中文}",
])
def test_complex_formula_shapes(tex):
    assert math_omml.latex_to_omml_lxml(tex) is not None


# ---------------------------------------------------------------------------
# Font injection
# ---------------------------------------------------------------------------
@needs_omml
def test_formula_font_defaults_to_times_new_roman():
    M_NS, W_NS = math_omml.M_NS, math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("U=IR")
    runs = omml.findall(f".//{{{M_NS}}}r")
    assert runs, "conversion produced no runs"

    for run in runs:
        rfonts = run.find(f"{{{W_NS}}}rPr/{{{W_NS}}}rFonts")
        assert rfonts is not None, "run missing w:rPr/w:rFonts"
        assert rfonts.get(f"{{{W_NS}}}ascii") == "Times New Roman"

    # CT_R content-model order: w:rPr must come before m:t
    for run in runs:
        assert [c.tag.rsplit("}", 1)[-1] for c in run][-1] == "t"


@needs_omml
def test_formula_font_env_override(monkeypatch):
    monkeypatch.setenv("MATH_FONT", "Cambria Math")
    W_NS = math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("x")
    rfonts = omml.find(f".//{{{W_NS}}}rFonts")
    assert rfonts is not None
    assert rfonts.get(f"{{{W_NS}}}ascii") == "Cambria Math"


@needs_omml
def test_formula_font_can_be_disabled(monkeypatch):
    monkeypatch.setenv("MATH_FONT", "")
    W_NS = math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("x")
    assert omml.find(f".//{{{W_NS}}}rFonts") is None
