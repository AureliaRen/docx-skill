# -*- coding: utf-8 -*-
r"""A power of ten is a numeric anchor: the unit after it is "per", not a fraction.

Reported in round 4::

    c=3×10⁸m/s   ->  c=3\times\frac{10^{8}m}{s}

The ``/`` inside ``m/s`` there reads as "per" — ``m/s`` is a unit, not a
quotient — so the whole thing must become ``3\times 10^{8}\,\text{m/s}``.

Root cause: ``_find_dimension_spans`` runs *before* ``_escape_unicode``, so it
sees the source the way a person typed it: a literal ``⁸``.  ``_anchor_span_start``
only accepted three left neighbours — a digit, a closing ``)``, or a ``^{n}``
group — and a literal ``⁸`` is none of them.  ``_left_boundary_ok`` then rejected
the span too, because the character before ``m`` is ``⁸``, which ``isalnum()``.

The fix adds a run of literal superscript **digits** to ``_ANCHOR_RE``, which
makes the two spellings of an exponent agree: ``x⁸m/s`` and ``x^{8}m/s`` now
produce the same LaTeX.  Two spellings of one fact giving two different results
was the bug, not a design choice — ``tests/test_units_ascii.py`` already pins the
same invariant for ``kg/m3`` vs ``kg/m³``.

What this file guards, and what must not move:

  A. the anchor fires for a literal exponent, with and without whitespace
  B. it fires identically for the already-braced spelling (parity, not a new rule)
  C. real fractions survive untouched — ``E=V/m``, ``v=s/t``, ``p=F/A``,
     ``I=U/R``, ``n=m/M``, ``ρ=m/V``, ``a/b``, ``10/10k``
  D. a letter-base power that is *not* attached to a unit chain stays a fraction
  E. ``_left_boundary_ok`` is unchanged: ``=`` and letters still block a span
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


# ---------------------------------------------------------------- A. 字面上标锚点
LITERAL_ANCHOR_CASES = [
    # 主人报的原始 bug
    ("c=3\u00d710\u2078m/s", r"c=3\times 10^{8}\,\text{m/s}"),
    ("3\u00d710\u2078m/s", r"3\times 10^{8}\,\text{m/s}"),
    # 数字与单位之间那个空格要一起吃掉（不留旧实现没有的多余空格）
    ("3.0\u00d710\u2078 m/s", r"3.0\times 10^{8}\,\text{m/s}"),
    # 负指数：`⁻` 是符号、`⁷` 才是数字，run 从 `⁷` 起算仍然锚得住
    ("1.0\u00d710\u207b\u2077mol/L", r"1.0\times 10^{-7}\,\text{mol/L}"),
    # 多位指数
    ("6.02\u00d710\u00b2\u00b3mol/L", r"6.02\times 10^{23}\,\text{mol/L}"),
    # 裸幂次（前面没有 ×10 也成立）
    ("10\u2078m/s", r"10^{8}\,\text{m/s}"),
    # 有幂次的量纲串走 \mathrm（幂次是真上标），与既有形态一致
    ("3\u00d710\u2078m/s\u00b2", r"3\times 10^{8}\,\mathrm{m/s^{2}}"),
    ("1\u00d710\u207b\u00b3kg/m3", r"1\times 10^{-3}\,\mathrm{kg/m^{3}}"),
]


@pytest.mark.parametrize("src, expected", LITERAL_ANCHOR_CASES)
def test_literal_superscript_anchors_the_unit_span(src: str, expected: str):
    assert conv(src) == expected


@pytest.mark.parametrize("src", [c[0] for c in LITERAL_ANCHOR_CASES])
def test_anchored_span_is_not_rewritten_as_a_fraction(src: str):
    """命中的量纲串整段被摘走，`/` 不许落成 `\\frac`。"""
    out = conv(src)
    assert "\\frac" not in out, out
    assert "/" in out, out


@pytest.mark.parametrize("src", ["3\u00d710\u2078m/s", "c=3\u00d710\u2078m/s",
                                 "10\u2078m/s"])
def test_the_span_is_found_where_the_bug_reported_it_missing(src: str):
    """复现口径：修复前 `_find_dimension_spans('3×10⁸m/s')` 返回 `[]`。"""
    assert L.unit_slash_chains(src), src


# ---------------------------------------------------------------- B. 两种拼写等价
PARITY_PAIRS = [
    ("3\u00d710\u2078m/s", "3\u00d710^{8}m/s"),
    ("c=3\u00d710\u2078m/s", "c=3\u00d710^{8}m/s"),
    ("3.0\u00d710\u2078 m/s", "3.0\u00d710^{8} m/s"),
    ("10\u2078m/s", "10^{8}m/s"),
    ("x\u2078m/s", "x^{8}m/s"),
    ("v\u00b2m/s", "v^{2}m/s"),
    ("v\u00b2r/g", "v^{2}r/g"),
    ("1.0\u00d710\u207b\u2077mol/L", "1.0\u00d710^{-7}mol/L"),
]


@pytest.mark.parametrize("literal, braced", PARITY_PAIRS)
def test_literal_and_braced_exponent_agree(literal: str, braced: str):
    """字面 `⁸` 与写成 `^{8}` 是同一个事实，输出必须逐字节相同。

    `^{n}` 这一支本来就认任意基数（`x^{8}m/s` 在修复前就已经是量纲串），
    所以新增的是**字面形态的等价**，不是一条更宽的判据。
    """
    assert conv(literal) == conv(braced), (conv(literal), conv(braced))


# ---------------------------------------------------------------- C. 真分式不变
REAL_FRACTION_CASES = [
    ("E=V/m", r"E=\frac{V}{m}"),
    ("v=s/t", r"v=\frac{s}{t}"),
    ("p=F/A", r"p=\frac{F}{A}"),
    ("I=U/R", r"I=\frac{U}{R}"),
    ("n=m/M", r"n=\frac{m}{M}"),
    ("\u03c1=m/V", r"\rho =\frac{m}{V}"),
    ("a/b", r"\frac{a}{b}"),
    ("10/10k", r"\frac{10}{10k}"),
]


@pytest.mark.parametrize("src, expected", REAL_FRACTION_CASES)
def test_real_fractions_are_never_eaten_by_the_new_anchor(src: str, expected: str):
    """最尖锐的反向用例：两边全是合法单位符号，仍然必须读作分数线。"""
    assert conv(src) == expected


@pytest.mark.parametrize("src", [c[0] for c in REAL_FRACTION_CASES])
def test_no_dimension_span_is_reported_for_real_fractions(src: str):
    assert L.unit_slash_chains(src) == [], L.unit_slash_chains(src)


# ---------------------------------------------------------------- D. 幂次也要分场合
def test_power_without_a_unit_chain_stays_a_fraction():
    """`V²/m` 里没有数值锚点 —— 幂次的基数不是数字，锚点不成立。"""
    assert conv("V\u00b2/m") == r"\frac{V^{2}}{m}"


def test_braced_power_without_a_unit_chain_stays_a_fraction():
    assert conv("V^{2}/m") == r"\frac{V^{2}}{m}"


def test_lowercase_power_without_a_unit_chain_stays_a_fraction():
    assert conv("x\u00b2/m") == r"\frac{x^{2}}{m}"


# ---------------------------------------------------------------- E. 边界判据没动
def test_left_boundary_still_blocks_a_letter_and_an_equals_sign():
    """`_left_boundary_ok` 一个字没改：无锚点时 `=` 与字母照样否决。"""
    assert L._left_boundary_ok("E=V/m", 2) is False
    assert L._left_boundary_ok("xm/s", 1) is False


def test_left_boundary_still_allows_a_clean_start():
    assert L._left_boundary_ok("m/s", 0) is True


def test_letter_superscript_is_not_a_numeric_anchor():
    """`ⁿ`/`ⁱ` 不是数值，不能进锚点字符类（否则 `xⁿm/s` 会被当成测量值）。"""
    assert L._ANCHOR_RE.search("x\u207f") is None
    assert conv("x\u207fm/s") == r"\frac{x^{n}m}{s}"


def test_sign_alone_is_not_a_numeric_anchor():
    """`⁻` 只有符号没有数字，锚点不成立（`10⁻` 与 `10^{-}` 同口径）。

    这条只在 `_ANCHOR_RE` 层面断言：`10⁻m/s` 整体仍会被 B 分支
    （`m/s` 登记在 COMPOUND_UNITS + 左边是干净边界）接住 —— 那是**修复前
    就存在**的行为，与本轮改动无关，故不做输出层断言。
    """
    assert L._ANCHOR_RE.search("10\u207b") is None
    assert L._ANCHOR_RE.search("10\u2078") is not None


def test_lone_sign_case_is_unchanged_by_the_new_anchor():
    """`10⁻m/s` 走的是 B 分支，修复前后逐字节一致。"""
    assert conv("10\u207bm/s") == "10^{-}\\,\\text{m/s}"


# ---------------------------------------------------------------- F. 既有形态不回退
@pytest.mark.parametrize("src, must_contain", [
    ("1.2kg/m3", r"\mathrm{kg/m^{3}}"),
    ("a=9.8m/s2", r"\mathrm{m/s^{2}}"),
    ("k=8.31J/(mol.K)", r"\text{J/(mol.K)}"),
    ("\u03c1=1.2kg/m\u00b3", r"\mathrm{kg/m^{3}}"),
    ("v=9.8m/s\u00b2", r"\mathrm{m/s^{2}}"),
])
def test_existing_dimension_forms_are_unchanged(src: str, must_contain: str):
    """对照组：这些是修复前就成立的，一条都不许动。"""
    assert must_contain in conv(src), conv(src)
