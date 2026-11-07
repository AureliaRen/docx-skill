# -*- coding: utf-8 -*-
r"""Consecutive Unicode scripts must become **one** script group, not many.

Reported in round 4: ``10⁻⁷`` came out as ``10^{-}^{7}``.

The old ``_escape_unicode`` translated one character at a time, so ``⁻`` became
``^{-}`` (an empty exponent) and ``⁷`` started a *second* superscript right
after it.  ``a^{b}^{c}`` is not valid LaTeX — ``latex2mathml`` keeps the first
group and drops or demotes the rest, so the document showed ``10`` with an empty
exponent, or ``10`` plus a stray ``7`` on the baseline.

The fix folds a maximal run of ``_SUPER`` characters into one ``^{...}`` and a
maximal run of ``_SUB`` characters into one ``_{...}`` before the per-character
pass runs.  ``⁻⁷`` therefore reads as the single exponent ``-7``.

What must not change: one-character scripts (``²``, ``₂``) still produce the
same output as before, and already-braced LaTeX (``10^{-7}``) is not re-wrapped.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402


def conv(s: str) -> str:
    return L.linear_to_latex(s)


# ---------------------------------------------------------------- A. 上标 run
SUPERSCRIPT_RUN_CASES = [
    # 主人报的原始 bug：负号与数字是**同一个**指数
    ("10\u207b\u2077", "10^{-7}"),
    ("x\u207b\u00b2", "x^{-2}"),
    ("2\u207b\u00b9", "2^{-1}"),
    # 正号同规则
    ("x\u207a\u00b3", "x^{+3}"),
    # 多位数：每一位都是独立的 Unicode 字符，必须整体成组
    ("10\u00b2\u00b3", "10^{23}"),
    ("2\u207b\u00b9\u2070", "2^{-10}"),
    # 单个上标：与修复前逐字节相同，这是不许动的既有行为
    ("10\u2078", "10^{8}"),
    ("x\u00b2", "x^{2}"),
    # 上标与乘号混排：乘号先被改写，指数整体仍要完整
    ("1.0\u00d710\u207b\u2077", r"1.0\times 10^{-7}"),
    ("3\u00d710\u2078", r"3\times 10^{8}"),
]


@pytest.mark.parametrize("src, expected", SUPERSCRIPT_RUN_CASES)
def test_consecutive_superscripts_form_one_group(src: str, expected: str):
    assert conv(src) == expected


# ---------------------------------------------------------------- B. 下标 run
SUBSCRIPT_RUN_CASES = [
    ("v\u2081\u2082", "v_{12}"),
    ("x\u208b\u2081", "x_{-1}"),
    ("a\u2080\u2081", "a_{01}"),
    # 单个下标不变
    ("H\u2082O", "H_{2}O"),
    # 维度串里的字面下标：仍要走通量纲判定那条路
    ("v\u2081", "v_{1}"),
]


@pytest.mark.parametrize("src, expected", SUBSCRIPT_RUN_CASES)
def test_consecutive_subscripts_form_one_group(src: str, expected: str):
    assert conv(src) == expected


# ---------------------------------------------------------------- C. 上标 + 下标
def test_superscript_then_subscript_stay_separate_groups():
    """`x²₁` 是一个真上标加一个真下标，顺序不能颠倒、也不能并成一组。"""
    assert conv("x\u00b2\u2081") == "x^{2}_{1}"


def test_subscript_then_superscript_stay_separate_groups():
    assert conv("x\u2081\u00b2") == "x_{1}^{2}"


# ---------------------------------------------------------------- D. 不变量
NO_EMPTY_GROUP = re.compile(r"[\^_]\{\s*\}")


@pytest.mark.parametrize("src", [
    "10\u207b\u2077", "x\u207b\u00b2", "2\u207b\u00b9", "10\u00b2\u00b3",
    "v\u2081\u2082", "x\u207b\u00b2\u2081", "2\u207a\u00b9",
])
def test_no_empty_script_group_is_ever_emitted(src: str):
    r"""`^{-}^{7}` 这种形态的根因就是空组；空组一个都不许出现。"""
    assert not NO_EMPTY_GROUP.search(conv(src)), conv(src)


@pytest.mark.parametrize("src", ["10\u207b\u2077", "x\u207b\u00b2", "v\u2081\u2082"])
def test_a_script_is_never_followed_by_a_second_script(src: str):
    r"""同一个基数的两个脚本必须各自带 `^`/`_`，不许出现 `^{..}^{..}` 的并排。"""
    out = conv(src)
    assert "}^{" not in out, out
    assert "}_{" not in out, out


# ---------------------------------------------------------------- E. 已是 LaTeX
@pytest.mark.parametrize("src", ["10^{-7}", "10^{+3}", "x^{23}", "v_{12}"])
def test_existing_latex_scripts_are_not_rewrapped(src: str):
    """源文档已经写成 LaTeX 的，原样穿过去（与本轮改动前逐字节一致）。"""
    assert conv(src) == src


# ---------------------------------------------------------------- F. 与单位共存
def test_script_run_does_not_break_the_unit_grammar():
    """`m/s²` 这类量纲串的幂次读法不能被 run 合并挤掉。"""
    out = conv("v=9.8m/s\u00b2")
    assert out == r"v=9.8\,\mathrm{m/s^{2}}", out


def test_script_run_in_a_superscript_power_still_parses_as_exponent():
    """`_read_exponent` 早就按 run 读字面上标，这里锁的是 `_escape_unicode` 追平它。"""
    got = L._read_exponent("m/s\u207b\u00b2", 3)
    assert got is not None
    assert got[1] == "-2"
