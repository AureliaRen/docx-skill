# -*- coding: utf-8 -*-
r"""Bare `^` / `_` must swallow a **signed number token**, not a single char.

Reported by the owner: `10^-7` rendered with an apparent gap between `-` and
`7`.  The old rule was ``\^([^\s{])`` — one character only — so ``10^-7``
became ``10^{-}7``: the minus sign went into the superscript while the ``7``
stayed on the baseline.  Two different baselines read as a space.

The new token is ``[-+−]?\d+(?:\.\d+)?`` *or* ``[-+−]?[A-Za-z]`` — a pure
number token or a single letter, never a span that crosses into the next
letter.  That last part is what keeps ``a^2b`` at ``a^{2}b`` instead of the
wrong ``a^{2b}``; both directions are pinned below.
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


# ---------------------------------------------------------------- 正例：符号指数
@pytest.mark.parametrize("src, expected", [
    # 主人报的原始 bug：负号必须与数字同进上标
    ("10^-7", "10^{-7}"),
    ("10^+3", "10^{+3}"),
    # U+2212 真减号（源文档里两种减号都有）
    ("10^\u22127", "10^{-7}"),
    # 科学记数法整串：乘号先被 _escape_unicode 改写，幂次仍要完整
    ("1.0\u00d710^-7", r"1.0\times 10^{-7}"),
    ("3\u00d710^+8", r"3\times 10^{+8}"),
    # 多位数指数（旧实现只会吃第一位，得到 10^{1}0）
    ("2^10", "2^{10}"),
    ("e^-12", "e^{-12}"),
    # 小数指数
    ("x^2.5", "x^{2.5}"),
])
def test_signed_and_multidigit_exponents(src, expected):
    assert conv(src) == expected


# ---------------------------------------------------------------- 反例：不许变宽
@pytest.mark.parametrize("src, expected", [
    # ★ 主人点名要锁死的：数字后紧跟字母时，字母必须留在正文
    ("a^2b", "a^{2}b"),
    ("x^2n", "x^{2}n"),
    ("2^3x", "2^{3}x"),
    # 单字母指数
    ("x^n", "x^{n}"),
    ("a^b", "a^{b}"),
])
def test_exponent_never_crosses_into_the_next_letter(src, expected):
    assert conv(src) == expected


# ---------------------------------------------------------------- 已经是 LaTeX
@pytest.mark.parametrize("src", [
    "10^{-7}",
    "10^{+3}",
    "x^{2}",
    "x^{n}",
    "a^{2}b",
])
def test_existing_latex_scripts_are_not_rewrapped(src):
    """`{` 挡住两个分支，所以已包裹的内容原样穿过去（不二次包裹）。"""
    assert conv(src) == src


# ---------------------------------------------------------------- 下划线同规则
@pytest.mark.parametrize("src, expected", [
    ("v_0", "v_{0}"),
    ("H_2O", "H_{2}O"),
    # 下划线的符号/多位数 token
    ("I_-1", "I_{-1}"),
    ("x_10", "x_{10}"),
    # 反例：不许跨字母
    ("m_1v_1", "m_{1}v_{1}"),
])
def test_subscript_token_follows_the_same_rule(src, expected):
    assert conv(src) == expected


# ---------------------------------------------------------------- 组合场景
def test_sign_and_unit_slash_still_coexist():
    """符号指数修好后，单位斜杠规则不能被挤掉（HANDOFF 前一轮成果）。"""
    out = conv("a=9.8m/s2")
    assert r"\mathrm{m/s^{2}}" in out, out


def test_scientific_notation_with_unit():
    out = conv("c=3\u00d710^8m/s")
    assert "^{8}" in out, out
