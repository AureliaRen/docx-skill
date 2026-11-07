# -*- coding: utf-8 -*-
r"""补齐 `test_unicode_scripts.py` / `test_dimension_anchor.py` 的边界取样缺口。

`REVIEW-CODE-R4` 应修 5 点名：「边界取样只挑恰好通过的 `x`/`t`」。
本文件补三类真边界：

**A. 连续上下标 run 的长度边界**（1/2/3/4 字符）
   原测试测了 `10⁻⁷`（2 字符）、`Fe^{3+}`，但没测长度 1 与长度 4。
   长度边界是 run 合并正则 `[...]+` 最容易写错的地方 ——
   写 `?` 就退化成单字符、写 `{2,}` 就漏单字符。

**B. 上下标混排**
   `x⁻²⁺` 这种上标里既有负号又有正号的；`SO_{4}^{2-}` 这种上下标都有的。

**C. 字面 `⁸` vs `^{8}` 等价性**
   对同一批输入逐条比对两种拼写。这条与 `test_super_unit_parity.py` 互补：
   那边测单位，这边测**不带单位**的纯幂次与化学电荷。

**D. 真分式保护的反向边界**
   「本该是分式的，不许因为新锚点变成量纲串」的相邻反例。
"""
from __future__ import annotations

import itertools
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402

SUPERS = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻ⁿⁱ"
SUBS = "₀₁₂₃₄₅₆₇₈₉₊₋"

# 非法形态标记：这些一旦出现，说明 run 合并把一组拆成了两组
ILLEGAL = ("^{-}^{", "^{}^{", "_{}_{", "^{}", "_{}")


def _illegal(out: str) -> list[str]:
    return [m for m in ILLEGAL if m in out]


# ---------------------------------------------------------------- A. 长度边界
@pytest.mark.parametrize("length", [1, 2, 3, 4])
def test_superscript_run_of_every_length_stays_one_group(length):
    """★ 连续上标 run 的**长度边界**：1~4 字符都必须收成一个成组。

    长度 1 是最容易漏的（正则写成 `{2,}` 就坏）；
    长度 4 是长指数（`10⁻¹⁰²³` 这类）。
    """
    src = "10" + "⁸" * length
    out = L.linear_to_latex(src)
    bad = _illegal(out)
    assert not bad, f"长度 {length}（{src!r}）出现非法形态 {bad}：{out!r}"
    assert out.count("^{") == 1, f"长度 {length} 被拆成多组：{out!r}"


@pytest.mark.parametrize("length", [1, 2, 3, 4])
def test_subscript_run_of_every_length_stays_one_group(length):
    """连续**下标** run 的长度边界：与上标组同构（正则写坏时两处一起坏）。"""
    src = "v" + "₁" * length
    out = L.linear_to_latex(src)
    bad = _illegal(out)
    assert not bad, f"长度 {length}（{src!r}）出现非法形态 {bad}：{out!r}"


def test_single_character_output_is_byte_identical_to_bare_form():
    """★ 单字符字面上标 == 花括号形态（逐字节）。

    R4 加 run 合并时的回归点是：`10⁻⁷` 会被逐字符拆成 `10^{-}^{7}`（非法）。
    这条锁的是它的温和面 —— 单字符输入（`10⁸` / `x²` / `v₁`）加了合并逻辑后
    输出必须**一个字节都不变**，仍等于作者直接写 `10^{8}` 的结果。
    """
    assert L.linear_to_latex("10⁸") == L.linear_to_latex("10^{8}")
    assert L.linear_to_latex("x²") == L.linear_to_latex("x^{2}")
    assert L.linear_to_latex("v₁") == L.linear_to_latex("v_{1}")


# ---------------------------------------------------------------- B. 上下标混排
@pytest.mark.parametrize("src", [
    "x⁻²", "x²⁻", "x⁻⁷⁸", "x⁸⁹⁰",
    "SO_{4}^{2-}", "Fe^{3+}", "MnO_{4}^{-}",
    "x_{1}^{2}", "a_{n+1}",
])
def test_mixed_scripts_never_produce_illegal_forms(src):
    """上下标混排：上标里带正负号、上下标都有的化学式，都不许出非法形态。"""
    out = L.linear_to_latex(src)
    bad = _illegal(out)
    assert not bad, f"{src!r} → {out!r} 含非法形态 {bad}"


@pytest.mark.parametrize("src,expect_fragment", [
    ("Fe^{3+}", "3+"),
    ("SO_{4}^{2-}", "2-"),
    ("MnO_{4}^{-}", "-"),
])
def test_chemical_charge_merges_sign_into_the_group(src, expect_fragment):
    """化学电荷：电荷符号必须并进**同一个**幂次组，不许单开一个 `^{+}`。"""
    out = L.linear_to_latex(src)
    assert "^{" in out
    assert r"^{+" not in out, f"{src!r} 的 + 被单独成组：{out!r}"
    assert expect_fragment in out, f"{src!r} 未包含 {expect_fragment!r}：{out!r}"


# ---------------------------------------------------------------- C. 纯幂次等价性
@pytest.mark.parametrize("literal,braced", [
    ("10⁸", "10^{8}"),
    ("10⁻⁷", "10^{-7}"),
    ("x²", "x^{2}"),
    ("x³", "x^{3}"),
    ("x⁻²", "x^{-2}"),
    ("2⁶", "2^{6}"),
    ("10⁹", "10^{9}"),
    ("x⁰", "x^{0}"),
])
def test_literal_superscript_equals_braced_without_units(literal, braced):
    """★ 不带单位时，字面上标与花括号必须逐字节等价。

    与 `test_super_unit_parity.py` 互补：那边测「幂次 + 单位」，
    这边测**纯幂次**（单位干扰不存在，能单独验证上标换算本身）。
    """
    assert L.linear_to_latex(literal) == L.linear_to_latex(braced), \
        f"{literal!r} vs {braced!r} 不等价"


# ---------------------------------------------------------------- D. 真分式反向边界
@pytest.mark.parametrize("src,why", [
    ("v=s/t", "纯量符号之间 —— 真分式"),
    ("p=F/A", "同上"),
    ("ρ=m/V", "同上"),
    ("E=V/m", "同上"),
    ("I=U/R", "同上"),
    ("n=m/M", "同上"),
    ("a/b", "字母除字母"),
    ("x/y", "字母除字母"),
])
def test_true_fractions_are_never_turned_into_dimension_strings(src, why):
    """★ 真分式保护的反向边界：这些**必须**是 `\\frac`，不许变成量纲串。

    审查 §三 对抗1 指出 `m s g h l r N T K C A V F H J L S` 这 18 个字母
    同时是变量符号。本组就是那个面的反向验证：它们是变量时必须是分式。
    """
    out = L.linear_to_latex(src)
    assert r"\frac" in out, f"{src!r}（{why}）没变成分式：{out!r}"
    assert r"\text{" not in out, f"{src!r}（{why}）被误判成量纲串：{out!r}"


# ---------------------------------------------------------------- E. 穷举兜底
def test_exhaustive_short_runs_never_illegal():
    """穷举长度 1~3 的全部上标字符组合 + 上下标混排，断言永不出现非法形态。

    这是「宁可多跑几秒，也不留一个非法形态」的兜底。
    """
    total = 0
    bad = []
    for length in (1, 2, 3):
        for combo in itertools.product(SUPERS, repeat=length):
            src = "10" + "".join(combo)
            out = L.linear_to_latex(src)
            total += 1
            hit = _illegal(out)
            if hit:
                bad.append((src, out, hit))
    for length in (1, 2):
        for combo in itertools.product(SUBS, repeat=length):
            src = "v" + "".join(combo)
            out = L.linear_to_latex(src)
            total += 1
            hit = _illegal(out)
            if hit:
                bad.append((src, out, hit))
    assert not bad, f"{len(bad)}/{total} 条出现非法形态，前几条：{bad[:5]}"
    # 穷举规模：14¹+14²+14³（上标）+ 12¹+12²（下标）= 2744 + 156 + 210 = 3110。
    # 阈值按**实数**定，不虚报 —— 写死 5000 只会让这条测试假红。
    assert total == 3110, f"穷举规模变了（{total}），检查字符集是否被改过"
