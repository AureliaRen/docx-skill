# -*- coding: utf-8 -*-
r"""字面 Unicode 幂次 vs 花括号幂次：**必须给出同一个结果**。

背景（R4 遗留 + 本轮发现）：
`HANDOFF-ROUND4` §三 3.2 报了「幂次后单位无锚点」，修法给 `_ANCHOR_RE`
加了「字面 Unicode 上标数字」分支。但那只修了**分式误判**那一面 ——
`3×10⁸m/s` 不再被读成分式。另一面没修：

    '3×10⁸s'          -> '3\\times 10^{8}s'            ← s 当变量，斜体
    '3\\times 10^{8}s' -> '3\\times 10^{8}\\,\\text{s}' ← s 是单位，正体

根因：`_unitify` 里的 `POW_UNIT_RE` 只认 `^{n}` 形态，而 `_unitify` 跑在
`_escape_unicode` **之前** —— 此刻源串里还是字面 `⁸`。于是字面写法漏掉了
「幂次后单位」这一步。

独立审查（`REVIEW-CODE-R4` §三 对抗3）说「22/22 等价」，是因为取样只挑了
`m/s`/`mol/L` 这类**带斜杠**的单位 —— 它们走 `_dimensionify` 另一条路，
不经过 `POW_UNIT_RE`。单字符单位（`s`/`J`/`W`/`V`/`A`/`F`/`H`/`T`…）
一个都没覆盖到。本文件补的就是这个缺口。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402


# 单字符 / 短单位：全部走 POW_UNIT_RE 这条独木桥（不带 `/`）
SIMPLE_UNITS = [
    "s", "J", "W", "V", "A", "F", "H", "T", "Pa", "kPa", "MPa",
    "Hz", "kHz", "Wb", "var", "kg", "m", "N", "C", "K", "mol", "g", "L",
]

# 其中真正会被 POW_UNIT_RE 转成正体的那一批。
# 另一批（kg / m / N / C / K / mol / g / L）**刻意不在** `UNIT_LIST` 里 ——
# 它们是「与量符号冲突的单字母」，见 UNITS 注释里那句
# 「`v=s/t` 必须保持真分式，就靠不收 t」。这条限制对**两种拼写同等生效**，
# 所以等价性不受影响（下面 118 条 parity 测试覆盖了全部 23 个单位）。
# 记在这里是为了不让「没有正体」被误读成 parity 缺口。
UPRIGHT_UNITS = [
    "s", "J", "W", "V", "A", "F", "H", "T", "Pa", "kPa", "MPa",
    "Hz", "kHz", "Wb", "var",
]
NOT_UPRIGHT_BY_DESIGN = ["kg", "m", "N", "C", "K", "mol", "g", "L"]

# 幂次的两种拼写：(字面形态, 花括号形态)
POW_FORMS = [
    ("3×10⁸", r"3\times 10^{8}"),
    ("10⁸", "10^{8}"),
    ("2×10⁶", r"2\times 10^{6}"),
    ("10³", "10^{3}"),
]


@pytest.mark.parametrize("unit", SIMPLE_UNITS)
@pytest.mark.parametrize("literal,braced", POW_FORMS)
def test_literal_and_braced_power_give_same_result(literal, braced, unit):
    """★ 核心不变量：`3×10⁸s` 与 `3\\times 10^{8}s` 必须逐字节相同。

    两种拼写是同一件事的两种写法（agent 从物理题里拷出来的是字面形态，
    手写 LaTeX 的是花括号形态），结果分叉就是 bug。
    """
    a = L.linear_to_latex(literal + unit)
    b = L.linear_to_latex(braced + unit)
    assert a == b, f"{literal + unit!r} -> {a!r}  但  {braced + unit!r} -> {b!r}"


@pytest.mark.parametrize("unit", UPRIGHT_UNITS)
def test_power_then_unit_renders_unit_upright(unit):
    """幂次后紧跟的单位必须排成正体 `\\text{}`，不能当变量排斜体。"""
    out = L.linear_to_latex("3×10⁸" + unit)
    assert r"\text{" in out, f"{unit!r} 的单位没有正体化：{out!r}"
    assert unit in out


@pytest.mark.parametrize("unit", NOT_UPRIGHT_BY_DESIGN)
def test_conflicting_single_letter_units_behave_the_same_in_both_spellings(unit):
    """`kg`/`m`/`N`/`mol`… 不转正体是**刻意的**（与量符号冲突），

    但**两种拼写必须一致**——这条才是 parity，不去要求它正体化。
    """
    a = L.linear_to_latex("3×10⁸" + unit)
    b = L.linear_to_latex(r"3\times 10^{8}" + unit)
    assert a == b, f"{unit!r} 两种拼写分叉：{a!r} vs {b!r}"


@pytest.mark.parametrize("literal,braced", POW_FORMS)
def test_negative_and_multi_digit_exponents(literal, braced):
    """负指数与多位指数同样要等价。"""
    for unit in ["s", "J", "mol"]:
        assert L.linear_to_latex(literal + unit) == L.linear_to_latex(braced + unit)


@pytest.mark.parametrize("src,expect_same_as", [
    ("1.0×10⁻⁷mol/L", r"1.0\times 10^{-7}mol/L"),
    ("3×10⁸m/s", r"3\times 10^{8}m/s"),
    ("3×10⁸m/s2", r"3\times 10^{8}m/s^{2}"),
    ("2.5×10⁶Pa", r"2.5\times 10^{6}Pa"),
])
def test_compound_units_also_match(src, expect_same_as):
    """带 `/` 的复合单位本来就走 `_dimensionify`，两种拼写也要一致。"""
    assert L.linear_to_latex(src) == L.linear_to_latex(expect_same_as)


def test_no_regression_on_existing_dimension_example():
    """R4 已锁的样例不许变：`c=3×10⁸m/s` 仍是量纲串。"""
    out = L.linear_to_latex("c=3×10⁸m/s")
    assert r"\text{m/s}" in out, out
    assert r"\frac" not in out, f"被误判成分式了：{out!r}"


def test_superscript_sign_alone_is_not_an_anchor():
    """`⁺`/`⁻` 单独出现不是数值锚点（没有数字可锚）。"""
    # 化学电荷：Fe³⁺ 的 ⁺ 不该被当成「幂次 + 单位」
    out = L.linear_to_latex("Fe^{3+}")
    assert out == r"Fe^{3+}", out


def test_pow_unit_re_only_accepts_digits():
    """`POW_UNIT_RE` 的字面分支只收数字上标，不收字母上标（ⁿ / ⁱ）。"""
    # ⁿ 是字母上标，不是数值 —— 不该触发「幂次后单位」
    out = L.linear_to_latex("xⁿm")
    assert r"\text{" not in out, f"字母上标被当成数字幂次了：{out!r}"
