# -*- coding: utf-8 -*-
r"""cases 左对齐：只改 cases，绝不碰矩阵。

## 为什么需要这条区分

OMML 里 `\begin{cases}` 与 `\begin{pmatrix}` / bmatrix / vmatrix / matrix
**完全同构** —— 都是 `m:m` + `baseJc=center` + `plcHide=on` +
`mcs/mc/mcPr/{count,mcJc=center}`。XSLT 统一给 `mcJc=center`，于是 cases 的
「值 & 条件」两列被居中，和数学排版惯例（左对齐）不符。

判据只能来自 MathML 层：`\begin{cases}` 的每个 `mtd` 都带
`columnalign="left"`，矩阵环境的 `mtd` 不带任何 columnalign。本文件用**硬反证**
钉住这件事 —— 矩阵那几条断言的是 `center`，不是「不等于 left」这种弱口径：
把矩阵也改成 left 是本单最容易犯的错（无条件改所有 `m:m` 一行代码就能做到），
所以必须有测试能立刻抓住它。

## 为什么单列 cases 也要测

`\begin{cases} 1 \\ 0 \end{cases}` 不产出 `m:m`（XSLT 对单列环境不建矩阵）。
此时 MathML 侧的 mtable 列表与 OMML 侧的 m:m 列表长度不等 —— 处理函数必须
**整段放弃**而不是硬对位，否则后面每一个 `m:m` 都会错位一格。这条钉的就是
「放弃」这个分支不炸、也不误改。

变异验证见 `docs/reports/P14-AB-FORMULA.md`：M3（无条件把所有 m:m 改 left）
让矩阵四条红；M4（长度不等也尽量对位）让嵌套那条红。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import math_omml  # noqa: E402
from _omml import needs_omml  # noqa: E402

M = math_omml.M_NS


def _m(tag: str) -> str:
    return "{%s}%s" % (M, tag)


def mc_jc(el) -> list[str]:
    """所有 `m:mcJc@val`（文档序，扁平）。判定「全 left / 全 center」用这个。"""
    return [jc.get(_m("val")) for jc in el.iter(_m("mcJc"))]


def per_matrix(el) -> list[list[str]]:
    """按 `m:m` 分组：每个矩阵自己的 mcJc 列表（文档序，外层先）。

    嵌套时扁平列表会把内层与外层的值混在一起（nested 得到
    `['left', 'center']`，读不出谁是谁），所以判断「外层 left、内层 center」
    必须用分组口径。
    """
    return [[jc.get(_m("val")) for jc in m.iter(_m("mcJc"))]
            for m in el.iter(_m("m"))]


CASES_TWO_COL = r"f(x)=\begin{cases} 0 & x<0 \\ 1 & x\geq 0\end{cases}"


# ------------------------------------------------------------------ 1. cases
@needs_omml
def test_cases_columns_become_left_aligned():
    """`\begin{cases}` 的两列改左对齐 —— 本单的主目标。"""
    el = math_omml.latex_to_omml_lxml(CASES_TWO_COL)
    assert mc_jc(el), "cases 没产出 m:m（转换退化？）"
    assert set(mc_jc(el)) == {"left"}, f"cases 的 mcJc 应全是 left，实际 {mc_jc(el)!r}"


# ------------------------------------------------------------------ 2-5. 矩阵反证
@needs_omml
def test_pmatrix_stays_centred():
    """紧反证：`pmatrix` 必须仍是 center。改成 left 就是本单最典型的误伤。"""
    el = math_omml.latex_to_omml_lxml(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}")
    assert set(mc_jc(el)) == {"center"}, f"pmatrix 的 mcJc 应全是 center，实际 {mc_jc(el)!r}"


@needs_omml
def test_bmatrix_stays_centred():
    el = math_omml.latex_to_omml_lxml(r"\begin{bmatrix} a & b \\ c & d \end{bmatrix}")
    assert set(mc_jc(el)) == {"center"}, f"bmatrix 的 mcJc 应全是 center，实际 {mc_jc(el)!r}"


@needs_omml
def test_vmatrix_stays_centred():
    el = math_omml.latex_to_omml_lxml(r"\begin{vmatrix} a & b \\ c & d \end{vmatrix}")
    assert set(mc_jc(el)) == {"center"}, f"vmatrix 的 mcJc 应全是 center，实际 {mc_jc(el)!r}"


@needs_omml
def test_bare_matrix_stays_centred():
    el = math_omml.latex_to_omml_lxml(r"\begin{matrix} a & b \\ c & d \end{matrix}")
    assert set(mc_jc(el)) == {"center"}, f"matrix 的 mcJc 应全是 center，实际 {mc_jc(el)!r}"


# ------------------------------------------------------------------ 6-7. 同式混排
@needs_omml
def test_matrix_then_cases_each_keeps_its_own_alignment():
    """同一条 LaTeX 里先 matrix 后 cases → 第一个 center，第二个 left。

    这条钉「按文档序对位」这件事本身：只看「有几个 m:m、有几个 cases」而不看顺序，
    在这条上会得到第一个 left、第二个 center（正好反了）。
    """
    latex = (r"A=\begin{pmatrix} 1&0\\0&1\end{pmatrix},\ "
             r"f(x)=\begin{cases}1&x>0\\0&x\leq0\end{cases}")
    el = math_omml.latex_to_omml_lxml(latex)
    groups = per_matrix(el)

    assert len(groups) == 2, f"应有两个 m:m，实际 {len(groups)}：{groups!r}"
    assert set(groups[0]) == {"center"}, f"第一个是矩阵，应 center：{groups!r}"
    assert set(groups[1]) == {"left"}, f"第二个是 cases，应 left：{groups!r}"


@needs_omml
def test_cases_then_matrix_each_keeps_its_own_alignment():
    """顺序反过来（先 cases 后 matrix）→ 各按各的。

    上一条 + 这条一起，把「盲目按位改」这条路堵死：无论改成「全改第一个」还是
    「全改最后一个」，都必有一条红。
    """
    latex = (r"f(x)=\begin{cases}1&x>0\\0&x\leq0\end{cases},\ "
             r"A=\begin{pmatrix} 1&0\\0&1\end{pmatrix}")
    el = math_omml.latex_to_omml_lxml(latex)
    groups = per_matrix(el)

    assert len(groups) == 2, f"应有两个 m:m，实际 {len(groups)}：{groups!r}"
    assert set(groups[0]) == {"left"}, f"第一个是 cases，应 left：{groups!r}"
    assert set(groups[1]) == {"center"}, f"第二个是矩阵，应 center：{groups!r}"


# ------------------------------------------------------------------ 8. 嵌套
@needs_omml
def test_nested_matrix_inside_cases_keeps_outer_left_and_inner_centre():
    """cases 里套 pmatrix → 外层 left、内层 center（最容易写错的一条）。

    两层的 m:m 都「看起来是矩阵」，只有 MathML 侧的 columnalign 能区分。按扁平
    顺序盲改会同时改掉内层，而这恰好是排版上最难一眼看出的错误。
    """
    latex = (r"\begin{cases} \begin{pmatrix} 1&0\\0&1\end{pmatrix} & x>0 \\ "
             r"0 & x\leq 0 \end{cases}")
    el = math_omml.latex_to_omml_lxml(latex)
    groups = per_matrix(el)

    assert len(groups) == 2, f"嵌套应产出 2 个 m:m（外层 cases + 内层矩阵）：{groups!r}"
    assert "left" in groups[0], f"外层 cases 应有 left：{groups!r}"
    assert "center" in groups[1], f"内层矩阵必须仍是 center：{groups!r}"


# ------------------------------------------------------------------ 9. 单列 cases
@needs_omml
def test_one_column_cases_produces_no_matrix_and_does_not_raise():
    r"""`\begin{cases} 1 \\ 0 \end{cases}`：不产 m:m，mcJc 列表为空，也不许抛。

    这条钉的是「长度不等就放弃」分支：MathML 侧有 mtable、OMML 侧没有 m:m。
    若实现改成「尽量对位」，这里会去改一个不存在的 m:m（IndexError）或把别的
    公式的 m:m 改掉。所以断言分三层：不抛、无 m:m、mcJc 为空。
    """
    el = math_omml.latex_to_omml_lxml(r"\begin{cases} 1 \\ 0 \end{cases}")
    assert [m for m in el.iter(_m("m"))] == [], "单列 cases 不该产出 m:m"
    assert mc_jc(el) == [], f"没有 m:m 就不该有 mcJc，实际 {mc_jc(el)!r}"


# ------------------------------------------------------------------ 10. 伪 cases
@needs_omml
def test_braced_matrix_is_still_a_matrix_not_cases():
    r"""`\left\{\begin{matrix} a & b \\ c & d \end{matrix}\right.`：看起来像 cases，实质是矩阵。

    手写花括号 + matrix 的写法在中文讲义里很常见（作者只想加个大括号）。MathML
    层它是 matrix 环境、mtd 不带 columnalign，所以必须保持 center —— 判据看环境
    名而不是「有没有花括号」。
    """
    el = math_omml.latex_to_omml_lxml(
        r"\left\{\begin{matrix} a & b \\ c & d \end{matrix}\right.")
    assert set(mc_jc(el)) == {"center"}, f"带 \\left\\{{ 的 matrix 应仍 center，实际 {mc_jc(el)!r}"


# ------------------------------------------------------------------ 11. 统计出口
@needs_omml
def test_stats_report_alignment_counts():
    """`stats` 出口：改了几个 cases；矩阵不计入。"""
    ok: dict = {}
    math_omml.latex_to_omml_lxml(CASES_TWO_COL, stats=ok)
    assert ok["cases_aligned"] == 1, f"应改 1 个 cases：{ok!r}"

    plain: dict = {}
    math_omml.latex_to_omml_lxml(r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", stats=plain)
    assert plain["cases_aligned"] == 0, f"矩阵不该被改：{plain!r}"


# ------------------------------------------------------------------ 12. 放弃分支
def test_mismatched_table_and_matrix_counts_abandon_instead_of_shifting():
    r"""长度不等 → **整段放弃**，只在 stats 里留原因（不依赖 Office，所以不挂 needs_omml）。

    为什么这条要用合成树而不是找一条能触发它的 LaTeX：`align_cases_left` 会先把
    ncols < 2 的 mtable 滤掉（它们本来就不产 `m:m`），所以自然输入下两个列表
    几乎总是等长。真出现不等，只可能是 XSLT 与判据分叉了 —— 那正是最该「宁可不改」
    的时刻。手搭一对 2 个 mtable / 1 个 m:m 的树，直接钉这个守卫：断言**一处都没改**
    且原因写进 stats，这样任何「尽量对位」的改动都会立刻红。
    """
    from lxml import etree

    MM = math_omml.MATHML_NS
    nsmap = {"m": MM}

    # MathML 侧：两个 2 列 mtable（一个 cases 样、一个矩阵样）
    def table(align):
        attrs = ' columnalign="%s"' % align if align else ""
        return ('<mtable><mtr><mtd%s>x</mtd><mtd>y</mtd></mtr></mtable>' % attrs)

    mml = etree.fromstring(
        ('<math xmlns="%s">%s%s</math>' % (MM, table("left"), table(""))).encode("utf-8"))

    # OMML 侧：只有一个 m:m，mcJc=center
    omml = etree.fromstring(
        ('<m:oMath xmlns:m="%s"><m:m><m:mPr><m:mcs><m:mc><m:mcPr>'
         '<m:count m:val="2"/><m:mcJc m:val="center"/>'
         '</m:mcPr></m:mc></m:mcs></m:mPr></m:m></m:oMath>' % M).encode("utf-8"))

    stats: dict = {}
    changed = math_omml.align_cases_left(mml, omml, stats)

    assert changed == 0, "长度不等时必须放弃，一个都不改"
    assert mc_jc(omml) == ["center"], f"放弃时原有样式不许动：{mc_jc(omml)!r}"
    reasons = stats.get("cases_align_skipped") or []
    assert reasons and "mtable 2 vs m:m 1" in reasons[0], f"放弃原因不对：{stats!r}"


# ------------------------------------------------------------------ 13. 结构复检
@needs_omml
@pytest.mark.parametrize("latex", [
    CASES_TWO_COL,
    r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}",
    r"\begin{cases} \begin{pmatrix} 1&0\\0&1\end{pmatrix} & x>0 \\ 0 & x\leq 0 \end{cases}",
])
def test_alignment_keeps_the_omml_legal(latex):
    """改 `mcJc` 之后仍须过 `assert_legal`，且 `mcPr` 子元素顺序仍是 count → mcJc。

    缺 `mcs` 要补建，补建位置错了 CT_MPr 就不合法（Word 只说「可能已损坏」）。
    """
    el = math_omml.latex_to_omml_lxml(latex)
    math_omml.assert_legal(el, latex)
    for mcpr in el.iter(_m("mcPr")):
        names = [c.tag.rsplit("}", 1)[-1] for c in mcpr]
        assert names == ["count", "mcJc"], f"mcPr 子元素顺序不对：{names}"
