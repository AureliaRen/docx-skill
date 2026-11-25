# -*- coding: utf-8 -*-
r"""OMML 直立补全：按语义拆 `m:r` 并给该直立的块补 `m:sty="p"`。

## 为什么这些断言长这样

Word 的数学区里字母默认**斜体**（变量），函数名 / 数字 / 运算符才该直立。
latex2mathml 只在少数场合自带 `m:sty="p"`（`\sin` / `\log` /
`\mathrm{…}`），而它给出的一个 `m:r` 经常同时装着两类字符
（`\frac{dy}{dx}` 的分子是单个 run "dy"、分母是单个 run "dx"）。所以
**只给 run 补属性是错的** —— 那会把 dx 的 x 一起变直立。必须先按语义分块再拆 run，
本文件钉的就是「分块判据」与「拆完的结构仍然合法」。

## 三条反直觉但正确的期望

1. `e^{-x}` 里的 `e` **不直立**（sty 是 None）。它是自然常数但不该按数字处理 ——
   数学排版里 e / i / π 都随变量走斜体，直立反而错。别把它改成 p。
2. `\max\{x,y\}` 改之前是**一个** run `'max{x,y}'`；改之后第一段必须是
   `('max', 'p')` —— 函数名从变量流里被切出来，这是本单的核心产出。
3. `\sin x` 的 run 数必须仍是 2。latex2mathml 已经给它 `m:sty="p"` 了，
   再拆一遍会重复处理上游样式，把 'sin' 切成 's'/'i'/'n'。

OMML 可用性判定统一走 `tests/_omml.py`（全仓唯一实现）：本机有 Office，CI 腿
没有，缺那一环时这些用例 skip 而不是红。

变异验证见 `docs/reports/P14-AB-FORMULA.md`：M1（去掉「已有 m:sty 就跳过」）
让 `\sin` 那条红；M2（拆 run 不 deepcopy w:rPr）让字体那条红。
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
W = math_omml.W_NS


def _m(tag: str) -> str:
    return "{%s}%s" % (M, tag)


def _w(tag: str) -> str:
    return "{%s}%s" % (W, tag)


def runs(el) -> list[tuple[str, str | None]]:
    """`m:r` 列表（文档序）→ `[(text, sty)]`，sty 取 `m:rPr/m:sty@val`。"""
    out = []
    for run in el.iter(_m("r")):
        text = "".join(t.text or "" for t in run.iter(_m("t")))
        sty = None
        rpr = run.find(_m("rPr"))
        if rpr is not None:
            sty_el = rpr.find(_m("sty"))
            if sty_el is not None:
                sty = sty_el.get(_m("val"))
        out.append((text, sty))
    return out


# ------------------------------------------------------------------ 1. e 不直立
@needs_omml
def test_euler_e_stays_italic_and_the_minus_turns_upright():
    r"""`e^{-x}`：e 是变量（None），减号是运算符（p），x 是变量（None）。

    **e 不该是 p。** 数学排版里 e / i / π 随变量走斜体；把它们按「数字/常数」直立
    是常见误改。这条专门防止有人为了「看着整齐」把单字母也塞进白名单。
    """
    latex = r"e^{-x}"
    el = math_omml.latex_to_omml_lxml(latex)
    got = runs(el)

    assert got[0] == ("e", None), f"e 必须是变量（斜体），实际 {got!r}"
    assert ("\u2212", "p") in got, f"U+2212 减号必须直立，实际 {got!r}"
    assert got[-1] == ("x", None), f"指数上的 x 必须是变量，实际 {got!r}"
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 2. 微分算子
@needs_omml
def test_differential_operator_d_is_upright_in_both_numerator_and_denominator():
    r"""`\frac{dy}{dx}`：拆成 4 段 d(p) y(None) d(p) x(None)。

    改之前是两个 run `('dy', None)` `('dx', None)`。整 run 补 p 会让 y/x 也直立；
    这条钉住「d 单独切出来」，同时钉住分子分母**各自**处理（不是只处理第一个）。
    """
    latex = r"\frac{dy}{dx}"
    el = math_omml.latex_to_omml_lxml(latex)
    got = runs(el)

    assert got == [("d", "p"), ("y", None), ("d", "p"), ("x", None)], \
        f"\\frac{{dy}}{{dx}} 的 run 序列不对：{got!r}"
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 3. 积分式
@needs_omml
def test_integral_digits_are_upright_while_the_integrand_variable_is_not():
    r"""`\int_0^1 2x\,dx = 1`：数字位与等号 p，x 变量，微分 d 直立。

    这条同时覆盖三类判据（纯数字、非字母非数字、微分算子），也覆盖 NBSP
    （`\,` 经 keep_spacing 变成控制空格）：NBSP 段上游已带 p，必须整 run 跳过，
    不能被再拆一遍。
    """
    latex = r"\int_0^1 2x\,dx = 1"
    el = math_omml.latex_to_omml_lxml(latex)
    got = runs(el)

    assert ("0", "p") in got and ("1", "p") in got, f"上下限数字必须直立：{got!r}"
    assert ("2", "p") in got, f"系数 2 必须直立：{got!r}"
    assert ("d", "p") in got, f"微分算子 d 必须直立：{got!r}"
    assert ("x", None) in got, f"积分变量 x 不该直立：{got!r}"
    assert ("=", "p") in got, f"等号必须直立：{got!r}"
    assert ("\u00a0", "p") in got, f"NBSP 段应保持上游的 p：{got!r}"
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 4. 函数名
@needs_omml
def test_max_gets_split_out_of_the_single_run_it_used_to_share():
    r"""`\max\{x,y\}`：改之前是**一个** run `'max{x,y}'`，现在第一段是 `('max','p')`。

    这是本单的核心用例 —— latex2mathml 把 `\max` 当普通字母排，整个 `max{x,y}`
    落进一个 run，Word 里连函数名都是斜体。只有按字符扫描才能把它切出来。
    """
    latex = r"\max\{x,y\}"
    el = math_omml.latex_to_omml_lxml(latex)
    got = runs(el)

    assert got[0] == ("max", "p"), f"第一段必须是直立的 max，实际 {got!r}"
    assert len(got) > 1, f"max{{x,y}} 必须被拆开（改之前是单 run），实际 {got!r}"
    assert ("x", None) in got and ("y", None) in got, f"x/y 仍是变量：{got!r}"
    assert ("{", "p") in got and ("}", "p") in got, f"花括号必须直立：{got!r}"
    math_omml.assert_legal(el, latex)


@needs_omml
def test_mathrm_max_and_bare_max_end_up_identical():
    r"""`\mathrm{max}\{x,y\}` 与 `\max\{x,y\}` 归一成同一 run 序列。

    两种写法本该同义；改之前一个是单 run、另一个是 p + 大括号两块，Word 里观感
    完全不同。归一之后谁写都行，用户不必记「要直立就加 \\mathrm」。
    """
    bare = runs(math_omml.latex_to_omml_lxml(r"\max\{x,y\}"))
    mathrm = runs(math_omml.latex_to_omml_lxml(r"\mathrm{max}\{x,y\}"))
    assert bare == mathrm, f"两种写法应归一：\n  bare   {bare!r}\n  mathrm {mathrm!r}"


# ------------------------------------------------------------------ 5. 不重复处理
@needs_omml
def test_sin_already_upright_is_not_split_again():
    r"""`\sin x`：'sin' 是 p，且 run 数仍是 2（不重复处理上游样式）。

    latex2mathml 已经给 `\sin` 补了 `m:sty="p"`。若无条件按字符重切，'sin'
    会变成 s/i/n 三段 —— 样式没变但 run 结构被无谓打碎，字距也会变。
    """
    latex = r"\sin x"
    stats: dict = {}
    el = math_omml.latex_to_omml_lxml(latex, stats=stats)
    got = runs(el)

    assert got == [("sin", "p"), ("x", None)], f"\\sin x 不该被重切：{got!r}"
    # stats 是「有没有被碰过」的可观察出口：尊重上游意味着 split 计数为 0。
    # 只比文本序列抓不到「原样重建」——重建后的 text/sty 看起来一样，但 m:rPr 里
    # 除 sty 之外的东西（nor / scr / …）会在重建时丢掉。所以这条必须钉计数。
    assert stats["upright_split"] == 0, f"上游已有 sty 的 run 不该被重建：{stats!r}"
    assert stats["upright_runs"] == 0, f"上游已有 sty 的 run 不该再补 sty：{stats!r}"
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 6. nor 不拆
@needs_omml
def test_text_run_with_nor_is_left_alone():
    r"""`\text{if }x>0`：带 `m:nor` 的 run 整段不动，'if ' 仍是单 run。

    `\text{}` 产出的是正文字符串（可能是中文、空格、标点），`m:nor` 已经强制
    非数学样式。若按字符拆它，一句正文会被切成十几个 run，编辑体验与字距都毁。
    """
    latex = r"\text{if }x>0"
    el = math_omml.latex_to_omml_lxml(latex)

    nor_runs = [r for r in el.iter(_m("r"))
                if r.find(_m("rPr")) is not None
                and r.find(_m("rPr")).find(_m("nor")) is not None]
    assert len(nor_runs) == 1, f"\\text{{}} 的 nor run 应恰好 1 个，实际 {len(nor_runs)}"
    nor_text = "".join(t.text or "" for t in nor_runs[0].iter(_m("t")))
    assert nor_text == "if\u00a0", f"nor run 被拆坏了：{nor_text!r}"
    # 同一公式里的数学部分照常处理（x>0 的 > 与 0 要直立）
    assert (">", "p") in runs(el)
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 7. 结构合法
@needs_omml
@pytest.mark.parametrize("latex", [
    r"e^{-x}",
    r"\frac{dy}{dx}",
    r"\int_0^1 2x\,dx = 1",
    r"\max\{x,y\}",
    r"\mathrm{max}\{x,y\}",
    r"\sin x",
    r"\text{if }x>0",
    r"\log_2 n",
    r"\det A",
])
def test_every_split_run_keeps_the_ct_r_content_model(latex):
    r"""`m:r` 的子元素必须是 `m:rPr?` → `w:rPr?` → `m:t`。

    位置错了 Word 不报错，只会**静默**把样式或字体丢掉 —— 合法 XML + 观感错，
    是最难查的一类。所以逐 run 断「m:rPr（若有）是首子元素、m:t 是最后子元素」。
    """
    el = math_omml.latex_to_omml_lxml(latex)
    for run in el.iter(_m("r")):
        kids = [k.tag for k in run]
        assert kids, f"空 m:r：{latex!r}"
        assert kids[-1] == _m("t"), f"m:t 必须是最后子元素：{latex!r} → {kids}"
        if _m("rPr") in kids:
            assert kids[0] == _m("rPr"), f"m:rPr 必须是首子元素：{latex!r} → {kids}"
        assert kids.count(_m("t")) == 1, f"一个 m:r 只能有一个 m:t：{latex!r} → {kids}"
    # 拆 run 之后结构必须仍然过自检
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 8. 字体不丢
@needs_omml
@pytest.mark.parametrize("latex", [r"\max\{x,y\}", r"\frac{dy}{dx}", r"\int_0^1 2x\,dx = 1"])
def test_split_runs_keep_the_injected_font(latex):
    r"""拆出来的每段都带 `w:rPr/w:rFonts`，ascii 是 `math_font()`（默认 Times New Roman）。

    字体是在 `apply_math_font` 阶段注入的，而拆 run 发生在它**之后**。若重建 run
    时不 deepcopy 原 `w:rPr`，拆出来的段会回落 Cambria Math —— 一条公式里字体半截
    TNR 半截 Cambria，比整条都错还难发现（所以这条单独钉）。
    """
    font = math_omml.math_font()
    assert font, "本用例假定默认会注入字体（MATH_FONT 未设成空串）"
    el = math_omml.latex_to_omml_lxml(latex)
    for run in el.iter(_m("r")):
        rpr = run.find(_w("rPr"))
        assert rpr is not None, f"拆出来的 run 丢了 w:rPr：{latex!r}"
        fonts = rpr.find(_w("rFonts"))
        assert fonts is not None, f"w:rPr 里没有 w:rFonts：{latex!r}"
        assert fonts.get(_w("ascii")) == font, \
            f"字体不是 {font!r}：{fonts.get(_w('ascii'))!r}（{latex!r}）"


# ------------------------------------------------------------------ 9. 统计出口
@needs_omml
def test_stats_report_how_many_runs_were_split_and_how_many_styles_added():
    """`stats` 是可选出口：传了就填 `upright_split` / `upright_runs`，不传不许炸。"""
    stats: dict = {}
    math_omml.latex_to_omml_lxml(r"\max\{x,y\}", stats=stats)
    assert stats["upright_split"] >= 1
    assert stats["upright_runs"] >= 1

    # 没有可拆内容的公式：键存在、值是 0 —— 调用方可以无脑累加
    plain: dict = {}
    math_omml.latex_to_omml_lxml("x", stats=plain)
    assert plain["upright_split"] == 0
    assert plain["upright_runs"] == 0

    # 不传 stats 也不抛
    assert math_omml.latex_to_omml_lxml(r"\max\{x,y\}") is not None


# ------------------------------------------------------------------ 10. 分块判据
def test_classify_run_text_is_the_documented_block_grammar():
    """直接钉 `classify_run_text` 的分块结果 —— 纯判据单测，不依赖 Office。

    函数名后接小写字母时不认（`maxx` 是变量名），后接 `{` / `(` / 空格 / 大写 /
    结尾都算命中；微分 d 只在后接字母时单独直立。
    """
    c = math_omml.classify_run_text

    assert c("max{x,y}") == [("max", "p"), ("{", "p"), ("x", None), (",", "p"),
                             ("y", None), ("}", "p")]
    assert c("maxx") == [("m", None), ("a", None), ("x", None), ("x", None)], \
        "maxx 是变量名，白名单不该命中"
    assert c("dx") == [("d", "p"), ("x", None)]
    assert c("3.14") == [("3.14", "p")], "小数点连写算一个数字块"
    assert c("1,2") == [("1,2", "p")], "逗号连写算一个数字块"
    assert c("xy") == [("x", None), ("y", None)]
    assert c("\u00a0") == [("\u00a0", None)], "NBSP 不设 sty"
    assert c("sinh") == [("sinh", "p")], "长名优先，不能被切成 sin + h"

# ------------------------------------------------------------------ 11. 接线
@needs_omml
def test_docx_mathify_passes_stats_into_the_converter():
    r"""接线不是空壳：@@docx_mathify._omml_with_phasor(tex, stats)@@ 必须把同一个 dict
    透传给 @@latex_to_omml_lxml@@，于是直立计数与相量点计数落在**同一份** stats 里。

    为什么单独钉：@@stats@@ 默认 None，接线漏了也不会报错、转换照样正确 —— 只是
    @@--report@@ 里永远看不到「直立补全修了几段」。这类静默丢统计最容易在后续重构中
    回退，所以用可观察的证据钉住：一份空 dict 进去，出来后应有 upright / phasor 两类键。
    """
    import docx_mathify

    stats: dict = {}
    # 一条公式同时触发两件事：dy/dx 走直立补全，\dot{I} 走相量点修复。
    docx_mathify._omml_with_phasor(r"\dot{I}+\frac{dy}{dx}", stats)

    assert stats.get("upright_split", 0) >= 1, f"stats 没透传到转换器：{stats!r}"
    assert "phasor_fixes" in stats, f"相量点键也应在同一份 stats 里：{stats!r}"

# ------------------------------------------------------------------ 12. nor 整 run 不动（L5 缺口）
@needs_omml
@pytest.mark.parametrize("latex, expected_text", [
    (r"\text{if }x>0", "if\u00a0"),
    (r"\text{max 3.14}", "max\u00a03.14"),
    (r"\text{sin 2}x", "sin\u00a02"),
], ids=["words", "func-and-float", "func-and-digit"])
def test_nor_run_stays_one_run_even_when_its_text_looks_upright(latex, expected_text):
    r"""带 `m:nor` 的 run **一个都不拆**，哪怕它自己的文本看着该直立。

    `\text{max 3.14}` 里 `max` 在函数白名单里、`3.14` 是数字 —— 不看 nor 的话
    `classify_run_text` 会把它切成 `[('max', 'p'), (NBSP, None), ('3.14', 'p')]`，
    正文被拆成 3 个 run，`m:nor` 只剩第一段带着（其余两段变成数学样式，正文不再直立）。

    为什么用例文本不能只用 `\text{if }`：`if ` 拆出来全是 None（没数字、没函数名、
    没运算符），`all(kind is None)` 那一关本来就会跳过它 —— 那样测的是「碰巧没可拆
    内容」，把 nor 判断整条删掉也照样绿。必须用**自身含可直立块**的 nor run 才钉得住。
    """
    el = math_omml.latex_to_omml_lxml(latex)

    nor_runs = [r for r in el.iter(_m("r"))
                if r.find(_m("rPr")) is not None
                and r.find(_m("rPr")).find(_m("nor")) is not None]
    assert len(nor_runs) == 1, \
        f"nor run 应恰好 1 个（被拆就会多出无 nor 的兄弟 run），实际 {len(nor_runs)}：{latex!r}"

    nor_run = nor_runs[0]
    texts = [t.text or "" for t in nor_run.iter(_m("t"))]
    assert len(texts) == 1, f"nor run 只能有一个 m:t（拆过就分家）：{texts!r}（{latex!r}）"
    assert texts[0] == expected_text, f"nor run 的文本被改了：{texts[0]!r} != {expected_text!r}"
    math_omml.assert_legal(el, latex)


# ------------------------------------------------------------------ 13. nor 不计入 stats
@needs_omml
def test_nor_runs_are_excluded_from_the_upright_counters():
    r"""`stats` 只数**非 nor** 的 run：`\text{max }\frac{dy}{dx}` → split 恰好 2。

    第 12 条只钉形态（nor run 还是一个、文本完整），这条钉「处理有没有发生」——
    两个方向都要压住：混排公式里计数必须等于数学 run 的条数（dy、dx 各一条），
    纯正文公式必须 0。于是「拆了 nor run 但统计不记」这种改法也跑不掉：形态侧
    第 12 条红，计数侧本条红。

    数字是实测读数（不是估算）：`\text{max }` 的 nor run 不拆，dy / dx 各拆一条，
    各补一段 sty，所以 (upright_split, upright_runs) == (2, 2)。
    """
    mixed: dict = {}
    el = math_omml.latex_to_omml_lxml(r"\text{max }\frac{dy}{dx}", stats=mixed)
    assert (mixed["upright_split"], mixed["upright_runs"]) == (2, 2), \
        f"nor run 不该计入直立计数，只该数 dy/dx：{mixed!r}"
    got = runs(el)
    assert got == [("max\u00a0", None), ("d", "p"), ("y", None), ("d", "p"), ("x", None)], \
        f"混排公式的 run 序列不对：{got!r}"
    math_omml.assert_legal(el, r"\text{max }\frac{dy}{dx}")

    nor_only: dict = {}
    math_omml.latex_to_omml_lxml(r"\text{max 3.14}", stats=nor_only)
    assert (nor_only["upright_split"], nor_only["upright_runs"]) == (0, 0), \
        f"纯 \\text{{}} 公式不该产生任何直立拆分：{nor_only!r}"
