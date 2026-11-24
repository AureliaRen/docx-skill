# -*- coding: utf-8 -*-
r"""``latex_to_omml_safe()``：手写调用方一直缺的那个出口，现在有名字了。

使用汇报缺陷 3（lead 已核实）：``math_omml.latex_to_omml_lxml`` 是**纯转换** ——
``latex2mathml`` 把 ``\dot{I}`` 翻成 ``m:limUpp``（上限），Word 里渲染成
「大字 + 飘着的点」；正确形态是 ``m:acc``（U+0307 组合点）。修复函数
``fix_phasor_dots`` 一直都在，但原 docstring 只说「返回带 xmlns:m 的元素」，
**零警示**，于是每个手写调用方都要自己踩一次这个坑。

本文件钉四件事：

1. ``latex_to_omml_safe`` 真的修了相量点（``limUpp`` 归零、``acc`` 出现），
   并把改了几处写进 ``stats["phasor_fixes"]``；
2. 它与裸调 ``latex_to_omml_lxml`` 在**同一输入**上结果不同 —— 证明包装不是空壳；
3. 返回值形状与裸调一致（``display`` 决定 ``m:oMath`` / ``m:oMathPara``）；
4. docstring 的警示还在（文案也是契约：删掉就等于退回「零提醒」）。

变异验证（证据见 ``docs/reports/P23-A-FORMULA-ENTRY.md``）：把 ``latex_to_omml_safe``
里的 ``fix_phasor_dots`` 调用注释掉，第 1、2 组断言立刻变红。

OMML 可用性判定统一走 ``tests/_omml.py``（全仓唯一实现）——本机有 Office，
CI 腿没有，缺那一环时这些用例 skip 而不是红。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from lxml import etree  # noqa: E402

import fix_phasor_dot  # noqa: E402
import math_omml  # noqa: E402
from _omml import needs_omml  # noqa: E402

# ``\dot{I}`` 的正确形态：``m:acc``（accent）+ U+0307；错误形态 ``m:limUpp``。
PHASOR = r"\dot{I}"


def _count(el, local: str) -> int:
    """数一个本地名在整棵树里出现几次（不数字符串，避免 start/end 标签算两次）。"""
    return sum(1 for e in el.iter()
               if isinstance(e.tag, str) and e.tag.rsplit("}", 1)[-1] == local)


def _root_name(el) -> str:
    return el.tag.rsplit("}", 1)[-1]


# ---------------------------------------------------------------- 4. 文案契约
def test_the_docstring_says_the_function_does_not_fix_phasor_dots():
    """警示本身是交付物的一部分：它被删掉，手写调用方就回到零提醒。

    这条不挂 ``needs_omml`` —— 它读的是字符串，不需要 Office。
    """
    doc = math_omml.latex_to_omml_lxml.__doc__ or ""
    assert "不修相量点" in doc, "警告第一句必须点明「不修相量点」"
    assert "latex_to_omml_safe" in doc, "必须指向推荐的包装器"
    assert "fix_phasor_dots" in doc, "必须点名手写调用方要补的那一刀"


# ---------------------------------------------------------------- 1. 真的修了
@needs_omml
def test_safe_wrapper_replaces_the_lim_upper_form_with_an_accent():
    r"""``\dot{I}``：safe 出来的是 accent，不是 limUpp。"""
    el = math_omml.latex_to_omml_safe(PHASOR)
    assert _count(el, "limUpp") == 0
    assert _count(el, "acc") >= 1


@needs_omml
def test_safe_wrapper_reports_how_many_dots_it_fixed():
    """``stats`` 是可选出口：传了就填 ``{"phasor_fixes": n}``，不传不许炸。"""
    stats: dict = {}
    math_omml.latex_to_omml_safe(PHASOR, stats=stats)
    assert stats["phasor_fixes"] >= 1

    # 没有 ``\dot`` 的公式：键仍然存在、值是 0 —— 调用方可以无脑累加。
    plain: dict = {}
    math_omml.latex_to_omml_safe("E = mc^2", stats=plain)
    assert plain["phasor_fixes"] == 0

    # 不传 stats 也不抛（默认参数路径）。
    assert _root_name(math_omml.latex_to_omml_safe(PHASOR)) == "oMath"


# ---------------------------------------------------------------- 2. 包装不是空壳
@needs_omml
def test_bare_call_and_safe_wrapper_disagree_on_the_same_input():
    """同一输入，裸调留 limUpp、safe 不留 —— 这是「包装真的生效」的判据。

    如果哪天有人把 safe 改成 ``return latex_to_omml_lxml(...)``（或把修复
    注释掉），这条立刻红：两个计数会相等。
    """
    bare = math_omml.latex_to_omml_lxml(PHASOR)
    safe = math_omml.latex_to_omml_safe(PHASOR)

    assert _count(bare, "limUpp") > 0, "裸调本该留下 limUpp（latex2mathml 的现状）"
    assert _count(safe, "limUpp") == 0
    assert _count(bare, "limUpp") != _count(safe, "limUpp")
    # 反向：acc 只在 safe 里出现，裸调一处都没有
    assert _count(bare, "acc") == 0
    assert _count(safe, "acc") >= 1


# ---------------------------------------------------------------- 3. 形状一致
@needs_omml
@pytest.mark.parametrize("display,expected", [(False, "oMath"), (True, "oMathPara")])
def test_safe_returns_the_same_shape_as_the_bare_call(display, expected):
    """``display`` 语义必须和裸调一致 —— 否则把它当 drop-in 替换会改版面。"""
    assert _root_name(math_omml.latex_to_omml_safe("E = mc^2", display=display)) == expected
    assert _root_name(math_omml.latex_to_omml_lxml("E = mc^2", display=display)) == expected


@needs_omml
def test_safe_keeps_the_m_namespace_so_the_element_can_be_appended():
    """返回值仍然是「自带 xmlns:m 的元素」，能直接 append 进 ``w:p``。"""
    el = math_omml.latex_to_omml_safe(PHASOR)
    assert el.nsmap.get("m") == math_omml.M_NS
    xml = etree.tostring(el, encoding="unicode")
    assert xml.startswith("<m:oMath")


# ---------------------------------------------------------------- 异常语义
@needs_omml
def test_safe_still_raises_for_latex_it_cannot_convert():
    r"""转不动的 LaTeX 照抛 ``RuntimeError`` —— safe 不吞转换失败。

    ⚠️ 坏样本必须挑**真失败**的：``r"\frac{a"`` 看着像坏括号，但
    ``latex2mathml`` 容忍它（实测产出一个正常 ``m:oMath``，不会抛），拿它当
    坏样本会得到一条永远绿、什么也没证明的测试。``r"\frac"`` 才会抛。
    """
    with pytest.raises(RuntimeError):
        math_omml.latex_to_omml_safe(r"\frac")


@needs_omml
def test_a_phasor_fix_failure_is_recorded_not_raised(monkeypatch):
    """修相量点出事 → 记 ``phasor_errors``、返回**未修但可用**的元素。

    与 ``docx_mathify._omml_with_phasor`` 同一套语义：公式本身是好的，相量点
    没修只是观感问题，不该让整篇转换失败。
    """
    def boom(_el):
        raise ValueError("boom")

    monkeypatch.setattr(fix_phasor_dot, "fix_phasor_dots", boom)
    stats: dict = {}
    el = math_omml.latex_to_omml_safe(PHASOR, stats=stats)

    assert _root_name(el) == "oMath", "公式仍然要返回"
    assert _count(el, "limUpp") > 0, "没修成，limUpp 原样留着"
    assert "phasor_fixes" not in stats
    assert len(stats["phasor_errors"]) == 1
    assert "ValueError" in stats["phasor_errors"][0]["error"]

