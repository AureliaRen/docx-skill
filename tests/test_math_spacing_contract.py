# -*- coding: utf-8 -*-
r"""显式间距契约（成品侧）：7 条间距命令必须活下来，落成 NBSP（P29-C）。

> 纯函数那一半（`keep_spacing()` 改写表本身，14 条 + 4 条反向）已拆到
> `tests/test_keep_spacing_contract.py`，**不带门控**，因为它在任何机器上都能
> 真跑。本文件剩下的全部要过 `latex2mathml` + `MML2OMML.XSL`，所以整份
> `needs_omml`。

## 为什么需要这条契约

XSLT 会丢弃间距（连普通空格也丢）。实测（`work/p29c/probe_spacing_matrix.py`）：
把改写关掉（`raw_spacing=True`）之后，`a\quad b` 转出来是 `m:t = ['a', 'b']`
—— 间距**静默消失**，整条链路没有任何一处报错，交付物里两个符号直接贴在一起。

所以「哪些间距命令能活下来」完全由 `scripts/math_omml.py` 里 `keep_spacing()`
那一条正则改写表决定：把它改窄（少列一条命令）不会有别的测试变红。改写表本身的
自洽性由 `test_keep_spacing_contract.py` 钉；**本文件钉的是它在成品里的效果** ——
表里列的命令，转出来的 OMML `m:t` 里必须真的出现 NBSP（U+00A0）。

## 判据为什么是「产物里出现 NBSP」

`\ `（控制空格）经 latex2mathml 落成 NBSP；`keep_spacing` 把其余间距命令统一
改写成 `\ `，于是「活下来」的可观测形态就是某个 `m:t` 的文本等于 U+00A0。
普通空格会被 XSLT 丢掉，所以 NBSP 出现在成品文本里 = 这条命令确实被改写、并且
在 XSLT 之后仍然存在。

## 两条与直觉相反、但按实测钉住的现状

1. `\!`（负细间距）同样被改写成控制空格 —— 负间距语义**丢了**，产物里是一个
   正的 NBSP。这是 `keep_spacing` 的现状（正则表不分正负），照实钉住：
   改成正解会让 `test_negative_thin_space_also_becomes_nbsp` 变红，届时应连同
   CHANGELOG 一起改，而不是让它悄悄漂移。
2. `\qquad` 与 `\quad` 产物**完全相同**（各一个 NBSP）：正则交替里 `qquad`
   先匹配、同样改写成单个 `\ `。LaTeX 语义上 `\qquad` 应是两个 `\quad`，
   现状不是。同样照实钉住。

实测矩阵见 `docs/reports/P29-C-TESTS.md` 第 2 节。
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

#: 本文件断言的是成品 OMML 文本，没有 XSLT 就只能拿到降级纯文字。
pytestmark = needs_omml

#: 间距命令全集 —— 与 keep_spacing() 正则里的交替完全一致（清单本体在
#: test_keep_spacing_contract.py，这里按同一份语义再列一遍，避免跨文件 import）。
SPACING_COMMANDS = (
    r"\quad",
    r"\qquad",
    r"\,",
    r"\;",
    r"\:",
    r"\!",
    r"\ ",      # 控制空格：本身就是要改写成的目标形态
)

NBSP = "\u00a0"


def _q(tag: str) -> str:
    return "{%s}%s" % (M, tag)


def _texts(el) -> list:
    return [t.text or "" for t in el.iter(_q("t"))]


def _convert(latex: str):
    return math_omml.latex_to_omml_safe(latex)


def _spaced(cmd: str) -> str:
    """把命令插进两个标识符之间：拿空格就得靠命令本身，不能靠源文本里的空格。"""
    return "a%sb" % cmd


# ── (1) 成品契约：每条间距命令 → 产物里有 NBSP ──────────────────────────


@pytest.mark.parametrize("cmd", SPACING_COMMANDS)
def test_spacing_command_survives_as_nbsp(cmd):
    r"""`a<cmd>b` 转出来，两个标识符之间必须有一个 NBSP。

    这是本文件的主判据。注意「两个标识符都还在」也要一起验证 —— 否则
    「把 b 吃掉」这种回归会让 NBSP 判据假绿。
    """
    latex = _spaced(cmd)
    el = _convert(latex)
    texts = _texts(el)
    assert "a" in texts, "%s 丢了左操作数：%r" % (latex, texts)
    assert "b" in texts, "%s 丢了右操作数：%r" % (latex, texts)
    assert any(NBSP in t for t in texts), (
        "%s 的间距没活下来：m:t=%r（期望含 U+00A0）" % (latex, texts))


@pytest.mark.parametrize("cmd", SPACING_COMMANDS)
def test_spacing_command_is_not_silently_dropped(cmd):
    r"""反面判据：间距不是被「抹平」了，而是真的落成了一个字符。

    与上一条互补 —— 上一条只看「有没有 NBSP」，这一条看「文本节点不是
    a、b 直接相邻」。两条同时绿，才说明命令被处理过而不是被删除。
    """
    texts = _texts(_convert(_spaced(cmd)))
    joined = "".join(texts)
    assert joined != "ab", "%s 的间距被静默删除（m:t=%r）" % (_spaced(cmd), texts)


# ── (2) 两条与直觉相反的现状（照实钉住，改则须连带改文档）───────────────


def test_negative_thin_space_also_becomes_nbsp():
    r"""`\!` 落成正 NBSP —— 负间距语义在 keep_spacing 这一步就丢了。

    LaTeX 里 `\!` 是负细间距；本仓库把它与 `\,` 一视同仁改写成控制空格。
    这是现状而非期望：钉住是为了让「修正负间距」变成一个显式决定（改本用例 +
    CHANGELOG），而不是顺手改正则之后无人知晓。
    """
    el = _convert(_spaced(r"\!") + " + 1")
    texts = _texts(el)
    assert NBSP in texts, r"\! 现状应产 NBSP：m:t=%r" % (texts,)
    assert not any("-" in t for t in texts), "不该凭空插入负号：%r" % (texts,)


def test_qquad_currently_equals_quad_single_nbsp():
    r"""`\qquad` 与 `\quad` 产物相同（各一个 NBSP），不是 LaTeX 语义。

    成因：正则交替是 `(?:qquad|quad|,|;|:|!)`，`qquad` 先命中、被改写成单个
    控制空格，所以两者等价。LaTeX 里 `\qquad` 应约等于两个 `\quad`。
    钉住现状；要修正则本用例一起改（报告 §8 记了这条作为 P30 候选）。
    """
    quad_texts = _texts(_convert(_spaced(r"\quad")))
    qquad_texts = _texts(_convert(_spaced(r"\qquad")))
    assert quad_texts == qquad_texts, "\\quad 与 \\qquad 产物分叉了：%r vs %r" % (
        quad_texts, qquad_texts)
    assert quad_texts.count(NBSP) == 1


# ── (3) 承重墙证据：关掉改写，间距就没了 ────────────────────────────────


def test_without_the_rewrite_spacing_vanishes():
    r"""`raw_spacing=True` 时 `a\quad b` 的间距直接消失 —— 这是 keep_spacing
    存在的理由，也是「这条契约不是形式主义」的证据。

    这条用例断言的是**当前**降级行为。若哪天 XSLT 或 latex2mathml 开始原生保留
    间距，它会红 —— 那时说明 keep_spacing 可以退休了，属预期变更。
    """
    el = math_omml.latex_to_omml_safe(r"a\quad b", raw_spacing=True)
    texts = _texts(el)
    assert not any(NBSP in t for t in texts), (
        "raw_spacing 下竟然还有 NBSP：%r" % (texts,))
    assert texts == ["a", "b"], "对照组的产物形状变了：%r" % (texts,)
