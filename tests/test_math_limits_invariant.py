# -*- coding: utf-8 -*-
r"""大算符 `\limits` 不变式：写与不写 `\limits` 必须落成**不同的 OMML 结构**（P29-C）。

## 这条测试钉的是什么

`\limits` 是排版开关：带它 = 上下限排在算子正上 / 正下，不带 = 排在右侧当上下标。
经 `latex2mathml` → `MML2OMML.XSL` 之后，这个区别落成两处**不同的结构**，
不是「看起来差不多」：

| 算子族 | 不带 `\limits` | 带 `\limits` |
|---|---|---|
| nary 族（`\sum` `\prod` `\int` …） | `m:nary` + `m:limLoc/@m:val="subSup"` | `m:nary` + `m:limLoc/@m:val="undOvr"` |
| limLow 族（`\max` `\lim` `\det` …） | `m:sSub` | `m:limLow` |

P28 有一轮把「写不写 `\limits` 结果一样」写进了文档，错了整整一轮才被推翻
（教训 179：要么实测结构差异，要么别写）。本文件是那条教训的落地 ——
**用 lxml 结构查询断言元素名与属性值，不比字符串**。

## 为什么不比字符串

序列化结果里混着字体注入（`w:rFonts`）、直立补全（`m:sty`）与 run 拆分，
任何一处排版微调都会让字符串比对变红 —— 那测的是排版实现，不是 `\limits` 的
语义。结构查询只认「哪个元素、哪个属性」，改排版不动它。

## 三态为什么必须一起测

只测「带 `\limits` 变成 `undOvr`」，会被「三态全变 `undOvr`」的回归骗过
（那种改动同样让它变绿，而 `\lim` 类公式在 Word 里会全部塌成上下排在正上）。
所以每条断言都配同算子的 plain / nolimits 对照，`test_limits_is_not_a_no_op`
直接把两态并排断言「不一样」。

`\nolimits` 实测与「什么都不写」完全一致（`subSup` / `m:sSub`），这是
latex2mathml 的既有行为，本文件照实钉住。

实测矩阵与结构 dump 见 `docs/reports/P29-C-TESTS.md` 第 2 节。
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

#: 整份文件都要真跑 XSLT：断言的是 OMML 结构，降级成纯文字时一条都成立不了。
pytestmark = needs_omml


def _q(tag: str) -> str:
    return "{%s}%s" % (M, tag)


# ── 算子分族（P29-C 实测，不是凭 LaTeX 知识猜的）─────────────────────────────
#
# 两族在 OMML 侧的可观测差异不同，所以断言方式不同：
#   * nary 族 → 靠 m:nary 的子元素 m:limLoc 的 @m:val 区分；
#   * limLow 族 → 结构本身换元素（m:sSub ↔ m:limLow），没有 limLoc 可看。
#
# ⚠️ 分族名单由 work/p29c/probe_limits_spacing.py 打表得出（25 个
# BIG_OPERATORS × 3 态 = 75 行，零异常零泄漏）。**不要按 LaTeX 直觉改这里** ——
# 直觉上 \bigoplus 该是 nary，实测它走 limLow 族（见报告第 4 节）。
NARY_OPERATORS = (
    "sum", "prod", "coprod", "int", "iint", "iiint", "oint",
    "bigcup", "bigcap", "bigvee", "bigwedge", "bigotimes",
)

LIMLOW_OPERATORS = (
    "bigoplus", "bigodot", "biguplus",
    "lim", "limsup", "liminf", "sup", "inf",
    "max", "min", "det", "gcd", "Pr",
)


def _convert(latex: str):
    return math_omml.latex_to_omml_safe(latex)


def _lim_locs(el) -> list:
    """全部 m:limLoc/@m:val，按文档顺序。"""
    return [n.get(_q("val")) for n in el.iter(_q("limLoc"))]


def _texts(el) -> list:
    return [t.text or "" for t in el.iter(_q("t"))]


def _has(el, tag: str) -> bool:
    return el.find(".//" + _q(tag)) is not None


def _nary_latex(op: str, state: str) -> str:
    return "\\%s%s_{i=1}^{n} x" % (op, "" if state == "plain" else "\\" + state)


def _limlow_latex(op: str, state: str) -> str:
    return "\\%s%s_{x}" % (op, "" if state == "plain" else "\\" + state)


# ── (1) 单个算子三态：\sum 走 nary，看 limLoc ────────────────────────────


def test_sum_limits_sets_limloc_undovr():
    r"""\sum\limits → m:nary 且 m:limLoc/@m:val 为 "undOvr"。"""
    el = _convert(r"\sum\limits_{i=1}^{n} a_i")
    assert _has(el, "nary"), "带 \\limits 的 \\sum 必须仍是 m:nary"
    assert _lim_locs(el) == ["undOvr"]


def test_sum_plain_sets_limloc_subsup():
    r"""\sum（不加修饰符）→ 同一个 m:limLoc 取 subSup。"""
    el = _convert(r"\sum_{i=1}^{n} a_i")
    assert _has(el, "nary")
    assert _lim_locs(el) == ["subSup"]


def test_sum_nolimits_sets_limloc_subsup():
    r"""\sum\nolimits 与不加修饰符一致（实测如此，不是推测）。"""
    el = _convert(r"\sum\nolimits_{i=1}^{n} a_i")
    assert _has(el, "nary")
    assert _lim_locs(el) == ["subSup"]


# ── (2) 单个算子三态：\max 走 limLow 族，看元素本身 ──────────────────────


def test_max_limits_becomes_limlow():
    r"""\max\limits → m:limLow，且不再有 m:sSub。"""
    el = _convert(r"\max\limits_{x}")
    assert _has(el, "limLow"), "带 \\limits 的 \\max 必须落成 m:limLow（上下排在正下）"
    assert not _has(el, "sSub"), "同一层不该同时留下右侧下标 m:sSub"


def test_max_plain_has_no_limlow():
    r"""\max_{x} 只有 m:sSub，m:limLow 必须不存在。"""
    el = _convert(r"\max_{x}")
    assert _has(el, "sSub")
    assert el.find(".//" + _q("limLow")) is None
    assert _lim_locs(el) == []


def test_max_nolimits_has_no_limlow():
    r"""\max\nolimits 与不加修饰符一致：仍是 m:sSub。"""
    el = _convert(r"\max\nolimits_{x}")
    assert _has(el, "sSub")
    assert el.find(".//" + _q("limLow")) is None


# ── (3) 总纲：两态确实不同（防「三态全一样」的回归）─────────────────────


def test_limits_is_not_a_no_op():
    r"""\limits 不是空操作：nary 族换 limLoc，limLow 族换元素。

    这一条就是 P28 那条错误文档的反面 —— 如果哪天有人把 \limits 擦成普通
    下标（三态归一），前六条里「带 \limits」那几条会一起变红，而这一条会同时
    红在两族上，定位最快。
    """
    plain = _convert(r"\sum_{i=1}^{n} a_i")
    limited = _convert(r"\sum\limits_{i=1}^{n} a_i")
    assert _lim_locs(plain) == ["subSup"]
    assert _lim_locs(limited) == ["undOvr"]
    assert _lim_locs(plain) != _lim_locs(limited)

    plain_max = _convert(r"\max_{x}")
    limited_max = _convert(r"\max\limits_{x}")
    assert _has(plain_max, "sSub") and not _has(plain_max, "limLow")
    assert _has(limited_max, "limLow") and not _has(limited_max, "sSub")


# ── (4) nary 族全量：12 个算子 × 3 态 ────────────────────────────────────


@pytest.mark.parametrize("op", NARY_OPERATORS)
def test_nary_operator_plain_is_subsup(op):
    el = _convert(_nary_latex(op, "plain"))
    assert _has(el, "nary"), "\\%s 应走 m:nary 族" % op
    assert _lim_locs(el) == ["subSup"]


@pytest.mark.parametrize("op", NARY_OPERATORS)
def test_nary_operator_limits_is_undovr(op):
    el = _convert(_nary_latex(op, "limits"))
    assert _has(el, "nary"), "\\%s\\limits 不应改掉 m:nary 本体" % op
    assert _lim_locs(el) == ["undOvr"], "\\%s\\limits 必须显式写 m:limLoc=undOvr" % op


@pytest.mark.parametrize("op", NARY_OPERATORS)
def test_nary_operator_nolimits_is_subsup(op):
    el = _convert(_nary_latex(op, "nolimits"))
    assert _has(el, "nary")
    assert _lim_locs(el) == ["subSup"]


# ── (5) limLow 族全量：13 个算子 × 3 态 ──────────────────────────────────


@pytest.mark.parametrize("op", LIMLOW_OPERATORS)
def test_limlow_operator_plain_stays_ssub(op):
    el = _convert(_limlow_latex(op, "plain"))
    assert _has(el, "sSub"), "\\%s 不带 \\limits 应为 m:sSub" % op
    assert el.find(".//" + _q("limLow")) is None


@pytest.mark.parametrize("op", LIMLOW_OPERATORS)
def test_limlow_operator_limits_becomes_limlow(op):
    el = _convert(_limlow_latex(op, "limits"))
    assert _has(el, "limLow"), "\\%s\\limits 必须落成 m:limLow" % op
    assert not _has(el, "sSub")


@pytest.mark.parametrize("op", LIMLOW_OPERATORS)
def test_limlow_operator_nolimits_stays_ssub(op):
    el = _convert(_limlow_latex(op, "nolimits"))
    assert _has(el, "sSub")
    assert el.find(".//" + _q("limLow")) is None


# ── (6) 跨全族：三态都转得出、不抛、不泄漏 ───────────────────────────────

ALL_STATES = ("plain", "limits", "nolimits")


@pytest.mark.parametrize("op", NARY_OPERATORS + LIMLOW_OPERATORS)
@pytest.mark.parametrize("state", ALL_STATES)
def test_no_operator_leaks_the_modifier(op, state):
    r"""\limits / \nolimits 不能作为字面量漏进正文（层 2 残留兜底的对照）。

    泄漏形态是成品 m:t 里出现字符串 limits —— 那会在 Word 里真印出来。
    三种状态各测一遍：\limits 合法不等于 \nolimits 也合法（P27 实测
    x\nolimits 就泄漏）。
    """
    latex = (_nary_latex(op, state) if op in NARY_OPERATORS
             else _limlow_latex(op, state))
    el = _convert(latex)
    leaked = [t for t in _texts(el) if "limits" in t]
    assert not leaked, "%s 在 %s 态泄漏修饰符字面量：%r" % (latex, state, leaked)


# ── (7) 自洽：分族名单必须覆盖 BIG_OPERATORS ─────────────────────────────


def test_operator_tables_cover_every_big_operator():
    r"""BIG_OPERATORS 与两个分族表必须互为子集（新增算子要当场分族）。

    BIG_OPERATORS 是 \limits 合法性的白名单（scripts/math_omml.py）。有人往里
    加算子而忘了分族时本用例会红 —— 比让新算子悄悄少一层覆盖好。
    """
    declared = set(math_omml.BIG_OPERATORS)
    covered = set(NARY_OPERATORS) | set(LIMLOW_OPERATORS)
    assert declared == covered, "未分族：%r；表里多余：%r" % (
        sorted(declared - covered), sorted(covered - declared))
    assert not (set(NARY_OPERATORS) & set(LIMLOW_OPERATORS)), "两族不许重叠"
