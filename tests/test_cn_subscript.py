# -*- coding: utf-8 -*-
r"""`S甲` → `S_{甲}`: Chinese single-character subscripts, behind an opt-in switch.

Reported by the owner: `S甲` / `S乙` came out at baseline height, so a paper
reading `S甲 > S乙` looks like plain text instead of two labelled variables.

The hard part is **not** the rewrite — it is not corrupting Chinese prose.
`m与`、`m高`、`t时`、`图甲`、`表乙` have exactly the same shape (ASCII letter +
one CJK char) and are ordinary text.  So:

* the feature is **off by default** (`linear_to_latex(s)` is unchanged);
* when on, a blacklist of connectives / positional words / copulas blocks the
  prose cases;
* the assertion requires the letter to be a *single* ASCII letter, so `ab甲`
  and `\alpha甲` never match.

Both directions are pinned below — a positive case for each rewrite and a
negative case for each prose pattern.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402


def conv_on(s: str) -> str:
    """Explicit per-call switch."""
    return L.linear_to_latex(s, cn_subscript=True)


def conv_off(s: str) -> str:
    return L.linear_to_latex(s)


# ---------------------------------------------------------------- 默认行为不变
@pytest.mark.parametrize("src", [
    "S甲", "v甲>v乙", "P甲=P乙", "S甲=S乙",
])
def test_default_call_leaves_chinese_untouched(src):
    """★ 默认必须是关的 —— 否则会悄悄改掉现有六卷的产出。"""
    assert conv_off(src) == src


# ---------------------------------------------------------------- 正例：该转的
@pytest.mark.parametrize("src, expected", [
    ("S甲", r"S_{甲}"),
    ("S乙", r"S_{乙}"),
    ("v甲>v乙", r"v_{甲}>v_{乙}"),
    ("P甲=P乙", r"P_{甲}=P_{乙}"),
    ("S甲>S乙", r"S_{甲}>S_{乙}"),
    # 物理卷常见：v甲 与 v乙 比较
    ("v甲=2v乙", r"v_{甲}=2v_{乙}"),
    # 单字母小写也应生效
    ("x丙", r"x_{丙}"),
])
def test_letter_plus_cjk_becomes_subscript_when_enabled(src, expected):
    assert conv_on(src) == expected


# ---------------------------------------------------------------- 反例：正文不许动
@pytest.mark.parametrize("src", [
    # ★ 主人点名要锁死的正文用例
    "m与",
    "m高",
    "t时",
    "图甲",
    "表乙",
    # 更多虚词 / 方位词 / 判断词
    "m和",
    "s的",
    "v为",
    "S是",
    "x中",
    "y内",
    "z外",
    "P上",
    "Q下",
    "f后",
    "g前",
    # 量名词
    "S值",
    "N数",
    "m量",
])
def test_prose_patterns_are_never_turned_into_subscripts(src):
    assert conv_on(src) == src


@pytest.mark.parametrize("src", [
    "ab甲",        # 字母不是「单个」：前置断言挡住
    r"\alpha甲",   # 命令尾字母后面也挡住
])
def test_multi_letter_or_command_tail_is_not_matched(src):
    out = conv_on(src)
    assert "_{" not in out, out


# ---------------------------------------------------------------- 与 ASCII 规则共存
def test_ascii_subscript_still_works_when_chinese_is_on():
    """开关只加中文规则，不能挤掉既有的 `v0` → `v_{0}`。"""
    assert conv_on("v0") == r"v_{0}"
    assert conv_on("H_2O") == r"H_{2}O"


def test_ascii_and_chinese_subscripts_in_one_formula():
    out = conv_on("S甲=S0")
    assert out == r"S_{甲}=S_{0}", out


# ---------------------------------------------------------------- 黑名单本身
def test_blacklist_is_not_empty_and_covers_the_reported_cases():
    for ch in "与和的时高长宽为是":
        assert ch in L.CN_SUBSCRIPT_BLACKLIST, ch


def test_module_level_switch_flips_the_default(monkeypatch):
    """模块级 CN_SUBSCRIPT 打开后，不带参数调用也生效（给 docx_create 批量用）。"""
    monkeypatch.setattr(L, "CN_SUBSCRIPT", True)
    assert L.linear_to_latex("S甲") == r"S_{甲}"
    # 显式参数仍能压过模块级开关
    assert L.linear_to_latex("S甲", cn_subscript=False) == "S甲"
