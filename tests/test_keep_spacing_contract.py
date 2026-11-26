# -*- coding: utf-8 -*-
r"""间距改写表的纯函数契约：`keep_spacing()` 的自洽性（P29-C 拆出）。

## 这个文件为什么**不加任何门控标记**

本文件只调 `scripts/math_omml.py` 的 `keep_spacing()` —— 一次正则替换，
不 import `latex2mathml`、不进 `xslt()`、不碰 `lxml`。所以它在**任何**
机器上都该真跑，包括没有 Office、没有 `MML2OMML.XSL` 的 CI 腿。

这条不是洁癖。P27 的教训 157：本机 +158 条、最弱 CI 腿只 +65，差 93 全是被
`needs_omml` 吃掉的 skip —— 用例写了等于没写。这批既然不依赖 XSLT，就不该
让它在 CI 上假装不存在。**加门控前先看调用链**：只要不进 `lconv.convert` /
`xslt()`，就不要 `needs_omml`。

## 钉的是什么

「哪些间距命令能活下来」完全由 `keep_spacing()` 那一条正则决定
（`scripts/math_omml.py` 里的 `re.sub(r"\\(?:qquad|quad|,|;|:|!)", r"\\ ", latex)`）。
把它**改窄**（少列一条命令）或**改宽**（误伤别的命令）都不会有别的测试变红，
所以这里两个方向各钉一遍：

* 窄：清单里每条命令都必须被改写成控制空格 `\ `，且原命令名不许残留；
* 宽：不相干的命令（`\alpha` / `\frac` / 上标 …）必须**原样不动**。

成品层面的契约（转出来 `m:t` 里真有 NBSP）在
`tests/test_math_spacing_contract.py`，那部分必须真跑 XSLT，仍带 `needs_omml`。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import math_omml  # noqa: E402

#: 间距命令全集 —— 与 keep_spacing() 正则里的交替完全一致。
#: 少一条就少一层覆盖（窄方向）；多一条正则不认的命令会被这里抓出来。
SPACING_COMMANDS = (
    r"\quad",
    r"\qquad",
    r"\,",
    r"\;",
    r"\:",
    r"\!",
    r"\ ",      # 控制空格：本身就是要改写成的目标形态
)

#: 不该被这条正则碰的输入（宽方向的反面钉子）。
UNTOUCHED = (
    r"\alpha",
    r"\frac{a}{b}",
    r"x^2 + y_1",
    r"\sum_{i=1}^{n} a_i",
)


def _spaced(cmd: str) -> str:
    """把命令插进两个标识符之间：拿空格就得靠命令本身，不能靠源文本里的空格。"""
    return "a%sb" % cmd


# ── 窄方向：清单里每条命令都必须被改写 ──────────────────────────────────


@pytest.mark.parametrize("cmd", SPACING_COMMANDS)
def test_keep_spacing_rewrites_to_control_space(cmd):
    r"""每条间距命令都被改写成 `\ `（控制空格），一条不漏。

    这是「活下来」的前提：漏改的命令会交给 latex2mathml 自己处理，而它要么丢
    （`\quad`）要么当未知命令泄漏（紧贴标识符时）。
    """
    assert math_omml.keep_spacing(cmd) == "\\ ", "keep_spacing 没改写 %r" % cmd


@pytest.mark.parametrize("cmd", SPACING_COMMANDS)
def test_keep_spacing_leaves_no_command_name(cmd):
    r"""改写后不再残留原命令名（`\quad` 之类不能被当正文留下），两操作数都在。"""
    out = math_omml.keep_spacing(_spaced(cmd))
    name = cmd[1:]
    if name.strip():                      # 控制空格没有命令名可查
        assert name not in out, "改写后仍残留 %r：%r" % (name, out)
    assert out.startswith("a") and out.endswith("b")


# ── 宽方向：不相干的命令必须原样不动 ────────────────────────────────────


@pytest.mark.parametrize("latex", UNTOUCHED)
def test_keep_spacing_does_not_touch_other_commands(latex):
    r"""正则不许改宽：非间距命令必须一个字符都不变。

    反向覆盖容易漏 —— 把交替写成 `\\.+` 之类同样能过「窄方向」那 14 条，
    却会把整条公式吃掉。这条用例把那种回归变成红的。
    """
    assert math_omml.keep_spacing(latex) == latex


# ── 逐字符级：替换只发生在命令位置，两侧文本不受影响 ───────────────────


def test_keep_spacing_only_replaces_the_command_itself():
    r"""`a\quad b` → `a\  b`：左操作数、右操作数、命令位置三者分别核对。"""
    assert math_omml.keep_spacing(r"a\quad b") == "a\\  b"
