# -*- coding: utf-8 -*-
r"""空 `m:oMath` 的两道防线（P29-A）。

## 修的是什么

`\rule{10pt}{2pt}` / `\hspace{1em}` 一族经 `latex2mathml` 变成 `<mspace>`，
而 `MML2OMML.XSL` **全文没有 mspace 模板**，于是整条公式被吞成
`<m:oMath/>`（零子元素）。旧链路四环全沉默：`insert_formula` rc=0、
`docx_validate` rc=0、`docx_mathify` 记成功、`assert_legal` 只看 `m:d` 的槽。
Word 里就是一片空白。

本文件钉两件事：

1. **层 1 硬拒**——间距 / 画线那一族在 `latex2mathml` 之前就抛（`needs_l2m`）；
2. **第 4 类结构判据**——`assert_legal` 见 `m:oMath` 自身无子元素就抛
   `OMMLStructureError`，兜住层 1 拦不到的空组 `{}` 与绕开正常链路的 OMML。

## 门控为什么这么标

- 层 1 的拒收发生在 `lconv.convert` 之前（`math_omml.py:1041` 早于 `:1043` 的
  `keep_spacing` 与 `:1045` 的 `lconv.convert`），
  **不需要** Office 的 XSL → `needs_l2m`；标成 `needs_omml` 会在无 Office 的
  runner 上白 skip，等于没测（P27 首轮四腿全红的教训）。
- 凡要真跑 XSLT 的（真转换产物）→ `needs_omml`。

`needs_l2m` 本文件**自建**，与 `tests/test_math_latex_guard.py:63-73` 同源：
不 import 那个模块，是为了避免测试文件之间互相 import 造成 collection 顺序耦合
（那边一旦改名 / 拆文件，这里会在收集期炸）。

## 不许误伤（这条比「全部拒」更重要）

`\sqrt{}` 产 `<m:rad><m:e/></m:rad>`、`\frac{}{}` 产空的 `m:num` / `m:den` ——
那是**别的语义下的空**，`m:oMath` 仍有子元素。判据只看 `m:oMath` 自身一个子元素
都没有，绝不外扩成「任何空 m:e 都非法」；`tests/test_math_delim_slots.py:322-341`
正钉着这一条。
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


def _has_latex2mathml() -> bool:
    """层 1 只依赖 latex2mathml（在它之前就抛），不需要 Office。"""
    try:
        import latex2mathml  # noqa: F401
    except Exception:
        return False
    return True


needs_l2m = pytest.mark.skipif(
    not _has_latex2mathml(), reason="needs latex2mathml")


def convert(latex: str, **kw):
    return math_omml.latex_to_omml_lxml(latex, **kw)


# ===========================================================================
# 1. 层 1：间距 / 画线一族硬拒（不需要 XSL）
# ===========================================================================
# 星号变体（\hspace*{1em}）刻意单列一条：`_CMD_RE` 是 `\\([A-Za-z]+)(\*?)`，
# 星号被第二组吃掉、命令名仍是 hspace —— 但这是**读正则得出的结论**，P29-A 实测
# 确认过：它同样命中表、同样被拒（另外它不产空公式，而是泄漏 `<m:t>1em</m:t>`，
# 见 docs/reports/P29-A-GUARD.md）。
SPACING_AND_RULE = [
    (r"\rule", r"\rule{10pt}{2pt}"),
    (r"\hspace", r"\hspace{1em}"),
    (r"\hspace", r"\hspace*{1em}"),
    (r"\hspace", r"a \hspace{1em} b"),
    (r"\hskip", r"\hskip 1em"),
    (r"\kern", r"\kern 1pt"),
    (r"\mspace", r"\mspace{1mu}"),
    (r"\enspace", r"\enspace"),
]


@needs_l2m
@pytest.mark.parametrize("command,latex", SPACING_AND_RULE,
                         ids=[repr(p[1]) for p in SPACING_AND_RULE])
def test_spacing_family_is_rejected_before_conversion(command, latex):
    r"""间距 / 画线一族必须由层 1 拒收，且消息里点名那条命令。

    只断言「抛了」不够：这些输入在旧链路上**完全不抛**（静默产空公式），所以
    抛 RuntimeError 本身就证明守卫生效；但仍断言命令名，防止将来某条被上游的
    语法异常顶替掉。
    """
    with pytest.raises(math_omml.LatexGuardError) as ei:
        convert(latex)
    msg = str(ei.value)
    assert command in msg, (command, msg[:400])


# ===========================================================================
# 2. 第 4 类结构判据：空 `m:oMath`
# ===========================================================================
@needs_omml
def test_empty_group_produces_empty_omath_and_raises():
    r"""`{}` 单独写 → `assert_legal` 抛 `OMMLStructureError`。

    层 1 拦不到它（`{}` 里没有控制序列），所以这是第 4 类判据的唯一真实入口。
    实测：`latex2mathml` 对 `{}` 给 `<mrow/>`，XSLT 之后就是零子元素的
    `<m:oMath/>`。
    """
    with pytest.raises(math_omml.OMMLStructureError) as ei:
        convert("{}")
    msg = str(ei.value)
    assert "空公式" in msg, "文案要说清是空公式"
    assert "oMath" in msg


@needs_omml
def test_assert_legal_rejects_bare_empty_omath():
    r"""手搭一个无子元素的 `m:oMath` → `assert_legal` 直接拒（不经转换）。

    这一条隔离得更彻底：它证明拦截点在 `assert_legal` 本身，而不是某个上游
    恰好在别处抛了。真实转换那条（上面）只证明链路末端会响。
    """
    from lxml import etree

    el = etree.Element(_m("oMath"))
    with pytest.raises(math_omml.OMMLStructureError) as ei:
        math_omml.assert_legal(el, "{}")
    assert "空公式" in str(ei.value)


@needs_omml
def test_empty_omath_in_display_wrapper_is_caught():
    r"""`display=True` 时根是 `m:oMathPara`，里面那个空 `m:oMath` 同样要被抓。

    `assert_legal` 用的是 `omml.iter()`，会遍历到嵌套的 `m:oMath` —— 但这条
    容易被「只看根元素」的实现漏掉，钉住。
    """
    from lxml import etree

    para = etree.Element(_m("oMathPara"))
    etree.SubElement(para, _m("oMathParaPr"))
    etree.SubElement(para, _m("oMath"))
    with pytest.raises(math_omml.OMMLStructureError):
        math_omml.assert_legal(para, "{}")


# ===========================================================================
# 3. 不许误伤：别的语义下的空 + 干净公式
# ===========================================================================
@needs_omml
@pytest.mark.parametrize("latex", [r"\sqrt{}", r"\frac{}{}", r"\sqrt{\ }"])
def test_empty_slots_of_other_semantics_are_not_flagged(latex):
    r"""`\sqrt{}` / `\frac{}{}` / `\sqrt{\ }` 一个都不许抛。

    它们产的是空的 `m:rad/m:e`、`m:num` / `m:den` —— 「别的语义下的空」，
    `m:oMath` 本身仍有子元素（`m:rad` / `m:f`）。收成「任何空 m:e 都非法」
    会把它们打成错误，`tests/test_math_delim_slots.py:322-341` 正钉着不许误伤。
    """
    el = convert(latex)
    math_omml.assert_legal(el, latex)      # 不抛即通过


@needs_omml
def test_clean_formula_is_not_flagged():
    r"""`E=mc^2` 这种干净公式必须过第 4 类判据（防守卫过严）。"""
    el = convert("E=mc^2")
    math_omml.assert_legal(el, "E=mc^2")


@needs_l2m
def test_quad_family_still_works():
    r"""`\quad` 一族仍能过层 1 —— 它们是本链路唯一能活的间距。

    不标 `needs_omml`：这条只验层 1 不误伤。真转换在别的文件里覆盖。
    """
    assert math_omml.find_unsupported_commands(r"a \quad b") == []

