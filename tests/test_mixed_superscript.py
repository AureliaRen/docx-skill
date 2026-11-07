# -*- coding: utf-8 -*-
r"""混合角标策略：纯上下标走「文字角标」，真结构公式留在 OMML。

主人的判据（HANDOFF-NEXT §一.2）：

    「选中角标，按 Ctrl+Shift+= 后**它应该返回到正常的大小**，这才是真的角标，
      但现在是按了之后变得更小了。」

OMML 的 `m:sSup` 做不到这一点 —— 它是公式引擎排的小号字，Word 里
`Font.Superscript` 恒为 `0`，再按上标只会在小号字上再叠一层。
真角标是 run 级 `<w:rPr><w:vertAlign>`：`Font.Superscript == -1`、字号不变、
`OMaths.Count == 0`（Word 16.0 真机取证，见 HANDOFF-NEXT §2.1）。

分流规则（`docx_mathify.MIXED_SUPERSCRIPT`，默认 True）：

* 带 `^`/`_` 且**无二维结构** → 普通 run + `w:vertAlign`；
* 有 `\frac` `\sqrt` `\int` `\sum` `\begin{` `\left|` 等 → 仍走
  `latex_to_omml_lxml`，行为不变（分式降级成一串平铺字符是不可接受的）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx.oxml.ns import qn  # noqa: E402

import docx_mathify as M  # noqa: E402
import linear_to_latex as L  # noqa: E402


def _omml_available() -> bool:
    """本机有没有 LaTeX→OMML 的最后一环（latex2mathml + MML2OMML.XSL）。

    Hosted CI runners have neither Office nor the stylesheet, so every test that
    asserts "this fragment became ``m:oMath``" has to skip there — otherwise the
    suite is green only on a developer machine with Word installed.  That is a
    false green light, and it hid a real failure once.
    """
    try:
        import math_omml
        math_omml.find_xslt()
    except Exception:
        return False
    return True


needs_omml = pytest.mark.skipif(
    not _omml_available(),
    reason="needs latex2mathml + Microsoft MML2OMML.XSL (local Office install)")


# ---------------------------------------------------------------- 判据：正例
@pytest.mark.parametrize("latex", [
    r"a^{2}",
    r"x^{3}",
    r"v_{0}",
    r"a_{n}",
    r"b_{n}=2^{n+1}",
    r"b_{n+1}=2a_{n}+6=2b_{n}",
    r"S_{n}=2^{n+2}-3n-4",
    r"10^{-7}",                       # 主人报的原始 bug
    r"10^{+3}",
    r"1.0\times 10^{-7}",
    r"S=\pi r^{2}=9\pi",              # 希腊字母可展开
    r"a=9.8\,\mathrm{m/s^{2}}",       # 量纲串也要降级（\mathrm 是一维正体）
    r"\sin30^{\circ}=0.50",           # 函数名 + 度数上标
])
def test_pure_script_is_recognised(latex):
    ok, why = M.is_pure_script_detail(latex)
    assert ok, why
    assert why == "pure_script"


# ---------------------------------------------------------------- 判据：反例
@pytest.mark.parametrize("latex, reason_start", [
    # ★ 二维结构一律不许降级（每条都带 ^ 或 _，确保测的是结构而非 no_script）
    (r"\frac{x^{2}}{16}+\frac{y^{2}}{9}=1", "structural:\\frac"),
    (r"y^{2}=\frac{81}{16}", "structural:\\frac"),
    (r"c=\frac{\sqrt{7}}{4}^{2}", "structural:\\frac"),
    (r"c=\sqrt{7}^{2}", "structural:\\sqrt"),
    (r"\int_{0}^{1}x\,dx", "structural:\\int"),
    (r"\sum_{n=1}^{N}n", "structural:\\sum"),
    (r"\begin{cases}x^{2}>0\end{cases}", "structural:\\begin"),
    (r"\left|x\right|^{2}", "structural:\\left"),
    (r"\overline{AB}^{2}", "structural:\\overline"),
])
def test_structural_formulas_stay_omml(latex, reason_start):
    ok, why = M.is_pure_script_detail(latex)
    assert not ok
    assert why == reason_start, why


@pytest.mark.parametrize("latex", [
    r"c=\sqrt{7}",
    r"\begin{cases}x>0\end{cases}",
    r"\left|x\right|",
    r"\overline{AB}",
])
def test_structure_without_scripts_is_still_omml(latex):
    r"""没有 `^`/`_` 的也能判成 OMML —— 理由记作 no_script，但结论一致。"""
    ok, why = M.is_pure_script_detail(latex)
    assert not ok
    assert why in ("no_script", "structural:" + latex.split("{")[0].lstrip("\\"))


@pytest.mark.parametrize("latex", [
    "y=f(x)",          # 没有角标：转了也只是换字体，白改
    "A=(-1,3)",
    "s=180km",
    "",
])
def test_spans_without_scripts_are_not_touched(latex):
    ok, why = M.is_pure_script_detail(latex)
    assert not ok
    assert why in ("no_script", "empty")


def test_infty_is_not_an_integral():
    r"""`\infty` 里含 `\in`，但命令边界判定不该把它当 `\int`。"""
    assert M.structural_reason(r"x\to+\infty") is None
    assert M.structural_reason(r"\int_0^1") == r"\int"


# ---------------------------------------------------------------- 展开：逐段对
@pytest.mark.parametrize("latex, expected", [
    (r"a^{2}", [("a", None), ("2", "superscript")]),
    (r"v_{0}", [("v", None), ("0", "subscript")]),
    (r"10^{-7}", [("10", None), ("-7", "superscript")]),
    (r"v_{0}^{2}", [("v", None), ("0", "subscript"), ("2", "superscript")]),
    # \mathrm 组展开成正体字符；`\,` 是显式窄空格，落成 NBSP
    (r"a=9.8\,\mathrm{m/s^{2}}",
     [("a=9.8\u00a0m/s", None), ("2", "superscript")]),
])
def test_latex_expands_to_text_pieces(latex, expected):
    assert M._latex_to_pieces(latex) == expected


# ---------------------------------------------------------------- run 生成
def test_script_runs_carry_vert_align():
    r"""产出的 run 必须带 `<w:vertAlign>`，且基线段不带。"""
    runs = M._script_runs(r"10^{-7}", None)
    assert runs is not None and len(runs) == 2
    assert runs[0].find(qn("w:rPr")) is None          # `10`：基线，无 rPr
    rpr = runs[1].find(qn("w:rPr"))
    va = rpr.find(qn("w:vertAlign"))
    assert va.get(qn("w:val")) == "superscript"
    assert "".join(t.text for t in runs[1].iter(qn("w:t"))) == "-7"


def test_script_runs_use_rpr_order():
    r"""`w:vertAlign` 必须按 CT_RPr 顺序插入，否则 Word 静默丢弃。

    这里复用 docx_crossref.with_vert_align：给它一个带 `w:sz` 的 rPr，
    `w:vertAlign` 要落在 `w:sz` **之后**（RPR_ORDER 里 sz 在 vertAlign 前）。
    """
    from docx.oxml import OxmlElement
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rpr.append(sz)

    runs = M._script_runs(r"x^{2}", rpr)
    out_rpr = runs[1].find(qn("w:rPr"))
    tags = [c.tag.split("}")[-1] for c in out_rpr]
    if "sz" in tags and "vertAlign" in tags:
        assert tags.index("sz") < tags.index("vertAlign"), tags


def test_subscript_run_uses_subscript_value():
    runs = M._script_runs(r"S_{n}", None)
    va = runs[1].find(qn("w:rPr")).find(qn("w:vertAlign"))
    assert va.get(qn("w:val")) == "subscript"


# ---------------------------------------------------------------- 空格口径
# LaTeX 里 `\times 10` 中间那个空格是**控制字的定界符**：它的唯一作用是结束
# `\times` 这个命令名，TeX 渲染时吞掉，屏幕上不存在。把它当文字空格搬进 run，
# 就会渲染出 `K=1.0× 10²`（Lead 在化学卷实测到的 bug）。
#
# 反过来，`\,` `\:` `\;` `\quad` 是**显式间距命令**，是作者有意留的，
# 必须落成真的空格（这里用 NBSP），不能吃成空串。
@pytest.mark.parametrize("latex, visible", [
    # 控制字后的定界空格必须消失
    (r"K=1.0\times 10^{2}", "K=1.0×102"),
    (r"3\times 10^{-7}", "3×10-7"),
    (r"a\leq 10^{2}", "a≤102"),
    (r"\sin 30^{\circ}", "sin30°"),
    (r"\alpha ^{2}", "α2"),
    # 显式间距必须保留（NBSP，不是普通空格、更不是空串）
    (r"a=9.8\,\mathrm{m/s^{2}}", "a=9.8\u00a0m/s2"),
    (r"M=18\,\text{g/mol}^{2}", "M=18\u00a0g/mol2"),
    (r"x\quad y^{2}", "x\u00a0y2"),
    # `\text{}` 内部的空格是正文，必须原样保留
    (r"\text{a b}^{2}", "a b2"),
    # `\ ` 是显式空格命令
    (r"v_{0}\ \text{m/s}", "v0 m/s"),
])
def test_spacing_follows_tex_semantics(latex, visible):
    pieces = M._latex_to_pieces(latex)
    assert pieces is not None
    assert "".join(t for t, _ in pieces) == visible


def test_control_word_boundary_space_is_dropped():
    r"""`\times 10` 那个空格不能出现在任何一段文字里。"""
    pieces = M._latex_to_pieces(r"K=1.0\times 10^{2}")
    joined = "".join(t for t, _ in pieces)
    assert "× 10" not in joined
    assert "\u00a0" not in joined          # 也不该退化成 NBSP
    assert joined == "K=1.0×102"


def test_explicit_space_is_not_dropped():
    r"""`\,` 不能被吃成空串 —— 那是作者写的窄空格。"""
    pieces = M._latex_to_pieces(r"a=9.8\,\mathrm{m/s^{2}}")
    base = "".join(t for t, v in pieces if v is None)
    assert "\u00a0" in base, repr(base)
    assert "9.8m/s" not in base, "窄空格被吞掉了"


def test_group_command_boundary_space_is_dropped():
    r"""`\mathrm{...}` 后面那个空格同样是定界符。"""
    pieces = M._latex_to_pieces(r"\mathrm{kg/m^{3}} 10")
    joined = "".join(t for t, _ in pieces)
    assert "kg/m" in joined
    assert " 10" not in joined, repr(joined)
    assert joined == "kg/m310"


# ---------------------------------------------------------------- 开关回退
def test_switch_off_restores_omml_only(monkeypatch):
    r"""`MIXED_SUPERSCRIPT = False` 时 `_om_or_scripts` 必须返回 None（走 OMML）。"""
    monkeypatch.setattr(M, "MIXED_SUPERSCRIPT", False)
    stats = {}
    assert M._om_or_scripts(r"a^{2}", None, stats) is None
    assert stats.get("text_scripts", 0) == 0


def test_switch_on_counts_conversions():
    stats = {}
    runs = M._om_or_scripts(r"a^{2}", None, stats)
    assert runs is not None
    assert stats["text_scripts"] == 1
    assert stats["text_script_rows"] == [r"a^{2}"]


# ---------------------------------------------------------------- 端到端
def _build(tmp_path, paragraphs):
    """用 python-docx 造一份最小 docx，再跑 mathify 全文。"""
    from docx import Document
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    src = tmp_path / "in.docx"
    doc.save(str(src))
    out = tmp_path / "out.docx"
    M.mathify(src, out, None, None)
    return out


def test_end_to_end_superscript_becomes_a_real_script(tmp_path):
    r"""`10^-7` 走完全程后：文本是 `10` + `-7`(上标)，且**不产生 OMML**。"""
    import zipfile
    out = _build(tmp_path, ["c=3×10^-7"])
    xml = zipfile.ZipFile(str(out)).read("word/document.xml").decode("utf-8")
    assert "<m:oMath" not in xml, "纯上下标不该再产 OMML"
    assert 'w:vertAlign w:val="superscript"' in xml
    assert "-7" in xml


@needs_omml
def test_end_to_end_fraction_still_produces_omml(tmp_path):
    r"""真分式必须仍然走 OMML —— 混合策略不能把 `\frac` 吃掉。"""
    import zipfile
    out = _build(tmp_path, ["x^2/16+y^2/9=1"])
    xml = zipfile.ZipFile(str(out)).read("word/document.xml").decode("utf-8")
    assert "<m:oMath" in xml, "分式必须还是 OMML"
    assert "<m:f>" in xml, "分式结构丢了"


@needs_omml
def test_end_to_end_mixed_document_has_both(tmp_path):
    r"""一份文档里两种形态并存：这是「混合策略」的本体。"""
    import zipfile
    out = _build(tmp_path, ["a_n=2a_{n-1}+3", "x^2/16+y^2/9=1"])
    xml = zipfile.ZipFile(str(out)).read("word/document.xml").decode("utf-8")
    assert "<m:oMath" in xml
    assert 'w:vertAlign w:val="subscript"' in xml
