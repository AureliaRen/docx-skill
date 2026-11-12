# -*- coding: utf-8 -*-
"""Regression suite for the formula chain.

Locks down three behaviours so future rule changes cannot silently break them:

  A. linear notation -> LaTeX (54 cases, including the bugs found while
     pressure-testing against physics / calculus / linear-algebra material)
  B. merged sub/superscript repair (6 cases)
  C. span detection — what must NOT be treated as a formula, and what must

The OMML tests additionally need ``latex2mathml`` plus Microsoft's
``MML2OMML.XSL`` from a local Office install. When either is missing they skip
instead of failing, so a bare CI runner still exercises everything else.

P3-3/T1a merge note — the local branch contributed four extra sections that
this file had no coverage for (marked ``MERGED-LOCAL`` below):
the phasor-dot patch (``m:limUpp`` -> ``m:acc``), ``assert_legal`` negative
cases, the ``docx_mathify`` end-to-end round trip, and a read-only replay of the
``_smoke/out_tnr.docx`` baseline. The single overlapping case
(``test_math_font_defaults_to_times_new_roman``) was dropped in favour of the
equivalent ``test_formula_font_defaults_to_times_new_roman`` above.

Run:  python -m pytest tests/ -v
"""
from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402
import math_omml  # noqa: E402
import fix_phasor_dot  # noqa: E402

# `_smoke` 基线在合并仓的**上一层**（docx-skill-build\_smoke\），不是
# docx-merged\_smoke\。本地版原按 `SKILL.parent / "_smoke"` 算（SKILL 是本地
# docx\ 根目录），直接搬过来会指到不存在的目录 → 从 parents[1] 再往上跳一级。
MERGED_ROOT = Path(__file__).resolve().parents[1]      # docx-skill-build\docx-merged
SMOKE = MERGED_ROOT.parent / "_smoke"                  # docx-skill-build\_smoke

# ---------------------------------------------------------------------------
# A. linear -> LaTeX
# ---------------------------------------------------------------------------
LATEX_CASES = [
    # original linear form,                expected LaTeX
    (r"a/b",                            r"\frac{a}{b}"),
    (r"6/(2+2)",                        r"\frac{6}{2+2}"),
    (r"10/10k",                         r"\frac{10}{10k}"),
    (r"I = U_{oc}/(R_{0} + R_{L})",     r"I = \frac{U_{oc}}{R_{0} + R_{L}}"),
    (r"U_{A} = (6/2 + 3)/(1/2 + 1/2)",  r"U_{A} = \frac{\frac{6}{2} + 3}{\frac{1}{2} + \frac{1}{2}}"),
    (r"U=IR",                           r"U=IR"),
    (r"ΣI=0",                           r"\sum I=0"),
    (r"f=50Hz",                         r"f=50\,\text{Hz}"),
    (r"i_{L}(0^{+}) = i_{L}(0^{-})",    r"i_{L}(0^{+}) = i_{L}(0^{-})"),
    # --- appended after pressure-testing against physics / calculus / linalg ---
    (r"lim_{x→0} (sinx)/x = 1",         r"\lim_{x\to 0} \frac{\sin x}{x} = 1"),
    (r"lim_{x→∞}(1 + 1/x)^{x} = e",     r"\lim_{x\to \infty}(1 + 1/x)^{x} = e"),
    (r"s = ∫_{a}^{b} √(1 + (y′)²) dx",  r"s = \int_{a}^{b} \sqrt{1 +(y')^{2}} dx"),
    (r"∫dx/(1 + x^{2}) = arctanx + C",
     r"\int \frac{dx}{1 + x^{2}} = \arctan x + C"),
    (r"tr(A) = λ_{1} + λ_{2}",          r"\mathrm{tr}(A) = \lambda_{1} + \lambda_{2}"),
    (r"diag(A, B)",                     r"\mathrm{diag}(A, B)"),
    (r"h = 6.63×10^{−34} J·s",          r"h = 6.63\times 10^{-34}\,\text{J·s}"),
    (r"R = 8.31 J/(mol·K)",             r"R = 8.31\,\text{J/(mol·K)}"),
    (r"A^{T}A = E",                     r"A^{T}A = E"),      # units must not eat the matrix name A
    (r"x^{2} = 4",                      r"x^{2} = 4"),
    # --- matrix / cases / multi-line environments: passed through verbatim ---
    (r"A = \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}",
     r"A = \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}"),
    (r"f(x) = \begin{cases} x^{2}, & x \ge 0 \\ -x, & x < 0 \end{cases}",
     r"f(x) = \begin{cases} x^{2}, & x \ge 0 \\ -x, & x < 0 \end{cases}"),
    (r"\begin{pmatrix} \frac{a}{b} & 0 \\ 0 & 1 \end{pmatrix}",
     r"\begin{pmatrix} \frac{a}{b} & 0 \\ 0 & 1 \end{pmatrix}"),
    # aligned/align without \tag -> starred form (latex2mathml inserts (1)(2) otherwise)
    (r"\begin{aligned} f &= x^{2} \\ &= y \end{aligned}",
     r"\begin{align*} f &= x^{2} \\ &= y \end{align*}"),
    (r"\begin{align} u &= v \\ m &= n \end{align}",
     r"\begin{align*} u &= v \\ m &= n \end{align*}"),
    # with \tag: left alone (docx_mathify handles numbering itself)
    (r"\begin{align} a &= b \tag{1} \\ c &= d \tag{2} \end{align}",
     r"\begin{align} a &= b \tag{1} \\ c &= d \tag{2} \end{align}"),
    (r"T = 2\pi\begin{vmatrix} a & b \\ c & d \end{vmatrix}",
     r"T = 2\pi\begin{vmatrix} a & b \\ c & d \end{vmatrix}"),
    # --- hand-written matrix text [a, b; c, d]  (';' rows, ',' columns) ---
    (r"[a, b; c, d]",                   r"\begin{bmatrix}a & b \\ c & d\end{bmatrix}"),
    (r"(1, 2; 3, 4)",                   r"\begin{pmatrix}1 & 2 \\ 3 & 4\end{pmatrix}"),
    (r"K = [0, -jωL; jωL, 0]",          r"K = \begin{bmatrix}0 & -j\omega L \\ j\omega L & 0\end{bmatrix}"),
    (r"[R₁, 0; 0, R₂]",                 r"\begin{bmatrix}R_{1} & 0 \\ 0 & R_{2}\end{bmatrix}"),
    (r"[1, 0, 0; 0, 1, 0; 0, 0, 1]",
     r"\begin{bmatrix}1 & 0 & 0 \\ 0 & 1 & 0 \\ 0 & 0 & 1\end{bmatrix}"),
    # --- 从本地 P2 版并入（T1b）：按输入串去重后本地独有的 23 条 ---
    (r"U_{n} = Σ(E/R) / Σ(1/R)",        r"U_{n} = \frac{\sum(\frac{E}{R})}{\sum(\frac{1}{R})}"),
    (r"√2",                             r"\sqrt{2}"),
    (r"√(LC)",                          r"\sqrt{LC}"),
    (r"f_{0} = 1/(2π√(LC))",            r"f_{0} = \frac{1}{2\pi \sqrt{LC}}"),
    (r"ω≈314 rad/s",                    r"\omega \approx 314\,\text{rad/s}"),
    (r"1.5A",                           r"1.5\,\text{A}"),
    (r"1μF=10^{-6}F",                   r"1\,\text{μF}=10^{-6}\,\text{F}"),
    (r"j^{2}=-1",                       r"j^{2}=-1"),
    (r"1/j=−j",                         r"\frac{1}{j}=-j"),
    (r"U_{L} = √3·U_{P}",               r"U_{L} = \sqrt{3}\cdot U_{P}"),
    (r"sin(ωt+ψ_{u})",                  r"\sin(\omega t+\psi_{u})"),
    (r"P(tanφ_{1} − tanφ_{2})/(ωU^{2})",
     r"\frac{P(\tan\varphi_{1} - \tan\varphi_{2})}{\omega U^{2}}"),
    (r"F = dp/dt",                      r"F = \frac{dp}{dt}"),
    (r"F = Gm_{1}m_{2}/r^{2}",          r"F = \frac{Gm_{1}m_{2}}{r^{2}}"),
    (r"dz = (∂z/∂x)dx + (∂z/∂y)dy",
     r"dz =(\frac{\partial z}{\partial x})dx +(\frac{\partial z}{\partial y})dy"),
    # --- T1b 漂移修正：远程量纲/单位子系统会把 ``N·m^{2}/kg^{2}`` 整串识别为
    #     单位串并整体 ``\mathrm`` 化，不再是 ``\frac{}{}`` 分式。
    #     按 D-011 接受远程输出，改期望值以跟随主干。 ---
    (r"G = 6.67×10^{−11} N·m^{2}/kg^{2}",
     r"G = 6.67\times 10^{-11}\,\mathrm{N\cdot m^{2}/kg^{2}}"),
    (r"v = √(GM/R) ≈ 7.9 km/s",
     r"v = \sqrt{\frac{GM}{R}} \approx 7.9\,\text{km/s}"),
    (r"c = 3.0×10^{8} m/s",             r"c = 3.0\times 10^{8}\,\text{m/s}"),
    (r"U̇ = U∠ψ_{u}",                   r"\dot{U} = U\angle \psi_{u}"),
    (r"j=1∠90°",                        r"j=1\angle 90^{\circ}"),
    (r"P∝I^{2}",                        r"P\propto I^{2}"),
    (r"τ = L/R_{eq}",                   r"\tau = \frac{L}{R_{eq}}"),
    (r"S^{2} = P^{2} + Q^{2}",          r"S^{2} = P^{2} + Q^{2}"),
]

# ---------------------------------------------------------------------------
# A2. merged sub/superscript repair
# ---------------------------------------------------------------------------
REPAIR_CASES = [
    (r"I^{2R}",      r"I^{2}R"),
    (r"R_{1R2}",     r"R_{1}R_{2}"),
    (r"\omega_{0L}", r"\omega_{0}L"),
    (r"\omega_{0CR}", r"\omega_{0}CR"),
    (r"U^{2}",       r"U^{2}"),        # must not change
    (r"10^{-6}",     r"10^{-6}"),      # must not change
]

# ---------------------------------------------------------------------------
# C. span detection
# ---------------------------------------------------------------------------
NO_SPAN = [
    "范围 = 前 5 周",        # CJK breaks it; no right-hand operand
    "b=", "n=", "=R",        # fragments
    "6V", "2Ω",              # bare quantities
    "2026-11-08",            # date
    "电工电子技术 · 电路分析部分",
    "[0, 1]",                # interval: single line, no ';', not a matrix
    "f(x, y)",               # function call: not a matrix
]
YES_SPAN = [
    "U=IR", "ΣI=0", "f=50Hz", "i_{L}(0^{+}) = i_{L}(0^{-})",
    "A = [a, b; c, d]",
]


def atoms_of(text: str) -> list[dict]:
    """Turn plain text into the atom list find_math_spans expects."""
    return [{"ch": c, "rpr": None, "script": None} for c in text]


# OMML 可用性判定收在 `tests/_omml.py` 一处（REVIEW-CODE-R4 应修 5）。
from _omml import needs_omml  # noqa: E402


# ---------------------------------------------------------------------------
# A. string level
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("linear,expected", LATEX_CASES)
def test_linear_to_latex(linear, expected):
    assert L.linear_to_latex(linear) == expected


@pytest.mark.parametrize("source,expected", REPAIR_CASES)
def test_repair_merged_scripts(source, expected):
    got, _fixes = L.repair_merged_scripts(source)
    assert got == expected


# ---------------------------------------------------------------------------
# C. span detection
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("text", NO_SPAN)
def test_text_is_not_treated_as_formula(text):
    assert not L.find_math_spans(atoms_of(text))


@pytest.mark.parametrize("text", YES_SPAN)
def test_text_is_treated_as_formula(text):
    assert L.find_math_spans(atoms_of(text))


# ---------------------------------------------------------------------------
# B. OMML legality
# ---------------------------------------------------------------------------
@needs_omml
@pytest.mark.parametrize("linear,_expected", LATEX_CASES)
def test_latex_converts_to_legal_omml(linear, _expected):
    latex = L.linear_to_latex(linear)
    # latex_to_omml_lxml runs assert_legal() internally and raises on illegal
    # structure such as a nested <m:oMath> or an <m:rad> missing <m:e>.
    omml = math_omml.latex_to_omml_lxml(latex)
    assert omml is not None
    assert omml.tag.rsplit("}", 1)[-1] in ("oMath", "oMathPara")


@needs_omml
@pytest.mark.parametrize("tex", [
    r"\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
    r"\sum_{i=1}^{n} x_i",
    r"\int_{0}^{\infty} e^{-x} dx",
    r"A = \begin{bmatrix}1 & 2\\3 & 4\end{bmatrix}",
    r"\text{中文}",
])
def test_complex_formula_shapes(tex):
    assert math_omml.latex_to_omml_lxml(tex) is not None


# ---------------------------------------------------------------------------
# Font injection
# ---------------------------------------------------------------------------
@needs_omml
def test_formula_font_defaults_to_times_new_roman():
    M_NS, W_NS = math_omml.M_NS, math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("U=IR")
    runs = omml.findall(f".//{{{M_NS}}}r")
    assert runs, "conversion produced no runs"

    for run in runs:
        rfonts = run.find(f"{{{W_NS}}}rPr/{{{W_NS}}}rFonts")
        assert rfonts is not None, "run missing w:rPr/w:rFonts"
        assert rfonts.get(f"{{{W_NS}}}ascii") == "Times New Roman"

    # CT_R content-model order: w:rPr must come before m:t
    for run in runs:
        assert [c.tag.rsplit("}", 1)[-1] for c in run][-1] == "t"


@needs_omml
def test_formula_font_env_override(monkeypatch):
    monkeypatch.setenv("MATH_FONT", "Cambria Math")
    W_NS = math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("x")
    rfonts = omml.find(f".//{{{W_NS}}}rFonts")
    assert rfonts is not None
    assert rfonts.get(f"{{{W_NS}}}ascii") == "Cambria Math"


@needs_omml
def test_formula_font_can_be_disabled(monkeypatch):
    monkeypatch.setenv("MATH_FONT", "")
    W_NS = math_omml.W_NS
    omml = math_omml.latex_to_omml_lxml("x")
    assert omml.find(f".//{{{W_NS}}}rFonts") is None


# ===========================================================================
# MERGED-LOCAL (§T1a): the four sections the local branch owned exclusively.
#
# namespace helpers are derived from `math_omml` itself so they can never drift
# from the script side; the OMML gate is `needs_omml` from `tests/_omml.py`
# (the single implementation) — no second copy of the availability check here.
# ===========================================================================
M = "{" + math_omml.M_NS + "}"
W = "{" + math_omml.W_NS + "}"
W_NS = math_omml.W_NS


# --------------------------------------------------------------- helpers
def run_cli(script: str, *args, env_extra: dict | None = None) -> subprocess.CompletedProcess:
    """以子进程跑 argparse CLI；环境对齐 hermes e2e 的口径。

    ``env_extra`` 用于单条用例改写环境（如把 ``MML2OMML_XSL`` 指到不存在的
    路径来复现「缺 XSL」）；传 ``None`` 表示删掉该变量，避免宿主机上的残留
    值影响判定。不改写时行为与旧签名完全一致。
    """
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    env.pop("MML2OMML_XSL", None)
    for k, v in (env_extra or {}).items():
        if v is None:
            env.pop(k, None)
        else:
            env[k] = v
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        capture_output=True, env=env, check=False,
    )


def doc_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def m_count(el, name: str) -> int:
    return len(list(el.iter(M + name)))


def plain_text(path: Path) -> str:
    """document.xml 里所有 ``w:t`` 的文本拼接（用于确认没有残留源写法）。"""
    from lxml import etree

    root = etree.fromstring(doc_xml(path).encode("utf-8"))
    return "".join(t.text or "" for t in root.iter("{%s}t" % W_NS))


# ---------------------------------------------------------------- 1. 相量点
@needs_omml
def test_dot_single_becomes_accent():
    """``\\dot{I}``：limUpp 1 → 0，acc 0 → ≥1，函数返回修了 1 处。"""
    el = math_omml.latex_to_omml_lxml(r"\dot{I}")

    # 补丁前：latex2mathml 产出的是 limUpp（点浮在字上方）
    assert m_count(el, "limUpp") == 1, "前置假设失效：latex2mathml 未产出 limUpp"
    assert m_count(el, "acc") == 0

    n = fix_phasor_dot.fix_phasor_dots(el)

    assert n == 1
    assert m_count(el, "limUpp") == 0, "相量点没改干净，Word 里会渲染成浮大的点"
    assert m_count(el, "acc") >= 1


@needs_omml
def test_dot_compound_two_occurrences():
    """``\\dot{I}_{L}+\\dot{I}_{C}``：两处点都要修，返回 2。"""
    el = math_omml.latex_to_omml_lxml(r"\dot{I}_{L}+\dot{I}_{C}")
    assert m_count(el, "limUpp") == 2

    n = fix_phasor_dot.fix_phasor_dots(el)

    assert n == 2
    assert m_count(el, "limUpp") == 0
    assert m_count(el, "acc") == 2


@needs_omml
def test_dot_keeps_run_font_properties():
    """acc 的 m:e 必须继承原 run 的 w:rPr——否则字体注入在这一处丢字。

    ``math_omml.apply_math_font`` 给每个 ``m:r`` 注了 Times New Roman；
    ``fix_phasor_dots`` 搬基的 run 时必须 deepcopy，不能重建裸 run。
    """
    el = math_omml.latex_to_omml_lxml(r"\dot{I}")
    fix_phasor_dot.fix_phasor_dots(el)

    acc = next(el.iter(M + "acc"))
    e = acc.find(M + "e")
    assert e is not None

    rfonts = [rf for rf in e.iter(W + "rFonts")]
    assert rfonts, "m:acc/m:e 里的 run 丢了 w:rPr/rFonts（字体注入在此处断链）"
    assert rfonts[0].get(W + "ascii") == "Times New Roman"


@needs_omml
def test_non_phasor_lim_is_left_alone():
    """长基（多字符）的 limUpp 是真正的``\\lim``上限，必须原样放行。

    这是补丁的取舍边界：只认「单拉丁字母基 + 单点上限」，改多了会把极限式改坏。
    """
    from lxml import etree

    el = etree.Element(M + "oMath")
    lim = etree.SubElement(el, M + "limUpp")
    base = etree.SubElement(lim, M + "e")
    r = etree.SubElement(base, M + "r")
    t = etree.SubElement(r, M + "t")
    t.text = "AB"                      # 多字符基
    lim_el = etree.SubElement(lim, M + "lim")
    r2 = etree.SubElement(lim_el, M + "r")
    t2 = etree.SubElement(r2, M + "t")
    t2.text = "\u0307"

    assert fix_phasor_dot.fix_phasor_dots(el) == 0
    assert m_count(el, "limUpp") == 1, "长基 limUpp 被误改"


# ---------------------------------------------------------------- 2. assert_legal
@needs_omml
def test_assert_legal_rejects_nested_omath_in_num():
    """``m:num`` 内嵌 ``<m:oMath>`` → 必须被拒。"""
    from lxml import etree

    el = etree.Element(M + "oMath")
    frac = etree.SubElement(el, M + "f")
    num = etree.SubElement(frac, M + "num")
    etree.SubElement(num, M + "oMath")          # 非法：内容模型不接受嵌套 oMath

    with pytest.raises(RuntimeError) as ei:
        math_omml.assert_legal(el, r"\frac{<bad>}{x}")
    assert "oMath" in str(ei.value)


@needs_omml
def test_assert_legal_rejects_rad_without_e():
    """``m:rad`` 缺 ``m:e`` → 必须被拒。"""
    from lxml import etree

    el = etree.Element(M + "oMath")
    rad = etree.SubElement(el, M + "rad")
    etree.SubElement(rad, M + "deg")            # 只有 deg，没有 e

    with pytest.raises(RuntimeError) as ei:
        math_omml.assert_legal(el, r"\sqrt{<bad>}")
    assert "rad" in str(ei.value)


@needs_omml
def test_assert_legal_accepts_valid_omml():
    """真实转换产物必须过自检（正向用例，防止守卫过严）。"""
    el = math_omml.latex_to_omml_lxml(r"U_R=I_R \times R_1")
    math_omml.assert_legal(el, r"U_R=I_R \times R_1")   # 不抛即通过


# ---------------------------------------------------------------- 3. 端到端
@needs_omml
def test_mathify_end_to_end(tmp_path: Path):
    r"""含**真分式** ``$...$`` 的 docx → docx_mathify → 零失败 + 有 oMath + validate 过。

    样例为什么必须用 ``\frac``：合并仓的 ``docx_mathify`` 按 ``MIXED_SUPERSCRIPT``
    混合策略分流，**纯上下标**片段（如 ``$U_R=I_R \times R_1$`` / ``$P = I^{2}R$``）
    在 ``_om_or_scripts()`` 第二阶就被判成「需要角标、但不需要二维排版」，
    直接落成 ``w:rPr/w:vertAlign`` 文字角标，产出 **0 个 ``<m:oMath>``**。
    这是既定行为（不是缺陷）：``m:sSup`` 排出来的小号字在 Word 里再按一次上标
    只会在小号字上叠一层，真角标必须是 run 级 vertAlign。

    所以**验 OMML 必须用真分式**——``\frac`` 属于 ``_STRUCTURAL_CMDS``，有二维
    排版，文字 run 表达不了，一定走 OMML 分支。
    """
    from docx import Document

    src = tmp_path / "in.docx"
    out = tmp_path / "out.docx"
    rep = tmp_path / "report.md"

    doc = Document()
    doc.add_paragraph("分压 $U=\\frac{U_0}{R_1}$ 结束")
    doc.add_paragraph("功率 $P = \\frac{U^{2}}{R}$")
    doc.save(str(src))

    proc = run_cli("docx_mathify.py", src, "-o", out, "--report", rep)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    assert out.is_file()

    # a) 失败数 0（报告里的「失败：0」是脚本自述，下面还有 XML 级证据）
    assert rep.is_file()
    text = rep.read_text(encoding="utf-8")
    assert "失败：0" in text, f"转换有失败项：\n{text}"

    # b) document.xml 里真出现 <m:oMath，且真是分式结构（不是被降级成文字）
    raw = doc_xml(out)
    assert raw.count("<m:oMath") >= 1, "真分式片段至少该产出一个 oMath"
    assert "<m:f>" in raw or "<m:f " in raw, "\\frac 没落成分式结构 <m:f>"

    # c) XML 级零失败证据：源写法不该残留在正文文本里
    #    注意 `$` 本身会留下——docx_mathify 用的是 find_math_spans 启发式，
    #    不剥 `$...$` 标记层（同一条段落里 `$` 留在公式两侧的 w:t 里）。
    #    这不是 bug，但必须锁死，否则以后误以为标记被消费掉了。
    body_text = plain_text(out)
    assert "\\frac" not in body_text, "源写法 \\frac 残留在正文，说明回退成了纯文本"
    assert "U_0" not in body_text, "下划线源写法残留，说明该片段没被转成公式"
    assert body_text.count("$") == 4, (
        f"两处 $...$ 应各留 2 个 $ 标记（当前 {body_text.count('$')} 个）：{body_text!r}")

    # d) 公式内容真进了 m:t（不是只把 $ 删掉）
    from lxml import etree

    mtext = "".join(t.text or "" for t in
                    etree.fromstring(raw.encode("utf-8")).iter(M + "t"))
    assert "U" in mtext and "0" in mtext, f"分子/分母没进 OMML：{mtext!r}"

    # e) 包健康 + python-docx 能开
    proc = run_cli("docx_validate.py", out)
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")
    assert '"ok": true' in proc.stdout.decode("utf-8").replace("'", '"')

    from docx import Document as D2

    D2(str(out))


@needs_omml
def test_mathify_leaves_plain_text_alone(tmp_path: Path):
    """没有公式的段落必须原样保留（防止改造过度）。"""
    from docx import Document

    src = tmp_path / "plain.docx"
    out = tmp_path / "plain_out.docx"

    doc = Document()
    doc.add_paragraph("纯文本段落，没有公式。")
    doc.save(str(src))

    proc = run_cli("docx_mathify.py", src, "-o", out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    assert "<m:oMath" not in doc_xml(out)
    assert "纯文本段落" in plain_text(out)


# ---------------------------------------------------------------- 4. 基线复现
@pytest.mark.skipif(
    not (SMOKE / "out_tnr.docx").is_file(),
    reason="合并仓上层没有 _smoke 基线样本（只读复现用，缺失不算失败）",
)
def test_smoke_baseline_out_tnr():
    """``_smoke/out_tnr.docx`` 基线复现：4 个公式、零 limUpp 残留、字体全 TNR。

    **_smoke 是只读回归基线**：本用例只读不写。它是「上一轮真机产出的样本」，
    用来兜住「整套链路在本机能跑出与当时一致的结果」这件事——不替代上面那些
    现场生成 fixture 的用例。
    """
    p = SMOKE / "out_tnr.docx"

    with zipfile.ZipFile(p) as z:
        raw = z.read("word/document.xml").decode("utf-8")
        names = z.namelist()

    assert raw.count("<m:oMath") == 4, "基线样本的公式数变了"
    assert "<m:limUpp" not in raw, "相量点没修干净（基线里不该有 limUpp）"

    from lxml import etree

    root = etree.fromstring(raw.encode("utf-8"))
    fonts = {rf.get(W + "ascii") for rf in root.iter(W + "rFonts")}
    assert fonts == {"Times New Roman"}, f"基线字体不是纯 TNR：{fonts}"

    # 基线包必须健康（顺带确认它没被别的流程改坏）
    proc = run_cli("docx_validate.py", p)
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")

    assert "word/document.xml" in names


# ===========================================================================
# 5. CLI 契约（P6-FIX-C5F）：落点标志互斥 + 退出码 3 真的可达
#
# 这一组锁的是**命令行行为**，不是公式内容。审计口径（CLI-AUDIT P2）：
# `docx_mathify.py in.docx --inplace -o out.docx` 曾静默忽略 `-o` —— out.docx
# 不生成、输入文件反被覆盖，退出码还是 0。所以三条断言缺一不可：退出码、
# 「两边文件都没被改」、以及备份文件没被创建（证明连备份分支都没进）。
#
# 三条落点用例刻意用**可降级片段**（`$P = I^{2}R$` 落成文字角标 run），
# 全程不碰 MML2OMML.XSL —— 这样它们在无 Office 的 CI 上也照跑，不 skip。
# 退出码 3 的用例反过来：用真分式 + 把 XSL 指到不存在的路径，任何机器上都
# 必然失败，同样不依赖宿主机装没装 Office。
# ===========================================================================
def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _degradable_docx(tmp_path: Path, name: str = "in.docx") -> Path:
    """纯上下标片段：落成文字 run + vertAlign，不需要 XSL 也不产 OMML。"""
    from docx import Document

    src = tmp_path / name
    doc = Document()
    doc.add_paragraph(r"功率 $P = I^{2}R$ 结束")
    doc.save(str(src))
    return src


def _structural_docx(tmp_path: Path, name: str = "in.docx") -> Path:
    r"""含真分式：必走 OMML 分支（``\frac`` 在 ``_STRUCTURAL_CMDS`` 里）。"""
    from docx import Document

    src = tmp_path / name
    doc = Document()
    doc.add_paragraph(r"分压 $U=\frac{U_0}{R_1}$ 结束")
    doc.save(str(src))
    return src


def test_mathify_rejects_inplace_with_out(tmp_path: Path):
    """``--inplace`` + ``-o`` 是用法错误：退 2，且**两个文件都不许被碰**。

    反向自证（做完已还原）：把互斥检查换成旧的分流写法，本用例会在
    ``returncode == 2`` 那一行红 —— 旧版退 0、输入文件被覆盖。
    """
    src = _degradable_docx(tmp_path)
    out = tmp_path / "out.docx"
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--inplace", "-o", out)

    assert proc.returncode == 2, (
        f"退出码应为 2（用法错误），实际 {proc.returncode}；"
        f"stdout={proc.stdout.decode('utf-8', 'replace')!r}")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "--inplace" in stderr and "-o" in stderr, (
        f"报错没说清是哪两个标志冲突：{stderr!r}")
    assert "error:" in stderr, f"用法错误必须走 stderr 且以 error: 开头：{stderr!r}"

    # a) 源文件字节不变（关键：旧版会就地覆盖它）
    assert _sha256(src) == before, "报错路径上源文件被改了 —— 判定必须发生在任何读写之前"
    # b) 没有凭空写出 -o 指定的文件
    assert not out.exists(), "-o 指定的文件被创建了，说明并没有真的报错退出"
    # c) 备份分支也没进（进了就说明它还是照着 --inplace 跑了一段）
    assert not (tmp_path / "in.bak.docx").exists(), "备份文件被创建，说明 --inplace 已开始执行"


def test_mathify_out_only_still_writes_out(tmp_path: Path):
    """单用 ``-o``：退 0，写出 out.docx，源文件一个字节都不动。"""
    src = _degradable_docx(tmp_path)
    out = tmp_path / "out.docx"
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "-o", out)

    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    assert out.is_file(), "单用 -o 必须写出输出文件"
    assert _sha256(src) == before, "单用 -o 不该改源文件"
    assert not (tmp_path / "in.bak.docx").exists(), "单用 -o 不该产生备份"


def test_mathify_inplace_only_still_rewrites(tmp_path: Path):
    """单用 ``--inplace``：退 0，先备份再就地改写，备份 == 原始字节。"""
    src = _degradable_docx(tmp_path)
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--inplace")

    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    bak = tmp_path / "in.bak.docx"
    assert bak.is_file(), "--inplace 必须先备份"
    assert _sha256(bak) == before, "备份内容与原文件不一致"
    assert _sha256(src) != before, "--inplace 没有真的改写源文件（转换没生效？）"
    # 角标真的落成了 run 级 vertAlign，而不是被整段跳过
    assert "vertAlign" in doc_xml(src), "纯上下标片段没落成文字角标"


def test_mathify_rejects_inplace_with_dry_run(tmp_path: Path):
    """``--inplace`` + ``--dry-run`` 是用法错误：退 2，且**一个字节都不许写**。

    P8-D3：两个标志的语义正好相反（``--inplace`` = 要写盘（原地）、
    ``--dry-run`` = 不写盘）。旧版 argparse 里它们是两个并列独立 flag，
    同传时 ``if args.inplace and not args.dry_run`` 静默择一 —— 表现为
    ``--dry-run`` 吃掉 ``--inplace``：不备份、不写盘、exit 0，用户无从判断
    到底按哪条跑了。收进 ``add_mutually_exclusive_group()`` 后由 argparse
    在解析阶段拒掉（与 ``docx_template.py`` 的 C5d、本脚本 C5f 同一口径）。

    本用例只碰 argparse 层（``_degradable_docx`` 是可降级片段，连 OMML 分支
    都不进），因此**不挂 ``needs_omml``**：无 XSL 的机器上也照样 pass。
    """
    src = _degradable_docx(tmp_path)
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--inplace", "--dry-run")

    assert proc.returncode == 2, (
        f"退出码应为 2（用法错误），实际 {proc.returncode}；"
        f"stdout={proc.stdout.decode('utf-8', 'replace')!r}")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "--inplace" in stderr and "--dry-run" in stderr, (
        f"报错没说清是哪两个标志冲突：{stderr!r}")
    assert "error:" in stderr, f"用法错误必须走 stderr 且以 error: 开头：{stderr!r}"

    # a) 源文件字节不变（旧版虽然也没写，但这条锁的是「判定先于任何读写」）
    assert _sha256(src) == before, "报错路径上源文件被改了 —— 判定必须发生在任何读写之前"
    # b) 备份文件没被创建（旧版 --dry-run 挡在前面，但组合一变就可能漏）
    assert not (tmp_path / "in.bak.docx").exists(), "备份文件被创建，说明已开始写盘"
    # c) 光有报错还不够：stdout 不许出现「输出:」那种「我写了」的自述
    assert "输出:" not in proc.stdout.decode("utf-8", "replace"), (
        "报错路径上还打印了「输出:」，与「零写盘」自相矛盾")


def test_mathify_dry_run_only_still_reports(tmp_path: Path):
    """单用 ``--dry-run``：退 0，出汇总报告，源文件不动、无备份。

    这是互斥组改动的**反向保护**：把 ``--inplace`` / ``--dry-run`` 收进同一个
    组时，最容易顺手把 ``--dry-run`` 单传也打成用法错误（比如误写成
    ``required=True`` 或 required 组）。这条钉住单传照旧。
    """
    src = _degradable_docx(tmp_path)
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--dry-run")

    assert proc.returncode == 0, (
        f"单用 --dry-run 应成功，实际 {proc.returncode}；"
        f"stderr={proc.stderr.decode('utf-8', 'replace')!r}")
    out_text = proc.stdout.decode("utf-8", "replace")
    assert "段落" in out_text and "失败" in out_text, f"stdout 汇总报告不见了：{out_text!r}"
    assert "输出:" not in out_text, "--dry-run 不写盘，不该打印「输出:」"
    assert _sha256(src) == before, "--dry-run 改了源文件"
    assert not (tmp_path / "in.bak.docx").exists(), "--dry-run 不该产生备份"


def test_mathify_exit_3_when_xsl_missing(tmp_path: Path):
    """缺 XSL → 文档仍写出，但退出码 3（部分成功），报告里失败数 > 0。

    这条用例是「退出码 3 不是死分支」的**可执行证据**：它把 ``MML2OMML_XSL``
    指到一个不存在的路径，构造出「环境不可用」这个真实触发条件。宿主机装没装
    Office 都一样 —— env 覆盖优先生效，挂了也走同一条失败路径。

    反向自证（做完已还原）：把 ``return 3`` 改成 ``return 0``，本用例在
    ``returncode == 3`` 处红。
    """
    src = _structural_docx(tmp_path)
    out = tmp_path / "out.docx"
    rep = tmp_path / "report.md"
    missing = tmp_path / "no_such_dir" / "MML2OMML.XSL"

    proc = run_cli("docx_mathify.py", src, "-o", out, "--report", rep,
                   env_extra={"MML2OMML_XSL": str(missing)})

    assert proc.returncode == 3, (
        f"缺 XSL 时退出码应为 3（部分成功），实际 {proc.returncode}；"
        f"stderr={proc.stderr.decode('utf-8', 'replace')!r}")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "failed to convert" in stderr, f"stderr 没报出失败数：{stderr!r}"

    # 文档照样写出去了（半成品保留原文，人工可补救）—— 这正是「部分成功」的含义
    assert out.is_file(), "退 3 时输出文件应当仍然存在"
    assert "\\frac" in plain_text(out), "失败片段应按原文留住，而不是被吃掉"

    # 报告独立佐证：失败数不是 0
    assert rep.is_file()
    text = rep.read_text(encoding="utf-8")
    assert "失败：0" not in text, "报告里失败数为 0，与退出码 3 自相矛盾"


def test_mathify_rejects_out_with_dry_run(tmp_path: Path):
    """``-o X`` + ``--dry-run`` 是用法错误：退 2，零写盘，且有产物 + 非零码不成立。

    P8-D6/B。旧版：``-o`` 与 ``--dry-run`` 同传时被静默择一 —— 不写盘、exit 0，
    stdout 却打出 ``输出: <X>``。那句回显是**假的**：文件根本没生成，用户读
    汇总行会以为 out.docx 已经好了。``--dry-run``（只出报告）与 ``-o``（要求
    写到那个路径）语义冲突，和 C5f 的 ``--inplace`` × ``-o`` 同一类，所以同样
    在 argparse 之后、任何读写之前显式拒（**不进互斥组**，否则会连带改掉 C5f
    的中文文案）。

    本用例只碰前置检查层（``_degradable_docx`` 连 OMML 分支都不进），所以
    **不挂 ``needs_omml``**：无 XSL 的机器上也必须 pass，不是 skip。
    """
    src = _degradable_docx(tmp_path)
    out = tmp_path / "out.docx"
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "-o", out, "--dry-run")

    assert proc.returncode == 2, (
        f"退出码应为 2（用法错误），实际 {proc.returncode}；"
        f"stdout={proc.stdout.decode('utf-8', 'replace')!r}")
    stdout = proc.stdout.decode("utf-8", "replace")
    stderr = proc.stderr.decode("utf-8", "replace")

    # a) 单行 error: 前缀
    err_lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(err_lines) == 1, f"用法错误应当只有一行，实际 {err_lines!r}"
    assert err_lines[0].startswith("error:"), f"必须以 error: 开头：{err_lines[0]!r}"
    assert "--dry-run" in err_lines[0] and ("-o" in err_lines[0] or "--out" in err_lines[0]), \
        f"报错没说清是哪两个标志冲突：{err_lines[0]!r}"

    # b) 零写盘：目标文件没被创建，源文件字节不变
    assert not out.exists(), "-o 指定的文件被创建了，与「--dry-run 不写盘」矛盾"
    assert _sha256(src) == before, "报错路径上源文件被改了 —— 判定必须发生在任何读写之前"
    assert not (tmp_path / "in.bak.docx").exists(), "备份文件被创建，说明已开始写盘"

    # c) 假回显必须消失：stdout 一行都不许有
    assert stdout == "", f"报错路径上 stdout 应为空（那句「输出:」是假的），实际 {stdout!r}"


def test_mathify_dry_run_only_stdout_unchanged(tmp_path: Path):
    """单用 ``--dry-run``：退 0，stdout 有汇总报告，**不含**「输出:」行。

    这是显式检查的反向保护：加检查时最容易顺手把 ``--dry-run`` 单传也拒了。
    """
    src = _degradable_docx(tmp_path)
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--dry-run")

    assert proc.returncode == 0, (
        f"单用 --dry-run 应成功，实际 {proc.returncode}；"
        f"stderr={proc.stderr.decode('utf-8', 'replace')!r}")
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "段落" in stdout and "失败" in stdout, f"stdout 汇总报告不见了：{stdout!r}"
    assert "输出:" not in stdout, "--dry-run 不写盘，不该打印「输出:」"
    assert _sha256(src) == before
    assert not (tmp_path / "in.bak.docx").exists()


def test_mathify_out_with_dry_run_keeps_c5f_wording_intact(tmp_path: Path):
    """C5f 回归保护：``-o`` × ``--inplace`` 的文案与退出码一个字都不许变。

    加了 ``-o`` × ``--dry-run`` 的新检查之后，这条锁住另一条组合没被动过。
    """
    src = _degradable_docx(tmp_path)
    out = tmp_path / "out.docx"
    before = _sha256(src)

    proc = run_cli("docx_mathify.py", src, "--inplace", "-o", out)

    assert proc.returncode == 2
    stderr = proc.stderr.decode("utf-8", "replace")
    expected = ("error: --inplace 与 -o/--out 互斥：--inplace 就地覆盖输入文件，"
                "-o/--out 写另一份输出文件，同时给无法判定落点；请只保留一个。")
    assert stderr.strip() == expected, f"C5f 文案被改动了：{stderr!r}"
    assert proc.stdout.decode("utf-8", "replace") == "", "用法错误路径 stdout 应为空"
    assert not out.exists()
    assert _sha256(src) == before
    assert not (tmp_path / "in.bak.docx").exists()


# ===========================================================================
# 输入错误契约（P9-C）
#
# ⚠️ 本节**刻意不挂 ``@needs_omml``**：这些用例只在 CLI 入口判定「这个路径是不是
#    一个能打开的 .docx」，全程不碰 ``latex2mathml`` / ``MML2OMML.XSL`` / Office。
#    给它们挂 ``needs_omml`` 会让 CI（无 Office）把**唯一能真跑**的输入校验整段
#    skip 掉，只在开发机上跑——正好反了。实测口径：把 ``MML2OMML_XSL`` 指向不存在的路径后
#    本节仍然 pass。
#
# 旧行为（实测 2026-11-11）：
#   * 缺文件 / 目录  → exit 1，stderr **1 行**（``error: <p> 不存在``，但语义错：
#     缺文件是「你给的输入不对」，不是「这个文档处理不了」）；
#   * 非 zip        → exit 1，**22 行 traceback**；
#   * 非 zip + ``--inplace`` → **先** ``shutil.copy2`` 写出 ``<名>.bak.docx``，
#     **再**在 ``Document()`` 上崩 —— 输入是垃圾却留下了一份垃圾备份，
#     stdout 还回了 93 字节的「备份: …」。
# ===========================================================================
def _nonzip(tmp_path: Path, name: str = "renamed.docx") -> Path:
    p = tmp_path / name
    p.write_text("i am not a zip package", encoding="utf-8")
    return p


def one_error_line(stderr: str) -> str:
    """stderr 必须**恰好一行**（``error: ...``），没有 traceback 前导。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def test_mathify_missing_input_is_a_usage_error(tmp_path: Path):
    """路径不存在 → exit 2 + 单行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such.docx"
    assert not missing.exists()

    proc = run_cli("docx_mathify.py", missing, "--dry-run")
    assert proc.returncode == 2, proc.stderr
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "Traceback" not in stderr
    line = one_error_line(stderr)
    assert line.startswith("error: ")
    assert "no such file" in line
    assert str(missing) in line
    assert proc.stdout == b""


def test_mathify_directory_is_a_usage_error(tmp_path: Path):
    """路径是目录 → exit 2 + 单行 ``not a file``。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()

    proc = run_cli("docx_mathify.py", adir, "--dry-run")
    assert proc.returncode == 2, proc.stderr
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "Traceback" not in stderr
    assert one_error_line(stderr) == f"error: not a file: {adir}"
    assert proc.stdout == b""


def test_mathify_non_zip_is_a_usage_error(tmp_path: Path):
    """存在的普通文件但不是 zip → exit 2 + 单行 ``not a ZIP package``，不再 22 行 traceback。"""
    bogus = _nonzip(tmp_path)

    proc = run_cli("docx_mathify.py", bogus, "--dry-run")
    assert proc.returncode == 2, proc.stderr
    stderr = proc.stderr.decode("utf-8", "replace")
    assert "Traceback" not in stderr
    assert one_error_line(stderr) == f"error: not a ZIP package: {bogus}"
    assert proc.stdout == b""


def test_mathify_non_zip_inplace_leaves_no_bak_behind(tmp_path: Path):
    """★ 非 zip + ``--inplace`` → **一个文件都不生成**（旧版留下 ``renamed.bak.docx``）。

    `--inplace` 的第一件事是 ``shutil.copy2(src, bak)``，校验若排在它后面，
    输入越烂、垃圾备份越确定会被写出来。这条同时钉三件事：
    rc=2、stdout 空、目录里没有 ``.bak.docx``。
    """
    bogus = _nonzip(tmp_path)
    before = {p.name for p in tmp_path.iterdir()}

    proc = run_cli("docx_mathify.py", bogus, "--inplace")
    assert proc.returncode == 2, proc.stderr
    assert proc.stdout == b"", "拒绝时不许回显「备份: …」"
    after = {p.name for p in tmp_path.iterdir()}
    assert after == before, f"拒绝时写盘了：新增 {sorted(after - before)}"
    assert not (tmp_path / "renamed.bak.docx").exists()
    assert bogus.read_text(encoding="utf-8") == "i am not a zip package"


def test_mathify_valid_document_dry_run_still_works(tmp_path: Path):
    """合法文档 ``--dry-run`` 仍是 exit 0 + 既有汇总行（加校验没把成功路径改坏）。"""
    src = _degradable_docx(tmp_path)

    proc = run_cli("docx_mathify.py", src, "--dry-run")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    out = proc.stdout.decode("utf-8", "replace")
    assert out.startswith("段落 ")
    assert "备份" not in out, "--dry-run 不许出现备份回显"
    assert not (tmp_path / "in.bak.docx").exists()
