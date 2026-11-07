# -*- coding: utf-8 -*-
r"""免排版分流：`s=10km` 这类简单等式落成普通文字，不再产 OMML。

主人 2026-11-07：

    「我发现很简单的完全不需要 omml 的，比如 `s=10km` 这种，
      也会 LaTeX 转成 OMML。」

上一轮的混合策略（`MIXED_SUPERSCRIPT`）只救了**带角标**的片段 ——
`is_pure_script` 要求片段里至少有一个 `^`/`_`。连角标都没有的简单等式
（`s=10km` / `t=2h` / `m=5kg`）因此全部落到 OMML 分支，在 Word 里点进去
是公式编辑器、量符号被排成数学斜体、单位也跟着斜。

本文件锁四件事：

1. `no_layout_reason()` 的判据边界：**该降级的降级**（PLAN-ROUND4 §2.1 正例）；
2. **不该降级的一条都不降**（`F=ma`、`v=s/t`、`E=V/m`、链式等式、带角标的单位）；
3. `NO_LAYOUT_NEEDED = False` 时**完全**回到旧行为；
4. `is_pure_script` 家族语义未被动过（它是「真角标」的判据，混用会同时坏两件事，
   见 HANDOFF-ROUND3 §11.4.5 ⚠️）。

端到端部分只用 tmp_path 现造 docx，不依赖 `work/` 下的任何产物。
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx.oxml.ns import qn  # noqa: E402

import docx_mathify as M  # noqa: E402
import linear_to_latex as L  # noqa: E402


# ------------------------------------------------------- OMML 是否可用（CI 无 Office）
def _omml_available() -> bool:
    try:
        import math_omml
        math_omml.find_xslt()
        import latex2mathml  # noqa: F401
    except Exception:
        return False
    return True


needs_omml = pytest.mark.skipif(
    not _omml_available(),
    reason="needs latex2mathml + Microsoft MML2OMML.XSL (local Office install)",
)


# ---------------------------------------------------------------- 判据：该降级
# 全部是「单个量符号 = 数值 [+ 单位]」，没有任何二维排版需求。
@pytest.mark.parametrize("latex", [
    "s=10km",                       # 主人报的原始样例
    "t=2h",
    "m=5kg",
    "T=300K",
    r"v=3\,\text{m/s}",             # linear_to_latex 对 `v=3m/s` 的实际输出
    r"P=100\,\text{W}",
    r"c=0.5\,\text{mol/L}",         # 主人点名允许降级的那条（RHS 带斜杠单位）
    r"c=0.5\mathrm{mol/L}",         # \mathrm 形态同一含义
    "x=3",                          # 无单位的纯数值等式
    "x=-2",
    "x=+3",
    "f'=2",                         # LHS 允许尾随撇号
])
def test_plain_quantity_needs_no_layout(latex):
    reason = M.no_layout_reason(latex)
    assert reason is not None, f"{latex!r} 应当降级"
    assert reason in ("plain_quantity", "plain_equation")


def test_reason_distinguishes_quantity_from_bare_value():
    r"""带单位的记 plain_quantity，纯数值的记 plain_equation —— 报告要能分开看。"""
    assert M.no_layout_reason("s=10km") == "plain_quantity"
    assert M.no_layout_reason("x=3") == "plain_equation"


# ---------------------------------------------------------------- 判据：必须保留
# 这些要么 RHS 不是纯数值（`F=ma`、`v=s/t`）、要么 LHS 不是单字母（`f(x)=0`）、
# 要么有二维结构、要么带角标（交给既有混合策略）。
@pytest.mark.parametrize("latex", [
    "F=ma",                                          # 两个量相乘
    "U=IR",
    r"I=\frac{U}{R}",                                 # 真分式
    r"n=\frac{m}{M}",
    r"v=\frac{s}{t}",                                 # `v=s/t` 经 linear_to_latex 后
    "f(x)=0",                                        # LHS 不是单字母
    "y=f(x)",
    "A=(-1,3)",                                      # 集合/坐标
    r"B=[2,+\infty)",                                # \infty 不是未知命令，但 RHS 不是数值
    r"n\geq 2",                                      # 没有 `=`
    r"m=\rho V=1.2\times 60=72kg",                   # 链式等式
    r"E=\frac{V}{m}",                                # 场强：RHS 是量符号相除
    r"a=9.8\,\mathrm{m/s^{2}}",                      # 带角标 → 混合策略（文字角标）
    r"c=3\times\frac{10^{8}m}{s}",                   # 科学计数法 + 分式
    "k=2.5e3",                                       # `e3` 不是单位
    "y=2x",                                          # `x` 不是单位 → 函数关系
    "n=2t",                                          # `t` 刻意不进 UNITS（见下）
    "s=10m/",                                        # 单位因子不完整
    "s=10/x",
    "",                                              # 空串
])
def test_formulas_that_need_layout_stay_omml(latex):
    assert M.no_layout_reason(latex) is None, f"{latex!r} 不该被降级"


def test_unit_tail_is_a_whitelist_not_any_letters():
    r"""数值后面的裸字母必须是真单位，否则 `y=2x` / `n=2t` 会被读成「2 米」「2 吨」。

    `t` 在 `linear_to_latex.UNITS` 里被**刻意排除**（`v=s/t` 必须保持真分式），
    这条白名单因此顺带护住了 `n=2t`。这是判据刻意保守的地方：拿不准就留 OMML。
    """
    assert M.no_layout_reason("s=10kg") is not None
    assert M.no_layout_reason("s=10kg/m") is not None
    assert M.no_layout_reason("y=2x") is None
    assert M.no_layout_reason("n=2t") is None


@pytest.mark.parametrize("latex", [
    r"\frac{1}{2}",
    r"\sqrt{2}",
    r"\int_{0}^{1}",
    r"\sum_{n=1}^{N}",
    r"\left|x\right|",
])
def test_structural_commands_block_the_downgrade(latex):
    r"""`structural_reason()` 命中的一律不降级（判据第一条）。"""
    assert M.structural_reason(latex) is not None
    assert M.no_layout_reason(latex) is None


@pytest.mark.parametrize("latex", [
    "x^{2}=9",
    "v_{0}=3",
])
def test_any_script_blocks_the_downgrade(latex):
    r"""带 `^` / `_` 的一律交给混合策略，免排版层不插手。

    这条是两层判据的**互斥契约**：`is_pure_script` 要求有角标，
    `no_layout_reason` 要求没角标，谁都别越界。
    """
    assert M.no_layout_reason(latex) is None


# ---------------------------------------------------------------- 既有语义未变
@pytest.mark.parametrize("latex, expected", [
    ("s=10km", (False, "no_script")),
    ("t=2h", (False, "no_script")),
    ("F=ma", (False, "no_script")),
    ("y=f(x)", (False, "no_script")),
    (r"a^{2}", (True, "pure_script")),
    # 无角标时 `no_script` **先于**结构判定返回 —— 这是既有的判据顺序，不许动
    (r"\frac{U}{R}", (False, "no_script")),
    # 有角标才轮到结构判定
    (r"\frac{x^{2}}{16}=1", (False, "structural:\\frac")),
    ("", (False, "empty")),
])
def test_is_pure_script_semantics_untouched(latex, expected):
    r"""`is_pure_script_detail` 的结论与理由必须与改动前逐字一致。

    它是「真角标」的判据，被 `_om_or_scripts` 的第二层直接使用；
    放宽它就等于同时改坏「简单等式」与「真角标」两件事（HANDOFF §11.4.5）。
    """
    assert M.is_pure_script_detail(latex) == expected


def test_structural_reason_untouched():
    r"""`\infty` 里含 `\in` 前缀，但命令边界判定仍不该把它当 `\int`。"""
    assert M.structural_reason(r"x\to+\infty") is None
    assert M.structural_reason(r"\int_0^1") == r"\int"


# ---------------------------------------------------------------- 分流：开关
def test_switch_on_downgrades_and_counts():
    stats: dict = {}
    runs = M._om_or_scripts("s=10km", None, stats)
    assert runs is not None
    assert stats["plain_no_layout"] == 1
    assert stats["plain_no_layout_rows"] == [
        {"latex": "s=10km", "reason": "plain_quantity"}]
    # 降级进的是「免排版」这个桶，不是「文字角标」那个桶
    assert stats.get("text_scripts", 0) == 0


def test_switch_off_restores_omml_only(monkeypatch):
    r"""`NO_LAYOUT_NEEDED = False` 时必须**完全**回到改动前的行为。"""
    monkeypatch.setattr(M, "NO_LAYOUT_NEEDED", False)
    stats: dict = {}
    assert M._om_or_scripts("s=10km", None, stats) is None
    assert M._om_or_scripts("t=2h", None, stats) is None
    assert "plain_no_layout" not in stats
    # 混合策略那层不受影响：带角标的照旧降级
    runs = M._om_or_scripts(r"a^{2}", None, stats)
    assert runs is not None
    assert stats["text_scripts"] == 1


def test_master_switch_off_also_stops_no_layout(monkeypatch):
    r"""`MIXED_SUPERSCRIPT = False` 是「全部 OMML」的总闸，免排版也一起停。

    否则「置 False 就等于旧行为」这句承诺对 MIXED_SUPERSCRIPT 就不成立。
    """
    monkeypatch.setattr(M, "MIXED_SUPERSCRIPT", False)
    stats: dict = {}
    assert M._om_or_scripts("s=10km", None, stats) is None
    assert "plain_no_layout" not in stats


def test_downgraded_runs_carry_the_original_rpr():
    r"""降级出来的 run 必须沿用原 rPr —— 正文字体字号不能被换掉。"""
    from docx.oxml import OxmlElement
    rpr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")
    rpr.append(sz)

    runs = M._om_or_scripts("s=10km", rpr, {})
    assert runs is not None
    assert runs[0].find(qn("w:rPr")) is not None
    assert "".join(t.text for t in runs[0].iter(qn("w:t"))) == "s=10km"


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
    return src, out


def _doc_xml(out: Path) -> str:
    with zipfile.ZipFile(str(out)) as z:
        return z.read("word/document.xml").decode("utf-8")


def test_end_to_end_plain_equations_produce_no_omml(tmp_path):
    r"""`s=10km` 这类走完全程后落在 `w:t` 里，且整篇**一个 OMML 都没有**。"""
    src, out = _build(tmp_path, [
        "s=10km",
        "t=2h",
        "m=5kg",
        "总路程 s=10km，求时间",
    ])
    stats = M.mathify(src, out, None, None)

    xml = _doc_xml(out)
    assert "<m:oMath" not in xml, "简单等式不该再产 OMML"
    assert "s=10km" in xml, "`s=10km` 没落进正文文字"
    assert "t=2h" in xml
    assert "m=5kg" in xml
    # 中文与公式混排的段落不能被吞掉
    assert "总路程 " in xml
    assert stats["plain_no_layout"] == 4, stats["plain_no_layout"]
    assert stats["failed"] == []


def test_end_to_end_quantity_with_unit_from_linearizer(tmp_path):
    r"""`v=3m/s` → `v=3\,\text{m/s}` 也降级，单位里的 `/` 不出分式。"""
    src, out = _build(tmp_path, ["v=3m/s"])
    stats = M.mathify(src, out, None, None)

    xml = _doc_xml(out)
    assert "<m:oMath" not in xml
    assert "<m:f>" not in xml, "单位斜杠被误转成分式"
    assert "m/s" in xml
    assert stats["plain_no_layout"] == 1


@needs_omml
def test_end_to_end_multiplication_equation_stays_omml(tmp_path):
    r"""`F=ma` 是「量符号相乘」，必须还是 OMML 公式对象。"""
    src, out = _build(tmp_path, ["F=ma"])
    M.mathify(src, out, None, None)

    xml = _doc_xml(out)
    assert "<m:oMath" in xml, "`F=ma` 被误降级了"


@needs_omml
def test_end_to_end_mixed_document_keeps_both_forms(tmp_path):
    r"""一份文档里两种形态并存：简单等式落文字，真分式留 OMML。"""
    src, out = _build(tmp_path, ["s=10km", "x^2/16+y^2/9=1"])
    stats = M.mathify(src, out, None, None)

    xml = _doc_xml(out)
    assert "s=10km" in xml
    assert "<m:oMath" in xml, "真分式必须还是 OMML"
    assert "<m:f>" in xml
    assert stats["plain_no_layout"] == 1


def test_end_to_end_switch_off_restores_omml(tmp_path, monkeypatch):
    r"""开关关掉后同一份输入必须重新产 OMML —— 回退路径要真的能回退。"""
    monkeypatch.setattr(M, "NO_LAYOUT_NEEDED", False)
    src, out = _build(tmp_path, ["s=10km"])
    stats = M.mathify(src, out, None, None)

    if not _omml_available():
        pytest.skip("needs latex2mathml + Microsoft MML2OMML.XSL")
    assert "<m:oMath" in _doc_xml(out)
    assert stats["plain_no_layout"] == 0


# ---------------------------------------------------------------- 报告
def test_report_lists_downgraded_spans(tmp_path):
    r"""`--report` 的 Markdown 里要有一节列出被降级的片段，方便人工核对。"""
    src, out = _build(tmp_path, ["s=10km", "F=ma"])
    report = tmp_path / "conv.md"
    M.mathify(src, out, report, None)

    txt = report.read_text(encoding="utf-8")
    assert "免排版降级" in txt
    assert "plain_quantity" in txt
    assert "`s=10km`" in txt
    assert "免排版降级：1" in txt
    # 未被降级的不能出现在这一节里
    head, _, tail = txt.partition("## 免排版降级")
    assert "F=ma" not in tail, "`F=ma` 不该出现在降级清单里"


def test_report_omits_the_section_when_nothing_is_downgraded(tmp_path):
    r"""没有任何降级时不该凭空长出一节空表（汇总行仍要报 0）。"""
    src, out = _build(tmp_path, ["F=ma"])
    report = tmp_path / "conv.md"
    M.mathify(src, out, report, None)

    txt = report.read_text(encoding="utf-8")
    assert "## 免排版降级" not in txt
    assert "免排版降级：0" in txt


def test_dry_run_report_counts_what_would_be_downgraded(tmp_path):
    r"""dry-run 不跑 rebuild，`stats['plain_no_layout']` 必然是 0。

    那时代码仍要写出**预览**清单（这正是 `--dry-run` 的用途：先看会改什么），
    所以汇总行的数字必须取明细条数，不能让它和表里的条目数打架。
    """
    src, _ = _build(tmp_path, ["s=10km", "F=ma"])
    report = tmp_path / "dry.md"
    stats = M.mathify(src, None, report, None, dry_run=True)

    assert stats["plain_no_layout"] == 0          # 没 rebuild，计数器就是 0
    txt = report.read_text(encoding="utf-8")
    assert "免排版降级：1" in txt, "dry-run 的汇总数字与明细表不一致"
    assert "## 免排版降级" in txt


# ---------------------------------------------------------------- 与线性式串联
# 判据吃的是 `linear_to_latex` 的输出，不是人手写的 LaTeX。这里逐条串起来，
# 保证「原线性写法 → 降级」这条链路真的通，而不是只有一个函数自测绿。
@pytest.mark.parametrize("linear, expect_down", [
    ("s=10km", True),
    ("t=2h", True),
    ("m=5kg", True),
    ("v=3m/s", True),
    ("c=0.5 mol/L", True),
    ("F=ma", False),
    ("U=IR", False),
    ("I=U/R", False),
    ("n=m/M", False),
    ("v=s/t", False),
    ("E=V/m", False),
    ("f(x)=0", False),
    ("y=f(x)", False),
    ("A=(-1,3)", False),
    ("m=ρV=1.2×60=72kg", False),
    ("a=9.8m/s²", False),
])
def test_linear_to_latex_then_no_layout(linear, expect_down):
    latex = L.linear_to_latex(linear)
    assert (M.no_layout_reason(latex) is not None) is expect_down, \
        f"{linear!r} -> {latex!r}"

# ------------------------------------------------- 后缀型单位（温度 / 角度 / 千分）
# linear_to_latex.UNITS 收的是「可独立成因子」的量纲符号；℃ 与 ° 是紧贴
# 数值的后缀，不在那张表里。不单独放行的话 T=25℃ 会白留一个 OMML 公式对象 ——
# 保守方向出错（不误伤），但温度是理科卷高频量，值得降。
@pytest.mark.parametrize("linear", [
    "T=25\u2103",
    "x=80%",
    "L=60dB",
])
def test_suffix_units_are_downgraded(linear):
    latex = L.linear_to_latex(linear)
    assert M.no_layout_reason(latex) == "plain_quantity", \
        f"{linear!r} -> {latex!r} 应降级"


@pytest.mark.parametrize("linear", [
    "y=2x",      # x 不是单位，是变量 —— 2 倍 x
    "n=2t",      # t 不在 UNITS（v=s/t 要靠它保持分式）
    "F=ma",      # 关系式，不是量值
])
def test_letters_that_are_not_units_are_not_downgraded(linear):
    latex = L.linear_to_latex(linear)
    assert M.no_layout_reason(latex) is None, f"{linear!r} -> {latex!r} 不该降级"
