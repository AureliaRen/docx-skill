# -*- coding: utf-8 -*-
r"""显式角标标注：`S_{甲}` / `x^{乙}` —— 「手动编辑等价物」路线。

主人 2026-11-07 拍板的目标：**任意字符可做上下标、零误报、等同手动编辑**。
自动猜（`cn_subscript` 黑名单制）做不到零误报 —— 六卷语料实测把
`M公司` 打成 `M_{公}司`、`操作员A展开` 打成 `A_{展}开`，所以那条路线关闭。
显式标注能做到：作者写了花括号才转，一个字符都不猜。

本文件锁三件事（缺一不可）：
  * **正例**：显式标注 → 真 `w:vertAlign` 角标，任意字符（CJK / 符号 / 多字符）；
  * **零误报**：不带标注的 `S甲` 与正文原样穿过，一个字符都不动；
  * **兜底**：漏写 `}` 时不许把整段吞成公式（上限 + 句末标点硬停）。
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import linear_to_latex as L  # noqa: E402
import docx_mathify as M  # noqa: E402


# ----------------------------------------------------------- 工具：端到端
def _run_mathify(tmp_path: Path, paragraphs: list[str]) -> Path:
    """建 docx → 跑 mathify → 返回产物路径。"""
    src = tmp_path / "e2e.docx"
    d = Document()
    for t in paragraphs:
        d.add_paragraph(t)
    d.save(str(src))
    M.mathify(src, src, None, None)
    return src


def _script_runs(path: Path) -> list[tuple[str, str]]:
    """从产物 XML 里抠出所有 (vertAlign, 文字)。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return re.findall(
        r'<w:vertAlign w:val="(\w+)"/>(?:(?!</w:r>).)*?'
        r'<w:t[^>]*>([^<]*)</w:t>', xml, re.S)


def _plain_text(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml))


# ----------------------------------------------------------- 1. 正例：任意字符
@pytest.mark.parametrize("para, script, expect_text", [
    ("位移为 S_{甲}，则下列说法正确的是", "subscript", "甲"),
    ("速度 v_{乙} 大于零", "subscript", "乙"),
    ("合力 F_{合} 沿斜面向上", "subscript", "合"),
    ("第 1 秒内位移 x_{1}", "subscript", "1"),
    ("第 2 秒内位移 x_{2}", "subscript", "2"),
    ("压力 P_{max} 已饱和", "subscript", "max"),
    ("下标可多名 x_{甲乙}", "subscript", "甲乙"),
    ("带逗号的下标 a_{i,j}", "subscript", "i,j"),
    ("上标写汉字 x^{甲}", "superscript", "甲"),
    ("上标混合 x^{乙2}", "superscript", "乙2"),
    ("正负号上标 Zn^{2+}", "superscript", "2+"),
    ("普通幂次仍是上标 m^{2}", "superscript", "2"),
])
def test_explicit_annotation_becomes_real_script(tmp_path, para, script, expect_text):
    """★ 显式标注 → 真 vertAlign，正文里的任意字符都算数。"""
    out = _run_mathify(tmp_path, [para])
    runs = _script_runs(out)
    assert any(v == script and t == expect_text for v, t in runs), \
        f"{para!r} 未产出 {script}='{expect_text}'，实际 {runs}"


def test_cjk_subscript_multi_paragraph(tmp_path):
    """多个段落各带一个中文下标，互不干扰。"""
    out = _run_mathify(tmp_path, [
        "甲的速度 v_{甲}", "乙的速度 v_{乙}", "合速度 v_{合}",
    ])
    runs = _script_runs(out)
    got = [t for v, t in runs if v == "subscript"]
    assert got == ["甲", "乙", "合"], got


# ----------------------------------------------------------- 2. 零误报：不标注就不动
@pytest.mark.parametrize("para", [
    "位移分别为 s甲、s乙，则下列说法正确的是（　）",
    "物体甲的速度 v甲 大于物体乙的速度 v乙。",
    "M公司的28M战术通信终端",
    "操作员A展开，操作员B固定地锚",
    "对照组：图甲、表乙",
    "变量 m与 时间的 t时 关系",
])
def test_unannotated_cjk_is_never_touched(tmp_path, para):
    """★★ 零误报的核心断言：没有花括号标注，正文原样穿过。

    这几条正是自动猜路线会打错的原句（`M公司` / `操作员A展开`）——
    显式路线必须一条都不动。
    """
    out = _run_mathify(tmp_path, [para])
    assert _script_runs(out) == [], f"{para!r} 被误改了角标"
    text = _plain_text(out)
    for ch in "甲乙公司展开固定":
        if ch in para:
            assert ch in text, f"{para!r} 的 {ch!r} 丢了"


# ----------------------------------------------------------- 3. 判据层：显式标注是接受信号
def _atoms(text: str) -> list[dict]:
    return [{"ch": c, "rpr": None, "script": None} for c in text]


@pytest.mark.parametrize("text", [
    "S_{甲}", "v_{乙}", "x^{乙}", "F_{合}", "x_{1}", "P_{max}",
])
def test_explicit_script_alone_is_enough_for_a_span(text):
    """显式标注不必靠关系符论证自己是公式 —— 它是作者的主动声明。"""
    assert L.find_math_spans(_atoms(text)) == [(0, len(text))]


@pytest.mark.parametrize("text", [
    "s甲", "v甲", "操作员A展开", "M公司", "图甲", "m与", "t时",
])
def test_unannotated_cjk_is_not_a_span(text):
    """零误报的另一面：没有花括号，就不该被认成公式片段。"""
    assert L.find_math_spans(_atoms(text)) == []


def test_ascii_script_without_braces_still_needs_an_operator():
    """`x^2` 无括号时仍按旧规矩走（要关系符），本次改动没放开它。"""
    assert L.find_math_spans(_atoms("x^2")) == []
    assert L.find_math_spans(_atoms("y=x^{2}")) == [(0, 7)]


# ----------------------------------------------------------- 4. 兜底：漏写 } 不许吞段落
@pytest.mark.parametrize("text", [
    "位移 S_{甲这是一段很长的中文后面还有很多字继续写下去",
    "变量 x^{乙未闭合的内容非常长非常长非常长非常长非常长",
])
def test_unclosed_group_does_not_swallow_the_whole_paragraph(text):
    """手滑漏写 `}` 时，上限兜底：不许把整段吞成一个公式片段。

    ⚠️ 这条只锁 span 长度；**输出级**（不丢字、不误排）由下面 §4b 锁 ——
    REVIEW-CODE-R5 的 A-1 正是从「只测 span 不测输出」这个洞溜过去的。
    """
    spans = L.find_math_spans(_atoms(text))
    for s, e in spans:
        assert e - s < len(text), f"{text!r} 被整段吞掉：{spans}"


# ------------------------------------------------- 4b. 兜底（输出级）：漏写 } 不丢字
# REVIEW-CODE-R5 A-1：未闭合的 `_{` 若放行进公式管线，latex2mathml 会自动补闭
# 括号 —— `_{` 两字符被吃进公式结构（document.xml 里消失）、其后正文被排成
# 巨型下角标。基线（无显式标注特性）对同输入是零丢字纯文本，所以这是回归。
# 修法：span 文本停在未闭组内 → 整段退回纯文本。以下断言全部在**输出**上做。
@pytest.mark.parametrize("para", [
    "位移S_{甲这是一段很长的话，后面还有正文。",
    "变量 x^{乙未闭合就开始写别的句子了。",
    "段末漏写S_{甲",
])
def test_unclosed_group_output_keeps_every_character(tmp_path, para):
    r"""★ 笔误输入的端到端输出：无公式对象，且**逐字符不丢**。"""
    out = _run_mathify(tmp_path, [para])
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert "<m:oMath" not in xml, f"{para!r} 漏写 }} 仍产出了公式对象"
    assert "<m:sSub" not in xml, f"{para!r} 正文被排成了下角标"
    text = _plain_text(out)
    assert para in text, f"丢字了：输入 {para!r} 不在输出 {text!r} 里"


def test_unclosed_group_mixed_with_valid_equation_also_safe(tmp_path):
    r"""混合段：合法等式 + 漏写括号 —— 整段保守退回纯文本，仍不丢字。"""
    para = "取 x=1 与 S_{甲未闭合"
    out = _run_mathify(tmp_path, [para])
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert "<m:oMath" not in xml
    assert para in _plain_text(out)


def test_over_length_legit_subscript_falls_back_to_text(tmp_path):
    r"""★ 26 字合法长角标：被上限截断后**整段退回纯文本**，不丢字不误排。

    这是审查 A-1 点名的第二个触发面（不只笔误）。25 字仍在豁免内、正常转角标
    （见下一条），26 字起保守回退 —— 宁可不渲染，不许丢字。
    """
    para = "S_{甲乙丙丁戊己庚辛壬癸甲乙丙丁戊己庚辛壬癸甲乙丙丁戊}"  # 25 字内容
    para26 = "S_{甲乙丙丁戊己庚辛壬癸甲乙丙丁戊己庚辛壬癸甲乙丙丁戊己}"  # 26 字
    out = _run_mathify(tmp_path, [para26])
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    assert "<m:oMath" not in xml, "26 字长角标被截断后仍进了公式对象（A-1 回归）"
    assert para26 in _plain_text(out), "26 字长角标回退时丢字了"


def test_at_cap_legit_subscript_still_converts(tmp_path):
    r"""25 字（恰好在上限内）的合法角标必须照常转成真下标 —— 修 A-1 不许误伤。"""
    para = "S_{" + "甲" * 25 + "}"
    out = _run_mathify(tmp_path, [para])
    runs = _script_runs(out)
    subs = [t for v, t in runs if v == "subscript"]
    assert "甲" * 25 in subs, f"25 字角标没转成下标：{subs}"


@pytest.mark.parametrize("text", [
    "位移 S_{甲。后面是另一句话。",
    "变量 x^{乙；分号也是句末",
])
def test_sentence_terminator_stops_the_group(text):
    """句末标点是硬停：漏写括号也不会跨句吞掉下文。"""
    spans = L.find_math_spans(_atoms(text))
    for s, e in spans:
        taken = text[s:e]
        for stop in "。；":
            assert stop not in taken, f"{text!r} 跨过了 {stop!r}：{taken!r}"


def test_group_length_helper_matches_nesting():
    """`_open_script_group_len`：只认 `^{` / `_{` 的组，别的花括号不算。"""
    assert L._open_script_group_len("S_{") == 0
    assert L._open_script_group_len("S_{甲") == 1
    assert L._open_script_group_len("S_{甲}") == -1
    assert L._open_script_group_len(r"\text{abc}") == -1     # 别的组不算
    assert L._open_script_group_len("x^{a{b}c") == 5          # 外层仍开着


# ----------------------------------------------------------- 5. 与既有行为共存
def test_explicit_cjk_and_ascii_subscripts_coexist(tmp_path):
    out = _run_mathify(tmp_path, ["既有 H_{2}O 又有 S_{甲}"])
    runs = _script_runs(out)
    subs = [t for v, t in runs if v == "subscript"]
    assert "2" in subs and "甲" in subs, subs


def test_auto_guess_route_stays_off_by_default(tmp_path):
    """自动猜路线保持关闭（主人拍板：做不到零误报就关闭）。"""
    assert L.CN_SUBSCRIPT is False
    out = _run_mathify(tmp_path, ["位移 s甲 与 s乙"])
    assert _script_runs(out) == []


def test_docx_create_style_atom_flow_keeps_explicit_group(tmp_path):
    """既有 vertAlign run（真角标）+ 显式标注混排：两条路都要通。"""
    d = Document()
    p = d.add_paragraph()
    r = p.add_run("S")
    r2 = p.add_run("甲")
    r2.font.subscript = True
    r3 = p.add_run(" 与 A_{乙} 并列")
    src = tmp_path / "mixed.docx"
    d.save(str(src))
    M.mathify(src, src, None, None)
    subs = [t for v, t in _script_runs(src) if v == "subscript"]
    assert "甲" in subs, subs
    assert "乙" in subs, subs
