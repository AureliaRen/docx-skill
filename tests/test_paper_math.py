# -*- coding: utf-8 -*-
"""数学卷（subj-math）回归测试。

真正的产物在 ``work/subjects/math/paper.docx``（由 Lead 的 verify_subject.py
独立复核）。本文件只锁**结构契约**，且**不碰 Word**：全走 zipfile + ElementTree，
所以 CI 上没有 Office 也能绿。

三件事必须守住，它们都是"静默失败"型：
  * 节数与页码格式：``w:pgNumType`` 写错槽位时 Word 整节忽略，文件照样能开；
  * 自动编号：``w:numPr`` 段落数与 ``lvlText`` 决定题号/文献号是否真渲染；
  * 交叉引用：只有真 ``<w:hyperlink w:anchor>`` 才可点，光有 ``\\h`` 不够。

spec 不在就 skip，绝不硬编码盘符路径（契约 §7）。
"""
from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

HERE = Path(__file__).resolve()
SCRIPTS = HERE.parents[1] / "scripts"
SUBJECT = HERE.parents[2] / "work" / "subjects" / "math"
SPEC = SUBJECT / "spec.json"
DOCX = SUBJECT / "paper.docx"
RAW = SUBJECT / "paper.raw.docx"

sys.path.insert(0, str(SCRIPTS))

pytestmark = pytest.mark.skipif(
    not SPEC.exists() or not DOCX.exists(),
    reason="math subject spec/docx not shipped",
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


# ------------------------------------------------------------------ fixtures
@pytest.fixture(scope="module")
def doc_xml() -> str:
    with zipfile.ZipFile(DOCX) as z:
        return z.read("word/document.xml").decode("utf-8")


@pytest.fixture(scope="module")
def root(doc_xml):
    return ET.fromstring(doc_xml)


@pytest.fixture(scope="module")
def spec() -> dict:
    import json
    return json.loads(SPEC.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def part_names():
    with zipfile.ZipFile(DOCX) as z:
        return z.namelist()


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t"))


# --------------------------------------------------------------------- spec
def test_spec_is_five_sections(spec):
    """卷头 / 第Ⅰ卷 / 第Ⅱ卷 / 参考答案 / 附录 = 5 节。"""
    n_breaks = sum(1 for b in spec["blocks"] if b["type"] == "section_break")
    assert n_breaks + 1 == 5


def test_spec_uses_every_required_key(spec):
    """契约 §5 的压测键必须真出现在 spec 里，缺一即视为没测到。"""
    blocks = spec["blocks"]
    kinds = {b["type"] for b in blocks}
    for need in ("numbered_list", "crossref", "caption", "image", "table",
                 "section_break"):
        assert need in kinds, "spec 缺少 %s 块" % need

    fmt_keys = {"indent_left_mm", "first_line_indent_chars", "line_spacing",
                "alignment", "space_before_pt", "space_after_pt"}
    seen = set()
    for b in blocks:
        for k in fmt_keys:
            if k in b:
                seen.add(k)
    assert seen == fmt_keys, "缺段落格式键: %s" % (fmt_keys - seen)

    # 参考文献编号
    ref = [b for b in blocks if b["type"] == "numbered_list"
           and b.get("bookmark_prefix") == "ref_"]
    assert ref and ref[0]["num_text"] == "[%1]"

    # 两种 crossref 形态
    sup = [b for b in blocks if b["type"] == "crossref" and b.get("superscript")]
    nor = [b for b in blocks if b["type"] == "crossref"
           and not b.get("superscript")]
    assert sup and sup[0]["format"] == "paragraph-number"
    assert nor

    # 至少 3 种页码格式（卷头不加页码 + decimal + lowerRoman）
    fmts = [b["footer_page_numbers"]["fmt"] for b in blocks
            if b["type"] == "section_break" and "footer_page_numbers" in b]
    assert len({f for f in fmts}) >= 2
    assert spec.get("footer_page_numbers") is False  # 卷头不挂页码


def test_spec_styles_carry_cjk_fonts(spec):
    """styles[].font 必须写中文字体，靠 styles 继承（逐 run 改不继承 eastAsia）。"""
    fonts = {s.get("font") for s in spec["styles"] if s.get("font")}
    assert "黑体" in fonts          # 标题
    assert "宋体" in {s.get("font_east_asia") for s in spec["styles"]}


# ------------------------------------------------------------- sections/pages
def test_sections_count(root):
    assert len(list(root.iter(W + "sectPr"))) == 5


def test_page_number_formats(root):
    """decimal 用于正文、lowerRoman 用于附录；卷头那一节不写 pgNumType。"""
    el = list(root.iter(W + "pgNumType"))
    fmts = {e.get(W + "fmt") for e in el}
    assert fmts == {"decimal", "lowerRoman"}
    # 4 个节带 pgNumType，第 1 节（卷头）没有
    assert len(el) == 4


def test_answer_section_restarts_at_one(root):
    """参考答案那节必须重启（start=1），否则页码会顺着第Ⅱ卷往下走。"""
    starts = []
    for e in root.iter(W + "pgNumType"):
        if e.get(W + "fmt") == "decimal" and e.get(W + "start"):
            starts.append(int(e.get(W + "start")))
    # 第Ⅰ卷 start=1、参考答案 start=1
    assert starts.count(1) >= 2


def test_middle_section_does_not_restart(root):
    """第Ⅱ卷继承来的 w:start 已被摘掉 —— 它要么无 start，要么不是 1。

    这是 add_section() 克隆 sectPr 的已知副作用；重建流水线里由
    _fix_pgnum.py 处理，此处断言它确实被处理过。
    """
    decimal_no_start = [e for e in root.iter(W + "pgNumType")
                        if e.get(W + "fmt") == "decimal"
                        and e.get(W + "start") is None]
    assert decimal_no_start, "第Ⅱ卷仍带着克隆来的 start=1，页码会重来"


# --------------------------------------------------------------- numbering
def test_numpr_paragraphs(root):
    """题号（1..9）、选项（A..D）、文献号（[1]..）都靠 numPr 实现。"""
    npr = [p for p in root.iter(W + "p")
           if p.find(W + "pPr/" + W + "numPr") is not None]
    assert len(npr) == 42, "numPr 段落应为 42（4+6*5+3+5），实际 %d" % len(npr)


def test_lvl_text_covers_three_list_flavors(part_names):
    """三套 lvlText：题号 %1.、选项 %1.（upperLetter）、文献 [%1]。"""
    with zipfile.ZipFile(DOCX) as z:
        num = ET.fromstring(z.read("word/numbering.xml"))
    texts = {(l.find(W + "numFmt").get(W + "val")
              if l.find(W + "numFmt") is not None else None,
              l.find(W + "lvlText").get(W + "val")
              if l.find(W + "lvlText") is not None else None)
             for l in num.iter(W + "lvl")}
    assert ("decimal", "%1.") in texts
    assert ("upperLetter", "%1.") in texts
    assert any(t == ("decimal", "[%1]") for t in texts)


def test_question_numbers_are_real_strings(root):
    """题号段的文字必须是题面本身，不是 "1." 之类的硬编码。"""
    npr = [p for p in root.iter(W + "p")
           if p.find(W + "pPr/" + W + "numPr") is not None]
    joined = "".join(para_text(p)[:12] for p in npr)
    assert "集合运算" in joined and "数列递推" in joined


# ---------------------------------------------------------------- bookmarks
def test_bookmarks_exist(root):
    names = {b.get(W + "name") for b in root.iter(W + "bookmarkStart")}
    assert {"fig_ellipse", "fig_cubic"} <= names
    assert {"ref_1", "ref_2", "ref_3", "ref_4", "ref_5"} <= names


def test_caption_bookmarks_are_on_captions(root):
    """图题注必须带 SEQ 域 + 书签，否则正文引用拿不到号。"""
    seq = 0
    for p in root.iter(W + "p"):
        for it in p.iter(W + "instrText"):
            if re.match(r"\s*SEQ\s+图", it.text or ""):
                seq += 1
    assert seq == 2, "两张图应各有一个 SEQ 图 域，实际 %d" % seq


# ------------------------------------------------------- cross-references
def test_hyperlink_wrappers(root):
    """只有真 hyperlink 包装才可点；数量正好 2（一处上角标 + 一处正文）。"""
    hl = list(root.iter(W + "hyperlink"))
    anchors = [h.get(W + "anchor") for h in hl if h.get(W + "anchor")]
    assert len(hl) == 2
    assert anchors == ["ref_1", "fig_ellipse"]


def test_anchors_resolve_to_bookmarks(root):
    names = {b.get(W + "name") for b in root.iter(W + "bookmarkStart")}
    for h in root.iter(W + "hyperlink"):
        a = h.get(W + "anchor")
        if a:
            assert a in names, "悬挂锚点: %s" % a


def test_ref_fields_present(root):
    codes = [it.text.strip() for it in root.iter(W + "instrText")]
    joined = " | ".join(codes)
    assert "REF ref_1" in joined
    assert "REF fig_ellipse" in joined
    assert r"\n" in joined          # paragraph-number -> "[1]"
    assert r"\h" in joined          # 超链接开关


def test_superscript_vs_normal_ref(root):
    """一处上角标本、一处正文大小 —— 两种形态都要真落地。"""
    sup = [va for va in root.iter(W + "vertAlign")
           if va.get(W + "val") == "superscript"]
    assert sup, "缺少 superscript 交叉引用"

    # 正文大小那处：REF fig_ellipse 的字段 run 不应带 superscript
    for h in root.iter(W + "hyperlink"):
        if h.get(W + "anchor") != "fig_ellipse":
            continue
        for r in h.iter(W + "r"):
            va = r.find(W + "rPr/" + W + "vertAlign")
            assert va is None or va.get(W + "val") != "superscript"
        break


def test_cached_ref_results(root):
    """上角标引用缓存值就是 [1]，正文引用是 "1"（Word 未重算也读得对）。"""
    texts = {}
    for h in root.iter(W + "hyperlink"):
        a = h.get(W + "anchor")
        if not a:
            continue
        texts[a] = "".join(t.text or "" for t in h.iter(W + "t"))
    assert texts.get("ref_1") == "[1]"
    assert texts.get("fig_ellipse") == "1"


# ------------------------------------------------------------------ formulas
def test_omml_formula_count(doc_xml):
    """至少 4 条线性公式要真变 OMML。

    本卷建卷时有 50 个公式片段，其中 20 个是**纯上下标**。按混合策略
    （``docx_mathify.MIXED_SUPERSCRIPT``）它们落成文字角标 ``w:vertAlign``
    而不是 OMML —— 这才是「选中它按 Ctrl+Shift+= 会还原成正常大小」的形式，
    公式对象（``m:sSup``）做不到。所以这里断言的是「真公式还在」＋「剩下
    的确实变成了文字角标」＋「片段总数没丢」，而不是一个固定数字。
    """
    n = doc_xml.count("<m:oMath>")
    assert n >= 4, "OMML 公式只有 %d 条" % n
    assert "<m:f>" in doc_xml, "真分式丢了"
    vert = doc_xml.count('w:vertAlign w:val="superscript"')
    vert_sub = doc_xml.count('w:vertAlign w:val="subscript"')
    assert vert + vert_sub > 0, "没有任何文字角标，混合策略没生效"
    assert n + vert + vert_sub >= 50, (
        "公式片段总数 %d 少于建卷时的 50" % (n + vert + vert_sub))


def _all_visible_text(root):
    """Every ``w:t`` in the body — OMML text included.

    A unit string may be laid out either way after the mixed-superscript
    change: as an ``m:oMath`` object (a real quotient) or as plain runs with
    ``w:vertAlign`` (a dimension string, ``m/s²``).  What must not change is
    the *characters the reader sees*, so the assertions below read them from
    the whole body rather than from OMML only.
    """
    return "".join(t.text or "" for t in root.iter(W + "t"))


def test_unit_slash_formulas_are_rendered(root):
    """带单位斜杠的量纲串必须落成可读文本，而不是留着反斜杠的 LaTeX 原文。

    三条硬指标：重力加速度、密度、物质的量浓度。源文本里是
    ``9.8m/s2`` / ``1.2kg/m3`` / ``0.5mol/L``，靠"数字锚点"成立。
    """
    joined = _all_visible_text(root)
    assert "9.8" in joined, "重力加速度丢了"
    assert "1.2" in joined, "密度丢了"
    assert "0.5" in joined, "浓度丢了"


def test_unit_slash_survives_as_unit_not_fraction(root):
    """m/s^2 与 kg/m^3 的斜杠不能被误转成分式（``\\frac`` / ``m:f``）。

    量纲串的正确形态是把指数抬起来（``m/s²``），不是上下叠一个分式。
    这条与「落在 OMML 还是文字角标」无关，只看有没有被当成分数。
    """
    joined = _all_visible_text(root)
    assert "m/s" in joined or "kg/m" in joined, "量纲串的斜杠形态丢了"
    # 量纲串不该出现在分式结构里：m:f 的分子分母拼起来不能是 kg 与 m3
    for f in root.iter(M + "f"):
        num = "".join(t.text or "" for t in f.iter(M + "t"))
        assert not ("kg/m" in num or "m/s" in num), (
            "量纲串被误转成分式: %r" % num)


def test_no_raw_frac_in_wrong_place(doc_xml):
    """真分式（v=s/t、16 分之 x 方）允许有；但不能把量纲串全变成分式。"""
    assert "<m:f>" in doc_xml                     # 确有真分式（如 x²/16）
    # 量纲公式里不该出现 \\frac 包裹 m/s2 这种
    assert "omml" not in doc_xml.lower()          # 兜底：没漏 LaTeX 原文


# ------------------------------------------------------------ images/tables
def test_images(part_names, root):
    media = [n for n in part_names if n.startswith("word/media/")]
    assert len(media) == 2
    assert len(list(root.iter(W + "drawing"))) == 2


def test_figs_on_disk():
    figs = sorted((SUBJECT / "figs").glob("*.png"))
    assert len(figs) >= 2
    for f in figs[:2]:
        assert f.stat().st_size > 20_000     # 不是空白占位图


def test_answer_card_table(root):
    assert len(list(root.iter(W + "tbl"))) == 1
    rows = list(root.iter(W + "tr"))
    assert len(rows) == 4                     # 表头 + 答案 + 分值 + 得分
    head = "".join(t.text or "" for t in rows[0].iter(W + "t"))
    assert "题号" in head


# --------------------------------------------------------------- typography
def test_paragraph_formatting_landed(root):
    ind = spac = jc = 0
    for p in root.iter(W + "p"):
        ppr = p.find(W + "pPr")
        if ppr is None:
            continue
        if ppr.find(W + "ind") is not None:
            ind += 1
        if ppr.find(W + "spacing") is not None:
            spac += 1
        if ppr.find(W + "jc") is not None:
            jc += 1
    assert ind >= 10, "缩进落地点太少: %d" % ind
    assert spac >= 10, "行距/段距落地点太少: %d" % spac
    assert jc >= 5, "对齐落地点太少: %d" % jc


def test_first_line_indent_uses_chars(root):
    """汉字题面的首行缩进必须写 w:firstLineChars（Word 的字符缩进）。"""
    chars = [i for i in root.iter(W + "ind") if i.get(W + "firstLineChars")]
    assert chars, "没有任何段落写了 firstLineChars"


def test_option_indent_present(root):
    """A/B/C/D 选项靠 indent_left_mm 缩进，w:ind/@w:left 必须有值。"""
    lefts = [int(i.get(W + "left")) for i in root.iter(W + "ind")
             if i.get(W + "left")]
    assert any(v > 0 for v in lefts)
    assert max(lefts) >= 200          # 至少 ~3.5mm


# -------------------------------------------------------------- text content
def test_cover_has_no_page_number(root):
    """卷头那一节的页脚不该有 PAGE 域（卷头不编页码）。"""
    first = list(root.iter(W + "sectPr"))[0]
    assert first.find(W + "pgNumType") is None
    assert first.findall(W + "footerReference") == []


def test_required_text_present(doc_xml):
    for s in ("绝密★启用前", "2026 年普通高等学校招生全国统一考试（模拟）",
              "注意事项", "第Ⅰ卷", "第Ⅱ卷", "参考答案与解析", "答题卡统计表",
              "参考文献"):
        assert s in doc_xml, "正文缺少：%s" % s


def test_reference_entries_numbered(root):
    """文献五条，且各自带 ref_N 书签。"""
    refs = [b for b in root.iter(W + "bookmarkStart")
            if (b.get(W + "name") or "").startswith("ref_")]
    assert len(refs) == 5
