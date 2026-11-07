# -*- coding: utf-8 -*-
"""物理卷（subj-physics）回归测试。

真正的产物在 ``work/subjects/physics/paper.docx``（由 Lead 的
``verify_subject.py`` 独立复核）。本文件只锁**结构契约**，且**不碰 Word**：
全走 ``zipfile`` + ``xml.etree.ElementTree``，所以 CI 上没有 Office 也能绿。

守的是四类"静默失败"——文件能打开，但语义已经丢了：

  * 节数与页码格式：``w:pgNumType`` 写错槽位时 Word 整节忽略页码设置；
  * 自动编号：``w:numPr`` 段落数 + ``lvlText`` 决定题号/文献号是否真渲染；
  * 交叉引用：只有真 ``<w:hyperlink w:anchor>`` 才可点，光有 ``\\h`` 不够；
  * 公式：``m:oMath`` 里的单位斜杠不能被误转成分式。

spec/docx 不在就 skip，绝不硬编码盘符路径（契约 §7）。
"""
from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

HERE = Path(__file__).resolve()
SCRIPTS = HERE.parents[1] / "scripts"
SUBJECT = HERE.parents[1].parent / "work" / "subjects" / "physics"
SPEC = SUBJECT / "spec.json"
DOCX = SUBJECT / "paper.docx"

sys.path.insert(0, str(SCRIPTS))

pytestmark = pytest.mark.skipif(
    not SPEC.exists() or not DOCX.exists(),
    reason="subject spec not shipped",
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


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
    return json.loads(SPEC.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def part_names():
    with zipfile.ZipFile(DOCX) as z:
        return z.namelist()


def para_text(p) -> str:
    """段落可见文字：``w:t`` 有字，OMML 的 ``m:t`` 也要算进去。"""
    out = []
    for n in p.iter():
        if n.tag == W + "t" or n.tag == M + "t":
            out.append(n.text or "")
    return "".join(out)


def bookmark_names(root) -> set:
    return {b.get(W + "name") for b in root.iter(W + "bookmarkStart")
            if b.get(W + "name")}


# --------------------------------------------------------------------- spec
def test_spec_has_at_least_three_sections(spec):
    """卷头 / 第Ⅰ卷 / 参考答案 —— 两个 section_break = 3 节。"""
    breaks = sum(1 for b in spec["blocks"] if b["type"] == "section_break")
    assert breaks + 1 >= 3, "spec 只有 %d 节" % (breaks + 1)


def test_spec_uses_every_required_key(spec):
    """契约 §5 的压测键必须真出现，缺一即视为没测到。"""
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

    ref = [b for b in blocks if b["type"] == "numbered_list"
           and b.get("bookmark_prefix") == "ref_"]
    assert ref, "参考文献没有带 ref_ 前缀的编号列表"
    assert ref[0]["num_text"] == "[%1]"

    sup = [b for b in blocks if b["type"] == "crossref" and b.get("superscript")]
    nor = [b for b in blocks if b["type"] == "crossref"
           and not b.get("superscript")]
    assert sup, "缺少上角标交叉引用"
    assert sup[0]["format"] == "paragraph-number"
    assert nor, "缺少正常大小的交叉引用"


def test_spec_declares_two_page_number_formats(spec):
    fmts = {b["footer_page_numbers"]["fmt"] for b in spec["blocks"]
            if b["type"] == "section_break" and "footer_page_numbers" in b}
    assert len(fmts) >= 2, "双套页码没测到：%s" % fmts
    # 卷头那一节是 spec 的第 1 节（没有任何 section_break 在它前面），
    # 所以两个 section_break 各自带一套页码 ⇒ 卷头天然没有页码。
    breaks = [b for b in spec["blocks"] if b["type"] == "section_break"]
    assert all("footer_page_numbers" in b for b in breaks), "有分节没写页码配置"


def test_spec_images_are_three_own_figs(spec):
    imgs = [b for b in spec["blocks"] if b["type"] == "image"]
    assert len(imgs) >= 3, "自产图不足 3 张"
    for b in imgs:
        assert Path(b["path"]).name.startswith("fig"), b["path"]


def test_spec_tables_at_least_two(spec):
    assert len([b for b in spec["blocks"] if b["type"] == "table"]) >= 2


def test_spec_carries_cjk_fonts(spec):
    """字体靠 styles 继承（段落级 runs 不继承 eastAsia）。"""
    fonts = {s.get("font") for s in spec["styles"] if s.get("font")}
    assert "黑体" in fonts, "标题字体缺失"
    assert "宋体" in fonts, "正文字体缺失"


# ------------------------------------------------------------- sections/pages
def test_sections_count(root):
    """三个 ``w:sectPr`` = 卷头 / 第Ⅰ卷 / 参考答案。"""
    assert len(list(root.iter(W + "sectPr"))) >= 3


def test_page_number_formats(root):
    """``lowerRoman`` 用于第Ⅰ、Ⅱ卷，``decimal`` 用于参考答案。"""
    fmts = {e.get(W + "fmt") for e in root.iter(W + "pgNumType")}
    fmts.discard(None)
    assert len(fmts) >= 2, "只有一种页码格式: %s" % fmts
    assert fmts == {"lowerRoman", "decimal"}, fmts


def test_cover_section_has_no_page_numbers(root):
    """卷头那节不写 ``w:pgNumType``、也不挂 footer。"""
    first = list(root.iter(W + "sectPr"))[0]
    assert first.find(W + "pgNumType") is None
    assert first.findall(W + "footerReference") == []


def test_body_sections_carry_page_fields(part_names):
    """后两节的页脚里必须有真 `` PAGE `` 域，否则页码根本不出。"""
    footers = [n for n in part_names if n.startswith("word/footer")]
    assert len(footers) >= 2, footers
    with zipfile.ZipFile(DOCX) as z:
        codes = " ".join(
            (ET.fromstring(z.read(n)).findtext(".//" + W + "instrText") or "")
            for n in footers)
    assert "PAGE" in codes, "页脚里没有 PAGE 域"


# --------------------------------------------------------------- numbering
def test_numpr_paragraphs(root):
    """注意事项 4 条 + 参考文献 3 条 = 7 段自动编号。"""
    npr = [p for p in root.iter(W + "p")
           if p.find(W + "pPr/" + W + "numPr") is not None]
    assert len(npr) >= 1, "题目没有自动编号"
    assert len(npr) == 7, "numPr 段落应为 7，实际 %d" % len(npr)


def test_numbering_definitions(part_names):
    """两套 lvlText：注意事项 ``%1.``、参考文献 ``[%1]``。"""
    assert "word/numbering.xml" in part_names, "缺 numbering 部件"
    with zipfile.ZipFile(DOCX) as z:
        num = ET.fromstring(z.read("word/numbering.xml"))
    pairs = {(
        (l.find(W + "numFmt").get(W + "val")
         if l.find(W + "numFmt") is not None else None),
        (l.find(W + "lvlText").get(W + "val")
         if l.find(W + "lvlText") is not None else None),
    ) for l in num.iter(W + "lvl")}
    assert ("decimal", "%1.") in pairs, "题号格式没定义"
    assert ("decimal", "[%1]") in pairs, "文献号格式没定义"


def test_numbered_paragraphs_carry_no_hardcoded_number(root):
    """编号段的文字必须是题面本身，不能把 "1." 硬写进 run。"""
    npr = [p for p in root.iter(W + "p")
           if p.find(W + "pPr/" + W + "numPr") is not None]
    for p in npr:
        txt = para_text(p).lstrip()
        assert not txt.startswith("1."), txt[:20]


# ------------------------------------------------------------ crossref/links
def test_hyperlink_wrappers_exist(root):
    """4 处交叉引用全部包成真 ``w:hyperlink``：3 图 + 1 文献。"""
    links = [h for h in root.iter(W + "hyperlink") if h.get(W + "anchor")]
    assert len(links) >= 4, "交叉引用没有真超链接：%d" % len(links)


def test_every_anchor_resolves_to_a_bookmark(root):
    names = bookmark_names(root)
    anchors = [h.get(W + "anchor") for h in root.iter(W + "hyperlink")
               if h.get(W + "anchor")]
    assert anchors, "没有内部锚点"
    dangling = [a for a in anchors if a not in names]
    assert not dangling, "悬挂引用: %s" % dangling


def test_bookmarks_cover_figs_and_refs(root):
    """3 图题注书签 + 3 条文献书签，另加每个题注一个 ``_Ref`` 锚点。"""
    names = bookmark_names(root)
    assert {"fig_vt", "fig_force", "fig_circuit"} <= names
    assert {"ref_1", "ref_2", "ref_3"} <= names
    # Captions also carry their own ``图_N`` / ``表_N`` bookmark and a
    # ``_Ref…`` whole-caption anchor; the answer-key grid stays uncaptioned.
    named = [n for n in names if not n.startswith("_Ref")]
    assert len(named) >= 6, sorted(names)
    assert sum(1 for n in names if n.startswith("_Ref")) == 4, sorted(names)


def test_hyperlink_wraps_a_ref_field(root):
    """wrapper 里必须包着 REF 域，否则跳转目标的缓存文本无从更新。"""
    ok = 0
    for h in root.iter(W + "hyperlink"):
        if not h.get(W + "anchor"):
            continue
        codes = " ".join((t.text or "") for t in h.iter(W + "instrText"))
        if "REF" in codes:
            ok += 1
    assert ok >= 4, "包着 REF 域的超链接只有 %d 个" % ok


def test_ref_field_switches(root):
    """``\\h`` 是可点开关；``\\n`` 让参考文献渲染成 ``[1]``。"""
    codes = " | ".join((t.text or "").strip()
                       for t in root.iter(W + "instrText"))
    assert r"\h" in codes, "缺 \\h 超链接开关"
    assert "REF ref_1" in codes
    assert r"\n" in codes, "参考文献引用没走 paragraph-number"


def test_captions_carry_seq_fields(root):
    """三个图题注各带一个 ``SEQ 图`` 计数器。"""
    seq = sum(1 for t in root.iter(W + "instrText")
              if (t.text or "").strip().startswith("SEQ 图"))
    assert seq == 3, "图题注 SEQ 域应为 3，实际 %d" % seq


def test_superscript_and_normal_both_present(root):
    """一处上角标（文献）+ 三处正文大小（引图）两种形态都要真落地。"""
    vals = [v.get(W + "val") for v in root.iter(W + "vertAlign")]
    assert "superscript" in vals, "没有上角标引用"
    links = [h for h in root.iter(W + "hyperlink") if h.get(W + "anchor")]
    plain = [h for h in links
             if not any(v.get(W + "val") == "superscript"
                        for v in h.iter(W + "vertAlign"))]
    assert len(plain) >= 3, "正文大小的引用只有 %d 处" % len(plain)


# ------------------------------------------------------------------ formulas
def test_omml_formulas_present(root):
    assert len(list(root.iter(M + "oMath"))) >= 1, "没有 OMML 公式"


def test_omath_count_matches_build(doc_xml):
    """建卷时 19 个公式片段，纯上下标的那些现在落成文字角标。

    混合策略（``docx_mathify.MIXED_SUPERSCRIPT``）把纯上下标片段写成
    ``w:vertAlign`` 的普通 run —— 这才是「选中按 Ctrl+Shift+= 会还原成正
    常大小」的形式。所以断言「仍有公式」+「片段总数没丢」，而不是固定 19。
    """
    n = doc_xml.count("<m:oMath>")
    assert n >= 1, "没有 OMML 公式"
    vert = doc_xml.count('w:vertAlign w:val="superscript"')
    vert_sub = doc_xml.count('w:vertAlign w:val="subscript"')
    assert vert + vert_sub > 0, "没有任何文字角标，混合策略没生效"
    assert n + vert + vert_sub >= 19, (
        "公式片段总数 %d 少于建卷时的 19" % (n + vert + vert_sub))


def test_unit_slash_survives_as_unit_not_fraction(root):
    """``m/s2`` / ``g/cm3`` / ``J/(mol.K)`` 必须留在正文里，不能变分式。

    量纲串可能留在 OMML 里（``m:t``），也可能落成普通 run 的铁文字
    （``w:t``，纯指数片段走这条）。两者都要读，否则会漏掉一半。
    """
    joined = "".join(t.text or "" for t in root.iter()
                     if t.tag in (W + "t", M + "t"))
    assert "/" in joined, "单位斜杠丢了"
    for unit in ("m/s", "g/cm", "J/(mol"):
        assert unit in joined, "量纲串 %s 丢了" % unit
    # 量纲串不该被塞进分式的分子/分母
    for f in root.iter(M + "f"):
        inner = "".join(t.text or "" for t in f.iter()
                        if t.tag in (W + "t", M + "t"))
        assert not any(u in inner for u in ("m/s", "g/cm", "J/(mol")), (
            "量纲串被误转成分式: %r" % inner)


def test_no_raw_latex_left_in_body(doc_xml):
    """没把 LaTeX 原文漏进正文。"""
    assert "\\frac" not in doc_xml
    assert "\\mathrm" not in doc_xml


# ----------------------------------------------------------- images / tables
def test_three_figures_embedded(part_names, root):
    media = [n for n in part_names if n.startswith("word/media/")]
    assert len(media) >= 3, "自产图不足 3 张: %s" % media
    assert len(list(root.iter(W + "drawing"))) >= 3


def test_figs_on_disk_are_real():
    figs = sorted((SUBJECT / "figs").glob("*.png"))
    assert len(figs) >= 3, figs
    for f in figs:
        assert f.stat().st_size > 10_000, "%s 像占位图" % f.name


def test_tables_present(root):
    """答题卡答案表 + 分值分布表 = 2 张。"""
    tables = list(root.iter(W + "tbl"))
    assert len(tables) >= 2, "表格只有 %d 张" % len(tables)
    rows = list(root.iter(W + "tr"))
    assert len(rows) == 8, "表格行数应为 8，实际 %d" % len(rows)


def test_answer_table_has_header(root):
    first = list(root.iter(W + "tbl"))[0]
    head = "".join(t.text or "" for t in first.iter(W + "t"))
    assert "题号" in head and "答案" in head


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
    """汉字题面的首行缩进必须写 ``w:firstLineChars``（Word 的字符缩进）。"""
    chars = [i for i in root.iter(W + "ind") if i.get(W + "firstLineChars")]
    assert chars, "没有任何段落写了 firstLineChars"


def test_option_indent_present(root):
    """A/B/C/D 选项靠 ``indent_left_mm`` 缩进。"""
    lefts = [int(i.get(W + "left")) for i in root.iter(W + "ind")
             if i.get(W + "left")]
    assert lefts and max(lefts) >= 200


def test_answer_rules_are_real_borders(root):
    """答题横线是 ``w:pBdr/w:bottom``，不是一串下划线字符。"""
    assert len(list(root.iter(W + "pBdr"))) >= 10


# -------------------------------------------------------------- text content
def test_required_text_present(doc_xml):
    for s in ("绝密★启用前", "注意事项", "第Ⅰ卷", "第Ⅱ卷",
              "参考答案与解析", "参考文献"):
        assert s in doc_xml, "正文缺少：%s" % s


def test_paper_has_two_exam_parts(spec):
    """第Ⅰ卷 / 第Ⅱ卷 都必须用真段落写上，且第Ⅱ卷另起一页。"""
    texts = [b.get("text", "") for b in spec["blocks"] if b["type"] == "paragraph"]
    assert any("第Ⅰ卷" in t for t in texts)
    assert any("第Ⅱ卷" in t for t in texts)
    assert any(b["type"] == "page_break" for b in spec["blocks"]), "第Ⅱ卷没另起一页"
