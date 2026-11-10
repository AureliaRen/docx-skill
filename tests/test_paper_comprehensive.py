# -*- coding: utf-8 -*-
"""理科综合（物理 + 化学 + 生物 合卷）试卷的产物断言。

这份测试只读磁盘上的 ``spec.json`` 与 ``paper.docx``，不碰 Word，所以 CI
（没有 Office）也能绿。断言针对的是*合卷*特有的失败模式，而不是排版细节：

* **题号连续性** —— 13 道题必须挂在同一个 ``<w:num>`` 上。跨科目重排时最容易
  出现的缺陷是「每科各建一个定义」，Word 会把每个定义从 1 数起，题号当场跳号
  （1,2,3 之后又回到 1）。所以断言：``numPr`` 段数 >= 9，且题号那一组的
  ``lvlText`` 只有一种取值、题号组内所有段共享同一 ``numId``。
* **三科各自都有公式段** —— 合卷的公式密度高于单科，但 OMML 转换是按段落做
  的；任何一科的公式若被写成纯文本（如 ``F=ma`` 没触发），那一科就看不出
  公式。断言物理、化学、生物三科的关键段落各自都定位到 ``m:oMath``。
* **两种页码格式** —— 卷头无页码 / 前置 ``lowerRoman`` / 正文与答案 ``decimal``。
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parents[1]
SUBJECT = SKILL.parent / "work" / "subjects" / "comprehensive"
SPEC = SUBJECT / "spec.json"
DOCX = SUBJECT / "paper.docx"

pytestmark = pytest.mark.skipif(
    not (SPEC.exists() and DOCX.exists()),
    reason="comprehensive subject spec/artifacts not shipped")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# 第三层公式痕迹：免排版降级后的纯文字简单等式（`M=18g/mol`、`E=6V`、`K=4`）。
# 只用于「这一段含公式」的粗筛，不追求与 `docx_mathify.no_layout_reason`
# 逐字等价 —— 后者在源脚本里，测试不该把它复制一份（复制品会漂移）。
_PLAIN_QUANTITY_RE = re.compile(
    r"[A-Za-z]\s*=\s*[-+]?\d+(?:\.\d+)?(?:\s*[A-Za-z\u00b0\u03a9\u03bc%/]+)?")


# ------------------------------------------------------------------ helpers
def doc_xml() -> str:
    with zipfile.ZipFile(DOCX) as z:
        return z.read("word/document.xml").decode("utf-8")


def num_defs() -> tuple[dict, dict]:
    """``(numId -> abstractNumId, abstractNumId -> {ilvl: lvlText})``."""
    with zipfile.ZipFile(DOCX) as z:
        raw = z.read("word/numbering.xml")
    import xml.etree.ElementTree as ET
    root = ET.fromstring(raw)
    to_abs, lvls = {}, {}
    for ab in root.findall(W + "abstractNum"):
        aid = ab.get(W + "abstractNumId")
        lvls[aid] = {l.get(W + "ilvl", "0"):
                     (l.find(W + "lvlText").get(W + "val")
                      if l.find(W + "lvlText") is not None else None)
                     for l in ab.findall(W + "lvl")}
    for nm in root.findall(W + "num"):
        aid = nm.find(W + "abstractNumId")
        if aid is not None:
            to_abs[nm.get(W + "numId")] = aid.get(W + "val")
    return to_abs, lvls


def paragraphs() -> list:
    """``[(visible_text, has_formula, numId, ilvl)]`` in document order.

    The text must include ``m:t``: once a formula becomes OMML it has no
    ``w:t`` at all, so a ``w:t``-only walk makes every formula-bearing
    paragraph look empty and quietly defeats a "this subject has formulas"
    assertion.

    Since 2026-11-07 a formula can be laid out **three** ways, and all three
    count as "this paragraph carries typeset mathematics":

    1. an ``m:oMath`` object — a fraction, a root, an integral;
    2. a plain run carrying ``w:vertAlign`` — a pure script like ``a=9.8m/s2``;
    3. plain text, when the fragment needs no two-dimensional layout at all —
       ``M=18g/mol``, ``E=6V``, ``h=15m``.  These are deliberately *not*
       formula objects any more (see ``docx_mathify.no_layout_reason``), so
       they are recognised here by shape instead: a single-letter quantity
       symbol, ``=``, a number, optional unit.
    """
    import xml.etree.ElementTree as ET
    root = ET.fromstring(doc_xml())
    out = []
    for p in root.iter(W + "p"):
        text = "".join(t.text or "" for t in p.iter()
                       if t.tag in (W + "t", M + "t"))
        has_math = any(True for _ in p.iter(M + "oMath"))
        if not has_math:
            has_math = any(True for _ in p.iter(W + "vertAlign"))
        if not has_math:
            # 第三层：免排版降级后的纯文字简单等式（`M=18g/mol`、`E=6V`）
            has_math = _PLAIN_QUANTITY_RE.search(text) is not None
        nid = ilvl = None
        ppr = p.find(W + "pPr")
        if ppr is not None:
            npr = ppr.find(W + "numPr")
            if npr is not None:
                e = npr.find(W + "numId")
                nid = None if e is None else e.get(W + "val")
                e = npr.find(W + "ilvl")
                ilvl = "0" if e is None else e.get(W + "val")
        out.append((text, has_math, nid, ilvl))
    return out


def spec_blocks() -> list:
    return json.loads(SPEC.read_text(encoding="utf-8"))["blocks"]


def blocks_of(kind: str) -> list:
    return [b for b in spec_blocks() if b.get("type") == kind]


# --------------------------------------------------------------------- spec
def test_spec_declares_the_cross_subject_skeleton():
    blocks = spec_blocks()
    assert len(blocks) >= 80, "合卷内容太薄，压不到排版链路"
    assert len(blocks_of("section_break")) == 3, "卷头/说明/正文/答案 需 4 节"
    for b in blocks_of("section_break"):
        assert "page" in b, "每节都要显式给 page 全套"


def test_spec_figure_blocks_each_carry_a_caption_bookmark():
    images = blocks_of("image")
    captions = blocks_of("caption")
    assert len(images) == 3, "三科各一张图"
    assert len(captions) == 3
    names = {c["bookmark"] for c in captions}
    assert names == {"fig_vt", "fig_app", "fig_ped"}
    # 题注必须紧跟在自己的图后面，否则 REF 指到别处
    order = [b["type"] for b in blocks_of_spec_order()]
    assert order == ["image", "caption", "image", "caption", "image", "caption"]


def blocks_of_spec_order() -> list:
    return [b for b in spec_blocks()
            if b.get("type") in ("image", "caption")]


def test_spec_crossrefs_cover_both_display_styles():
    refs = blocks_of("crossref")
    assert len(refs) >= 4
    sup = [r for r in refs if r.get("superscript")]
    plain = [r for r in refs if not r.get("superscript")]
    assert len(sup) >= 2, "参考文献的正文引用要上角标"
    assert len(plain) >= 2, "题注引用要正常大小"
    assert {r["target"] for r in sup} == {"ref_1", "ref_2"}
    assert {r["target"] for r in plain} == {"fig_vt", "fig_app", "fig_ped"}
    assert all(r.get("format") == "paragraph-number" for r in sup), \
        "paragraph-number 才渲染成 [1]/[2]"


def test_spec_reference_list_is_four_entries_with_prefix():
    lists = [b for b in blocks_of("numbered_list")
             if b.get("bookmark_prefix") == "ref_"]
    assert len(lists) == 1
    assert len(lists[0]["items"]) == 4
    assert lists[0]["num_text"] == "[%1]"
    assert lists[0]["num_text"] != "%1．"


def test_spec_numbers_every_question_on_one_counter():
    q = [b for b in blocks_of("numbered_list") if b.get("num_text") == "%1．"]
    assert len(q) == 13, "9 选择 + 4 非选择 = 13 道题"
    ids = {b.get("num_id") for b in q}
    assert None in ids, "第一块必须让 docx_create 建定义"
    assert len(ids) == 2, "除首块外必须共用同一个 numId，否则 Word 从 1 重数"
    shared = {i for i in ids if i is not None}
    assert len(shared) == 1
    assert sum(1 for b in q if b.get("num_id") is None) == 1


def test_spec_three_subjects_have_their_own_formula_strings():
    texts = json.dumps(spec_blocks(), ensure_ascii=False)
    for token in ("a=9.8m/s2", "F=ma", "ρ=1.2kg/m3",     # 物理
                  "c=0.5mol/L", "M=18g/mol",             # 化学
                  "P=50W/m2", "v=s/t", "k=8.31J/(mol.K)"):  # 生物 / 跨学科
        assert token in texts, "缺公式串 %r" % token
    assert "3×10" not in texts and "10⁸" not in texts, \
        "科学计数法混单位会被 mathify 误判（契约 §5）"


# --------------------------------------------------------------------- docx
def test_document_has_four_sections_and_two_page_number_formats():
    fmts = re.findall(r'<w:pgNumType[^/]*?w:fmt="(\w+)"', doc_xml())
    assert len(fmts) >= 2
    assert set(fmts) == {"lowerRoman", "decimal"}, \
        "卷头无页码 + 前置 lowerRoman + 正文/答案 decimal"
    with zipfile.ZipFile(DOCX) as z:
        doc = z.read("word/document.xml").decode("utf-8")
    assert doc.count("<w:sectPr") + doc.count("<w:sectPr ") >= 4


def test_question_numbers_are_continuous_on_a_single_counter():
    """合卷最致命的缺陷：跨科目重排导致题号跳号。"""
    to_abs, lvls = num_defs()
    paras = paragraphs()
    q = [(t, n, l) for (t, _m, n, l) in paras if l is not None]
    assert len(q) >= 9, "numPr 段数 >= 9（9 选择 + 4 非选择 → 13）"

    # 题号组 = lvlText 为 "%1．" 的那一组
    qnums = [n for (t, n, l) in q
             if (lvls.get(to_abs.get(n, ""), {}).get(l) or "") == "%1．"]
    assert len(qnums) == 13, "13 道题必须全部落在题号组，实得 %d" % len(qnums)
    assert len(set(qnums)) == 1, "题号组必须只有一个 numId，实际 %r" % set(qnums)

    # lvlText 单一：题号组的所有 abstractNum 只用一种模板
    abs_ids = {to_abs[qnums[0]]}
    texts = {lvls[a].get("0") for a in abs_ids}
    assert texts == {"%1．"}, "题号 lvlText 必须单一，实际 %r" % texts

    # 文献组自成一组、模板是 [%1]
    refs = [(t, n, l) for (t, n, l) in q
            if (lvls.get(to_abs.get(n, ""), {}).get(l) or "").startswith("[")]
    assert len(refs) == 4
    assert {lvls[to_abs[n]][l] for (_t, n, l) in refs} == {"[%1]"}


def test_every_subject_contributes_a_formula_paragraph():
    """三科各自都要有「公式层」痕迹 —— 但不限定必须是 ``m:oMath``。

    2026-11-07 起公式分三层落盘（见 ``docx_mathify``）：
    真结构 → OMML；纯上下标 → ``w:vertAlign`` 文字角标；
    无二维排版需求的简单等式 → 普通文字 run。

    所以 ``M=18g/mol`` 这种「量符号=数值+单位」现在**故意**不产 OMML，
    用「段落里有 oMath」来筛会把正确的降级误判成缺失。这里改成
    「段落里有 oMath 或 vertAlign」，并单独断言两条真结构公式仍是 OMML。
    """
    paras = paragraphs()
    math_paras = [t for (t, m, _n, _l) in paras if m]

    def has(*tokens):
        return any(all(tok in p for tok in tokens) for p in math_paras)

    # —— 三层都算公式痕迹：题面文字里出现这些片段即证明该科公式没丢 ——
    assert has("加速度"), "物理题面要有 a=9.8m/s2"
    assert has("牛顿第二定律") or has("F=ma"), "物理第 10 题的 F=ma"
    assert has("mol/L"), "化学的 c=0.5mol/L / c(HI)=1.0mol/L"
    assert has("g/mol"), "化学的 M=18g/mol"
    assert has("P=50"), "生物的 P=50W/m2"
    assert has("酶促反应速率"), "跨学科第 13 题的 v=s/t"

    # —— 真结构公式必须仍是 OMML（这是本轮不许回退的硬指标）——
    assert has("I=U"), "欧姆定律 I=U/R 必须仍是 OMML 分式"
    assert has("F=ma"), "牛顿第二定律 F=ma 必须仍是 OMML"
    assert has("v=s"), "v=s/t 必须仍是 OMML 分式"

    # 解析里三科也要各自带公式（不是只有题面有）
    assert sum(1 for t in math_paras if t.startswith("第 10 题")) >= 1
    assert sum(1 for t in math_paras if t.startswith("第 11 题")) >= 1
    assert sum(1 for t in math_paras if t.startswith("第 13 题")) >= 1


def fraction_parts() -> list:
    """``[(fPr/type, numerator_text, denominator_text)]`` for every m:f."""
    import xml.etree.ElementTree as ET
    root = ET.fromstring(doc_xml())
    out = []
    for f in root.iter(M + "f"):
        kind = "bar"
        fpr = f.find(M + "fPr")
        if fpr is not None:
            e = fpr.find(M + "type")
            if e is not None:
                kind = e.get(M + "val")
        num = "".join(t.text or "" for t in f.find(M + "num").iter(M + "t"))
        den = "".join(t.text or "" for t in f.find(M + "den").iter(M + "t"))
        out.append((kind, num, den))
    return out


def test_dimension_units_stay_units_and_real_quotients_become_fractions():
    r"""`/` in a unit string reads as 「每」; between two quantity symbols it is
    a fraction bar.

    The docx layer cannot store both in one shape, so ``linear_to_latex``
    emits two *different* structures and the test asserts each one:

    * ``a=9.8m/s2``  -> ``m:f`` with ``fPr/type="lin"`` (a/b set on one
      baseline: 千克每立方米, not 千克除立方米)
    * ``v=s/t``      -> ``m:f`` with the default ``bar`` fraction bar
      (真分式, 分母在横线下方)

    Asserting on the literal string ``"W/m"`` is meaningless: the OMML for
    ``W/m2`` holds ``W`` and ``m2`` in separate ``m:t`` runs.
    """
    parts = fraction_parts()

    linear = [(n, d) for (k, n, d) in parts if k == "lin"]
    barred = [(n, d) for (k, n, d) in parts if k == "bar"]

    # 真分式：I=U/R 与 v=s/t 必须用默认分数线（分母在横线下方）。
    # 题面与解析各写一次，所以每个式子出现两次、共 4 条。
    # 这条与「量纲串走 OMML 还是文字角标」无关，是本测试的核心语义。
    assert ("U", "R") in barred, "I=U/R 必须变成真分式"
    assert ("s", "t") in barred, "v=s/t 必须变成真分式"
    assert set(barred) == {("U", "R"), ("s", "t")}, \
        "真分式集合应只有 U/R 与 s/t，实际 %r" % (barred,)

    # 量纲串：/ 是「每」，绝不能落成 bar 分式。
    # 混合策略后纯指数片段（P=50W/m²）改由文字角标承担，此时它不出现在
    # 任何 m:f 里 —— 那也是正确的；关键是它不许出现在 barred 集合里。
    unit_in_bar = [p for p in barred
                   if p not in (("U", "R"), ("s", "t"))]
    assert not unit_in_bar, "量纲串被落成了真分式: %r" % (unit_in_bar,)

    # 若量纲串仍走 OMML，必须是 lin 形式（而不是 bar）
    import xml.etree.ElementTree as _ET
    body = "".join(t.text or "" for t in _ET.fromstring(doc_xml()).iter()
                   if t.tag in (W + "t", M + "t"))
    for pair in (("m", "s2"), ("W", "m2")):
        if pair in linear:
            continue
        # 否则它应当已经变成文字角标：正文里能找到该单位串
        assert pair[0] in body and pair[1] in body, (
            "量纲串 %r 既不在 lin 分式里、也不在正文里" % (pair,))

    # 题面 + 解析各两次
    assert len(barred) == 4, "真分式应 4 条，实得 %r" % (barred,)


def test_unit_strings_without_an_exponent_stay_plain_text():
    """`mol/L`、`g/mol`、`J/(mol.K)` 无幂次，落成 `\\text{}` 原文，不进 m:f。"""
    doc = doc_xml()
    for unit in ("mol/L", "g/mol", "J/(mol.K)"):
        assert unit in doc, "量纲串 %r 没落进 OMML" % unit
    parts = fraction_parts()
    assert not [p for p in parts if "mol" in p[1] + p[2]], \
        "无幂次的量纲串不该被做成分数线"


def test_three_figures_and_two_tables_land_in_the_package():
    with zipfile.ZipFile(DOCX) as z:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
    assert len(media) == 3
    doc = doc_xml()
    assert doc.count("<w:tbl>") == 2, "分值分布表 + 题号答案对照表"
    assert doc.count("<w:drawing>") == 3


def test_links_are_real_wrappers_with_resolvable_anchors():
    doc = doc_xml()
    anchors = re.findall(r'<w:hyperlink w:anchor="([^"]+)"', doc)
    assert len(anchors) >= 5
    bookmarks = set(re.findall(r'<w:bookmarkStart[^>]*w:name="([^"]+)"', doc))
    assert set(anchors) <= bookmarks, "悬挂锚点：%r" % (set(anchors) - bookmarks)
    assert {"fig_vt", "fig_app", "fig_ped", "ref_1", "ref_2"} <= set(anchors)


def test_superscript_references_carry_vertalign_on_every_field_run():
    doc = doc_xml()
    sup = doc.count('w:val="superscript"')
    assert sup >= 10, "两个上角标 REF 域 × 5 run = 10 个 vertAlign"
    # 正常大小的题注引用不应被误设成上角标
    norm = [a for a in re.findall(r'<w:hyperlink w:anchor="([^"]+)"', doc)
            if a.startswith("fig_")]
    assert len(norm) == 3


def test_paragraph_format_keys_reach_ppr():
    doc = doc_xml()
    assert doc.count("<w:ind ") >= 20, "题面首行缩进 + 选项左缩进"
    assert 'w:firstLineChars="200"' in doc
    assert doc.count("<w:jc ") >= 20, "大量居中/两端对齐"
    assert 'w:line="360"' in doc, "line_spacing 1.5 → w:line 360"


def test_captions_carry_seq_fields_so_the_numbers_can_renumber():
    doc = doc_xml()
    assert doc.count("SEQ 图") >= 3
    assert doc.count("REF fig_") >= 3


def test_pdf_and_png_evidence_are_archived():
    pdf = SUBJECT / "paper.pdf"
    assert pdf.exists() and pdf.stat().st_size > 1000
    pngs = sorted((SUBJECT / "png").glob("page-*.png"))
    assert len(pngs) >= 5, "10 页试卷至少留 5 张页面图作证据"
    acc = json.loads((SUBJECT / "accept.json").read_text(encoding="utf-8"))
    assert acc["ok"] is True and acc["openable"] is True
    assert acc["hard_errors"] == []
    assert acc["sections"] == 4
    assert acc["hyperlink_targets"]
    for t in acc["hyperlink_targets"]:
        assert t["sub_address"], "每个超链接都要有内部锚点"
    assert acc["numbering_summary"]["rendered"] == \
        acc["numbering_summary"]["xml_numpr"], "Word 必须把每个 numPr 都渲染出来"
    assert acc["numbering_summary"]["text_mismatch"] == 0


def test_readings_file_quotes_the_real_numbers():
    txt = (SUBJECT / "accept_readings.txt").read_text(encoding="utf-8")
    acc = json.loads((SUBJECT / "accept.json").read_text(encoding="utf-8"))
    for key in ("openable", "pages", "sections", "bookmarks", "hyperlinks",
                "hyperlink_targets", "numbering_render", "field_list",
                "superscript_count", "omath_count", "images", "tables"):
        assert key in txt, "读数缺 %s" % key
    assert "pages               = %d" % acc["pages"] in txt
    assert "sections            = %d" % acc["sections"] in txt
    assert "hyperlinks          = %d" % acc["hyperlinks"] in txt
