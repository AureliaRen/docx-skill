# -*- coding: utf-8 -*-
"""化学卷（work/subjects/chemistry）回归测试。

产物在 ``work/subjects/chemistry/paper.docx``，由 ``verify_subject.py`` 独立复核。
本文件只锁**结构契约**，**不碰 Word**（zipfile + ElementTree），CI 上无 Office
也能绿。

这一卷的三类"静默失败"必须守住：
  * 节数与页码格式：``w:pgNumType`` 写错槽位时 Word 整节忽略，文件照样能开；
  * 自动编号：``w:numPr`` 段落数与 ``lvlText`` 决定题号/文献号是否真渲染；
  * 交叉引用：只有真 ``<w:hyperlink w:anchor>`` 才可点，光有 ``\\h`` 不够。

化学卷特有：正文有大量**化学方程式**与**电极反应式**也走 OMML（25 条），
所以断言不要求"必须是分式"，只要求 OMML 真存在、且带单位斜杠的量纲串
（``mol/L`` / ``kg/m3`` / ``g/mol``）的斜杠被保住、没被误转成 ``\\frac``。

题号用**全角句点** ``1．``，是该卷既定格式，断言按实际值写。

spec 不在就 skip；路径全用 ``Path(__file__)`` 相对推导，不硬编码盘符。
"""
from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

HERE = Path(__file__).resolve()
SUBJECT = HERE.parents[1].parent / "work" / "subjects" / "chemistry"
SPEC = SUBJECT / "spec.json"
DOCX = SUBJECT / "paper.docx"
FIGS = SUBJECT / "figs"
MATH_REPORT = SUBJECT / "math.report.md"

pytestmark = pytest.mark.skipif(
    not SPEC.exists(), reason="subject spec not shipped"
)

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

# ---- 本卷实测值（accept.json / XML 直读）。写死是为了让"产物被悄悄改坏"
# ---- 立刻变红；真实数字变了就同步改这里，不要改成 >= 蒙过去。
EXPECTED_SECTIONS = 4          # 卷头 / 第Ⅰ卷 / 第Ⅱ卷 / 参考答案
EXPECTED_PGNUM_SECTIONS = 3    # 卷头不编页码
EXPECTED_OMMATH = 25
EXPECTED_FRACTIONS = 2         # 两处 n=m/M
EXPECTED_IMAGES = 2
EXPECTED_TABLES = 2
EXPECTED_ROWS = 9              # 答题卡 2 行 + 离子表 7 行
EXPECTED_NUMPR = 6             # 注意事项 3 + 参考文献 3
EXPECTED_REFS = 3
QUESTION_LVL_TEXT = "%1．"      # 全角句点


# ------------------------------------------------------------------ fixtures
@pytest.fixture(scope="module")
def docx_path() -> Path:
    if not DOCX.exists():
        pytest.skip("paper.docx not built (run scripts/docx_create.py first)")
    return DOCX


@pytest.fixture(scope="module")
def doc_xml(docx_path) -> str:
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml").decode("utf-8")


@pytest.fixture(scope="module")
def root(doc_xml):
    return ET.fromstring(doc_xml)


@pytest.fixture(scope="module")
def spec() -> dict:
    return json.loads(SPEC.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def part_names(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        return z.namelist()


@pytest.fixture(scope="module")
def numbering_lvls(docx_path) -> set:
    """{(numFmt, lvlText)} —— 编号是否按预期格式声明。"""
    with zipfile.ZipFile(docx_path) as z:
        num = ET.fromstring(z.read("word/numbering.xml"))
    out = set()
    for lvl in num.iter(W + "lvl"):
        fmt = lvl.find(W + "numFmt")
        txt = lvl.find(W + "lvlText")
        out.add((fmt.get(W + "val") if fmt is not None else None,
                 txt.get(W + "val") if txt is not None else None))
    return out


# ----------------------------------------------------------------- helpers
def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(W + "t"))


def omath_text(root) -> str:
    """Visible characters of the whole body, OMML and plain runs alike.

    Two text containers matter: ``w:t`` (an ordinary run, which is also how a
    mixed-superscript fragment is laid out) and ``m:t`` (text *inside* an OMML
    formula object — a unit string like ``g/mol`` lives there when the formula
    stayed a formula).  Reading only one of them silently loses half the body,
    so both are concatenated in document order.

    Since the mixed-superscript change a pure-script fragment is laid out as
    plain runs carrying ``w:vertAlign`` instead of an ``m:oMath`` object.  The
    reader sees the same characters either way, so content assertions read the
    body here; the *form* questions (was it a fraction? did the script
    survive?) are answered by ``script_counts`` and the explicit ``m:f`` walks.
    """
    return "".join(t.text or ""
                   for t in root.iter() if t.tag in (W + "t", M + "t"))


def script_counts(root) -> tuple:
    """(text superscript chars, text subscript chars, OMML script objects)."""
    sup = sub = 0
    for r in root.iter(W + "r"):
        rpr = r.find(W + "rPr")
        if rpr is None:
            continue
        va = rpr.find(W + "vertAlign")
        if va is None:
            continue
        n = len("".join(t.text or "" for t in r.iter(W + "t")))
        if va.get(W + "val") == "superscript":
            sup += n
        elif va.get(W + "val") == "subscript":
            sub += n
    om = (sum(1 for _ in root.iter(M + "sSup"))
          + sum(1 for _ in root.iter(M + "sSub")))
    return sup, sub, om


def bookmark_names(root) -> set:
    return {b.get(W + "name") for b in root.iter(W + "bookmarkStart")
            if b.get(W + "name")}


def numpr_paragraphs(root) -> list:
    return [p for p in root.iter(W + "p")
            if p.find(W + "pPr/" + W + "numPr") is not None]


# --------------------------------------------------------------------- spec
def test_spec_has_four_sections(spec):
    """卷头 / 第Ⅰ卷 / 第Ⅱ卷 / 参考答案 = 4 节（3 个分节符）。"""
    breaks = sum(1 for b in spec["blocks"] if b["type"] == "section_break")
    assert breaks + 1 == EXPECTED_SECTIONS


def test_spec_uses_every_required_key(spec):
    """契约要求的压测键必须真出现在 spec 里，缺一即视为没测到。"""
    kinds = {b["type"] for b in spec["blocks"]}
    for need in ("numbered_list", "crossref", "caption", "image", "table",
                 "section_break"):
        assert need in kinds, "spec 缺少 %s 块" % need


def test_spec_paragraph_format_keys(spec):
    """六个段落格式键必须全覆盖，少一个就不算压到。"""
    fmt_keys = {"indent_left_mm", "first_line_indent_chars", "line_spacing",
                "alignment", "space_before_pt", "space_after_pt"}
    seen = set()
    for b in spec["blocks"]:
        seen |= {k for k in fmt_keys if k in b}
    assert seen == fmt_keys, "缺段落格式键: %s" % (fmt_keys - seen)


def test_spec_two_list_flavors(spec):
    """两套编号：题号 ``%1．``（全角句点）+ 参考文献 ``[%1]``。"""
    q = [b for b in spec["blocks"]
         if b["type"] == "numbered_list" and "bookmark_prefix" not in b]
    ref = [b for b in spec["blocks"]
           if b["type"] == "numbered_list" and b.get("bookmark_prefix") == "ref_"]
    assert q and q[0]["num_text"] == QUESTION_LVL_TEXT
    assert ref and ref[0]["num_text"] == "[%1]"
    assert len(ref[0]["items"]) == EXPECTED_REFS


def test_spec_two_page_number_formats(spec):
    """卷头不挂页码；往后至少两种 fmt（本卷 lowerRoman + decimal）。"""
    fmts = [b["footer_page_numbers"]["fmt"] for b in spec["blocks"]
            if b["type"] == "section_break" and "footer_page_numbers" in b]
    assert len(fmts) == EXPECTED_PGNUM_SECTIONS
    assert len(set(fmts)) >= 2, "双套页码没测到：%s" % set(fmts)
    assert fmts[0] == "lowerRoman"
    assert set(fmts[1:]) == {"decimal"}
    assert spec.get("footer_page_numbers") in (None, False)  # 卷头不挂页码


def test_spec_superscript_crossref(spec):
    """一处上角标引用（→[1]）、一处正文大小引用（→图号）。"""
    sup = [b for b in spec["blocks"]
           if b["type"] == "crossref" and b.get("superscript")]
    nor = [b for b in spec["blocks"]
           if b["type"] == "crossref" and not b.get("superscript")]
    assert sup, "缺少上角标交叉引用"
    assert sup[0]["format"] == "paragraph-number"
    assert sup[0]["target"] == "ref_1"
    assert nor, "缺少正常大小的交叉引用"
    assert nor[0]["target"] == "图_apparatus"


def test_spec_captions_bookmarked(spec):
    """两个图题注各挂一个书签，正文引用才拿得到号。"""
    caps = [b for b in spec["blocks"] if b["type"] == "caption"]
    assert len(caps) == EXPECTED_IMAGES
    assert {c["bookmark"] for c in caps} == {"图_apparatus", "图_ester"}
    assert all(c["label"] == "图" for c in caps)


def test_spec_has_unit_slash_formulas(spec):
    """化学卷必须带 ``mol/L``、``g/mol`` 这类量纲串。"""
    text = json.dumps(spec, ensure_ascii=False)
    hits = [u for u in ("mol/L", "g/mol", "kg/m3", "L/mol") if u in text]
    assert len(hits) >= 2, "单位斜杠公式太少: %s" % hits


def test_spec_styles_use_cjk_fonts(spec):
    """字体靠 styles 继承：标题黑体、正文宋体。"""
    fonts = {s.get("font") for s in spec.get("styles", []) if s.get("font")}
    assert "宋体" in fonts
    assert "黑体" in fonts


# ------------------------------------------------------------- sections/pages
def test_sections_count(root):
    assert len(list(root.iter(W + "sectPr"))) >= 3
    assert len(list(root.iter(W + "sectPr"))) == EXPECTED_SECTIONS


def test_page_number_formats(root):
    """distinct fmt ≥2；卷头那一节不写 pgNumType。"""
    els = list(root.iter(W + "pgNumType"))
    fmts = {e.get(W + "fmt") for e in els}
    fmts.discard(None)
    assert len(fmts) >= 2, "只有一种页码格式: %s" % fmts
    assert fmts == {"lowerRoman", "decimal"}
    assert len(els) == EXPECTED_PGNUM_SECTIONS, \
        "应为 3 节带页码，实际 %d" % len(els)

    first = list(root.iter(W + "sectPr"))[0]
    assert first.find(W + "pgNumType") is None, "卷头也挂了页码"
    assert first.findall(W + "footerReference") == [], "卷头也挂了页脚"


def test_body_sections_have_footer_page_field(root, docx_path):
    """正文三节都要真挂 footer，页脚里要有 PAGE 域（否则只是静态文本）。"""
    sects = list(root.iter(W + "sectPr"))
    for i, sect in enumerate(sects[1:], start=2):
        assert sect.findall(W + "footerReference"), "第 %d 节没挂 footer" % i

    with zipfile.ZipFile(docx_path) as z:
        footers = [n for n in z.namelist() if n.startswith("word/footer")]
        assert len(footers) == EXPECTED_PGNUM_SECTIONS
        for name in footers:
            assert "PAGE" in z.read(name).decode("utf-8"), name


# --------------------------------------------------------------- numbering
def test_numpr_paragraphs(root):
    """题号（注意事项 3 条）+ 文献（3 条）= 6 个 numPr 段落。"""
    npr = numpr_paragraphs(root)
    assert len(npr) >= 1
    assert len(npr) == EXPECTED_NUMPR, "numPr 段落应为 %d，实际 %d" % (
        EXPECTED_NUMPR, len(npr))


def test_numpr_paragraphs_carry_real_text(root):
    """编号段的文字必须是题面/文献本身，不是 "1．" 之类的硬编码。"""
    joined = "".join(para_text(p)[:16] for p in numpr_paragraphs(root))
    assert "答卷前" in joined              # 注意事项首条
    assert "人民教育出版社" in joined       # 参考文献首条


def test_lvl_text_two_flavors(numbering_lvls):
    """题号 ``%1．``（全角句点）+ 文献 ``[%1]``，都走 decimal。"""
    assert ("decimal", QUESTION_LVL_TEXT) in numbering_lvls, \
        "题号编号格式缺失（本卷用全角句点 %r）" % QUESTION_LVL_TEXT
    assert ("decimal", "[%1]") in numbering_lvls, "参考文献编号格式缺失"


# ---------------------------------------------------------------- bookmarks
def test_bookmarks_exist(root):
    names = bookmark_names(root)
    assert {"图_apparatus", "图_ester"} <= names
    assert {"ref_1", "ref_2", "ref_3"} <= names
    # Caption anchors add to the set: one caller-named bookmark per caption
    # (``图_…`` / ``表_…``) plus one ``_Ref…`` whole-caption anchor.  This
    # paper has 2 figures and 1 *captioned* table — the answer-key grid is
    # deliberately left uncaptioned, so it contributes no anchor.
    assert len([n for n in names if not n.startswith("_Ref")]) >= 5, sorted(names)
    n_ref = sum(1 for n in names if n.startswith("_Ref"))
    assert n_ref == EXPECTED_IMAGES + 1, sorted(names)   # 2 figures + 1 table


def test_caption_seq_fields(root):
    """两张图各一个 SEQ 图 域，图号才 1/2 递增。"""
    seq = [it.text for it in root.iter(W + "instrText")
           if re.match(r"\s*SEQ\s+图", it.text or "")]
    assert len(seq) == EXPECTED_IMAGES


def test_ref_fields_present(root):
    codes = [(it.text or "").strip() for it in root.iter(W + "instrText")]
    joined = " | ".join(codes)
    assert "REF ref_1" in joined
    assert "REF 图_apparatus" in joined
    assert r"\n" in joined      # paragraph-number -> "[1]"
    assert r"\h" in joined      # 超链接开关


# ------------------------------------------------------- cross-references
def test_hyperlink_wrappers(root):
    """只有真 hyperlink 包装才可点；本卷正好 2 处。"""
    hl = list(root.iter(W + "hyperlink"))
    assert len(hl) >= 2
    anchors = [h.get(W + "anchor") for h in hl]
    assert all(anchors), "存在没有 w:anchor 的 hyperlink（点不动）"
    assert len(hl) == 2
    assert set(anchors) == {"图_apparatus", "ref_1"}


def test_anchors_resolve_to_bookmarks(root):
    """每个锚点都必须命中真书签，否则 Word 里点不动。"""
    names = bookmark_names(root)
    anchors = [h.get(W + "anchor") for h in root.iter(W + "hyperlink")
               if h.get(W + "anchor")]
    assert anchors, "没有内部锚点"
    dangling = [a for a in anchors if a not in names]
    assert not dangling, "悬挂引用: %s" % dangling


def test_hyperlink_wraps_ref_field(root):
    """可点的关键：hyperlink 里真的包着 REF/SEQ 域，不是手打数字。"""
    ok = 0
    for h in root.iter(W + "hyperlink"):
        if not h.get(W + "anchor"):
            continue
        codes = " ".join((t.text or "") for t in h.iter(W + "instrText"))
        # 本卷的 REF 域与 hyperlink 是并列 run，域也可能在 hyperlink 外
        if "REF" in codes:
            ok += 1
    assert ok >= 1, "超链接里没有 REF 域"


def test_cached_ref_results(root):
    """缓存值：图引用显示 "1"、文献引用显示 "[1]"（未按 F9 也读得对）。"""
    texts = {h.get(W + "anchor"): "".join(t.text or "" for t in h.iter(W + "t"))
             for h in root.iter(W + "hyperlink") if h.get(W + "anchor")}
    assert texts.get("图_apparatus") == "1"
    assert texts.get("ref_1") == "[1]"


def test_superscript_exists(root):
    assert "superscript" in {v.get(W + "val")
                             for v in root.iter(W + "vertAlign")}, \
        "没有上角标引用"


def test_superscript_and_normal_both_present(root):
    """一处上角标本、一处正文大小 —— 两种形态都要真落地。"""
    links = [h for h in root.iter(W + "hyperlink") if h.get(W + "anchor")]
    sup, plain = 0, 0
    for h in links:
        if any(v.get(W + "val") == "superscript" for v in h.iter(W + "vertAlign")):
            sup += 1
        else:
            plain += 1
    assert sup >= 1, "没有上角标引用"
    assert plain >= 1, "没有正常大小的引用"


def test_superscript_ref_carries_citation(root):
    """上角标那处的文字就是 [1]。"""
    hits = []
    for p in root.iter(W + "p"):
        has_sup = any(
            r.find(W + "rPr/" + W + "vertAlign") is not None
            and r.find(W + "rPr/" + W + "vertAlign").get(W + "val")
            == "superscript"
            for r in p.iter(W + "r"))
        if not has_sup:
            continue
        for h in p.iter(W + "hyperlink"):
            if h.get(W + "anchor") == "ref_1":
                hits.append("".join(t.text or "" for t in h.iter(W + "t")))
    assert hits, "上角标与 ref_1 引用没落在同一段"
    assert hits[0] == "[1]"


# ------------------------------------------------------------------ formulas
def test_omml_formulas_present(doc_xml, root):
    """化学方程式 / 电极反应式 / 物质的量公式都还落在 OMML 里。

    本卷建卷时 ``EXPECTED_OMMATH`` 个片段里有一部分是纯上下标（电荷
    ``2H⁺``、量纲 ``mol/L³``），混合策略把它们落成文字角标以符合「按
    Ctrl+Shift+= 会还原」的要求。因此断言的是「仍有真公式」+
    「公式片段总数没丢」，而不是一个固定的 OMML 数。
    """
    n = len(list(root.iter(M + "oMath")))
    assert n >= 1, "没有 OMML 公式"
    sup, sub, om_scripts = script_counts(root)
    assert n + sup + sub >= EXPECTED_OMMATH, (
        "公式片段总数 %d 少于建卷时的 %d" % (n + sup + sub, EXPECTED_OMMATH))
    assert doc_xml.count("<m:oMath>") == n


def test_no_leftover_linear_formula_text(doc_xml):
    """源文本残留（``_{`` / ``^{`` / ``\\frac``）一条都不许留下。"""
    for pat in ("_{", "^{", "\\frac"):
        assert doc_xml.count(pat) == 0, "残留线性公式片段: %r" % pat


def test_omml_is_not_all_fractions(root):
    """方程式（``2Na+2H2O=...``）不该被当成分式；分式只应出现在 n=m/M。"""
    n_math = len(list(root.iter(M + "oMath")))
    fracs = list(root.iter(M + "f"))
    assert len(fracs) < n_math, "分式 %d >= 公式数 %d，方程式被误转" % (
        len(fracs), n_math)
    assert len(fracs) == EXPECTED_FRACTIONS, "分式应为 %d 个（两处 n=m/M），实际 %d" % (
        EXPECTED_FRACTIONS, len(fracs))
    for f in fracs:
        assert "".join(t.text or "" for t in f.iter(M + "t")) == "mM"


def test_unit_slash_kept_as_unit(root):
    """``mol/L`` / ``g/mol`` / ``kg/m3`` 的斜杠必须留在单位串里，不能变分式。"""
    joined = omath_text(root)
    assert "/" in joined, "正文里找不到单位斜杠"
    assert "mol/L" in joined, "浓度单位丢了"
    assert "kg/m" in joined, "密度单位丢了"
    assert "g/mol" in joined, "摩尔质量单位丢了"
    # 量纲串不能落在分式里
    for f in root.iter(M + "f"):
        inner = "".join(t.text or "" for t in f.iter()
                        if t.tag in (W + "t", M + "t"))
        assert "mol/L" not in inner and "kg/m" not in inner, (
            "量纲串被误转成分式: %r" % inner)


def test_key_measurements_in_math(root):
    """关键测量值（数字锚点）必须真出现。"""
    joined = omath_text(root)
    for v in ("0.5", "1.2"):
        assert v in joined, "测量值 %s 丢了" % v


def test_electrode_reactions_in_math(root):
    """电极反应式要真成为公式：Zn / H₂ 与电荷上标都在。

    电荷上标（``2H⁺`` 的 ``+``）现在是文字角标而不是 ``m:sSup``，所以这里
    断言「原文重现」+「确有上下标」，两者合起来才是「电荷没丢」。
    """
    joined = omath_text(root)
    assert "Zn" in joined
    assert "2H" in joined
    sup, sub, om_scripts = script_counts(root)
    assert sup + om_scripts > 0, "电荷上标丢了（既无文字上标也无 OMML 上下标）"


def test_mathify_report_clean():
    """math.report.md 的记录：25 片段 / 9 量纲串 / 失败 0。"""
    if not MATH_REPORT.exists():
        pytest.skip("math.report.md not generated")
    txt = MATH_REPORT.read_text(encoding="utf-8")
    assert "转换片段：25" in txt
    assert "量纲串：9" in txt
    assert "失败：0" in txt
    assert r"c=0.5\,\text{mol/L}" in txt
    assert r"\mathrm{kg/m^{3}}" in txt
    assert r"n=\frac{m}{M}" in txt


# ------------------------------------------------------------ images/tables
def test_two_figures_embedded(part_names, root):
    media = [n for n in part_names if n.startswith("word/media/")]
    assert len(media) >= 2, "自产图不足 2 张: %s" % media
    assert len(media) == EXPECTED_IMAGES
    assert len(list(root.iter(W + "drawing"))) == EXPECTED_IMAGES


def test_figs_on_disk():
    figs = sorted(FIGS.glob("*.png")) if FIGS.is_dir() else []
    assert len(figs) >= 2, "figs/ 里只有 %d 张图" % len(figs)
    for f in figs[:2]:
        assert f.stat().st_size > 20_000, "%s 像空白占位图" % f.name


def test_tables_present(root):
    tbls = list(root.iter(W + "tbl"))
    assert len(tbls) >= 2
    assert len(tbls) == EXPECTED_TABLES
    assert len(list(root.iter(W + "tr"))) == EXPECTED_ROWS, \
        "表格总行数应为 %d" % EXPECTED_ROWS


def test_answer_card_table_head(root):
    """答题卡表头是 题号 1..6，答案行必须真填了选项。"""
    tbl = list(root.iter(W + "tbl"))[0]
    rows = list(tbl.iter(W + "tr"))
    assert len(rows) == 2
    head = "".join(t.text or "" for t in rows[0].iter(W + "t"))
    body = "".join(t.text or "" for t in rows[1].iter(W + "t"))
    assert "题号" in head and "6" in head
    assert "答案" in body and "A" in body


def test_ion_table_has_embedded_math(root):
    """离子对照表里两行电极反应式仍是公式（表内混排，不是纯文本）。

    它们可能落成 OMML 对象，也可能落成带 ``w:vertAlign`` 的普通 run
    （纯上下标的电极式走后者）。两种情况都算「表内混排成公式」，判据是
    表格单元格里出现了 OMML 或文字角标，而不是纯文本。
    """
    embedded = 0
    for t in root.iter(W + "tbl"):
        embedded += len(list(t.iter(M + "oMath")))
        for r in t.iter(W + "r"):
            rpr = r.find(W + "rPr")
            if rpr is not None and rpr.find(W + "vertAlign") is not None:
                embedded += 1
    assert embedded >= 2, "表内嵌公式应至少 2 处，实际 %d" % embedded


# --------------------------------------------------------------- typography
def test_paragraph_formatting_landed(root):
    """缩进 / 行距 / 对齐三件套都要有足够落地点（实测 76 / 90 / 17）。"""
    ind = spac = jc = 0
    for p in root.iter(W + "p"):
        ppr = p.find(W + "pPr")
        if ppr is None:
            continue
        ind += ppr.find(W + "ind") is not None
        spac += ppr.find(W + "spacing") is not None
        jc += ppr.find(W + "jc") is not None
    assert ind >= 70, "缩进落地点太少: %d" % ind
    assert spac >= 80, "行距/段距落地点太少: %d" % spac
    assert jc >= 15, "对齐落地点太少: %d" % jc


def test_first_line_indent_uses_chars(root):
    """汉字题面的首行缩进必须是 w:firstLineChars（Word 的字符缩进）。"""
    chars = [i for i in root.iter(W + "ind") if i.get(W + "firstLineChars")]
    assert len(chars) >= 30, "首行缩进落地点太少: %d" % len(chars)


def test_option_indent_present(root):
    """选项 / 编号列表靠 indent_left_mm，w:ind/@w:left 必须有正值。"""
    lefts = [int(i.get(W + "left")) for i in root.iter(W + "ind")
             if i.get(W + "left")]
    assert lefts and max(lefts) > 0


# -------------------------------------------------------------- text content
def test_required_text_present(doc_xml):
    for s in ("绝密★启用前", "2026 年普通高等学校招生全国统一考试（模拟）",
              "注意事项", "第Ⅰ卷", "第Ⅱ卷", "参考答案与解析",
              "参考文献", "答题卡"):
        assert s in doc_xml, "正文缺少：%s" % s


def test_chemistry_subject_matter(doc_xml):
    """内容必须真是化学卷，不是占位行。"""
    for s in ("物质的量", "原电池", "酯化", "蒸馏", "阿伏加德罗常数"):
        assert s in doc_xml, "正文缺少化学内容：%s" % s


def test_reference_entries_numbered(root):
    refs = [b for b in root.iter(W + "bookmarkStart")
            if (b.get(W + "name") or "").startswith("ref_")]
    assert len(refs) == EXPECTED_REFS
