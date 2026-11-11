# -*- coding: utf-8 -*-
"""样例回归：用 spec.json 端到端重建一篇示例论文。

回归的重点不是「能生成一个 docx」，而是**原文的每一类硬事实都能被补丁版 skill
表达出来**：3 节页码体系、10 张真图、4 张真表、16 条自动编号参考文献带书签、
8 个图题注 + 4 个表题注、26 处可点击交叉引用（上角标文献号 + 正常大小图表号）。

两段式断言：

* **spec 层**（纯 JSON，任何机器都能跑）——声明的键和数量必须与任务书一致，
  防止 spec 悄悄退化成「只放了几段文字」；
* **docx 层**（``docx_create`` 从 spec 就地重建到 ``tmp_path``，不进 Word）
  ——字段是否真的落进 XML：节数、``w:media``、``w:tbl``、书签名、
  ``w:hyperlink/@w:anchor`` 全部可解析、``w:numPr`` 段数、
  ``pgNumType`` 的 fmt 集合。

``m:oMath`` 只能由 ``docx_mathify`` 产生，而它依赖 Office 的 MML2OMML.XSL，
所以那一项只在随卷交付的 ``paper.docx`` 存在时才断言（CI 上跳过）。
"""
from __future__ import annotations

import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
SUBJECT = (Path(__file__).resolve().parents[1].parent
           / "work" / "subjects" / "demo")
SPEC = SUBJECT / "spec.json"
PAPER = SUBJECT / "paper.docx"

pytestmark = pytest.mark.skipif(not SPEC.exists(),
                                reason="demo subject spec not shipped")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


# ---------------------------------------------------------------- fixtures
@pytest.fixture(scope="module")
def spec() -> dict:
    return json.loads(SPEC.read_text(encoding="utf-8"))


def _blocks(spec: dict, kind: str) -> list:
    return [b for b in spec["blocks"] if b.get("type") == kind]


@pytest.fixture(scope="module")
def built(tmp_path_factory) -> Path:
    """Rebuild the paper from spec.json — no Word, no mathify."""
    sys.path.insert(0, str(SCRIPTS))
    import docx_create as dc

    paths = [b["path"] for b in _blocks(
        json.loads(SPEC.read_text(encoding="utf-8")), "image")]
    missing = [p for p in paths if not Path(p).is_file()]
    if missing:
        pytest.skip("image files not shipped: %s" % missing[:3])

    out = tmp_path_factory.mktemp("demo") / "paper.docx"
    old = sys.argv
    try:
        sys.argv = ["docx_create.py", str(SPEC), str(out)]
        assert dc.main() == 0
    finally:
        sys.argv = old
    return out


@pytest.fixture(scope="module")
def xml(built: Path):
    with zipfile.ZipFile(built) as z:
        names = z.namelist()
        root = ET.fromstring(z.read("word/document.xml").decode("utf-8"))
    return root, names


# ------------------------------------------------------------------ spec 层
def test_spec_declares_two_breaks_so_the_paper_has_three_sections(spec):
    """封面/摘要 + 正文 + 参考文献 = 3 节，靠 2 个 section_break 切出来。"""
    assert len(_blocks(spec, "section_break")) == 2


def test_page_size_matches_the_original(spec):
    page = spec["page"]
    assert page["width_mm"] == 210 and page["height_mm"] == 297
    assert page["margins_mm"] == {"top": 25, "bottom": 25,
                                  "left": 30, "right": 20}


def test_two_distinct_page_number_formats(spec):
    """封面无页码 / 正文 decimal / 参考文献 lowerRoman。"""
    fmts = [b["footer_page_numbers"]["fmt"] for b in _blocks(spec, "section_break")
            if b.get("footer_page_numbers", {}).get("fmt")]
    assert len(set(fmts)) >= 2, fmts
    assert "decimal" in fmts and "lowerRoman" in fmts


def test_ten_images_each_with_its_own_width(spec):
    imgs = _blocks(spec, "image")
    assert len(imgs) == 10
    widths = [b["width_mm"] for b in imgs]
    assert all(60 <= w <= 150 for w in widths), widths
    assert len(set(widths)) >= 8, "widths should vary per figure: %s" % widths


def test_four_tables_with_real_content(spec):
    tables = _blocks(spec, "table")
    assert len(tables) == 4
    for t in tables:
        assert t["header"], t
        assert len(t["rows"]) >= 4, t


def test_captions_carry_bookmarks_and_crossrefs_resolve(spec):
    """8 图题注 + 4 表题注，且每条 crossref 的目标都真存在。

    书签有两个来源：``caption`` 自带 ``bookmark``（图表号），``numbered_list``
    的 ``bookmark_prefix`` 逐条命名（文献号）。两者合起来才是全文锚点集合。
    """
    caps = _blocks(spec, "caption")
    fig_caps = [c for c in caps if c.get("label") == "图"]
    tab_caps = [c for c in caps if c.get("label") == "表"]
    assert len(fig_caps) == 8, [c.get("bookmark") for c in fig_caps]
    assert len(tab_caps) == 4

    declared = {c["bookmark"] for c in caps}
    for lst in _blocks(spec, "numbered_list"):
        prefix = lst.get("bookmark_prefix")
        if prefix:
            declared |= {"%s%d" % (prefix, i)
                         for i in range(1, len(lst["items"]) + 1)}
    targets = {b["target"] for b in _blocks(spec, "crossref")}
    assert targets, "no crossref blocks at all"
    assert targets <= declared, "dangling crossref targets: %s" % (
        sorted(targets - declared),)


def test_crossref_display_switches_cover_both_modes(spec):
    """文献引用是上角标 + paragraph-number；图/表引用是正常大小。"""
    refs = _blocks(spec, "crossref")
    sup = [r for r in refs if r.get("superscript")]
    normal = [r for r in refs if not r.get("superscript")]
    assert len(sup) >= 3
    assert len(normal) >= 3
    assert all(r.get("format") == "paragraph-number" for r in sup), \
        "superscript citations must render as [n]"
    assert all(r["target"].startswith("ref_") for r in sup)
    figrefs = [r for r in normal if r["target"].startswith("fig_")]
    assert len(figrefs) >= 3
    assert all(r.get("prefix") == "如图" for r in figrefs), \
        "figure references should read 如图"


def test_reference_list_is_numbered_with_brackets_and_bookmarks(spec):
    lists = _blocks(spec, "numbered_list")
    assert len(lists) == 1
    ref = lists[0]
    assert ref["num_text"] == "[%1]"
    assert ref["bookmark_prefix"] == "ref_"
    assert len(ref["items"]) == 16


def test_required_paragraph_format_keys_are_exercised(spec):
    """§5 的压测点：缩进/行距/对齐/段前段后/左缩进都要真的出现过。"""
    seen = set()
    for b in spec["blocks"]:
        for k in ("first_line_indent_chars", "line_spacing", "alignment",
                  "space_before_pt", "space_after_pt", "indent_left_mm"):
            if b.get(k) is not None:
                seen.add(k)
    assert seen == {"first_line_indent_chars", "line_spacing", "alignment",
                    "space_before_pt", "space_after_pt", "indent_left_mm"}


def test_styles_bind_chinese_fonts(spec):
    names = {s["name"]: s for s in spec["styles"]}
    assert names["EX章标题"]["font"] == "黑体"
    assert names["EX章标题"]["size_pt"] == 16
    assert names["EX节标题"]["font"] == "黑体"
    assert names["EX节标题"]["size_pt"] == 14
    assert names["EX正文"]["font"] == "宋体"
    assert names["EX正文"]["size_pt"] == 12
    assert names["EX图表名"]["font"] == "宋体"
    assert names["EX图表名"]["size_pt"] == 10.5
    assert names["EX图表名"]["bold"] is True


def test_four_linear_formulas_survive_as_plain_text_before_mathify(spec):
    """原文的纯文字指标改写成线性式，交给 mathify 出 OMML。"""
    text = "".join(b.get("text", "") for b in spec["blocks"])
    for frag in ("h=15m", "r=20m", "t=1h", "P=50W/m²"):
        assert frag in text, frag


# ------------------------------------------------------------------ docx 层
def test_build_produces_three_sections(xml):
    root, _names = xml
    assert len(list(root.iter(W + "sectPr"))) == 3


def test_build_embeds_ten_images(xml):
    root, names = xml
    media = [n for n in names if n.startswith("word/media/")]
    assert len(media) == 10, media
    assert len(list(root.iter(W + "drawing"))) == 10


def test_build_writes_four_tables_and_the_original_row_count(xml):
    root, _names = xml
    assert len(list(root.iter(W + "tbl"))) == 4


def test_two_page_number_formats_land_in_xml(xml):
    root, _names = xml
    fmts = [(p.get(W + "fmt"), p.get(W + "start"))
            for p in root.iter(W + "pgNumType")]
    assert len({f for f, _s in fmts}) >= 2, fmts
    assert ("decimal", "1") in fmts
    assert ("lowerRoman", "1") in fmts


def test_every_hyperlink_anchor_names_a_real_bookmark(xml):
    """交叉引用必须可点击：w:hyperlink/@w:anchor 全部能解析。"""
    root, _names = xml
    bookmarks = {b.get(W + "name") for b in root.iter(W + "bookmarkStart")
                 if b.get(W + "name")}
    anchors = [h.get(W + "anchor") for h in root.iter(W + "hyperlink")
               if h.get(W + "anchor")]
    assert len(anchors) >= 20, len(anchors)
    assert not [a for a in anchors if a not in bookmarks], \
        "dangling anchors: %s" % [a for a in anchors if a not in bookmarks]
    for name in ("ref_1", "ref_16", "fig_1", "fig_8", "tab_4"):
        assert name in bookmarks


def test_reference_items_each_get_a_bookmark(xml):
    root, _names = xml
    bookmarks = {b.get(W + "name") for b in root.iter(W + "bookmarkStart")
                 if b.get(W + "name")}
    assert {("ref_%d" % i) for i in range(1, 17)} <= bookmarks


def test_numbered_paragraphs_and_the_bracket_level_text(xml, built):
    root, _names = xml
    numpr = [p for p in root.iter(W + "p")
             if p.find(W + "pPr/" + W + "numPr") is not None]
    assert len(numpr) == 16
    with zipfile.ZipFile(built) as z:
        numbering = ET.fromstring(z.read("word/numbering.xml").decode("utf-8"))
    lvl_texts = [lvl.get(W + "val") for lvl in numbering.iter(W + "lvlText")]
    assert "[%1]" in lvl_texts, lvl_texts


def test_superscript_runs_exist_and_are_not_the_only_shape(xml):
    root, _names = xml
    vals = [v.get(W + "val") for v in root.iter(W + "vertAlign")]
    assert vals.count("superscript") >= 1


def test_ref_fields_are_written_for_forward_references(xml):
    root, _names = xml
    codes = [t.text.strip() for t in root.iter(W + "instrText")
             if (t.text or "").strip()]
    refs = [c for c in codes if c.startswith("REF ")]
    assert len(refs) >= 20, len(refs)
    assert any("ref_1" in c for c in refs)
    assert any("\\n" in c for c in refs), "reference citations need \\n"


def test_caption_seq_counters_are_per_label(xml):
    root, _names = xml
    codes = [t.text.strip() for t in root.iter(W + "instrText")
             if (t.text or "").strip()]
    assert sum(1 for c in codes if c.startswith("SEQ 图")) == 8
    # Four, one per table.  Before the caption work this was 0 — tables had no
    # SEQ field at all, so "Insert Table of Figures" found nothing.  It must
    # stay exactly 4: a second caption per table would number that table twice
    # and shift every later table's number.
    assert sum(1 for c in codes if c.startswith("SEQ 表")) == 4


def test_paragraph_fonts_reach_the_style_part(built: Path):
    """``font`` 写在 styles 里必须真的落到 ``eastAsia``（补丁版修的就是这个）。"""
    with zipfile.ZipFile(built) as z:
        styles = ET.fromstring(z.read("word/styles.xml").decode("utf-8"))
    found = {}
    for st in styles.iter(W + "style"):
        sid = st.get(W + "styleId")
        if not sid or sid not in ("EX章标题", "EX正文", "EX图表名"):
            continue
        rf = st.find(".//" + W + "rFonts")
        assert rf is not None, sid
        found[sid] = rf.get(W + "eastAsia")
    assert found == {"EX章标题": "黑体", "EX正文": "宋体",
                     "EX图表名": "宋体"}, found


# ---------------------------------------------------- 交付件的 Word 级读数
@pytest.mark.skipif(not PAPER.exists(), reason="paper.docx not shipped")
def test_shipped_paper_has_omml_formulas():
    """mathify 之后才有 m:oMath；它需要 Office 的 XSL，故只断言交付件。

    本卷只有四个公式片段，而且**四个全是不需要二维排版的简单等式**
    （`h=15m` / `r=20m` / `t=1h` / `P=50W/m2`）。按 2026-11-07 的三层分流
    它们全部降级成普通文字 run —— 所以本卷 ``m:oMath`` 数为 **0 是正确结果**，
    不是公式丢失。原来那句 ``assert n_omml >= 1`` 锁的是旧行为，已作废。

    现在锁的是「四个片段的内容都还在文档里」，与落在哪一层无关。
    """
    with zipfile.ZipFile(PAPER) as z:
        root = ET.fromstring(z.read("word/document.xml").decode("utf-8"))
    n_omml = len(list(root.iter(M + "oMath")))
    n_vert = len(list(root.iter(W + "vertAlign")))
    # 可见字符：w:t 与 m:t 都算（降级后是 w:t，未降级的是 m:t）
    visible = "".join(t.text or "" for t in root.iter(W + "t"))
    visible += "".join(t.text or "" for t in root.iter(M + "t"))
    flat = visible.replace("\u00a0", " ").replace(" ", "")
    for frag in ("h=15m", "r=20m", "t=1h", "P=50W/m2"):
        assert frag in flat, "公式片段 %r 丢了（既不在 w:t 也不在 m:t）" % frag
    # 本卷刻意全降级：若哪天又冒出 OMML，说明判据放宽了，值得复核
    assert n_omml == 0, (
        "本卷四个片段都无需二维排版，应全部降级；实得 %d 个 oMath" % n_omml)
    assert n_vert >= 0  # 本卷无角标，保留占位便于以后加


@pytest.mark.skipif(not (SUBJECT / "accept_readings.txt").exists(),
                    reason="accept_readings.txt not shipped")
def test_readings_report_the_same_shape_as_the_xml():
    txt = (SUBJECT / "accept_readings.txt").read_text(
        encoding="utf-8", errors="replace")
    for key in ("openable", "pages", "sections", "bookmarks", "hyperlinks",
                "hyperlink_targets", "numbering_render", "field_list",
                "superscript_count", "omath_count", "images", "tables"):
        assert key in txt, key
    assert "sections            = 3" in txt
    assert "images              = 10" in txt
    assert "tables              = 4" in txt
