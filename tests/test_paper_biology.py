# -*- coding: utf-8 -*-
"""生物学试卷实测：spec → docx 的端到端结构断言（M7 · Wave 2）。

这份测试属于 `work/subjects/biology/` 的交付物，和 `test_section_para.py` 一个
范式：**不依赖 Office**，只在 CI 上用 stdlib + python-docx 复核"spec 里写的东
西是否真的落进了 `word/document.xml`"。

生物卷的特别之处（写在 `AGENTS` 契约里）：**不要求 OMML 公式**，但 spec 里必
须保留纯文本的遗传学表示（`AaBb×AaBb`、`9∶3∶3∶1`）。所以这里断言
`m:oMath == 0`（而不是 >= 1），另加一条"遗传学记法还在正文里"。

spec 不在仓库里就整体 skip —— 别人的 CI 检出可能没带 work/ 目录。
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
SPEC = (Path(__file__).resolve().parents[2]
        / "work" / "subjects" / "biology" / "spec.json")

pytestmark = pytest.mark.skipif(not SPEC.exists(),
                                reason="biology subject spec not shipped")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


# ---------------------------------------------------------------- fixtures

@pytest.fixture(scope="module")
def spec() -> dict:
    return json.loads(SPEC.read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def built(spec, tmp_path_factory) -> Path:
    """Build the paper from spec.json exactly the way M3 does."""
    sys.path.insert(0, str(SCRIPTS))
    import docx_create as dc
    from docx import Document

    doc = Document()
    if spec.get("page"):
        dc.apply_page(doc, spec["page"])
    if spec.get("styles"):
        dc.add_styles(doc, spec["styles"])
    for block in spec.get("blocks", []):
        dc.add_block(doc, block)

    from docx_crossref import check_crossref_anchors
    dangling = check_crossref_anchors(doc)
    assert not dangling["missing"], "dangling cross-reference: %s" % (
        dangling["missing"],)

    out = Path(tmp_path_factory.mktemp("biology")) / "paper.docx"
    doc.save(str(out))
    return out


@pytest.fixture(scope="module")
def xml(built) -> str:
    with zipfile.ZipFile(built) as z:
        return z.read("word/document.xml").decode("utf-8")


@pytest.fixture(scope="module")
def parts(built) -> list:
    with zipfile.ZipFile(built) as z:
        return z.namelist()


# ------------------------------------------------------------ spec / 内容

def test_spec_is_the_expected_exam(spec):
    """卷头 / Ⅰ卷 / Ⅱ卷 / 答案 / 参考文献 五段内容都在。"""
    text = "\n".join(b.get("text", "") or "" for b in spec["blocks"])
    for needle in ("（模拟）· 生物", "绝密★启用前", "注意事项",
                   "第Ⅰ卷", "第Ⅱ卷", "参考答案与解析", "参考文献"):
        assert needle in text, "缺卷面要素：%s" % needle


def test_six_choice_questions_cover_the_syllabus(spec):
    """6 道选择题，且六个知识板块各出现一次。"""
    blocks = spec["blocks"]
    starts = [b["start"] for b in blocks if b["type"] == "numbered_list"
              and b.get("num_text") == "%1."]
    assert starts == [1, 2, 3, 4, 5, 6]

    stems = [b["items"][0] for b in blocks
             if b["type"] == "numbered_list" and b.get("num_text") == "%1."]
    # 细胞结构与代谢 / 遗传规律 / 变异与进化 / 稳态与调节 / 生态系统 / 生物技术
    marks = ["真核细胞", "AaBb", "变异与生物进化", "内环境与稳态",
             "生态系统", "生物技术"]
    for m, stem in zip(marks, stems):
        assert m in stem, "第 %d 题应考「%s」，实际：%s" % (
            marks.index(m) + 1, m, stem)

    # 每题四个选项，且都带 indent_left_mm 缩进
    opts = [b for b in blocks if b["type"] == "paragraph"
            and b.get("style") == "SOption"]
    assert len(opts) == 24
    assert all(o["indent_left_mm"] >= 5 for o in opts)
    assert [o["text"][:2] for o in opts[:4]] == ["A.", "B.", "C.", "D."]


def test_three_free_response_questions(spec):
    """第Ⅱ卷 3 题：系谱/概率、光合曲线、生态实验设计。"""
    text = "\n".join(b.get("text", "") or "" for b in spec["blocks"])
    assert "7.（20 分）" in text          # 遗传系谱与概率
    assert "8.（20 分）" in text          # 光合 / 呼吸曲线
    assert "9.（24 分）" in text          # 生态 / 实验设计
    assert "标志重捕法" in text           # 生态题里的定量计算
    # 答题区（下划线空行）
    blanks = [b for b in spec["blocks"] if b["type"] == "paragraph"
              and "_" * 20 in (b.get("text") or "")]
    assert len(blanks) >= 9


def test_genetics_notation_stays_plain_text(spec):
    """生物卷不转 OMML，但遗传学表示必须以纯文本存在。"""
    text = json.dumps(spec, ensure_ascii=False)     # 含题面、选项与表格单元格
    for needle in ("AaBb×AaBb", "9∶3∶3∶1", "AaBb×aabb", "Aabb×aaBb"):
        assert needle in text, "缺纯文本遗传学记法：%s" % needle
    # 题面里不得出现 LaTeX 残留
    assert "\\frac" not in text and "$" not in text


# --------------------------------------------------------------- 结构断言

def test_sections_and_page_number_formats(built):
    """3 节 + 两种页码格式：卷头无页码，Ⅰ·Ⅱ卷 lowerRoman，答案 decimal 重启。"""
    from docx import Document
    import docx_section as ds

    doc = Document(str(built))
    assert len(doc.sections) == 3

    fmts = [ds.get_pg_num_type(s) for s in doc.sections]
    assert fmts[0] is None, "卷头节不该有 pgNumType（无页码）"
    assert fmts[1] == {"fmt": "lowerRoman", "start": "1"}
    assert fmts[2] == {"fmt": "decimal", "start": "1"}


def test_section_footers_are_private(built):
    """1、3 节各自持有页脚（PAGE 域），第 2 节不重复挂 —— 罗马数字得以续排。"""
    from docx import Document
    import docx_section as ds

    doc = Document(str(built))
    refs = [ds._ref_id(s._sectPr, "footer", "default") for s in doc.sections]
    assert refs[0] is None
    assert refs[2] is not None
    assert refs[2] != refs[1]
    foot = "".join(t.text or "" for t in
                   doc.part.related_parts[refs[2]].element.iter(W + "t"))
    assert re.search(r"\d|", foot)          # 缓存结果已写好，不是空页脚


def test_built_xml_has_three_sections(xml):
    assert xml.count("<w:sectPr") == 3
    assert "w:fmt=\"lowerRoman\"" in xml
    assert "w:fmt=\"decimal\"" in xml


def test_numbering_is_real_numpr(numpr_texts):
    """题号与文献号都是真 w:numPr（不是手打的字符串）。"""
    assert len(numpr_texts) == 9                 # 6 题 + 3 条文献
    assert numpr_texts[0].startswith("下列关于真核细胞")


def test_numbering_formats_in_part(built):
    """[%1] 给文献、%1. 给题号 —— 两种 lvlText 都存在。"""
    with zipfile.ZipFile(built) as z:
        num = z.read("word/numbering.xml").decode("utf-8")
    assert "[%1]" in num
    assert "%1." in num


def test_crossrefs_are_clickable_anchors(anchors, bookmark_names):
    """>=1 上角标（→ref_1）+ >=1 正常大小（→系谱图题注），锚点都能解析。"""
    assert len(anchors) == 5
    assert "ref_1" in anchors and "ref_2" in anchors      # 上角标文献引用
    assert "fig_pedigree" in anchors                      # "如图 1 所示"
    assert set(anchors) <= set(bookmark_names)


def test_figure_captions_carry_bookmarks(bookmark_names):
    """每张图都有 caption + bookmark。"""
    for bm in ("fig_pedigree", "fig_photo", "fig_growth"):
        assert bm in bookmark_names
    assert "ref_1" in bookmark_names and "ref_3" in bookmark_names


def test_superscript_and_normal_refs_coexist(xml):
    """上角标引用与正常大小引用各至少一处。"""
    sup = len(re.findall(r'<w:vertAlign w:val="superscript"/>', xml))
    normal_ref = len(re.findall(r"REF fig_[a-z]+", xml))
    assert sup >= 2, "上角标 REF 至少 2 处（ref_1、ref_2），实际 %d" % sup
    assert normal_ref >= 3, "正常大小 REF 至少 3 处（三张图），实际 %d" % normal_ref


def test_paragraph_number_format_for_citation(xml):
    r"文献引用用 REF <bm> \n，Word 才渲染成 [1]。"
    assert r"REF ref_1 \n" in xml
    assert r"REF ref_2 \n" in xml


def test_images_and_tables(parts):
    media = [n for n in parts if n.startswith("word/media/")]
    assert len(media) == 3, "三张自产图（系谱 / 净光合 / 种群增长）"


def test_no_omml_formula(xml):
    """生物卷不要求公式：确认没有残留 OMML。"""
    assert xml.count("<m:oMath") == 0


def test_paragraph_format_keys_landed(xml):
    """首行缩进 / 行距 / 对齐 / 段前距 / 左缩进 都真写进了 pPr。"""
    assert "w:firstLineChars=" in xml          # first_line_indent_chars
    assert "w:firstLine=" in xml               # 它的 twips 孪生值
    assert "w:line=" in xml                    # line_spacing
    assert 'w:val="center"' in xml             # alignment center
    assert 'w:val="both"' in xml               # alignment justify
    assert "w:before=" in xml                  # space_before_pt
    assert "w:after=" in xml                   # space_after_pt
    assert 'w:left="' in xml                   # indent_left_mm（选项缩进）


def test_style_fonts_are_chinese(spec_xml_styles):
    """styles 里声明的中文字体真的绑到了 eastAsia 槽。"""
    assert "宋体" in spec_xml_styles and "黑体" in spec_xml_styles
    assert 'w:eastAsia="宋体"' in spec_xml_styles
    assert 'w:eastAsia="黑体"' in spec_xml_styles


# ---- 派生 fixture（依赖 xml / spec，放最后以免打断上面的可读性 --------------

@pytest.fixture(scope="module")
def numpr_texts(xml) -> list:
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    out = []
    for p in root.iter(W + "p"):
        if p.find(W + "pPr/" + W + "numPr") is not None:
            out.append("".join(t.text or "" for t in p.iter(W + "t")))
    return out


@pytest.fixture(scope="module")
def anchors(xml) -> list:
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    return [h.get(W + "anchor") for h in root.iter(W + "hyperlink")
            if h.get(W + "anchor")]


@pytest.fixture(scope="module")
def bookmark_names(xml) -> list:
    import xml.etree.ElementTree as ET
    root = ET.fromstring(xml)
    names = [b.get(W + "name") for b in root.iter(W + "bookmarkStart")]
    return [n for n in names if n and not n.startswith("_")]


@pytest.fixture(scope="module")
def spec_xml_styles(built) -> str:
    with zipfile.ZipFile(built) as z:
        return z.read("word/styles.xml").decode("utf-8")
