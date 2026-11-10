# -*- coding: utf-8 -*-
"""六卷「原生 outline_level → 静态目录」回归测试（R7 / Phase 3-A）。

背景：R7 Phase 1/2 给六卷真实 spec 落了 ``outline_level``、并插了静态目录 +
页码闭环，但 ``tests\\test_paper_*.py`` 断言的是 ``spec → docx_create →
docx_mathify`` 链路产物、**不含 TOC 后处理**，所以「六卷目录」这件事完全没有
测试视野（pytest 1440 全绿 ≠ 目录被测）。本文件补上这一段。

两层断言，边界很清楚：

* **XML 层（CI 可跑，必测）** —— 只读 ``word/document.xml`` /
  ``word/styles.xml``，不需要 Office：
    1. ``collect_outline`` 条目数与条目文本逐卷等于 ``work\\R7-MAPPING.md`` §八
       定稿清单（9/9/9/7/10/27）；条目级别序列同样锁死；
    2. 条目段样式命中 ``^toc\\s*[1-3]$``、缩进 0/2/4 字符（``w:leftChars``
       与 ``w:left`` 双写一致）、右对齐点线制表位落在该节版心
       （math 170 mm / 其余 160 mm）；
    3. ``styles.xml`` 里各样式 ``w:outlineLvl`` 与 §八 全表**逐卷相等**
       （0-based），且 ``SSub`` / ``SMark`` / ``卷头说明`` / ``EX目录壳``
       这些「语义转版式层」的样式**没有** ``outlineLvl``；
    4. 幂等：同一文档重复插目录不增殖（条目数、段数都不变）；
    5. demo 专有：主目录 27 条 + 图目录 8 行 + 表目录 4 行三层共存，
       且旧目录壳（``EX目录壳`` 标题 + 27 行手写条目）是真替换而非并列；
    6. 分页：五卷分页目标段带 ``w:pageBreakBefore``；demo 不靠 pbr，
       靠节断（目录区之后的 ``sectPr``）另起页。

* **Word 层（本机有 Word 才跑，默认 skip）** —— 页码闭环与「目录独占一页」
  必须由真实排版给出，XML 层给不了。三条纪律来自 R7 Phase 2 的四条方法学
  纠正（PLAN-ROUND7 硬约束 #28–#32）：
    * 闭环比 ``displayed_text``（保留 ``i``/``iii`` 字形），**不比**
      ``displayed_number``（已数值化，``iii``→3）；
    * ``--page-probe`` 的索引空间是 Word ``Paragraphs(n)``（含表内段与表行
      尾标记），必须用「全部 ``w:p`` + 表格行数」映射，不能拿
      ``doc.paragraphs`` 索引直加；
    * ``paragraph_pages`` 首位可能是目录标题段，比对按索引键取值、不按位置硬凑。

  本机虽有 Office，本轮（Phase 3-A）纪律是**不跑 Word**（会与别的代理抢
  ``run_accept.py`` 的 COM 令牌锁），所以再加一道 opt-in 开关：只有显式设
  ``DOCX_TOC_WORD_TESTS=1`` 才进入。CI 无 Word 时由 ``skipif`` 干净跳过。

测试对象是 ``work\\r7_final\\<卷>.docx``，由 ``work\\r7_pipeline\\
make_six_volumes.py`` 一条命令产出（不在就现场生成；spec 与流水线都不在则
整文件 skip）。全程不写 ``skill-patched\\``。
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

HERE = Path(__file__).resolve()
SCRIPTS = HERE.parents[1] / "scripts"
PROJECT = HERE.parents[2]
WORK = PROJECT / "work"
FINAL = WORK / "r7_final"
PIPELINE = WORK / "r7_pipeline"
SUBJECTS = WORK / "subjects"
RUN_ACCEPT = SUBJECTS / "_tools" / "run_accept.py"

sys.path.insert(0, str(SCRIPTS))

from docx import Document                                    # noqa: E402
from docx.oxml.ns import qn                                  # noqa: E402

import docx_create as dc                                     # noqa: E402
import docx_toc as toc                                       # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

VOLUMES = ("math", "physics", "chemistry", "biology", "comprehensive", "demo")

# 逐卷定稿期望值。出处：
#   entries / levels / styles_outline —— work\R7-MAPPING.md §八（六卷全表）
#   titles —— 该表 + spec 正文（以「空白折叠后逐字相等」口径比对）
#   tab_mm —— 该卷目录所在节的版心宽（math 是 170 mm 的那一卷）
#   page_break —— 五卷用 w:pageBreakBefore 另起页；demo 靠节断，故为 False
EXPECT = {
    "math": {
        "entries": 9,
        "levels": [1, 1, 2, 2, 2, 3, 3, 2, 2],
        "titles": [
            "2026年普通高等学校招生全国统一考试（模拟）",
            "数学",
            "第Ⅰ卷（选择题共30分）",
            "第Ⅱ卷（非选择题共120分）",
            "参考答案与解析",
            "一、选择题（每小题5分，共30分）",
            "二、解答题（共120分）",
            "参考文献",
            "附录：评分细则与命题说明",
        ],
        "styles_outline": {"卷标题": 0, "卷节标题": 1, "答案小节标题": 2},
        "styles_flat": ["卷头标签", "正文"],
        "tab_mm": 170.0,
        "page_break": True,
        "anchor_text": "第Ⅰ卷（选择题",
    },
    "physics": {
        "entries": 9,
        "levels": [1, 1, 2, 2, 2, 3, 3, 3, 3],
        "titles": [
            "2026年普通高等学校招生全国统一考试（模拟）·物理",
            "物理",
            "第Ⅰ卷（选择题共42分）",
            "第Ⅱ卷（非选择题共58分）",
            "参考答案与解析",
            "一、选择题（每小题7分，共42分）",
            "二、非选择题（共58分）",
            "三、分值分布",
            "四、命题依据与参考文献",
        ],
        "styles_outline": {"STitle": 0, "SSubject": 0, "SSection": 1,
                           "SHead": 2},
        "styles_flat": ["SMark"],
        "tab_mm": 160.0,
        "page_break": True,
        "anchor_text": "第Ⅰ卷（选择题",
    },
    "chemistry": {
        "entries": 9,
        "levels": [1, 1, 2, 2, 2, 2, 2, 2, 2],
        "titles": [
            "2026年普通高等学校招生全国统一考试（模拟）",
            "化学",
            "第Ⅰ卷（选择题共36分）",
            "第Ⅱ卷（非选择题共64分）",
            "参考答案与解析",
            "一、选择题（每小题6分，共36分）",
            "二、非选择题（共64分）",
            "三、常见离子颜色与电极反应对照表",
            "参考文献",
        ],
        "styles_outline": {"卷标题": 0, "科目标题": 0, "小标题": 1,
                           "小节标题": 1},
        "styles_flat": ["卷头说明", "正文"],
        "tab_mm": 160.0,
        "page_break": True,
        "anchor_text": "第Ⅰ卷（选择题",
    },
    "biology": {
        "entries": 7,
        "levels": [1, 2, 2, 1, 2, 2, 2],
        "titles": [
            "2026年普通高等学校招生全国统一考试（模拟）·生物",
            "第Ⅰ卷（选择题共36分）",
            "第Ⅱ卷（非选择题共64分）",
            "参考答案与解析",
            "一、选择题（每小题6分，共36分）",
            "二、非选择题（共64分）",
            "参考文献",
        ],
        "styles_outline": {"STitle": 0, "SHead": 1},
        "styles_flat": ["SSub"],
        "tab_mm": 160.0,
        "page_break": True,
        "anchor_text": "第Ⅰ卷（选择题",
    },
    "comprehensive": {
        "entries": 10,
        "levels": [1, 2, 2, 2, 2, 1, 3, 3, 3, 3],
        "titles": [
            "2026年普通高等学校招生全国统一考试（模拟）·理科综合",
            "（物理·化学·生物合卷）",
            "试卷结构与分值分配",
            "第Ⅰ卷选择题",
            "第Ⅱ卷非选择题",
            "参考答案与解析",
            "一、分值分布",
            "二、三科题号—答案对照",
            "三、解析",
            "参考文献",
        ],
        "styles_outline": {"STitle": 0, "SChapter": 1, "SHead": 2},
        "styles_flat": ["卷头说明"],
        "tab_mm": 160.0,
        "page_break": True,
        "anchor_text": "试卷结构与分值分配",
    },
    "demo": {
        "entries": 27,
        "levels": [1, 1, 1, 1, 2, 2, 1, 2, 2, 2, 1, 2, 3, 3, 3, 2, 3, 3, 2, 3,
                   3, 1, 2, 2, 2, 1, 1],
        "titles": [
            "摘要",
            "Abstract",
            "引言",
            "一、文献综述",
            "（一）国内外相关技术研究现状",
            "（二）现有问题与创新点",
            "二、设计原理",
            "（一）系统平台原理",
            "（二）信号覆盖原理",
            "（三）结构优化设计原理",
            "三、系统总体设计与方法",
            "（一）系统总体架构设计",
            "1.系统组成",
            "2.系统工作流程",
            "3.主要技术指标",
            "（二）结构优化设计方法",
            "1.部件选型与优化",
            "2.结构优化验证",
            "（三）部署流程设计方法",
            "1.模块化拆装设计",
            "2.部署流程优化",
            "四、指标分析验证",
            "（一）结构优化效果",
            "（二）部署效率",
            "（三）性能与环境适应性",
            "结论",
            "参考文献",
        ],
        "styles_outline": {"EX章标题": 0, "EX节标题": 1, "EX小标题": 2},
        "styles_flat": ["EX目录壳"],
        "tab_mm": 160.0,
        "page_break": False,        # 该卷目录区之后是节断（sectPr nextPage）
        "anchor_text": "目录",
    },
}

# 目录缩进（docx_toc.DEFAULT_INDENT_CHARS）：一级 0 字、二级 2 字、三级 4 字。
INDENT_CHARS = {1: 0, 2: 2, 3: 4}


# ---------------------------------------------------------------- 可用性判定
def _inputs_available() -> bool:
    """成品在 → 直接测；成品不在但流水线 + 六卷 spec 都在 → 现场生成。"""
    if all((FINAL / ("%s.docx" % vol)).exists() for vol in VOLUMES):
        return True
    return ((PIPELINE / "make_six_volumes.py").exists()
            and all((SUBJECTS / vol / "spec.json").exists()
                    for vol in VOLUMES))


pytestmark = pytest.mark.skipif(
    not _inputs_available(),
    reason="六卷成品 work\\r7_final 与 spec/流水线均不可用（本机未建卷）",
)


def _run_pipeline() -> None:
    """用流水线现场生成六卷成品（写入 work\\r7_final\\）。"""
    import importlib.util

    path = PIPELINE / "make_six_volumes.py"
    spec = importlib.util.spec_from_file_location("make_six_volumes", str(path))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    rc = module.main([])
    assert rc == 0, "六卷流水线返回 %s" % rc


@pytest.fixture(scope="session")
def products() -> dict:
    """``{卷: work\\r7_final\\<卷>.docx}``；缺货时现场跑流水线。"""
    paths = {vol: FINAL / ("%s.docx" % vol) for vol in VOLUMES}
    if not all(path.exists() for path in paths.values()):
        _run_pipeline()
    missing = [vol for vol, path in paths.items() if not path.exists()]
    assert not missing, "流水线未能产出：%s" % missing
    return paths


@pytest.fixture(scope="session")
def docs(products) -> dict:
    return {vol: Document(str(path)) for vol, path in products.items()}


# ---------------------------------------------------------------- 小工具
def norm(text) -> str:
    """折叠全部空白（含全角空格 U+3000、不换行空格）后比较。

    目录条目与标题段的差别只在制表位与页码槽，字符本身应当一致；把空白归一
    是为了不让 ``w:t`` 的拆 run 方式（``第Ⅰ卷（选择题`` + 全角空格 + ``共 36 分）``）
    影响判定。
    """
    return re.sub(r"\s+", "", str(text or ""))


def entries_of(doc):
    return [p for p in doc.paragraphs if toc.is_toc_paragraph(p)]


def split_at_tab(para):
    """按第一个 run 级 ``<w:tab/>`` 把条目切成（标题, 页码槽）。

    真实制表符在 XML 里是元素而不是 ``w:t`` 里的 ``\\t``，所以按元素切；
    ``w:tabs`` 里的 ``w:tab`` 是「声明制表位」，不是字符，必须排除。
    """
    pieces, current = [], ""
    for child in para._p.iter():
        if child.tag == qn("w:tab") and child.getparent().tag != qn("w:tabs"):
            pieces.append(current)
            current = ""
        elif child.tag == qn("w:t"):
            current += child.text or ""
    pieces.append(current)
    if len(pieces) < 2:
        return (pieces[0] if pieces else ""), ""
    return pieces[0], pieces[-1].strip()


def run_tabs(para):
    return [t for t in para._p.iter(qn("w:tab"))
            if t.getparent().tag != qn("w:tabs")]


def tabs_of(para):
    """``[(val, pos_twips, leader), ...]``；没有 ``w:tabs`` 时为 ``[]``。"""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return []
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        return []
    return [(t.get(qn("w:val")), t.get(qn("w:pos")), t.get(qn("w:leader")))
            for t in tabs.findall(qn("w:tab"))]


def indent_of(para):
    """``(leftChars, left_twips)``，缺 ``w:ind`` 时为 ``(None, None)``。"""
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return None, None
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return None, None
    left_chars = ind.get(qn("w:leftChars"))
    left = ind.get(qn("w:left"))
    return (int(left_chars) if left_chars is not None else None,
            int(left) if left is not None else None)


def has_page_break_before(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    if pPr is None:
        return False
    element = pPr.find(qn("w:pageBreakBefore"))
    if element is None:
        return False
    return (element.get(qn("w:val")) or "true").lower() not in ("0", "false")


def has_sectpr(para) -> bool:
    pPr = para._p.find(qn("w:pPr"))
    return pPr is not None and pPr.find(qn("w:sectPr")) is not None


def style_outline_levels(docx_path: Path) -> dict:
    """``styles.xml`` 里每个样式名 -> ``w:outlineLvl/@w:val``（0-based 字符串）。

    只看段落样式（``w:type="paragraph"`` 或缺省）；没有 ``outlineLvl`` 的样式
    不出现在结果里 —— 这正是「语义转版式层」的判据。
    """
    with zipfile.ZipFile(str(docx_path)) as archive:
        root = ET.fromstring(archive.read("word/styles.xml"))
    out = {}
    for style in root.iter(W + "style"):
        kind = style.get(W + "type")
        if kind not in (None, "paragraph"):
            continue
        name = style.find(W + "name")
        name = name.get(W + "val") if name is not None else style.get(W + "styleId")
        if not name:
            continue
        pPr = style.find(W + "pPr")
        level = None if pPr is None else pPr.find(W + "outlineLvl")
        if level is not None:
            out[name] = level.get(W + "val")
    return out


def toc_block_of(doc):
    """目录块在 ``doc.paragraphs`` 里的 ``(首行, 末行)``；无目录时返回 ``None``。"""
    lines = [i for i, p in enumerate(doc.paragraphs)
             if toc.is_toc_paragraph(p) or toc.is_our_toc_title(p)]
    if not lines:
        return None
    return lines[0], lines[-1]


# ================================================================ 一、条目清单
@pytest.mark.parametrize("vol", VOLUMES)
def test_entry_count_matches_mapping(vol, products, docs):
    """条目数与 R7-MAPPING §八 定稿逐卷相等：9/9/9/7/10/27（合计 71）。"""
    want = EXPECT[vol]["entries"]
    outline = toc.collect_outline(docs[vol])
    assert len(outline) == want, (
        "%s 大纲条目 %d != 预期 %d（spec 的 outline_level 被动过？）"
        % (vol, len(outline), want))
    assert len(entries_of(docs[vol])) == want, "目录条目段数 != 大纲条目数"


@pytest.mark.parametrize("vol", VOLUMES)
def test_entry_titles_match_mapping(vol, products, docs):
    """条目文本逐条等于 §八 清单（空白折叠后逐字相等）。

    这条是「目录跟着 spec 走」的锁：spec 里改一个标题字，这里必须同步改，
    否则就是无声的正文改动。
    """
    got = [norm(split_at_tab(p)[0]) for p in entries_of(docs[vol])]
    want = EXPECT[vol]["titles"]
    assert got == want, "%s 条目文本不一致：\n  实测 %s\n  期望 %s" % (
        vol, got, want)


@pytest.mark.parametrize("vol", VOLUMES)
def test_entry_levels_match_mapping(vol, products, docs):
    """条目级别序列与 §八 全表相等（含 math 答案小节 L3、physics SHead L3）。"""
    got = [e["level"] for e in toc.collect_outline(docs[vol])]
    assert got == EXPECT[vol]["levels"], (
        "%s 条目级别序列不一致：%s != %s"
        % (vol, got, EXPECT[vol]["levels"]))


# ============================================================ 二、条目段版式
@pytest.mark.parametrize("vol", VOLUMES)
def test_entry_styles_are_word_builtin_toc_styles(vol, products, docs):
    """条目段样式必须命中 ``^toc\\s*[1-3]$``，且与条目级别一一对应。

    样式名不是 ``toc N`` 时 Word 的「更新目录」不再认这些行；写成
    ``toc1``／自定义名同样会失效（styleId 才是内置身份，见 docx_toc 注释）。
    """
    entries = entries_of(docs[vol])
    levels = [e["level"] for e in toc.collect_outline(docs[vol])]
    names = [p.style.name for p in entries]
    for name in names:
        assert re.match(r"^toc\s*[1-3]$", name), "非法条目样式 %r" % name
    assert names == ["toc %d" % level for level in levels], (
        "%s 条目样式与级别不匹配：%s vs %s" % (vol, names, levels))


@pytest.mark.parametrize("vol", VOLUMES)
def test_entry_indent_is_zero_two_four_chars(vol, products, docs):
    """缩进按级别 0/2/4 字符，``w:leftChars`` 与 ``w:left`` 双写一致。

    只写 twips 会被 Word 下次保存按字体「修」一遍；只写 chars 别的解析器读不到。
    """
    entries = entries_of(docs[vol])
    levels = [e["level"] for e in toc.collect_outline(docs[vol])]
    for para, level in zip(entries, levels):
        want_chars = INDENT_CHARS[level]
        left_chars, left = indent_of(para)
        assert left_chars == want_chars * 100, (
            "%s %r 缩进 leftChars=%r，期望 %d"
            % (vol, para.text[:12], left_chars, want_chars * 100))
        assert left == int(round(want_chars * dc.DEFAULT_CHAR_PT * 20)), (
            "%s %r 缩进 left=%r twips，期望 %d"
            % (vol, para.text[:12], left,
               int(round(want_chars * dc.DEFAULT_CHAR_PT * 20))))


@pytest.mark.parametrize("vol", VOLUMES)
def test_tab_stop_is_right_dot_at_content_width(vol, products, docs):
    """右制表位 = 该节版心：math 170 mm（9638 twips），其余 160 mm（9071）。"""
    want_mm = EXPECT[vol]["tab_mm"]
    want_twips = int(round(want_mm * 1440 / 25.4))
    doc = docs[vol]
    for para in entries_of(doc):
        stops = tabs_of(para)
        assert len(stops) == 1, "%r 制表位个数 %d != 1" % (
            para.text[:16], len(stops))
        val, pos, leader = stops[0]
        assert (val, leader) == ("right", "dot"), "制表位不是右对齐点线：%r" % (
            (val, leader),)
        assert int(pos) == want_twips, (
            "tab 位 %s twips != %d（%.1f mm）" % (pos, want_twips, want_mm))
        assert len(run_tabs(para)) == 1, "条目缺 run 级 <w:tab/>"
    # 版心取自真实 section 几何，不是硬编码常数
    content_mm = min(round(dc.content_width_mm(s), 3)
                     for s in doc.sections)
    assert content_mm == pytest.approx(want_mm, abs=0.05), (
        "%s 版心 %.3f mm != %.1f mm" % (vol, content_mm, want_mm))


# ========================================================= 三、styles.xml 层级
@pytest.mark.parametrize("vol", VOLUMES)
def test_style_outline_levels_match_mapping(vol, products):
    """各样式 ``w:outlineLvl``（0-based）与 §八 全表逐卷相等。"""
    got = style_outline_levels(products[vol])
    for name, level in EXPECT[vol]["styles_outline"].items():
        assert name in got, "%s 缺少样式 %r 的 outlineLvl" % (vol, name)
        assert int(got[name]) == level, (
            "%s 样式 %r outlineLvl=%s，期望 %d（level %d）"
            % (vol, name, got[name], level, level + 1))


@pytest.mark.parametrize("vol", VOLUMES)
def test_noise_styles_carry_no_outline_level(vol, products):
    """``SSub`` / ``SMark`` / ``卷头说明`` / ``EX目录壳`` 等**没有** outlineLvl。

    R7 把这些「原本被 outline_level 误表达」的版式样式降级为纯版式样式；它们
    一旦重新带上 outlineLvl，目录就会把这些噪声段列进来（R6 的 clean 逻辑
    正是为此存在，R7 要让它在 spec 层面成为不必要）。
    """
    got = style_outline_levels(products[vol])
    for name in EXPECT[vol]["styles_flat"]:
        assert name not in got, (
            "%s 的 %r 带上了 outlineLvl=%s（应为纯版式样式）"
            % (vol, name, got.get(name)))


# ================================================================ 四、幂等
@pytest.mark.parametrize("vol", VOLUMES)
def test_rebuild_does_not_accumulate(vol, products, tmp_path):
    """同一文档重复插目录不增殖：段数与条目数都不变。

    幂等靠 ``_DocxTocEntry…`` 书签（``drop_existing_toc`` 只删带标记的段落），
    不靠样式名 —— 用户自己拿 Word 生成的 ``toc 1`` 段落不会被误删。
    """
    source = products[vol]
    doc = Document(str(source))
    before_paras = len(doc.paragraphs)
    before_entries = len(entries_of(doc))
    first = toc.build_static_toc(doc, levels=(1, 2, 3), title="目录")
    second = toc.build_static_toc(doc, levels=(1, 2, 3), title="目录")
    assert first["entries"] == second["entries"] == EXPECT[vol]["entries"]
    assert len(doc.paragraphs) == before_paras, (
        "%s 重插后段数 %d != %d" % (vol, len(doc.paragraphs), before_paras))
    assert len(entries_of(doc)) == before_entries == EXPECT[vol]["entries"]
    out = tmp_path / ("%s.rebuild.docx" % vol)
    doc.save(str(out))
    again = Document(str(out))
    assert len(entries_of(again)) == EXPECT[vol]["entries"]
    assert [norm(split_at_tab(p)[0]) for p in entries_of(again)] == \
        EXPECT[vol]["titles"]


# ========================================================= 五、demo 三层共存
def test_demo_main_fig_tab_directories_coexist(products, docs):
    """demo：主目录 27 条 + 图目录 8 行 + 表目录 4 行三层共存。

    图／表目录是**手写**行（页码值来自作者，不由分页器回填），R7 只把它们的
    壳标题从 ``EX章标题``（L1，会污染主目录）降到 ``EX目录壳``（无层级）。
    """
    doc = docs["demo"]
    paragraphs = doc.paragraphs
    texts = [p.text.strip() for p in paragraphs]

    assert texts.count("目录") == 1, "目录标题应恰好一个"
    assert texts.count("图目录") == 1
    assert texts.count("表目录") == 1

    main_toc = entries_of(doc)
    assert len(main_toc) == 27

    fig_shell = texts.index("图目录")
    tab_shell = texts.index("表目录")
    assert fig_shell < tab_shell, "图目录应在表目录之前"

    fig_rows = paragraphs[fig_shell + 1:fig_shell + 1 + 8]
    assert all("\t" in p.text for p in fig_rows), "图目录应有 8 行带 tab 的条目"
    assert paragraphs[fig_shell + 9].text.strip() == "表目录", (
        "图目录 8 行之后应紧跟表目录壳")

    tab_rows = paragraphs[tab_shell + 1:tab_shell + 1 + 4]
    assert all("\t" in p.text for p in tab_rows), "表目录应有 4 行带 tab 的条目"

    # 三个壳标题都不进主目录（样式无层级）
    outline_titles = [norm(e["text"]) for e in toc.collect_outline(doc)]
    for shell in (fig_shell, tab_shell):
        assert paragraphs[shell].style.name == "EX目录壳"
        assert norm(paragraphs[shell].text) not in outline_titles, (
            "%r 被收进了主目录" % paragraphs[shell].text.strip())

    # 手写条目页码是「值」不是空槽：这是作者原始页码，不经流水线回填
    for row in fig_rows + tab_rows:
        assert split_at_tab(row)[1], "手写图/表目录条目不应有空页码槽"


# ================================================================ 六、分页
@pytest.mark.parametrize("vol", ["math", "physics", "chemistry", "biology",
                                 "comprehensive"])
def test_body_section_starts_with_page_break_before(vol, products, docs):
    """五卷：目录块之后的第一个正文段带 ``w:pageBreakBefore``，目录独占一页。

    实现比「独立成节」轻得多：段数与节数都不变，只改分页行为，所以这条断言
    同时约束「节数没变」（节数由流水线的段/节不变量守住）。
    """
    doc = docs[vol]
    first, last = toc_block_of(doc)
    target = last + 1
    para = doc.paragraphs[target]
    assert has_page_break_before(para), (
        "%s 分页目标段 #%d %r 没有 pageBreakBefore" % (vol, target,
                                                     para.text[:20]))
    assert para.text.strip().startswith("第Ⅰ卷") or \
        para.text.strip() == "试卷结构与分值分配", (
            "%s 分页目标段不是卷首标题：%r" % (vol, para.text[:24]))
    # 目录块自身不该带分页属性（分页在正文段上，不在目录标题上）
    assert not has_page_break_before(doc.paragraphs[first])


def test_demo_uses_section_break_not_page_break_before(products, docs):
    """demo 不靠 pbr：目录区之后的 ``sectPr`` 就是分页机制（``nextPage`` 缺省）。

    这是有意的差异：该卷目录区（目录 + 图目录 + 表目录 + 表 4）本来就跨页，
    节断比段级 pbr 更贴合作者原稿。
    """
    doc = docs["demo"]
    first, last = toc_block_of(doc)
    # 目录条目里也有「引言」这一行，所以从目录块之后往后找正文那一段。
    intro = next(i for i, p in enumerate(doc.paragraphs)
                 if i > last and p.text.strip() == "引言")
    assert intro > last
    assert any(has_sectpr(p) for p in doc.paragraphs[last:intro]), (
        "目录区与「引言」之间没有 sectPr，正文不会另起页")
    assert not has_page_break_before(doc.paragraphs[last + 1]), (
        "demo 不应额外加 pageBreakBefore")


# ============================================================ 七、Word 层（opt-in）
WORD_ENV = "DOCX_TOC_WORD_TESTS"


def _word_layer_enabled() -> bool:
    """Word 层默认**不跑**（Phase 3-A 纪律），需显式 opt-in。

    三道闸：显式环境变量、Windows、``win32com`` 可导入、``run_accept.py`` 在位。
    三者任一不满足就 skip —— CI 上没有 Office，跳过而不是失败。
    """
    if os.environ.get(WORD_ENV, "").strip().lower() not in ("1", "true", "yes", "on"):
        return False
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return RUN_ACCEPT.exists()


needs_word = pytest.mark.skipif(
    not _word_layer_enabled(),
    reason="Word 层默认跳过：需本机 Office + win32com + %s=1" % WORD_ENV)


def word_index_map(doc) -> list:
    """``doc.paragraphs`` 索引 -> Word ``Paragraphs(n)`` 的 1-based 序号。

    约束 #29：Word 的索引空间含**表格单元格里的段落**与**每行行尾标记**，
    与 ``doc.paragraphs`` 不同源；拿 0-based 索直加 1 会静默读到别的段落
    （实测 biology 148 vs 108、comprehensive 209 vs 100、demo 297 vs 167）。
    """
    mapping, counter = [], 1

    def walk_table(table):
        nonlocal counter
        for row in table.findall(qn("w:tr")):
            for cell in row.findall(qn("w:tc")):
                for child in cell:
                    if child.tag == qn("w:p"):
                        counter += 1
                    elif child.tag == qn("w:tbl"):
                        walk_table(child)
            counter += 1                    # 行尾标记

    for child in doc.element.body:
        if child.tag == qn("w:p"):
            mapping.append(counter)
            counter += 1
        elif child.tag == qn("w:tbl"):
            walk_table(child)
        elif child.tag == qn("w:sdt"):
            content = child.find(qn("w:sdtContent"))
            for sub in ([] if content is None else content):
                if sub.tag == qn("w:p"):
                    mapping.append(counter)
                    counter += 1
    return mapping


def probe_paragraph_pages(docx_path: Path, indices, tmp_path: Path) -> dict:
    """走 ``run_accept.py``（跨进程 Word 令牌锁）跑 ``--page-probe``。"""
    out = tmp_path / ("probe-%s.json" % Path(docx_path).stem)
    command = [
        sys.executable, "-X", "utf8", str(RUN_ACCEPT), str(docx_path),
        "--json", str(out),
        "--page-probe", ",".join(str(i) for i in indices),
        "--no-update-fields", "--quiet",
    ]
    proc = subprocess.run(command, cwd=str(RUN_ACCEPT.parent), timeout=1800,
                          capture_output=True, text=True, encoding="utf-8",
                          errors="replace")
    assert out.exists(), ("word_accept 没有产出 JSON（rc=%s）\n%s"
                          % (proc.returncode, (proc.stderr or "")[-1500:]))
    return json.loads(out.read_text(encoding="utf-8"))


@needs_word
@pytest.mark.parametrize("vol", VOLUMES)
def test_word_page_numbers_close_the_loop(vol, products, docs, tmp_path):
    """页码闭环：目录里的页码 == 该条标题段在 Word 里的**显示**页号。

    约束 #28：比 ``displayed_text``（保留 ``i``/``iii`` 字形）；
    ``displayed_number`` 已数值化，拿来比会把能源卷的 ``i`` 判成「不闭环」。
    约束 #31：``--pages`` 回填取的是「标题段」页码 —— 所以这里探的也是标题段，
    不是 ``toc N`` 条目段（条目段全落目录页，首遍行末是空槽）。
    """
    doc = docs[vol]
    outline = toc.collect_outline(doc)
    entries = entries_of(doc)
    assert len(outline) == len(entries) == EXPECT[vol]["entries"]
    mapping = word_index_map(doc)
    indices = [mapping[e["para_index"]] for e in outline]

    report = probe_paragraph_pages(products[vol], indices, tmp_path)
    if report.get("paragraph_pages_source") != "word":
        pytest.skip("Word 未给出实测页码（%s）" % report.get(
            "paragraph_pages_source"))

    # 约束 #32：按索引键取行，不按位置硬凑 —— 首位可能是目录标题段。
    rows = {int(row["index"]): row
            for row in (report.get("paragraph_pages") or [])}
    missing = [i for i in indices if i not in rows]
    assert not missing, "Word 未测到这些段：%s" % missing
    pages = [rows[i].get("displayed_text") for i in indices]
    if any(page is None for page in pages):
        pytest.skip("有段落没测到页码：%s" % pages)

    tokens = [split_at_tab(para)[1] for para in entries]
    assert pages == tokens, (
        "%s 目录页码与标题段实测页号不一致：\n  目录 %s\n  Word %s"
        % (vol, tokens, pages))


@needs_word
@pytest.mark.parametrize("vol", VOLUMES)
def test_word_toc_owns_its_page(vol, products, docs, tmp_path):
    """目录独占一页：目录标题所在页 < 正文首段所在页（物理页比较）。

    五卷靠 ``w:pageBreakBefore``，demo 靠节断；两者都必须成立，所以这条
    不看 pbr 属性，只看真实排版结果。
    """
    doc = docs[vol]
    first, last = toc_block_of(doc)
    mapping = word_index_map(doc)
    title_index = mapping[first]
    body_index = mapping[last + 1]

    report = probe_paragraph_pages(products[vol],
                                   [title_index, body_index], tmp_path)
    if report.get("paragraph_pages_source") != "word":
        pytest.skip("Word 未给出实测页码")
    rows = {int(row["index"]): row
            for row in (report.get("paragraph_pages") or [])}
    title_row = rows.get(title_index) or {}
    body_row = rows.get(body_index) or {}
    if title_row.get("physical") is None or body_row.get("physical") is None:
        pytest.skip("物理页未测到：%s / %s" % (title_row, body_row))
    assert int(body_row["physical"]) > int(title_row["physical"]), (
        "%s 目录与正文同页：目录 %r 页 %r，正文首段 %r 页 %r"
        % (vol, doc.paragraphs[first].text.strip(),
           title_row.get("physical"), doc.paragraphs[last + 1].text[:16],
           body_row.get("physical")))