# -*- coding: utf-8 -*-
"""``docx_section list`` 的三个派生字段（P17-A A2）。

## 本文件存在的理由

`docx_section list` 原来只报**输入**，不报**推论**：给一个
``page_number: {fmt, start}`` 和一条 ``header_footer[].preview`` 字符串，让主人
自己算「这一节显示第几页到第几页」「这一节的页码到底有没有落到页脚上」。

主人的真实排版作业（封面无页码 + 目录单独编号 + 正文从 1 起）里，这两件事都得
自己推，推错一次就是交稿后目录页码对不上，或者整篇没有页码而没人说过。

所以 `section_info()` 新增三个派生字段：

* ``pages`` — ``{from, to, from_text, source, reason}``；**能算就算，算不出就
  ``null`` 并写明为什么**，绝不瞎猜一个数；
* ``footer_kind`` — ``own`` / ``inherited`` / ``none``；
* ``footer_has_page_field`` — 本节实际渲染的页脚里有没有 PAGE 域。

## 硬约束：只许新增，不许改既有

`docx_section list` 是既有 CLI，已有消费者（`test_section_para.py`、
`test_section_pagenum_zh.py` 与文档）。本文件除了测新字段，还要钉住**旧字段一
个字都没变** —— 那是「增强」和「改坏」的分界线。

Run:  python -m pytest tests/test_docx_section.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from lxml import etree

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_section.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

import docx_section as ds  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_NS = "{%s}" % W

# 复用 footer 测试里的真实页脚 fixture（同一套写法，别造第二份）
sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_docx_footer import (  # noqa: E402
    NSDECL, ZH_PAGE_LINE, _splice_footer, build_plain_footer,
)


# ---- helpers -----------------------------------------------------------------

def run_list(path) -> dict:
    """以真子进程跑 ``docx_section list``，拿它真正打出来的 JSON。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), "list", str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    return json.loads(proc.stdout)


def three_sections(tmp_path: Path, name: str = "three.docx") -> Path:
    """封面（无页脚、无 pgNumType）/ 目录（罗马数字、罗马页脚）/ 正文（从 1 起）.

    这就是主人那份排版作业的骨架，也是 ``footer_kind`` 三档能同时出现的最小
    文档。
    """
    doc = Document()
    doc.add_paragraph("封面")
    ds.add_section_break(doc)
    doc.add_paragraph("目录")
    ds.set_pg_num_type(doc.sections[1], fmt="lowerRoman", start=1)
    ds.add_page_number_field(doc.sections[1], template="page")
    ds.add_section_break(doc)
    doc.add_paragraph("正文")
    ds.set_pg_num_type(doc.sections[2], fmt="decimal", start=1)
    ds.add_page_number_field(doc.sections[2], template="zh")
    path = tmp_path / name
    doc.save(str(path))
    return path


# =============================================================================
# pages
# =============================================================================
def test_pages_from_start_and_to_is_null_with_a_reason(tmp_path: Path):
    """★ ``pages.from`` 取自 ``pgNumType/@start``；``to`` 一律 null + reason。

    「本节末页」等于「起页 + 本节页数 − 1」，而 python-docx 不排版：文档里没有
    任何地方存着「这一节占几页」。给一个像模像样的数字就是编造，所以这里断言
    它**没有**编造。
    """
    src = three_sections(tmp_path)
    detail = run_list(src)["section_detail"]

    assert detail[1]["pages"]["from"] == 1
    assert detail[1]["pages"]["source"] == "pgNumType.start"
    assert detail[2]["pages"]["from"] == 1
    assert detail[2]["pages"]["to"] is None
    assert "page count" in detail[2]["pages"]["reason"]
    assert "layout" in detail[2]["pages"]["reason"]


def test_pages_from_text_renders_the_counter_format(tmp_path: Path):
    """``from_text`` 是**渲染出来**的字：``lowerRoman`` 的 1 是 ``i`` 不是 ``1``。

    只报数字会让「目录显示 i」和「目录显示 1」两件事在报告里长得一样。
    """
    src = three_sections(tmp_path)
    detail = run_list(src)["section_detail"]
    assert detail[1]["pages"]["from"] == 1
    assert detail[1]["pages"]["from_text"] == "i"          # lowerRoman
    assert detail[2]["pages"]["from_text"] == "1"          # decimal


def test_pages_without_start_continues_the_previous_counter_and_says_so(
        tmp_path: Path):
    """第 1 节起没有 ``pgNumType.start`` → ``from`` 是 **null**，不是猜出来的数。

    这一节接着上一节的计数器走，起页 = 上一节末页 + 1，而上一节末页本来就未知。
    链条断在这里就必须说断，不许拿「上一节从 1 起、大概 3 页」这种估计顶上。
    """
    doc = Document()
    doc.add_paragraph("封面")
    ds.add_section_break(doc)
    doc.add_paragraph("续页")            # 第 1 节：no pgNumType at all
    path = tmp_path / "nostart.docx"
    doc.save(str(path))

    detail = run_list(path)["section_detail"]
    assert detail[0]["pages"]["from"] == 1        # section 0 starts at 1
    assert detail[0]["pages"]["source"].startswith("document default")
    assert detail[1]["pages"]["from"] is None
    assert detail[1]["pages"]["source"] is None
    assert "continues the previous section" in detail[1]["pages"]["reason"]


# =============================================================================
# footer_kind
# =============================================================================
def test_footer_kind_is_own_inherited_and_none(tmp_path: Path):
    """三档同时出现：封面 ``none``、目录与正文 ``own``。

    再叠一个「中间节没自己的页脚」的场景，逼出 ``inherited``。
    """
    src = three_sections(tmp_path)
    detail = run_list(src)["section_detail"]
    assert detail[0]["footer_kind"] == "none"
    assert detail[1]["footer_kind"] == "own"
    assert detail[2]["footer_kind"] == "own"

    # A section inserted between two sections that already have footers: it has
    # no reference of its own, so Word reuses the preceding one.
    doc = Document(src)
    ds.add_section_break(doc, reuse_last=True)
    doc.sections[-1].footer.is_linked_to_previous = True
    mid = tmp_path / "mid.docx"
    doc.save(str(mid))
    detail = run_list(mid)["section_detail"]
    assert detail[-1]["footer_kind"] == "inherited"
    assert detail[-2]["footer_kind"] == "own"


def test_footer_kind_none_when_nothing_anywhere(tmp_path: Path):
    """全篇没有页脚 → 每节都是 ``none``（不是 ``inherited`` 的 0 号特例）。"""
    doc = Document()
    doc.add_paragraph("只有正文")
    ds.add_section_break(doc)
    doc.add_paragraph("第二节")
    path = tmp_path / "nofooter.docx"
    doc.save(str(path))
    detail = run_list(path)["section_detail"]
    assert [i["footer_kind"] for i in detail] == ["none", "none"]


# =============================================================================
# footer_has_page_field
# =============================================================================
def test_footer_has_page_field_true_for_complex_field(tmp_path: Path):
    """complex 形态（``fldChar begin/instrText/separate/end``）→ true。"""
    src = three_sections(tmp_path)
    detail = run_list(src)["section_detail"]
    assert detail[0]["footer_has_page_field"] is False
    assert detail[1]["footer_has_page_field"] is True
    assert detail[2]["footer_has_page_field"] is True


def test_footer_has_page_field_true_for_fldsimple(tmp_path: Path):
    """★ ``<w:fldSimple w:instr=" PAGE "/>`` 形态也必须算「有页码域」。

    只认 complex 的实现会把半个世界的文档判成没页码 —— 而 ``fldSimple`` 正是
    Word 与 WPS 在某些路径下真正写出来的东西。
    """
    doc = Document()
    doc.add_paragraph("正文")
    path = tmp_path / "simple.docx"
    doc.save(str(path))
    line = ('<w:p><w:r><w:t xml:space="preserve">第 </w:t></w:r>'
            '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
            '<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>')
    _splice_footer(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<w:ftr %s>%s</w:ftr>' % (NSDECL, line))
    detail = run_list(path)["section_detail"]
    assert detail[0]["footer_has_page_field"] is True


def test_footer_has_page_field_false_for_text_only_footer(tmp_path: Path):
    """页脚只有文字、没有域 → false（不能因为「页脚非空」就报 true）。"""
    doc = Document()
    doc.add_paragraph("正文")
    path = tmp_path / "textonly.docx"
    doc.save(str(path))
    _splice_footer(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<w:ftr %s><w:p><w:r><w:t>机密</w:t></w:r></w:p>'
                         '</w:ftr>' % NSDECL)
    detail = run_list(path)["section_detail"]
    assert detail[0]["footer_has_page_field"] is False


def test_broken_footer_part_fails_at_the_python_docx_boundary(tmp_path: Path):
    """**已知边界**：页脚部件不是良构 XML 时，``docx_section list`` 起不来。

    这条测的不是「我们去修好了」——``docx_section`` 走 ``Document(path)``，而
    python-docx 在打开包时就把**每一个**部件都 parse 了，所以坏页脚会在
    ``Package.open`` 里先炸，轮不到本模块。把它写成绿测试只能是假的。

    钉住两件真事：

    * 这条路确实崩（``list`` 继承 python-docx 的拒绝），报错是 python-docx 的，
      不是本模块的；
    * **读原始 XML 的那条路不受影响** —— ``docx_footer check`` 照样给报告
      （``footer-part-unreadable`` warning）。所以「想知道页脚坏没坏」有工具可
      用，不必指望 ``list``。

    真正的健壮性要求（``docx_validate`` 不能崩）由
    ``tests/test_docx_validate.py`` 覆盖，那边走的是 raw-zip 路径。
    """
    import zipfile

    doc = Document()
    doc.add_paragraph("正文")
    path = tmp_path / "brokenfooter.docx"
    doc.save(str(path))
    _splice_footer(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<w:ftr %s><w:p><w:r><w:t>x</w:t></w:r></w:p></w:ftr>'
                   % NSDECL)
    tmp = path.with_suffix(".t.docx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for n in zin.namelist():
            data = zin.read(n)
            if n == "word/footer1.xml":
                data = b"<w:ftr><unclosed>"
            zout.writestr(n, data)
    tmp.replace(path)

    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    listed = subprocess.run(
        [sys.executable, str(SCRIPT), "list", str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    assert listed.returncode != 0
    assert "python-docx-open-failed" in listed.stderr or "XMLSyntaxError" in \
        listed.stderr, listed.stderr

    # 同一条路径上的另一支：raw-zip 读取不受影响
    checked = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_footer.py"), "check", str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    assert checked.returncode == 0, checked.stderr
    rep = json.loads(checked.stdout)
    assert "footer-part-unreadable" in {i["code"] for i in rep["issues"]}


# =============================================================================
# 零回归：既有字段的名字 / 类型 / 语义
# =============================================================================
def test_existing_fields_are_untouched(tmp_path: Path):
    """旧字段逐条比对：``page_number`` 仍是 ``{fmt, start}``、``header_footer``
    仍带 ``kind/type/r_id/preview``、其余键名与类型不变。

    这是「只许新增」的唯一可执行证明 —— 少了它，一次「顺手重命名」不会有任何
    测试变红。
    """
    src = three_sections(tmp_path)
    detail = run_list(src)["section_detail"]
    sec = detail[2]

    for key in ("index", "start_type", "page_width_mm", "page_height_mm",
                "margins_mm", "page_number", "cols", "different_first_page",
                "sectPr_location", "header_footer"):
        assert key in sec, "existing key %r disappeared" % key

    assert sec["index"] == 2
    assert sec["page_number"] == {"fmt": "decimal", "start": "1"}
    assert isinstance(sec["cols"], int)
    assert isinstance(sec["different_first_page"], bool)
    assert sec["sectPr_location"] == "body"
    assert isinstance(sec["margins_mm"]["top"], float)

    foot = sec["header_footer"][0]
    assert set(foot) == {"kind", "type", "r_id", "preview"}
    assert foot["kind"] == "footer"
    assert foot["type"] == "default"
    assert foot["r_id"].startswith("rId")
    assert "[PAGE]" in foot["preview"]


def test_top_level_shape_is_unchanged(tmp_path: Path):
    """``{"ok", "sections", "section_detail"}`` 三项不变、``sections`` 仍是整数。"""
    src = three_sections(tmp_path)
    rep = run_list(src)
    assert set(rep) == {"ok", "sections", "section_detail"}
    assert rep["ok"] is True
    assert rep["sections"] == 3
    assert len(rep["section_detail"]) == 3


def test_new_fields_are_additive_only(tmp_path: Path):
    """新增的三个键就是这三个，没有夹带别的东西一起冒出来。"""
    src = three_sections(tmp_path)
    sec = run_list(src)["section_detail"][0]
    added = {"pages", "footer_kind", "footer_has_page_field"}
    assert added <= set(sec)
    assert set(sec["pages"]) == {"from", "to", "from_text", "source", "reason"}
    assert sec["footer_kind"] in ("own", "inherited", "none")
    assert isinstance(sec["footer_has_page_field"], bool)