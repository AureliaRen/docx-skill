# -*- coding: utf-8 -*-
"""`docx_footer.py` 的验收：页脚形态、框宽、页码是否被裁，以及退出码契约。

## 本文件存在的理由

主人拿来的真实文档里，页脚是 **WPS 浮动文本框 + ``spAutoFit``**，框宽被「1」
定成 **4.56 pt**，页码进位到两位就被裁。实测两件事：

1. ``docx_validate`` 对那份文档返 ``{"ok": true, "issues": []}`` —— 完全没报；
2. ``docx_section list`` 只给 ``header_footer[].preview`` 一个字符串
   （``第 1 页　共 1 页 [PAGE] [NUMPAGES]``），读不出页脚是浮动框、框有多宽。

「那 1 个字符的数字把 88 pt 的行塞进 4.56 pt 的框」这件事，只有把页脚 XML 里
的 ``wp:extent/@cx``（或 VML ``style="width:4.56pt"``）读出来、再按页脚**实际**
字号估算行宽，才能说出口。本文件就是这句话的护栏。

## fixture 是**真 XML**，不是 mock

``build_floating_footer`` 用 python-docx 起包、再用 lxml **直接构造**页脚
XML（VML ``w:pict``/``v:shape`` 与 DrawingML ``w:drawing``/``wp:anchor`` 两支），
写进真 .docx 的 ``word/footer1.xml``。用 mock 只能证明「我调了这个函数」，测不出
「4.56 pt 的框配 88 pt 的行算不算溢出」—— 而那正是本单要钉的东西。

## 为什么**不**挂 ``@needs_omml`` / ``needs_word``

全程只用 ``zipfile`` + ``lxml`` + ``python-docx``，不碰 Office COM、LibreOffice、
``MML2OMML.XSL``。无 Office 的 CI runner 上必须照常 pass。

Run:  python -m pytest tests/test_docx_footer.py -q
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_footer.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

import docx_footer as F  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
V = "urn:schemas-microsoft-com:vml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

#: The width the real template's box ended up with, in EMU: 4.56 pt.
BOX_EMU = 57912

NSDECL = ('xmlns:w="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:wps="%s" '
          'xmlns:mc="%s" xmlns:v="%s" xmlns:r="%s" '
          'xmlns:o="urn:schemas-microsoft-com:office:office"'
          % (W, WP, A, WPS, MC, V, R))

ZH_PAGE_LINE = """<w:p><w:r><w:t xml:space="preserve">第 </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:t xml:space="preserve"> 页&#x3000;共 </w:t></w:r>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>"""

VML_FOOTER = """<w:ftr {ns}>
  <w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:pict>
      <v:shape id="ftr1" type="#_x0000_t202"
               style="width:{width}pt;height:20pt"
               o:allowincell="f" filled="f" stroked="f"
               mso-fit-shape-to-text="t">
        <v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:t">
          <w:txbxContent>{line}</w:txbxContent>
        </v:textbox>
      </v:shape>
    </w:pict></w:r>
  </w:p>
</w:ftr>"""

DRAWING_FOOTER = """<w:ftr {ns}>
  <w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    <w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps">
          <w:drawing>
            <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
                       simplePos="0" relativeHeight="251658240" behindDoc="0"
                       locked="0" layoutInCell="1" allowOverlap="1">
              <wp:simplePos x="0" y="0"/>
              <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
              <wp:extent cx="{cx}" cy="254000"/>
              <wp:effectExtent l="0" t="0" r="0" b="0"/>
              <wp:wrapNone/>
              <wp:docPr id="100" name="Text Box 1"/>
              <wp:cNvGraphicFramePr/>
              <a:graphic><a:graphicData uri="{wps}">
                <wps:wsp>
                  <wps:cNvSpPr txBox="1"/>
                  <wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="254000"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    <a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>
                  <wps:txbx><w:txbxContent>{line}</w:txbxContent></wps:txbx>
                  <wps:bodyPr rot="0" vert="horz" wrap="none" lIns="0" tIns="0"
                              rIns="0" bIns="0" anchor="t" anchorCtr="0">
                    <a:spAutoFit/>
                  </wps:bodyPr>
                </wps:wsp>
              </a:graphicData></a:graphic>
            </wp:anchor>
          </w:drawing>
        </mc:Choice>
      </mc:AlternateContent>
    </w:r>
  </w:p>
</w:ftr>"""


# ---- helpers -----------------------------------------------------------------

def run_cli(*args) -> subprocess.CompletedProcess:
    """真子进程：退出码属 CLI 契约，in-process 测不到。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), *[str(a) for a in args]],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def _next_rid(rels: bytes) -> str:
    got = [int(m) for m in re.findall(rb'Id="rId(\d+)"', rels)]
    return "rId%d" % (max(got) + 1 if got else 1)


def _add_footer_ref(doc_xml: bytes, rid: str) -> bytes:
    text = doc_xml.decode("utf-8")
    idx = text.rfind("</w:sectPr>")
    head, tail = text[:idx], text[idx:]
    open_idx = head.rfind("<w:sectPr")
    close_idx = head.find(">", open_idx) + 1
    ref = '<w:footerReference w:type="default" r:id="%s"/>' % rid
    return (head[:close_idx] + ref + head[close_idx:] + tail).encode("utf-8")


def _add_rel(rels: bytes, rid: str) -> bytes:
    text = rels.decode("utf-8")
    rel = ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/footer" '
           'Target="footer1.xml"/>' % rid)
    return text.replace("</Relationships>", rel + "</Relationships>").encode("utf-8")


def _add_ctype(data: bytes) -> bytes:
    text = data.decode("utf-8")
    over = ('<Override PartName="/word/footer1.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>')
    return text.replace("</Types>", over + "</Types>").encode("utf-8")


def _splice_footer(path: Path, footer_xml: str) -> Path:
    """Replace ``word/footer1.xml`` with *footer_xml*, adding the part if absent.

    One rId, computed from the existing rels and used in both places it has to
    appear.  Computing it twice yields an id already taken by ``customXml`` and
    points the section at the wrong part — the fixture then "passes" for a
    reason that has nothing to do with the code under test.
    """
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w") as dst:
        names = src.namelist()
        have = "word/footer1.xml" in names
        rid = _next_rid(src.read("word/_rels/document.xml.rels"))
        for name in names:
            data = src.read(name)
            if name == "word/footer1.xml":
                data = footer_xml.encode("utf-8")
            elif name == "word/document.xml" and not have:
                data = _add_footer_ref(data, rid)
            elif name == "word/_rels/document.xml.rels" and not have:
                data = _add_rel(data, rid)
            elif name == "[Content_Types].xml" and not have:
                data = _add_ctype(data)
            dst.writestr(name, data)
        if not have:
            dst.writestr("word/footer1.xml", footer_xml.encode("utf-8"))
    tmp.replace(path)
    return path


def build_floating_footer(tmp_path: Path, kind: str = "drawing",
                          width_pt: float = 4.56, name: str = "floating.docx",
                          pg_num: bool = True,
                          line: str | None = None) -> Path:
    """A one-section document whose footer is a REAL floating text box.

    ``kind="drawing"`` writes ``w:drawing`` → ``wp:anchor`` → ``wps:txbx`` (box
    width from ``wp:extent/@cx``); ``kind="vml"`` writes ``w:pict`` →
    ``v:shape`` → ``v:textbox`` (box width from ``style="width:…pt"``).  Both are
    what real writers emit, and both must be readable.

    *line* overrides the footer body.  The default is the Chinese page line
    (``第 {PAGE} 页　共 {NUMPAGES} 页``); pass plain text to build a box that
    overflows **without carrying a numeric field** — the P18 shape that made
    ``page_number_widths_pt`` ``None`` while ``overflow`` was True, and crashed
    the CLI on ``widths[digits]``.
    """
    body_line = ZH_PAGE_LINE if line is None else line
    doc = Document()
    doc.add_paragraph("正文一行。")
    sectPr = doc.sections[0]._sectPr
    if pg_num:
        el = etree.SubElement(sectPr, "{%s}pgNumType" % W)
        el.set("{%s}start" % W, "1")
    path = tmp_path / name
    doc.save(str(path))

    if kind == "vml":
        body = VML_FOOTER.format(ns=NSDECL, width="%g" % width_pt,
                                 line=body_line)
    else:
        body = DRAWING_FOOTER.format(ns=NSDECL,
                                     cx=int(round(width_pt * 12700)),
                                     wps=WPS, line=body_line)
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + body)
    return _splice_footer(path, footer)


def build_plain_footer(tmp_path: Path, name: str = "plain.docx",
                       pg_num: bool = False) -> Path:
    """A one-section document with a normal (non-text-box) Chinese page footer."""
    doc = Document()
    doc.add_paragraph("正文一行。")
    if pg_num:
        el = etree.SubElement(doc.sections[0]._sectPr, "{%s}pgNumType" % W)
        el.set("{%s}start" % W, "1")
    path = tmp_path / name
    doc.save(str(path))
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:ftr %s>%s</w:ftr>' % (NSDECL, ZH_PAGE_LINE))
    return _splice_footer(path, footer)


def build_fldsimple_footer(tmp_path: Path, name: str = "simple.docx") -> Path:
    """A footer whose page field is the ``w:fldSimple`` form, not complex."""
    doc = Document()
    doc.add_paragraph("正文一行。")
    path = tmp_path / name
    doc.save(str(path))
    line = ('<w:p><w:r><w:t xml:space="preserve">第 </w:t></w:r>'
            '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>'
            '<w:r><w:t xml:space="preserve"> 页</w:t></w:r></w:p>')
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:ftr %s>%s</w:ftr>' % (NSDECL, line))
    return _splice_footer(path, footer)


# =============================================================================
# ① 普通页脚 + complex 形态的页码域 → 无 issue
# =============================================================================
def test_plain_footer_with_page_field_has_no_issues(tmp_path: Path):
    """普通页脚（无文本框）+ ``field begin/instrText/separate/end`` → ok、exit 0。

    这是「不该报」的那一边：加了检查之后，最普通的页脚一行都不许响。
    """
    src = build_plain_footer(tmp_path)
    proc = run_cli("check", src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert rep["issues"] == []

    foot = rep["sections"][0]["footers"][0]
    assert foot["shape"] == "plain"
    assert foot["type"] == "default"
    assert foot["part"] == "word/footer1.xml"
    assert foot["has_page_field"] is True
    assert foot["has_numpages_field"] is True
    assert foot["overflow"] is False
    # 第 1 页　共 1 页 / 两位 / 三位都要给，且必须单调变宽
    widths = foot["page_number_widths_pt"]
    assert widths["1"] < widths["12"] < widths["123"]
    assert foot["page_number_width_pt"] == widths["12"]


# =============================================================================
# ② 浮动文本框页脚、框宽 4.56 pt → overflow true + exit 1
# =============================================================================
@pytest.mark.parametrize("kind", ["drawing", "vml"])
def test_narrow_floating_textbox_overflows_and_exits_1(tmp_path: Path, kind: str):
    """★ 本单的核心真问题：4.56 pt 的浮动框装不下 ``第 1 页　共 1 页``。

    两条写入路径都要抓：DrawingML（``wp:extent/@cx="57912"``）与 VML
    （``style="width:4.56pt"``）。框宽必须是**实测**出来的 4.56 —— 不是从字号
    反推、不是四舍五入成 5，就是 4.56。
    """
    src = build_floating_footer(tmp_path, kind=kind, width_pt=4.56,
                                name="%s-narrow.docx" % kind)
    proc = run_cli("check", src)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False

    foot = rep["sections"][0]["footers"][0]
    assert foot["shape"] == "floating-textbox"
    assert foot["box_width_pt"] == pytest.approx(4.56, abs=0.01)
    assert foot["overflow"] is True
    # 一位数字就已经装不下：行宽 ~77 pt 对 4.56 pt 的框
    assert foot["overflow_at_digits"] == "1"
    assert foot["page_number_width_pt"] > foot["box_width_pt"]

    codes = [i["code"] for i in rep["issues"]]
    assert "footer-textbox-overflow" in codes
    errs = [i for i in rep["issues"] if i["severity"] == "error"]
    assert errs and errs[0]["code"] == "footer-textbox-overflow"
    assert "4.56" in errs[0]["detail"]


def test_wide_floating_textbox_does_not_overflow(tmp_path: Path):
    """反向对照：同样的浮动框、200 pt 宽 → 不溢出、exit 0。

    没有这一条，上面那条「溢出」可能只是「凡是浮动框都报」的同义反复。
    """
    src = build_floating_footer(tmp_path, kind="drawing", width_pt=200,
                                name="wide.docx")
    proc = run_cli("check", src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    foot = rep["sections"][0]["footers"][0]
    assert foot["shape"] == "floating-textbox"
    assert foot["box_width_pt"] == pytest.approx(200.0, abs=0.01)
    assert foot["overflow"] is False
    assert rep["issues"] == []


def test_floating_box_reports_spautofit_and_wrap(tmp_path: Path):
    """``spAutoFit`` 与 ``wrap`` 是这套 bug 的两个成因，必须报出来。

    真实模板就是 ``spAutoFit``（框按内容缩，于是「1」把框定成 4.56 pt）配
    ``wrap="none"``（浮动不参与重排，裁掉就是裁掉）。
    """
    src = build_floating_footer(tmp_path, kind="drawing", width_pt=4.56)
    rep = json.loads(run_cli("check", src).stdout)
    foot = rep["sections"][0]["footers"][0]
    assert foot["spAutoFit"] is True
    assert foot["wrap"] == "none"
    box = foot["textboxes"][0]
    assert box["floating"] is True
    assert box["spAutoFit"] is True
    assert box["wrap"] == "none"


# =============================================================================
# ③ --section 越界 → exit 2 + stderr 一行
# =============================================================================
def test_section_out_of_range_is_a_usage_error(tmp_path: Path):
    """``--section 9``（只有 1 节）→ exit 2、stderr 恰好一行、**stdout 空**。

    越界是用法错误而不是文档问题：用户点了一节不存在的节。这也是本项目 exit 2
    的定义（``an --index / --insert-at outside the legal range``）。
    """
    src = build_plain_footer(tmp_path)
    proc = run_cli("check", src, "--section", 9)
    assert proc.returncode == 2, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stdout == ""
    line = one_error_line(proc.stderr)
    assert line.startswith("error: ")
    assert "no such section 9" in line


def test_inherited_footer_issue_is_attributed_to_the_inheriting_section(
        tmp_path: Path):
    """★ 页脚是**共享**的：继承它的那一节也必须报，``--section N`` 不许说谎。

    实测到的假阴性：第 1 节没有自己的页脚，渲染的是第 0 节那份 4.56 pt 的窄框，
    页码照样被裁。原先按「部件名」去重，于是缺陷只记在第 0 节名下，``--section
    1`` 回答 ``ok: true`` —— 那一页的页码仍然是断的。这正是本包要消掉的那类
    假阴性，所以单独钉一条。
    """
    src = build_floating_footer(tmp_path, kind="drawing", width_pt=4.56,
                                name="shared.docx")
    doc = Document(str(src))
    from docx.enum.section import WD_SECTION
    doc.add_section(WD_SECTION.NEW_PAGE)
    doc.add_paragraph("第二节")
    two = tmp_path / "shared2.docx"
    doc.save(str(two))

    rep = json.loads(run_cli("check", two).stdout)
    kinds = [s["footer_kind"] for s in rep["sections"]]
    assert kinds == ["own", "inherited"]
    assert rep["ok"] is False
    assert sorted(i["section"] for i in rep["issues"]) == [0, 1]

    only0 = json.loads(run_cli("check", two, "--section", "0").stdout)
    assert only0["ok"] is False
    assert [i["section"] for i in only0["issues"]] == [0]

    only1 = json.loads(run_cli("check", two, "--section", "1").stdout)
    assert only1["ok"] is False, "an inherited clipped footer is still clipped"
    assert [i["section"] for i in only1["issues"]] == [1]


def test_bare_pgnumtype_element_is_not_a_finding(tmp_path: Path):
    """裸 ``<w:pgNumType/>``（什么也没设）+ 只有文字的页脚 → **不报**。

    实测到的误报：``docx_section set_pg_num_type`` 的契约是「保证元素存在且顺序
    合法」，不传 fmt/start 时写出来的就是这个裸元素。把它当成「编号设了却不显
    示」，会让每一个经过那条路径的文档都吃一条 warning —— 而它根本什么也没配
    置。命中条件因此要求 ``w:fmt`` 或 ``w:start`` 至少有一个。
    """
    doc = Document()
    doc.add_paragraph("正文。")
    etree.SubElement(doc.sections[0]._sectPr, "{%s}pgNumType" % W)
    path = tmp_path / "bare.docx"
    doc.save(str(path))
    line = '<w:p><w:r><w:t>机密</w:t></w:r></w:p>'
    _splice_footer(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<w:ftr %s>%s</w:ftr>' % (NSDECL, line))

    proc = run_cli("check", path)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep["issues"] == []
    # the element is still reported as present — it is only the *finding* that
    # must not fire
    assert rep["sections"][0]["pg_num_type"] == {"fmt": None, "start": None}


def test_section_flag_selects_one_and_keeps_exit_semantics(tmp_path: Path):
    """``--section 0`` 在单节文档上照常工作，退出码仍由 error 级问题决定。"""
    src = build_floating_footer(tmp_path, width_pt=4.56)
    proc = run_cli("check", src, "--section", 0)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert len(rep["sections"]) == 1
    assert rep["issues"], "filtering to the offending section must keep its issue"


# =============================================================================
# ④ w:fldSimple 形态的页码域也要认出来
# =============================================================================
def test_fldsimple_page_field_is_recognised(tmp_path: Path):
    """``<w:fldSimple w:instr=" PAGE "/>`` 与 complex 形态必须同样算「有页码域」。

    只认 complex 的实现会把这半个世界的文档判成「没有页码」。
    """
    src = build_fldsimple_footer(tmp_path)
    proc = run_cli("check", src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    foot = rep["sections"][0]["footers"][0]
    assert foot["has_page_field"] is True
    assert "PAGE" in foot["fields"]
    assert foot["page_number_widths_pt"]["12"] > 0


def test_field_keyword_handles_switches_and_nesting():
    """``field_keyword`` 只取第一段，指令里的开关（``\\* MERGEFORMAT``）不影响。"""
    assert F.field_keyword(" PAGE ") == "PAGE"
    assert F.field_keyword(" NUMPAGES  \\* MERGEFORMAT ") == "NUMPAGES"
    assert F.field_keyword("PAGEREF _Toc1 \\h") == "PAGEREF"
    assert F.field_keyword("") == ""
    assert F.field_keyword(None) == ""


def test_nested_complex_fields_are_read_per_level():
    """嵌套域（``IF`` 里套 ``PAGE``）必须各归各位，不能拼成一个关键字。

    拼错的形态是 ``IFPAGE``/``PAGEIF`` 之类 —— 正则做法一定出这个结果，所以这
    条钉的是「用栈而不是用字符串拼接」。
    """
    xml = ('<w:ftr %s><w:p>'
           '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
           '<w:r><w:instrText> IF </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
           '<w:r><w:instrText> PAGE </w:instrText></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
           '<w:r><w:t>1</w:t></w:r>'
           '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
           '<w:r><w:t> = 1 "a" "b"</w:t></w:r>'
           '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
           '<w:r><w:t>a</w:t></w:r>'
           '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
           '</w:p></w:ftr>' % NSDECL)
    root = etree.fromstring(xml.encode("utf-8"))
    assert F.field_keywords(root) == ["PAGE", "IF"]


def test_no_page_field_warning_when_numbering_configured(tmp_path: Path):
    """``pgNumType`` 设了、页脚有文字但没有 PAGE 域 → ``footer-no-page-field``。

    severity 是 **warning**（不是 error）：这是合法中间态（编号先设、页脚后填），
    而且它不该把文档判成 error 级 —— 真造成内容丢失的是溢出那条。
    """
    doc = Document()
    doc.add_paragraph("正文。")
    el = etree.SubElement(doc.sections[0]._sectPr, "{%s}pgNumType" % W)
    el.set("{%s}start" % W, "1")
    path = tmp_path / "nofield.docx"
    doc.save(str(path))
    line = '<w:p><w:r><w:t>机密</w:t></w:r></w:p>'
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<w:ftr %s>%s</w:ftr>' % (NSDECL, line))
    _splice_footer(path, footer)

    proc = run_cli("check", path)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    hits = [i for i in rep["issues"] if i["code"] == "footer-no-page-field"]
    assert len(hits) == 1
    assert hits[0]["severity"] == "warning"
    assert rep["ok"] is True, "a warning must not flip ok to False"


# =============================================================================
# ⑤ 非 docx 输入 → exit 2
# =============================================================================
def test_missing_path_is_a_usage_error(tmp_path: Path):
    """缺文件 → exit 2 + 一行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such.docx"
    proc = run_cli("check", missing)
    assert proc.returncode == 2, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stdout == ""
    line = one_error_line(proc.stderr)
    assert "no such file" in line and str(missing) in line


def test_directory_is_a_usage_error(tmp_path: Path):
    """传目录 → exit 2 + 一行 ``not a file`` + stdout 空。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()
    proc = run_cli("check", adir)
    assert proc.returncode == 2, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stdout == ""
    assert one_error_line(proc.stderr) == f"error: not a file: {adir}"


def test_non_zip_is_a_usage_error(tmp_path: Path):
    """改名成 .docx 的纯文本 → exit 2 + 一行 ``not a ZIP package`` + stdout 空。"""
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("this is not a zip file", encoding="utf-8")
    proc = run_cli("check", bogus)
    assert proc.returncode == 2, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stdout == ""
    assert one_error_line(proc.stderr) == f"error: not a ZIP package: {bogus}"


def test_zip_without_word_document_is_a_usage_error(tmp_path: Path):
    """合法 ZIP、但不是 Word 包（缺 ``word/document.xml``）→ exit 2。

    与「包是真的、页脚坏了」是两件事：前者是路径给错，后者要给报告。
    """
    notword = tmp_path / "notword.docx"
    with zipfile.ZipFile(notword, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
    proc = run_cli("check", notword)
    assert proc.returncode == 2, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stdout == ""


# =============================================================================
# 只读性 / 不崩 / 语料零回归
# =============================================================================
def test_check_is_read_only(tmp_path: Path):
    """``check`` 不许碰输入：跑完字节完全一致，目录里也不许多出文件。"""
    src = build_floating_footer(tmp_path, width_pt=4.56)
    before = src.read_bytes()
    listing = sorted(p.name for p in tmp_path.iterdir())
    run_cli("check", src)
    assert src.read_bytes() == before
    assert sorted(p.name for p in tmp_path.iterdir()) == listing


def test_unreadable_footer_part_warns_instead_of_crashing(tmp_path: Path):
    """页脚部件不是良构 XML → warning ``footer-part-unreadable``，不 traceback。

    ``docx_validate`` 的定位是「读得懂就报，读不懂也不能崩」；这条保证坏页脚不
    会把它的 JSON 报告变成 lxml 栈帧。
    """
    src = build_plain_footer(tmp_path, name="broken.docx")
    broken = tmp_path/"broken2.docx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(broken, "w") as zout:
        for name in zin.namelist():
            data = zin.read(name)
            if name == "word/footer1.xml":
                data = b"<w:ftr><unclosed>"
            zout.writestr(name, data)

    proc = run_cli("check", broken)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    codes = [i["code"] for i in rep["issues"]]
    assert "footer-part-unreadable" in codes
    assert all(i["severity"] == "warning" for i in rep["issues"])


def test_estimate_model_is_declared_in_the_report(tmp_path: Path):
    """估算模型必须随报告一起输出 —— 引用宽度时要能一起引它的口径。"""
    src = build_plain_footer(tmp_path)
    rep = json.loads(run_cli("check", src).stdout)
    assert "estimate_model" in rep
    assert "em" in rep["estimate_model"]


def test_char_width_model_is_what_it_says():
    """宽度模型自检：全角 1 em、数字 0.5 em、空格 0.25 em。"""
    assert F.char_em("第") == 1.0
    assert F.char_em("\u3000") == 1.0
    assert F.char_em("1") == 0.5
    assert F.char_em(" ") == 0.25
    assert F.text_width_pt("第", 11.0) == 11.0
    assert F.text_width_pt("12", 10.0) == 10.0


def test_pt_from_measure_converts_every_unit_word_writes():
    """框宽的三种真实写法都要读对：EMU、pt、cm。"""
    assert F._pt_from_measure("57912") == pytest.approx(4.56, abs=0.01)
    assert F._pt_from_measure("4.56pt") == pytest.approx(4.56, abs=0.01)
    assert F._pt_from_measure("0.16cm") == pytest.approx(4.54, abs=0.05)
    assert F._pt_from_measure(None) is None
    assert F._pt_from_measure("") is None


def test_unknown_box_width_is_not_treated_as_zero(tmp_path: Path):
    """没有 ``wp:extent`` 的框：``box_width_pt`` 是 null，**不许**当 0 报溢出。

    未知不等于零。把未知当零会让每一个没量到宽度的框都变成假溢出 —— 这正是
    「加了检查就误报」最常见的那一档。
    """
    no_extent = DRAWING_FOOTER.replace('<wp:extent cx="{cx}" cy="254000"/>', "")
    footer = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              + no_extent.format(ns=NSDECL, cx=BOX_EMU, wps=WPS,
                                 line=ZH_PAGE_LINE))
    doc = Document()
    doc.add_paragraph("正文。")
    src = tmp_path / "no-extent.docx"
    doc.save(str(src))
    _splice_footer(src, footer)

    proc = run_cli("check", src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    foot = rep["sections"][0]["footers"][0]
    assert foot["box_width_pt"] is None
    assert foot["overflow"] is False


def test_narrow_box_overflow_is_reported_by_docx_validate(tmp_path: Path):
    """A3 接线：同一份文档，``docx_validate`` 也要把溢出报成 error。

    主人原来的痛点是「docx_validate 完全没报」。这条钉的就是那条闭环：同一个
    检测、两个入口，结论必须一致。
    """
    validate = SCRIPTS / "docx_validate.py"
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    src = build_floating_footer(tmp_path, width_pt=4.56, name="wire.docx")
    proc = subprocess.run(
        [sys.executable, str(validate), str(src)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    hits = [i for i in rep["issues"] if i["code"] == "footer-textbox-overflow"]
    assert len(hits) == 1
    assert hits[0]["severity"] == "error"
# ---------------------------------------------------------------------------
# P18 regression: a box that overflows WITHOUT carrying a numeric field
# ---------------------------------------------------------------------------
#
# The shape that crashed the CLI.  ``overflow`` is decided from ``widths`` for
# every box, but the widths used to be published only when the box contained a
# PAGE/NUMPAGES field.  So a floating box holding plain text, too narrow for that
# text, produced ``overflow: True`` together with ``page_number_widths_pt: None``,
# and the reporting path indexed the dict anyway:
#
#     docx_footer check : exit=1, 13-line traceback, stdout 0 bytes
#     docx_validate     : exit=0, {"ok": true} + one warning (the crash, swallowed)
#
# Both halves are pinned below: the CLI must report, and validate must escalate.

PLAIN_LONG_LINE = (
    '<w:p><w:r><w:t xml:space="preserve">'
    '\u8fd9\u662f\u4e00\u884c\u5f88\u957f\u7684\u9875\u811a\u6587\u5b57'
    '</w:t></w:r></w:p>'
)


def test_text_only_box_overflow_reports_instead_of_crashing(tmp_path: Path):
    """No numeric field + too narrow => a finding, never a traceback."""
    src = build_floating_footer(tmp_path, width_pt=4.56, name="plain.docx",
                                line=PLAIN_LONG_LINE)
    proc = run_cli("check", src)

    # D-030: an expected failure is exit 1 with a report, not a traceback.
    assert "Traceback" not in (proc.stderr or ""), proc.stderr
    assert proc.stdout.strip(), "stdout must carry the report"
    assert proc.returncode == 1, (proc.returncode, proc.stderr)

    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    box = rep["sections"][0]["footers"][0]["textboxes"][0]
    assert box["overflow"] is True
    # The widths must accompany the verdict that was derived from them.
    assert box["page_number_widths_pt"], (
        "overflow was decided from widths, so widths must be published")
    assert box["page_number_widths_pt"][box["overflow_at_digits"]] is not None
    assert box["has_page_field"] is False      # the shape under test
    hits = [i for i in rep["issues"] if i["code"] == "footer-textbox-overflow"]
    assert len(hits) == 1
    assert "width unknown" not in hits[0]["detail"]


def test_text_only_overflow_is_an_error_in_docx_validate(tmp_path: Path):
    """The crash used to be downgraded to ``{"ok": true}``. It must be an error."""
    src = build_floating_footer(tmp_path, width_pt=4.56, name="plain2.docx",
                                line=PLAIN_LONG_LINE)
    validate = SCRIPTS / "docx_validate.py"
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    proc = subprocess.run(
        [sys.executable, str(validate), str(src)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    codes = [i["code"] for i in rep["issues"]]
    assert "footer-textbox-overflow" in codes, codes
    # And the swallowed-crash code must NOT be what surfaced it.
    assert "footer-check-skipped" not in codes, codes


def test_widths_are_published_whenever_overflow_is_decided(tmp_path: Path):
    """The invariant behind the P0: verdict and evidence come from one source.

    Checks both branches — with a field and without — because the bug was
    exactly a disagreement between them.
    """
    for name, line, expect_field in (
            ("with.docx", None, True),
            ("without.docx", PLAIN_LONG_LINE, False)):
        src = build_floating_footer(tmp_path, width_pt=4.56, name=name, line=line)
        rep = json.loads(run_cli("check", src).stdout)
        box = rep["sections"][0]["footers"][0]["textboxes"][0]
        assert box["has_page_field"] is expect_field
        if box["overflow"]:
            assert box["page_number_widths_pt"], name
            assert box["page_number_width_pt"] is not None, name
