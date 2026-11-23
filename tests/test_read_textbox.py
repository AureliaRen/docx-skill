# -*- coding: utf-8 -*-
"""``docx_read`` 读出**浮动文本框**里的文字（P19-B）。

## 本文件存在的理由

真实事故：一份 WPS 做的阅读批注作业，43 条批注全部写在**浮动文本框**里。
有人用 ``python-docx`` 遍历 ``paragraph.runs`` 读不到任何批注文字，于是判定
「43 处标记全是色块、没有一处批注」，把一项核心成果误判成缺陷。他事后想用
``docx_read --blocks`` 交叉验证 —— **``--blocks`` 一样读不到**。

根因：文本框的段落住在 ``w:txbxContent`` 里，它的父链只可能是
``wps:txbx``（DrawingML ``w:drawing`` 内）或 ``v:textbox``（VML ``w:pict`` 内），
**永远不是 ``w:body`` 的直属子元素**。所以 ``body.iterchildren()`` 看不见它，
``Paragraph.text`` 也只拼自己的 ``w:r`` 子元素 —— 两条路都不通。

## 去重是本文件的第一判据

Word / WPS 写一个框时会同时写 ``mc:Choice``（DrawingML）与 ``mc:Fallback``
（VML），**两支各存一份同样的文本**。直接遍历容器会把 43 条读成 86 条 ——
这比读不到更糟，因为它看起来"读到了"。所以：

* 一个 ``mc:AlternateContent`` 只读**一支**（有文本的 ``mc:Choice`` 优先）；
* 读成 86 条是失败，读成 172 条（两层都遍历）也是失败。

## fixture 是**真 XML**，不是 mock

``_splice_body`` 用 python-docx 起包、再用 lxml/字符串**直接构造**正文 XML
（三种容器形态 + AlternateContent 双支），写进真 .docx。用 mock 只能证明
「我调了这个函数」，测不出「两支都收会变 86」。

## 为什么**不**挂 ``@needs_omml`` / ``needs_word``

全程只用 ``zipfile`` + ``python-docx``，不碰 Office COM、LibreOffice、
``MML2OMML.XSL``。无 Office 的 CI runner 上必须照常 pass。

Run:  python -m pytest tests/test_read_textbox.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_read.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
V = "urn:schemas-microsoft-com:vml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

NSDECL = ('xmlns:w="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:wps="%s" '
          'xmlns:mc="%s" xmlns:v="%s" xmlns:r="%s"'
          % (W, WP, A, WPS, MC, V, R))

#: 注入片段的根元素必须自带声明：python-docx 生成的 ``document.xml`` 只声明了
#: ``w`` / ``wp`` / ``wps`` / ``mc`` / ``v``，**没有** DrawingML 主命名空间 ``a``。
#: 少了它，``a:graphic`` 会让整个 document.xml 解析失败（XMLSyntaxError），
#: 被测脚本连文档都打不开 —— 那是 fixture 的错，不是被测代码的错。
FRAG_NS = (' xmlns:wp="%s" xmlns:a="%s" xmlns:wps="%s" xmlns:mc="%s" '
           'xmlns:v="%s"' % (WP, A, WPS, MC, V))


# =============================================================================
# XML 片段：三种容器形态，各自"到达 Word"的那条链
# =============================================================================
def box_para(text: str) -> str:
    """One box paragraph: a single run holding *text*."""
    if text is None:
        return "<w:p/>"
    return ('<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % text)


def drawing_box(content: str, wrap: str = "anchor") -> str:
    """A DrawingML box: ``w:drawing`` -> ``wp:anchor``/``wp:inline`` -> ``wps:txbx``.

    ``wrap="anchor"`` is the floating form (what annotation tools emit);
    ``wrap="inline"`` is the in-line form. Both reach Word, and both must be read.
    """
    host_open = ('<wp:anchor distT="0" distB="0" distL="0" distR="0" '
                 'simplePos="0" relativeHeight="251659264" behindDoc="0" '
                 'locked="0" layoutInCell="1" allowOverlap="1">'
                 '<wp:simplePos x="0" y="0"/>'
                 '<wp:positionH relativeFrom="page"><wp:posOffset>35560</wp:posOffset></wp:positionH>'
                 '<wp:positionV relativeFrom="line"><wp:posOffset>0</wp:posOffset></wp:positionV>')
    host_close = "</wp:anchor>"
    if wrap == "inline":
        host_open = '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        host_close = "</wp:inline>"
    return ('<w:p><w:r><w:drawing>%s'
            '<wp:extent cx="1151890" cy="720090"/>'
            '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            '<wp:wrapNone/><wp:docPr id="5001" name="Note5001"/>'
            '<a:graphic><a:graphicData uri="%s"><wps:wsp>'
            '<wps:cNvSpPr txBox="1"/>'
            '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1152000" cy="720000"/></a:xfrm></wps:spPr>'
            '<wps:txbx><w:txbxContent>%s</w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="none"/>'
            '</wps:wsp></a:graphicData></a:graphic>%s'
            '</w:drawing></w:r></w:p>'
            % (host_open, WPS, content, host_close))


def vml_box(content: str) -> str:
    """A VML box: ``w:pict`` -> ``v:shape`` -> ``v:textbox`` -> ``w:txbxContent``.

    The ``w:txbxContent`` wrapper is not decoration: it is the element the reader
    keys on, and it is what real VML (measured on a WPS-produced file) always
    has. A fixture that skipped it would test a shape that never occurs.
    """
    return ('<w:p><w:r><w:pict>'
            '<v:shape id="Note5001" type="#_x0000_t202" '
            'style="position:absolute;left:0pt;height:56.7pt;width:90.7pt" '
            'filled="f" stroked="f">'
            '<v:textbox inset="0mm,0mm,0mm,0mm">'
            '<w:txbxContent>%s</w:txbxContent>'
            '</v:textbox></v:shape></w:pict></w:r></w:p>'
            % content)


def alternate(choice_inner: str, fallback_inner: str) -> str:
    """One box written the way Word/WPS actually write it: two branches.

    Both branches carry the **same** text. A reader that walks every container
    reports this as two boxes; a reader that honours ``mc:AlternateContent``
    reports one.
    """
    return ('<w:p><w:r><mc:AlternateContent>'
            '<mc:Choice Requires="wps">%s</mc:Choice>'
            '<mc:Fallback>%s</mc:Fallback>'
            '</mc:AlternateContent></w:r></w:p>'
            % (choice_inner, fallback_inner))


# =============================================================================
# fixture 装配
# =============================================================================
def _declare_ns(text: str) -> str:
    """Add the DrawingML main namespace to the ``w:document`` root if absent.

    Only ``xmlns:a`` is ever missing: python-docx already declares ``wp`` /
    ``wps`` / ``mc`` / ``v``. Adding one declaration at the root is what makes
    the injected ``a:graphic`` resolvable without touching the fragment.
    """
    if "xmlns:a=" in text:
        return text
    marker = "<w:document "
    at = text.find(marker)
    assert at >= 0, "not a w:document part"
    end = at + len(marker)
    return text[:end] + 'xmlns:a="%s" ' % A + text[end:]


def _splice_body(path: Path, fragment: str) -> Path:
    """Insert *fragment* (whole ``w:p`` elements) before the final ``w:sectPr``.

    Byte-in / byte-out on every other ZIP entry, so the fixture differs from a
    plain document only by the injected XML -- the control in each test stays a
    real comparison rather than "some other document".

    The namespaces go on the ``w:document`` root, not on the injected paragraph:
    python-docx's generated part declares ``w`` / ``wp`` / ``wps`` / ``mc`` / ``v``
    but **not** the DrawingML main namespace ``a``, and a fragment that declares
    it locally poisons the sibling paragraphs that follow (a ``w:p`` carrying a
    stray ``xmlns:a`` is still fine, but repeating the trick per fragment makes
    the part unreadable in practice). Declaring once at the root keeps every
    injected element in one namespace scope.
    """
    tmp = path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w",
                                                       zipfile.ZIP_DEFLATED) as dst:
        for name in src.namelist():
            data = src.read(name)
            if name == "word/document.xml":
                text = data.decode("utf-8")
                text = _declare_ns(text)
                idx = text.rfind("<w:sectPr")
                if idx < 0:
                    idx = text.rfind("</w:body>")
                text = text[:idx] + fragment + text[idx:]
                data = text.encode("utf-8")
            dst.writestr(name, data)
    tmp.replace(path)
    return path


def build(tmp_path: Path, fragment: str = "", name: str = "doc.docx",
          paragraphs=("正文第一段。",)) -> Path:
    """A real .docx whose body optionally carries *fragment*."""
    doc = Document()
    for line in paragraphs:
        doc.add_paragraph(line)
    path = tmp_path / name
    doc.save(str(path))
    if fragment:
        _splice_body(path, fragment)
    return path


def run_read(*args):
    """真子进程：退出码与 stdout/stderr 都是 CLI 契约的一部分。"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), *[str(a) for a in args]],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def payload(*args) -> dict:
    cp = run_read(*args)
    assert cp.returncode == 0, (cp.returncode, cp.stderr, cp.stdout)
    return json.loads(cp.stdout)


def textbox_blocks(doc_path: Path) -> list:
    return [b for b in payload(doc_path, "--blocks")["blocks"]
            if b["type"] == "textbox"]


# =============================================================================
# ① 三种容器形态：各自都能读到
# =============================================================================
def test_floating_drawing_box_is_read(tmp_path: Path):
    """DrawingML 浮动框（``wp:anchor`` → ``wps:txbx``）——批注工具的主力形态。"""
    src = build(tmp_path, drawing_box(box_para("使某人措手不及")))
    boxes = textbox_blocks(src)
    assert len(boxes) == 1
    assert boxes[0]["text"] == "使某人措手不及"


def test_inline_drawing_box_is_read(tmp_path: Path):
    """DrawingML 内联框（``wp:inline`` → ``wps:txbx``）也必须读到。

    内联框和浮动框走的是 ``w:drawing`` 下的两个不同宿元素；只认 ``wp:anchor``
    的写法在内联框上会静默返回 0 条。
    """
    src = build(tmp_path, drawing_box(box_para("内联框文字"), wrap="inline"))
    boxes = textbox_blocks(src)
    assert len(boxes) == 1
    assert boxes[0]["text"] == "内联框文字"


def test_vml_box_is_read(tmp_path: Path):
    """VML 框（``w:pict`` → ``v:shape`` → ``v:textbox``）——WPS 默认写法。"""
    src = build(tmp_path, vml_box(box_para("尾流湍流：大机身后的旋涡")))
    boxes = textbox_blocks(src)
    assert len(boxes) == 1
    assert boxes[0]["text"] == "尾流湍流：大机身后的旋涡"


def test_flavor_names_the_container(tmp_path: Path):
    """报告里要能分清是哪种框：浮动 / 内联 / VML 三种名字各不相同。"""
    frag = (drawing_box(box_para("a"))
            + drawing_box(box_para("b"), wrap="inline")
            + vml_box(box_para("c")))
    boxes = textbox_blocks(build(tmp_path, frag))
    assert [b["flavor"] for b in boxes] == [
        "drawing-floating", "drawing-inline", "vml"]


# =============================================================================
# ② AlternateContent 去重：Choice / Fallback 双写只算一次
# =============================================================================
def test_alternate_content_pair_counts_once(tmp_path: Path):
    """**本包最重要的判据**：Choice 与 Fallback 各存一份，只许读一条。

    读成 2 条就是去重没做；它比读不到更糟 —— 数字看起来是"成功"的。
    """
    frag = alternate(
        drawing_box(box_para("晴空湍流：无云、无雷达回波")),
        vml_box(box_para("晴空湍流：无云、无雷达回波")),
    )
    boxes = textbox_blocks(build(tmp_path, frag))
    assert len(boxes) == 1, f"Choice/Fallback 双写被读成了 {len(boxes)} 条"
    assert boxes[0]["text"] == "晴空湍流：无云、无雷达回波"


def test_alternate_prefers_the_choice_branch(tmp_path: Path):
    """两支文本不同时取 ``mc:Choice`` —— 那是 Word 真正渲染的那一支。"""
    frag = alternate(drawing_box(box_para("Choice 分支")),
                     vml_box(box_para("Fallback 分支")))
    boxes = textbox_blocks(build(tmp_path, frag))
    assert [b["text"] for b in boxes] == ["Choice 分支"]
    assert boxes[0]["flavor"] == "drawing-floating"


def test_empty_choice_falls_back_to_the_fallback(tmp_path: Path):
    """``mc:Choice`` 是形状但没文本时，才轮到 ``mc:Fallback``。

    这是「取 Choice」与「取有内容的那一支」的分界：一分支被清空时，另一分支
    的文字仍然可读，丢掉它等于把文档读空。
    """
    frag = alternate(drawing_box(box_para(None)), vml_box(box_para("只剩 Fallback")))
    boxes = textbox_blocks(build(tmp_path, frag))
    assert [b["text"] for b in boxes] == ["只剩 Fallback"]
    assert boxes[0]["flavor"] == "vml"


def test_many_alternate_pairs_do_not_double(tmp_path: Path):
    """43 对双写 → 43 条，不是 86（去重是逐对生效，不是整体碰运气）。"""
    frag = "".join(
        alternate(drawing_box(box_para("批注%d" % i)),
                  vml_box(box_para("批注%d" % i)))
        for i in range(43))
    boxes = textbox_blocks(build(tmp_path, frag))
    assert len(boxes) == 43, f"43 对双写读成了 {len(boxes)} 条"
    assert boxes[0]["text"] == "批注0"
    assert boxes[-1]["text"] == "批注42"


def test_all_three_modes_agree_on_the_deduped_count(tmp_path: Path):
    """``--blocks`` / ``--text`` / ``--runs`` 三个模式对同一条数必须一致。

    三个模式各写一套遍历是最容易分叉的地方：一个模式去重了、另一个没有，
    调用方交叉验证时会得到两个互相矛盾的"事实"。
    """
    frag = "".join(
        alternate(drawing_box(box_para("批注%d" % i)),
                  vml_box(box_para("批注%d" % i)))
        for i in range(5))
    src = build(tmp_path, frag)

    assert len(textbox_blocks(src)) == 5
    assert len(payload(src, "--text")["textboxes"]) == 5
    run_boxes = [p for p in payload(src, "--runs")["paragraphs"]
                 if p["textbox"] is not None]
    assert len(run_boxes) == 5
    assert len({p["textbox"] for p in run_boxes}) == 5


# =============================================================================
# ③ 文本语义：多段、嵌套、与宿主段落分开
# =============================================================================
def test_multi_paragraph_box_joins_with_newlines(tmp_path: Path):
    """框内两段 → 一个条目、两行；不是两个条目，也不是拼成一长串。"""
    content = box_para("第一行") + box_para("第二行")
    boxes = textbox_blocks(build(tmp_path, drawing_box(content)))
    assert len(boxes) == 1
    assert boxes[0]["text"] == "第一行\n第二行"
    assert boxes[0]["paragraphs"] == ["第一行", "第二行"]


def test_nested_box_is_not_folded_into_its_host(tmp_path: Path):
    """框里再套一个框：外层文本不许把内层文本吞进去。

    这正是 ``Paragraph.text`` 在处理宿主段落时做对、而朴素 ``iter(w:t)``
    会做错的那一步；框内再套一层时必须同样成立。
    """
    inner = drawing_box(box_para("内层"))
    outer = drawing_box(box_para("外层") + inner)
    boxes = textbox_blocks(build(tmp_path, outer))
    assert [b["text"] for b in boxes] == ["外层", "内层"]


def test_host_paragraph_text_is_untouched(tmp_path: Path):
    """宿主段落自己的 ``text`` 不含框里的字 —— 框是独立条目，不是段落的一部分。

    宿主段落的 ``text`` 是空串（它只装了一个 drawing），框里那句必须走
    ``textbox`` 块，而不是被塞进宿主段落。
    """
    src = build(tmp_path, drawing_box(box_para("框里的字")))
    blocks = payload(src, "--blocks")["blocks"]
    paragraph = [b for b in blocks if b["type"] == "paragraph"]
    assert [b["text"] for b in paragraph] == ["正文第一段。", ""]
    assert all("框里的字" not in b["text"] for b in paragraph)
    assert textbox_blocks(src)[0]["text"] == "框里的字"


def test_box_ordinals_are_document_wide_and_sequential(tmp_path: Path):
    """``index`` 从 0 连号，且 ``--blocks`` 与 ``--runs`` 指同一个框用同一个号。"""
    frag = "".join(drawing_box(box_para("批注%d" % i)) for i in range(3))
    src = build(tmp_path, frag)
    assert [b["index"] for b in textbox_blocks(src)] == [0, 1, 2]

    run_boxes = [p for p in payload(src, "--runs")["paragraphs"]
                 if p["textbox"] is not None]
    assert [p["textbox"] for p in run_boxes] == [0, 1, 2]
    assert [p["path"] for p in run_boxes] == [
        "body textbox[0]", "body textbox[1]", "body textbox[2]"]


def test_runs_inside_a_box_are_reported(tmp_path: Path):
    """``--runs`` 要给框内 run 的字符格式，不能只给一段纯文本。"""
    content = ('<w:p><w:r><w:rPr><w:b/><w:sz w:val="26"/></w:rPr>'
               '<w:t>粗体批注</w:t></w:r></w:p>')
    src = build(tmp_path, drawing_box(content))
    row = [p for p in payload(src, "--runs")["paragraphs"]
           if p["textbox"] is not None][0]
    assert row["text"] == "粗体批注"
    run = row["runs"][0]
    assert run["text"] == "粗体批注"
    assert run["bold"] is True
    assert run["size_pt"] == 13.0
    assert row["mixed_format"] is False, "单 run 不是 trap，不许报警"


# =============================================================================
# ④ 普通文档零影响：既有契约一个字都不许变
# =============================================================================
def test_plain_document_has_no_textbox_blocks(tmp_path: Path):
    """无框文档：``--blocks`` 里一条 ``textbox`` 都没有，块序列与旧版同形。"""
    src = build(tmp_path, "", paragraphs=("一段", "两段"))
    blocks = payload(src, "--blocks")["blocks"]
    assert [b["type"] for b in blocks] == ["paragraph", "paragraph"]
    assert [b["text"] for b in blocks] == ["一段", "两段"]


def test_plain_document_text_keeps_body_semantics(tmp_path: Path):
    """``--text`` 的 ``body`` 语义不变，只多一个空的 ``textboxes`` 键。"""
    src = build(tmp_path, "", paragraphs=("一段", "两段"))
    data = payload(src, "--text")
    assert sorted(data) == ["body", "footers", "headers", "tables", "textboxes"]
    assert data["body"] == ["一段", "两段"]
    assert data["textboxes"] == []


def test_plain_document_runs_keep_their_fields(tmp_path: Path):
    """``--runs`` 的段落字段只多一个恒为 ``null`` 的 ``textbox``，其余不变。"""
    src = build(tmp_path)
    row = payload(src, "--runs")["paragraphs"][0]
    assert sorted(row) == ["index", "mixed_format", "path", "runs", "style",
                           "table_path", "text", "textbox"]
    assert row["textbox"] is None
    assert row["path"] == "body"
    assert row["index"] == 0


def test_markdown_renders_box_text(tmp_path: Path):
    """``--md`` 也要带上框里的字，否则 round trip 又把它们丢了。"""
    src = build(tmp_path, drawing_box(box_para("被引用的批注")))
    out = tmp_path / "out.md"
    cp = run_read(src, "--md", out)
    assert cp.returncode == 0, cp.stderr
    summary = json.loads(cp.stdout)
    assert summary["textboxes"] == 1
    assert "被引用的批注" in out.read_text(encoding="utf-8")


def test_markdown_of_a_plain_document_is_unchanged(tmp_path: Path):
    """无框文档的 markdown 与旧版逐字符相同 —— 新增块类型不许带出空行。"""
    src = build(tmp_path, "", paragraphs=("标题一样的一段",))
    out = tmp_path / "out.md"
    assert run_read(src, "--md", out).returncode == 0
    assert out.read_text(encoding="utf-8") == "标题一样的一段\n"