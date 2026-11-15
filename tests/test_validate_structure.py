# -*- coding: utf-8 -*-
"""``docx_validate`` 的**结构合法性**检查 + XML 非法不再裸 traceback（P13-B）。

## 本文件存在的理由

``docx_validate`` 原来是纯「引用面」检查：包件存在 / rels 解析 / ``r:id`` 解析 /
媒体 magic / style 存在。这五项都只问「这个名字能不能解析到东西」，
**不问标记本身的形状**。于是两种 Word 直接拒绝打开的文档全绿：

* 删掉 inline shape、又换过图片 part 之后留下的空 ``<w:drawing/>``
  （内部既无 ``wp:inline`` 也无 ``wp:anchor``）—— 主人实战踩到，Word 报
  「无法读取的内容」，而校验器返回 ``{"ok": true, "issues": []}``、exit 0；
* 块级元素 ``<w:p>`` / ``<w:tbl>`` 直接嵌在 ``<w:r>`` 里。

第二种更隐蔽：python-docx 打开它毫无怨言，所以连最后的兜底检查也看不见。

同时钉住 `word/document.xml` **不是良构 XML** 时的形态：P9-C 已把退出码定在 1，
但当时它会从 ``etree.fromstring`` 直接抛出去 —— 十九条 lxml 栈帧 + 空 stdout，
而本脚本的契约是「stdout 上永远有一份 JSON 报告」。现在它是一条
``bad-document-xml`` issue。

## 为什么钉「正常文档零 issue」

本脚本的能力边界是 health-check，不是 XSD。加检查最容易犯的错是把合法文档判坏，
所以**仓内真实 fixture**（`_smoke` 目录下三份，只读）和「带真图片的文档」都断言
``issues == []``：任何新增检查只要开始误报，这几条立刻红。

Run:  python -m pytest tests/test_validate_structure.py -q
"""
from __future__ import annotations

import json
import os
import struct
import subprocess
import sys
import zipfile
import zlib
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_validate.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

SMOKE = SKILL.parent / "_smoke"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---- helpers -----------------------------------------------------------------

def run_validate(path) -> subprocess.CompletedProcess:
    """真子进程：退出码是 CLI 契约的一部分，in-process 测不到。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def make_png(path: Path, w: int = 4, h: int = 4) -> None:
    """A minimal valid PNG — enough for ``a:blip`` + magic-byte checks."""
    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                     + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def rewrite_part(src: Path, dst: Path, part: str, transform) -> None:
    """Copy *src* to *dst*, replacing ``part`` with ``transform(bytes) -> bytes``."""
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = transform(data)
            zout.writestr(item, data)


def docx_with(tmp_path: Path, name: str, body_xml: str = "") -> Path:
    """A one-paragraph docx with *body_xml* appended to ``w:body``.

    Built by string surgery on a real python-docx file rather than by
    ``OxmlElement``: the point is to produce markup python-docx's own API
    refuses to create, which is exactly what a half-finished edit leaves behind.
    """
    base = tmp_path / "_base.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(base))
    target = tmp_path / name
    rewrite_part(
        base, target, "word/document.xml",
        lambda data: data.replace(
            b"</w:body>",
            body_xml.encode("utf-8") + b"</w:body>", 1))
    base.unlink()
    return target


def picture_docx(tmp_path: Path) -> Path:
    png = tmp_path / "dot.png"
    make_png(png)
    d = Document()
    d.add_paragraph("with a picture")
    d.add_picture(str(png))
    p = tmp_path / "with_picture.docx"
    d.save(str(p))
    return p


# ===========================================================================
# 判据 1：空 <w:drawing/> → ok:false + 一条 error issue + exit 1
# ===========================================================================
def test_empty_drawing_is_an_error(tmp_path: Path):
    """空 ``<w:drawing/>``（无 ``wp:inline`` / ``wp:anchor``）→ error + exit 1。

    这是主人实际踩到的形状：换掉封面校徽的图片 part、又删了 inline shape，
    段落里留下一个空 drawing。所有 ``r:embed`` 都没了，引用面五项检查一条都不
    会问它，而 Word 直接拒绝打开整个文档。
    """
    src = docx_with(
        tmp_path, "empty-drawing.docx",
        '<w:p><w:r><w:drawing/></w:r>'
        '<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>')

    proc = run_validate(src)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == "", "a document problem must not print to stderr"
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    codes = [i["code"] for i in rep["issues"]]
    assert "empty-drawing" in codes, rep
    bad = [i for i in rep["issues"] if i["code"] == "empty-drawing"]
    assert len(bad) == 1, "one issue per offending drawing, not per scan"
    assert bad[0]["severity"] == "error"
    assert "wp:inline" in bad[0]["detail"] and "wp:anchor" in bad[0]["detail"]


def test_empty_drawing_inside_alternate_content_is_still_caught(tmp_path: Path):
    """``mc:AlternateContent`` 里也不能藏空 drawing —— 检查看 drawing 自己，不看父级。

    这条防的是「只在 ``w:drawing`` 的直接父级是 ``w:r`` 时才查」这种取巧实现。
    同时反证：``mc:Choice`` 里那个**带** ``wp:inline`` 的 drawing 是合法的，
    不能被误判 —— 合并场景里 ``mc:AlternateContent`` 到处都是。
    """
    src = docx_with(
        tmp_path, "alt-content.docx",
        '<w:p><w:r><mc:AlternateContent>'
        '<mc:Choice Requires="wps"><w:drawing><wp:inline/></w:drawing></mc:Choice>'
        '<mc:Fallback><w:drawing/></mc:Fallback>'
        '</mc:AlternateContent></w:r></w:p>')

    proc = run_validate(src)
    rep = json.loads(proc.stdout)
    assert proc.returncode == 1
    assert [i["code"] for i in rep["issues"]] == ["empty-drawing"], rep


# ===========================================================================
# 判据 1b：非法嵌套 <w:p> in <w:r> → error（可选覆盖项，已实现）
# ===========================================================================
def test_paragraph_inside_run_is_invalid_nesting(tmp_path: Path):
    src = docx_with(
        tmp_path, "bad-nesting.docx",
        '<w:p><w:r><w:t>x</w:t>'
        '<w:p><w:r><w:t>inner</w:t></w:r></w:p></w:r></w:p>')

    proc = run_validate(src)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    bad = [i for i in rep["issues"] if i["code"] == "invalid-nesting"]
    assert len(bad) == 1, rep
    assert bad[0]["severity"] == "error"
    assert "<w:p>" in bad[0]["detail"]


def test_table_inside_run_is_invalid_nesting(tmp_path: Path):
    """``<w:tbl>`` 同样不是 run content —— 块级两种元素都覆盖到了。"""
    src = docx_with(
        tmp_path, "tbl-nesting.docx",
        '<w:p><w:r><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:r></w:p>')

    proc = run_validate(src)
    rep = json.loads(proc.stdout)
    assert proc.returncode == 1
    assert [i["code"] for i in rep["issues"]] == ["invalid-nesting"]
    assert "<w:tbl>" in rep["issues"][0]["detail"]


# ===========================================================================
# 判据 2 / 3：正常文档零 issue（含真图片、含仓内真实 fixture）
# ===========================================================================
def test_real_picture_document_has_no_issues(tmp_path: Path):
    """带真图片（``wp:inline`` + ``a:blip``）的文档 → ``ok:true``、issues 空。

    这是「加检查不判坏正常文档」的第一条护栏：新检查碰到的合法 drawing
    是最常见的 ``wp:inline`` 那种，它必须一条 issue 都不产生。
    """
    src = picture_docx(tmp_path)
    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep == {"ok": True, "issues": []}


def test_anchored_drawing_is_not_empty(tmp_path: Path):
    """只有 ``wp:anchor`` 的浮动对象是合法的 —— 别把 anchor 当成空 drawing。

    ``wp:inline`` 与 ``wp:anchor`` 是二选一：要求两者都有，会把所有
    「四周环绕」排版的图误判成坏文档。
    """
    src = docx_with(
        tmp_path, "anchored.docx",
        '<w:p><w:r><w:drawing><wp:anchor/></w:drawing></w:r></w:p>')
    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stdout)
    assert json.loads(proc.stdout)["issues"] == []


@pytest.mark.parametrize("name", ["in.docx", "out.docx", "out_tnr.docx"])
def test_smoke_fixtures_stay_completely_clean(name):
    """仓内真实 fixture（只读）→ issues **空**，与改动前逐字节一致。

    P13-B 的硬约束：「加检查绝不能把正常文档判坏」。这三份是真实产出的
    docx（公式 / 字体 / 版式都在里面），改动前后 ``docx_validate`` 的输出都是
    ``{"ok": true, "issues": []}`` —— 本用例就是那条不变量的常驻护栏。
    """
    src = SMOKE / name
    if not src.exists():  # pragma: no cover - 装机副本没有 _smoke 父目录
        pytest.skip(f"{src} not present in this checkout")
    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    assert json.loads(proc.stdout) == {"ok": True, "issues": []}


# ===========================================================================
# 判据 4：document.xml 非法 XML → 单行/JSON 形态 + exit 1，无 traceback
# ===========================================================================
def test_malformed_document_xml_reports_instead_of_raising(tmp_path: Path):
    """非法 XML → exit **1** + ``bad-document-xml`` + exit≠2、stderr 空、无栈帧。

    旧行为：``etree.fromstring`` 的 ``XMLSyntaxError`` 没人接 → 十九条 lxml 栈帧
    写到 stderr、stdout 空、exit 1。退出码语义本来就对（文档坏了），但形态违约：
    本脚本承诺 stdout 上永远有一份 JSON。
    """
    broken = tmp_path / "bad-xml.docx"
    with zipfile.ZipFile(broken, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "<w:document><unclosed>")

    proc = run_validate(broken)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    assert "Traceback" not in proc.stderr, proc.stderr
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    assert [i["code"] for i in rep["issues"]] == ["bad-document-xml"]
    assert rep["issues"][0]["severity"] == "error"


def test_malformed_styles_xml_reports_instead_of_raising(tmp_path: Path):
    """``word/styles.xml`` 坏了也一样：issue，不是 traceback。"""
    base = tmp_path / "_base.docx"
    d = Document()
    d.add_paragraph("hello")
    d.save(str(base))
    broken = tmp_path / "bad-styles.docx"
    rewrite_part(base, broken, "word/styles.xml",
                 lambda _data: b"<w:styles><unclosed>")

    proc = run_validate(broken)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    assert "Traceback" not in proc.stderr
    rep = json.loads(proc.stdout)
    codes = [i["code"] for i in rep["issues"]]
    assert "bad-styles-xml" in codes, rep


# ===========================================================================
# 判据 5：输入本身不是包 → 仍是 exit 2（新分支别把它带回 exit 1）
# ===========================================================================
def test_missing_path_is_still_a_usage_error(tmp_path: Path):
    proc = run_validate(tmp_path / "nope.docx")
    assert proc.returncode == 2, proc.stderr
    assert "no such file" in one_error_line(proc.stderr)
    assert proc.stdout == ""


def test_directory_is_still_a_usage_error(tmp_path: Path):
    proc = run_validate(tmp_path)
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a file: {tmp_path}"
    assert proc.stdout == ""


def test_non_zip_is_still_a_usage_error(tmp_path: Path):
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("this is not a zip file", encoding="utf-8")
    proc = run_validate(bogus)
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a ZIP package: {bogus}"
    assert proc.stdout == ""


def test_bad_xml_is_one_not_two(tmp_path: Path):
    """「zip 里 XML 坏了」≠「输入不是包」：必须 exit 1，不能退成 2。

    这两档混起来最伤：分不清的话，调用方会把「你自己路径写错了」当成
    「文档要修」，或者反过来去修一个根本没读到的文件。
    """
    broken = tmp_path / "bad-xml.docx"
    with zipfile.ZipFile(broken, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "<w:document><unclosed>")
    proc = run_validate(broken)
    assert proc.returncode == 1
    assert proc.returncode != 2


# ===========================================================================
# 判据 6：多个 .rels 不再共用一把钥匙（既有 bug，且随 PYTHONHASHSEED 摇摆）
# ===========================================================================
def two_rels_docx(tmp_path: Path) -> Path:
    """一份**合法**文档，包内同时存在 ``document.xml.rels`` 与 ``header1.xml.rels``。

    两者同住 ``word/_rels/``。旧实现按「父目录」给 rels 建键，于是这两个——
    以及 ``footer1`` / ``comments`` / ``footnotes`` / ``endnotes`` / ``fontTable``
    ——全部塌进同一个 ``"word"`` 槽，后读的覆盖先读的。而读的顺序来自一个
    `set`，所以「谁赢」按 ``PYTHONHASHSEED`` 变。症状不是崩溃，是**同一个文件
    在不同进程里给出不同答案**：某些 seed 报两条 ``unresolved-reference``，
    另一些 seed 报 ``ok:true``。

    这张 fixture 用 python-docx 造（含一张真图片，保证 document.xml 里有
    ``r:embed``），再用 zip 手术塞一个空 ``header1.xml.rels``。两个 rels 都
    合法，所以正确答案恒为 ``issues == []``。
    """
    base = tmp_path / "_base.docx"
    png = tmp_path / "_dot.png"
    make_png(png)
    d = Document()
    d.add_paragraph("hello")
    d.add_picture(str(png))
    d.save(str(base))

    target = tmp_path / "two-rels.docx"
    empty_rels = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                  b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                  b'package/2006/relationships"></Relationships>')
    with zipfile.ZipFile(base) as zin, \
            zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("word/_rels/header1.xml.rels", empty_rels)
    base.unlink()
    png.unlink()
    return target


@pytest.mark.parametrize("seed", ["0", "1", "2", "3", "4", "5", "6", "7"])
def test_two_rels_files_do_not_collide(tmp_path: Path, seed: str):
    """8 个 seed 必须给出同一个答案 —— 而且都是 ``ok:true``。

    钉 seed 而不是「多跑几次」：旧实现在 0/3/4/5/7 上稳定报错、1/2/6 上稳定
    通过，跑一次有一半概率测不出来。seed 0 是旧实现**稳定失败**的那个，所以
    这条用例在旧代码上必红、在新代码上必绿，不留「碰巧过」的空间。
    """
    path = two_rels_docx(tmp_path)
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8",
               PYTHONUTF8="1", PYTHONHASHSEED=seed)
    proc = subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )
    assert proc.returncode == 0, (seed, proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    assert rep == {"ok": True, "issues": []}, (seed, rep)


def test_a_header_rels_still_resolves_its_own_targets(tmp_path: Path):
    """修键不能只是「别再塌在一起」——每个 rels 仍要解析到**它自己**的部件。

    ``header1.xml.rels`` 指向 ``header1.xml``（同目录，相对目标）。按文件主干
    建键时 owner 是 ``word/header1.xml``，相对目标要解到 ``word/header1.xml``
    才对；解到 ``word/document.xml`` 就会误报 dangling-rel。
    """
    base = tmp_path / "_base.docx"
    d = Document()
    d.add_paragraph("hello")
    d.save(str(base))

    header_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/'
        b'package/2006/relationships">'
        b'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
        b'officeDocument/2006/relationships/header" Target="header1.xml"/>'
        b'</Relationships>')
    header = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              b'<w:hdr xmlns:w="http://schemas.openxmlformats.org/'
              b'wordprocessingml/2006/main"><w:p/></w:hdr>')

    target = tmp_path / "with-header.docx"
    with zipfile.ZipFile(base) as zin, \
            zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, zin.read(item.filename))
        zout.writestr("word/header1.xml", header)
        zout.writestr("word/_rels/header1.xml.rels", header_rels)
    base.unlink()

    proc = run_validate(target)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert json.loads(proc.stdout) == {"ok": True, "issues": []}
