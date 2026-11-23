# -*- coding: utf-8 -*-
"""``docx_validate`` 的页脚接线（P17-A A3）与「不崩」保证。

## 本文件存在的理由

主人原来的痛点原话是「``docx_validate`` 完全没报」：一份页脚被塞进 4.56 pt 浮动
文本框的文档，它返 ``{"ok": true, "issues": []}``，而页码进位到两位就被裁。

A3 把 `docx_footer` 的两条检查接进来，所以这里要钉四件事：

1. **真问题要报**：4.56 pt 浮动框 → ``footer-textbox-overflow``（severity
   ``error``）+ ``ok:false`` + exit 1；
2. **合法文档不许响**：宽框、普通页脚、真实语料 → 零命中；
3. **不崩**：页脚部件缺失 / 不是良构 XML 时给 ``warning``，绝不把 JSON 报告换
   成 lxml 栈帧 —— 这是本脚本的定位（「读得懂就报，读不懂也不能崩」）；
4. **severity 是有理由的**：溢出=error（内容永久丢失）、缺号=warning（合法中间
   态）。理由写在 ``docx_validate.py`` 的模块 docstring 里，测试钉的是行为。

## 为什么**不**挂 ``@needs_omml`` / ``needs_word``

只用 ``zipfile`` + ``lxml`` + ``python-docx`` 造输入与读取，不碰 Office COM /
LibreOffice / ``MML2OMML.XSL``。无 Office 的 runner 必须照常 pass。

Run:  python -m pytest tests/test_docx_validate.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_validate.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_docx_footer import (  # noqa: E402
    NSDECL, ZH_PAGE_LINE, _splice_footer, build_floating_footer,
    build_plain_footer,
)

SMOKE = SKILL.parent / "_smoke"


# ---- helpers -----------------------------------------------------------------

def run_validate(path) -> subprocess.CompletedProcess:
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)


def codes(rep: dict) -> list:
    return [i["code"] for i in rep["issues"]]


def footer_codes(rep: dict) -> list:
    return [c for c in codes(rep) if c.startswith("footer-")]


def corrupt_part(src: Path, dst: Path, part: str, data: bytes) -> Path:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for name in zin.namelist():
            body = zin.read(name)
            if name == part:
                body = data
            zout.writestr(name, body)
    return dst


# =============================================================================
# 1. 真问题必须报
# =============================================================================
@pytest.mark.parametrize("kind", ["drawing", "vml"])
def test_narrow_floating_textbox_is_an_error(tmp_path: Path, kind: str):
    """★ ``4.56 pt`` 浮动框 → ``footer-textbox-overflow``，severity **error**。

    error 的理由：这不是「可能难看」，是内容**永久丢失** —— 浮动框不参与重排，
    超出框宽的字画到框外被页边裁掉，读者拿到 ``第 1`` 而不是 ``第 12 页``，而文
    档里没有任何一处会报错。
    """
    src = build_floating_footer(tmp_path, kind=kind, width_pt=4.56,
                                name="%s-narrow.docx" % kind)
    proc = run_validate(src)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    hits = [i for i in rep["issues"] if i["code"] == "footer-textbox-overflow"]
    assert len(hits) == 1
    assert hits[0]["severity"] == "error"
    assert "4.56" in hits[0]["detail"]


def test_wide_floating_textbox_stays_clean(tmp_path: Path):
    """反向对照：200 pt 宽的浮动框 → 零命中、exit 0。

    没有这一条，「溢出」可能只是「凡是浮动框都报」的同义反复。
    """
    src = build_floating_footer(tmp_path, kind="drawing", width_pt=200,
                                name="wide.docx")
    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert footer_codes(json.loads(proc.stdout)) == []


def test_plain_footer_with_page_field_is_clean(tmp_path: Path):
    """最普通的页脚（无文本框 + complex PAGE 域）→ 一条页脚 issue 都不许有。"""
    src = build_plain_footer(tmp_path)
    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert footer_codes(json.loads(proc.stdout)) == []


# =============================================================================
# 2. footer-no-page-field：warning，不是 error
# =============================================================================
def test_numbering_without_page_field_is_a_warning(tmp_path: Path):
    """``pgNumType`` 在、页脚有文字但无 PAGE 域 → warning，``ok`` 仍为 true。

    severity 定 warning 的两条理由（都实测过）：

    * 这是**合法中间态** —— 编号先设、页脚留给后续工序填；
    * 真实语料里「``pgNumType`` 在、页脚没 PAGE」有 4 份，全部是**整篇没有页
      脚**（``footer_kind: none``），那是「没做页脚」而不是「页脚里漏了号」。把这
      一档判 error 会让 4 份正常文档变红。
    """
    from docx import Document
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    doc = Document()
    doc.add_paragraph("正文。")
    el = etree.SubElement(doc.sections[0]._sectPr, "{%s}pgNumType" % W)
    el.set("{%s}start" % W, "1")
    path = tmp_path / "nofield.docx"
    doc.save(str(path))
    _splice_footer(path, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                         '<w:ftr %s><w:p><w:r><w:t>机密</w:t></w:r></w:p></w:ftr>'
                   % NSDECL)

    proc = run_validate(path)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    rep = json.loads(proc.stdout)
    hits = [i for i in rep["issues"] if i["code"] == "footer-no-page-field"]
    assert len(hits) == 1
    assert hits[0]["severity"] == "warning"
    assert rep["ok"] is True


def test_numbering_without_any_footer_is_silent(tmp_path: Path):
    """设了 ``pgNumType`` 但整篇没有页脚 → **不报**。

    这是真实语料里的那 4 份。区分「没做页脚」与「页脚里漏了号」正是这条
    ``footer-no-page-field`` 能被判成 warning 而不是 error 的前提。
    """
    from docx import Document
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    doc = Document()
    doc.add_paragraph("正文。")
    el = etree.SubElement(doc.sections[0]._sectPr, "{%s}pgNumType" % W)
    el.set("{%s}start" % W, "1")
    path = tmp_path / "nofooter.docx"
    doc.save(str(path))

    proc = run_validate(path)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert footer_codes(json.loads(proc.stdout)) == []


# =============================================================================
# 3. 「读不懂」也不能崩
# =============================================================================
def test_broken_footer_part_warns_instead_of_raising(tmp_path: Path):
    """页脚部件不是良构 XML → ``footer-part-unreadable`` warning，报告照样出来。

    ``docx_validate`` 的定位是「读得懂就报，读不懂也不能崩」。页脚这一层把
    「读不动」收成一条 warning，而不是让 lxml 的栈帧顶掉整份 JSON。
    """
    src = build_plain_footer(tmp_path, name="ok.docx")
    broken = corrupt_part(src, tmp_path / "broken.docx", "word/footer1.xml",
                          b"<w:ftr><unclosed>")
    proc = run_validate(broken)
    assert proc.stderr == "", proc.stderr
    assert proc.returncode in (0, 1)
    rep = json.loads(proc.stdout)
    hits = [i for i in rep["issues"] if i["code"] == "footer-part-unreadable"]
    assert len(hits) == 1
    assert hits[0]["severity"] == "warning"


def test_missing_footer_part_warns_instead_of_raising(tmp_path: Path):
    """页脚引用指向的部件**不在包里** → warning/error 收成 issue，不 traceback。

    这是「rels 指向一个不存在的 footer2.xml」那种真实损坏：`dangling-rel` 已经
    会报，页脚检查这一层只负责不要因此抛异常。
    """
    src = build_plain_footer(tmp_path, name="ok2.docx")
    # drop the part but keep the reference: rels now dangle
    tmp = tmp_path / "nopart.docx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w") as zout:
        for name in zin.namelist():
            if name == "word/footer1.xml":
                continue
            zout.writestr(name, zin.read(name))

    proc = run_validate(tmp)
    assert proc.stderr == "", proc.stderr
    rep = json.loads(proc.stdout)
    assert "Traceback" not in proc.stderr
    assert rep["ok"] is False


def test_footer_check_never_changes_the_exit_code_contract(tmp_path: Path):
    """健康文档仍 exit 0 + ``ok:true`` + stderr 空：接线没把成功路径改坏。"""
    from docx import Document

    doc = Document()
    doc.add_paragraph("hello")
    path = tmp_path / "healthy.docx"
    doc.save(str(path))
    proc = run_validate(path)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert rep["issues"] == []


# =============================================================================
# 4. 真实语料零误报
# =============================================================================
@pytest.mark.parametrize("name", ["in.docx", "out.docx", "out_tnr.docx"])
def test_smoke_corpus_has_no_footer_findings(name: str):
    """``_smoke`` 三份真实文档：新增检查**零命中**（PLAN §四-2 的硬要求）。

    这是「加了检查就误报」最早会炸的地方 —— 这三份是仓内真实 fixture，任何新检
    查只要开始误报，这条立刻红。
    """
    path = SMOKE / name
    if not path.exists():
        pytest.skip("_smoke/%s not present" % name)
    proc = run_validate(path)
    rep = json.loads(proc.stdout)
    assert footer_codes(rep) == [], rep["issues"]
    assert proc.returncode == 0, (proc.returncode, proc.stderr)


def test_smoke_footer_check_agrees_with_validate():
    """两个入口对同一份真实文档必须同口径：``check`` 的 ok == ``validate`` 无重问题。"""
    path = SMOKE / "in.docx"
    if not path.exists():
        pytest.skip("_smoke/in.docx not present")
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    chk = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_footer.py"), "check", str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)
    val = run_validate(path)
    assert chk.returncode == 0, chk.stderr
    assert val.returncode == 0, val.stdout
    assert json.loads(chk.stdout)["ok"] is True
    assert footer_codes(json.loads(val.stdout)) == []