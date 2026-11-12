# -*- coding: utf-8 -*-
"""``docx_validate.py`` 的**双契约**：输入错误退 2，内容问题仍退 1（P9-C）。

## 本文件存在的理由

`docx_validate` 是全仓唯一一个「`1` 有正当语义」的脚本：它报告的是**文档健康度**，
`ok:false` + `issues` 才是它的正常产出。P8 已定案：

    exit 0  文档健康（可能有 warning）
    exit 1  文档是真的，但**内容**有 error 级问题（stdout 是 JSON，带 issues）
    exit 2  你给的路径根本不是一个可读的 .docx（stdout 空，stderr 一行）

所以本单对它的改法**与其余六个脚本不同**：只把「输入根本不是包」这三档从 1 提到 2，
**内容问题那一档的 JSON 契约一个字都不许动**。本文件的用例就是这两条边的护栏：

* 前三条（缺文件 / 是目录 / 非 zip）→ exit **2** + stderr 单行 + **stdout 空**；
* 中两条（合法 zip 但缺 `[Content_Types].xml`、XML 非法）→ 仍 exit **1**，
  且缺 part 那条必须照样打出 ``issues``（那是它的输出面）；
* 最后一条（健康文档）→ exit **0** + ``ok:true``。

## 为什么**不**挂 ``@needs_omml``

本文件只用 ``zipfile`` + ``python-docx`` 造输入，不碰 ``math_omml`` /
``MML2OMML.XSL`` / Office COM。无 Office 的 CI runner 上必须照常 pass。

Run:  python -m pytest tests/test_validate_errors.py -q
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
SCRIPT = SCRIPTS / "docx_validate.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402


def run_validate(path) -> subprocess.CompletedProcess:
    """以真子进程跑 ``docx_validate.py``：退出码属 CLI 契约，不能 in-process 测。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def healthy_docx(tmp_path: Path) -> Path:
    d = Document()
    d.add_paragraph("hello")
    p = tmp_path / "healthy.docx"
    d.save(str(p))
    return p


# ===========================================================================
# 第一边：输入错误 → exit 2（stdout 必须是空的）
# ===========================================================================
def test_missing_path_is_a_usage_error(tmp_path: Path):
    """缺文件 → exit 2 + 单行 ``no such file`` + **stdout 空**。

    旧行为：``zipfile.ZipFile`` 抛 ``FileNotFoundError``（``OSError`` 子类）被
    收成一条 ``not-a-zip`` issue，于是**退 1 并打 JSON** —— 调用方读到的
    ``issues`` 里混着「你的路径写错了」这种根本不是文档问题的东西。
    """
    missing = tmp_path / "no-such.docx"
    proc = run_validate(missing)
    assert proc.returncode == 2, proc.stderr
    line = one_error_line(proc.stderr)
    assert "no such file" in line
    assert str(missing) in line
    assert proc.stdout == ""


def test_directory_is_a_usage_error(tmp_path: Path):
    """传目录 → exit 2 + 单行 ``not a file`` + stdout 空。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()
    proc = run_validate(adir)
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a file: {adir}"
    assert proc.stdout == ""


def test_non_zip_is_a_usage_error(tmp_path: Path):
    """非 zip → exit 2 + 单行 ``not a ZIP package`` + stdout 空。

    ★ 这条与 ``test_docx_skill.py::TestValidate::test_not_a_zip`` 是同一件事：
    那条老用例断言的是**旧语义**（exit 1 + ``not-a-zip`` JSON），已随本单改口。
    """
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("this is not a zip file", encoding="utf-8")
    proc = run_validate(bogus)
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a ZIP package: {bogus}"
    assert proc.stdout == ""


# ===========================================================================
# 第二边：内容问题 → **仍是 exit 1**，JSON 契约不动
# ===========================================================================
def test_zip_without_content_types_is_still_exit_1_with_issues(tmp_path: Path):
    """合法 zip、缺 ``[Content_Types].xml`` → exit **1** + JSON 里带 error 级 issue。

    这是「文档有问题」那一档：路径没错、确实是 zip，只是包不完整。它**必须**
    留在 exit 1，否则「输入错误」与「文档坏死」就再也分不开了 —— 这正是
    本单要保住的东西。
    """
    broken = tmp_path / "no-content-types.docx"
    with zipfile.ZipFile(broken, "w") as zf:
        zf.writestr("word/document.xml", "<w:document/>")

    proc = run_validate(broken)
    assert proc.returncode == 1, proc.stderr
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    codes = {i["code"] for i in rep["issues"]}
    assert "missing-part" in codes
    assert any(i["severity"] == "error" for i in rep["issues"])


def test_malformed_document_xml_is_still_exit_1(tmp_path: Path):
    """合法 zip、``word/document.xml`` 不是良构 XML → 仍是 exit 1（不是 2）。

    这里**只钉退出码**：该分支当前在 ``etree.fromstring`` 上直接抛
    ``XMLSyntaxError``（裸 traceback），形态不理想但语义正确（文档坏了 → 1）。
    改成「收成一条 ``bad-rels-xml`` 式 issue」会动它的输出面，超出本单范围，
    已作为遗留写进报告。
    """
    broken = tmp_path / "bad-xml.docx"
    with zipfile.ZipFile(broken, "w") as zf:
        zf.writestr("[Content_Types].xml", "<Types/>")
        zf.writestr("word/document.xml", "<w:document><unclosed>")

    proc = run_validate(broken)
    assert proc.returncode == 1, (proc.returncode, proc.stderr, proc.stdout)


# ===========================================================================
# 反向保险：健康文档照常 exit 0
# ===========================================================================
def test_healthy_document_still_passes(tmp_path: Path):
    """合法文档 → exit 0 + ``ok:true`` + stderr 空（加校验没把成功路径改坏）。"""
    src = healthy_docx(tmp_path)
    proc = run_validate(src)
    assert proc.returncode == 0, proc.stderr
    assert proc.stderr == ""
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert all(i["severity"] != "error" for i in rep["issues"])