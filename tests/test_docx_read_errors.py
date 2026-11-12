# -*- coding: utf-8 -*-
"""``docx_read.py`` 的**输入错误**契约（P8 / D2b）。

## 为什么单独成文件

`docx_read.py` 的成功路径已被 `test_docx_skill.py` / `test_outline_level.py`
覆盖，但「输入根本不是一个可读的 .docx」这条路径**从来没有被断言过**。实测旧行为：

    docx_read.py no-such.docx --text
    → exit 1，stderr 是十四行 traceback，末行
      ``docx.opc.exceptions.PackageNotFoundError: Package not found at '...'``

调用方拿到的是 python-docx 内部异常，而不是「路径写错了」。P8 定案的全仓退出码
口径是：**`2` = 用法/输入错误 → stderr 单行 `error: ...` + 零写盘**，`1` 留给
「输入合法但处理不了」的语义（例如 `docx_validate` 发现文档本身有问题）。
本文件把这条口径钉在 `docx_read` 上。

## 一条额外的硬约束：错误路径不许污染 stdout

`docx_read` 成功时 stdout 是 **JSON**（`--text` / `--structure` / `--styles` /
`--blocks` / `--revisions` / `--images` / `--md` 全是），所以错误**必须**走
stderr。若把错误也打成 JSON 到 stdout，所有 `json.loads(stdout)` 的调用方都会
把一次失败读成一次成功——本文件对每个错误用例都断 `stdout == ""`。

## 为什么**不**挂 `@needs_omml`

本文件只用 `python-docx` 造一个正常的 .docx、再用 `pathlib` 造目录与非 zip 文件；
全程不碰 `math_omml` / `latex2mathml` / `MML2OMML.XSL` / Office COM。因此在无
Office 的 CI runner 上这些用例**必须照常 pass**，写成 skip 就是写错了。

Run:  python -m pytest tests/test_docx_read_errors.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_read.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

# 每个「模式」都要独立走一遍输入校验：`--images` / `--revisions` 在
# `Document(args.path)` **之前**就分叉（前者直接开 zipfile、后者直接读 zip），
# 所以它们必须各自挡住坏输入，不能指望后面那句 `Document(...)` 兜底。
ALL_MODES = [
    ("--text",),
    ("--structure",),
    ("--styles",),
    ("--blocks",),
    ("--revisions",),
]


def run_read(*args) -> subprocess.CompletedProcess:
    """以真子进程跑 ``docx_read.py``：退出码与 stdout/stderr 都属 CLI 契约。"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    """stderr 必须**恰好一行**（`error: ...`），没有 traceback 前导。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def good_docx(tmp_path: Path) -> Path:
    """一份最小的合法 .docx（1 段正文），用于「合法路径没被改坏」的对照。"""
    d = Document()
    d.add_paragraph("hello")
    p = tmp_path / "good.docx"
    d.save(str(p))
    return p


# ===========================================================================
# 路径不存在
# ===========================================================================
@pytest.mark.parametrize("mode", ALL_MODES, ids=[m[0] for m in ALL_MODES])
def test_missing_path_is_a_usage_error(tmp_path: Path, mode):
    """路径不存在 → exit 2 + stderr 单行 `error: ...no such file` + stdout 空。

    **每个输出模式都要单独验**：`--images` 与 `--revisions` 走的是
    `extract_images` / `detect_revisions`，它们直接 `zipfile.ZipFile(path)`，
    抛的是 `FileNotFoundError`；而 `--text` 等走 `Document(path)`，抛的是
    `PackageNotFoundError`。两条异常路径的兜底必须都成立。
    """
    missing = tmp_path / "no-such.docx"
    assert not missing.exists()

    cp = run_read(missing, *mode)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line.startswith("error: ")
    assert "no such file" in line
    assert str(missing) in line
    assert cp.stdout == ""


def test_missing_path_is_reported_before_anything_is_created(tmp_path: Path):
    """`--images DIR` 的 DIR 在输入非法时**不许**被创建（零写盘）。"""
    missing = tmp_path / "no-such.docx"
    outdir = tmp_path / "img_out"
    assert not outdir.exists()

    cp = run_read(missing, "--images", outdir)
    assert cp.returncode == 2, cp.stderr
    assert cp.stdout == ""
    assert not outdir.exists()


def test_missing_path_for_md_does_not_write_the_markdown(tmp_path: Path):
    """`--md FILE` 在输入非法时**不许**写出 FILE（零写盘）。"""
    missing = tmp_path / "no-such.docx"
    md = tmp_path / "out.md"

    cp = run_read(missing, "--md", md)
    assert cp.returncode == 2, cp.stderr
    assert cp.stdout == ""
    assert not md.exists()


# ===========================================================================
# 路径存在，但不是文件 / 不是 zip
# ===========================================================================
@pytest.mark.parametrize("mode", ALL_MODES, ids=[m[0] for m in ALL_MODES])
def test_directory_path_is_a_usage_error(tmp_path: Path, mode):
    """路径是目录 → exit 2 + stderr 单行 `error: not a file: ...` + stdout 空。

    目录**存在**，所以「存在性检查」放行；必须再有一道「是文件」的检查，
    否则 `Document(dir)` 会抛 `IsADirectoryError`（同属裸 traceback 那一类）。
    """
    adir = tmp_path / "a-directory"
    adir.mkdir()

    cp = run_read(adir, *mode)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line == f"error: not a file: {adir}"
    assert cp.stdout == ""


@pytest.mark.parametrize("mode", ALL_MODES, ids=[m[0] for m in ALL_MODES])
def test_non_zip_file_is_a_usage_error(tmp_path: Path, mode):
    """存在的普通文件但不是 zip → exit 2 + stderr 单行 `error: not a ZIP package`。

    这类输入最容易从「用户把 .doc / .rtf / .txt 改名成 .docx」产生，旧行为是
    `PackageNotFoundError` / `zipfile.BadZipFile` 裸 traceback。
    """
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("i am not a zip package", encoding="utf-8")

    cp = run_read(bogus, *mode)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line == f"error: not a ZIP package: {bogus}"
    assert cp.stdout == ""


def test_truncated_zip_is_a_usage_error(tmp_path: Path):
    """合法 zip 头但内容截断 → 同样是 `not a ZIP package`，不是裸 `BadZipFile`。

    只截前 200 字节：`zipfile.is_zipfile` 只看**尾部**的中央目录，所以这里刻意
    留着 `PK\\x03\\x04` 开头而砍掉尾部——确保走的是「判定不是 zip」而不是
    「判定是 zip 但读崩」。两种都归 exit 2，但只有后者需要 `BadZipFile` 兜底。
    """
    src = good_docx(tmp_path)
    truncated = tmp_path / "truncated.docx"
    truncated.write_bytes(src.read_bytes()[:200])

    cp = run_read(truncated, "--text")
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == f"error: not a ZIP package: {truncated}"
    assert cp.stdout == ""


def test_zip_without_word_document_part_is_a_usage_error(tmp_path: Path):
    """是 zip、但不是 Word 包（缺 `word/document.xml`）→ exit 2，信息点名缺的部件。

    这条路走的是 zip 判定**通过之后** `Document(...)` 才报错，因此它验的是
    `PackageNotFoundError` / `BadZipFile` 兜底那一段（防御性接住，防 TOCTOU：
    文件在 `is_zipfile` 与 `Document()` 之间被换掉）。
    """
    import zipfile

    fake = tmp_path / "empty-package.docx"
    with zipfile.ZipFile(fake, "w") as zf:
        zf.writestr("hello.txt", "not a word package")

    cp = run_read(fake, "--text")
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line.startswith("error: ")
    assert "not a Word package" in line
    assert str(fake) in line
    assert cp.stdout == ""


# ===========================================================================
# 反向保险：合法路径逐字段不变
# ===========================================================================
def test_valid_document_text_output_is_unchanged(tmp_path: Path):
    """合法文档 `--text` 仍是 exit 0 + 纯 JSON + 四个既有键，一个不多一个不少。

    加输入校验最容易的副作用是「顺手在校验里 print 点什么」或把 JSON 结构
    改掉；这条把 stdout 当 JSON 直接解析，钉住键集合与取值。
    """
    src = good_docx(tmp_path)
    cp = run_read(src, "--text")
    assert cp.returncode == 0, cp.stderr
    assert cp.stderr == ""

    out = json.loads(cp.stdout)
    assert set(out) == {"body", "tables", "headers", "footers"}
    assert out["body"] == ["hello"]


def test_valid_document_structure_and_styles_are_unchanged(tmp_path: Path):
    """`--structure` / `--styles` / `--blocks` 的键集合原样。"""
    src = good_docx(tmp_path)

    structure = json.loads(run_read(src, "--structure").stdout)
    assert set(structure) == {"outline", "paragraph_count", "table_count",
                              "tables", "section_count"}
    assert structure["paragraph_count"] == 1

    styles = json.loads(run_read(src, "--styles").stdout)
    assert set(styles) == {"styles"}
    assert "Normal" in styles["styles"]

    blocks = json.loads(run_read(src, "--blocks").stdout)
    assert set(blocks) == {"blocks"}
    assert blocks["blocks"][0]["type"] == "paragraph"


def test_valid_document_revisions_and_images_are_unchanged(tmp_path: Path):
    """`--revisions` 的键集合与 `--images` 的空列表原样（这两条走另一条分支）。"""
    src = good_docx(tmp_path)

    rev = json.loads(run_read(src, "--revisions").stdout)
    assert rev["has_tracked_changes"] is False
    assert rev["comments"] is False

    outdir = tmp_path / "imgs"
    images = json.loads(run_read(src, "--images", outdir).stdout)
    assert images == {"images": []}
    assert outdir.is_dir()
