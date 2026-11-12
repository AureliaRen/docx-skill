# -*- coding: utf-8 -*-
"""``docx_numbering.py`` 的**输入错误**契约（P9-C）。

## 为什么单独成一个文件

`docx_numbering.py` 的行为面已被 `test_numbering_crossref.py` /
`test_section_para.py` 覆盖（define / apply / detach / restart 的 XML 副作用），
但「输入根本不是一个可读的 .docx」这条路径**从来没有被断言过**。实测旧行为：

    docx_numbering.py no-such.docx list
    → exit 1，stderr 是十九行 traceback，末行
      ``docx.opc.exceptions.PackageNotFoundError: Package not found at '...'``

目录与非 zip 分别抛 ``IsADirectoryError`` / ``PackageNotFoundError``，同样是裸
traceback。全仓退出码口径（P8 定案）：**`2` = 用法/输入错误 → stderr 单行
`error: ...` + stdout 空 + 零写盘**；`1` 留给「输入合法但处理不了」。

## 为什么每个子命令都单独验

`main()` 里 `Document(args.path)` 在**参数交叉校验之后、分派之前**，所以坏输入对
每个子命令都该在同一个位置被挡住。把 `list` / `define` / `apply` 各走一遍，就是
在钉「校验没被塞进某个子命令的私有分支」。`define` / `apply` 是**会写盘**的
（`doc.save(out)`），它们是「零写盘」的关键用例。

## 为什么**不**挂 ``@needs_omml``

本文件只用 `python-docx` 造一份正常 .docx，再用 `pathlib` 造目录与非 zip 文件；
全程不碰 `math_omml` / `latex2mathml` / `MML2OMML.XSL` / Office COM。无 Office 的
CI runner 上这些用例**必须照常 pass**，写成 skip 就是写错了。

Run:  python -m pytest tests/test_numbering_errors.py -q
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
SCRIPT = SCRIPTS / "docx_numbering.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

# (子命令, 追加参数) —— 三条覆盖「只读 / 追加定义 / 批量改写」三种落点。
MODES = [
    ("list", ()),
    ("define", ("--text", "%1.")),
    ("apply", ("--index", "0", "--num-id", "1")),
]


def run_numbering(*args) -> subprocess.CompletedProcess:
    """以真子进程跑 ``docx_numbering.py``：退出码与 stdout/stderr 都属 CLI 契约。"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    """stderr 必须**恰好一行**（``error: ...``），没有 traceback 前导。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def good_docx(tmp_path: Path) -> Path:
    """一份最小的合法 .docx（零个编号定义），用于成功对照与「零写盘」检查。"""
    d = Document()
    d.add_paragraph("hello")
    p = tmp_path / "good.docx"
    d.save(str(p))
    return p


# ===========================================================================
# 四档坏输入
# ===========================================================================
@pytest.mark.parametrize("cmd,extra", MODES, ids=[m[0] for m in MODES])
def test_missing_path_is_a_usage_error(tmp_path: Path, cmd, extra):
    """路径不存在 → exit 2 + stderr 单行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such.docx"
    assert not missing.exists()

    cp = run_numbering(cmd, missing, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line.startswith("error: ")
    assert "no such file" in line
    assert str(missing) in line
    assert cp.stdout == ""


@pytest.mark.parametrize("cmd,extra", MODES, ids=[m[0] for m in MODES])
def test_directory_path_is_a_usage_error(tmp_path: Path, cmd, extra):
    """路径是目录 → exit 2 + 单行 ``not a file``。

    目录**存在**，所以存在性检查会放行；没有第二道「是文件」的检查就会落到
    ``Document(dir)`` 抛 ``IsADirectoryError``。
    """
    adir = tmp_path / "a-directory"
    adir.mkdir()

    cp = run_numbering(cmd, adir, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == f"error: not a file: {adir}"
    assert cp.stdout == ""


@pytest.mark.parametrize("cmd,extra", MODES, ids=[m[0] for m in MODES])
def test_non_zip_file_is_a_usage_error(tmp_path: Path, cmd, extra):
    """存在的普通文件但不是 zip → exit 2 + 单行 ``not a ZIP package``。"""
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("i am not a zip package", encoding="utf-8")

    cp = run_numbering(cmd, bogus, *extra)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == f"error: not a ZIP package: {bogus}"
    assert cp.stdout == ""


def test_non_zip_input_writes_nothing_even_with_output_flag(tmp_path: Path):
    """``-o out.docx`` + 坏输入 → **不生成** out.docx（零写盘）。

    `define` 是本脚本里唯一「先读后写」的子命令：一旦校验被放到 `doc.save()`
    之后，这里就会出现一个 0 字节或被截断的输出文件，而退出码仍是 2。
    """
    bogus = tmp_path / "renamed.docx"
    bogus.write_text("i am not a zip package", encoding="utf-8")
    out = tmp_path / "out.docx"

    cp = run_numbering("define", bogus, "-o", out)
    assert cp.returncode == 2, cp.stderr
    assert cp.stdout == ""
    assert not out.exists()


# ===========================================================================
# 反向保险：合法路径照常工作
# ===========================================================================
def test_valid_document_list_is_unchanged(tmp_path: Path):
    """合法文档 ``list`` 仍是 exit 0 + 纯 JSON + 既有三个键。

    加输入校验最容易的副作用是「顺手在校验里 print 点什么」或把 `_open_document`
    之后的对象换掉；这条把 stdout 当 JSON 解析，钉住键集合与空文档的取值。
    """
    src = good_docx(tmp_path)

    cp = run_numbering("list", src)
    assert cp.returncode == 0, cp.stderr
    assert cp.stderr == ""

    out = json.loads(cp.stdout)
    assert set(out) == {"abstract_nums", "nums", "usage"}
    # python-docx 的默认模板自带一套 numbering 定义，所以这里**不能**断言空表：
    # 空白文档的 abstract_nums / nums 本来就有内容（实测 9 个 abstractNum）。
    # 钉的是结构而不是具体条数——条数随 python-docx 版本变，结构不会。
    assert out["abstract_nums"], "默认模板应当自带 abstractNum 定义"
    assert all("abstractNumId" in a for a in out["abstract_nums"])
    assert all({"numId", "abstractNumId"} <= set(n) for n in out["nums"])
    assert out["usage"] == {}


def test_valid_document_define_still_writes_the_counter(tmp_path: Path):
    """合法文档 ``define`` 仍是 exit 0，且真的把 `<w:num>` 写进了输出文件。"""
    src = good_docx(tmp_path)
    out = tmp_path / "out.docx"

    cp = run_numbering("define", src, "-o", out, "--text", "[%1]")
    assert cp.returncode == 0, cp.stderr
    payload = json.loads(cp.stdout)
    assert payload["output"] == str(out)
    assert out.is_file()