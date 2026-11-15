# -*- coding: utf-8 -*-
"""``docx_toc`` 的**参数拒绝面**：``--levels`` / ``--indent-chars`` / ``--pages``。

## 本文件存在的理由

三个参数各有一种「静默成功」或「裸 traceback」的失败方式，全部由主人实测确认：

* ``--levels 0`` —— 解出空集合，选中 0 个标题，于是写出一个 0 条目的空 TOC，
  **exit 0、stdout 一份看着像成功的 JSON**。用户拿到的是「命令成功了，
  目录里什么都没有」，而真相是这个参数根本没有意义（Word 只有 1–9 级）。
* ``--indent-chars a,b,c`` —— ``_levels_arg`` 里 ``int('a')`` 抛 ``ValueError``，
  没人接，于是十九行栈帧 + exit 1。
* ``--pages`` 的数量与条目数不一致 —— 这条**已经**是 exit 2（P9 已修），本文件
  把它一起钉住，因为它在同一个 ``main()`` 里，回归时最容易被顺手改坏。

三个参数共用 ``_levels_arg``，所以修法必须区分**下界**：``--levels`` 的 0 非法，
``--indent-chars`` 的 0 合法（``0,2,4`` 是默认值）。本文件的正反两面都断言：
拒绝成立的输入，同时**不拒绝**合法输入 —— 只测「拦住了」会漏掉「拦过头了」。

## 为什么每条都查「零写盘」

``docx_toc`` 是少数会**覆盖输入文件**（``-o`` 省略时）的脚本。一个用法错误如果
在 ``doc.save()`` 之后才被发现，用户的原文就没了。所以每条拒绝路径都用
SHA256 钉住「文件字节没变」+「``-o`` 目标没被创建」，而不只是退出码。

Run:  python -m pytest tests/test_toc_args.py -q
"""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

import docx_toc as toc  # noqa: E402


# ---- helpers -----------------------------------------------------------------

def sha256(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def headings_doc(tmp_path: Path, name: str = "src.docx") -> Path:
    """Three headings — one per outline level — so the counts are unambiguous."""
    doc = Document()
    doc.add_heading("Chapter One", level=1)
    doc.add_paragraph("body text")
    doc.add_heading("Section 1.1", level=2)
    doc.add_paragraph("more body")
    doc.add_heading("Sub 1.1.1", level=3)
    doc.add_paragraph("deep body")
    p = tmp_path / name
    doc.save(str(p))
    return p


def err_lines(capsys) -> list:
    return [ln for ln in capsys.readouterr().err.strip().splitlines() if ln.strip()]


# ===========================================================================
# 判据 6：--levels 0 → exit 2 + 单行 error: + 零写盘
# ===========================================================================
def test_levels_zero_is_a_usage_error(tmp_path, capsys):
    """``--levels 0`` → exit 2、stderr 恰好一行 ``error:``、**输入字节不变**。

    旧行为：exit 0 + ``{"entries": 0, "levels": []}`` + 写出一个空 TOC。
    「成功但什么都没做」比报错更难发现 —— 交付物看着是好的。
    """
    src = headings_doc(tmp_path)
    out = tmp_path / "out.docx"
    before = sha256(src)

    code = toc.main([str(src), "--levels", "0", "-o", str(out)])

    assert code == 2, "an outline level of 0 selects nothing; that is a usage error"
    lines = err_lines(capsys)
    assert len(lines) == 1, f"expected one line, got {lines}"
    assert lines[0].startswith("error: ")
    assert "minimum of 1" in lines[0]
    assert not out.exists(), "a rejected run must not create the output file"
    assert sha256(src) == before, "a rejected run must not touch the input"


def test_levels_zero_does_not_overwrite_input_when_output_omitted(tmp_path, capsys):
    """省略 ``-o`` 时输入文件就是输出文件 —— 这条最危险，必须零写盘。"""
    src = headings_doc(tmp_path, "inplace.docx")
    before = sha256(src)

    code = toc.main([str(src), "--levels", "0"])

    assert code == 2
    assert sha256(src) == before, "in-place mode must leave the document byte-identical"
    assert err_lines(capsys)[0].startswith("error: ")


@pytest.mark.parametrize("bad", ["0", "0,1", "1-0", "-1"])
def test_levels_out_of_range_are_refused(tmp_path, capsys, bad):
    """下界附近的几种写法全部拒绝：``0`` / ``0,1`` / ``1-0`` / ``-1``。

    ``1-0`` 是 ``range(1, 1)`` → 空集合，与 ``0`` 同一种「静默空目录」；
    ``-1`` 走的是另一条语法分支（连字符被当成区间分隔符）。
    """
    src = headings_doc(tmp_path)
    out = tmp_path / "never.docx"
    code = toc.main([str(src), "--levels", bad, "-o", str(out)])
    assert code == 2, f"--levels {bad} must be refused"
    assert not out.exists()
    assert len(err_lines(capsys)) == 1


def test_levels_above_nine_are_refused(tmp_path, capsys):
    """Word 只有 9 级纲目，``--levels 10`` 不可能选中任何标题。"""
    src = headings_doc(tmp_path)
    out = tmp_path / "nope.docx"
    assert toc.main([str(src), "--levels", "10", "-o", str(out)]) == 2
    assert "out of range" in err_lines(capsys)[0]
    assert not out.exists()


# ===========================================================================
# 判据 7：--indent-chars 非数字 → exit 2 + 单行 error:，无 traceback
# ===========================================================================
def test_indent_chars_non_numeric_is_a_usage_error(tmp_path, capsys):
    """``--indent-chars a,b,c`` → exit 2 + 单行 ``error:``，**不许 traceback**。

    旧行为：``int('a')`` 的 ``ValueError`` 一路穿出 ``main()``，
    ``__main__`` 的 ``raise SystemExit(main())`` 接不住，于是十九行栈帧 + exit 1。
    栈帧里既没有「这个参数写错了」也没有「该怎么写」，而 stderr 首行是
    ``Traceback (most recent call last)``。
    """
    src = headings_doc(tmp_path)
    out = tmp_path / "indent_bad.docx"
    before = sha256(src)

    code = toc.main([str(src), "--indent-chars", "a,b,c", "-o", str(out)])

    assert code == 2
    lines = err_lines(capsys)
    assert len(lines) == 1, f"expected one line, got {lines}"
    assert lines[0].startswith("error: ")
    assert "'a'" in lines[0], "the message must quote the offending token"
    assert "not a whole number" in lines[0]
    assert not out.exists()
    assert sha256(src) == before


def test_indent_chars_rejection_prints_no_traceback(tmp_path, capsys):
    """stderr 里不能出现 ``Traceback`` —— 这是「不再裸 traceback」的直接断言。"""
    src = headings_doc(tmp_path)
    assert toc.main([str(src), "--indent-chars", "x"]) == 2
    assert "Traceback" not in capsys.readouterr().err


# ===========================================================================
# 判据 8：合法输入必须**继续**工作（反证：别把好输入拦掉）
# ===========================================================================
def test_indent_chars_zero_two_four_still_works(tmp_path, capsys):
    """``--indent-chars 0,2,4``（默认值本身）→ exit 0 + 写出真 TOC。

    ``0`` 作为缩进值是合法的 —— 它与 ``--levels`` 共用解析函数，但下界不同。
    只测「拦住 a,b,c」而不测这条，就会漏掉「把合法输入一起拦掉」。
    """
    src = headings_doc(tmp_path)
    out = tmp_path / "indent_ok.docx"

    code = toc.main([str(src), "--indent-chars", "0,2,4", "-o", str(out)])

    assert code == 0, capsys.readouterr().err
    assert out.exists()
    report = json.loads(capsys.readouterr().out)
    assert report["entries"] == 3
    assert report["levels"] == [1, 2, 3]


def test_indent_chars_accepts_leading_zero_and_range(tmp_path):
    """``0-2`` 这种区间写法也得继续能用（``_levels_arg`` 的另一条分支）。"""
    src = headings_doc(tmp_path)
    out = tmp_path / "range.docx"
    assert toc.main([str(src), "--indent-chars", "0-2", "-o", str(out)]) == 0
    assert out.exists()


def test_levels_range_notation_still_works(tmp_path, capsys):
    """``--levels 1-3`` 是文档里写明的写法，不能被新下界检查误伤。"""
    src = headings_doc(tmp_path)
    code = toc.main([str(src), "--levels", "1-3", "--dry-run"])
    assert code == 0, capsys.readouterr().err
    assert json.loads(capsys.readouterr().out)["entries"] == 3


# ===========================================================================
# 判据 6b：--pages 数量不匹配 —— 核实「它是不是真的报」（已核实：是）
# ===========================================================================
def test_pages_count_mismatch_is_already_a_usage_error(tmp_path, capsys):
    """``--pages`` 数量不匹配 → exit 2（P9 已修，这里钉住别回归）。

    ``--pages`` 按位置对号：第 n 个值落在第 n 条条目上。数量不一致时多出来的
    被丢弃、缺的变成空白槽，两者都静默 —— 交付文件里一个空页码与「从没分页过」
    完全无法区分。所以这一条**已经**是 exit 2，本用例只是核实并钉住。
    """
    src = headings_doc(tmp_path)
    out = tmp_path / "pages_bad.docx"

    code = toc.main([str(src), "--pages", "1,2,3,4,5", "-o", str(out)])

    assert code == 2
    line = err_lines(capsys)[0]
    assert line.startswith("error: ")
    assert "5 value" in line and "3 entries" in line
    assert not out.exists()


def test_pages_matching_count_still_works(tmp_path, capsys):
    """数量对上 → exit 0（别把 ``--pages`` 一起拦死）。"""
    src = headings_doc(tmp_path)
    out = tmp_path / "pages_ok.docx"
    assert toc.main([str(src), "--pages", "1,2,3", "-o", str(out)]) == 0
    assert json.loads(capsys.readouterr().out)["entries"] == 3


# ===========================================================================
# 单元面：_levels_arg 的下界参数化（--levels 与 --indent-chars 共用）
# ===========================================================================
def test_levels_arg_shared_parser_keeps_the_two_floors_apart():
    """同一个解析器、两个下界：``0`` 对 indent 合法、对 levels 非法。

    这条直接对函数断言，因为 CLI 层的两条用例各自只看自己那一面，
    谁也证不了「它们真的共用同一个解析器、只是下界不同」。
    """
    assert toc._levels_arg("0,2,4", what="--indent-chars", minimum=0) == (0, 2, 4)
    with pytest.raises(ValueError):
        toc._levels_arg("0,2,4", what="--levels", minimum=1)
    assert toc._levels_arg("1,2,3", what="--levels", minimum=1) == (1, 2, 3)


def test_levels_arg_error_names_the_parameter_and_the_token():
    """拒绝消息必须同时给出「哪个参数」和「哪个值」 —— 否则用户还得猜。"""
    with pytest.raises(ValueError) as exc:
        toc._levels_arg("a,b", what="--indent-chars", minimum=0)
    msg = str(exc.value)
    assert "--indent-chars" in msg and "'a'" in msg
