# -*- coding: utf-8 -*-
"""``docx_edit.py`` 八个子命令的深覆盖（P5 / A-2）。

## 为什么单独成文件

`test_scripts.py` 只是**顺手**用到了 `docx_edit` 背后的 `replace_in_paragraph`，
八个子命令的 **CLI 契约**（argparse 参数、JSON stdout、退出码、写盘时机）与
**位置语义**（index 是 body 段落序号、跨 run 替换、field 三件套）从来没有被断言过。
「文件里能搜到那个词就算过」这类断言对「改错了段落」「字段少一个 separate」
「merge 时吃掉一个字」完全免疫——所以本文件把每条断言钉在**具体元素 / 具体下标 /
逐字文本**上。

## 覆盖表（子命令 → 承重断言）

| 子命令 | 承重断言 |
|---|---|
| `replace` | 跨 run 命中（词被拆进两个 run 也要替换掉，且继承**匹配起点那个 run** 的格式）；命中数按**原文出现次数**计；`--body-only` 真的跳过页眉；默认范围连**表格单元格**一起扫 |
| `set-cell` | 只有目标 (table,row,col) 变，其余单元格逐格不变；多段落单元格被**压成一段** |
| `insert` | 新段落在**指定 index 之前**（= 该 index 的位置），前后段落顺序不变 |
| `delete` | 只删指定 index，段落总数 -1，`w:sectPr` 仍在 body 末尾 |
| `style` | `w:pStyle` 真写进目标段（`w:val` = styleId），相邻段**没有** pPr |
| `normalize` | 合并后 `w:r` 数量下降且**文本逐字不变**；`w:rPr` 不同不合并；带 `w:br` 的 run 是隔离带；空白保留（`xml:space="preserve"`） |
| `toc` | 段落里有 `TOC` 域，且 `fldChar begin → instrText → separate → 占位文本 → end` 顺序齐全 |
| `page-numbers` | 页脚出现 `PAGE` 与 `NUMPAGES`，两个域各一套 begin/separate/end |

## 边界用例（非零退出 + 源文件 SHA256 不变）

`delete` / `style` / `set-cell` 的越界与未知样式名都走「先算结果、后 save」的路径，
所以异常必须在 `doc.save()` 之前抛出——`SHA256 不变` 就是在断这件事。

## 为什么**不**挂 `@needs_omml`

读源码确认过：`docx_edit.py` 只 import `_bootstrap` / `argparse` / `json` / `sys` /
`python-docx` / `docx_common`，`docx_common` 也不碰 `math_omml` / `latex2mathml` /
`MML2OMML.XSL`。八个子命令全程没有 LaTeX→OMML 转换路径，因此在无 Office 的 CI 上
这些用例**必须照常运行**，挂 skip 标记反而会白丢覆盖。

Run:  python -m pytest tests/test_docx_edit_subcommands.py -q
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_edit.py"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.enum.text import WD_BREAK  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------- helpers
def run_edit(*args) -> subprocess.CompletedProcess:
    """以子进程跑 ``docx_edit.py``：退出码与 stdout 都属 CLI 契约，必须真进程验。"""
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def payload(cp: subprocess.CompletedProcess) -> dict:
    """断言退出码为 0 并把 stdout 当 **JSON** 解析。

    `SKILL.md` 把 `docx_edit` 归到 "JSON on stdout" 那一类，所以这里不做字符串
    包含判断，直接 `json.loads`——顺带把「有没有多打一行日志污染 stdout」也验了。
    """
    assert cp.returncode == 0, (
        f"rc={cp.returncode}\nstdout={cp.stdout}\nstderr={cp.stderr}")
    return json.loads(cp.stdout.strip())


def sha256(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def freeze_docx(path: Path) -> Path:
    """把 zip 条目的时间戳钉成 1980-01-01，让「文件 SHA256」只反映**内容**。

    为什么必须这么做（实测坑，见 `P5-A2-REPORT.md` §5）：

    python-docx 保存时用 `zipfile` 默认的 DOS 时间戳，**粒度是 2 秒**。于是
    「造 fixture 的 save」和「被测脚本里那次 save」如果落在同一个 2 秒刻度内，
    两次输出**逐字节完全相同**——哪怕脚本真的重写了整个包。结果是
    「SHA256 不变 → 没被改动」这条断言会**随机**放行一个「save 提到操作之前」的
    真实回归（本文件第一次反向自证就撞上了：同一变异跑两次，一次红一次绿）。

    把 fixture 的时间戳钉到 1980 之后，任何一次新 save 的字节流必然不同，
    SHA256 才成为确定性的内容指纹。
    """
    src = Path(path)
    raw = src.read_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(out, "w") as zout:
            for info in infos:
                new = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
                new.compress_type = info.compress_type
                new.external_attr = info.external_attr
                zout.writestr(new, zin.read(info.filename))
    src.write_bytes(out.getvalue())
    return src


def save(tmp_path: Path, name: str, paragraphs, builder=None) -> Path:
    d = Document()
    for text in paragraphs:
        d.add_paragraph(text)
    if builder is not None:
        builder(d)
    p = tmp_path / name
    d.save(str(p))
    return freeze_docx(p)


def body_paragraph_elements(doc) -> list:
    """``<w:body>`` 的**直接** ``<w:p>`` 子元素（不含表格单元格里的段落）。"""
    return [c for c in doc.element.body if c.tag == w("p")]


def field_sequence(p_el) -> list:
    """按文档顺序抽出段落里的域结构：``('fldChar', type)`` / ``('instrText', text)`` / ``('t', text)``。"""
    seq = []
    for el in p_el.iter():
        tag = el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""
        if tag == "fldChar":
            seq.append(("fldChar", el.get(w("fldCharType"))))
        elif tag == "instrText":
            seq.append(("instrText", el.text))
        elif tag == "t" and el.text is not None:
            seq.append(("t", el.text))
    return seq


def instr_texts(p_el) -> list:
    return [t.text or "" for t in p_el.iter(w("instrText"))]


def pstyle_val(para) -> str | None:
    pPr = para._p.find(w("pPr"))
    if pPr is None:
        return None
    ps = pPr.find(w("pStyle"))
    return None if ps is None else ps.get(w("val"))


# ===========================================================================
# replace
# ===========================================================================
def test_replace_hits_a_word_split_across_two_runs(tmp_path: Path):
    """跨 run 命中：`"hello"` 被拆成 `"hel"` + `"lo world"`，仍必须整体替换。

    这是本子命令**相对普通 XML 整串替换的核心价值**：Word 文档里同一个词经常被
    拆进多个 run（拼写检查、格式微调都会拆），只按单 run 找的实现会在这里静默失效
    ——返回 0 命中、文档一字不改，而且不报错。

    同时断言格式归属：替换文本落在**匹配起点那个 run** 里，因此必须仍是 bold。
    """
    d = Document()
    d.add_paragraph("alpha one")
    p = d.add_paragraph()
    r1 = p.add_run("hel")
    r1.bold = True
    p.add_run("lo world")
    d.add_paragraph("omega three")
    src = tmp_path / "xrun.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "xrun_out.docx"
    res = payload(run_edit("replace", src, "--find", "hello",
                           "--replace", "HI", "-o", out))
    assert res["replacements"] == 1
    assert res["ok"] is True
    assert res["output"] == str(out)

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == [
        "alpha one", "HI world", "omega three"]
    # 格式继承自匹配起点所在的 run（hel 是 bold）
    runs = got.paragraphs[1].runs
    assert runs[0].text == "HI"
    assert runs[0].bold is True
    assert runs[1].text == " world"


def test_replace_counts_every_cross_run_occurrence(tmp_path: Path):
    """一次调用里的**多处**跨 run 命中都要换掉，命中数按原文出现次数计。

    fixture 全文 = `"abcd abcd"`（被切成 `ab` / `cd ab` / `cd`），`abcd` 出现 2 次。
    Pass 2 的 budget 是从**原文**出现次数算的，所以这里必须正好 2；漏一处就是回归。
    """
    d = Document()
    p = d.add_paragraph()
    p.add_run("ab")
    p.add_run("cd ab")
    p.add_run("cd")
    src = tmp_path / "xrun2.docx"
    d.save(str(src))
    freeze_docx(src)
    assert Document(str(src)).paragraphs[0].text == "abcd abcd"

    out = tmp_path / "xrun2_out.docx"
    res = payload(run_edit("replace", src, "--find", "abcd",
                           "--replace", "X", "-o", out))
    assert res["replacements"] == 2
    assert Document(str(out)).paragraphs[0].text == "X X"


def test_replace_cross_run_terminates_when_replacement_contains_needle(tmp_path: Path):
    """跨 run 命中 + `new` 含 `old`：必须**终止**且不重复膨胀。

    单 run 版本的死循环回归已由 `test_scripts.py` 守着；这里是同一条逻辑在
    跨 run 分支上的对照——`"ca"`+`"t"` → `cat` 换 `catcat`，结果只能是 `catcat`，
    不能变成 `catcatcat...`。
    """
    d = Document()
    p = d.add_paragraph()
    p.add_run("ca")
    p.add_run("t")
    src = tmp_path / "catcat.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "catcat_out.docx"
    res = payload(run_edit("replace", src, "--find", "cat",
                           "--replace", "catcat", "-o", out))
    assert res["replacements"] == 1
    assert Document(str(out)).paragraphs[0].text == "catcat"


def test_replace_default_scope_covers_body_tables_and_header(tmp_path: Path):
    """默认范围 = body + 表格单元格（递归）+ 页眉；三处都要换到。"""
    d = Document()
    d.add_paragraph("tick body")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "tick cell"
    sec = d.sections[0]
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "tick header"
    src = tmp_path / "scope.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "scope_all.docx"
    res = payload(run_edit("replace", src, "--find", "tick",
                           "--replace", "TOCK", "-o", out))
    assert res["replacements"] == 3
    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["TOCK body"]
    assert got.tables[0].cell(0, 0).text == "TOCK cell"
    assert [p_.text for p_ in got.sections[0].header.paragraphs] == ["TOCK header"]


def test_replace_body_only_leaves_header_untouched(tmp_path: Path):
    """`--body-only` 只跳过页眉页脚，**不**跳过表格。"""
    d = Document()
    d.add_paragraph("tick body")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "tick cell"
    sec = d.sections[0]
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "tick header"
    src = tmp_path / "scope_b.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "scope_body.docx"
    res = payload(run_edit("replace", src, "--find", "tick",
                           "--replace", "TOCK", "--body-only", "-o", out))
    assert res["replacements"] == 2
    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["TOCK body"]
    assert got.tables[0].cell(0, 0).text == "TOCK cell"
    assert [p_.text for p_ in got.sections[0].header.paragraphs] == ["tick header"]


# ===========================================================================
# set-cell
# ===========================================================================
def _grid(doc, t_index: int = 0) -> list[list[str]]:
    t = doc.tables[t_index]
    return [[t.cell(r, c).text for c in range(len(t.columns))]
            for r in range(len(t.rows))]


def test_set_cell_changes_only_the_target_cell(tmp_path: Path):
    """改的确实是 (0,1,2)，其余 5 格逐格不变。

    「文档里出现 NEW 就算过」对**行列写反**完全免疫（把 row/col 传反了照样有 NEW），
    所以必须整张表逐格比对。
    """
    d = Document()
    d.add_paragraph("intro")
    t = d.add_table(rows=2, cols=3)
    for r in range(2):
        for c in range(3):
            t.cell(r, c).text = f"r{r}c{c}"
    src = tmp_path / "cell.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "cell_out.docx"
    res = payload(run_edit("set-cell", src, "--table", 0,
                           "--row", 1, "--col", 2, "--text", "NEW", "-o", out))
    assert (res["table"], res["row"], res["col"]) == (0, 1, 2)
    assert res["ok"] is True

    got = Document(str(out))
    assert _grid(got) == [["r0c0", "r0c1", "r0c2"],
                          ["r1c0", "r1c1", "NEW"]]
    # 正文段落不受影响
    assert [p_.text for p_ in got.paragraphs] == ["intro"]


def test_set_cell_collapses_a_multi_paragraph_cell(tmp_path: Path):
    """原单元格有两段时，`cell.text = ...` 会把它压成一段——这是本子命令的既有语义。"""
    d = Document()
    t = d.add_table(rows=1, cols=2)
    c = t.cell(0, 0)
    c.paragraphs[0].text = "old1"
    c.add_paragraph("old2")
    t.cell(0, 1).text = "keep"
    src = tmp_path / "cell2.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "cell2_out.docx"
    payload(run_edit("set-cell", src, "--table", 0, "--row", 0,
                     "--col", 0, "--text", "ONLY", "-o", out))

    got = Document(str(out))
    assert [p_.text for p_ in got.tables[0].cell(0, 0).paragraphs] == ["ONLY"]
    assert got.tables[0].cell(0, 1).text == "keep"


def test_set_cell_bad_table_index_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    """表下标越界 → 非零退出，源文件 SHA256 不变（异常在 save 之前抛出）。"""
    d = Document()
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "r0c0"
    src = tmp_path / "cellbad.docx"
    d.save(str(src))
    freeze_docx(src)
    before = sha256(src)

    cp = run_edit("set-cell", src, "--table", 3, "--row", 0,
                  "--col", 0, "--text", "x")
    assert cp.returncode != 0
    assert "IndexError" in cp.stderr
    assert sha256(src) == before


def test_set_cell_bad_row_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    """行下标越界同理。"""
    d = Document()
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "r0c0"
    src = tmp_path / "cellbad2.docx"
    d.save(str(src))
    freeze_docx(src)
    before = sha256(src)

    cp = run_edit("set-cell", src, "--table", 0, "--row", 9,
                  "--col", 0, "--text", "x")
    assert cp.returncode != 0
    assert "IndexError" in cp.stderr
    assert sha256(src) == before


# ===========================================================================
# insert
# ===========================================================================
def test_insert_lands_before_the_given_index(tmp_path: Path):
    """`--index 1` 的新段落**占住第 1 位**（= 插在原第 1 段之前），其余顺序不变。

    只用「文档里有 INSERTED」会漏掉「插到文末」的回归——所以按下标断。
    """
    src = save(tmp_path, "ins.docx", ["a", "b", "c"])
    out = tmp_path / "ins_out.docx"
    res = payload(run_edit("insert", src, "--index", 1,
                           "--text", "INSERTED", "-o", out))
    assert res["inserted_at"] == 1
    assert res["text"] == "INSERTED"

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["a", "INSERTED", "b", "c"]
    assert len(body_paragraph_elements(got)) == 4


def test_insert_applies_the_requested_style(tmp_path: Path):
    """`--style` 真落在新段上，邻段样式不变。"""
    src = save(tmp_path, "ins2.docx", ["a", "b"])
    out = tmp_path / "ins2_out.docx"
    payload(run_edit("insert", src, "--index", 1, "--text", "HEAD",
                     "--style", "Heading 1", "-o", out))

    got = Document(str(out))
    assert [(p_.text, p_.style.name) for p_ in got.paragraphs] == [
        ("a", "Normal"), ("HEAD", "Heading 1"), ("b", "Normal")]


def test_insert_beyond_end_appends_at_tail(tmp_path: Path):
    """`--index` 超出段落数 → 追加到文末（既有 fallback），JSON 回显的是入参。

    注意 `inserted_at` 是**参数回显**（99），不是新段真实落点（3）；这条差异已在
    `docs/reports/P5-A2-REPORT.md` 登记为行为意外，不在此处断言实现应该改。
    """
    src = save(tmp_path, "ins3.docx", ["a", "b", "c"])
    out = tmp_path / "ins3_out.docx"
    res = payload(run_edit("insert", src, "--index", 99,
                           "--text", "TAIL", "-o", out))
    assert res["inserted_at"] == 99
    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["a", "b", "c", "TAIL"]


def test_insert_with_unknown_style_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    """未知样式名 → 非零退出（KeyError），源文件 SHA256 不变、输出文件不产生。"""
    src = save(tmp_path, "insbad.docx", ["a", "b"])
    before = sha256(src)
    out = tmp_path / "insbad_out.docx"

    cp = run_edit("insert", src, "--index", 1, "--text", "X",
                  "--style", "Nope", "-o", out)
    assert cp.returncode != 0
    assert "KeyError" in cp.stderr
    assert sha256(src) == before
    assert not out.exists()


# ===========================================================================
# delete
# ===========================================================================
def test_delete_removes_only_the_indexed_paragraph(tmp_path: Path):
    """删中间那段：总数 -1，其余段原文不变。"""
    src = save(tmp_path, "del.docx", ["a", "b", "c"])
    out = tmp_path / "del_out.docx"
    res = payload(run_edit("delete", src, "--index", 1, "-o", out))
    assert res["deleted_index"] == 1

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["a", "c"]
    assert len(body_paragraph_elements(got)) == 2


def test_delete_last_paragraph_keeps_section_properties(tmp_path: Path):
    """删最后一段时 `w:sectPr` 必须还在 body 末尾——否则文档版面直接丢。

    python-docx 把节属性挂在 body 的最后一个 `w:sectPr` 子元素上（不是段落里），
    所以「删到只剩 0 段」这种极端操作最容易把它一起带走。
    """
    src = save(tmp_path, "del_last.docx", ["a", "b", "c"])
    out = tmp_path / "del_last_out.docx"
    payload(run_edit("delete", src, "--index", 2, "-o", out))

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["a", "b"]
    tags = [c.tag.rsplit("}", 1)[-1] for c in got.element.body]
    assert tags == ["p", "p", "sectPr"]
    assert got.sections[0].page_width is not None


def test_delete_out_of_range_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    """index 越界 → 非零退出（IndexError），源文件 SHA256 不变。"""
    src = save(tmp_path, "delbad.docx", ["a", "b"])
    before = sha256(src)

    cp = run_edit("delete", src, "--index", 99)
    assert cp.returncode != 0
    assert "IndexError" in cp.stderr
    assert sha256(src) == before


# ===========================================================================
# style
# ===========================================================================
def test_style_writes_pstyle_and_leaves_neighbours_alone(tmp_path: Path):
    """`w:pStyle@w:val` = styleId（`Heading1`，无空格），相邻段**没有** pPr。"""
    src = save(tmp_path, "sty.docx", ["p0", "p1", "p2"])
    out = tmp_path / "sty_out.docx"
    res = payload(run_edit("style", src, "--index", 1,
                           "--style", "Heading 1", "-o", out))
    assert (res["index"], res["style"]) == (1, "Heading 1")

    got = Document(str(out))
    assert pstyle_val(got.paragraphs[1]) == "Heading1"
    assert got.paragraphs[1].style.name == "Heading 1"
    assert pstyle_val(got.paragraphs[0]) is None
    assert pstyle_val(got.paragraphs[2]) is None
    # 相邻段连 pPr 都不该凭空出现
    assert got.paragraphs[0]._p.find(w("pPr")) is None


def test_style_overwrites_an_existing_pstyle(tmp_path: Path):
    """已有 pStyle 时是**改写**而不是再挂一个（同一段只能有一个 w:pStyle）。"""
    src = save(tmp_path, "sty2.docx", ["p0", "p1"])
    mid = tmp_path / "sty2_a.docx"
    payload(run_edit("style", src, "--index", 1,
                     "--style", "Heading 2", "-o", mid))
    assert pstyle_val(Document(str(mid)).paragraphs[1]) == "Heading2"

    out = tmp_path / "sty2_b.docx"
    payload(run_edit("style", mid, "--index", 1,
                     "--style", "Heading 3", "-o", out))
    got = Document(str(out))
    pPr = got.paragraphs[1]._p.find(w("pPr"))
    assert len(pPr.findall(w("pStyle"))) == 1
    assert pstyle_val(got.paragraphs[1]) == "Heading3"
    assert got.paragraphs[1].style.name == "Heading 3"


def test_style_out_of_range_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    src = save(tmp_path, "stybad.docx", ["a", "b"])
    before = sha256(src)

    cp = run_edit("style", src, "--index", 99, "--style", "Heading 1")
    assert cp.returncode != 0
    assert "IndexError" in cp.stderr
    assert sha256(src) == before


def test_style_unknown_name_exits_nonzero_and_keeps_bytes(tmp_path: Path):
    src = save(tmp_path, "stybad2.docx", ["a", "b"])
    before = sha256(src)

    cp = run_edit("style", src, "--index", 0, "--style", "No Such Style")
    assert cp.returncode != 0
    assert "KeyError" in cp.stderr
    assert sha256(src) == before


# ===========================================================================
# normalize
# ===========================================================================
def test_normalize_merges_same_format_runs_and_preserves_text(tmp_path: Path):
    """三条同格式相邻 run → 1 条，**文本逐字不变**。

    「合并不能吃字」是这个子命令唯一真正要命的地方：合并逻辑自己拼
    `pt.text + ct.text`，只要少一次拼接或把 `w:t` 找错，文本就会静默丢字。
    所以既断 run 数下降，也断逐字文本相等。
    """
    d = Document()
    p = d.add_paragraph()
    for txt in ("aa", "bb", "cc"):
        p.add_run(txt)
    src = tmp_path / "norm.docx"
    d.save(str(src))
    freeze_docx(src)
    before_text = Document(str(src)).paragraphs[0].text
    assert before_text == "aabbcc"

    out = tmp_path / "norm_out.docx"
    res = payload(run_edit("normalize", src, "-o", out))
    assert res["runs_merged"] == 2

    got = Document(str(out))
    assert len(got.paragraphs[0].runs) == 1
    assert got.paragraphs[0].text == before_text
    assert got.paragraphs[0].runs[0].text == "aabbcc"


def test_normalize_keeps_format_boundary_between_runs(tmp_path: Path):
    """`w:rPr` 不同的相邻 run 不能合并（bold 与 plain 是两条）。"""
    d = Document()
    q = d.add_paragraph()
    q.add_run("dd")
    b = q.add_run("ee")
    b.bold = True
    q.add_run("ff")
    src = tmp_path / "norm_fmt.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "norm_fmt_out.docx"
    res = payload(run_edit("normalize", src, "-o", out))
    assert res["runs_merged"] == 0

    got = Document(str(out))
    assert [(r_.text, r_.bold) for r_ in got.paragraphs[0].runs] == [
        ("dd", None), ("ee", True), ("ff", None)]
    assert got.paragraphs[0].text == "ddeeff"


def test_normalize_does_not_merge_across_a_break_run(tmp_path: Path):
    """带子元素的 run（这里是一条 `w:br`）是隔离带：它两侧不得合并。

    实现里 `kids = {子标签} - {rPr, t}` 非空就把 prev 清掉。如果这条 guard 失效，
    换行会被吞进相邻 run 的文本里——`para.text` 看起来还对，`w:br` 却没了。
    """
    d = Document()
    p = d.add_paragraph()
    p.add_run("aa")
    rb = p.add_run("bb")
    rb.add_break(WD_BREAK.LINE)
    p.add_run("cc")
    p.add_run("dd")
    src = tmp_path / "norm_br.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "norm_br_out.docx"
    res = payload(run_edit("normalize", src, "-o", out))
    # 只有 cc+dd 这一对是安全的
    assert res["runs_merged"] == 1

    got = Document(str(out))
    assert len(got.paragraphs[0].runs) == 3
    assert [r_.text for r_ in got.paragraphs[0].runs] == ["aa", "bb\n", "ccdd"]
    assert got.paragraphs[0]._p.find(".//" + w("br")) is not None, "w:br 被合并吞掉了"


def test_normalize_preserves_leading_whitespace(tmp_path: Path):
    """合并后要给 `w:t` 打 `xml:space="preserve"`，否则前导空格在 Word 里会被吃掉。"""
    d = Document()
    p = d.add_paragraph()
    p.add_run("aa")
    p.add_run(" b")
    src = tmp_path / "norm_ws.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "norm_ws_out.docx"
    payload(run_edit("normalize", src, "-o", out))

    got = Document(str(out))
    assert got.paragraphs[0].text == "aa b"
    t = got.paragraphs[0].runs[0]._r.find(w("t"))
    assert t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve"
    assert t.text == "aa b"


def test_normalize_empty_run_does_not_eat_text(tmp_path: Path):
    """空 run 合并进前一条时不能把后一条吃掉（fixture：`xx` / `` / `yy`）。"""
    d = Document()
    p = d.add_paragraph()
    p.add_run("xx")
    p.add_run("")
    p.add_run("yy")
    src = tmp_path / "norm_empty.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "norm_empty_out.docx"
    res = payload(run_edit("normalize", src, "-o", out))
    assert res["runs_merged"] == 2

    got = Document(str(out))
    assert len(got.paragraphs[0].runs) == 1
    assert got.paragraphs[0].text == "xxyy"


# ===========================================================================
# toc
# ===========================================================================
def test_toc_inserts_a_complete_toc_field_at_the_given_index(tmp_path: Path):
    """`--index 1` 前插入一个含 TOC 域的段落，域结构必须是
    `begin → instrText(TOC) → separate → 占位文本 → end` 五件齐**且顺序正确**。

    只断 `w:instrText` 里含 "TOC" 是不够的：缺 `separate` 时 Word 打开后不是
    「未更新域」而是直接显示原始域代码；缺 `end` 会让后面整篇文档都被当成域内容。
    """
    src = save(tmp_path, "toc.docx", ["alpha one", "beta two"])
    out = tmp_path / "toc_out.docx"
    res = payload(run_edit("toc", src, "--index", 1, "-o", out))
    assert res["toc_inserted_at"] == 1

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs][0] == "alpha one"
    assert got.paragraphs[0].text == "alpha one"
    assert got.paragraphs[2].text == "beta two"

    paras = body_paragraph_elements(got)
    assert len(paras) == 3
    toc_para = paras[1]
    seq = field_sequence(toc_para)
    assert [kind for kind, _ in seq][:1] == ["fldChar"]
    assert seq[0] == ("fldChar", "begin")
    assert seq[1][0] == "instrText" and "TOC" in seq[1][1]
    assert seq[2] == ("fldChar", "separate")
    assert seq[3][0] == "t"
    assert seq[4] == ("fldChar", "end")
    assert len(seq) == 5
    # 三个 fldChar 一个不少
    assert [v for k, v in seq if k == "fldChar"] == [
        "begin", "separate", "end"]
    # 域参数是 1-3 级标题 + 超链接
    assert instr_texts(toc_para) == [r' TOC \o "1-3" \h \z \u ']


def test_toc_defaults_to_index_zero(tmp_path: Path):
    """`--index` 缺省 0 → 插在**最前面**（文档第一段变成 TOC 占位段）。"""
    src = save(tmp_path, "toc0.docx", ["alpha one", "beta two"])
    out = tmp_path / "toc0_out.docx"
    res = payload(run_edit("toc", src, "-o", out))
    assert res["toc_inserted_at"] == 0

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 3
    assert "TOC" in "".join(instr_texts(paras[0]))
    assert got.paragraphs[1].text == "alpha one"
    assert got.paragraphs[2].text == "beta two"


def test_toc_beyond_end_appends_at_tail(tmp_path: Path):
    """`--index` 超出段落数 → 追加到文末（既有 fallback）。"""
    src = save(tmp_path, "toc99.docx", ["alpha one", "beta two"])
    out = tmp_path / "toc99_out.docx"
    res = payload(run_edit("toc", src, "--index", 99, "-o", out))
    assert res["toc_inserted_at"] == 99

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 3
    assert "TOC" in "".join(instr_texts(paras[-1]))
    assert got.paragraphs[0].text == "alpha one"


# ===========================================================================
# page-numbers
# ===========================================================================
def test_page_numbers_puts_page_and_numpages_fields_in_the_footer(tmp_path: Path):
    """页脚出现 `PAGE` 与 `NUMPAGES`，两个域各一套 begin/separate/end，且
    段数不增加（复用页脚第一段，而不是新开一段把自己顶到下一页）。"""
    src = save(tmp_path, "pn.docx", ["body text"])
    out = tmp_path / "pn_out.docx"
    res = payload(run_edit("page-numbers", src, "-o", out))
    assert res["footer_fields"] == ["PAGE", "NUMPAGES"]

    got = Document(str(out))
    footer = got.sections[0].footer
    assert len(footer.paragraphs) == 1

    fp = footer.paragraphs[0]
    assert instr_texts(fp._p) == [" PAGE ", " NUMPAGES "]
    fld_types = [f.get(w("fldCharType")) for f in fp._p.iter(w("fldChar"))]
    assert fld_types == ["begin", "separate", "end",
                         "begin", "separate", "end"]
    assert fp.text == "Page 1 of 1"
    # 正文一个字不动
    assert [p_.text for p_ in got.paragraphs] == ["body text"]


def test_page_numbers_appends_to_an_existing_footer_paragraph(tmp_path: Path):
    """页脚原有文字不被清掉，页码接在它后面（既有语义：追加，不覆盖）。"""
    d = Document()
    d.add_paragraph("x")
    d.sections[0].footer.paragraphs[0].text = "prev"
    src = tmp_path / "pn_prev.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "pn_prev_out.docx"
    payload(run_edit("page-numbers", src, "-o", out))

    got = Document(str(out))
    assert got.sections[0].footer.paragraphs[0].text == "prevPage 1 of 1"


def test_page_numbers_run_twice_duplicates_the_fields(tmp_path: Path):
    """**非幂等**：跑两次页码会追加第二套域。

    这不是本子命令的回归，而是当前实现的既有行为（`cmd_page_numbers` 无去重判断）。
    钉住它有两个用处：① 任何人加去重逻辑时会在这里红，逼他确认是不是有意的；
    ② 报告里登记为「重复执行会叠加」的证据。
    """
    src = save(tmp_path, "pn_twice.docx", ["body text"])
    once = tmp_path / "pn_once.docx"
    payload(run_edit("page-numbers", src, "-o", once))
    twice = tmp_path / "pn_twice_out.docx"
    payload(run_edit("page-numbers", once, "-o", twice))

    got = Document(str(twice))
    fp = got.sections[0].footer.paragraphs[0]
    assert fp.text == "Page 1 of 1Page 1 of 1"
    assert instr_texts(fp._p) == [" PAGE ", " NUMPAGES ", " PAGE ", " NUMPAGES "]


# ===========================================================================
# 行为记录（当前实现如此，钉住是为了「谁改了必须显式知道」）
# ===========================================================================
@pytest.mark.parametrize("cmd,extra,expected", [
    ("delete", ["--index", -1], ["a", "b"]),
    ("style", ["--index", -1, "--style", "Heading 1"], ["a", "b", "c"]),
    ("insert", ["--index", -1, "--text", "NEG"], ["a", "b", "NEG", "c"]),
])
def test_negative_index_uses_python_semantics(tmp_path: Path, cmd, extra, expected):
    """负数 `--index` 不报错，按 Python 负索引解释（-1 = 最后一段）。

    `docx_edit.py` 的按 index 取段的子命令都是直接 `doc.paragraphs[args.index]`
    与 `args.index < len(paras)` 判断，CLI 层没有任何 `>= 0` 校验。于是：
    `delete --index -1` 删**最后一段**、`style --index -1` 给**最后一段**套样式、
    `insert --index -1` 落在最后一段**之前**。这是既有行为，不是本轮的期望值；
    钉住它是为了让「哪天有人加了负索引校验」时这里有红点，逼他确认是不是有意的。
    """
    src = save(tmp_path, f"neg_{cmd}.docx", ["a", "b", "c"])
    out = tmp_path / f"neg_{cmd}_out.docx"
    res = payload(run_edit(cmd, src, *extra, "-o", out))
    assert res["ok"] is True

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == expected
    if cmd == "style":
        assert pstyle_val(got.paragraphs[-1]) == "Heading1"
    if cmd == "delete":
        # 删掉的是末段 c，不是首段 a
        assert "c" not in expected


def test_toc_negative_index_inserts_before_last_paragraph(tmp_path: Path):
    """`toc --index -1` 同样落进 Python 负索引语义：TOC 段插在**最后一段之前**。"""
    src = save(tmp_path, "neg_toc.docx", ["a", "b", "c"])
    out = tmp_path / "neg_toc_out.docx"
    payload(run_edit("toc", src, "--index", -1, "-o", out))

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 4
    assert "TOC" in "".join(instr_texts(paras[2]))
    assert [p_.text for p_ in got.paragraphs][0] == "a"
    assert got.paragraphs[-1].text == "c"


def test_paragraph_index_ignores_table_cells(tmp_path: Path):
    """`--index` 数的是 `doc.paragraphs`（**只含 body 段落**），表格单元格里的段落不计入。

    fixture：body 段 `p0` → 1×1 表（格内文本 `cell`）→ body 段 `p1`。
    `doc.paragraphs` 只有 2 项。两向都断：

    * `delete --index 1` 删的是 `p1`，表格必须完好；
    * `delete --index 2` 必须**越界报错**——如果 index 改成走
      `iter_all_paragraphs`（会递归进表格），第 2 项恰好是格内段落，就会静默删掉
      表格内容并 exit 0。这一向才是真正钉住「index 语义」的那条。
    """
    d = Document()
    d.add_paragraph("p0")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "cell"
    d.add_paragraph("p1")
    src = tmp_path / "withtable.docx"
    d.save(str(src))
    freeze_docx(src)
    assert [p_.text for p_ in Document(str(src)).paragraphs] == ["p0", "p1"]

    out = tmp_path / "withtable_out.docx"
    payload(run_edit("delete", src, "--index", 1, "-o", out))

    got = Document(str(out))
    assert [p_.text for p_ in got.paragraphs] == ["p0"]
    assert len(got.tables) == 1
    assert got.tables[0].cell(0, 0).text == "cell"

    # index 2 越界：表格段落不在 doc.paragraphs 里
    before = sha256(src)
    cp = run_edit("delete", src, "--index", 2)
    assert cp.returncode != 0
    assert "IndexError" in cp.stderr
    assert sha256(src) == before
    assert Document(str(src)).tables[0].cell(0, 0).text == "cell"


def test_replace_with_empty_needle_is_a_noop(tmp_path: Path):
    """`--find ""` → 0 命中且正文逐字不变（`replace_in_paragraph` 对空针直接返回 0）。

    空针如果没有这道 guard，"在每个位置插入"会让整篇文档膨胀成灾难。
    """
    src = save(tmp_path, "empty_find.docx", ["abc", "def"])
    out = tmp_path / "empty_find_out.docx"
    res = payload(run_edit("replace", src, "--find", "", "--replace", "X", "-o", out))
    assert res["replacements"] == 0
    assert [p_.text for p_ in Document(str(out)).paragraphs] == ["abc", "def"]


# ===========================================================================
# 全子命令共同契约
# ===========================================================================
def test_every_subcommand_reports_json_with_ok_true(tmp_path: Path):
    """八个子命令都必须：退出 0 + stdout 是纯 JSON + `ok: true` + `output` 指向产物。

    `SKILL.md` 把 `docx_edit` 归到 "JSON on stdout"，任何一行 `print` 调试残留都会
    把这条打红——这正是要守的。
    """
    src = tmp_path / "all.docx"
    d = Document()
    d.add_paragraph("alpha x")
    d.add_paragraph("beta y")
    t = d.add_table(rows=1, cols=1)
    t.cell(0, 0).text = "c00"
    d.save(str(src))
    freeze_docx(src)

    cases = {
        "replace": ("replace", "--find", "alpha", "--replace", "ALPHA"),
        "set-cell": ("set-cell", "--table", 0, "--row", 0, "--col", 0,
                     "--text", "NEWCELL"),
        "insert": ("insert", "--index", 1, "--text", "INS"),
        "delete": ("delete", "--index", 1),
        "style": ("style", "--index", 0, "--style", "Heading 1"),
        "normalize": ("normalize",),
        "toc": ("toc", "--index", 0),
        "page-numbers": ("page-numbers",),
    }
    assert len(cases) == 8

    for name, argv in cases.items():
        out = tmp_path / f"all_{name}.docx"
        cp = run_edit(*argv, src, "-o", out)
        res = payload(cp)
        assert res["ok"] is True, name
        assert res["output"] == str(out), name
        assert out.is_file(), name
        # 产物必须还能被 python-docx 打开（写盘没写坏 zip）
        Document(str(out))
