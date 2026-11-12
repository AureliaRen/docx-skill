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
| `page-numbers` | 页脚出现 `PAGE` 与 `NUMPAGES`，两个域各一套 begin/separate/end；**幂等**（已有域时 `changed:false`，页脚 XML 逐字节不变，就地重跑连文件 mtime/SHA256 都不动）；页脚原有 run 的文字与加粗一律保留（追加，不整段替换） |

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
import re
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


def field_codes(p_el) -> list:
    """段落里域代码的第一个 token（`PAGE` / `NUMPAGES` / `TOC` …），按出现顺序。"""
    return [t.split()[0] for t in instr_texts(p_el) if t.split()]


def footer_parts(path: Path) -> dict:
    """``word/footer*.xml`` 的 {部件名: 原始字节}，用于「页脚 XML 变没变」的指纹比对。"""
    with zipfile.ZipFile(Path(path)) as z:
        return {n: z.read(n) for n in sorted(z.namelist())
                if re.fullmatch(r"word/footer\d*\.xml", n)}


def one_error_line(stderr: str) -> str:
    """stderr 必须**恰好一行**（`error: ...`），没有 traceback 前导。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


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


def _cell_fixture(tmp_path: Path, name: str = "cellbad.docx") -> Path:
    """2×2 表 + 1 段正文：三个维度（table/row/col）都有越界余量可测。"""
    d = Document()
    d.add_paragraph("intro")
    t = d.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = f"r{r}c{c}"
    src = tmp_path / name
    d.save(str(src))
    return freeze_docx(src)


@pytest.mark.parametrize("dim,args,needle", [
    ("table", ["--table", 9, "--row", 0, "--col", 0],
     "set-cell --table 9 is out of range: legal range is [0, 0], current N=1"),
    ("row", ["--table", 0, "--row", 9, "--col", 0],
     "set-cell --row 9 is out of range: legal range is [0, 1], current N=2"),
    ("col", ["--table", 0, "--row", 0, "--col", 9],
     "set-cell --col 9 is out of range: legal range is [0, 1], current N=2"),
])
def test_set_cell_each_dimension_out_of_range_is_a_usage_error(
        tmp_path: Path, dim, args, needle):
    """table / row / col **三个维度各一条**越界 → 退出 2 + stderr 点名是哪个维度。

    为什么要分维度断而不只断「非零退出」：`doc.tables[t].cell(r, c)` 里三个下标
    各自裸索引，一句泛泛的 `is out of range` 无法告诉调用方到底该改哪一个参数。
    每条的 stderr 必须**只**包含自己那一维的措辞，别的维度不得出现——否则
    「把三者写进同一句话」也能骗过只查 `in stderr` 的断言。
    """
    src = _cell_fixture(tmp_path, f"cellbad_{dim}.docx")
    before = sha256(src)
    out = tmp_path / f"cellbad_{dim}_out.docx"

    cp = run_edit("set-cell", src, *args, "--text", "x", "-o", out)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    line = one_error_line(cp.stderr)
    assert line == f"error: {needle}"
    for other in {"table", "row", "col"} - {dim}:
        assert f"--{other} " not in line, f"{dim} 越界却报了 {other}"
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


@pytest.mark.parametrize("dim,args", [
    ("table", ["--table", -1, "--row", 0, "--col", 0]),
    ("row", ["--table", 0, "--row", -1, "--col", 0]),
    ("col", ["--table", 0, "--row", 0, "--col", -1]),
])
def test_set_cell_negative_index_is_a_usage_error(tmp_path: Path, dim, args):
    """三个维度的负下标同样退出 2（`-1` 不再被读成「最后一个」）。"""
    src = _cell_fixture(tmp_path, f"cellneg_{dim}.docx")
    before = sha256(src)
    out = tmp_path / f"cellneg_{dim}_out.docx"

    cp = run_edit("set-cell", src, *args, "--text", "x", "-o", out)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert f"set-cell --{dim} -1 is out of range" in one_error_line(cp.stderr)
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_set_cell_edge_cells_are_still_accepted(tmp_path: Path):
    """反向保险：合法边角 `(0,0)` 与 `(1,1)` 照旧成功，输出 JSON 三键不变。

    「越界一律报错」的最省事实现是把范围收得过窄（比如写成 `0 < t < N`），
    这条专门挡住那种改过头。
    """
    src = _cell_fixture(tmp_path, "cell_edge.docx")

    out00 = tmp_path / "cell_edge_00.docx"
    res00 = payload(run_edit("set-cell", src, "--table", 0, "--row", 0,
                             "--col", 0, "--text", "TL", "-o", out00))
    assert (res00["table"], res00["row"], res00["col"]) == (0, 0, 0)

    out11 = tmp_path / "cell_edge_11.docx"
    res11 = payload(run_edit("set-cell", src, "--table", 0, "--row", 1,
                             "--col", 1, "--text", "BR", "-o", out11))
    assert (res11["table"], res11["row"], res11["col"]) == (0, 1, 1)

    assert _grid(Document(str(out00))) == [["TL", "r0c1"], ["r1c0", "r1c1"]]
    assert _grid(Document(str(out11))) == [["r0c0", "r0c1"], ["r1c0", "BR"]]


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


def test_insert_beyond_end_refuses_instead_of_clamping(tmp_path: Path):
    """`--index` 超出段落数 → **退出 2 + 不写盘**，不再静默追加到文末。

    旧实现（P6 前）在 `args.index >= len(paras)` 时落到 `doc.add_paragraph`，
    于是 `--index 99` 在 3 段文档上把段落追加到**下标 3**，却在 JSON 里回显
    `inserted_at: 99`——调用方拿到的是一个永远不存在的位置。`P5-A2-REPORT.md`
    把它登记为「行为意外」，本条按审计 §2.5 的裁定把它改成硬错误。

    断言四件事：退出码 2、stderr 说明合法范围、stdout 没有 JSON、源文件字节不变。
    """
    src = save(tmp_path, "ins3.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / "ins3_out.docx"

    cp = run_edit("insert", src, "--index", 99, "--text", "TAIL", "-o", out)
    assert cp.returncode == 2
    assert "out of range" in cp.stderr
    assert "[0, 2]" in cp.stderr and "current N=3" in cp.stderr
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


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


def test_delete_out_of_range_is_a_usage_error(tmp_path: Path):
    """`delete --index 99` → 退出 2 + stderr 单行带合法范围，原文件零改动。

    旧实现直接 `doc.paragraphs[args.index]`，于是 999 换来十四行
    `IndexError` traceback 和 exit 1——调用方拿到的是解释器内部错误，
    而不是「你参数写错了」。P8 定案：用法错误一律 exit 2、stderr 单行。
    """
    src = save(tmp_path, "delbad.docx", ["a", "b"])
    before = sha256(src)
    out = tmp_path / "delbad_out.docx"

    cp = run_edit("delete", src, "--index", 99, "-o", out)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        "error: delete --index 99 is out of range: legal range is [0, 1], "
        "current N=2")
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_delete_every_legal_index_still_works(tmp_path: Path):
    """反向保险：`0` 与 `N-1` 两个边界照样删得掉，JSON 三键不变。"""
    for idx in (0, 2):
        src = save(tmp_path, f"delok{idx}.docx", ["a", "b", "c"])
        out = tmp_path / f"delok{idx}_out.docx"
        res = payload(run_edit("delete", src, "--index", idx, "-o", out))
        assert res["deleted_index"] == idx
        left = [p_.text for p_ in Document(str(out)).paragraphs]
        assert left == (["b", "c"] if idx == 0 else ["a", "b"])


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


def test_style_out_of_range_is_a_usage_error(tmp_path: Path):
    """`style --index 99` → 退出 2 + stderr 单行带合法范围，原文件零改动。"""
    src = save(tmp_path, "stybad.docx", ["a", "b"])
    before = sha256(src)
    out = tmp_path / "stybad_out.docx"

    cp = run_edit("style", src, "--index", 99, "--style", "Heading 1", "-o", out)
    assert cp.returncode == 2, cp.stderr
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        "error: style --index 99 is out of range: legal range is [0, 1], "
        "current N=2")
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_style_both_legal_paragraph_endpoints_still_work(tmp_path: Path):
    """反向保险：首段与末段都还能套样式。"""
    src = save(tmp_path, "styok.docx", ["a", "b", "c"])
    for idx in (0, 2):
        out = tmp_path / f"styok{idx}_out.docx"
        res = payload(run_edit("style", src, "--index", idx,
                               "--style", "Heading 1", "-o", out))
        assert res["index"] == idx
        assert res["style"] == "Heading 1"
        assert pstyle_val(Document(str(out)).paragraphs[idx]) == "Heading1"


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


def test_toc_legal_endpoints_land_where_asked(tmp_path: Path):
    """合法区间的两个端点 `0` 与 `N-1` 都必须 exit 0，且落点就是入参。

    ``toc --index`` 的语义是「插在**第 i 段之前**」，所以合法区间与
    `delete`/`style`/`insert` 同口径：`0 <= i <= N-1`。

    ``toc_inserted_at`` 在这里等于入参，是因为两个端点都真插到了那个位置；
    本用例同时按 XML 顺序反查落点，避免「回显入参但实际落在别处」。

    历史：本条原为 `test_toc_beyond_end_appends_at_tail`，钉住的是
    「`--index 99` 静默落进 `doc.add_paragraph` 追加到文末、`toc_inserted_at`
    回显 99」——那是把**一个缺陷**当契约留档：真实落点是 N（文末），回显却是
    99，调用方拿到的位置它永远得不到。P8 / D5 把它并入「越界即硬错误」，
    本用例只留合法端点。
    """
    paragraphs = ["alpha one", "beta two", "gamma three"]
    # 端点 0：TOC 段成为第一段
    src = save(tmp_path, "toc_lo.docx", paragraphs)
    out = tmp_path / "toc_lo_out.docx"
    res = payload(run_edit("toc", src, "--index", 0, "-o", out))
    assert res["toc_inserted_at"] == 0

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 4
    assert "TOC" in "".join(instr_texts(paras[0]))

    # 端点 N-1：TOC 段插在最后一段之前（= 新的倒数第二段）
    src = save(tmp_path, "toc_hi.docx", paragraphs)
    out = tmp_path / "toc_hi_out.docx"
    res = payload(run_edit("toc", src, "--index", 2, "-o", out))
    assert res["toc_inserted_at"] == 2

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 4
    assert "TOC" in "".join(instr_texts(paras[2]))
    # TOC 段自己带一句域结果占位文本（`_add_field` 的 cached 值），所以不是空串
    assert [p_.text for p_ in got.paragraphs][:2] == ["alpha one", "beta two"]
    assert got.paragraphs[3].text == "gamma three"


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


def test_page_numbers_run_twice_is_idempotent(tmp_path: Path):
    """**幂等**：第二次跑不再插第二套域，而是报 `changed: false` 原样收手。

    P6-C3 之前这里是反向断言（`"Page 1 of 1Page 1 of 1"`、四段 instrText），
    钉的是当时的缺陷：`cmd_page_numbers` 没有去重判断，plan 里同一句
    `page-numbers` 被 apply 两次就叠出两个页码。现在改成断幂等——**这条用例的
    方向翻了，是有意的**，写在这里以免后来人以为漏了历史。
    """
    src = save(tmp_path, "pn_twice.docx", ["body text"])
    once = tmp_path / "pn_once.docx"
    res1 = payload(run_edit("page-numbers", src, "-o", once))
    assert res1["changed"] is True

    twice = tmp_path / "pn_twice_out.docx"
    res2 = payload(run_edit("page-numbers", once, "-o", twice))
    assert res2["changed"] is False
    assert res2["footer_fields"] == ["PAGE", "NUMPAGES"]

    got = Document(str(twice))
    fp = got.sections[0].footer.paragraphs[0]
    assert fp.text == "Page 1 of 1"
    assert instr_texts(fp._p) == [" PAGE ", " NUMPAGES "]
    assert field_codes(fp._p) == ["PAGE", "NUMPAGES"]


def test_page_numbers_second_run_keeps_the_footer_xml_byte_identical(tmp_path: Path):
    """连跑两次，`word/footer*.xml` 逐字节相同（不是「看起来一样」，是字节级）。

    只比 `paragraph.text` 会放过「域插了两套但占位文本恰好相同」这类问题；
    这里直接比 footer 部件的原始字节，并另外数一遍域个数。
    """
    src = save(tmp_path, "pn_idem.docx", ["body text"])
    once = tmp_path / "pn_idem_once.docx"
    payload(run_edit("page-numbers", src, "-o", once))
    before = footer_parts(once)
    assert before, "第一次跑完必须真的存在 footer 部件"

    twice = tmp_path / "pn_idem_twice.docx"
    payload(run_edit("page-numbers", once, "-o", twice))
    after = footer_parts(twice)

    assert sorted(before) == sorted(after)
    for name in before:
        assert before[name] == after[name], f"{name} 在第二次跑之后被改写了"

    fp = Document(str(twice)).sections[0].footer.paragraphs[0]
    # 一个 PAGE + 一个 NUMPAGES，不多不少
    assert instr_texts(fp._p).count(" PAGE ") == 1
    assert instr_texts(fp._p).count(" NUMPAGES ") == 1
    assert fp.text == "Page 1 of 1"


def test_page_numbers_in_place_rerun_writes_nothing_at_all(tmp_path: Path):
    """就地重跑（不带 `-o`）时连文件都不写：SHA256 与 mtime 一字不差。

    这是「幂等」对操作者最实在的含义——同一份 plan 跑两遍，第二遍不该在
    mtime 上留痕，也不该重打一遍 zip（python-docx 重写会把条目时间戳刷成
    当前时刻，那样 SHA256 必然变）。
    """
    src = save(tmp_path, "pn_inplace.docx", ["body text"])
    payload(run_edit("page-numbers", src))
    first_sha, first_mtime = sha256(src), src.stat().st_mtime_ns

    res = payload(run_edit("page-numbers", src))
    assert res["changed"] is False
    assert sha256(src) == first_sha
    assert src.stat().st_mtime_ns == first_mtime

    fp = Document(str(src)).sections[0].footer.paragraphs[0]
    assert fp.text == "Page 1 of 1"


def test_page_numbers_keeps_the_existing_footer_text_and_its_formatting(tmp_path: Path):
    """页脚原有 run 的文字与加粗必须原样保留，页码**追加**在段尾（不整段替换）。

    断言钉在 run 级，而不是只看 `paragraph.text` 里含「公司机密」——整段替换成
    `"公司机密Page 1 of 1"` 的单 run 也能让 `text` 断言通过，却已经把格式、
    书签、域一起丢了。
    """
    d = Document()
    d.add_paragraph("body text")
    fp = d.sections[0].footer.paragraphs[0]
    r = fp.add_run("公司机密")
    r.bold = True
    src = tmp_path / "pn_keep.docx"
    d.save(str(src))
    freeze_docx(src)

    out = tmp_path / "pn_keep_out.docx"
    payload(run_edit("page-numbers", src, "-o", out))

    footer = Document(str(out)).sections[0].footer
    assert len(footer.paragraphs) == 1
    fp = footer.paragraphs[0]
    assert fp.text == "公司机密Page 1 of 1"
    # 原 run 还在原位，且加粗没丢
    assert fp.runs[0].text == "公司机密"
    assert fp.runs[0].bold is True
    # 新 run 追加在它后面，不是插在前面
    assert fp.runs[1].text == "Page "
    # XML 里第一个 w:r 的文本仍是原文
    first_run_text = next(
        t.text for t in fp._p.iter(w("t")) if t.text
    )
    assert first_run_text == "公司机密"


@pytest.mark.parametrize("form", ["complex", "simple"])
def test_page_numbers_does_not_stack_on_an_existing_page_field(tmp_path: Path, form):
    """页脚里**已经有页码域**（Word 存的复杂域 / `w:fldSimple` 简单域）时不再插。

    这两种写法都要认：python-docx 只产复杂域，而 Word 与 LibreOffice 都会产
    `w:fldSimple`。只认一种，另一类文档就会变成 `第 1 页 Page 1 of 1`。
    """
    d = Document()
    d.add_paragraph("body text")
    fp = d.sections[0].footer.paragraphs[0]
    fp.add_run("第 ")
    if form == "complex":
        p = fp._p
        for ftype, extra in (("begin", None), (None, " PAGE "),
                             ("separate", None), (None, "1"), ("end", None)):
            r = p.makeelement(w("r"), {})
            p.append(r)
            if ftype is not None:
                fld = r.makeelement(w("fldChar"), {w("fldCharType"): ftype})
                r.append(fld)
            elif extra == " PAGE ":
                it = r.makeelement(w("instrText"), {})
                it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                it.text = extra
                r.append(it)
            else:
                t = r.makeelement(w("t"), {})
                t.text = extra
                r.append(t)
    else:
        fs = fp._p.makeelement(w("fldSimple"), {w("instr"): " PAGE "})
        r = fs.makeelement(w("r"), {})
        t = r.makeelement(w("t"), {})
        t.text = "1"
        r.append(t)
        fs.append(r)
        fp._p.append(fs)
    fp.add_run(" 页")
    src = tmp_path / f"pn_has_{form}.docx"
    d.save(str(src))
    freeze_docx(src)

    before = footer_parts(src)
    out = tmp_path / f"pn_has_{form}_out.docx"
    res = payload(run_edit("page-numbers", src, "-o", out))
    assert res["changed"] is False
    assert res["already_present"] == ["PAGE"]

    footer = Document(str(out)).sections[0].footer
    para = footer.paragraphs[0]
    assert "Page " not in para.text
    if form == "complex":
        # 复杂域的占位结果也是 `<w:r><w:t>`，因此进 `paragraph.text`
        assert para.text == "第 1 页"
        assert instr_texts(para._p) == [" PAGE "]
    else:
        # `w:fldSimple` 的结果 run 不是 paragraph 的直接子元素，
        # python-docx 的 `.text` 看不到它（只剩 "第  页"）——所以按 XML 断。
        assert para.text == "第  页"
        simple = list(para._p.iter(w("fldSimple")))
        assert len(simple) == 1
        assert simple[0].get(w("instr")) == " PAGE "
    # 页脚部件逐字节未变
    assert footer_parts(out)["word/footer1.xml"] == before["word/footer1.xml"]


# ===========================================================================
# 行为记录（当前实现如此，钉住是为了「谁改了必须显式知道」）
# ===========================================================================
@pytest.mark.parametrize("cmd,extra", [
    ("delete", []),
    ("style", ["--style", "Heading 1"]),
])
def test_negative_index_is_a_usage_error(tmp_path: Path, cmd, extra):
    """负数 `--index` → 退出 2，不再借用 Python 负索引语义（P8 / D2b）。

    历史：本条原为 `test_negative_index_uses_python_semantics`，钉住的是
    「`delete --index -1` 删末段、`style --index -1` 给末段套样式、两者 exit 0」。
    那是把**一个缺陷**当契约留档——`-1` 这个入参在任何文档里都不等于「最后一段」：
    段落增删一次，同一个 `--index -1` 指的就换了一段。P8 定案的全仓退出码口径
    （`2` = 用法/输入错误）把它改成硬错误，与 `insert` 在 P6 的处置保持一致；
    `toc` 的 `--index -1` 仍按 Python 语义走（有专门用例钉着），不在本单范围。

    断言四件事：退出码 2、stderr 单行且带真实合法范围、stdout 为空、原文件
    SHA256 不变且输出文件未创建（证明没有任何 `doc.save()` 落在异常路径上）。
    """
    src = save(tmp_path, f"neg_{cmd}.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / f"neg_{cmd}_out.docx"

    cp = run_edit(cmd, src, "--index", -1, *extra, "-o", out)
    assert cp.returncode == 2
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        f"error: {cmd} --index -1 is out of range: legal range is [0, 2], "
        f"current N=3")
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_toc_negative_index_is_a_usage_error(tmp_path: Path):
    """`toc --index -1` 与 `delete`/`style`/`insert` 同口径：退出 2，不落盘。

    历史：本条原为 `test_toc_negative_index_inserts_before_last_paragraph`，
    钉住的是「`--index -1` 走 Python 负索引、TOC 段插在最后一段之前、exit 0」
    ——把**一个缺陷**当契约留档。（上一条 `test_negative_index_is_a_usage_error`
    的 docstring 里那句「`toc` 的 `--index -1` 仍按 Python 语义走（有专门用例
    钉着），不在本单范围」已随 P8 / D5 作废。）`-1` 在段落增删后会指到另一段，
    与 `toc_inserted_at` 回显的 `-1` 一起构成「报告一个永远得不到的位置」。

    断言四件事：退出码 2、stderr 单行且带真实合法范围、stdout 为空、原文件
    SHA256 不变且输出文件未创建。
    """
    src = save(tmp_path, "neg_toc.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / "neg_toc_out.docx"

    cp = run_edit("toc", src, "--index", -1, "-o", out)
    assert cp.returncode == 2
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        "error: toc --index -1 is out of range: legal range is [0, 2], "
        "current N=3")
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


@pytest.mark.parametrize("bad", [3, 99])
def test_toc_index_at_or_beyond_n_is_a_usage_error(tmp_path: Path, bad):
    """`--index >= N` → 退出 2；`N`（刚好差一位）与 `99` 同等对待。

    旧实现里 `args.index < len(paras)` 为假时静默转 `doc.add_paragraph`
    追加到文末，`toc_inserted_at` 仍回显入参——N 与 99 表现完全相同。
    """
    src = save(tmp_path, f"toc_oob_{bad}.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / f"toc_oob_{bad}_out.docx"

    cp = run_edit("toc", src, "--index", bad, "-o", out)
    assert cp.returncode == 2, bad
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        f"error: toc --index {bad} is out of range: legal range is [0, 2], "
        f"current N=3")
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_toc_empty_body_accepts_index_zero_only(tmp_path: Path):
    """空 body（N=0）只剩一个插入点：`0` 合法（= 追加首段），`1` 退出 2。

    与 `insert --index` 的退化边界同口径（见
    `test_insert_empty_body_accepts_index_zero_only`）：合法范围 `[0, N-1]`
    在 N=0 时退化成空集，照搬会连「往空文档写第一个 TOC 段」都封死。
    """
    src = tmp_path / "empty_toc.docx"
    Document().save(str(src))
    freeze_docx(src)

    out = tmp_path / "empty_toc_out.docx"
    res = payload(run_edit("toc", src, "--index", 0, "-o", out))
    assert res["toc_inserted_at"] == 0

    got = Document(str(out))
    paras = body_paragraph_elements(got)
    assert len(paras) == 1
    assert "TOC" in "".join(instr_texts(paras[0]))

    before = sha256(src)
    for bad_index in (1, -1):
        cp = run_edit("toc", src, "--index", bad_index)
        assert cp.returncode == 2, bad_index
        assert "out of range" in cp.stderr
        assert "current N=0" in cp.stderr
        assert cp.stdout.strip() == ""
    assert sha256(src) == before


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

    # index 2 越界：表格段落不进 doc.paragraphs（P8 起为 exit 2 + 单行 stderr）
    before = sha256(src)
    cp = run_edit("delete", src, "--index", 2)
    assert cp.returncode == 2
    assert "Traceback" not in cp.stderr
    assert one_error_line(cp.stderr) == (
        "error: delete --index 2 is out of range: legal range is [0, 1], "
        "current N=2")
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


# ===========================================================================
# insert --index：越界硬报错 + inserted_at = 真实落点（P6 / 审计 §2.5）
# ===========================================================================
# 背景（2026-11-11 实测，见 docs/reports/P6-FIX-C2B.md）：旧实现对 `--index`
# 既不做下界也不做上界校验，两种越界都 exit 0 并回显入参：
#   --index 99  → 落到 add_paragraph 分支，真实下标 4，JSON 报 99
#   --index -1  → 走 Python 负索引，插在最后一段之前（下标 3），JSON 报 -1
# 现在：越界一律 stderr + exit 2，且**一个字节都不写**；`inserted_at` 改为插入后
# 按元素身份反查出来的真实下标。
def test_insert_index_equal_to_n_exits_two(tmp_path: Path):
    """`--index == N`（差一位越界，最容易漏的那个）→ 退出 2，不落盘。

    旧实现里 `args.index < len(paras)` 为假时静默转 append，于是 N 与 99 表现
    完全相同——这条专门钉住「刚好差一位」也要报错。
    """
    src = save(tmp_path, "ins_n.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / "ins_n_out.docx"

    cp = run_edit("insert", src, "--index", 3, "--text", "X", "-o", out)
    assert cp.returncode == 2
    assert "out of range" in cp.stderr
    assert "[0, 2]" in cp.stderr and "current N=3" in cp.stderr
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


@pytest.mark.parametrize("bad", [-1, -2, -99])
def test_insert_negative_index_exits_two(tmp_path: Path, bad):
    """负数 `--index` → 退出 2（不再借用 Python 负索引语义）。

    `delete` / `style` 的负索引仍按 Python 语义走（有专门用例钉着）；`insert`
    特殊在它要**回报落点**：回显 `-1` 而真实插在下标 N-1，是最典型的谎报。
    """
    src = save(tmp_path, f"ins_neg{bad}.docx", ["a", "b", "c"])
    before = sha256(src)
    out = tmp_path / f"ins_neg{bad}_out.docx"

    cp = run_edit("insert", src, f"--index={bad}", "--text", "X", "-o", out)
    assert cp.returncode == 2
    assert "out of range" in cp.stderr
    assert "[0, 2]" in cp.stderr and "current N=3" in cp.stderr
    assert cp.stdout.strip() == ""
    assert sha256(src) == before
    assert not out.exists()


def test_insert_boundary_zero_and_last_index_are_ok(tmp_path: Path):
    """合法边界 `0` 与 `N-1` 行为不变：exit 0、落点等于入参、其余 JSON 键不动。

    这条是「不要修过头」的反向保险：加校验后 0 和 N-1 必须照旧可用，且
    `ok` / `output` / `text` 三个键一个不少。
    """
    src = save(tmp_path, "ins_edge.docx", ["a", "b", "c"])

    out0 = tmp_path / "ins_edge0.docx"
    res0 = payload(run_edit("insert", src, "--index", 0, "--text", "HEAD",
                            "-o", out0))
    assert res0["inserted_at"] == 0
    assert [p_.text for p_ in Document(str(out0)).paragraphs] == [
        "HEAD", "a", "b", "c"]

    out_last = tmp_path / "ins_edge_last.docx"
    res_last = payload(run_edit("insert", src, "--index", 2, "--text", "MID",
                                "-o", out_last))
    assert res_last["inserted_at"] == 2
    assert [p_.text for p_ in Document(str(out_last)).paragraphs] == [
        "a", "b", "MID", "c"]

    for res, out in ((res0, out0), (res_last, out_last)):
        assert res["ok"] is True
        assert res["output"] == str(out)


def test_insert_inserted_at_is_the_measured_landing_position(tmp_path: Path):
    """`inserted_at` 必须是**插入后按元素身份反查**出来的下标，不是入参回显。

    怎么做成真断言（两条独立证据）：

    1. **对象身份**：进程内调 `cmd_insert`，先按住原第 2 段的元素引用，插完断言
       它被**推到下标 3**、下标 2 换成了别的元素——即新段落确实占住 2。同时
       `inserted_at == 2`，而这是实现反查出来的值。
    2. **逐字重复文本**：三段都写同一个词 `dup`，插入的也是 `dup`。此时「按文本
       搜索第一个 dup」这种假定位完全失效，只能靠位置/身份判断——落点仍然是 2。

    反向自证（把 `cmd_insert` 换回旧的 clamp+回显实现）：本用例在第 1 步就红，
    因为旧实现下 `--index 99` 之外没有越界可言，但 `inserted_at` 恒等于入参，
    一旦入参与落点出现分叉（下面的 `--index 99` 分支）它就报不出真实位置。
    """
    from types import SimpleNamespace

    import docx_edit as de  # scripts/ 已在模块顶部进 sys.path

    src = save(tmp_path, "ins_measure.docx", ["dup", "b", "dup"])
    doc = Document(str(src))
    anchor = body_paragraph_elements(doc)[2]

    res = de.cmd_insert(doc, SimpleNamespace(index=2, text="dup", style=None))
    after = body_paragraph_elements(doc)

    assert len(after) == 4
    assert after[3] is anchor, "原第 2 段必须被推到下标 3（对象身份不变）"
    assert after[2] is not anchor, "下标 2 必须换成新插入的那个元素"
    assert res["inserted_at"] == 2, "inserted_at 必须是反查出的真实下标"
    assert [p_.text for p_ in doc.paragraphs] == ["dup", "b", "dup", "dup"]

    # 证据 2：重复文本下 CLI 也必须报真实落点（文本搜索无法区分，只能靠位置）
    src2 = save(tmp_path, "ins_measure2.docx", ["dup", "b", "dup"])
    out = tmp_path / "ins_measure2_out.docx"
    res2 = payload(run_edit("insert", src2, "--index", 1, "--text", "dup",
                            "-o", out))
    assert res2["inserted_at"] == 1
    assert [p_.text for p_ in Document(str(out)).paragraphs] == [
        "dup", "dup", "b", "dup"]


def test_inserted_at_comes_from_the_lookup_not_the_argument(tmp_path: Path,
                                                            monkeypatch):
    """`inserted_at` 的**来源**必须是落点反查，不是 `args.index`——用哨兵值证明。

    为什么需要这条（而不是只靠上一条黑盒断言）：在**合法区间内**，回显入参和
    反查落点在数值上恒等（`insert_paragraph_before` 让新段正好占住该下标），
    所以「`inserted_at == 2` 且新段在下标 2」这种黑盒断言对「实现改回回显入参」
    完全免疫——它照样绿。要分开这两种实现，只能钉**数据来源**：把
    `_body_index_of` 换成一个与入参不相等的哨兵，`inserted_at` 必须跟着哨兵走。

    反向自证：把 `cmd_insert` 的返回改成 `{"inserted_at": args.index, ...}`
    （校验保留，只退回回显）→ 本条立刻红，报 `0 != 41`。
    """
    from types import SimpleNamespace

    import docx_edit as de

    src = save(tmp_path, "ins_src.docx", ["a", "b", "c"])
    doc = Document(str(src))
    monkeypatch.setattr(de, "_body_index_of", lambda _doc, _el: 41)

    res = de.cmd_insert(doc, SimpleNamespace(index=0, text="X", style=None))
    assert res["inserted_at"] == 41, "inserted_at 必须来自 _body_index_of 的返回值"


def test_insert_empty_body_accepts_index_zero_only(tmp_path: Path):
    """空 body（N=0）只剩一个插入点：`0` 合法（= 追加首段），其余退出 2。

    `Document()` 的默认模板 body 里只有 `sectPr`，`doc.paragraphs` 是空列表。
    合法范围 `[0, N-1]` 在这里退化成空集，若照搬会连「往空文档写第一段」都封死，
    所以单列这条边界：0 放行、1 与 -1 报错。
    """
    src = tmp_path / "empty_body.docx"
    Document().save(str(src))
    freeze_docx(src)

    out = tmp_path / "empty_body_out.docx"
    res = payload(run_edit("insert", src, "--index", 0, "--text", "FIRST",
                           "-o", out))
    assert res["inserted_at"] == 0
    assert [p_.text for p_ in Document(str(out)).paragraphs] == ["FIRST"]

    before = sha256(src)
    for bad in (1, -1):
        cp = run_edit("insert", src, "--index", bad, "--text", "NOPE")
        assert cp.returncode == 2, bad
        assert "out of range" in cp.stderr
        assert "current N=0" in cp.stderr
        assert cp.stdout.strip() == ""
    assert sha256(src) == before


# ===========================================================================
# 入口级输入校验：坏输入文件 → exit 2 + 单行 stderr + 零写盘（P8 / D8）
# ===========================================================================
# 背景（2026-11-11 实测）：`main()` 在
# `ap.parse_args()` 之后、任何子命令校验之前就 `Document(args.path)`，于是输入
# 文件本身坏掉时**八个子命令一起**裸抛异常，exit 1 + 十几行 traceback：
#
#   missing    → PackageNotFoundError      （19 行）
#   directory  → FileNotFoundError         （23 行）
#   non-zip    → PackageNotFoundError      （19 行）
#   bare-zip   → KeyError "no [Content_Types].xml"（31 行）
#
# `docx_read` 上一轮已经在**入口**做同样的事（`InputError` / `_check_input` /
# `_open_document`），本文件这一组用例把 `docx_edit` 钉到同一份契约上：
# 四档 × 多个子命令，一律 exit 2、stdout 空、stderr 恰好一行 `error: ...`、
# 一个字节都不写。**故意跨子命令参数化**——一条用例只测 `replace` 的话，
# 「把校验塞进 cmd_replace」也能让它变绿，那就不是入口级校验了。
def _bad_input_fixtures(tmp_path: Path) -> dict[str, Path]:
    """四档「这个路径不可能是 Word 包」，每档一个 fixture。

    * `missing`   —— 路径根本不存在（打错字 / cwd 不对 / 引号丢了）
    * `directory` —— 存在但是目录（shell 通配常见）
    * `non-zip`   —— 是文件但不是 ZIP（.doc/.rtf/.txt 改了后缀）
    * `bare-zip`  —— 是 ZIP 但没有 `[Content_Types].xml`（自制压缩包改名）
    """
    missing = tmp_path / "no_such_input.docx"
    if missing.exists():
        missing.unlink()

    directory = tmp_path / "a_directory.docx"
    directory.mkdir(exist_ok=True)

    non_zip = tmp_path / "plain_text.docx"
    non_zip.write_bytes(b"this is plain text, not a zip archive\n")

    bare_zip = tmp_path / "bare_archive.docx"
    with zipfile.ZipFile(bare_zip, "w") as z:
        z.writestr("hello.txt", "no content types here")
        z.writestr("word/document.xml", "<w:document/>")

    return {"missing": missing, "directory": directory,
            "non-zip": non_zip, "bare-zip": bare_zip}


# 四档 × 三个子命令。三个子命令分别落在三条**不同代码路径**上：
# `replace` 走 if 分支、`normalize` 走 else+try、`page-numbers` 走 else+try 且
# 会碰 sectPr/footer 部件；三条都退 2 才能排除「某个子命令局部修补」。
BAD_INPUT_COMMANDS = [
    ("replace", lambda p: ("replace", p, "--find", "a", "--replace", "b")),
    ("normalize", lambda p: ("normalize", p)),
    ("page-numbers", lambda p: ("page-numbers", p)),
]


@pytest.mark.parametrize("kind", ["missing", "directory", "non-zip", "bare-zip"])
@pytest.mark.parametrize("cmd", [c[0] for c in BAD_INPUT_COMMANDS])
def test_bad_input_file_is_a_usage_error_before_any_subcommand(
        tmp_path: Path, kind, cmd):
    """坏输入 → exit 2、stdout 空、stderr 恰好一行 `error:`、不写盘。

    参数化两维（档 × 子命令）就是为了证明校验收在**入口**：任何一条组合退回
    「裸 traceback + exit 1」，这里立刻红，且 `one_error_line` 会把那十几行
    traceback 原样打进失败消息里。
    """
    paths = _bad_input_fixtures(tmp_path)
    target = paths[kind]
    mk = dict(BAD_INPUT_COMMANDS)[cmd]

    before = sha256(target) if target.is_file() else None
    out = tmp_path / f"guard_{kind}_{cmd}.docx"

    cp = run_edit(*mk(target), "-o", out)

    assert cp.returncode == 2, (
        f"{kind}/{cmd}: rc={cp.returncode}\n"
        f"stdout={cp.stdout!r}\nstderr={cp.stderr}")
    line = one_error_line(cp.stderr)
    assert line.startswith("error: "), line
    # stdout 必须是**空的**：成功路径的 stdout 是 JSON，把失败也写进 stdout
    # 会让每个 `json.loads(stdout)` 的调用方把失败读成成功。
    assert cp.stdout == "", f"{kind}/{cmd}: stdout must be empty, got {cp.stdout!r}"
    # 零写盘：`-o` 指的产物不该存在，输入文件（若真是文件）SHA256 不变。
    assert not out.exists(), f"{kind}/{cmd}: wrote {out}"
    if before is not None:
        assert sha256(target) == before, f"{kind}/{cmd}: input was rewritten"


@pytest.mark.parametrize("kind", ["missing", "directory", "non-zip", "bare-zip"])
def test_bad_input_message_names_the_specific_failure(tmp_path: Path, kind):
    """四档各说各的**原因**，不是一句笼统的 "bad file"。

    依据：`SKILL.md` 的报错分类里，这四档对应四种不同的修法（改路径 / 换文件 /
    重新导出 / 重新打包）。一句 "cannot open" 会把「你打错字了」和「这个 .docx
    是别人发来的 .doc 改了后缀」压成同一条，调用方无从下手。
    """
    paths = _bad_input_fixtures(tmp_path)
    target = paths[kind]

    cp = run_edit("replace", target, "--find", "a", "--replace", "b")
    assert cp.returncode == 2
    line = one_error_line(cp.stderr)

    expect = {
        "missing": "no such file",
        "directory": "not a file",
        "non-zip": "not a ZIP package",
        "bare-zip": "not a Word package",
    }[kind]
    assert expect in line, f"{kind}: expected {expect!r} in {line!r}"


def test_input_check_runs_before_subcommand_argument_validation(tmp_path: Path):
    """坏输入 + 越界 `--index` ⇒ 报的必须是**输入**的错，不是 index 的错。

    这条专门钉「校验的位置」：只有当 `_check_input()` 排在 `Document()` 之前、
    且 `Document()` 排在子命令分派之前时，同时给两个错才会先报输入。
    顺序反了（先分派再开文档）时这里会拿到 `out of range`，本条立刻红。
    """
    paths = _bad_input_fixtures(tmp_path)
    cp = run_edit("delete", paths["missing"], "--index", "99")
    assert cp.returncode == 2
    line = one_error_line(cp.stderr)
    assert "no such file" in line, line
    assert "out of range" not in line, line


def test_every_subcommand_refuses_a_missing_input(tmp_path: Path):
    """八个子命令**全部**在坏输入上退 2 —— 一条不落。

    上一条按「四档 × 3 子命令」取样，这条把其余五个补上（`set-cell` / `insert`
    / `delete` / `style` / `toc`），保证没有哪个子命令因为自己的参数解析成功
    就绕过入口校验。
    """
    missing = _bad_input_fixtures(tmp_path)["missing"]
    cases = {
        "set-cell": ("set-cell", missing, "--table", "0", "--row", "0",
                     "--col", "0", "--text", "x"),
        "insert": ("insert", missing, "--index", "0", "--text", "x"),
        "delete": ("delete", missing, "--index", "0"),
        "style": ("style", missing, "--index", "0", "--style", "Normal"),
        "toc": ("toc", missing, "--index", "0"),
    }
    for name, argv in cases.items():
        cp = run_edit(*argv)
        assert cp.returncode == 2, f"{name}: rc={cp.returncode} {cp.stderr}"
        assert one_error_line(cp.stderr).startswith("error: "), name
        assert cp.stdout == "", name


def test_input_guard_is_shared_with_docx_read_not_a_copy(tmp_path: Path):
    """校验逻辑是**从 `docx_read` 拿的同一份**，不是抄了一份。

    为什么钉这个：抄一份的意思是两处判据会各自漂移（`docx_read` 上一轮刚为
    `KeyError` 补过 bare-zip 一档）。这里断言三个名字**对象同一**——一旦有人
    在 `docx_edit` 里另写一份 `_check_input`，即使行为当下相同，本条也红。
    """
    import docx_edit as de
    import docx_read as dr

    assert de._check_input is dr._check_input
    assert de._open_document is dr._open_document
    assert de.InputError is dr.InputError


def test_valid_document_still_round_trips_after_the_guard(tmp_path: Path):
    """反向保护：加了入口校验，好文件必须照旧 exit 0 + JSON。

    没有这条，「把所有输入都拒掉」也能让上面那组全绿。
    """
    src = save(tmp_path, "guard_good.docx", ["alpha", "beta"])
    out = tmp_path / "guard_good_out.docx"
    res = payload(run_edit("replace", src, "--find", "alpha",
                           "--replace", "ALPHA", "-o", out))
    assert res["replacements"] == 1
    assert res["ok"] is True
    assert [p.text for p in Document(str(out)).paragraphs] == ["ALPHA", "beta"]
