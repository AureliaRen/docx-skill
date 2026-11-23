# -*- coding: utf-8 -*-
r"""``table_width.py check`` — 表格越界检测（只读）。

## 为什么要有这个 action

主人原话：「表格越界不自动检测 — 表 3 宽 19.05cm 超 16cm 版心，靠子代理手工算
gridCol 才发现。」

`report` 打印 `total = 10800 dxa (19.05 cm)` 却 **exit 0**：它不是漏了告警，而是
**根本没有「太宽」这个概念** —— 没人告诉它版心多宽。`check` 补的就是这半边。

## 本文件锁死什么

| # | 契约 | 为什么容易坏 |
|---|---|---|
| 1 | 19.05cm 表 + `--page-width 16` → `overflow: true` + **exit 1** | 「有发现」必须与非零退出码绑定，否则 CI 里 grep 不到 |
| 2 | 合规表 → exit 0 | 全红等于没红 |
| 3 | **不给 `--page-width` → `overflow: null` + exit 0** | `null` 不是 `false`：没判 ≠ 判合格 |
| 4 | `16cm` 与 `16` 结果相同 | 带单位是主人一定会踩的格式坑 |
| 5 | 非法 `--page-width` → exit 2 + stderr 恰一行 + stdout 空 | D-030 口径 |
| 6 | `--table` 越界 → exit 2 | 同上 |
| 7 | **autofit 表也判得出越界**（按 tcW 和算） | 只按 tblGrid 算会放过 autofit 文档 |

## 反例（本文件第二条隐藏轴）

第 4 条的正向是「`16cm` 能用」，它的反面是「`--page-width` 不能被当成
`--content-width` 的默认值 14.6」。所以有一条用例专门断言：同一张 15cm 的表，
对 `--page-width 16` 合规、对 `--page-width 14` 越界 —— 阈值真的在起作用，
而不是恒真/恒假。
"""
from __future__ import annotations

import ast
import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import table_width as tw  # noqa: E402

SCRIPT = SCRIPTS / "table_width.py"
MODULE_SRC = SCRIPT.read_text(encoding="utf-8")

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


# --------------------------------------------------------------------- helpers
def run_cli(argv):
    """Run the CLI in a subprocess: exit codes and separate streams must be real."""
    return subprocess.run([sys.executable, str(SCRIPT), *argv],
                          capture_output=True, text=True,
                          encoding="utf-8", errors="replace")


def unpacked(tmp_path: Path, xml: str, name: str = "up") -> Path:
    """Write a minimal unpacked body part: ``<name>/word/document.xml``."""
    d = tmp_path / name / "word"
    d.mkdir(parents=True)
    (d / "document.xml").write_text(xml, encoding="utf-8")
    return tmp_path / name


def cell(w: int) -> str:
    return ('<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="%d"/></w:tcPr><w:p/></w:tc>'
            % w)


def table_xml(widths, layout="fixed", tbl_w="auto", grid=None) -> str:
    """A one-row table with explicit ``tcW`` per column and an optional grid.

    ``grid`` defaults to the same values as ``widths`` (a self-consistent
    table); pass a different list to model the stale-equal-split trap.
    """
    grid = list(grid) if grid is not None else list(widths)
    grid_xml = "".join('<w:gridCol w:w="%d"/>' % w for w in grid)
    layout_xml = ('<w:tblLayout w:type="%s"/>' % layout) if layout else ""
    return ('<w:tbl><w:tblPr><w:tblW w:type="%s" w:w="0"/>%s</w:tblPr>'
            '<w:tblGrid>%s</w:tblGrid><w:tr>%s</w:tr></w:tbl>'
            % (tbl_w, layout_xml, grid_xml,
               "".join(cell(w) for w in widths)))


def doc_with(*tables) -> str:
    return ('<w:document %s><w:body>%s</w:body></w:document>'
            % (_NS, "".join(tables)))


WIDE_3COL = 10800           # 3 x 3600 dxa = 19.05 cm — the owner's actual case
COMPLIANT_3COL = 8505       # 3 x 2835 dxa = 15.00 cm


@pytest.fixture()
def wide(tmp_path):
    """19.05 cm table, fixed layout, grid consistent with tcW."""
    return unpacked(tmp_path, doc_with(table_xml([3600, 3600, 3600])), "wide")


@pytest.fixture()
def ok(tmp_path):
    """15.00 cm table — compliant against a 16 cm text block."""
    return unpacked(tmp_path, doc_with(table_xml([2835, 2835, 2835])), "ok")


def read_json(proc):
    return json.loads(proc.stdout)


def stderr_lines(proc):
    return [ln for ln in proc.stderr.replace("\r", "").split("\n") if ln]


# ---------------------------------------------------------------- 1. fires
def test_overwide_table_is_flagged_and_exits_1(wide):
    """★ 19.05 cm 的三列表 + ``--page-width 16`` → overflow + exit 1.

    这是主人实际遇到的那张表。它必须**既**在 JSON 里说 true，**又**用非零
    退出码让 CI 能 grep —— 只有其中一个都算没交付。
    """
    proc = run_cli(["check", str(wide), "--page-width", "16", "--json"])
    assert proc.returncode == 1, proc.stderr
    report = read_json(proc)
    assert report["overflow"] == [1]
    row = report["tables"][0]
    assert row["overflow"] is True
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)
    assert row["page_width_cm"] == pytest.approx(16.0)
    assert row["overflow_by_cm"] == pytest.approx(3.05, abs=0.01)
    assert row["cols"] == 3
    assert row["layout"] == "fixed"


def test_human_output_names_the_overflow(wide):
    """The default text output must state the overrun in cm, not just "bad"."""
    proc = run_cli(["check", str(wide), "--page-width", "16"])
    assert proc.returncode == 1
    assert "overflow" in proc.stdout
    assert "19.05" in proc.stdout
    assert "+3.05" in proc.stdout


def test_threshold_actually_moves(ok):
    """★ 同一张 15 cm 表：对 16 合规，对 14 越界 —— 阈值不是恒真恒假。

    没有这条，``overflow`` 写死 ``True`` 也能让上面两条变绿。
    """
    passed = run_cli(["check", str(ok), "--page-width", "16", "--json"])
    assert passed.returncode == 0, passed.stderr
    assert read_json(passed)["tables"][0]["overflow"] is False

    failed = run_cli(["check", str(ok), "--page-width", "14", "--json"])
    assert failed.returncode == 1
    assert read_json(failed)["tables"][0]["overflow"] is True


# -------------------------------------------------------------- 2. quiet
def test_compliant_table_exits_0(ok):
    """A table that fits must stay silent about overflow and exit 0."""
    proc = run_cli(["check", str(ok), "--page-width", "16", "--json"])
    assert proc.returncode == 0, proc.stderr
    report = read_json(proc)
    assert report["overflow"] == []
    assert report["tables"][0]["overflow"] is False
    assert report["tables"][0]["total_cm"] == pytest.approx(15.0, abs=0.01)


def test_no_tables_exits_0(tmp_path):
    """A document with no tables is a real empty answer, like ``report``."""
    up = unpacked(tmp_path, doc_with(), "empty")
    proc = run_cli(["check", str(up), "--page-width", "16", "--json"])
    assert proc.returncode == 0, proc.stderr
    assert read_json(proc)["tables"] == []


# --------------------------------------------- 3. no --page-width -> null
def test_missing_page_width_is_null_not_false(wide):
    """★ 不给 ``--page-width`` → ``overflow: null`` + exit 0.

    ``null`` 与 ``false`` 的差别是本条的全部意义：``false`` 会让调用方以为
    「量过了，没超」。而且**绝不能**拿 ``auto`` 的 ``--content-width`` 默认
    值 14.6 顶上来 —— 那张 19.05 cm 的表会在 14.6 下报越界，一个没给版心的
    调用方就凭空得到一个「结论」。
    """
    proc = run_cli(["check", str(wide), "--json"])
    assert proc.returncode == 0, proc.stderr
    report = read_json(proc)
    row = report["tables"][0]
    assert row["overflow"] is None
    assert row["overflow_by_cm"] is None
    assert row["page_width_cm"] is None
    assert report["page_width_given"] is False
    # The table's own widths are still reported — that reading needs no page.
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)


def test_missing_page_width_says_so_in_text(wide):
    """The human report must name the missing input instead of staying quiet."""
    proc = run_cli(["check", str(wide)])
    assert proc.returncode == 0
    assert "not provided" in proc.stdout
    assert "cannot be judged" in proc.stdout


# ------------------------------------------------------- 4. unit suffix
@pytest.mark.parametrize("spelling", ["16", "16cm", "16CM", "16 cm", "160mm"])
def test_cm_suffix_and_plain_number_agree(wide, spelling):
    """★ ``16`` / ``16cm`` / ``16 cm`` / ``160mm`` 必须是同一个 16 cm。

    带单位是主人一定会踩的格式坑：argparse 的 ``type=float`` 会直接
    "invalid value" 并甩一整屏 usage，而不是给出这个脚本承诺的单行 error。
    """
    proc = run_cli(["check", str(wide), "--page-width", spelling, "--json"])
    assert proc.returncode == 1, (spelling, proc.stderr)
    row = read_json(proc)["tables"][0]
    assert row["page_width_cm"] == pytest.approx(16.0)
    assert row["overflow_by_cm"] == pytest.approx(3.05, abs=0.01)


def test_inch_and_point_spellings_parse():
    """The parser understands the other units too — 16 cm is 6.299 in / 453.5 pt."""
    for text, expect in (("6.299in", 16.0), ("453.54pt", 16.0), ("16", 16.0)):
        cm, why = tw._parse_page_width(text)
        assert why is None, (text, why)
        assert cm == pytest.approx(expect, abs=0.01)


# --------------------------------------------- 5. bad --page-width -> exit 2
@pytest.mark.parametrize("bad", ["abc", "-5", "0", "16x", "", "nan", "inf"])
def test_bad_page_width_is_exit_2_one_line(wide, bad):
    """★ D-030：参数错 → exit 2、stdout 空、stderr **恰一行** ``error:``。"""
    proc = run_cli(["check", str(wide), "--page-width", bad])
    assert proc.returncode == 2, (bad, proc.stdout, proc.stderr)
    assert proc.stdout == "", "stdout must stay empty on a refusal"
    lines = stderr_lines(proc)
    assert len(lines) == 1, lines
    assert lines[0].startswith("error:"), lines[0]


def test_bad_page_width_writes_nothing(wide):
    """A refused argument must not touch the XML part — ``check`` is read-only."""
    doc = wide / "word" / "document.xml"
    before = doc.read_bytes()
    proc = run_cli(["check", str(wide), "--page-width", "abc"])
    assert proc.returncode == 2
    assert doc.read_bytes() == before


# --------------------------------------------------------- 6. --table range
@pytest.mark.parametrize("index", ["0", "2", "9", "-1"])
def test_table_index_out_of_range_is_exit_2(tmp_path, index):
    """★ ``--table`` 越界 → exit 2（不是 1：参数指错，什么都没读）。"""
    up = unpacked(tmp_path, doc_with(table_xml([2835, 2835, 2835])), "one")
    proc = run_cli(["check", str(up), "--page-width", "16", "--table", index])
    assert proc.returncode == 2, (index, proc.stdout, proc.stderr)
    assert proc.stdout == ""
    lines = stderr_lines(proc)
    assert len(lines) == 1, lines
    assert "out of range" in lines[0]


def test_table_index_selects_one_table(tmp_path):
    """``--table 2`` reports table 2 (and only table 2), keeping its real index."""
    xml = doc_with(table_xml([2835, 2835, 2835]),
                   table_xml([3600, 3600, 3600]))
    up = unpacked(tmp_path, xml, "two")
    proc = run_cli(["check", str(up), "--page-width", "16", "--table", "2",
                    "--json"])
    assert proc.returncode == 1, proc.stderr
    report = read_json(proc)
    assert len(report["tables"]) == 1
    assert report["tables"][0]["table"] == 2
    assert report["overflow"] == [2]


# ------------------------------------------------------------ 7. autofit
def test_autofit_table_is_judged_from_tcw(tmp_path):
    """★ autofit（无 ``tblLayout``）时按 tcW 和算——grid 可能还是陈旧等分。

    Word 在 autofit 下按 tcW 比例分配，此时 tblGrid 不是权威。若 ``check``
    一律读 grid，这张表会读到 grid 的 15 cm 而报「合规」，而它实际渲染
    19.05 cm。这条就是防那个假阴性。
    """
    # grid says 15 cm (stale), tcW says 19.05 cm (what renders)
    xml = doc_with(table_xml([3600, 3600, 3600], layout=None,
                             grid=[2835, 2835, 2835]))
    up = unpacked(tmp_path, xml, "autofit")
    proc = run_cli(["check", str(up), "--page-width", "16", "--json"])
    assert proc.returncode == 1, proc.stderr
    row = read_json(proc)["tables"][0]
    assert row["layout"] == "autofit"
    assert row["source"] == "tcW"
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)
    assert row["overflow"] is True


def test_fixed_layout_with_stated_tblw_reads_tblw(tmp_path):
    """★ fixed + **显式** tblW 带宽 → tblW 权威，grid 与 tcW 都被它覆盖。

    这条在 P18 审查时被 Word 实测改写过一次，原始版本断言的是「fixed 下
    tblGrid 权威」，而那个规则是错的。实测（Word 16.0，``Table.Columns(i)``）：

        tblW dxa 15.00, grid 15.00, tcW 19.05  →  Word 15.00   (tcW 被忽略)
        tblW dxa 15.00, grid 45.00, tcW 45.00  →  Word 15.00   (grid 也被忽略)
        tblW dxa 19.05, grid 15.00, tcW 15.00  →  Word 19.05

    所以显式 tblW 存在时，读 grid 或 tcW 都会给出与 Word 不符的结论。
    ``table_xml`` 的 ``tbl_w`` 参数只写 ``w:type``，宽度固定为 ``w:w="0"``，
    因此这里显式构造一个真正带宽度值的 tblW。
    """
    xml = doc_with(table_xml([2835, 2835, 2835], layout="fixed",
                             grid=[3600, 3600, 3600]))
    xml = xml.replace('<w:tblW w:type="auto" w:w="0"/>',
                      '<w:tblW w:type="dxa" w:w="10800"/>')
    up = unpacked(tmp_path, xml, "fixedtblw")
    proc = run_cli(["check", str(up), "--page-width", "16", "--json"])
    assert proc.returncode == 1, proc.stderr
    row = read_json(proc)["tables"][0]
    assert row["source"] == "tblW"
    # 10800 dxa = 19.05 cm: tblW wins over the 19.05 grid and the 15 cm tcW.
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)
    assert row["overflow"] is True


def test_fixed_layout_with_auto_tblw_reads_tcw(tmp_path):
    """★ fixed + ``tblW=auto`` → **tcW** 权威（Word 实测 19.05，grid 只有 15）。

    这是 P17 那条假阴性的正身：``tblW=auto`` 是 python-docx 的默认，也是本仓库
    ``docx_create`` 生成表格的形态，而旧代码在 fixed 下只读 grid，于是把一张
    Word 渲染 19.05 cm 的表报成「15 cm 合规、exit 0」。
    """
    xml = doc_with(table_xml([3600, 3600, 3600], layout="fixed",
                             tbl_w="auto", grid=[2835, 2835, 2835]))
    up = unpacked(tmp_path, xml, "fixedauto")
    proc = run_cli(["check", str(up), "--page-width", "16", "--json"])
    assert proc.returncode == 1, proc.stderr
    row = read_json(proc)["tables"][0]
    assert row["source"] == "tcW"
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)
    assert row["overflow"] is True


# ------------------------------------------------------- the required hint
def test_fixed_layout_hint_is_in_text_and_json(wide):
    """★ 主人点名要求的那句话，两个输出通道都要有。

    ``tblLayout=fixed 时只认 tblGrid，改 tcW 无效`` —— 少了它，``check`` 只会
    说「超了」，不会说「所以别去改 tcW」。
    """
    text = run_cli(["check", str(wide), "--page-width", "16"])
    assert "只认 tblGrid" in text.stdout
    assert "改 tcW 无效" in text.stdout

    js = run_cli(["check", str(wide), "--page-width", "16", "--json"])
    hint = read_json(js)["hint"]
    assert "只认 tblGrid" in hint
    assert "改 tcW 无效" in hint


# ------------------------------------------------------------- read-only
def test_check_never_writes(wide):
    """``check`` 是只读 action：跑一遍 XML 必须逐字节不变。"""
    doc = wide / "word" / "document.xml"
    before = doc.read_bytes()
    for argv in (["--page-width", "16"], ["--page-width", "16", "--json"],
                 []):
        run_cli(["check", str(wide), *argv])
    assert doc.read_bytes() == before


def test_check_refuses_non_unpacked_dirs(tmp_path):
    """``check`` 沿用既有的输入三档：目录不在 / 是文件 / 缺 body part → 2."""
    missing = tmp_path / "absent"
    proc = run_cli(["check", str(missing), "--page-width", "16"])
    assert proc.returncode == 2, proc.stderr
    assert proc.stdout == ""

    a_file = tmp_path / "a.docx"
    a_file.write_text("not a zip", encoding="utf-8")
    proc = run_cli(["check", str(a_file), "--page-width", "16"])
    assert proc.returncode == 2
    assert proc.stdout == ""


def test_check_on_a_real_docx_roundtrip(tmp_path):
    """End-to-end: build a real .docx, unpack it, widen it, catch it.

    The synthetic XML above proves the arithmetic; this proves the action works
    against a document python-docx actually produced (tblStyle, tblLook,
    ``tblW type=auto`` and all), which is the shape real files arrive in.
    """
    docx = pytest.importorskip("docx")
    from docx.shared import Mm

    doc = docx.Document()
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.left_margin = section.right_margin = Mm(25)
    t = doc.add_table(rows=2, cols=3)
    for i in range(3):
        t.cell(0, i).text = "c%d" % (i + 1)
    src = tmp_path / "src.docx"
    doc.save(str(src))

    dest = tmp_path / "unpacked"
    with zipfile.ZipFile(src) as zf:
        zf.extractall(dest)

    # 15 cm first: compliant
    set_ok = run_cli(["set", str(dest), "--all", "--widths", "5,5,5"])
    assert set_ok.returncode == 0, set_ok.stderr
    proc = run_cli(["check", str(dest), "--page-width", "16", "--json"])
    assert proc.returncode == 0, proc.stdout
    assert read_json(proc)["tables"][0]["overflow"] is False

    # now 19.05 cm: the real defect
    set_wide = run_cli(["set", str(dest), "--all", "--widths", "6.35,6.35,6.35"])
    assert set_wide.returncode == 0, set_wide.stderr
    proc = run_cli(["check", str(dest), "--page-width", "16", "--json"])
    assert proc.returncode == 1, proc.stdout
    row = read_json(proc)["tables"][0]
    assert row["overflow"] is True
    assert row["total_cm"] == pytest.approx(19.05, abs=0.01)


# ---------------------------------------------------------------- layering
def test_module_still_bootstraps_first():
    """``scripts/*.py`` 的**第一条可执行语句**必须是 ``import _bootstrap``。

    按 AST 判而不是按行号判：本模块以 docstring 开头（``table_width.py`` 没有
    shebang），``docx_preview.py`` 又先有 ``from __future__ import annotations``
    —— 那条 ``__future__`` 必须是**第一条语句**才能生效，所以「第一行」这条
    硬规矩的真实内容只能是「除 ``__future__`` 外第一条 import 是 _bootstrap」。
    """
    tree = ast.parse(MODULE_SRC)
    stmts = list(tree.body)
    # A docstring is a bare string expression at index 0; `__future__` must
    # precede everything executable.  Skip both, then the bootstrap import is
    # required to be the first thing left.
    if stmts and isinstance(stmts[0], ast.Expr) \
            and isinstance(stmts[0].value, ast.Constant) \
            and isinstance(stmts[0].value.value, str):
        stmts = stmts[1:]
    if stmts and isinstance(stmts[0], ast.ImportFrom) \
            and stmts[0].module == "__future__":
        stmts = stmts[1:]
    assert stmts, "module has no statements"
    first = stmts[0]
    assert isinstance(first, ast.Import), ast.dump(first)[:120]
    assert [a.name for a in first.names] == ["_bootstrap"], \
        [a.name for a in first.names]


def test_parses_as_python_310():
    """CI 有 py3.10 腿：新代码不许用 3.11+ 语法。"""
    ast.parse(MODULE_SRC, feature_version=(3, 10))


def test_line_endings_are_lf():
    """``scripts/*.py`` 一律 LF（``.gitattributes`` 承诺）。"""
    raw = SCRIPT.read_bytes()
    assert b"\r\n" not in raw, "CRLF found in scripts/table_width.py"
