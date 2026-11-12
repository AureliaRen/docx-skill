# -*- coding: utf-8 -*-
"""表格几何 + ``w:zoom`` 补丁的回归（不需要真 docx 二进制）。

来源：``source-snapshot/scripts/office/test_table_layout.py``（原属上游 office
脚本目录，import 走 ``helpers.docx_utils``）。平铺进本仓库后只改 import 两处：
``helpers.docx_utils`` → ``docx_utils``，路径指向同级 ``../scripts``。
测试类与断言原样保留（unittest 类 pytest 能直接收）。

锁死 SKILL.md 承诺的两件事：

1. ``tblLayout=fixed`` 让渲染器**忽略 tcW**，所以宽度必须落进
   ``tblGrid/gridCol``；只设 ``cell.width`` 会被静默吃掉。
2. python-docx 写 ``<w:zoom w:val="bestFit"/>`` 缺 schema 强制的 ``w:percent``；
   而裸正则 ``r"<w:zoom\\s*/>"`` **抓不到**带属性的形式——最后一条用例专门
   把「朴素正则会失败」这件事固化成断言。
"""
import os
import subprocess
import sys
import unittest
import zipfile
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx_utils import (
    ET, wtag, insert_ordered, set_w_val, fix_child_order, cm_to_dxa,
    TBLPR_ORDER, TRPR_ORDER, TCPR_ORDER,
)
from fix_zoom import fix_settings_xml
import table_width as tw


class TestInsertOrdered(unittest.TestCase):
    def test_schema_order_respected(self):
        """tblLayout must land between tblW and tblLook, not at the end."""
        tblPr = ET.Element(wtag("tblPr"))
        for tag in ("tblLook", "tblW", "jc"):
            insert_ordered(tblPr, ET.Element(wtag(tag)))
        insert_ordered(tblPr, ET.Element(wtag("tblLayout")))
        got = [c.tag.split("}")[-1] for c in tblPr]
        self.assertEqual(got, ["tblW", "jc", "tblLayout", "tblLook"])

    def test_no_duplicate_element(self):
        """Inserting the same child twice must not duplicate it."""
        tblPr = ET.Element(wtag("tblPr"))
        insert_ordered(tblPr, ET.Element(wtag("tblW")))
        insert_ordered(tblPr, ET.Element(wtag("tblW")))
        self.assertEqual(len(tblPr.findall(wtag("tblW"))), 1)

    def test_existing_attrs_updated(self):
        """Re-inserting keeps one element but applies new attributes."""
        tblPr = ET.Element(wtag("tblPr"))
        first = ET.Element(wtag("tblW"))
        first.set(wtag("w"), "0")
        insert_ordered(tblPr, first)
        second = ET.Element(wtag("tblW"))
        second.set(wtag("w"), "8277")
        insert_ordered(tblPr, second)
        self.assertEqual(tblPr.find(wtag("tblW")).get(wtag("w")), "8277")
        self.assertEqual(len(tblPr.findall(wtag("tblW"))), 1)

    def test_fix_child_order_sorts_existing(self):
        """Hand-appended children get re-sorted into schema order."""
        tblPr = ET.Element(wtag("tblPr"))
        for tag in ("tblLook", "tblLayout", "tblW"):
            tblPr.append(ET.Element(wtag(tag)))
        moved = fix_child_order(tblPr, TBLPR_ORDER)
        self.assertGreater(moved, 0)
        got = [c.tag.split("}")[-1] for c in tblPr]
        self.assertEqual(got, ["tblW", "tblLayout", "tblLook"])

    def test_order_lists_match_xsd(self):
        """Spot-check the first/last entries against wml.xsd."""
        self.assertEqual(TBLPR_ORDER[0], "tblStyle")
        self.assertEqual(TBLPR_ORDER[6], "tblW")
        self.assertEqual(TBLPR_ORDER[12], "tblLayout")
        self.assertEqual(TRPR_ORDER[6], "cantSplit")
        self.assertEqual(TCPR_ORDER[1], "tcW")


class TestUnits(unittest.TestCase):
    def test_cm_to_dxa(self):
        self.assertEqual(cm_to_dxa(2.54), 1440)
        self.assertEqual(cm_to_dxa(1.0), 567)


class TestZoomPatch(unittest.TestCase):
    def test_attributed_tag_is_patched(self):
        """The form python-docx actually writes must gain w:percent."""
        xml = '<w:settings><w:zoom w:val="bestFit"/></w:settings>'
        out, n = fix_settings_xml(xml, 100)
        self.assertEqual(n, 1)
        self.assertIn('w:percent="100"', out)
        self.assertIn('w:val="bestFit"', out)   # keep the original attribute

    def test_bare_tag_is_patched(self):
        out, n = fix_settings_xml("<w:zoom/>", 100)
        self.assertEqual(n, 1)
        self.assertIn('w:percent="100"', out)

    def test_already_valid_is_untouched(self):
        xml = '<w:zoom w:val="bestFit" w:percent="100"/>'
        out, n = fix_settings_xml(xml, 100)
        self.assertEqual(n, 0)
        self.assertEqual(out, xml)

    def test_naive_regex_would_have_missed_it(self):
        """Documents the old SKILL.md bug: prove the naive pattern fails."""
        import re
        tag = '<w:zoom w:val="bestFit"/>'
        naive = re.sub(r"<w:zoom\s*/>", "<w:zoom w:percent=\"100\"/>", tag)
        self.assertEqual(naive, tag)            # unchanged == bug
        fixed, n = fix_settings_xml(tag, 100)
        self.assertEqual(n, 1)
        self.assertNotEqual(fixed, tag)


class TestTableWidths(unittest.TestCase):
    def _table(self, n=3, grid=2880, tcws=(1361, 2381, 2835)):
        tbl = ET.Element(wtag("tbl"))
        tblPr = ET.SubElement(tbl, wtag("tblPr"))
        layout = ET.SubElement(tblPr, wtag("tblLayout"))
        layout.set(wtag("type"), "fixed")
        grid_el = ET.SubElement(tbl, wtag("tblGrid"))
        for _ in range(n):
            gc = ET.SubElement(grid_el, wtag("gridCol"))
            gc.set(wtag("w"), str(grid))
        tr = ET.SubElement(tbl, wtag("tr"))
        for w in tcws:
            tc = ET.SubElement(tr, wtag("tc"))
            tcPr = ET.SubElement(tc, wtag("tcPr"))
            tcW = ET.SubElement(tcPr, wtag("tcW"))
            tcW.set(wtag("type"), "dxa")
            tcW.set(wtag("w"), str(w))
        return tbl

    def test_read_tcw_prefers_cell_widths(self):
        tbl = self._table()
        self.assertEqual(tw._read_tcw(tbl), [1361, 2381, 2835])
        self.assertEqual(tw._read_grid(tbl), [2880, 2880, 2880])

    def test_apply_widths_writes_all_three_places(self):
        """The whole point: grid, tblW and tcW must all carry the same values."""
        tbl = self._table()
        widths = [cm_to_dxa(2.4), cm_to_dxa(4.2), cm_to_dxa(5.0)]
        tw._apply_widths(tbl, widths)

        grid = [int(gc.get(wtag("w"))) for gc in tbl.iter(wtag("gridCol"))]
        self.assertEqual(grid, widths)                  # grid fixed

        tblPr = tbl.find(wtag("tblPr"))
        self.assertEqual(int(tblPr.find(wtag("tblW")).get(wtag("w"))), sum(widths))
        self.assertEqual(tblPr.find(wtag("tblLayout")).get(wtag("type")), "fixed")

        for tc, w in zip(tbl.iter(wtag("tc")), widths):
            self.assertEqual(int(tc.find(wtag("tcPr")).find(wtag("tcW")).get(wtag("w"))), w)

    def test_apply_widths_keeps_schema_order(self):
        tbl = self._table()
        tw._apply_widths(tbl, [1000, 1000, 1000])
        tblPr = tbl.find(wtag("tblPr"))
        got = [c.tag.split("}")[-1] for c in tblPr]
        self.assertEqual(got.index("tblW") < got.index("tblLayout"), True)


# ---------------------------------------------------------------------------
# 静默失败契约（P6 / 审计 §2.9 + §2.10）
#
# 旧实现有两种「说成功但其实没干成」：
#   1. ``word/document.xml`` 存在但不是 Word body 部件（解包错目录、拿错部件、
#      HTML 错误页存成 .docx）时，照样打印 ``no tables found`` 并 exit 0 ——
#      与「文档真没有表格」完全无法区分，上游解包错误被吃掉。
#   2. ``--widths`` 个数多于列数时静默截断（丢列）；少于列数时静默补末列
#      （凭空造宽度，且调用方毫不知情）。
#
# 三条契约：
#   * 解包/解析失败 → stderr + exit 2，stdout 绝不能出现 ``no tables found``；
#   * 文档合法且真无表格 → ``report`` 保持 ``no tables found`` + exit 0
#     （只读命令的空结果是真答案），但改文档的动作 exit 非零；
#   * ``--widths`` **多于**列数 → exit 2 且一个字节都不写（不许丢数据）；
#     **少于**列数 → 保留「补末列」这一既有承诺（lead 2026-11-11 裁定），
#     但 stderr 必须出现 ``padded N column(s)``；相等 → 输出与行为完全不变。
# ---------------------------------------------------------------------------
import contextlib
import io
import tempfile

_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
_DOC_NO_TBL = f'<w:document {_NS}><w:body><w:p/></w:body></w:document>'
_WRONG_PART = f'<w:styles {_NS}><w:style/></w:styles>'
_GARBAGE = "this is not xml at all\n"
_CELL = '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2000"/></w:tcPr><w:p/></w:tc>'
_TBL3 = ('<w:tbl><w:tblPr><w:tblW w:type="auto" w:w="0"/></w:tblPr>'
         '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/>'
         '<w:gridCol w:w="2000"/></w:tblGrid>'
         f'<w:tr>{"".join([_CELL] * 3)}</w:tr></w:tbl>')
_DOC_3COL = f'<w:document {_NS}><w:body>{_TBL3}</w:body></w:document>'


class TestTableWidthFailureContract(unittest.TestCase):
    """A failure must be loud, and an empty result must not look like success."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def _unpacked(self, xml, name="up"):
        d = self.tmp / name / "word"
        d.mkdir(parents=True)
        (d / "document.xml").write_text(xml, encoding="utf-8")
        return self.tmp / name

    def _run(self, argv):
        """Call main() the way the CLI does; return (exit_code, stdout, stderr)."""
        out, err = io.StringIO(), io.StringIO()
        argv_saved = sys.argv
        try:
            sys.argv = ["table_width.py", *argv]
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
                code = tw.main()
        finally:
            sys.argv = argv_saved
        return code, out.getvalue(), err.getvalue()

    # -- §2.9: unpack/parse failure must never masquerade as "no tables" ----

    def test_report_unparseable_document_exits_2(self):
        """Truncated unzip / HTML error page saved as .docx."""
        up = self._unpacked(_GARBAGE, "garbage")
        code, out, err = self._run(["report", str(up)])
        self.assertEqual(code, 2)
        self.assertNotIn("no tables found", out)
        self.assertIn("not well-formed XML", err)

    def test_report_wrong_part_exits_2_not_no_tables(self):
        """Parses fine but is not <w:document><w:body> — the masked-error case."""
        up = self._unpacked(_WRONG_PART, "wrongpart")
        code, out, err = self._run(["report", str(up)])
        self.assertEqual(code, 2)
        self.assertNotIn("no tables found", out)
        self.assertIn("not a Word document part", err)

    def test_report_genuinely_no_tables_still_exit_0(self):
        """The normal empty answer stays exactly as it was, and stays distinct."""
        up = self._unpacked(_DOC_NO_TBL, "notables")
        code, out, err = self._run(["report", str(up)])
        self.assertEqual(code, 0)
        self.assertEqual(out.strip(), "no tables found")
        self.assertEqual(err, "")

    def test_set_on_document_without_tables_is_not_success(self):
        """Nothing was written, so exit 0 would be a lie."""
        up = self._unpacked(_DOC_NO_TBL, "notables")
        code, out, err = self._run(["set", str(up), "--widths", "2.4,4.2"])
        self.assertNotEqual(code, 0)
        self.assertEqual(code, 1)
        self.assertNotIn("written:", out)
        self.assertIn("no tables found; nothing written", err)

    # -- §2.10: width count vs column count --------------------------------
    # The two mismatches are deliberately asymmetric (lead ruling 2026-11-11):
    # too many = silent data loss -> refuse; too few = documented "repeat the
    # last value" convenience -> keep, but announce it on stderr.

    def test_set_too_many_widths_exits_2_and_writes_nothing(self):
        up = self._unpacked(_DOC_3COL, "toomany")
        doc = up / "word" / "document.xml"
        before = doc.read_bytes()
        code, out, err = self._run(["set", str(up), "--widths", "1,2,3,4,5"])
        self.assertEqual(code, 2, f"{out}{err}")
        self.assertIn("refusing to drop", err)
        self.assertNotIn("written:", out)
        self.assertEqual(doc.read_bytes(), before, "a rejected call still wrote")

    def test_set_fewer_widths_pads_last_column_and_announces_it(self):
        """``列数不足补末列`` still works, but it is never silent."""
        up = self._unpacked(_DOC_3COL, "fewer")
        code, out, err = self._run(["set", str(up), "--widths", "1,2"])
        self.assertEqual(code, 0)
        self.assertIn("set 3 cols -> 1.00, 2.00, 2.00 cm", out)
        self.assertIn("padded 1 column(s)", err)
        self.assertIn("with 2.00 cm", err)
        root = ET.parse(str(up / "word" / "document.xml")).getroot()
        tbl = root.find(wtag("body")).find(wtag("tbl"))
        self.assertEqual(tw._read_grid(tbl),
                         [cm_to_dxa(1.0), cm_to_dxa(2.0), cm_to_dxa(2.0)])

    def test_set_width_count_match_is_unchanged(self):
        up = self._unpacked(_DOC_3COL, "tbl3")
        code, out, err = self._run(["set", str(up), "--widths", "1,2,3"])
        self.assertEqual(code, 0)
        self.assertIn("set 3 cols -> 1.00, 2.00, 3.00 cm", out)
        self.assertEqual(err, "", "an exact match must stay silent on stderr")
        root = ET.parse(str(up / "word" / "document.xml")).getroot()
        tbl = root.find(wtag("body")).find(wtag("tbl"))
        self.assertEqual(tw._read_grid(tbl),
                         [cm_to_dxa(1.0), cm_to_dxa(2.0), cm_to_dxa(3.0)])


# ===========================================================================
# 输入错误契约（P9-C）：`unpacked_dir` 不是「解包后的 docx」时退 2
#
# 本脚本的输入**不是 .docx 而是解包目录**，所以「缺文件 / 是目录 / 非 zip」
# 三档在这里平移成：
#
#   缺        → 路径根本不存在
#   是目录    → 传进来的是**文件**而不是目录（`<file>/word/document.xml` 必然不存在）
#   非 zip    → 目录在，但里面没有 `word/document.xml`（解包没解全 / 解包错了目录）
#
# 旧版三者一律 `error: <p>/word/document.xml not found` + **exit 1**，与
# 「文档有问题」混为一谈。本脚本对「解包目录对、但 XML 坏了」已经退 2
# （`ET.ParseError` / 非 `<w:document>` 根元素那两条），所以输入本身不对
# 更该退 2。
# ===========================================================================
def run_table_width(*args):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "table_width.py"), *map(str, args)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def _unpacked_docx(tmp_path: Path) -> Path:
    """造一个真的解包目录（从一份最小的 .docx 解出来）。"""
    from docx import Document

    src = tmp_path / "src.docx"
    d = Document()
    d.add_paragraph("hello")
    d.save(str(src))
    dst = tmp_path / "unpacked"
    dst.mkdir()
    with zipfile.ZipFile(src) as zf:
        zf.extractall(dst)
    return dst


def test_missing_unpacked_dir_is_a_usage_error(tmp_path: Path):
    """解包目录不存在 → exit 2 + 单行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such-unpacked"
    proc = run_table_width("report", missing)
    assert proc.returncode == 2, proc.stderr
    line = one_error_line(proc.stderr)
    assert "no such file" in line
    assert str(missing) in line
    assert proc.stdout == ""


def test_file_instead_of_unpacked_dir_is_a_usage_error(tmp_path: Path):
    """传进来的是文件而不是目录 → exit 2 + 单行 ``not a file``... 语义换成「不是目录」。"""
    from docx import Document

    f = tmp_path / "a-file.docx"
    d = Document()
    d.add_paragraph("hello")
    d.save(str(f))

    proc = run_table_width("report", f)
    assert proc.returncode == 2, proc.stderr
    line = one_error_line(proc.stderr)
    assert "not a directory" in line
    assert str(f) in line
    assert proc.stdout == ""


def test_dir_without_document_part_is_a_usage_error(tmp_path: Path):
    """目录在但缺 ``word/document.xml`` → exit 2 + 单行 ``not an unpacked .docx``。"""
    empty = tmp_path / "empty-unpack"
    (empty / "word").mkdir(parents=True)

    proc = run_table_width("report", empty)
    assert proc.returncode == 2, proc.stderr
    line = one_error_line(proc.stderr)
    assert "document.xml" in line
    assert str(empty) in line
    assert proc.stdout == ""


def test_valid_unpacked_dir_still_reports(tmp_path: Path):
    """合法解包目录 ``report`` 仍是 exit 0（空表读数是真答案，不是失败）。"""
    up = _unpacked_docx(tmp_path)

    proc = run_table_width("report", up)
    assert proc.returncode == 0, proc.stderr
    assert proc.stderr == ""
    assert proc.stdout.strip() == "no tables found"


if __name__ == "__main__":
    unittest.main(verbosity=2)
