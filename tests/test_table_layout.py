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
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

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


if __name__ == "__main__":
    unittest.main(verbosity=2)
