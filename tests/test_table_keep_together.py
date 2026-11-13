# -*- coding: utf-8 -*-
"""``table`` keeps together on one page, and a table caption follows its table.

Word has **no table-level "keep with next"**.  The property that exists is
per row (``w:cantSplit``), and the one that glues consecutive blocks together is
per *paragraph* (``w:keepNext``).  So "this table must not break across a page"
is written in two layers, and asserting only one of them would be a false green:
every XML checker accepts ``<w:trPr><w:keepNext/></w:trPr>`` while Word silently
drops it, because ``keepNext`` is not among ``trPr``'s children (``TRPR_ORDER``
in ``docx_utils`` has no such entry).

These tests therefore assert the two layers **separately and in their own
places**, plus the row that must be exempt (the last one), plus the automatic
caption-to-table link and its counter-proof.
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest

SKILL = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL / "scripts"
CREATE = SCRIPTS / "docx_create.py"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

ROWS12 = [["行%d" % i, "值%d" % i, "备注%d" % i] for i in range(1, 13)]
A4 = {"width_mm": 210, "height_mm": 297,
      "margins_mm": {"top": 26, "bottom": 26, "left": 26, "right": 26}}


# ---- helpers -----------------------------------------------------------------

def run_create(tmp_path, spec):
    """Run the CLI; return (returncode, stdout payload, stderr, out path)."""
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.docx"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(CREATE), str(spec_path), str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SKILL))
    payload = None
    for line in (proc.stdout or "").splitlines():
        if line.strip().startswith("{"):
            payload = json.loads(line)
    return proc.returncode, payload, proc.stderr or "", out


def table_of(doc):
    assert doc.tables, "no table was built"
    return doc.tables[-1]


def trpr_tags(row):
    trPr = row._tr.find(qn("w:trPr"))
    return [] if trPr is None else [c.tag.split("}")[-1] for c in trPr]


def has_cant_split(row):
    trPr = row._tr.find(qn("w:trPr"))
    return trPr is not None and trPr.find(qn("w:cantSplit")) is not None


def row_para_keep_next(row):
    """How many of the row's cell paragraphs carry ``w:keepNext``."""
    count = 0
    for cell in row.cells:
        for para in cell.paragraphs:
            pPr = para._p.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:keepNext")) is not None:
                count += 1
    return count


def row_para_count(row):
    return sum(len(cell.paragraphs) for cell in row.cells)


# ---- 1. keep_together: both layers, in both places ---------------------------

def test_every_row_gets_cant_split(tmp_path):
    """The row-level half: all 13 rows (header included), none exempt."""
    rc, _, stderr, _ = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12,
         "keep_together": True}]})
    assert rc == 0, stderr
    doc = Document(str(tmp_path / "out.docx"))
    table = table_of(doc)
    assert len(table.rows) == 13, "12 data rows plus the header row"
    for i, row in enumerate(table.rows):
        assert "cantSplit" in trpr_tags(row), (
            "row %d has no w:cantSplit; trPr=%s" % (i, trpr_tags(row)))


def test_every_row_except_the_last_keeps_with_next(tmp_path):
    """The paragraph-level half: every cell paragraph, all rows but the last."""
    rc, _, stderr, _ = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12,
         "keep_together": True}]})
    assert rc == 0, stderr
    table = table_of(Document(str(tmp_path / "out.docx")))
    for i, row in enumerate(table.rows[:-1]):
        assert row_para_keep_next(row) == row_para_count(row), (
            "row %d: only %d of %d cell paragraphs keep_with_next"
            % (i, row_para_keep_next(row), row_para_count(row)))
    # The last row is the one that must NOT carry it: keepNext on the final row
    # would glue the table to whatever follows the table, which is not what
    # "keep the table together" means.
    assert row_para_keep_next(table.rows[-1]) == 0


def test_cant_split_is_inside_trpr_and_keep_next_is_inside_ppr(tmp_path):
    """Red line 1, as two separate assertions about two different parents.

    Both halves are checked against the *serialised* XML, not the object graph,
    because the object graph would happily accept a ``keepNext`` that Word then
    drops: the defect is in where the element lands, so the test has to look at
    where it landed.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12,
         "keep_together": True}]})
    assert rc == 0, stderr
    with zipfile.ZipFile(out) as zf:
        root = ElementTree.fromstring(zf.read("word/document.xml"))
    tbl = root.find(".//" + W + "tbl")
    assert tbl is not None
    tr = tbl.find(W + "tr")

    # (a) cantSplit lives in trPr — this is the one Word honours.
    trPr = tr.find(W + "trPr")
    assert trPr is not None
    assert trPr.find(W + "cantSplit") is not None
    assert tr.find(W + "keepNext") is None, (
        "keepNext must not be a direct child of w:tr")

    # (b) keepNext lives in the cell paragraph's pPr — a different parent.
    tc = tr.find(W + "tc")
    para = tc.find(W + "p")
    pPr = para.find(W + "pPr")
    assert pPr is not None
    assert pPr.find(W + "keepNext") is not None
    # (c) and it is *not* inside trPr: that is the silently-ignored spelling.
    assert trPr.find(W + "keepNext") is None, (
        "w:keepNext inside w:trPr is silently ignored by Word; it must go on "
        "the cell paragraphs instead")


def test_trpr_stays_in_schema_order(tmp_path):
    """``trPr`` is an xsd:sequence; cantSplit goes before trHeight/tblHeader."""
    rc, _, stderr, _ = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["a"], "rows": [["1"]],
         "keep_together": True}]})
    assert rc == 0, stderr
    table = table_of(Document(str(tmp_path / "out.docx")))
    from docx_utils import TRPR_ORDER
    tags = trpr_tags(table.rows[0])
    assert "cantSplit" in tags
    # Every present child must appear in TRPR_ORDER's relative order.
    positions = [TRPR_ORDER.index(t) for t in tags if t in TRPR_ORDER]
    assert positions == sorted(positions), tags


# ---- 2. backward compatibility ----------------------------------------------

def test_without_keep_together_no_cant_split_is_written(tmp_path):
    """Default False: a spec that does not ask gets the old document."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12}]})
    assert rc == 0, stderr
    table = table_of(Document(str(out)))
    for i, row in enumerate(table.rows):
        assert not has_cant_split(row), "row %d grew a cantSplit unasked" % i
        assert row_para_keep_next(row) == 0, "row %d grew keepNext unasked" % i


def test_keep_together_false_is_the_same_as_omitting_it(tmp_path):
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["a"], "rows": [["1"]],
         "keep_together": False}]})
    assert rc == 0, stderr
    with zipfile.ZipFile(out) as zf:
        xml = zf.read("word/document.xml").decode("utf-8")
    assert "cantSplit" not in xml


def test_keep_together_must_be_a_real_bool(tmp_path):
    """``"yes"`` would work by accident and ``"false"`` would too — reject both."""
    rc, payload, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table", "header": ["a"], "rows": [["1"]],
         "keep_together": "yes"}]})
    assert rc == 2, (rc, stderr)
    assert payload["ok"] is False
    assert "keep_together" in payload["error"]
    assert "true or false" in payload["error"]
    assert not out.exists()
    lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(lines) == 1, lines


# ---- 3. the automatic caption-to-table link ---------------------------------

def test_a_table_caption_directly_above_a_table_links_by_default(tmp_path):
    """No key asked for it: D-029 made the link the default."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "paragraph", "text": "下表列出十二行。"},
        {"type": "table_caption", "text": "十二行数据"},
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12},
    ]})
    assert rc == 0, stderr
    doc = Document(str(out))
    caption = doc.paragraphs[-1]        # the caption is the last body paragraph
    assert caption.text.strip(), "the caption paragraph should carry text"
    assert caption.paragraph_format.keep_with_next is True, (
        "the caption must stay with the paragraph that follows it")
    assert has_cant_split(doc.tables[-1].rows[0]), (
        "the table's first row must not split away from its caption")


def test_a_paragraph_between_caption_and_table_breaks_the_link(tmp_path):
    """The counter-proof: adjacency really is being tested, not assumed.

    If the implementation glued every caption to every following table these
    assertions fail — so this is the test that proves the "next block" rule is
    doing work rather than the implementation turning the feature on globally.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "paragraph", "text": "下表列出十二行。"},
        {"type": "table_caption", "text": "十二行数据"},
        {"type": "paragraph", "text": "（这里隔了一段）"},
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12},
    ]})
    assert rc == 0, stderr
    doc = Document(str(out))
    caption = doc.paragraphs[-2]        # the separating paragraph is now last
    assert caption.paragraph_format.keep_with_next in (False, None), (
        "a caption with a paragraph in between must not claim keep_with_next")
    assert not has_cant_split(doc.tables[-1].rows[0]), (
        "a table two blocks below a caption must not inherit the link")


def test_no_table_after_the_caption_means_no_link(tmp_path):
    """A trailing caption is legal and must not crash on a missing next block."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table_caption", "text": "孤零零的题注"},
    ]})
    assert rc == 0, stderr
    doc = Document(str(out))
    assert doc.paragraphs[-1].paragraph_format.keep_with_next in (False, None)


@pytest.mark.parametrize("blocks,expected", [
    ([{"type": "table_caption"}, {"type": "table"}], True),
    ([{"type": "table_caption"}, {"type": "paragraph"},
      {"type": "table"}], False),
    ([{"type": "table_caption"}], False),                     # nothing follows
    ([{"type": "table"}, {"type": "table_caption"}], False),   # wrong order
    ([{"type": "table_caption"}, {"type": "table_caption"}], False),
])
def test_the_adjacency_rule_itself(blocks, expected):
    """``_caption_glues_to_next_table`` is the single statement of "directly after"."""
    assert dc._caption_glues_to_next_table(blocks, 0) is expected


def test_the_adjacency_rule_handles_a_non_dict_neighbour():
    """A malformed neighbour is not a table, and must not raise."""
    assert dc._caption_glues_to_next_table(
        [{"type": "table_caption"}, "not a dict"], 0) is False
    assert dc._caption_glues_to_next_table([], 0) is False
    assert dc._caption_glues_to_next_table(
        [{"type": "table_caption"}, {"type": "table"}], 5) is False


def test_keep_together_and_the_caption_link_do_not_replace_each_other(tmp_path):
    """Both on: the caption links *and* every row is locked. Neither implies the other.

    ``table_caption`` + ``table`` gives one row worth of ``cantSplit``; asking
    for ``keep_together`` as well must give the whole table, and the caption
    link must still be there.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4, "blocks": [
        {"type": "table_caption", "text": "十二行数据"},
        {"type": "table", "header": ["名称", "数值", "说明"], "rows": ROWS12,
         "keep_together": True},
    ]})
    assert rc == 0, stderr
    doc = Document(str(out))
    assert doc.paragraphs[-1].paragraph_format.keep_with_next is True
    for i, row in enumerate(doc.tables[-1].rows):
        assert has_cant_split(row), "row %d lost its cantSplit" % i


# ---- 4. the helper is a helper, not a side effect ---------------------------

def test_keep_rows_together_is_idempotent():
    """Calling it twice must not duplicate the element (xsd:sequence allows one)."""
    doc = Document()
    table = doc.add_table(rows=0, cols=2)
    for i in range(3):
        cells = table.add_row().cells
        cells[0].text = "a%d" % i
        cells[1].text = "b%d" % i
    dc._keep_rows_together(table.rows)
    dc._keep_rows_together(table.rows)
    for row in table.rows:
        assert trpr_tags(row).count("cantSplit") == 1


def test_keep_rows_together_tolerates_an_empty_table():
    doc = Document()
    table = doc.add_table(rows=0, cols=2)
    dc._keep_rows_together(table.rows)       # must not raise
    assert list(table.rows) == []
