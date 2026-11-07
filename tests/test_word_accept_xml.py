# -*- coding: utf-8 -*-
"""Tests for ``tools/word_accept.py`` — the XML layer and the report contract.

Word cannot run on CI, so the tests pin down everything that can be decided from
the bytes: structure counts, bookmark names, field codes, the ``w:hyperlink``
wrapper that makes a cross-reference clickable, ``numPr`` + ``lvlText``, the
table-of-contents page numbers, style-id resolution, and the exact shape of the
merged report when Word is off.

Three properties of the layering are asserted directly, because they are the
reason the file exists:

1. importing the module must not import pywin32 — checked with ``ast``, so a
   well-meaning ``import win32com.client`` at module scope fails the suite
   instead of failing on a colleague's laptop;
2. a broken pywin32 must produce ``ok: null, skipped: true`` and exit 0, never a
   traceback;
3. the top-level contract keys stay populated when Word is off, falling back to
   the XML value with ``sources`` recording which layer answered.
"""
from __future__ import annotations

import ast
import json
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools"
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(TOOLS))
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import word_accept as wa  # noqa: E402

XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

# every key the Wave-2 callers are promised
CONTRACT_KEYS = {
    "ok", "openable", "skipped", "pages", "paragraphs", "sections", "bookmarks",
    "inline_shapes", "fields", "hyperlinks", "tables", "bookmark_names",
    "hyperlink_targets", "field_list", "numbering_render", "numbering_summary",
    "page_numbers", "page_numbers_source", "pdf", "pngs", "errors",
    "hard_errors", "warnings",
    "sources", "xref_summary", "xml", "word", "expectations",
}

MODULE_SRC = (TOOLS / "word_accept.py").read_text(encoding="utf-8")


# --------------------------------------------------------------------------- helpers
@pytest.fixture(autouse=True)
def _clean_word_cache():
    """No test may leak cached COM availability into the next one."""
    wa.reset_word_cache()
    yield
    wa.reset_word_cache()


def _run_element(paragraph, child, superscript=False):
    r = OxmlElement("w:r")
    if superscript:
        rpr = OxmlElement("w:rPr")
        va = OxmlElement("w:vertAlign")
        va.set(qn("w:val"), "superscript")
        rpr.append(va)
        r.append(rpr)
    r.append(child)
    paragraph._p.append(r)
    return r


def add_ref_hyperlink(paragraph, anchor, result="1", superscript=False):
    """Append a REF field wrapped in ``<w:hyperlink w:anchor=anchor>``.

    This is the BRIEF-COMMON §7.2 shape verbatim: a real wrapper around the
    whole field.  ``\\h`` alone would leave Hyperlinks.Count at 0.
    """
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("w:anchor"), anchor)
    hl.set(qn("w:history"), "1")
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(XMLSPACE, "preserve")
    instr.text = " REF " + anchor + " \\h "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = result
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for child in (begin, instr, sep, text, end):
        r = OxmlElement("w:r")
        r.append(child)
        hl.append(r)
    paragraph._p.append(hl)
    return hl


def add_style(styles_element, style_id, name):
    st = OxmlElement("w:style")
    st.set(qn("w:type"), "paragraph")
    st.set(qn("w:styleId"), style_id)
    nm = OxmlElement("w:name")
    nm.set(qn("w:val"), name)
    st.append(nm)
    styles_element.append(st)


def set_pstyle(paragraph, style_id):
    ppr = paragraph._p.get_or_add_pPr()
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), style_id)
    ppr.insert(0, ps)


def save(doc, tmp_path, name="t.docx"):
    path = tmp_path / name
    doc.save(str(path))
    return path


@pytest.fixture()
def plain(tmp_path):
    """A document with a table, so the row-mark offset is exercised."""
    doc = Document()
    doc.add_heading("Title", level=1)
    for i in range(3):
        doc.add_paragraph("段落 %d" % i)
    t = doc.add_table(rows=2, cols=2)
    for r in range(2):
        for c in range(2):
            t.cell(r, c).text = "c%d%d" % (r, c)
    return save(doc, tmp_path)


# --------------------------------------------------------------------------- layering
def test_no_toplevel_office_import():
    """The XML layer must stay importable without Office or pywin32.

    A module-level ``import win32com.client`` would break ``probe_xml`` — and
    with it every CI run — so it is banned by inspection rather than by hope.
    """
    tree = ast.parse(MODULE_SRC)
    banned = {"win32com", "win32com.client", "pythoncom", "fitz", "lxml"}
    found = []
    for node in tree.body:                       # top level only, deliberately
        if isinstance(node, ast.Import):
            found += [a.name for a in node.names if a.name in banned]
        elif isinstance(node, ast.ImportFrom):
            if node.module in banned:
                found.append(node.module)
    assert found == [], "top-level office imports: %s" % found


def test_probe_xml_plain_document(plain):
    rep = wa.probe_xml(plain)
    assert rep["xml_ok"] is True
    assert rep["xml_error"] is None
    assert rep["paragraphs_xml"] == 8            # 4 body + 4 table-cell
    assert rep["paragraphs_body_direct"] == 4
    assert rep["paragraphs_in_tables"] == 4      # 2x2 cells; row marks are Word's
    assert rep["table_rows"] == 2
    assert rep["tables_xml"] == 1
    assert rep["sections_xml"] == 1
    assert rep["parts"] > 0
    assert rep["xml_warnings"] == []


def test_probe_xml_counts_every_section(tmp_path):
    doc = Document()
    doc.add_paragraph("first")
    doc.add_section()
    doc.add_paragraph("second")
    rep = wa.probe_xml(save(doc, tmp_path))
    assert rep["sections_xml"] == 2
    assert rep["xml_ok"] is True


# --------------------------------------------------------------------------- failure modes
def test_probe_xml_missing_file(tmp_path):
    rep = wa.probe_xml(tmp_path / "nope.docx")
    assert rep["xml_ok"] is False
    assert "not readable" in rep["xml_error"]
    assert rep["paragraphs_xml"] == 0


def test_probe_xml_not_a_zip(tmp_path):
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"definitely not a docx")
    rep = wa.probe_xml(bad)
    assert rep["xml_ok"] is False
    assert "zip" in rep["xml_error"]


def test_probe_xml_zip_without_document_part(tmp_path):
    bad = tmp_path / "nodoc.docx"
    with zipfile.ZipFile(str(bad), "w") as z:
        z.writestr("hello.txt", "x")
    rep = wa.probe_xml(bad)
    assert rep["xml_ok"] is False
    assert "word/document.xml" in rep["xml_error"]


def test_probe_xml_malformed_document_xml(tmp_path):
    bad = tmp_path / "malformed.docx"
    with zipfile.ZipFile(str(bad), "w") as z:
        z.writestr("word/document.xml", "<w:document><w:body>")
    rep = wa.probe_xml(bad)
    assert rep["xml_ok"] is False
    assert "not well-formed" in rep["xml_error"]


def test_probe_xml_empty_file(tmp_path):
    empty = tmp_path / "empty.docx"
    empty.write_bytes(b"")
    rep = wa.probe_xml(empty)
    assert rep["xml_ok"] is False
    assert "empty" in rep["xml_error"]


# --------------------------------------------------------------------------- clickability
def test_hyperlink_wrapper_is_the_clickability_proxy(tmp_path):
    """§7.3: a real wrapper is what CI can see; ``\\h`` alone leaves a bare field.

    The bare-field document is the negative control — it is exactly the shape
    that produced the owner's "交叉引用没有生效" report.
    """
    doc = Document()
    p = doc.add_paragraph("引用：")
    add_ref_hyperlink(p, "图_1", result="1")           # wrapped
    p2 = doc.add_paragraph("裸域：")
    bare = OxmlElement("w:fldChar")
    bare.set(qn("w:fldCharType"), "begin")
    _run_element(p2, bare)
    instr = OxmlElement("w:instrText")
    instr.text = " REF 图_2 \\h "
    _run_element(p2, instr)
    rep = wa.probe_xml(save(doc, tmp_path))

    assert rep["fields_xml"] == 2                       # both fields counted
    assert rep["hyperlinks_xml"] == 1                   # only the wrapper
    assert rep["anchor_hyperlinks_xml"] == 1
    assert rep["hyperlink_anchors"] == [
        {"anchor": "图_1", "text": "1", "wrapper": True}]
    assert "图_2" not in [a["anchor"] for a in rep["hyperlink_anchors"]]


def test_field_inside_wrapper_is_found(tmp_path):
    doc = Document()
    add_ref_hyperlink(doc.add_paragraph("引用："), "图_1", result="3")
    rep = wa.probe_xml(save(doc, tmp_path))
    assert rep["fields_xml"] == 1
    assert rep["field_list_xml"] == [{"code": "REF 图_1 \\h", "result": "3"}]


def test_xref_summary_reports_clickable(tmp_path):
    doc = Document()
    add_ref_hyperlink(doc.add_paragraph("a"), "图_1", result="1")
    add_ref_hyperlink(doc.add_paragraph("b"), "图_1", result="1",
                      superscript=True)
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(save(doc, tmp_path)), "--word", "off"]))
    xs = rep["xref_summary"]
    assert xs["clickable"] is True
    assert xs["internal_anchors"] == 2
    assert xs["wrapper_count_xml"] == 2
    assert xs["ref_fields"] == 2
    # superscript is a Word-only measurement; without Word it stays unknown
    assert xs["superscript"] == 0
    assert xs["normal_size"] == 0


def test_external_hyperlink_counts_as_clickable(tmp_path):
    doc = Document()
    p = doc.add_paragraph("site")
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), "rId99")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "link"
    r.append(t)
    hl.append(r)
    p._p.append(hl)
    rep = wa.probe_xml(save(doc, tmp_path))
    assert rep["hyperlinks_xml"] == 1
    assert rep["external_hyperlinks_xml"] == 1
    assert rep["hyperlink_anchors"] == []


# --------------------------------------------------------------------------- text / styles
def test_tab_element_is_preserved(tmp_path):
    """``<w:tab/>`` is an element, not a character; dropping it loses the page
    number of every TOC entry."""
    import xml.etree.ElementTree as ET

    doc = Document()
    p = doc.add_paragraph()
    p.add_run("摘要").add_tab()
    p.add_run("i")
    path = save(doc, tmp_path)
    rep = wa.probe_xml(path)
    assert rep["paragraphs_xml"] == 1
    with zipfile.ZipFile(str(path)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    para = next(iter(root.iter(wa.q("p"))))
    assert wa._para_text(para) == "摘要\ti"


def test_toc_page_numbers_resolved_through_style_id(tmp_path):
    """The real document stores ``w:pStyle w:val="9"``; the style's *name* is
    ``toc 1``.  Matching on the raw value finds nothing — the same trap as the
    upstream 6.5 % heading recall."""
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    add_style(doc.styles.element, "27", "正文")
    for label, page in (("摘要", "i"), ("Abstract", "ii"), ("1.系统组成", "6")):
        p = doc.add_paragraph()
        set_pstyle(p, "9")
        p.add_run(label).add_tab()
        p.add_run(page)
    body = doc.add_paragraph("正文")
    set_pstyle(body, "27")
    rep = wa.probe_xml(save(doc, tmp_path))

    assert rep["page_numbers"] == ["i", "ii", "6"]
    assert rep["page_numbers_source"] == "toc"
    assert rep["toc_styles"] == ["toc 1"]


def test_style_usage_resolves_ids_to_names(tmp_path):
    doc = Document()
    add_style(doc.styles.element, "27", "正文")
    add_style(doc.styles.element, "24", "章标题")
    for _ in range(3):
        set_pstyle(doc.add_paragraph("x"), "27")
    set_pstyle(doc.add_paragraph("heading"), "24")
    doc.add_paragraph("no style at all")
    rep = wa.probe_xml(save(doc, tmp_path))
    usage = {u["name"]: u["count"] for u in rep["style_usage"]}
    assert usage == {"正文": 3, "章标题": 1, "(no pStyle)": 1}
    by_name = {u["name"]: u["style_id"] for u in rep["style_usage"]}
    assert by_name["正文"] == "27"
    assert by_name["(no pStyle)"] is None


def test_page_number_token_only_for_toc_styles(tmp_path):
    """A caption ending in a digit must not be mistaken for a TOC entry."""
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    add_style(doc.styles.element, "11", "table of figures")
    p = doc.add_paragraph()
    set_pstyle(p, "11")
    p.add_run("图 1 系统组成").add_tab()
    p.add_run("2")
    rep = wa.probe_xml(save(doc, tmp_path))
    assert rep["page_numbers"] == []
    assert rep["page_numbers_source"] == "none"


# --------------------------------------------------------------------------- bookmarks / numbering
def test_bookmark_names_exclude_internal(tmp_path):
    doc = Document()
    p = doc.add_paragraph("target")
    for bid, name in (("0", "图_1"), ("1", "_GoBack"), ("2", "ref_2")):
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bid)
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bid)
        p._p.append(start)
        p._p.append(end)
    rep = wa.probe_xml(save(doc, tmp_path))
    assert rep["bookmark_names_xml"] == ["图_1", "ref_2"]
    assert rep["bookmarks_xml"] == 3
    assert rep["bookmarks_internal_xml"] == 1


def test_numpr_paragraphs_carry_lvl_text(tmp_path):
    """``numPr`` detection must resolve numId -> abstractNum -> lvlText, so a
    consumer can assert the reference list really is ``[%1]``."""
    from docx_numbering import cmd_apply, cmd_define

    class Args:
        def __init__(self, **kw):
            self.__dict__.update(kw)

    doc = Document()
    for i in range(3):
        doc.add_paragraph("参考文献 %d" % i)
    nid = cmd_define(doc, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                               indent_left=425, hanging=425))["numId"]
    cmd_apply(doc, Args(index=["0:3"], style=None, match=None,
                        num_id=nid, ilvl=0))
    rep = wa.probe_xml(save(doc, tmp_path))

    assert len(rep["numpr_paragraphs"]) == 3
    first = rep["numpr_paragraphs"][0]
    assert first["index_xml"] == 1
    assert first["num_id"] == str(nid)
    assert first["ilvl"] == "0"
    assert first["lvl_text"] == "[%1]"
    assert first["text"].startswith("参考文献 0")
    assert rep["number_formats"] == {str(nid): {"0": "[%1]"}}
    assert rep["fields_xml"] == 0


def test_numpr_without_definition_still_reported(tmp_path):
    """A dangling numId is a real bug shape and must not vanish."""
    doc = Document()
    p = doc.add_paragraph("orphan")
    ppr = p._p.get_or_add_pPr()
    numpr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), "0")
    numid = OxmlElement("w:numId")
    numid.set(qn("w:val"), "77")
    numpr.append(ilvl)
    numpr.append(numid)
    ppr.append(numpr)
    rep = wa.probe_xml(save(doc, tmp_path))
    entry = rep["numpr_paragraphs"][0]
    assert entry["num_id"] == "77"
    assert entry["lvl_text"] is None            # nothing to resolve it against


# --------------------------------------------------------------------------- merge logic
def _xml_entry(idx, text, num_id="10"):
    return {"index_xml": idx, "text": text, "num_id": num_id, "ilvl": "0",
            "lvl_text": "[%1]"}


def _word_entry(idx, list_string, text):
    return {"index": idx, "list_string": list_string, "list_type": 3,
            "ilvl": 1, "text": text}


def test_merge_numbering_pairs_by_ordinal_not_by_index():
    """XML index and Word index legitimately differ (one extra paragraph per
    table row), so the pairing must be positional."""
    xml_entries = [_xml_entry(246, "盛万兴,段青"), _xml_entry(247, "Krausman")]
    word_lists = [_word_entry(272, "[1]", "盛万兴,段青"),
                  _word_entry(273, "[2]", "Krausman")]
    out, mism = wa._merge_numbering(xml_entries, word_lists, 200)
    assert mism == 0
    assert [e["list_string"] for e in out] == ["[1]", "[2]"]
    assert [e["index"] for e in out] == [272, 273]
    assert [e["index_xml"] for e in out] == [246, 247]
    assert all(e["rendered"] and e["text_match"] for e in out)
    assert all(e["lvl_text"] == "[%1]" for e in out)


def test_merge_numbering_flags_text_mismatch():
    xml_entries = [_xml_entry(10, "甲"), _xml_entry(11, "乙")]
    word_lists = [_word_entry(20, "[1]", "完全不相干"), _word_entry(21, "[2]", "乙")]
    out, mism = wa._merge_numbering(xml_entries, word_lists, 200)
    assert mism == 1
    assert [e["text_match"] for e in out] == [False, True]


def test_merge_numbering_keeps_unpaired_entries_visible():
    xml_entries = [_xml_entry(10, "甲"), _xml_entry(11, "乙")]
    word_lists = [_word_entry(20, "[1]", "甲")]
    out, mism = wa._merge_numbering(xml_entries, word_lists, 200)
    assert len(out) == 2
    assert mism == 0                     # the one pair that exists does match
    assert out[1]["rendered"] is False
    assert out[1]["list_string"] is None
    assert "no list paragraph" in out[1]["note"]

    out2, _ = wa._merge_numbering([_xml_entry(10, "甲")],
                                  [_word_entry(20, "[1]", "甲"),
                                   _word_entry(21, "[2]", "乙")], 200)
    assert out2[1]["num_id"] is None
    assert "no matching numPr" in out2[1]["note"]


def test_merge_numbering_respects_limit():
    xml_entries = [_xml_entry(i, "条目") for i in range(1, 11)]
    word_lists = [_word_entry(100 + i, "[%d]" % i, "条目") for i in range(1, 11)]
    out, _ = wa._merge_numbering(xml_entries, word_lists, 4)
    assert len(out) == 4
    assert [e["list_string"] for e in out] == ["[1]", "[2]", "[3]", "[4]"]


# --------------------------------------------------------------------------- no-Word contract
def test_word_off_report_keeps_every_contract_key(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off"]))
    assert CONTRACT_KEYS <= set(rep)
    assert rep["skipped"] is True
    assert rep["ok"] is None
    assert rep["openable"] is None
    assert rep["word"] is None
    assert rep["paragraphs"] == rep["xml"]["paragraphs_xml"]
    assert rep["sources"]["paragraphs"] == "xml"
    assert rep["sources"]["pages"] == "none"
    assert rep["pages"] is None
    assert rep["xml"]["xml_ok"] is True
    # numbering is still reported, with rendered=None meaning "unknown"
    assert rep["numbering_summary"]["rendered"] == 0


def test_import_failure_is_skipped_not_a_traceback(plain, monkeypatch):
    """The real "no Word here" path: pywin32 missing or Word not installed.

    The contract wants ``ok: null`` and exit 0 — the skip reason is reported in
    ``errors`` but must not be mistaken for a defect in the document.
    """
    def boom():
        raise ModuleNotFoundError("No module named 'win32com'")

    monkeypatch.setattr(wa, "_import_win32", boom)
    wa.reset_word_cache()
    rep = wa.build_report(wa.build_parser().parse_args([str(plain)]))
    assert rep["skipped"] is True
    assert rep["ok"] is None
    assert rep["openable"] is None
    assert rep["word_available"] is False
    assert rep["hard_errors"] == []
    assert any("unavailable" in e for e in rep["errors"])


def test_word_off_produces_no_com_error_even_when_pywin32_is_broken(
        plain, monkeypatch):
    def boom():
        raise RuntimeError("COM exploded")

    monkeypatch.setattr(wa, "_import_win32", boom)
    wa.reset_word_cache()
    rc = wa.main([str(plain), "--word", "off", "--quiet"])
    assert rc == 0


def test_broken_file_is_never_handed_to_word(plain, monkeypatch):
    """A structurally broken file must short-circuit, not open a repair dialog."""
    def forbidden():                     # pragma: no cover - must not run
        raise AssertionError("Word layer must not be reached")

    monkeypatch.setattr(wa, "probe_word", forbidden)
    bad = plain.parent / "broken.docx"
    bad.write_bytes(b"not a zip")
    rep = wa.build_report(wa.build_parser().parse_args([str(bad)]))
    assert rep["openable"] is False
    assert rep["ok"] is False
    assert rep["skipped"] is True
    assert any("not a zip" in e for e in rep["errors"])


# --------------------------------------------------------------------------- expectations
def test_expect_mismatch_fails_the_report(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off", "--expect-sections", "99"]))
    assert rep["ok"] is False
    assert rep["expectations"][0]["pass"] is False
    assert rep["expectations"][0]["actual"] == 1
    assert any("expect-sections" in e for e in rep["errors"])


def test_expect_paragraphs_alias(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off", "--expect-params",
         str(wa.probe_xml(plain)["paragraphs_xml"])]))
    assert rep["expectations"][0]["name"] == "paragraphs"
    assert rep["expectations"][0]["pass"] is True
    assert rep["ok"] is None            # still unverifiable, not a fake "true"


def test_expect_pages_without_word_is_a_warning_not_a_failure(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off", "--expect-pages", "3"]))
    assert rep["ok"] is None
    assert rep["expectations"][0]["pass"] is False
    assert any("not checkable without Word" in w for w in rep["warnings"])


def test_expect_openable_yes_without_word_stays_unknown(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off", "--expect-openable", "yes"]))
    assert rep["ok"] is None
    assert rep["expectations"][0]["actual"] is None


# --------------------------------------------------------------------------- CLI
def test_main_stdout_is_exactly_one_json_document(plain, capsys):
    rc = wa.main([str(plain), "--word", "off", "--quiet"])
    captured = capsys.readouterr()
    assert rc == 0
    assert captured.out.strip().startswith("{")
    assert captured.out.strip().endswith("}")
    report = json.loads(captured.out)
    assert report["tool"] == "word_accept"
    assert report["skipped"] is True
    assert captured.err == ""           # --quiet really is quiet


def test_summary_goes_to_stderr_not_stdout(plain, capsys):
    rc = wa.main([str(plain), "--word", "off"])
    captured = capsys.readouterr()
    assert rc == 0
    json.loads(captured.out)            # stdout is still pure JSON
    assert "[word_accept]" in captured.err


def test_main_returns_one_for_a_missing_file(tmp_path, capsys):
    rc = wa.main([str(tmp_path / "ghost.docx"), "--word", "off", "--quiet"])
    captured = capsys.readouterr()
    assert rc == 1
    report = json.loads(captured.out)
    assert report["ok"] is False
    assert report["errors"]


def test_json_output_file_is_utf8_without_bom(plain, tmp_path, capsys):
    out = tmp_path / "report.json"
    rc = wa.main([str(plain), "--word", "off", "--json", str(out), "--quiet"])
    assert rc == 0
    capsys.readouterr()
    raw = out.read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf")
    data = json.loads(raw.decode("utf-8"))
    assert data["docx"] == str(plain)


def test_limit_list_caps_every_list_key(tmp_path):
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    for i in range(6):
        p = doc.add_paragraph()
        set_pstyle(p, "9")
        p.add_run("条目%d" % i).add_tab()
        p.add_run(str(i + 1))
    path = save(doc, tmp_path)
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(path), "--word", "off", "--limit-list", "3"]))
    assert len(rep["page_numbers"]) == 3
    assert rep["page_numbers"] == ["1", "2", "3"]


def test_render_summary_never_raises(plain):
    """The summary is what a human reads after a bad run; it must not be the
    thing that crashes."""
    ok_rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off"]))
    assert "[word_accept]" in wa.render_summary(ok_rep)

    bad_rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off", "--expect-sections", "99"]))
    text = wa.render_summary(bad_rep)
    assert "ERROR" in text
    assert "ok=False" in text
