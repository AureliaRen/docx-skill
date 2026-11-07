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
    # page-number truth layer (T3)
    "page_number_formats", "page_number_restarts", "section_specs",
    "paragraph_pages", "paragraph_pages_source", "toc_audit",
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


# ===========================================================================
# page-number truth layer (T3)
# ===========================================================================
# The distinction these tests exist for, measured on
# ``work/subjects/math/paper.docx`` (a 5-section paper whose last section is
# ``lowerRoman start=1``):
#
#   physical page 7  ->  WD_ACTIVE_END_PAGE          = 7
#   the same page    ->  WD_ACTIVE_END_ADJUSTED_PAGE = 2   -> glyph "ii"
#   last section     ->  physical 8 / adjusted 1     -> glyph "i"
#   the PDF's footer band for page 8 reads "- i -"
#
# So "page 7" and "i" are both true statements about *different* pages, and the
# only way to keep them apart is to record which measurement produced which
# number.  Everything below pins down one half of that.

def set_pgnum(section, fmt=None, start=None):
    """Write ``<w:pgNumType>`` on a section's ``sectPr``, *replacing* any old one.

    Replacement rather than append is the point: ``Document.add_section()``
    clones the running ``sectPr``, so appending leaves two ``pgNumType``
    elements and Word silently obeys only the last (measured on Word 16: a
    section holding ``lowerRoman start=1`` then ``upperRoman start=5`` reads
    ``V``).  The duplicate case is exercised separately, on purpose.
    """
    for old in section._sectPr.findall(qn("w:pgNumType")):
        section._sectPr.remove(old)
    el = OxmlElement("w:pgNumType")
    if fmt is not None:
        el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), start)
    section._sectPr.append(el)
    return el


def test_wd_active_end_adjusted_page_constant():
    """1 and 3 are different questions with the same name in passing."""
    assert wa.WD_ACTIVE_END_ADJUSTED_PAGE == 1
    assert wa.WD_ACTIVE_END_PAGE == 3
    assert wa.WD_ACTIVE_END_ADJUSTED_PAGE != wa.WD_ACTIVE_END_PAGE


def test_section_specs_reads_pgnumtype(tmp_path):
    """Each section's raw fmt/start, in document order, straight from the XML."""
    doc = Document()
    doc.add_paragraph("cover")
    s1 = doc.sections[0]
    doc.add_section()
    s2 = doc.sections[1]
    set_pgnum(s2, "lowerRoman", "1")
    doc.add_paragraph("front matter")
    doc.add_section()
    s3 = doc.sections[2]
    set_pgnum(s3, "decimal")            # fmt only: continues, does not restart
    doc.add_paragraph("body")

    rep = wa.probe_xml(save(doc, tmp_path))
    specs = rep["section_specs"]
    assert rep["sections_xml"] == 3
    assert [s["index"] for s in specs] == [1, 2, 3]
    assert [(s["fmt"], s["start"]) for s in specs] == [
        (None, None), ("lowerRoman", "1"), ("decimal", None)]
    # ``start`` stays the raw string so "absent" and "explicit 1" differ
    assert specs[1]["start"] == "1"
    assert specs[2]["start"] is None
    assert rep["page_number_formats"] == ["lowerRoman", "decimal"]
    # only an explicit start is a restart
    assert rep["page_number_restarts"] == [
        {"section": 2, "start": "1", "fmt": "lowerRoman"}]


def test_section_specs_carry_paragraph_spans(tmp_path):
    """``first_para_index`` / ``last_para_index`` make 段落→节 answerable.

    A ``sectPr`` lives on the last paragraph of the section it closes — for
    ``Document.add_section()`` that is the empty paragraph it inserts, which is
    still body text of that section.  So spans are inclusive at both ends and
    must tile the body without gaps or overlaps.
    """
    doc = Document()
    for i in range(3):
        doc.add_paragraph("cover %d" % i)      # 0,1,2
    doc.add_section()                          # empty paragraph 3 carries sectPr
    for i in range(2):
        doc.add_paragraph("body %d" % i)       # 4,5
    doc.add_section()                          # empty paragraph 6 carries sectPr
    doc.add_paragraph("appendix")              # 7

    rep = wa.probe_xml(save(doc, tmp_path))
    specs = rep["section_specs"]
    spans = [(s["first_para_index"], s["last_para_index"]) for s in specs]
    assert spans == [(0, 3), (4, 6), (7, 7)]
    assert rep["paragraphs_body_direct"] == 8
    # contiguous coverage of the whole body, in order, no gaps and no overlap
    assert [a for a, _ in spans] == [0] + [b + 1 for _, b in spans[:-1]]
    assert spans[-1][1] == rep["paragraphs_body_direct"] - 1


def test_section_specs_survive_a_document_without_pgnumtype(plain):
    rep = wa.probe_xml(plain)
    assert rep["sections_xml"] == 1
    spec = rep["section_specs"][0]
    assert (spec["fmt"], spec["start"], spec["pgNumType_count"]) == (None, None, 0)
    assert spec["index"] == 1
    assert spec["first_para_index"] == 0
    assert spec["last_para_index"] == rep["paragraphs_body_direct"] - 1
    assert spec["last_para_index_xml"] == rep["paragraphs_xml"] - 1
    assert rep["page_number_formats"] == []
    assert rep["page_number_restarts"] == []


def test_duplicate_pgnumtype_is_reported_and_last_one_wins(tmp_path):
    """Word obeys the *last* ``pgNumType``; the duplicate is a warning, not a fix.

    This is the shape ``python-docx`` produces when a caller appends instead of
    replacing after ``add_section()`` cloned the sectPr.  Measured on Word 16: a
    section holding ``lowerRoman start=1`` then ``upperRoman start=5`` prints
    ``V`` — the earlier element is dead, so reading it would report the wrong
    alphabet.
    """
    doc = Document()
    doc.add_paragraph("cover")
    doc.add_section()
    sec = doc.sections[1]
    first = OxmlElement("w:pgNumType")
    first.set(qn("w:fmt"), "lowerRoman")
    first.set(qn("w:start"), "1")
    sec._sectPr.append(first)
    second = OxmlElement("w:pgNumType")
    second.set(qn("w:fmt"), "upperRoman")
    second.set(qn("w:start"), "5")
    sec._sectPr.append(second)
    doc.add_paragraph("body")

    rep = wa.probe_xml(save(doc, tmp_path))
    spec = rep["section_specs"][1]
    assert (spec["fmt"], spec["start"]) == ("upperRoman", "5")
    assert spec["pgNumType_count"] == 2
    assert any("more than one w:pgNumType" in w for w in rep["xml_warnings"])


def test_page_glyph_uses_the_section_counter_alphabet():
    """A displayed number is a *number*; a reader sees a *glyph*.

    ``3`` under ``lowerRoman`` is ``iii`` — the reported evidence is exactly this
    conversion, so it is pinned with the real counter values rather than mocked.
    """
    assert wa.page_glyph("lowerRoman", 3) == "iii"
    assert wa.page_glyph("upperRoman", 4) == "IV"
    assert wa.page_glyph("lowerLetter", 1) == "a"
    assert wa.page_glyph("upperLetter", 27) == "AA"
    assert wa.page_glyph("decimalZero", 7) == "07"
    assert wa.page_glyph("decimal", 7) == "7"
    assert wa.page_glyph(None, 5) == "5"       # no pgNumType: plain decimal
    assert wa.page_glyph("lowerRoman", 1) == "i"


def test_page_glyph_degrades_without_docx_section(monkeypatch):
    """``scripts/docx_section`` is another agent's file — losing it must not throw.

    ``page_glyph`` is called from the Word layer, which must never raise; the
    documented fallback is "the number as text", and that is what this asserts.
    """
    monkeypatch.setattr(wa, "_section_module", lambda: None)
    assert wa.page_glyph("lowerRoman", 3) == "3"
    assert wa.page_glyph(None, 12) == "12"
    assert wa.page_glyph("decimal", "7") == "7"
    assert wa.page_glyph("decimal", None) == ""


def test_roman_token_parsing():
    """TOC tokens are Roman as often as decimal; ``ii`` must sort after ``i``."""
    assert wa._token_number("i") == 1
    assert wa._token_number("iii") == 3
    assert wa._token_number("iv") == 4
    assert wa._token_number("xi") == 11
    assert wa._token_number("12") == 12
    assert wa._token_number("0") == 0
    assert wa._token_number("abc") is None
    assert wa._token_number(None) is None
    assert wa._token_number("") is None


def test_toc_audit_counts_levelled_entries(tmp_path):
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    add_style(doc.styles.element, "10", "toc 2")
    for style_id, label, page in (("9", "第一章", "1"), ("9", "第二章", "5"),
                                  ("10", "1.1 节", "2")):
        p = doc.add_paragraph()
        set_pstyle(p, style_id)
        p.add_run(label).add_tab()
        p.add_run(page)
    doc.add_paragraph("正文无样式")

    ta = wa.probe_xml(save(doc, tmp_path))["toc_audit"]
    assert ta["entries"] == 3
    assert ta["entries_toc_style"] == 3
    assert ta["styles"] == ["toc 1", "toc 2"]
    assert ta["levels"] == [1, 2]
    assert ta["page_tokens"] == ["1", "5", "2"]
    assert ta["page_parse_failures"] == []
    assert ta["source"] == "toc"


def test_toc_audit_flags_non_monotonic_pages(tmp_path):
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    for label, page in (("甲", "1"), ("乙", "7"), ("丙", "3")):
        p = doc.add_paragraph()
        set_pstyle(p, "9")
        p.add_run(label).add_tab()
        p.add_run(page)
    ta = wa.probe_xml(save(doc, tmp_path))["toc_audit"]
    assert ta["page_tokens"] == ["1", "7", "3"]
    assert ta["monotonic"] is False


def test_toc_audit_monotonic_handles_roman(tmp_path):
    """``i, ii, iii`` increases; as text it would not.  Convert before comparing."""
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    for label, page in (("摘要", "i"), ("Abstract", "ii"), ("目录", "iii")):
        p = doc.add_paragraph()
        set_pstyle(p, "9")
        p.add_run(label).add_tab()
        p.add_run(page)
    ta = wa.probe_xml(save(doc, tmp_path))["toc_audit"]
    assert ta["page_tokens"] == ["i", "ii", "iii"]
    assert ta["monotonic"] is True
    assert ta["monotonic_partial"] is False


def test_toc_audit_records_entries_without_a_page_number(tmp_path):
    """An entry that renders but carries no number is a broken TOC, not an empty one."""
    doc = Document()
    add_style(doc.styles.element, "9", "toc 1")
    p = doc.add_paragraph()
    set_pstyle(p, "9")
    p.add_run("没有页码的一行")
    q_ = doc.add_paragraph()
    set_pstyle(q_, "9")
    q_.add_run("有页码").add_tab()
    q_.add_run("4")

    ta = wa.probe_xml(save(doc, tmp_path))["toc_audit"]
    assert ta["entries"] == 2
    assert ta["page_tokens"] == ["4"]
    assert len(ta["page_parse_failures"]) == 1
    assert ta["page_parse_failures"][0]["text"].startswith("没有页码")
    assert ta["source"] == "toc"          # the TOC exists; only one row is unreadable


def test_toc_audit_reports_absence_as_computed(plain):
    ta = wa.probe_xml(plain)["toc_audit"]
    assert ta["entries"] == 0
    assert ta["page_tokens"] == []
    assert ta["monotonic"] is True        # vacuously; entries=0 says why
    assert ta["source"] == "computed"


def test_paragraph_pages_absent_without_a_request(plain):
    """No request and no Word -> no rows, and no invented page numbers."""
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off"]))
    assert rep["paragraph_pages"] == []
    assert rep["paragraph_pages_source"] == "none"
    assert rep["sources"]["paragraph_pages"] == "none"


def test_paragraph_pages_computed_fallback_is_labelled(tmp_path):
    """Without Word the rows exist for structure only, and say so on every row.

    This is the guard against the failure this layer was built to stop: a
    "computed" number being handed to a reader as the page a paragraph is on.

    Body layout here is ``0 cover`` / ``1 (sectPr)`` / ``2 front`` /
    ``3 (sectPr)`` / ``4 body``; ``--page-probe`` counts Word paragraphs, which
    are 1-based, so 1, 3, 5 name those three paragraphs in order.

    The third section sets ``decimal`` explicitly.  It has to: a bare
    ``add_section()`` *clones* the previous ``sectPr``, so section 3 would
    otherwise inherit ``lowerRoman`` and this test would be asserting Word's
    inheritance rather than the formatter's behaviour.
    """
    doc = Document()
    doc.add_paragraph("cover")
    doc.add_section()
    set_pgnum(doc.sections[1], "lowerRoman", "1")
    doc.add_paragraph("front")
    doc.add_section()
    set_pgnum(doc.sections[2], "decimal")
    doc.add_paragraph("body")
    path = save(doc, tmp_path)

    rep = wa.build_report(wa.build_parser().parse_args(
        [str(path), "--word", "off", "--page-probe", "1,3,5"]))
    assert rep["paragraph_pages_source"] == "computed"
    assert rep["sources"]["paragraph_pages"] == "computed"
    rows = rep["paragraph_pages"]
    assert [r["index"] for r in rows] == [1, 3, 5]
    assert all(r["source"] == "computed" for r in rows)
    assert all(r["exact"] is False for r in rows)
    assert all(r["physical"] is None for r in rows)   # nothing measured
    assert all(r["index_space"] == "xml" for r in rows)
    # section assignment is real even when the page is not
    assert [r["section"] for r in rows] == [1, 2, 3]
    assert [r["page_format"] for r in rows] == [None, "lowerRoman", "decimal"]
    # the glyph comes from that section's fmt, from the section's start number
    assert rows[1]["displayed_number"] == 1
    assert rows[1]["displayed_text"] == "i"
    assert rows[2]["displayed_text"] == "1"
    assert any("computed" in w for w in rep["warnings"])


def test_paragraph_pages_unknown_index_is_not_a_page(tmp_path):
    """An index outside the body gets ``None``, never a neighbouring page."""
    doc = Document()
    doc.add_paragraph("only")
    path = save(doc, tmp_path, "one.docx")
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(path), "--word", "off", "--page-probe", "1,99"]))
    by_index = {r["index"]: r for r in rep["paragraph_pages"]}
    assert by_index[1]["section"] == 1
    assert by_index[1]["displayed_text"] == "1"
    assert by_index[99]["section"] is None
    assert by_index[99]["displayed_number"] is None
    assert by_index[99]["displayed_text"] is None
    assert by_index[99]["basis"] == "unknown-section"


def test_estimated_paragraph_pages_shape():
    """The helper is usable direct, and its warning names the trap.

    ``indices`` here follow the same 1-based convention as ``--page-probe``:
    ``1`` is the first paragraph of section 1, ``5`` a paragraph of section 2.
    """
    specs = [
        {"index": 1, "first_para_index": 0, "last_para_index": 2,
         "first_para_index_xml": 0, "last_para_index_xml": 2,
         "fmt": None, "start": None},
        {"index": 2, "first_para_index": 3, "last_para_index": 9,
         "first_para_index_xml": 3, "last_para_index_xml": 9,
         "fmt": "upperRoman", "start": "1"},
    ]
    est = wa.estimated_paragraph_pages(specs, [{"index": 2, "last_page": 7}],
                                       [1, 5, 9])
    assert est["source"] == "computed"
    assert "computed 不得当作交付页码" in est["warning"]
    assert [r["section"] for r in est["rows"]] == [1, 2, 2]
    assert est["rows"][0]["displayed_text"] == "1"
    assert est["rows"][1]["page_format"] == "upperRoman"
    assert est["rows"][1]["displayed_text"] == "I"
    assert est["rows"][1]["section_last_page"] == 7
    assert all(r["exact"] is False and r["physical"] is None
               for r in est["rows"])


def test_table_rows_warn_about_the_index_space(tmp_path):
    """XML ordinal != Word paragraph index once a table exists.

    Measured on ``work/subjects/math/paper.docx``: XML ordinal 95 maps to section
    5 while Word's ``Paragraphs(96)`` is in section 4 — one table is enough to
    shift everything after it.  So the estimate says which index space it is in
    and the report warns when that space can drift.
    """
    doc = Document()
    doc.add_paragraph("有表格")
    t = doc.add_table(rows=2, cols=1)
    t.cell(0, 0).text = "c0"
    t.cell(1, 0).text = "c1"
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(save(doc, tmp_path)), "--word", "off", "--page-probe", "1"]))
    assert rep["paragraph_pages_source"] == "computed"
    assert rep["paragraph_pages"][0]["index_space"] == "xml"
    assert any("table row" in w for w in rep["warnings"])


def test_page_probe_cli_parsing():
    parser = wa.build_parser()
    args = parser.parse_args(["x.docx", "--page-probe", "1, 9,9, 43"])
    assert args.page_probe_indices == [1, 9, 43]     # deduped, order kept

    args2 = parser.parse_args(["x.docx"])
    assert args2.page_probe_indices is None

    with pytest.raises(SystemExit):
        parser.parse_args(["x.docx", "--page-probe", "3,abc"])
    with pytest.raises(SystemExit):
        parser.parse_args(["x.docx", "--page-probe", "0"])


def test_paragraph_pages_with_word_are_marked_word(plain, monkeypatch):
    """When Word answers, ``source`` flips to ``word`` and both numbers are kept.

    Word is stubbed at the ``probe_word`` boundary, because CI has no Office: what
    is asserted is the *merge contract* — a layer that returns readings must not
    be relabelled as an estimate, and ``physical`` and ``displayed_number`` must
    both survive into the report.
    """
    def fake_word(docx, **kw):
        assert kw.get("page_probe_indices") == [1]
        return {
            "word_ok": True, "skipped": False, "skipped_reason": None,
            "errors": [], "warnings": [], "pages": 8, "paragraphs": 8,
            "sections": 1, "bookmarks": 0, "bookmarks_total": 0,
            "inline_shapes": 0, "floating_shapes": 0, "fields": 0,
            "hyperlinks": 0, "tables": 0, "bookmark_names": [],
            "bookmark_names_all": [], "hyperlink_targets": [], "field_list": [],
            "page_numbers": [], "page_numbers_source": "none",
            "section_details": [{"index": 1, "last_page": 7,
                                 "last_page_displayed_number": 2,
                                 "last_page_displayed_text": "ii",
                                 "page_format": "lowerRoman"}],
            "paragraph_pages": [{"index": 1, "physical": 7,
                                 "displayed_number": 2, "displayed_text": "ii",
                                 "section": 1, "page_format": "lowerRoman"}],
            "paragraph_pages_source": "word",
            "pdf": None, "pngs": [], "pdf_pages": None, "pdf_page_footers": [],
            "word_version": "16.0",
        }

    monkeypatch.setattr(wa, "probe_word", fake_word)
    wa.reset_word_cache()
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "on", "--page-probe", "1"]))
    assert rep["paragraph_pages_source"] == "word"
    assert rep["sources"]["paragraph_pages"] == "word"
    row = rep["paragraph_pages"][0]
    assert row["physical"] == 7          # sheet ordinal
    assert row["displayed_number"] == 2  # the number the footer shows
    assert row["displayed_text"] == "ii"
    assert not any("computed" in w for w in rep["warnings"])
    assert rep["section_details"][0]["last_page_displayed_text"] == "ii"


def test_report_keeps_page_truth_keys_when_word_is_off(plain):
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(plain), "--word", "off"]))
    assert rep["page_number_formats"] == []
    assert rep["page_number_restarts"] == []
    assert len(rep["section_specs"]) == 1
    assert rep["toc_audit"]["source"] == "computed"
    assert rep["xml"]["toc_audit"]["entries"] == 0


def test_render_summary_shows_page_format_and_restarts(tmp_path):
    """The human summary must state the numbering scheme, not only the counts.

    "pages=8" hides the fact that sections print 1,2,3 then i — the line a
    reader needs is the fmt list plus where the counter restarts.
    """
    doc = Document()
    doc.add_paragraph("cover")
    doc.add_section()
    set_pgnum(doc.sections[1], "lowerRoman", "1")
    doc.add_paragraph("front")
    rep = wa.build_report(wa.build_parser().parse_args(
        [str(save(doc, tmp_path)), "--word", "off"]))
    text = wa.render_summary(rep)
    assert "page format" in text
    assert "lowerRoman" in text
    assert "restarts at" in text
    assert "[word_accept]" in text
