# MIT License. End-to-end tests for the docx skill.
"""Pytest suite proving create / read / edit / template round-trips.

Runs the scripts as subprocesses (argparse CLIs) and also verifies the
outputs with python-docx directly. Stdlib + python-docx only; all
fixtures are generated on the fly; no network.

Merged-repo note (P3-3 / T1c): this file was transplanted from the local
branch into the merged tree. ``SKILL`` resolves to the ``docx-merged``
root and ``SCRIPTS`` to ``docx-merged/scripts`` -- the layout is
identical, so no fixture adaptation was needed. Two deliberate
adjustments were made to follow the merged (upstream) behaviour, which is
the project's basis:

1. ``numbered_list`` now emits real Word automatic numbering
   (``numbering.xml`` + ``w:numPr``) instead of relying on the built-in
   ``List Number`` style, so ``test_styles_used`` asserts the numbering
   structure rather than the style name.
2. ``docx_template.py`` invocations now spell out the ``fill``
   subcommand instead of leaning on the implicit default-subcommand
   compatibility shim.
"""
from __future__ import annotations

import json
import os
import struct
import subprocess
import sys
import zlib
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import pytest
from docx import Document

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"

NON_ASCII = "Фамилия — ‘test’"


def make_png(path: Path) -> None:
    """Write a tiny valid 2x2 red PNG using only stdlib."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", 2, 2, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * 2 for _ in range(2))
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    path.write_bytes(png)


def run(script: str, *args: str) -> dict:
    env = dict(os.environ)
    env["LC_ALL"] = "C"  # prove no locale-default text reads
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        capture_output=True, env=env, check=False)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    return json.loads(proc.stdout.decode("utf-8"))


# ------------------------------------------------- automatic numbering probe
#
# ``docx_create``'s ``numbered_list`` block writes real Word numbering: each
# item is a ``<w:p>`` whose ``<w:pPr>`` carries ``<w:numPr>`` with an
# ``<w:ilvl>`` and a ``<w:numId>``, and ``word/numbering.xml`` holds the
# matching ``<w:abstractNum>`` / ``<w:num>`` definition.  These two helpers
# read that structure straight out of the package so tests can assert on it.

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def numbering_paragraphs(path: Path) -> list:
    """Return ``[(text, numId, ilvl), ...]`` for paragraphs carrying numPr.

    ``numId``/``ilvl`` are ``None`` when the corresponding child is absent.
    """
    with zipfile.ZipFile(str(path)) as zf:
        root = ElementTree.fromstring(zf.read("word/document.xml"))
    found = []
    for para in root.iter(q("p")):
        num_pr = para.find(f"{q('pPr')}/{q('numPr')}")
        if num_pr is None:
            continue
        text = "".join(t.text or "" for t in para.iter(q("t")))
        num_id = num_pr.find(q("numId"))
        ilvl = num_pr.find(q("ilvl"))
        found.append((
            text,
            None if num_id is None else num_id.get(q("val")),
            None if ilvl is None else ilvl.get(q("val")),
        ))
    return found


def numbering_definitions(path: Path) -> dict:
    """Map ``numId -> {abstractNumId, numFmt, lvlText, start}``."""
    with zipfile.ZipFile(str(path)) as zf:
        if "word/numbering.xml" not in zf.namelist():
            return {}
        root = ElementTree.fromstring(zf.read("word/numbering.xml"))

    abstract = {}
    for an in root.findall(q("abstractNum")):
        lvl = an.find(q("lvl"))          # ilvl 0 is what create emits
        abstract[an.get(q("abstractNumId"))] = {
            "numFmt": None if lvl is None or lvl.find(q("numFmt")) is None
            else lvl.find(q("numFmt")).get(q("val")),
            "lvlText": None if lvl is None or lvl.find(q("lvlText")) is None
            else lvl.find(q("lvlText")).get(q("val")),
            "start": None if lvl is None or lvl.find(q("start")) is None
            else lvl.find(q("start")).get(q("val")),
        }

    defs = {}
    for num in root.findall(q("num")):
        aid = num.find(q("abstractNumId"))
        aid = None if aid is None else aid.get(q("val"))
        entry = {"abstractNumId": aid, "numFmt": None, "lvlText": None,
                 "start": None}
        if aid in abstract:
            entry.update(abstract[aid])
        defs[num.get(q("numId"))] = entry
    return defs


@pytest.fixture(scope="module")
def workdir(tmp_path_factory) -> Path:
    return tmp_path_factory.mktemp("docxskill")


@pytest.fixture(scope="module")
def created(workdir: Path) -> Path:
    """Create a document exercising every create feature."""
    png = workdir / "pic.png"
    make_png(png)
    spec = {
        "page": {"width_mm": 210, "height_mm": 297,
                 "margins_mm": {"top": 25, "bottom": 25,
                                "left": 20, "right": 20}},
        "header": "Report header",
        "footer": "Page footer",
        "styles": [{"name": "FancyNote", "base": "Normal", "font": "Arial",
                    "size_pt": 11, "italic": True, "color": "1F4E79"}],
        "blocks": [
            {"type": "heading", "text": "Main Title", "level": 1},
            {"type": "heading", "text": "Section One", "level": 2},
            {"type": "paragraph", "runs": [
                {"text": "plain "},
                {"text": "boldbit", "bold": True},
                {"text": " italicbit", "italic": True},
                {"text": " underbit", "underline": True}]},
            {"type": "paragraph", "text": "Styled note.",
             "style": "FancyNote"},
            {"type": "bullet_list", "items": ["alpha", "beta"]},
            {"type": "numbered_list", "items": ["first", "second"]},
            {"type": "table", "header": ["Name", "Qty"],
             "rows": [["Widget", "3"], ["Gadget", "5"]],
             "style": "Table Grid", "header_bold": True},
            {"type": "image", "path": str(png), "width_mm": 30},
            {"type": "page_break"},
            {"type": "paragraph", "text": "After the break."},
        ],
    }
    spec_path = workdir / "spec.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    out = workdir / "created.docx"
    res = run("docx_create.py", spec_path, out)
    assert res["ok"] and out.exists()
    return out


class TestCreateAndRead:
    def test_text_roundtrip(self, created: Path):
        text = run("docx_read.py", created, "--text")
        body = "\n".join(text["body"])
        for expected in ("Main Title", "plain boldbit italicbit underbit",
                         "Styled note.", "alpha", "second",
                         "After the break."):
            assert expected in body
        assert text["tables"] == [[["Name", "Qty"], ["Widget", "3"],
                                   ["Gadget", "5"]]]
        assert "Report header" in text["headers"]
        assert "Page footer" in text["footers"]

    def test_structure(self, created: Path):
        st = run("docx_read.py", created, "--structure")
        outline = [(h["level"], h["text"]) for h in st["outline"]]
        assert (1, "Main Title") in outline
        assert (2, "Section One") in outline
        assert st["table_count"] == 1
        assert st["tables"][0] == {"rows": 3, "cols": 2}

    def test_styles_used(self, created: Path):
        styles = run("docx_read.py", created, "--styles")["styles"]
        # Merged-repo behaviour: ``numbered_list`` no longer leans on Word's
        # built-in "List Number" style.  ``docx_create`` attaches a real
        # ``w:numPr`` pointing at a ``w:numId`` defined in
        # ``word/numbering.xml`` (abstractNum + num), because the "List
        # Number" style carries no numPr and paragraphs styled with it
        # render with no numbers at all.  So "List Number" is legitimately
        # absent from the used-style list, and asserting the numbering
        # structure is *stronger* than asserting a style name: presence of
        # the style would prove nothing about the numbering actually
        # appearing on the page.
        for s in ("Heading 1", "FancyNote", "List Bullet", "Table Grid"):
            assert s in styles
        assert "List Number" not in styles, (
            "merged docx_create numbers lists structurally; a 'List Number' "
            "paragraph style reappearing here means it fell back to the old "
            "style-only (unnumbered) behaviour")

        # -- the real numbering evidence ----------------------------------
        with zipfile.ZipFile(str(created)) as zf:
            assert "word/numbering.xml" in zf.namelist()

        numbered = numbering_paragraphs(created)
        assert [t for (t, _nid, _lvl) in numbered] == ["first", "second"]
        num_ids = {nid for (_t, nid, _lvl) in numbered}
        assert len(num_ids) == 1 and None not in num_ids, num_ids
        assert {lvl for (_t, _n, lvl) in numbered} == {"0"}

        defs = numbering_definitions(created)
        num_id = num_ids.pop()
        assert num_id in defs, f"numId {num_id} not defined in numbering.xml"
        assert defs[num_id]["abstractNumId"] is not None
        assert defs[num_id]["numFmt"] == "decimal"
        assert defs[num_id]["lvlText"] == "%1."

    def test_images_extracted(self, created: Path, workdir: Path):
        outdir = workdir / "media"
        res = run("docx_read.py", created, "--images", outdir)
        assert len(res["images"]) == 1
        img = Path(res["images"][0])
        assert img.read_bytes().startswith(b"\x89PNG")

    def test_run_formatting_persisted(self, created: Path):
        doc = Document(str(created))
        para = next(p for p in doc.paragraphs if "boldbit" in p.text)
        flags = {r.text.strip(): (r.bold, r.italic, r.underline)
                 for r in para.runs if r.text.strip()}
        assert flags["boldbit"][0] is True
        assert flags["italicbit"][1] is True
        assert flags["underbit"][2] is True

    def test_page_setup(self, created: Path):
        sec = Document(str(created)).sections[0]
        assert round(sec.page_width.mm) == 210
        assert round(sec.top_margin.mm) == 25

    def test_revisions_detection(self, created: Path):
        rev = run("docx_read.py", created, "--revisions")
        assert rev["has_tracked_changes"] is False
        assert rev["comments"] is False

    def test_text_includes_nested_tables_and_header_variants(self, workdir: Path):
        path = workdir / "nested-and-headers.docx"
        doc = Document()
        outer = doc.add_table(rows=1, cols=1)
        outer.cell(0, 0).text = "outer"
        nested = outer.cell(0, 0).add_table(rows=1, cols=1)
        nested.cell(0, 0).text = "nested-only"
        outer.cell(0, 0).add_paragraph("after-nested")
        section = doc.sections[0]
        section.different_first_page_header_footer = True
        section.first_page_header.paragraphs[0].text = "first-header"
        section.even_page_header.paragraphs[0].text = "even-header"
        doc.save(path)

        text = run("docx_read.py", path, "--text")
        assert text["tables"][0][0][0].split("\n") == ["outer", "nested-only", "after-nested"]
        assert "first-header" in text["headers"]
        assert "even-header" in text["headers"]


class TestEdit:
    def test_replace_preserves_formatting(self, created: Path, workdir: Path):
        out = workdir / "edited.docx"
        res = run("docx_edit.py", "replace", created, "--find", "boldbit",
                  "--replace", "REPLACED", "-o", out)
        assert res["replacements"] == 1
        doc = Document(str(out))
        para = next(p for p in doc.paragraphs if "REPLACED" in p.text)
        run_ = next(r for r in para.runs if "REPLACED" in r.text)
        assert run_.bold is True  # formatting survived

    def test_set_cell(self, created: Path, workdir: Path):
        out = workdir / "cell.docx"
        run("docx_edit.py", "set-cell", created, "--table", "0", "--row",
            "1", "--col", "1", "--text", "99", "-o", out)
        assert Document(str(out)).tables[0].cell(1, 1).text == "99"

    def test_insert_and_delete(self, created: Path, workdir: Path):
        out = workdir / "ins.docx"
        run("docx_edit.py", "insert", created, "--index", "0", "--text",
            "Inserted first", "-o", out)
        doc = Document(str(out))
        assert doc.paragraphs[0].text == "Inserted first"
        out2 = workdir / "del.docx"
        run("docx_edit.py", "delete", out, "--index", "0", "-o", out2)
        assert Document(str(out2)).paragraphs[0].text != "Inserted first"

    def test_apply_style(self, created: Path, workdir: Path):
        out = workdir / "styled.docx"
        doc = Document(str(created))
        idx = next(i for i, p in enumerate(doc.paragraphs)
                   if p.text == "After the break.")
        run("docx_edit.py", "style", created, "--index", str(idx),
            "--style", "Heading 2", "-o", out)
        doc2 = Document(str(out))
        assert doc2.paragraphs[idx].style.name == "Heading 2"


class TestTemplate:
    def test_fill_everywhere_non_ascii(self, workdir: Path):
        # Build a template: tokens in body, split runs, table, header, footer.
        tpl = workdir / "tpl.docx"
        doc = Document()
        doc.sections[0].header.paragraphs[0].text = "H: {{name}}"
        doc.sections[0].footer.paragraphs[0].text = "F: {{date}}"
        p = doc.add_paragraph()
        p.add_run("Dear {{na")          # token split across runs
        p.add_run("me}}, hello.")
        t = doc.add_table(rows=1, cols=2)
        t.cell(0, 0).text = "{{name}}"
        t.cell(0, 1).text = "{{ date }}"   # spaced variant
        doc.add_paragraph("Unfilled: {{missing}}")
        doc.save(str(tpl))

        values = workdir / "values.json"
        values.write_text(
            json.dumps({"name": NON_ASCII, "date": "2026-08-08"},
                       ensure_ascii=False), encoding="utf-8")
        out = workdir / "filled.docx"
        # Explicit ``fill`` subcommand -- the bare 3-positional form still
        # works through docx_template's compatibility shim, but naming the
        # subcommand keeps the test pinned to the documented CLI.
        res = run("docx_template.py", "fill", tpl, values, out)
        assert res["ok"] is True
        assert res["unfilled_tokens"] == ["missing"]

        text = run("docx_read.py", out, "--text")
        assert f"Dear {NON_ASCII}, hello." in text["body"]
        assert text["tables"][0][0] == [NON_ASCII, "2026-08-08"]
        assert f"H: {NON_ASCII}" in text["headers"]
        assert "F: 2026-08-08" in text["footers"]

    def test_strict_fails_on_unfilled(self, workdir: Path):
        tpl = workdir / "tpl2.docx"
        doc = Document()
        doc.add_paragraph("{{gone}}")
        doc.save(str(tpl))
        values = workdir / "empty.json"
        values.write_text("{}", encoding="utf-8")
        env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS / "docx_template.py"), "fill",
             str(tpl), str(values), str(workdir / "out2.docx"), "--strict"],
            capture_output=True, env=env, check=False)
        assert proc.returncode == 1
        payload = json.loads(proc.stdout.decode("utf-8"))
        assert payload["unfilled_tokens"] == ["gone"]


# --------------------------------------------------------------- new parity
# (``W`` and ``q`` are defined near the top, beside the numbering probe.)


def run_raw(script: str, *args: str):
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / script), *map(str, args)],
        capture_output=True, env=env, check=False)


def _add_ins(para, rev_id: int, text: str, author="Editor"):
    from lxml import etree
    ins = etree.SubElement(para._p, q("ins"))
    ins.set(q("id"), str(rev_id))
    ins.set(q("author"), author)
    ins.set(q("date"), "2026-01-02T03:04:05Z")
    r = etree.SubElement(ins, q("r"))
    t = etree.SubElement(r, q("t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _add_del(para, rev_id: int, text: str, author="Editor"):
    from lxml import etree
    dele = etree.SubElement(para._p, q("del"))
    dele.set(q("id"), str(rev_id))
    dele.set(q("author"), author)
    dele.set(q("date"), "2026-01-02T03:04:05Z")
    r = etree.SubElement(dele, q("r"))
    t = etree.SubElement(r, q("delText"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


@pytest.fixture()
def tracked(tmp_path: Path) -> Path:
    """Doc with a tracked insertion + deletion in body AND in a table."""
    doc = Document()
    p = doc.add_paragraph("Base ")
    _add_ins(p, 1, "ADDED")
    _add_del(p, 2, "REMOVED")
    table = doc.add_table(rows=1, cols=1)
    cp = table.cell(0, 0).paragraphs[0]
    cp.add_run("Cell ")
    _add_ins(cp, 3, "CELLADD")
    _add_del(cp, 4, "CELLGONE")
    path = tmp_path / "tracked.docx"
    doc.save(str(path))
    return path


class TestRevisions:
    def test_list(self, tracked: Path):
        res = run("docx_revisions.py", "list", tracked)
        revs = {r["id"]: r for r in res["revisions"]}
        assert len(revs) == 4
        assert revs["1"] == {"id": "1", "author": "Editor",
                             "date": "2026-01-02T03:04:05Z",
                             "type": "insertion", "text": "ADDED"}
        assert revs["2"]["type"] == "deletion"
        assert revs["2"]["text"] == "REMOVED"
        assert revs["3"]["text"] == "CELLADD"  # inside table
        assert revs["4"]["type"] == "deletion"

    def test_accept_all(self, tracked: Path, tmp_path: Path):
        out = tmp_path / "acc.docx"
        res = run("docx_revisions.py", "accept-all", tracked, "-o", out)
        assert res["resolved"] == 4
        doc = Document(str(out))
        assert doc.paragraphs[0].text == "Base ADDED"
        assert doc.tables[0].cell(0, 0).text == "Cell CELLADD"
        assert run("docx_revisions.py", "list", out)["revisions"] == []

    def test_reject_all(self, tracked: Path, tmp_path: Path):
        out = tmp_path / "rej.docx"
        run("docx_revisions.py", "reject-all", tracked, "-o", out)
        doc = Document(str(out))
        assert doc.paragraphs[0].text == "Base REMOVED"
        assert doc.tables[0].cell(0, 0).text == "Cell CELLGONE"

    def test_accept_single_by_id(self, tracked: Path, tmp_path: Path):
        out = tmp_path / "one.docx"
        res = run("docx_revisions.py", "accept", tracked, "--id", "1",
                  "-o", out)
        assert res["resolved"] == 1
        doc = Document(str(out))
        assert doc.paragraphs[0].text == "Base ADDED"  # del 2 unresolved
        remaining = run("docx_revisions.py", "list", out)["revisions"]
        assert sorted(r["id"] for r in remaining) == ["2", "3", "4"]

    def test_reject_single_by_id(self, tracked: Path, tmp_path: Path):
        out = tmp_path / "rone.docx"
        run("docx_revisions.py", "reject", tracked, "--id", "2", "-o", out)
        doc = Document(str(out))
        assert doc.paragraphs[0].text == "Base REMOVED"  # ins 1 unresolved

    def test_unknown_id_fails(self, tracked: Path, tmp_path: Path):
        proc = run_raw("docx_revisions.py", "accept", tracked, "--id",
                       "999", "-o", tmp_path / "x.docx")
        assert proc.returncode == 1


class TestComments:
    @pytest.fixture()
    def base(self, tmp_path: Path) -> Path:
        doc = Document()
        doc.add_paragraph("The quarterly revenue rose sharply.")
        doc.add_paragraph("Second paragraph.")
        path = tmp_path / "base.docx"
        doc.save(str(path))
        return path

    def test_add_list_delete(self, base: Path, tmp_path: Path):
        out = tmp_path / "com.docx"
        res = run("docx_comments.py", "add", base, "--target",
                  "quarterly revenue", "--text", "Needs a source",
                  "--author", "Reviewer", "--initials", "R", "-o", out)
        assert res["ok"] is True
        cid = res["comment_id"]

        listed = run("docx_comments.py", "list", out)["comments"]
        assert len(listed) == 1
        c = listed[0]
        assert c["id"] == cid
        assert c["author"] == "Reviewer"
        assert c["text"] == "Needs a source"
        assert c["anchored_text"] == "quarterly revenue"
        assert c["date"]

        # document text unchanged by anchoring
        text = run("docx_read.py", out, "--text")
        assert "The quarterly revenue rose sharply." in text["body"]

        out2 = tmp_path / "nocom.docx"
        run("docx_comments.py", "delete", out, "--id", cid, "-o", out2)
        assert run("docx_comments.py", "list", out2)["comments"] == []
        text2 = run("docx_read.py", out2, "--text")
        assert "The quarterly revenue rose sharply." in text2["body"]

    def test_xml_fallback_path(self, base: Path, tmp_path: Path):
        out = tmp_path / "xmlcom.docx"
        res = run("docx_comments.py", "add", base, "--target",
                  "Second paragraph", "--text", "fallback note",
                  "--author", "Bot", "--xml", "-o", out)
        assert res["native_api"] is False
        listed = run("docx_comments.py", "list", out)["comments"]
        assert listed[0]["text"] == "fallback note"
        assert listed[0]["anchored_text"] == "Second paragraph"
        # file still opens cleanly
        assert Document(str(out)).paragraphs[1].text == "Second paragraph."

    def test_missing_target_fails(self, base: Path, tmp_path: Path):
        proc = run_raw("docx_comments.py", "add", base, "--target",
                       "not present", "--text", "x", "-o",
                       tmp_path / "y.docx")
        assert proc.returncode == 1


class TestValidate:
    def test_healthy_file_passes(self, created: Path):
        res = run("docx_validate.py", created)
        assert res["ok"] is True
        assert all(i["severity"] != "error" for i in res["issues"])

    def test_not_a_zip(self, tmp_path: Path):
        """非 zip → **exit 2**（P9-C 改口；旧断言是 exit 1 + ``not-a-zip`` JSON）。

        旧语义把「你给的路径根本不是 docx」也塞进 ``issues`` 里，退 1 —— 于是
        「路径写错了」和「文档坏了」在调用方眼里是同一件事，任何
        ``if rc: 去修文档`` 的自动化都会去修一份根本不是 docx 的文件。

        P9-C 定的双契约：**输入错误退 2（stdout 空）**，**内容问题仍退 1
        （stdout 是带 ``issues`` 的 JSON）**。后一半由
        ``tests/test_validate_errors.py`` 单独钉住（缺 ``[Content_Types].xml``
        的那条），所以这里只验前一半。
        """
        bad = tmp_path / "bad.docx"
        bad.write_bytes(b"this is not a zip file")
        proc = run_raw("docx_validate.py", bad)
        assert proc.returncode == 2
        assert proc.stdout == b"", "输入错误路径 stdout 必须为空"
        assert b"not a ZIP package" in proc.stderr
        assert b"Traceback" not in proc.stderr

    def test_dangling_rel_and_empty_image(self, created: Path,
                                          tmp_path: Path):
        import shutil
        import zipfile
        broken = tmp_path / "broken.docx"
        shutil.copy(created, broken)
        # rebuild the zip: drop the image part, zero out nothing else
        src = zipfile.ZipFile(str(created))
        with zipfile.ZipFile(str(broken), "w") as dst:
            for item in src.infolist():
                if item.filename.startswith("word/media/"):
                    dst.writestr(item.filename, b"")  # empty image
                else:
                    dst.writestr(item, src.read(item.filename))
        proc = run_raw("docx_validate.py", broken)
        assert proc.returncode == 1
        rep = json.loads(proc.stdout.decode("utf-8"))
        codes = {i["code"] for i in rep["issues"]}
        assert "empty-image" in codes

    def test_missing_style(self, tmp_path: Path):
        import zipfile
        doc = Document()
        doc.add_paragraph("styled", style="Heading 1")
        path = tmp_path / "styles.docx"
        doc.save(str(path))
        # rewrite document.xml to reference a style id that doesn't exist
        src = zipfile.ZipFile(str(path))
        broken = tmp_path / "badstyle.docx"
        with zipfile.ZipFile(str(broken), "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "word/document.xml":
                    data = data.replace(b'w:val="Heading1"',
                                        b'w:val="GhostStyle"')
                dst.writestr(item, data)
        proc = run_raw("docx_validate.py", broken)
        assert proc.returncode == 1
        rep = json.loads(proc.stdout.decode("utf-8"))
        assert any(i["code"] == "missing-style" and "GhostStyle"
                   in i["detail"] for i in rep["issues"])


class TestNormalize:
    def test_merges_split_runs(self, tmp_path: Path):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Hel")            # identical (no) formatting, split
        p.add_run("lo wo")
        p.add_run("rld")
        b = p.add_run("BOLD1")
        b.bold = True
        b2 = p.add_run("BOLD2")
        b2.bold = True
        i = p.add_run("ital")
        i.italic = True
        path = tmp_path / "split.docx"
        doc.save(str(path))

        out = tmp_path / "norm.docx"
        res = run("docx_edit.py", "normalize", path, "-o", out)
        assert res["runs_merged"] == 3  # 2 plain merges + 1 bold merge

        doc2 = Document(str(out))
        para = doc2.paragraphs[0]
        assert para.text == "Hello worldBOLD1BOLD2ital"
        assert [r.text for r in para.runs] == \
            ["Hello world", "BOLD1BOLD2", "ital"]
        assert para.runs[1].bold is True
        assert para.runs[2].italic is True


class TestFields:
    def test_toc_and_page_numbers_via_edit(self, created: Path,
                                           tmp_path: Path):
        out = tmp_path / "fields.docx"
        run("docx_edit.py", "toc", created, "--index", "0", "-o", out)
        run("docx_edit.py", "page-numbers", out)

        import zipfile
        doc_xml = zipfile.ZipFile(str(out)).read(
            "word/document.xml").decode("utf-8")
        assert "TOC \\o" in doc_xml
        assert "fldChar" in doc_xml
        footer_names = [n for n in zipfile.ZipFile(str(out)).namelist()
                        if n.startswith("word/footer")]
        footers = "".join(zipfile.ZipFile(str(out)).read(n).decode("utf-8")
                          for n in footer_names)
        assert "PAGE" in footers and "NUMPAGES" in footers
        # still a valid document
        assert run("docx_validate.py", out)["ok"] is True

    def test_toc_and_footer_in_create_spec(self, tmp_path: Path):
        spec = {
            "footer_page_numbers": True,
            "blocks": [
                {"type": "toc"},
                {"type": "heading", "text": "Chapter", "level": 1},
            ],
        }
        spec_path = tmp_path / "fspec.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        out = tmp_path / "fcreate.docx"
        run("docx_create.py", spec_path, out)

        import zipfile
        z = zipfile.ZipFile(str(out))
        assert "TOC \\o" in z.read("word/document.xml").decode("utf-8")
        footers = "".join(z.read(n).decode("utf-8") for n in z.namelist()
                          if n.startswith("word/footer"))
        assert "NUMPAGES" in footers


# ------------------------------------------------- spec validation (P6-FIX-C5E)
#
# ``docx_create`` reads every key with ``dict.get`` and never complained, so the
# only possible outcome of a misspelt key was silence: ``{"blockz": [...]}``
# built a valid, empty document, reported ``"blocks": 0`` and exited 0.  The
# author had no way to see that the content never went in.
#
# The contract these tests pin down:
#
#   0  the document was written (JSON on stdout)
#   1  it was built but is not deliverable (dangling reference, bad bookmark)
#   2  the *spec* was rejected before anything was written — one readable line
#      on stderr, no traceback, and no file on disk to mistake for output


def _create_raw(tmp_path: Path, spec, name: str = "spec.json"):
    """Run ``docx_create.py`` on *spec* (a dict, or raw text); return (proc, out).

    ``spec`` as a string is written verbatim, which is how the malformed-JSON
    case gets its breakage past ``json.dumps``.
    """
    spec_path = tmp_path / name
    if isinstance(spec, str):
        spec_path.write_text(spec, encoding="utf-8")
    else:
        spec_path.write_text(json.dumps(spec, ensure_ascii=False),
                             encoding="utf-8")
    out = tmp_path / (name + ".docx")
    return run_raw("docx_create.py", spec_path, out), out


def _create_error(tmp_path: Path, spec, name: str = "spec.json") -> str:
    """Assert the spec is *rejected* (exit 2, no traceback, no file); return why."""
    proc, out = _create_raw(tmp_path, spec, name)
    stderr = proc.stderr.decode("utf-8", "replace")
    stdout = proc.stdout.decode("utf-8", "replace")
    assert proc.returncode == 2, "expected exit 2, got %d\nstdout: %s\nstderr: %s" \
        % (proc.returncode, stdout, stderr)
    assert "Traceback" not in stderr, (
        "a rejected spec must not raise a traceback:\n" + stderr)
    assert stderr.strip(), "exit 2 must say what is wrong on stderr"
    payload = json.loads(stdout)          # stdout stays machine-readable
    assert payload["ok"] is False
    assert payload["output"] is None
    assert not out.exists(), (
        "exit 2 promises nothing was written, but %s exists" % out)
    return payload["error"]


class TestSpecValidation:
    """Unknown keys, unknown types and out-of-range values are errors, not no-ops."""

    def test_unknown_top_level_key_is_rejected(self, tmp_path: Path):
        # The exact failure mode from the audit: ``blocks`` misspelt, so the
        # whole document was silently empty.
        error = _create_error(tmp_path, {
            "blockz": [{"type": "paragraph", "text": "第一章"}],
        })
        assert "blockz" in error
        assert 'did you mean "blocks"' in error, error
        assert "valid keys here" in error

    def test_unknown_block_key_is_rejected_and_suggests(self, tmp_path: Path):
        # Both typos in one run: the message must name each one, say which block
        # it is in, and suggest the real spelling — one fix pass, not two.
        error = _create_error(tmp_path, {"blocks": [
            {"type": "heading", "text": "第一章", "levl": 1},
            {"type": "paragraph", "text": "正文", "stylish": "Body"},
        ]})
        assert "blocks[0]" in error and "heading" in error
        assert "levl" in error and 'did you mean "level"' in error
        assert "blocks[1]" in error
        assert "stylish" in error and 'did you mean "style"' in error

    def test_unknown_nested_key_is_rejected(self, tmp_path: Path):
        # Not only top level / block level: page margins, styles, runs and tab
        # stops are the same trap, one level deeper.
        error = _create_error(tmp_path, {
            "page": {"margins_mm": {"top": 25, "tp": 25}},
            "blocks": [],
        })
        assert "margins_mm" in error and '"tp"' in error, error

    def test_bad_json_is_rejected_without_traceback(self, tmp_path: Path):
        # Truncated file: json.load used to propagate JSONDecodeError.
        error = _create_error(
            tmp_path, '{\n  "blocks": [\n    {"type": "paragraph", "text": "x"}\n',
            name="broken.json")
        assert "not valid JSON" in error
        assert "line" in error and "column" in error

    def test_unknown_block_type_is_rejected(self, tmp_path: Path):
        error = _create_error(tmp_path, {"blocks": [
            {"type": "paragraph", "text": "第一段"},
            {"type": "callout", "text": "未知块类型"},
        ]})
        assert "blocks[1]" in error
        assert "callout" in error
        # the message lists the vocabulary that would have worked
        assert "paragraph" in error and "heading" in error

    def test_heading_level_out_of_range_is_rejected(self, tmp_path: Path):
        # python-docx takes 0..9; 12 used to surface as its own ValueError with
        # a traceback.  Reported with the range, and *not* clamped.
        error = _create_error(tmp_path, {"blocks": [
            {"type": "heading", "text": "越界标题", "level": 12},
        ]})
        assert "out of range" in error and "12" in error
        assert "0..9" in error, error
        assert "not clamped" in error

    def test_non_numeric_heading_level_is_rejected(self, tmp_path: Path):
        error = _create_error(tmp_path, {"blocks": [
            {"type": "heading", "text": "第一章", "level": "1"},
        ]})
        assert "whole number" in error, error

    def test_table_row_wider_than_header_is_rejected(self, tmp_path: Path):
        # ``cells[i]`` used to raise IndexError on the third cell of a 2-column
        # table — a bare traceback for a one-character spec mistake.
        error = _create_error(tmp_path, {"blocks": [
            {"type": "table", "header": ["A", "B"], "rows": [["1", "2", "3"]]},
        ]})
        assert "rows[0]" in error and "3 cells" in error
        assert "2 column" in error

    def test_valid_spec_still_exits_zero(self, tmp_path: Path):
        """Guard: the strictness must not reject a spec that was always legal."""
        png = tmp_path / "pic.png"
        make_png(png)
        spec = {
            "page": {"width_mm": 210, "height_mm": 297,
                     "margins_mm": {"top": 25, "bottom": 25,
                                    "left": 20, "right": 20}},
            "header": "Header",
            "footer_page_numbers": True,
            "styles": [{"name": "Note", "base": "Normal", "font": "Arial",
                        "size_pt": 11, "italic": True,
                        "outline_level": 2}],
            "blocks": [
                {"type": "heading", "text": "Title", "level": 1},
                {"type": "paragraph", "style": "Note", "text": "Body.",
                 "first_line_indent_chars": 2, "line_spacing": 1.5},
                {"type": "bullet_list", "items": ["a", "b"]},
                {"type": "numbered_list", "items": ["one", "two"],
                 "num_text": "[%1]", "bookmark_prefix": "ref_"},
                {"type": "table", "header": ["H1", "H2"],
                 "rows": [["1", "2"]], "header_bold": True},
                {"type": "image", "path": str(png), "width_mm": 20},
                {"type": "page_break"},
            ],
        }
        proc, out = _create_raw(tmp_path, spec, name="ok.json")
        assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
        payload = json.loads(proc.stdout.decode("utf-8"))
        assert payload["ok"] is True and payload["blocks"] == 7
        assert out.exists()
        text = run("docx_read.py", out, "--text")
        assert "Body." in "\n".join(text["body"])

    def test_every_documented_block_key_is_accepted(self, tmp_path: Path):
        """Whitelist reconciliation: every documented key, every block type.

        This is the half of the contract that protects existing specs — the
        strict check must not mistake a legitimate key for a typo.  Every key
        named in the module docstring (and read by the builder) appears here at
        least once; if one is ever dropped from ``BLOCK_KEYS`` this test turns
        red instead of a user's spec turning red.
        """
        png = tmp_path / "ks.png"
        make_png(png)
        spec = {
            "page": {"width_mm": 210, "height_mm": 297, "orient": "portrait",
                     "margins_mm": {"top": 25, "bottom": 25, "left": 20,
                                    "right": 20, "header": 12, "footer": 12,
                                    "gutter": 0}},
            "header": "Header",
            "footer": "Footer",
            "footer_page_numbers": {"template": "page_of", "fmt": "decimal",
                                    "start": 1, "align": "center",
                                    "total": "numpages", "prefix": "p ",
                                    "suffix": " / ",
                                    "font": "Times New Roman", "size_pt": 10,
                                    "bold": True, "page_type": "default"},
            "styles": [{"name": "KS", "base": "Normal", "font": "Arial",
                        "font_east_asia": "SimSun", "size_pt": 12,
                        "bold": False, "italic": False, "color": "1F4E79",
                        "outline_level": 2}],
            "blocks": [
                {"type": "heading", "text": "第一章", "level": 1,
                 "alignment": "left", "keep_with_next": True,
                 "bookmark": "h_book"},
                {"type": "paragraph", "style": "KS", "text": "正文",
                 "indent_left_mm": 5, "indent_right_mm": 5,
                 "first_line_indent_chars": 2, "char_width_pt": 12,
                 "first_line_indent_font_pt": 12, "line_spacing": 1.5,
                 "space_before_pt": 3, "space_after_pt": 3,
                 "alignment": "justify", "keep_with_next": False,
                 "page_break_before": False, "bookmark": "p_book",
                 "tab_stops": [{"pos_mm": 100, "align": "right",
                                "leader": "dot"},
                               {"pos_cm": 5.0, "align": "left",
                                "leader": "none"}],
                 "bottom_border": {"style": "single", "size": 8, "space": 1,
                                   "color": "000000",
                                   "sides": ["bottom"]}},
                {"type": "paragraph", "runs": [
                    {"text": "plain "},
                    {"text": "biu", "bold": True, "italic": True,
                     "underline": True},
                    {"text": "sup", "vert_align": "superscript"},
                    {"text": "sup2", "superscript": True},
                    {"text": "sub", "subscript": True}]},
                {"type": "bullet_list", "items": ["alpha", {"text": "beta"}],
                 "style": "List Bullet"},
                {"type": "numbered_list", "items": ["one", "two"],
                 "num_text": "[%1]", "num_fmt": "decimal", "start": 1,
                 "indent_left": 425, "hanging": 425,
                 "bookmark_prefix": "ref_", "ilvl": 0},
                {"type": "caption", "label": "图", "text": "示意",
                 "style": "Caption", "bookmark": "cap_book",
                 "bookmark_span": "number", "ref_bookmark": True,
                 "drop_text": False},
                {"type": "table_caption", "label": "表", "text": "表格题注",
                 "bookmark": "tcap_book"},
                {"type": "crossref", "target": "cap_book", "prefix": "如图",
                 "suffix": "", "format": "default", "cached": "1",
                 "superscript": True, "vert_align": "superscript",
                 "hyperlink": True, "own_paragraph": True},
                {"type": "table", "header": ["列一", "列二"],
                 "rows": [["1", "2"]], "style": "Table Grid",
                 "header_bold": True, "caption": "表格题注二",
                 "caption_label": "表", "caption_bookmark": "tbl_book"},
                {"type": "image", "path": str(png), "width_mm": 20,
                 "alignment": "center", "bookmark": "img_book"},
                {"type": "page_break", "bookmark": "pb_book"},
                {"type": "toc", "instr": ' TOC \\o "1-3" \\h \\z \\u '},
                {"type": "section_break", "start_type": "new_page",
                 "page": {"width_mm": 210, "height_mm": 297,
                          "orient": "portrait",
                          "margins_mm": {"top": 25, "bottom": 25,
                                         "left": 20, "right": 20}},
                 "header": "第二章页眉", "header_align": "center",
                 "header_font": "Arial", "header_size_pt": 9,
                 "header_bold": False,
                 "footer": "第二节页脚", "footer_align": "center",
                 "footer_font": "Arial", "footer_size_pt": 9,
                 "footer_bold": False,
                 "footer_page_numbers": {"template": "page",
                                         "fmt": "lowerRoman", "start": 1,
                                         "align": "right"},
                 "different_first_page": False, "cols": 1,
                 "no_footer": False, "reuse_last_paragraph": True,
                 "apply_to": "next"},
            ],
        }
        proc, out = _create_raw(tmp_path, spec, name="all-keys.json")
        assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
        payload = json.loads(proc.stdout.decode("utf-8"))
        assert payload["ok"] is True
        assert payload["blocks"] == len(spec["blocks"]) == 13
        assert payload["sections"] == 2
        assert out.exists()
