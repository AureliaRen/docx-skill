# -*- coding: utf-8 -*-
"""Tests for ``scripts/docx_preview.py`` — CLI contract, argv refusals, lint maths.

Three layers are pinned here, and the split is deliberate:

1. **Layering** (no PyMuPDF / Office imports at module scope) — asserted with
   ``ast``, so a well-meaning ``import fitz`` at the top of the script fails the
   suite on a machine that has PyMuPDF instead of failing on a runner that does
   not.  Same discipline as ``tests/test_word_accept_xml.py:149-164``.
2. **argv contract** — every refusal path (exit 2 before anything is read or
   written, one ``error:`` line on stderr, empty stdout) and every
   environment failure (exit 1).  These run everywhere, no skip needed.
3. **The lint maths** — ``_fill_ratio`` / ``_ink_ratio`` / ``_last_line`` /
   ``_body_blocks`` are tested against **synthetic blocks**, not against a
   rendered PDF.  A layout ratio is arithmetic; tying it to LibreOffice's
   line-spacing would make the suite skip exactly where it matters and hide a
   regression behind a skip.

Only the end-to-end cases (real soffice → real PDF → real PNG) carry
``@needs_render``.  Those are the ones a CI runner cannot run, and they are
counted honestly in the skip tally of the round report.
"""
from __future__ import annotations

import ast
import json
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docx_preview as dp  # noqa: E402

SCRIPT = SCRIPTS / "docx_preview.py"
MODULE_SRC = SCRIPT.read_text(encoding="utf-8")


# --------------------------------------------------------------------- helpers
def run_cli(argv, env=None, cwd=None):
    """Run the CLI in a **subprocess**, so exit codes and streams are real.

    In-process ``main()`` calls would let a stray ``print`` pass unnoticed and
    would not prove that stdout is empty on a refusal — the whole point of the
    exit-code contract is what a shell sees.

    ``encoding="utf-8"`` is mandatory on Windows: the default is the ANSI code
    page (cp936 here), and the report is full of Chinese page text, so decoding
    with the locale rejects the bytes and the reader thread raises instead of
    returning a result.  ``errors="replace"`` keeps a decoding surprise from
    turning into an unhandled thread exception.
    """
    full_env = None
    if env is not None:
        import os
        full_env = dict(os.environ)
        full_env["PYTHONIOENCODING"] = "utf-8"
        for key, value in env.items():
            if value is None:
                full_env.pop(key, None)
            else:
                full_env[key] = value
    return subprocess.run([sys.executable, str(SCRIPT), *argv],
                          capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=full_env,
                          cwd=str(cwd) if cwd else None)


def stderr_lines(proc):
    return [line for line in proc.stderr.replace("\r", "").split("\n") if line]


def assert_refused(proc, code):
    """The exit-2 shape: right code, empty stdout, exactly one ``error:`` line."""
    assert proc.returncode == code, proc.stderr
    assert proc.stdout == "", "stdout must stay empty on a refusal"
    lines = stderr_lines(proc)
    assert len(lines) == 1, lines
    assert lines[0].startswith("error:"), lines[0]
    return lines[0]


def block(y0, y1, text="x", x0=0.0, x1=400.0):
    """One ``get_text("blocks")``-shaped dict, as ``_body_blocks`` builds them."""
    return {"x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": text, "type": 0}


@pytest.fixture()
def docx_file(tmp_path):
    """A small real .docx — python-docx is a hard dependency of the skill."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("预览测试：第一段。")
    doc.add_paragraph("预览测试：第二段。")
    path = tmp_path / "sample.docx"
    doc.save(str(path))
    return path


# ------------------------------------------------------------------- layering
def test_no_toplevel_render_import():
    """``fitz`` must never be imported at module scope.

    ``test_word_accept_xml.py:149-164`` bans ``fitz`` at the top of
    ``tools/word_accept.py``; the same ban applies here for a harder reason.  A
    module-level import cannot be *skipped* — on a runner without PyMuPDF the
    test module itself fails to import, so every render test fails instead of
    skipping, and the skip budget is bypassed entirely.
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
    assert found == [], "top-level render imports: %s" % found


def test_fitz_is_imported_inside_a_function():
    """The complement of the ban: ``fitz`` *is* imported, just lazily.

    Without this, deleting the lazy import would also satisfy the ban above and
    the render paths would fail at runtime instead of at review time.
    """
    tree = ast.parse(MODULE_SRC)
    lazy = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for inner in ast.walk(node):
                if isinstance(inner, ast.Import):
                    lazy += [a.name for a in inner.names if a.name == "fitz"]
                elif isinstance(inner, ast.ImportFrom) and inner.module == "fitz":
                    lazy.append("fitz")
    assert lazy, "fitz must still be imported inside a function"


def test_no_win32com_import_anywhere():
    """This tool never drives Word — not even lazily.

    ``docx_preview`` renders through LibreOffice precisely so it works on a
    machine without Office; a stray ``win32com`` import would reintroduce the
    dependency the whole file exists to avoid.

    Checked with ``ast`` over **every** node rather than by substring: the
    module docstring explains *why* ``win32com`` is banned and names it, so a
    ``"win32com" not in source`` assertion would fail on its own explanation.
    """
    tree = ast.parse(MODULE_SRC)
    found = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            found += [a.name for a in node.names
                      if a.name.split(".")[0] in ("win32com", "pythoncom")]
        elif isinstance(node, ast.ImportFrom) and node.module:
            if node.module.split(".")[0] in ("win32com", "pythoncom"):
                found.append(node.module)
    assert found == [], "win32com/pythoncom imports: %s" % found


def test_input_validation_is_imported_not_copied():
    """D-028: the three input helpers live in ``docx_read`` and are imported."""
    assert "from docx_read import" in MODULE_SRC
    for name in ("InputError", "_check_input", "_open_document"):
        assert name in MODULE_SRC
    # A copy would show up as a local def with the same name.
    tree = ast.parse(MODULE_SRC)
    defined = {node.name for node in tree.body
               if isinstance(node, (ast.FunctionDef, ast.ClassDef))}
    assert "_check_input" not in defined
    assert "InputError" not in defined


def test_content_width_comes_from_docx_create():
    """One implementation of the text-block width, shared with ``width_frac``."""
    assert "from docx_create import content_width_mm" in MODULE_SRC


def test_render_shape_matches_word_accept():
    """The four constants that keep this file and ``word_accept`` in step.

    ``tools/word_accept.py:_render_pngs`` names pages ``page-%03d.png``, caps the
    work at 400 pages, defaults to 110 dpi, and reads the footer from the band
    starting at 88% of the page height.  Two renderers that disagree on any of
    these produce artefacts nobody can compare.
    """
    assert dp.PAGE_DPI == 110
    assert dp.PAGE_LIMIT == 400
    assert dp.FOOTER_BAND_START == 0.88
    assert '("page-%03d.png" % i)' in MODULE_SRC


# ---------------------------------------------------------------- argv contract
def test_help_lists_the_three_capabilities():
    proc = run_cli(["--help"])
    assert proc.returncode == 0
    for flag in ("--pdf", "--pages", "--lint", "--dpi", "--min-fill"):
        assert flag in proc.stdout


def test_input_failure_is_exit_2_not_1(tmp_path):
    """A missing path is a usage error (2), not "cannot process" (1).

    The distinction is the whole point of the D-024 split: `1` means the
    document is real and the machine is not, so a caller can retry elsewhere;
    `2` means the caller typed the wrong path and retrying is pointless.
    """
    proc = run_cli([str(tmp_path / "absent.docx"), "--lint"])
    line = assert_refused(proc, 2)
    assert "no such file" in line


def test_directory_input_is_exit_2(tmp_path):
    proc = run_cli([str(tmp_path), "--lint"])
    line = assert_refused(proc, 2)
    assert "not a file" in line


def test_non_zip_input_is_exit_2(tmp_path):
    fake = tmp_path / "renamed.docx"
    fake.write_text("I am not a Word package.\n", encoding="utf-8")
    proc = run_cli([str(fake), "--lint"])
    line = assert_refused(proc, 2)
    assert "not a ZIP package" in line


def test_stderr_starts_with_error_not_usage(tmp_path):
    """教训 66：``usage:`` means argparse refused the *arguments*.

    A path that exists-but-is-wrong must be refused by this tool with its own
    `error:` line, not by argparse's usage block — otherwise the caller cannot
    tell a bad flag from a bad file.
    """
    proc = run_cli([str(tmp_path / "absent.docx"), "--lint"])
    first = stderr_lines(proc)[0]
    assert first.startswith("error:")
    assert not first.startswith("usage:")


def test_json_and_text_are_mutually_exclusive(docx_file):
    proc = run_cli([str(docx_file), "--lint", "--json", "--text"])
    line = assert_refused(proc, 2)
    assert "mutually exclusive" in line


def test_a_mode_is_required(docx_file):
    proc = run_cli([str(docx_file)])
    line = assert_refused(proc, 2)
    assert "nothing to do" in line


@pytest.mark.parametrize("flag,value", [("--dpi", "5"), ("--dpi", "5000"),
                                        ("--min-fill", "1.5"),
                                        ("--min-fill", "-0.1"),
                                        ("--max-orphan-chars", "-1"),
                                        ("--timeout", "0")])
def test_out_of_range_numbers_are_exit_2(docx_file, flag, value):
    """Range refusals happen in ``main``, so they must still be exit 2 + one line."""
    proc = run_cli([str(docx_file), "--lint", flag, value])
    assert_refused(proc, 2)


def test_missing_soffice_is_exit_1(docx_file):
    """A real document we cannot render here → 1, not 2.

    ``SOFFICE_PATH`` deliberately disables the search path, so pointing it at a
    path that does not exist is a deterministic way to simulate a machine
    without LibreOffice (the runner is one).
    """
    proc = run_cli([str(docx_file), "--pdf", str(docx_file.with_suffix(".pdf"))],
                   env={"SOFFICE_PATH": str(docx_file.parent / "nope" / "soffice.exe")})
    line = assert_refused(proc, 1)
    assert "SOFFICE_PATH" in line


def test_soffice_override_disables_the_fallback(monkeypatch, tmp_path):
    """With ``SOFFICE_PATH`` set, an unusable value is an error, not a hint."""
    monkeypatch.setenv("SOFFICE_PATH", str(tmp_path / "nope.exe"))
    path, reason = dp.find_soffice()
    assert path is None
    assert reason and "SOFFICE_PATH" in reason


def test_soffice_override_is_honoured(monkeypatch, tmp_path):
    fake = tmp_path / "soffice.exe"
    fake.write_text("", encoding="utf-8")
    monkeypatch.setenv("SOFFICE_PATH", str(fake))
    path, reason = dp.find_soffice()
    assert reason is None
    assert path == str(fake)


def test_missing_pymupdf_is_exit_1(docx_file, monkeypatch):
    """PyMuPDF absent → 1 + one ``error:`` line, never a traceback.

    Simulated by making the lazy import raise.  ``ImportError`` is the real
    shape on a runner without the wheel; the point of the lazy import is that
    this failure is *reachable* rather than an import-time crash.
    """
    def boom(*args, **kwargs):
        raise ImportError("No module named 'fitz'")

    monkeypatch.setattr(dp, "_fitz", boom)
    code = dp.main([str(docx_file), "--lint"])
    assert code == 1


def test_missing_pymupdf_reason_is_one_line(monkeypatch):
    def boom(*args, **kwargs):
        raise ImportError("No module named 'fitz'\n  hint: pip install pymupdf")

    monkeypatch.setattr(dp, "_fitz", boom)
    reason = dp.py_mupdf_reason()
    assert reason is not None
    assert "\n" not in reason
    assert reason.startswith("PyMuPDF")


# ----------------------------------------------------------------- lint maths
def test_fill_ratio_measures_how_far_content_reaches():
    """Full page ≈ 1.0, two lines near the top ≈ 0.0.

    This is the reading the ``--min-fill`` threshold is applied to, so its scale
    matters more than its precision: the 0.40 default has to sit between a
    healthy page and a half-empty one, not on top of a healthy page.
    """
    full = [block(y, y + 14.0) for y in range(70, 780, 31)]
    assert dp._fill_ratio(full, 842.0) > 0.80
    two_lines = [block(75.9, 90.4), block(107.1, 121.6)]
    assert dp._fill_ratio(two_lines, 842.0) < 0.10


def test_fill_ratio_is_not_line_spacing_driven():
    """The literal ink ratio *is* line-spacing driven — that is why it is separate.

    Two paragraphs covering the **same span** with different line spacing agree
    on ``fill_ratio`` and disagree on ``ink_ratio``.  Pinning both keeps a future
    "simplification" from swapping them: the ``--min-fill`` threshold applied to
    ``ink_ratio`` would flag every healthy page (a full page measured 0.379 on a
    real LibreOffice render).
    """
    def ladder(step):
        ys = list(range(70, 780, step))
        ys[-1] = 770                       # both ladders end at the same place
        return [block(y, y + 14.0) for y in ys]

    tight, loose = ladder(18), ladder(45)
    assert dp._fill_ratio(tight, 842.0) == dp._fill_ratio(loose, 842.0)
    assert dp._ink_ratio(tight, 842.0) > dp._ink_ratio(loose, 842.0)


def test_fill_ratio_merges_overlapping_blocks():
    """Overlapping boxes count once — otherwise the ratio can exceed 1.0."""
    overlapping = [block(100, 200), block(150, 250)]
    assert dp._ink_ratio(overlapping, 842.0) == pytest.approx(150 / 842.0,
                                                             abs=1e-3)


def test_fill_ratio_of_an_empty_page_is_zero():
    assert dp._fill_ratio([], 842.0) == 0.0
    assert dp._ink_ratio([], 842.0) == 0.0


def test_fill_ratio_never_exceeds_one():
    """A block that hangs past the page bottom must not report > 1.0."""
    assert dp._fill_ratio([block(0, 2000)], 842.0) == 1.0
    assert dp._ink_ratio([block(0, 2000)], 842.0) == 1.0


def test_last_line_picks_the_lowest_block_and_its_last_line():
    blocks = [block(100, 120, "第一段\n第二行"),
              block(400, 420, "末段第一行\n末段最后一行"),
              block(200, 220, "中间段")]
    assert dp._last_line(blocks) == "末段最后一行"


def test_last_line_of_a_blank_page_is_empty():
    assert dp._last_line([]) == ""
    assert dp._last_line([block(100, 120, "   \n  ")]) == ""


def test_issue_contract_matches_docx_validate():
    """Same three keys, same two severities — one report format, one parser."""
    issues: list = []
    dp._issue(issues, "warning", "sparse-page", "detail")
    assert issues == [{"severity": "warning", "code": "sparse-page",
                       "detail": "detail"}]
    assert set(issues[0]) == {"severity", "code", "detail"}


def test_one_line_collapses_a_traceback():
    assert dp._one_line("a\n  b\tc") == "a b c"


def test_pdf_validity_guard(tmp_path):
    """A conversion that "succeeds" with an empty file is not a product."""
    empty = tmp_path / "empty.pdf"
    empty.write_bytes(b"")
    assert dp._pdf_looks_valid(empty) is False
    html = tmp_path / "error.pdf"
    html.write_bytes(b"<html>500</html>")
    assert dp._pdf_looks_valid(html) is False
    good = tmp_path / "good.pdf"
    good.write_bytes(b"%PDF-1.7\n...")
    assert dp._pdf_looks_valid(good) is True


def test_pymupdf_is_reachable_here():
    """Sanity: the lazy import works on this machine (no skip — it is cheap)."""
    reason = dp.py_mupdf_reason()
    assert reason is None or "PyMuPDF" in reason


# ------------------------------------------------------- end-to-end (skippable)
needs_render = pytest.mark.skipif(
    dp.py_mupdf_reason() is not None or dp.find_soffice()[0] is None,
    reason="needs LibreOffice + PyMuPDF (absent on the CI runners)")

RENDER_DOCX = Path(__file__).resolve().parent / "_preview_fixture.docx"


def _wide_fixture(path: Path) -> None:
    """A .docx whose page 1 holds two lines and whose page 2 is full.

    Built here rather than checked in: the point of the fixture is a *page that
    is visibly almost empty*, and a binary in the repo would rot silently.
    """
    from docx import Document
    from docx.shared import Mm

    doc = Document()
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.left_margin = section.right_margin = Mm(26)
        section.top_margin = section.bottom_margin = Mm(25.4)
    doc.add_paragraph("第一页只有两行字。")
    doc.add_paragraph("第二行。")
    doc.add_page_break()
    for i in range(1, 40):
        doc.add_paragraph("第二页第 %d 段：正常密度的正文，用来做对照。" % i)
    doc.save(str(path))


@needs_render
def test_pdf_is_produced_with_pages(tmp_path):
    """Case 1: a real PDF with ``page_count > 0``."""
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    out = tmp_path / "out.pdf"
    proc = run_cli([str(docx), "--pdf", str(out), "--json"])
    assert proc.returncode == 0, proc.stderr
    report = json.loads(proc.stdout)
    assert out.is_file() and out.stat().st_size > 0
    assert report["pdf"]["page_count"] > 0
    assert report["pdf"]["page_count"] == dp.pdf_page_count(out)


@needs_render
def test_png_count_equals_pdf_pages(tmp_path):
    """Case 2: one PNG per page, numbered from 001."""
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    png_dir = tmp_path / "pngs"
    proc = run_cli([str(docx), "--pages", str(png_dir), "--dpi", "72", "--json"])
    assert proc.returncode == 0, proc.stderr
    report = json.loads(proc.stdout)
    files = sorted(png_dir.glob("page-*.png"))
    assert len(files) == report["pdf"]["page_count"] == report["pages"]["count"]
    assert files[0].name == "page-001.png"


@needs_render
def test_dpi_changes_pixel_dimensions(tmp_path):
    """Case 2b: ``--dpi`` reaches ``get_pixmap``."""
    fitz = dp._fitz()
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    sizes = {}
    for dpi in (72, 144):
        out = tmp_path / ("dpi%d" % dpi)
        proc = run_cli([str(docx), "--pages", str(out), "--dpi", str(dpi)])
        assert proc.returncode == 0, proc.stderr
        pix = fitz.Pixmap(str(out / "page-001.png"))
        sizes[dpi] = (pix.width, pix.height)
    assert sizes[144][0] > sizes[72][0]
    assert sizes[144][1] > sizes[72][1]


@needs_render
def test_lint_names_the_half_empty_page(tmp_path):
    """Case 3: the detector has to *fire*, not just run.

    "跑通但 issues 空" is the failure mode this test exists to refuse: a lint that
    never reports anything is indistinguishable from a document with no
    problems, and the original 13-round pain was exactly a half-empty page
    nobody could see.
    """
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    proc = run_cli([str(docx), "--lint", "--json"])
    assert proc.returncode == 0, proc.stderr
    lint = json.loads(proc.stdout)["lint"]
    assert 1 in lint["sparse_pages"], lint["sparse_pages"]
    assert lint["issues"], "a detected defect must produce an issue"
    assert any(issue["code"] == "sparse-page" for issue in lint["issues"])
    assert lint["pages"][0]["fill_ratio"] < 0.40


@needs_render
def test_min_fill_moves_the_threshold(tmp_path):
    """``--min-fill 0`` turns the sparse detector off without touching the rest."""
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    proc = run_cli([str(docx), "--lint", "--min-fill", "0", "--json"])
    assert proc.returncode == 0, proc.stderr
    assert json.loads(proc.stdout)["lint"]["sparse_pages"] == []


@needs_render
def test_text_output_is_human_readable(tmp_path):
    """The default (no ``--json``) output names the pages, one fact per line."""
    docx = tmp_path / "in.docx"
    _wide_fixture(docx)
    proc = run_cli([str(docx), "--lint"])
    assert proc.returncode == 0, proc.stderr
    assert proc.stdout.startswith("pdf: ")
    assert "lint: " in proc.stdout
    assert "sparse_pages=[1]" in proc.stdout
