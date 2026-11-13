#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Render a .docx the way a reader sees it: PDF, per-page PNG, layout lint JSON.

Usage::

    docx_preview.py IN.docx [--pdf OUT.pdf] [--pages DIR] [--dpi 110]
                            [--lint] [--min-fill 0.40] [--json | --text]

Three capabilities, all optional and composable (at least one is required):

``--pdf OUT.pdf``
    Convert with a headless LibreOffice (``soffice --convert-to pdf``).  The
    conversion is **synchronous**: ``subprocess.run`` waits for the child, so
    the PDF exists (or does not) when the call returns.  The old PowerShell
    advice ``Start-Process soffice ...`` without ``-Wait`` is the reason this
    bullet exists: it returns immediately, the exit code is empty, and the
    "converted" file is still being written.
``--pages DIR``
    One PNG per PDF page, named ``page-001.png`` … (``page-%03d.png``), at
    ``--dpi`` (default 110, the same default as
    ``tools/word_accept.py:probe_word``).
``--lint``
    A layout report: fill ratio per page, half-empty pages, blank pages,
    images wider than the text block, tables split across a page break, and a
    lone character on a page's last line.

Why this file exists
--------------------
Before it, the only way to *see* a generated .docx was: export to PDF by hand,
convert each page to PNG by hand, then look at the pictures.  That loop ate
most of a 13-round iteration on a real assignment -- every round paid the same
tax, and nothing in the repository could answer "is page 4 half empty?".  The
lint half of this tool answers it in one command.

Rendering shape is deliberately identical to
``tools/word_accept.py:_render_pngs`` (``:1585-1611``) and
``tools/word_accept.py:_pdf_from_word`` (``:1568-1582``): the same
``page-%03d.png`` names, the same ``limit=400`` overflow guard, the same
``get_pixmap(dpi=...)`` call, and the same footer text band
(``y >= 0.88 * page height``).  Two independent ``fitz`` implementations drift;
one shape does not.  See ``docs/reports/P12-R1B-PREVIEW.md`` §2 for the exact
差异表.

Exit codes (D-024 four-tier convention, ``3`` is not used here)
--------------------------------------------------------------
``0``  success.  Lint findings are *output*, not failure: a report that names
       three half-empty pages is a successful run.
``2``  usage / input error -- the path does not exist, is a directory, is not a
       ZIP package, is not a Word package, or the argv contradicts itself.
       Refused **before** anything is read or written: stdout stays empty and
       stderr holds exactly one ``error:``-prefixed line.
``1``  the input is a real document but this machine cannot process it: no
       LibreOffice, no PyMuPDF, the conversion failed, or the PDF LibreOffice
       produced has zero pages.

Layering
--------
``fitz`` is imported **inside** functions, never at module scope.  The reason is
pinned by ``tests/test_word_accept_xml.py:149-164``: a module-level
``import fitz`` cannot even be *imported* on a runner without PyMuPDF, so the
render tests could not be skipped -- they would fail at collection time.  The
same ban covers ``win32com`` / ``pythoncom`` / ``lxml``: this tool never drives
Word and never parses DOCX XML.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import contextlib
import io
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

# 输入校验三件套由 `docx_read` 单一拥有（D-028）：这里只 import、不复制判定。
# `docx_validate.py` / `docx_edit.py` / `docx_mathify.py` 同形。
from docx_read import InputError, _check_input, _open_document

# 正文宽度的唯一实现（`docx_create.py:199`）。P12-C 的 `width_frac` 也用它，
# 两处必须同一个数，否则「图片超宽」与「图片按 fraction 缩放」会互相打脸。
from docx_create import content_width_mm

# --------------------------------------------------------------------- constants
PAGE_DPI = 110
"""Default PNG resolution.  Same as ``word_accept.probe_word(dpi=110)``."""

PAGE_LIMIT = 400
"""Overflow guard copied from ``word_accept._render_pngs(limit=400)``: a
document with 900 pages is a mistake, not something to rasterise for six
minutes."""

FOOTER_BAND_START = 0.88
"""Body text ends above ``0.88 * page height``; below that is the footer.

Copied from ``word_accept._render_pngs`` (which reads the footer *from* that
band).  Here the same constant is used in the opposite direction: footer
content is **excluded** from the fill ratio and from the "last line" check,
because a page number sitting at the bottom of an otherwise empty page would
otherwise make that page look full.
"""

DEFAULT_MIN_FILL = 0.40
"""``--min-fill`` default: below this share of the page is "half empty".

0.40 is the number the P12 report asked for.  A dense A4 body page measures
roughly 0.75-0.90 by :func:`_fill_ratio`, a page holding two lines of text
roughly 0.02-0.05, so the threshold sits in a wide gap rather than on a cliff.
"""

DEFAULT_MARGIN_MM = 25.4
"""Fallback top/bottom margin when the section does not report one."""

PT_PER_MM = 72.0 / 25.4
"""1 mm in PostScript points (PDF user units)."""

WIDTH_TOLERANCE_PT = 1.0
"""Images are flagged only when they exceed the text block by more than a
point: margins round-trip through twips (a 20 mm margin reads back as
20.0015 mm), so an exactly-full-width image must not be reported."""

SPLIT_TOLERANCE_PT = 8.0
"""How close to the margin a table edge must sit to count as "ran off the
page".  LibreOffice shifts a table by a few points depending on row padding."""

DEFAULT_TIMEOUT = 180
"""Seconds allowed for one ``soffice`` conversion."""

SOFFICE_ENV = "SOFFICE_PATH"

SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


class PreviewError(Exception):
    """The document is fine, this machine is not (or the conversion failed).

    Distinct from :class:`docx_read.InputError` on purpose: ``InputError`` means
    "you named the wrong thing" (exit 2, refuse before doing anything), while
    this means "your document is real but I cannot render it here" (exit 1,
    after the input was accepted).  Keeping them apart is what makes the two
    exit codes honest instead of decorative.
    """


# ------------------------------------------------------------------- dependencies
def _fitz():
    """Import PyMuPDF lazily and return the module.

    **Never** at module scope -- see the module docstring.  A local import also
    means the "PyMuPDF missing" path is a clean ``exit 1`` with one line on
    stderr instead of an import-time traceback that no ``skipif`` can catch.
    """
    import fitz

    return fitz


def py_mupdf_reason() -> str | None:
    """``None`` when PyMuPDF is importable, else the one-line reason.

    Returning the reason rather than a bool keeps the diagnostic honest: "no
    PyMuPDF" and "PyMuPDF is installed but broken" are different repairs.
    """
    try:
        _fitz()
    except Exception as exc:  # ImportError, or a broken shared library
        return "PyMuPDF (fitz) is not importable: %s" % _one_line(exc)
    return None


def find_soffice() -> tuple[str | None, str | None]:
    """Locate the LibreOffice binary.  Returns ``(path, reason_when_missing)``.

    ``SOFFICE_PATH`` wins over every candidate and **disables the fallback**:
    when it is set, an unusable value is an error rather than a hint.  That is
    what makes the failure path testable -- point it at a path that does not
    exist and the tool must refuse instead of quietly finding the real
    LibreOffice next to it.
    """
    override = os.environ.get(SOFFICE_ENV, "").strip()
    if override:
        if Path(override).is_file():
            return override, None
        return None, ("%s=%s is not an existing file" % (SOFFICE_ENV, override))
    for candidate in SOFFICE_CANDIDATES:
        if Path(candidate).is_file():
            return candidate, None
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found, None
    return None, (
        "no LibreOffice found; looked at %s and at PATH. Set %s to the soffice "
        "binary, or install LibreOffice." % (", ".join(SOFFICE_CANDIDATES),
                                             SOFFICE_ENV))


def _one_line(exc_or_text) -> str:
    """Collapse anything to a single line -- stderr must be exactly one line."""
    return " ".join(str(exc_or_text).split())


def _tail(proc, limit: int = 300) -> str:
    """Last ``limit`` characters of a finished process' output, as one line.

    LibreOffice is chatty and sometimes writes the useful sentence in the middle
    of a Java stack trace; keep the tail and drop everything else so the error
    stays one line.
    """
    text = " ".join(((proc.stdout or "") + " " + (proc.stderr or "")).split())
    return text[-limit:] if len(text) > limit else text


# ----------------------------------------------------------------------- convert
def convert_to_pdf(docx_path: str, out_dir, soffice: str,
                   timeout: int = DEFAULT_TIMEOUT) -> tuple[str, float]:
    """``soffice --headless --convert-to pdf`` into *out_dir*.

    Returns ``(pdf_path, seconds)``.  Raises :class:`PreviewError` when the
    conversion fails, when LibreOffice exits 0 without writing a PDF, or when
    what it wrote is not a PDF (a zero-byte file happens when the profile is
    locked by another instance -- hence the throw-away
    ``-env:UserInstallation`` below).

    ``subprocess.run`` is *the* synchronous wait: it does not return until the
    child has exited.  The P10 landmine was
    ``Start-Process soffice -ArgumentList ...`` without ``-Wait``, which returns
    at once, leaves ``$LASTEXITCODE`` empty, and hands you a half-written file.
    There is no equivalent mistake available here.
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    # A throw-away profile per run: a second soffice sharing the default profile
    # blocks (or silently converts nothing).  as_uri() gives file:///C:/... on
    # Windows and file:///tmp/... elsewhere.
    profile = Path(tempfile.mkdtemp(prefix="docx-preview-lo-"))
    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nologo",
        "--nolockcheck",
        "-env:UserInstallation=" + profile.as_uri(),
        "--convert-to", "pdf",
        "--outdir", str(out),
        str(docx_path),
    ]
    started = time.time()
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              timeout=timeout, check=False)
    except subprocess.TimeoutExpired:
        raise PreviewError(
            "soffice did not finish within %d s converting %s"
            % (timeout, Path(docx_path).name)) from None
    except OSError as exc:
        raise PreviewError("cannot run soffice (%s): %s"
                           % (soffice, _one_line(exc))) from None
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    seconds = round(time.time() - started, 2)

    produced = out / (Path(docx_path).stem + ".pdf")
    if not produced.is_file():
        others = sorted(out.glob("*.pdf"))
        produced = others[0] if others else None
    if produced is None:
        raise PreviewError(
            "soffice exited %s but wrote no PDF: %s"
            % (proc.returncode, _tail(proc) or "(no output)"))
    if not _pdf_looks_valid(produced):
        raise PreviewError(
            "soffice wrote %s but it is not a PDF (%d bytes): %s"
            % (produced.name, produced.stat().st_size, _tail(proc) or "(no output)"))
    return str(produced), seconds


def _pdf_looks_valid(path) -> bool:
    """Cheap structural check: non-empty and starts with ``%PDF-``.

    This is the "exit code is not evidence" guard.  A conversion that returns 0
    while leaving a 0-byte or HTML-error file behind is the exact false green
    this check exists to refuse.
    """
    p = Path(path)
    try:
        if p.stat().st_size <= 0:
            return False
        with p.open("rb") as fh:
            return fh.read(5) == b"%PDF-"
    except OSError:
        return False


def pdf_page_count(pdf_path) -> int:
    """Page count, read through PyMuPDF.  Raises when the PDF cannot be opened."""
    fitz = _fitz()
    with fitz.open(str(pdf_path)) as pdf:
        return pdf.page_count


# ------------------------------------------------------------------------- pages
def render_pngs(pdf_path, png_dir, dpi: int = PAGE_DPI,
                limit: int = PAGE_LIMIT):
    """Per-page PNG.  Returns ``(paths, page_count, footers, truncated)``.

    Shape copied from ``tools/word_accept.py:_render_pngs`` (``:1585-1611``) so
    the two stay recognisably the same code:

    * the file name is ``page-%03d.png``, so page 7 sorts before page 10 and a
      shell glob is already in reading order;
    * ``page.get_pixmap(dpi=dpi).save(...)`` -- the same one-call raster;
    * ``limit`` caps the work and reports ``truncated`` instead of pretending
      the document had fewer pages;
    * the footer band is ``fitz.Rect(0, 0.88h, w, h)`` read with
      ``get_text("text", clip=band)``, which is also what the lint uses to keep
      page numbers out of the fill ratio.

    Differences from that function, both deliberate: this one *raises*
    :class:`PreviewError` instead of collecting an ``errors`` list (a CLI has
    somewhere to put the exit code), and it reports ``truncated``.
    """
    fitz = _fitz()
    out_dir = Path(png_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    footers = []
    pages = 0
    truncated = False
    try:
        with fitz.open(str(pdf_path)) as pdf:
            pages = pdf.page_count
            for i, page in enumerate(pdf, 1):
                if i > limit:
                    truncated = True
                    break
                target = out_dir / ("page-%03d.png" % i)
                page.get_pixmap(dpi=dpi).save(str(target))
                paths.append(str(target))
                band = fitz.Rect(0, page.rect.height * FOOTER_BAND_START,
                                 page.rect.width, page.rect.height)
                footers.append(" ".join(page.get_text("text", clip=band).split()))
    except Exception as exc:
        raise PreviewError("PNG rendering failed: %s" % _one_line(exc)) from None
    return paths, pages, footers, truncated


# -------------------------------------------------------------------------- lint
def _issue(issues, severity: str, code: str, detail: str) -> None:
    """One finding, in ``docx_validate._issue()``'s contract.

    ``{"severity": "error"|"warning", "code": ..., "detail": ...}`` -- the same
    three keys, the same two severities, so a caller that already understands a
    ``docx_validate`` report needs no second parser.  ``severity`` here is about
    how a reader would feel, not about the exit code: this tool exits 0 whenever
    it produced the report.
    """
    issues.append({"severity": severity, "code": code, "detail": detail})


def _body_blocks(page) -> list[dict]:
    """Text/image blocks above the footer band.

    ``get_text("blocks")`` returns ``(x0, y0, x1, y1, text, block_no, type)``.
    A block that *starts* inside the footer band is footer content (page number,
    "第 3 页 共 8 页") and is dropped -- see :data:`FOOTER_BAND_START`.
    """
    height = page.rect.height
    cutoff = height * FOOTER_BAND_START
    blocks = []
    for block in page.get_text("blocks"):
        x0, y0, x1, y1, text, _no, btype = block[:7]
        if y0 >= cutoff:
            continue
        blocks.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1,
                       "text": text, "type": btype})
    return blocks


def _merged_height(blocks) -> float:
    """Total height of the **merged** bbox intervals of *blocks*, in points.

    Merged, not summed: overlapping boxes inside one paragraph must count once,
    otherwise a page reports more than 1.0 and two pages cannot be compared.
    """
    spans = sorted((b["y0"], b["y1"]) for b in blocks if b["y1"] > b["y0"])
    covered = 0.0
    cur0 = cur1 = None
    for y0, y1 in spans:
        if cur1 is None:
            cur0, cur1 = y0, y1
        elif y0 > cur1:
            covered += cur1 - cur0
            cur0, cur1 = y0, y1
        else:
            cur1 = max(cur1, y1)
    if cur1 is not None:
        covered += cur1 - cur0
    return covered


def _fill_ratio(blocks, page_height: float) -> float:
    """How far down the page its content reaches: ``span / page height``.

    ``span`` is the vertical extent of the **union** of the block boxes
    (:func:`_merged_height` applied to the whole set: bottommost ``y1`` minus
    topmost ``y0``), so this answers "how much of the page does the content
    occupy" -- the question a half-empty page fails.

    Why not ``sum(block heights) / page height``: that number is dominated by
    **line spacing**, which has nothing to do with how full the page is.  A
    completely full A4 page set at 1.5 spacing measures ``0.38`` that way, so an
    absolute threshold of ``0.40`` (the :data:`DEFAULT_MIN_FILL` the P12 report
    asked for) would flag every healthy page and the detector would be useless.
    Measured on a real LibreOffice render: full page ``0.958``, two-line page
    ``0.017`` -- the threshold sits in a wide valley between them.  The ink
    measure is still reported, as ``ink_ratio``, and is explicitly *not* the
    number the threshold is applied to.
    """
    if page_height <= 0 or not blocks:
        return 0.0
    top = min(b["y0"] for b in blocks)
    bottom = max(b["y1"] for b in blocks)
    span = max(0.0, bottom - top)
    return round(min(1.0, span / page_height), 3)


def _ink_ratio(blocks, page_height: float) -> float:
    """Share of the page height covered by ink, ignoring the gaps between lines.

    Informational only: this is the literal "bbox union height / page height"
    reading, and it is line-spacing dependent (see :func:`_fill_ratio`), so it is
    reported but never thresholded.
    """
    if page_height <= 0 or not blocks:
        return 0.0
    return round(min(1.0, _merged_height(blocks) / page_height), 3)



def _last_line(blocks) -> str:
    """The last text line of the page body, stripped.  ``""`` when there is none.

    "Last" is decided by the lowest ``y1``, and within that block by the last
    non-empty ``splitlines()`` entry -- a paragraph block holds several lines, so
    taking ``text`` whole would answer a different question.
    """
    text_blocks = [b for b in blocks if b["text"].strip()]
    if not text_blocks:
        return ""
    last = max(text_blocks, key=lambda b: b["y1"])
    lines = [line.strip() for line in last["text"].splitlines() if line.strip()]
    return lines[-1] if lines else ""


def _image_boxes(page) -> list[tuple]:
    """Bounding boxes of the images placed on this page, in points."""
    try:
        infos = page.get_image_info()
    except Exception:
        return []
    boxes = []
    for info in infos:
        bbox = info.get("bbox")
        if bbox:
            boxes.append(tuple(float(v) for v in bbox))
    return boxes


def _table_boxes(page) -> list[tuple]:
    """Bounding boxes of the tables PyMuPDF's table finder sees, in points.

    ``find_tables()`` writes an advisory line to **stdout** (``Consider using the
    pymupdf_layout package ...``) the first time it runs in a process.  That
    would land in the middle of this tool's JSON, so stdout is captured around
    the call.  Returns ``[]`` on any failure: this is a heuristic layer, and a
    PDF whose ruling lines confuse the finder must degrade to "no table
    evidence" rather than fail the whole lint.
    """
    try:
        with contextlib.redirect_stdout(io.StringIO()):
            finder = page.find_tables()
    except Exception:
        return []
    boxes = []
    for table in getattr(finder, "tables", []):
        bbox = getattr(table, "bbox", None)
        if bbox:
            boxes.append(tuple(float(v) for v in bbox))
    return boxes


def _page_geometry(doc) -> dict:
    """Text-block width and top/bottom margins of the first section, in mm.

    Multi-section documents are approximated by their first section.  The
    alternative -- mapping every PDF page back to the section that produced it
    -- needs a page-to-section map Word itself would have to supply, and the
    error it introduces is a few millimetres on the *tolerance* of a width
    check, not on the check itself.

    ``content_width_mm`` comes from ``docx_create.content_width_mm`` (the single
    implementation, also used by ``width_frac``); a document whose section has
    no page width at all reports ``None`` and the width check is skipped with a
    warning instead of inventing a number.
    """
    geometry = {"content_width_mm": None, "top_margin_mm": None,
                "bottom_margin_mm": None}
    try:
        section = doc.sections[0]
    except (AttributeError, IndexError):
        return geometry
    try:
        geometry["content_width_mm"] = content_width_mm(section)
    except Exception:
        geometry["content_width_mm"] = None
    for key, attr in (("top_margin_mm", "top_margin"),
                      ("bottom_margin_mm", "bottom_margin")):
        margin = getattr(section, attr, None)
        try:
            geometry[key] = round(float(margin.mm), 2) if margin is not None else None
        except Exception:
            geometry[key] = None
    return geometry


def lint_pdf(pdf_path, geometry=None, min_fill: float = DEFAULT_MIN_FILL,
             max_orphan_chars: int = 2) -> dict:
    """Layout report for a PDF.  Pure reading: nothing is written.

    Checks, and what each one can and cannot see:

    ``fill_ratio`` / ``sparse_pages``
        Per page, the union of body-block heights divided by page height (see
        :func:`_fill_ratio`).  ``sparse_pages`` lists the pages below
        ``min_fill``.  **This is the half-empty-page detector** -- the defect
        that cost eight of thirteen rounds, because a page holding two lines of
        text looks fine in a paragraph-by-paragraph reading and looks broken in
        a PDF.
    ``blank_pages``
        No text **and** no image.  A full-page figure is therefore not blank
        (it has an image block); a page with only a stray paragraph mark is.
    ``image_over_width``
        An image wider than the text block, measured against
        ``content_width_mm`` from the DOCX section.  Needs the geometry; when the
        document reports no page width the check is skipped and one
        ``content-width-unknown`` warning says so, rather than guessing.
    ``table_split``
        **Heuristic, and it can be wrong both ways.**  A table is reported as
        split when, on page *n*, some table's bottom edge reaches the bottom
        margin (within :data:`SPLIT_TOLERANCE_PT`) *and* page *n+1* has a table
        whose top edge reaches the top margin.  False positive: a table that
        happens to end exactly at the bottom margin, followed by a *different*
        table that starts at the top of the next page.  False negative: a table
        split without ruled lines (PyMuPDF finds tables by ruling lines and text
        alignment), or a split hidden by a repeated header row that the finder
        merges.  Verify by looking at the two PNGs before acting on it.
    ``orphan_last``
        The last body line of a page holds at most ``max_orphan_chars``
        characters while the page holds more text than that line.  This catches
        the classic widow -- one character carried onto the next page -- and
        will also catch a legitimate two-character closing line (``完。``,
        ``"]"``).  **Heuristic.**  The footer band is excluded first, so a page
        number is never mistaken for an orphan.

    Severity: ``blank-page`` and ``image-over-width`` are ``error`` (a reader
    sees a broken page); ``sparse-page``, ``table-split``, ``orphan-last`` and
    ``content-width-unknown`` are ``warning``.
    """
    fitz = _fitz()
    geometry = geometry or {}
    content_width_mm_value = geometry.get("content_width_mm")
    content_width_pt = (float(content_width_mm_value) * PT_PER_MM
                        if content_width_mm_value else None)
    top_pt = (float(geometry.get("top_margin_mm") or DEFAULT_MARGIN_MM)
              * PT_PER_MM)
    # Only the top margin is needed: `table_split` asks "does the next page open
    # with a table", which is a top-edge question.  The bottom edge was tried and
    # discarded -- LibreOffice stops a split table wherever its last row ends,
    # measured 11 pt short of the bottom margin -- so `bottom_margin_mm` stays in
    # the geometry dict for callers but is not consumed here.

    pages: list[dict] = []
    issues: list[dict] = []
    sparse: list[int] = []
    blanks: list[int] = []
    over: list[dict] = []
    splits: list[dict] = []
    orphans: list[dict] = []
    heights: list[float] = []
    lowest_bottom: list[float | None] = []
    table_boxes: list[list[tuple]] = []

    if content_width_pt is None:
        _issue(issues, "warning", "content-width-unknown",
               "the document reports no section page width, so image widths "
               "were not checked against the text block")

    with fitz.open(str(pdf_path)) as pdf:
        count = pdf.page_count
        for number, page in enumerate(pdf, 1):
            height = page.rect.height
            heights.append(height)
            blocks = _body_blocks(page)
            images = _image_boxes(page)
            ratio = _fill_ratio(blocks, height)
            ink = _ink_ratio(blocks, height)
            text = "".join(b["text"] for b in blocks)
            chars = len(text.strip())
            last_line = _last_line(blocks)
            boxes = _table_boxes(page)
            table_boxes.append(boxes)
            lowest_bottom.append(max((b["y1"] for b in blocks), default=None))

            is_blank = chars == 0 and not images
            pages.append({
                "page": number,
                "fill_ratio": ratio,
                "ink_ratio": ink,
                "text_chars": chars,
                "block_count": len(blocks),
                "image_count": len(images),
                "last_line": last_line,
                "blank": is_blank,
            })

            if is_blank:
                blanks.append(number)
                _issue(issues, "error", "blank-page",
                       "page %d holds neither text nor an image" % number)
            elif ratio < min_fill:
                sparse.append(number)
                _issue(issues, "warning", "sparse-page",
                       "page %d: fill_ratio %.3f < %.2f (content covers %.1f%% "
                       "of the page height; %d char(s) in %d block(s))"
                       % (number, ratio, min_fill, ratio * 100.0, chars,
                          len(blocks)))

            if content_width_pt is not None:
                for index, box in enumerate(images):
                    width_pt = box[2] - box[0]
                    if width_pt > content_width_pt + WIDTH_TOLERANCE_PT:
                        width_mm = round(width_pt / PT_PER_MM, 2)
                        over.append({
                            "page": number,
                            "image": index,
                            "width_mm": width_mm,
                            "content_width_mm": content_width_mm_value,
                        })
                        _issue(issues, "error", "image-over-width",
                               "page %d image %d is %.1f mm wide but the text "
                               "block is %.1f mm" % (number, index, width_mm,
                                                     float(content_width_mm_value)))

            if 0 < len(last_line) <= max_orphan_chars and chars > len(last_line):
                orphans.append({"page": number, "text": last_line})
                _issue(issues, "warning", "orphan-last",
                       "page %d ends with a lone %r on its last line"
                       % (number, last_line))

        for index in range(count - 1):
            # "page n ends with a table" = the lowest body block on page n sits
            # inside one of page n's table boxes.  Comparing the table's own
            # bottom edge against the bottom margin does **not** work: LibreOffice
            # stops the table where the rows stop, which measured 11 pt short of
            # the margin on a real split.
            ends_in_table = (
                lowest_bottom[index] is not None
                and any(box[1] <= lowest_bottom[index] <= box[3] + SPLIT_TOLERANCE_PT
                        for box in table_boxes[index]))
            starts_with_table = any(
                box[1] <= top_pt + SPLIT_TOLERANCE_PT
                for box in table_boxes[index + 1])
            if ends_in_table and starts_with_table:
                splits.append({"page": index + 1, "next": index + 2})
                _issue(issues, "warning", "table-split",
                       "page %d ends with a table and page %d opens with one "
                       "(heuristic: verify against the two PNGs)" % (index + 1,
                                                                    index + 2))

    return {
        "page_count": count,
        "min_fill": min_fill,
        "content_width_mm": content_width_mm_value,
        "pages": pages,
        "sparse_pages": sparse,
        "blank_pages": blanks,
        "image_over_width": over,
        "table_split": splits,
        "orphan_last": orphans,
        "issues": issues,
    }


# --------------------------------------------------------------------------- CLI
def _parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        prog="docx_preview.py",
        description="Render a .docx as PDF + per-page PNG and lint the layout.")
    ap.add_argument("path", help="the .docx file to preview")
    ap.add_argument("--pdf", metavar="OUT.pdf",
                    help="write the converted PDF here (default: a temporary "
                         "file, deleted afterwards)")
    ap.add_argument("--pages", metavar="DIR",
                    help="write one page-%%03d.png per page into DIR")
    ap.add_argument("--dpi", type=int, default=PAGE_DPI,
                    help="PNG resolution (default %d)" % PAGE_DPI)
    ap.add_argument("--lint", action="store_true",
                    help="report the layout: fill ratio, half-empty pages, "
                         "blank pages, oversized images, split tables, "
                         "orphaned last lines")
    ap.add_argument("--min-fill", type=float, default=DEFAULT_MIN_FILL,
                    metavar="RATIO",
                    help="pages below this share of ink are 'sparse' "
                         "(default %.2f)" % DEFAULT_MIN_FILL)
    ap.add_argument("--max-orphan-chars", type=int, default=2,
                    metavar="N",
                    help="a last line of at most N characters is an orphan "
                         "(default 2)")
    ap.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT,
                    metavar="SECONDS",
                    help="soffice conversion timeout (default %d)"
                         % DEFAULT_TIMEOUT)
    ap.add_argument("--json", action="store_true",
                    help="print one JSON object on stdout (machine-readable)")
    ap.add_argument("--text", action="store_true",
                    help="print human-readable lines on stdout (the default)")
    return ap


def _usage_error(message: str) -> int:
    """Exit 2: one ``error:``-prefixed line on stderr, empty stdout."""
    print("error: " + message, file=sys.stderr)
    return 2


def _cannot_process(message: str) -> int:
    """Exit 1: the document is real, the environment or the conversion failed."""
    print("error: " + message, file=sys.stderr)
    return 1


def main(argv=None) -> int:
    ap = _parser()
    args = ap.parse_args(argv)

    # ---- refusals first: nothing below this line has read or written anything
    if args.json and args.text:
        return _usage_error("--json and --text are mutually exclusive")
    if not (args.pdf or args.pages or args.lint):
        return _usage_error("nothing to do: give at least one of "
                            "--pdf OUT.pdf, --pages DIR, --lint")
    if not 20 <= args.dpi <= 600:
        return _usage_error("--dpi must be between 20 and 600, got %d" % args.dpi)
    if not 0.0 <= args.min_fill <= 1.0:
        return _usage_error("--min-fill must be within 0..1, got %r"
                            % args.min_fill)
    if args.max_orphan_chars < 0:
        return _usage_error("--max-orphan-chars must be >= 0, got %d"
                            % args.max_orphan_chars)
    if args.timeout <= 0:
        return _usage_error("--timeout must be > 0, got %d" % args.timeout)

    try:
        _check_input(args.path)
        doc = _open_document(args.path)
    except InputError as exc:
        return _usage_error(str(exc))

    # ---- preconditions: PyMuPDF first, because every mode needs it (the PDF is
    # verified through it even when only --pdf was asked for).
    reason = py_mupdf_reason()
    if reason:
        return _cannot_process(reason)
    soffice, why = find_soffice()
    if soffice is None:
        return _cannot_process(why or "no LibreOffice found")

    work = Path(tempfile.mkdtemp(prefix="docx-preview-"))
    try:
        geometry = _page_geometry(doc) if args.lint else {}
        try:
            produced, seconds = convert_to_pdf(args.path, work, soffice,
                                               args.timeout)
        except PreviewError as exc:
            return _cannot_process(str(exc))

        # Verify before anything is moved to a user-supplied path: a PDF with
        # zero pages is a failed conversion, not a product.
        try:
            page_count = pdf_page_count(produced)
        except Exception as exc:
            return _cannot_process("cannot open the converted PDF: %s"
                                   % _one_line(exc))
        if page_count <= 0:
            return _cannot_process("the converted PDF has 0 pages")

        pdf_path = produced
        if args.pdf:
            target = Path(args.pdf)
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(produced, target)
            pdf_path = str(target)

        report: dict = {
            "ok": True,
            "input": str(Path(args.path).resolve()),
            "pdf": {
                "path": str(Path(pdf_path).resolve()) if args.pdf else None,
                "page_count": page_count,
                "bytes": Path(pdf_path).stat().st_size,
                "seconds": seconds,
                "kept": bool(args.pdf),
            },
        }

        if args.pages:
            try:
                paths, pages_seen, footers, truncated = render_pngs(
                    pdf_path, args.pages, args.dpi)
            except PreviewError as exc:
                return _cannot_process(str(exc))
            report["pages"] = {
                "dir": str(Path(args.pages).resolve()),
                "count": len(paths),
                "page_count": pages_seen,
                "dpi": args.dpi,
                "truncated": truncated,
                "files": paths,
                "footers": footers,
            }

        if args.lint:
            try:
                report["lint"] = lint_pdf(pdf_path, geometry=geometry,
                                          min_fill=args.min_fill,
                                          max_orphan_chars=args.max_orphan_chars)
            except PreviewError as exc:
                return _cannot_process(str(exc))
            except Exception as exc:
                return _cannot_process("layout lint failed: %s"
                                       % _one_line(exc))

        if args.json:
            print(json.dumps(report, ensure_ascii=False))
        else:
            _print_text(report)
        return 0
    finally:
        shutil.rmtree(work, ignore_errors=True)


def _print_text(report: dict) -> None:
    """Human-readable summary.  One fact per line, greppable, no decoration."""
    pdf = report.get("pdf") or {}
    if pdf.get("path"):
        print("pdf: %d page(s), %d bytes, %.2fs -> %s"
              % (pdf["page_count"], pdf["bytes"], pdf["seconds"], pdf["path"]))
    else:
        print("pdf: %d page(s) converted (not kept; pass --pdf to keep it)"
              % pdf.get("page_count", 0))
    pages = report.get("pages")
    if pages:
        print("pages: %d png(s), dpi=%d -> %s"
              % (pages["count"], pages["dpi"], pages["dir"]))
        if pages["truncated"]:
            print("pages: TRUNCATED at %d of %d page(s)" % (pages["count"],
                                                            pages["page_count"]))
    lint = report.get("lint")
    if lint:
        print("lint: %d issue(s), %d page(s), content_width_mm=%s"
              % (len(lint["issues"]), lint["page_count"],
                 lint["content_width_mm"]))
        print("lint: sparse_pages=%s blank_pages=%s"
              % (lint["sparse_pages"], lint["blank_pages"]))
        for issue in lint["issues"]:
            print("  %s %s %s" % (issue["severity"], issue["code"],
                                  issue["detail"]))


if __name__ == "__main__":
    sys.exit(main())
