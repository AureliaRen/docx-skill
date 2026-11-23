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

DEFAULT_MAX_TAIL_GAP = 60.0
"""``--max-tail-gap`` default, in mm: a page whose body ink stops more than
60 mm above the bottom of the text block is reported as a ``gap_page``.

60 mm is the value the P17 report asked for, and it was chosen against the real
corpus, not from taste: an A4 page with 25.4 mm margins has a 246 mm text block,
so 60 mm of tail white means the last quarter of the page is empty — visible to
a reader, and well past the few millimetres a normal last line leaves.  Measured
on the corpus in ``docs/reports/P17-B-LAYOUT.md``: full pages sat at 0-25 mm,
the genuinely half-used pages at 120-170 mm, so the threshold sits in a wide
valley.  The companion ``fill_ratio`` (a *ratio*) is deliberately left alone:
that one is scale-free and already has consumers.
"""

PT_PER_MM = 72.0 / 25.4
"""1 mm in PostScript points (PDF user units)."""

HEADER_TOLERANCE_PT = 1.0
"""How far **above** the text-block top a block must end to count as a header.

A block whose bottom edge lands on the text-block top is body text starting at
the margin; requiring a full point of clearance keeps a first line that
coincides with the margin (float rounding puts ``70.85 <= 70.85`` on a knife
edge) and still catches every real header, which sits in the margin itself --
the measured corpus header ended 12.15 pt above it.  Erring this way is safe:
a header that escapes the cut produces a *visible* negative ``head_gap_mm``,
while a wrongly deleted body line is silent.

(Note: the string above is documentation for a *reader*, not a real docstring —
these are floats, and a float has no ``__doc__`` slot.  Same convention as
``PT_PER_MM`` above: the literal sits directly under its assignment so the
constant and its rationale stay together.)
"""

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
        # ``encoding="utf-8"`` is mandatory, exactly as in
        # ``docx_template.py:_run_node``.  With bare ``text=True`` subprocess
        # decodes the child's pipes with ``locale.getpreferredencoding(False)``
        # -- cp936 on a Chinese Windows -- and soffice echoes the *path* it is
        # converting, so a Chinese directory name lands in the pipe as UTF-8
        # and the decode raises inside subprocess' own reader thread:
        #
        #     UnicodeDecodeError: 'gbk' codec can't decode byte 0xaf in
        #     position 46: illegal multibyte sequence
        #
        # The damage is worse than a crash: the exception kills the reader
        # thread, so ``_tail(proc)`` then sees an empty pipe, the real error
        # message is lost, and the process still exits 0 -- a false green.
        #
        # Measured on this machine (LibreOffice 25.8, cp936 console, UTF-8
        # mode off): soffice stdout and stderr are **UTF-8**, e.g.
        # ``b'...\\xe4\\xb8\\xad\\xe6\\x96\\x87\\xe7\\x9b\\xae\\xe5\\xbd\\x95...'``
        # decodes cleanly as UTF-8 and fails as cp936.  ``errors="replace"``
        # is the belt-and-braces rule: whatever a locale-localised soffice
        # build emits, this call must never raise UnicodeDecodeError.
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              encoding="utf-8", errors="replace",
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


def _body_blocks(page, header_band: bool = True,
                 body_top_pt: float | None = None) -> list[dict]:
    """Body text/image blocks, with the running header and footer removed.

    ``get_text("blocks")`` returns ``(x0, y0, x1, y1, text, block_no, type)``.
    A block that *starts* inside the footer band is footer content (page number,
    "第 3 页 共 8 页") and is dropped -- see :data:`FOOTER_BAND_START`.

    A block that ends **strictly above the text-block top** is a running header
    and is dropped too.  That cut is anchored on the document's own top margin
    (``body_top_pt``) and not on a fixed fraction of the page, because a
    fraction cannot tell a header from a body line: on A4 the 12% band ends at
    101 pt while a legitimate first body line starting at a 36 pt margin ends at
    92 pt, so a fraction-based cut would silently delete the first line of every
    tightly-margined document.  Measured, and pinned by
    ``test_a_body_line_near_the_top_is_not_mistaken_for_a_header``.

    When *body_top_pt* is None the header cut is skipped rather than guessed:
    the only safe fallback is the pre-P18 reading, and a missed header shows up
    as a negative ``head_gap_mm`` which is visible, whereas a deleted body line
    is not.

    *header_band* is a testing seam: pass ``False`` to reproduce the pre-P18
    reading and prove the cut is what changed it.
    """
    height = page.rect.height
    cutoff = height * FOOTER_BAND_START
    blocks = []
    for block in page.get_text("blocks"):
        x0, y0, x1, y1, text, _no, btype = block[:7]
        if y0 >= cutoff:
            continue
        if header_band and body_top_pt is not None \
                and y1 <= body_top_pt - HEADER_TOLERANCE_PT:
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


def _first_line(blocks) -> str:
    """The first text line of the page body, stripped.  ``""`` when there is none.

    The mirror of :func:`_last_line`: "first" is the smallest ``y0``, and within
    that block the first non-empty ``splitlines()`` entry.  Used to answer the
    question a bare gap measurement cannot — *is the following page starting a
    fresh paragraph, or is it the rest of something that got pushed over?*
    """
    text_blocks = [b for b in blocks if b["text"].strip()]
    if not text_blocks:
        return ""
    first = min(text_blocks, key=lambda b: b["y0"])
    lines = [line.strip() for line in first["text"].splitlines() if line.strip()]
    return lines[0] if lines else ""


def _text_bottom_mm(blocks) -> float | None:
    """Lowest ink on the page, in mm from the page's top edge.

    ``None`` for a page with no body blocks at all -- "nothing was printed here"
    is not the same reading as "ink ends at 0 mm", and a caller that folds them
    together reports a blank page as a huge tail gap.

    *Blocks* come from :func:`_body_blocks`, so footer content (page numbers)
    is already excluded; a page number at the bottom of an otherwise empty page
    must not count as body ink.
    """
    if not blocks:
        return None
    return round(max(b["y1"] for b in blocks) / PT_PER_MM, 2)


def _head_gap_mm(blocks, body_top_pt: float) -> float | None:
    """Blank between the top of the text block and the first line, in mm.

    Positive when the first line starts below the text-block top (normal), and
    **negative** when content starts above it — a paragraph with negative
    space-before, or a header word LibreOffice placed in the body layer.  The
    sign is kept rather than clamped: a negative value is evidence, not noise.
    """
    if not blocks:
        return None
    top = min(b["y0"] for b in blocks)
    return round((top - body_top_pt) / PT_PER_MM, 2)


def _tail_gap_mm(blocks, body_bottom_pt: float) -> float | None:
    """Blank between the lowest ink and the bottom of the text block, in mm.

    This is the number a person measures with a ruler when a page looks
    half-used: 8.3 cm of white paper above the bottom margin.  ``None`` when the
    page has no body blocks — see :func:`_text_bottom_mm`.

    Can be **negative** when content overflows the bottom margin (a table that
    would not split, a figure pushed down); kept signed for the same reason as
    :func:`_head_gap_mm`.

    Units: ``body_bottom_pt`` is points (say, 720.0 for a 792 pt page with a
    25.4 mm bottom margin) and the result is millimetres, so the conversion
    happens once, here.  Doing it at the call site mixes a millimetre reading
    with a point reading and flips the sign — which is exactly how the first
    version of this function reported a full page as ``-170 mm``.
    """
    bottom = _text_bottom_mm(blocks)
    if bottom is None:
        return None
    return round(body_bottom_pt / PT_PER_MM - bottom, 2)


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
    # Flag the one case the approximation is actually wrong for.  `_page_geometry`
    # returns the FIRST section because the PDF gives no page→section map, and
    # that is fine while every section shares the same margins: the header cut
    # and the text-block bottom are then the same lines on every page.  Once
    # margins differ, `head_gap_mm` on a later section is measured against the
    # wrong text-block top and a running header can read as body ink.  Reporting
    # the condition is honest; silently approximating is what got missed.
    try:
        margins = set()
        for s in doc.sections:
            m = (getattr(s, "top_margin", None),
                 getattr(s, "bottom_margin", None),
                 getattr(s, "left_margin", None),
                 getattr(s, "right_margin", None))
            margins.add(tuple(None if v is None else round(v.mm, 2)
                              for v in m))
        geometry["sections"] = len(doc.sections)
        geometry["section_margins_differ"] = len(margins) > 1
    except Exception:
        geometry["sections"] = None
        geometry["section_margins_differ"] = None
    return geometry


def lint_pdf(pdf_path, geometry=None, min_fill: float = DEFAULT_MIN_FILL,
             max_orphan_chars: int = 2, gaps: bool = False,
             max_tail_gap: float = DEFAULT_MAX_TAIL_GAP,
             tail_gap_skip_last: bool = False,
             header_band: bool = True) -> dict:
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

    ``gaps=True`` additionally reports **absolute** lengths, which
    ``fill_ratio`` cannot express (a ratio says "this page is 30% full"; the
    question that gets asked is "how many centimetres of paper are blank"):

    ``text_bottom_mm`` / ``tail_gap_mm`` / ``head_gap_mm``
        Per page: the lowest ink measured from the page's top edge, the blank
        between that ink and the bottom of the text block, and the blank
        between the top of the text block and the first line.  All in mm, all
        signed where a page overflows its margins.  ``None`` -- never ``0`` --
        for a page with no body blocks, because "no ink" and "ink at 0 mm" are
        different readings.
    ``next_page_lead``
        The first 40 characters of the next page's first body line.  This is
        the *association* a bare number cannot give: a page with 120 mm of tail
        white whose successor opens mid-sentence is content that ran out, while
        one whose successor opens with a fresh heading (or a table) is a block
        that got pushed whole onto the next page.  Same reading, different fix.
    ``gap_pages`` / ``tail-gap`` issues
        Pages whose ``tail_gap_mm`` exceeds ``max_tail_gap``.  Thresholded on an
        absolute length on purpose: ``fill_ratio`` is scale-free and therefore
        cannot say "8.3 cm", which is the number a reader actually measures.

    Severity: ``blank-page`` and ``image-over-width`` are ``error`` (a reader
    sees a broken page); ``sparse-page``, ``table-split``, ``orphan-last``,
    ``tail-gap`` and ``content-width-unknown`` are ``warning``.

    **Every key added by** ``gaps=True`` **is absent when it is** ``False``, so a
    caller that already parses a ``--lint`` report sees byte-identical output.
    That is why the two switches are separate rather than one flag: ``--lint``
    has consumers (this repository's own round reports among them), and
    silently widening its schema would break them for no gain.
    """
    fitz = _fitz()
    geometry = geometry or {}
    content_width_mm_value = geometry.get("content_width_mm")
    content_width_pt = (float(content_width_mm_value) * PT_PER_MM
                        if content_width_mm_value else None)
    top_margin_mm = float(geometry.get("top_margin_mm")
                          or DEFAULT_MARGIN_MM)
    top_pt = top_margin_mm * PT_PER_MM
    bottom_margin_mm = float(geometry.get("bottom_margin_mm")
                             or DEFAULT_MARGIN_MM)
    # Only the top margin is needed by `table_split`, which asks "does the next
    # page open with a table" -- a top-edge question.  The bottom margin is not
    # used there: LibreOffice stops a split table wherever its last row ends,
    # measured 11 pt short of the bottom margin on a real split.  `gaps` does
    # need it, and for a different question: where the *text block* bottoms out
    # (page height - bottom margin), which is the line a person holds a ruler
    # against when they say "8.3 cm of blank paper".

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
    gap_pages: list[dict] = []
    first_lines: list[str] = []

    if content_width_pt is None:
        _issue(issues, "warning", "content-width-unknown",
               "the document reports no section page width, so image widths "
               "were not checked against the text block")

    # Geometry is read from the first section only (see `_page_geometry`); once
    # sections disagree on margins that approximation affects `head_gap_mm` and
    # the text-block bottom on later sections.  Say so instead of letting a
    # plausible number stand for a measurement it is not.
    if geometry.get("section_margins_differ"):
        _issue(issues, "warning", "section-margins-differ",
               "sections %s have different margins, but the layout geometry is "
               "taken from section 1; head_gap_mm and tail_gap_mm on later "
               "sections may be measured against the wrong text-block edges"
               % geometry.get("sections", "?"))

    with fitz.open(str(pdf_path)) as pdf:
        count = pdf.page_count
        for number, page in enumerate(pdf, 1):
            height = page.rect.height
            heights.append(height)
            blocks = _body_blocks(page, header_band=header_band,
                                  body_top_pt=top_pt)
            images = _image_boxes(page)
            ratio = _fill_ratio(blocks, height)
            ink = _ink_ratio(blocks, height)
            text = "".join(b["text"] for b in blocks)
            chars = len(text.strip())
            last_line = _last_line(blocks)
            boxes = _table_boxes(page)
            table_boxes.append(boxes)
            lowest_bottom.append(max((b["y1"] for b in blocks), default=None))

            text_bottom = _text_bottom_mm(blocks) if gaps else None
            head_gap = (_head_gap_mm(blocks, top_pt) if gaps else None)
            tail_gap = (_tail_gap_mm(blocks, height - bottom_margin_mm * PT_PER_MM)
                        if gaps else None)
            if gaps:
                first_lines.append(_first_line(blocks))

            is_blank = chars == 0 and not images
            page_row = {
                "page": number,
                "fill_ratio": ratio,
                "ink_ratio": ink,
                "text_chars": chars,
                "block_count": len(blocks),
                "image_count": len(images),
                "last_line": last_line,
                "blank": is_blank,
            }
            if gaps:
                page_row["text_bottom_mm"] = text_bottom
                page_row["tail_gap_mm"] = tail_gap
                page_row["head_gap_mm"] = head_gap
            pages.append(page_row)

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

            # Absolute tail white.  Kept apart from `sparse-page` on purpose:
            # a page can be genuinely full in ratio terms and still leave a big
            # tail if the text starts far down the page, and a two-line page
            # leaves a huge tail while scoring a tiny ratio -- the two numbers
            # answer different questions, so both are reported and neither
            # suppresses the other.  A blank page is excluded: its tail gap is
            # the whole text block, which is already reported as `blank-page`
            # and would otherwise double-count the same defect.
            if gaps and not is_blank and tail_gap is not None \
                    and tail_gap > max_tail_gap \
                    and not (tail_gap_skip_last and number == count):
                gap_pages.append({"page": number, "tail_gap_mm": tail_gap,
                                  "text_bottom_mm": text_bottom})
                _issue(issues, "warning", "tail-gap",
                       "page %d: body ink stops at %.1f mm from the page top, "
                       "leaving %.1f mm blank above the bottom margin "
                       "(> %.1f mm)"
                       % (number, text_bottom or 0.0, tail_gap, max_tail_gap))

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

    # Cross-page association for `gaps`: attach each page's successor lead line.
    # Done after the loop because the lead of page n is page n+1's first line,
    # which is not known while page n is being summarised.
    if gaps:
        for index, page_row in enumerate(pages):
            if index + 1 < len(first_lines):
                lead = first_lines[index + 1]
                page_row["next_page_lead"] = lead[:40] if lead else ""
            else:
                page_row["next_page_lead"] = None   # last page: no successor

    result = {
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
    if gaps:
        # Only present when asked for -- see the docstring: `--lint` already has
        # consumers and its schema must not grow under them.
        result["gap_pages"] = gap_pages
        result["max_tail_gap_mm"] = max_tail_gap
        result["bottom_margin_mm"] = bottom_margin_mm
        # `head_gap_mm` is measured against the top of the text block, so the
        # top margin is what lets a caller check the number instead of trusting
        # it: first_line_y = top_margin_mm + head_gap_mm.  Without it the only
        # way to reconcile a surprising `head_gap_mm` was to re-render the page,
        # which is exactly the work `--gaps` exists to remove.
        result["top_margin_mm"] = top_margin_mm
    return result


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
    ap.add_argument("--gaps", action="store_true",
                    help="also report absolute page-tail lengths in mm "
                         "(text_bottom_mm / tail_gap_mm / head_gap_mm), the "
                         "next page's opening line, and pages whose tail white "
                         "exceeds --max-tail-gap")
    ap.add_argument("--max-tail-gap", type=float, default=DEFAULT_MAX_TAIL_GAP,
                    metavar="MM",
                    help="a page whose body ink stops more than MM above the "
                         "bottom of the text block is a gap_page "
                         "(default %.0f)" % DEFAULT_MAX_TAIL_GAP)
    ap.add_argument("--skip-last-page-gap", action="store_true",
                    help="do not report the final page's tail white as a "
                         "gap_page.  A document ends where its last sentence "
                         "ends, so the last page is half empty by definition; "
                         "whether that counts as a defect is the caller's "
                         "call, which is why it is a flag and not the default")
    ap.add_argument("--no-header-band", action="store_true",
                    help="count running headers as body ink (pre-P18 reading). "
                         "Only useful to reproduce the old negative "
                         "head_gap_mm; the default excludes them")
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
    # `--gaps` counts as a capability in its own right: it produces the whole
    # lint report (the gap fields are extra keys on it), so demanding `--lint`
    # alongside it would be asking for the same flag twice.  It also *implies*
    # --lint's work, so `args.lint` is set below rather than making `--gaps`
    # alone emit a reduced report that no caller expects.
    if args.gaps:
        args.lint = True
    if not (args.pdf or args.pages or args.lint or args.gaps):
        return _usage_error("nothing to do: give at least one of "
                            "--pdf OUT.pdf, --pages DIR, --lint, --gaps")
    if not 20 <= args.dpi <= 600:
        return _usage_error("--dpi must be between 20 and 600, got %d" % args.dpi)
    if not 0.0 <= args.min_fill <= 1.0:
        return _usage_error("--min-fill must be within 0..1, got %r"
                            % args.min_fill)
    if args.max_orphan_chars < 0:
        return _usage_error("--max-orphan-chars must be >= 0, got %d"
                            % args.max_orphan_chars)
    if args.max_tail_gap < 0:
        return _usage_error("--max-tail-gap must be >= 0, got %r"
                            % args.max_tail_gap)
    if args.timeout <= 0:
        return _usage_error("--timeout must be > 0, got %d" % args.timeout)

    try:
        _check_input(args.path)
        doc = _open_document(args.path)
    except InputError as exc:
        return _usage_error(str(exc))

    # ---- preconditions, in the order the work happens: convert first, then
    # read the PDF back. The order is load-bearing, not cosmetic. Asking for
    # PyMuPDF first reported the *second* blocker on a machine that had
    # neither, so the one line the user got named the wrong thing to install
    # (a GitHub runner has no LibreOffice and no PyMuPDF; it said "no fitz").
    # Whichever comes first is the one they must fix first, so it has to be
    # the one the message names.
    soffice, why = find_soffice()
    if soffice is None:
        return _cannot_process(why or "no LibreOffice found")
    reason = py_mupdf_reason()
    if reason:
        return _cannot_process(reason)

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
                                          max_orphan_chars=args.max_orphan_chars,
                                          gaps=args.gaps,
                                          max_tail_gap=args.max_tail_gap,
                                          tail_gap_skip_last=args.skip_last_page_gap,
                                          header_band=not args.no_header_band)
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
        # `gap_pages` only exists when --gaps was passed; printing an empty
        # list for a --lint run would imply the check ran and found nothing.
        if "gap_pages" in lint:
            print("gaps: max_tail_gap_mm=%.1f gap_pages=%s"
                  % (lint["max_tail_gap_mm"],
                     [row["page"] for row in lint["gap_pages"]]))
            for row in lint["pages"]:
                tail = row.get("tail_gap_mm")
                print("  page %d: text_bottom_mm=%s tail_gap_mm=%s"
                      " head_gap_mm=%s next_page_lead=%r"
                      % (row["page"], row.get("text_bottom_mm"), tail,
                         row.get("head_gap_mm"), row.get("next_page_lead", "")))
        for issue in lint["issues"]:
            print("  %s %s %s" % (issue["severity"], issue["code"],
                                  issue["detail"]))


if __name__ == "__main__":
    sys.exit(main())
