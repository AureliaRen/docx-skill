#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Edit an existing .docx in place (or to a new file).

Subcommands:
  replace   find-and-replace text, preserving run formatting
  set-cell  set the text of a table cell
  insert    insert a paragraph before a given body paragraph index
  delete    delete a body paragraph by index
  style     apply a paragraph style to a body paragraph by index
  normalize merge adjacent runs with identical formatting
  toc       insert a Table of Contents field at a body paragraph index
  page-numbers  add "Page X of Y" (PAGE/NUMPAGES fields) to the footer;
                idempotent - a footer that already carries a PAGE/NUMPAGES
                field is left byte-for-byte alone, and an in-place re-run
                writes nothing at all
  char-format   set character formatting (highlight/colour/bold/italic) on
                the runs that carry some anchor text; runs are split when the
                anchor is only part of them, so untouched text keeps its format

Examples:
  docx_edit.py replace in.docx --find old --replace new -o out.docx
  docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text "42"
  docx_edit.py insert in.docx --index 3 --text "New para" --style Normal
  docx_edit.py delete in.docx --index 3
  docx_edit.py style in.docx --index 0 --style "Heading 1"
  docx_edit.py normalize in.docx -o out.docx
  docx_edit.py toc in.docx --index 1 -o out.docx
  docx_edit.py page-numbers in.docx -o out.docx
  docx_edit.py char-format in.docx --find "见附录" --highlight darkMagenta
  docx_edit.py char-format in.docx --find "重点" --highlight DARK_MAGENTA --color FF0000
  docx_edit.py char-format in.docx --find "关键词" --bold --italic -o out.docx

Field results (TOC entries, page numbers) are computed by Word or
LibreOffice when the document is opened, not by python-docx; until then
the fields show placeholder text.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import copy
import json
import os
import re
import sys

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_common import iter_all_paragraphs, replace_in_paragraph

# Input validation is **not** reimplemented here. `docx_read` owns the single
# definition (`InputError` / `_check_input` / `_open_document`) so the two CLIs
# cannot drift apart — `docx_read` already learned the fourth case the hard way
# (a ZIP with no `[Content_Types].xml` dies on a bare `KeyError`). Reusing the
# underscore names is the established shape in this repo, not a new liberty:
# `docx_toc.py` imports `docx_create._apply_tab_stops` /
# `docx_create._bookmark_para` the same way. Both halves are in `scripts/`,
# which `_bootstrap` already put on `sys.path`.
from docx_read import InputError, _check_input, _open_document
# Run-level inspection is **also** taken from `docx_read`, not reimplemented:
# `--runs` has to report the same `vert_align` / `bold` / `italic` / `size_pt` /
# `font` that this file then decides to keep or drop. Two copies of that reader
# would drift, and the drift would show up as `--runs` promising "single run,
# safe to edit" for a cell `set-cell` then treats as mixed.
from docx_read import describe_run, is_mixed_format

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def cmd_replace(doc, args) -> dict:
    n = 0
    for para in iter_all_paragraphs(doc):
        n += replace_in_paragraph(para, args.find, args.replace)
    return {"replacements": n}


def _check_range(what: str, value: int, n: int) -> None:
    """Refuse a ``value`` that does not name a real element of ``n`` items.

    ``what`` is the ``--option`` name, so the message can say which of
    ``--table`` / ``--row`` / ``--col`` / ``--index`` the caller got wrong.
    ``set-cell`` needs that distinction: one call indexes three dimensions, and
    a message that does not name the bad one forces the caller to guess which
    argument to fix.

    Negative values are rejected rather than read through Python's negative
    indexing. ``-1`` is not "the last paragraph" in any durable sense — add or
    delete one paragraph and the same argument means a different one — so a
    caller that wrote ``-1`` has a bug, and echoing success hides it.
    """
    if n == 0:
        raise IndexOutOfRangeError(
            f"{what} {value} is out of range: there is nothing to index, "
            f"so no index is legal (current N=0)")
    if not 0 <= value <= n - 1:
        raise IndexOutOfRangeError(
            f"{what} {value} is out of range: "
            f"legal range is [0, {n - 1}], current N={n}")


def _writable_run(run) -> bool:
    """True for a run a flat string can be written into.

    A field run (``w:fldChar`` / ``w:instrText``) and a drawing run are not
    plain text carriers: overwriting one deletes a field or a picture, which the
    caller never asked for. ``docx_read --runs`` reports the same two flags, so
    the two CLIs agree on what counts as writable text.
    """
    r_el = run._r
    for tag in ("fldChar", "instrText", "drawing", "pict"):
        if r_el.find(_q(tag)) is not None:
            return False
    return True


def _strip_vert_align(run):
    """Drop ``w:rPr/w:vertAlign`` from one run; return its previous value.

    ``None`` when the run stated nothing. Only called on the *mixed* path: with
    one surviving run over several disagreeing ones there is no way to know
    which characters were meant to be subscript, and carrying the first run's
    ``vertAlign`` onto the whole new string is exactly the bug this rewrite
    exists to stop (a first run of ``F`` with ``vertAlign=subscript`` turned a
    whole ``FS`` into subscript).
    """
    rpr = run._r.find(_q("rPr"))
    if rpr is None:
        return None
    el = rpr.find(_q("vertAlign"))
    if el is None:
        return None
    val = el.get(_q("val"))
    rpr.remove(el)
    return val


def _fill_cell_text(cell, text: str, where: str) -> list[str]:
    """Overwrite one cell's text, keeping every format a flat string allows.

    Layered on purpose, because the caller hands over a **flat string** and a
    flat string does not say which character is a subscript. Guessing would be
    making the caller's decision for them, badly:

    1. **Paragraph level is always kept.** Alignment, spacing and the paragraph
       style live in ``w:pPr`` and are not touched at all. This is the layer
       that can be preserved losslessly in every case.
    2. **One text run is lossless.** Reassigning ``run.text`` goes through
       ``CT_R.clear_content()``, which removes the run's content but *keeps*
       ``w:rPr`` -- bold, size, font and even ``vertAlign`` survive untouched.
    3. **Several text runs cannot be mapped from a flat string.** The first
       writable run is kept as the formatting carrier, its ``vertAlign`` is
       stripped when the runs disagree (see :func:`_strip_vert_align`), the
       others are removed, and ``warnings`` says so out loud.
    4. **Nothing is silent.** Every lossy outcome appends one ``warning: ...``
       line; a one-run cell appends nothing, so the warning stays a signal.

    Returns the warning lines (already prefixed); ``cmd_set_cell`` prints them
    to stderr so stderr and the JSON agree.
    """
    notes: list[str] = []
    paragraphs = list(cell.paragraphs)
    if not paragraphs:
        # A `w:tc` with no `w:p` at all is malformed but openable; give it the
        # paragraph the caller's text needs rather than IndexError.
        paragraphs = [cell.add_paragraph()]
    target, extras = paragraphs[0], paragraphs[1:]

    writable = [r for r in target.runs if _writable_run(r)]
    text_runs = [r for r in writable if r.text]

    if len(text_runs) <= 1:
        # Lossless path. Reuse the run that already carries text, else the
        # first empty formatting carrier, else create one.
        run = text_runs[0] if text_runs else (writable[0] if writable else None)
        if run is None:
            kept = len(target.runs)
            target.add_run(text)
            if kept:
                # The cell's only runs are fields or drawings. Appending keeps
                # them (destroying a field or a picture to satisfy a text write
                # is not a trade this command should make silently), but the
                # cell's resulting text is old+new, which the caller must know.
                notes.append(
                    f"warning: {where}: the paragraph held {kept} run(s) that "
                    f"are fields or drawings; they were kept and the new text "
                    f"was appended after them, so the cell now reads "
                    f"{target.text!r}. Use --replace-cell to drop them.")
        else:
            run.text = text
    else:
        first = text_runs[0]
        info = [describe_run(r) for r in text_runs]
        if is_mixed_format(info):
            dropped = _strip_vert_align(first)
            tail = (f" Its vertAlign={dropped!r} was dropped: a flat string "
                    f"cannot say which characters were subscript/superscript."
                    if dropped else "")
            notes.append(
                f"warning: {where}: cell held {len(text_runs)} text runs with "
                f"different run formatting, so the whole new string now uses "
                f"the first run's formatting; paragraph formatting "
                f"(alignment/spacing/style) was kept, the other "
                f"{len(text_runs) - 1} run(s) were removed.{tail} "
                f"Edit the runs one by one, or pass --replace-cell to rewrite "
                f"the cell.")
        first.text = text
        for run in text_runs[1:]:
            run._r.getparent().remove(run._r)

    if extras:
        # Only the first paragraph can hold a single flat string. Collapsing the
        # rest is the one structural loss this path makes, so it is announced.
        nested = bool(cell.tables)
        for para in extras:
            if nested:
                # Deleting paragraphs around a nested table would re-order the
                # cell's blocks; empty their text and leave the `w:p` in place.
                for run in [r for r in para.runs if _writable_run(r)]:
                    run._r.getparent().remove(run._r)
            else:
                para._p.getparent().remove(para._p)
        notes.append(
            f"warning: {where}: cell held {len(paragraphs)} paragraphs and now "
            f"holds 1; the first paragraph's formatting was kept and the "
            f"{'other paragraphs were emptied' if nested else 'rest were removed'}"
            f" (a flat --text is a single paragraph).")
    return notes


def cmd_set_cell(doc, args) -> dict:
    """Set one table cell's text, refusing any index that names nothing.

    All three dimensions are checked, in order, before a cell is touched; the
    message names the dimension that was wrong together with its own legal
    range, so ``--row 9`` never reports the column count.

    Column count comes from the row itself rather than ``table.columns``: a
    cell grid can be ragged, and ``table.cell(row, col)`` indexes the row.

    Two modes:

    * default -- **preserve**. Paragraph formatting is kept, a single text run
      is rewritten in place (its ``w:rPr`` survives), and a multi-run cell
      inherits the first run's formatting with ``vertAlign`` stripped plus one
      ``warning:`` line on stderr. See :func:`_fill_cell_text`.
    * ``--replace-cell`` -- the old ``cell.text = text``: clear the whole cell
      and write one paragraph with one run. Kept for the legitimate "I really
      do want this cell emptied and rewritten" case, and it is the escape hatch
      the warning points at.

    Exit code is unchanged either way: a warning is not a failure.
    """
    tables = doc.tables
    _check_range("set-cell --table", args.table, len(tables))
    table = tables[args.table]
    rows = table.rows
    _check_range("set-cell --row", args.row, len(rows))
    _check_range("set-cell --col", args.col, len(rows[args.row].cells))

    where = f"table {args.table} row {args.row} col {args.col}"
    cell = table.cell(args.row, args.col)
    if getattr(args, "replace_cell", False):
        cell.text = args.text
        return {"table": args.table, "row": args.row, "col": args.col,
                "mode": "replace-cell", "warnings": []}

    notes = _fill_cell_text(cell, args.text, where)
    for line in notes:
        print(line, file=sys.stderr)
    return {"table": args.table, "row": args.row, "col": args.col,
            "mode": "preserve", "warnings": notes}


class IndexOutOfRangeError(ValueError):
    """A `--index` / `--table` / `--row` / `--col` that names nothing real.

    Raised *before* the document is touched, so `main` can refuse with exit 2
    and nothing is written — the file is exactly as it was.

    Covers ``delete``, ``style`` and ``set-cell``. ``insert`` keeps its own
    subclass (:class:`InsertIndexError`) because its message has to explain an
    extra rule — an empty body still has one legal insertion point — but both
    are caught by the same handler in `main`, so every one of these becomes
    ``exit 2`` plus one line on stderr.
    """


class InsertIndexError(IndexOutOfRangeError):
    """`insert --index` cannot be honoured honestly.

    Raised *before* the document is touched, so `main` can refuse with exit 2
    and nothing is written — the file is exactly as it was.
    """


class CharFormatError(ValueError):
    """`char-format` was given something it cannot write as asked.

    Two shapes, both refused **before** the file is saved, so `main` returns
    exit 2 with one ``error:`` line and the input is byte-for-byte untouched:

    * **a value that is not legal** — an unknown highlight colour, a colour that
      is not six hex digits, or a call that asks for no format at all.  Exit 2,
      because the caller named an argument value that does not exist (the same
      class as a bad ``--index``; see D-024 / D-030).
    * **an anchor that selects nothing** — ``--find`` matched no run.  Exit 2 as
      well, and for the same reason: the argument names something that is not
      there.  Exit 1 is reserved for "the document is real but has a problem"
      (`docx_validate`'s territory), and a text anchor this command cannot find
      says nothing about the document's health.  Reporting 0 here would be the
      worst outcome: the caller asked for a visible change, got none, and was
      told it worked.
    """


def _apply_char_format_args(ap) -> None:
    """The ``char-format`` option block, kept next to its handler."""
    ap.add_argument("--find", required=True, metavar="TEXT",
                    help="anchor text; only the runs carrying it are changed "
                         "(runs are split when the anchor is part of one)")
    ap.add_argument("--highlight", metavar="COLOR",
                    help="w:highlight value: darkMagenta / DARK_MAGENTA / "
                         "dark_magenta all work; also none (remove) and "
                         "inherit (leave the element out). Legal set: "
                         + ", ".join(HIGHLIGHT_VALUES))
    ap.add_argument("--color", metavar="RRGGBB",
                    help="font colour, six hex digits (FF0000 or #00a0ff) or "
                         "'auto'; a theme colour on the run is replaced, since "
                         "a theme attribute would win over the RGB value")
    ap.add_argument("--bold", dest="bold", action="store_true", default=None,
                    help="set bold on")
    ap.add_argument("--no-bold", dest="bold", action="store_false",
                    help="set bold explicitly off (w:val=\"0\", not absent)")
    ap.add_argument("--italic", dest="italic", action="store_true", default=None,
                    help="set italic on")
    ap.add_argument("--no-italic", dest="italic", action="store_false",
                    help="set italic explicitly off")


def _body_index_of(doc, p_el) -> int | None:
    """Position of the paragraph element ``p_el`` in ``doc.paragraphs``.

    Identity (`is`), not text equality: two paragraphs may legitimately hold
    the same text, and the point of this lookup is to report the position of
    *the object we just inserted*, not of the first paragraph that looks like
    it. Returns None when the element is not a body paragraph at all.
    """
    for i, para in enumerate(doc.paragraphs):
        if para._p is p_el:
            return i
    return None


def _check_insert_index(index: int, n: int) -> None:
    """Refuse an ``--index`` that does not name a real insertion point.

    ``index`` is a 0-based position in ``doc.paragraphs`` (body only — table
    cell paragraphs are not counted), and must name a paragraph that exists:
    the new paragraph is inserted *before* it, so it lands at exactly that
    index. Anything else is a caller bug, and clamping it silently is worse
    than failing, because the caller is told a position it will never get:

    * ``--index 99`` on a 4-paragraph body used to fall through to
      ``add_paragraph`` and land at index 4 while reporting ``inserted_at: 99``;
    * ``--index -1`` used to hit Python's negative indexing, insert before the
      *last* paragraph (landing at N-1) and echo ``-1`` back.

    An empty body (N=0) is the one degenerate case with a single valid
    insertion point: 0, which appends the document's first paragraph.
    """
    if n == 0:
        if index != 0:
            raise InsertIndexError(
                f"insert --index {index} is out of range: the body has no "
                f"paragraphs, so the only valid index is [0, 0] (current N=0)")
        return
    _check_range("insert --index", index, n)


def cmd_insert(doc, args) -> dict:
    """Insert a paragraph at body index ``index`` and report where it landed.

    ``inserted_at`` in the returned JSON is the paragraph's **measured**
    position after the insert (looked up by element identity), never an echo
    of the argument: an echoed argument is precisely how the old code
    reported ``99`` for a paragraph that had been appended at index 4.
    """
    paras = doc.paragraphs
    n = len(paras)
    _check_insert_index(args.index, n)

    if n == 0:
        new_para = doc.add_paragraph(args.text, style=args.style)
    else:
        new_para = paras[args.index].insert_paragraph_before(
            args.text, style=args.style)

    landed = _body_index_of(doc, new_para._p)
    if landed is None:
        raise InsertIndexError(
            "the paragraph was inserted but is not reachable as a body "
            "paragraph, so its position cannot be reported")
    return {"inserted_at": landed, "text": new_para.text}


def cmd_delete(doc, args) -> dict:
    """Delete one body paragraph, refusing an ``--index`` that names nothing.

    The check runs *before* the element is touched, so a refused call leaves
    the document object — and therefore the file on disk — exactly as it was.
    """
    paras = doc.paragraphs
    _check_range("delete --index", args.index, len(paras))
    el = paras[args.index]._element
    el.getparent().remove(el)
    return {"deleted_index": args.index}


def cmd_style(doc, args) -> dict:
    """Apply a paragraph style, refusing an ``--index`` that names nothing.

    The check runs *before* ``doc.styles[...]`` is consulted, so an
    out-of-range index is reported as such rather than being masked by a
    ``KeyError`` on a style name the caller never got to.
    """
    paras = doc.paragraphs
    _check_range("style --index", args.index, len(paras))
    paras[args.index].style = doc.styles[args.style]
    return {"index": args.index, "style": args.style}


def _run_format_key(r_el) -> str:
    """Canonical string for a run's w:rPr (None when absent)."""
    from lxml import etree
    rpr = r_el.find(_q("rPr"))
    return "" if rpr is None else etree.tostring(rpr).decode("utf-8")


def cmd_normalize(doc) -> dict:
    """Merge adjacent sibling runs with identical formatting."""
    merged = 0
    for para in iter_all_paragraphs(doc):
        prev = None
        for r_el in list(para._p):
            if r_el.tag != _q("r"):
                prev = None
                continue
            # only merge plain-text runs (no breaks, tabs, drawings...)
            kids = {c.tag for c in r_el} - {_q("rPr"), _q("t")}
            if kids:
                prev = None
                continue
            if (prev is not None
                    and _run_format_key(prev) == _run_format_key(r_el)):
                pt = prev.find(_q("t"))
                ct = r_el.find(_q("t"))
                if pt is None:
                    pt = prev.makeelement(_q("t"), {})
                    prev.append(pt)
                pt.text = (pt.text or "") + ((ct.text or "")
                                             if ct is not None else "")
                pt.set("{http://www.w3.org/XML/1998/namespace}space",
                       "preserve")
                r_el.getparent().remove(r_el)
                merged += 1
            else:
                prev = r_el
    return {"runs_merged": merged}


def _add_field(para, instr: str, placeholder: str) -> None:
    """Append a complex field (begin/instrText/separate/result/end)."""
    p = para._p
    for ftype, extra in (("begin", None), (None, instr),
                         ("separate", None), (None, placeholder),
                         ("end", None)):
        r = p.makeelement(_q("r"), {})
        p.append(r)
        if ftype is not None:
            fld = r.makeelement(_q("fldChar"), {_q("fldCharType"): ftype})
            r.append(fld)
        elif extra is instr:
            it = r.makeelement(_q("instrText"), {})
            it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            it.text = instr
            r.append(it)
        else:
            t = r.makeelement(_q("t"), {})
            t.text = extra
            r.append(t)


def _check_toc_index(index: int, n: int) -> None:
    """Refuse a ``toc --index`` that does not name a real insertion point.

    Same contract as ``delete`` / ``style`` / ``insert``: the legal range is
    ``0 <= index <= n - 1``, because the TOC paragraph is inserted *before*
    paragraph ``index`` and therefore lands at exactly that index. Two silent
    fallbacks used to hide caller bugs here:

    * ``--index -1`` hit Python's negative indexing, inserted before the *last*
      paragraph and echoed ``-1`` back as ``toc_inserted_at``;
    * ``--index >= N`` fell through to ``doc.add_paragraph`` and appended to the
      end, while still reporting ``99`` — a position the caller never got.

    An empty body (N=0) is the one degenerate case with a single valid
    insertion point: 0, which appends the document's first paragraph. Local to
    ``toc`` rather than folded into ``_check_range`` because that helper is
    shared with ``delete`` / ``style`` / ``set-cell``, where N=0 has no legal
    index at all.
    """
    if n == 0:
        if index != 0:
            raise IndexOutOfRangeError(
                f"toc --index {index} is out of range: the body has no "
                f"paragraphs, so the only valid index is [0, 0] (current N=0)")
        return
    _check_range("toc --index", index, n)


def cmd_toc(doc, args) -> dict:
    """Insert a TOC field before body paragraph ``index``.

    The index is checked *before* any element is touched, so a refused call
    leaves the document — and therefore the file on disk — exactly as it was.
    """
    paras = doc.paragraphs
    n = len(paras)
    _check_toc_index(args.index, n)
    if n == 0:
        para = doc.add_paragraph("")
    else:
        para = paras[args.index].insert_paragraph_before("")
    _add_field(para, r' TOC \o "1-3" \h \z \u ',
               "Table of contents - open in Word/LibreOffice and update "
               "fields to populate.")
    return {"toc_inserted_at": args.index}


PAGE_FIELD_CODES = ("PAGE", "NUMPAGES")


def _field_codes(p_el) -> set:
    """Field codes a footer paragraph already carries (`PAGE`, `NUMPAGES`).

    Both spellings count. python-docx only ever writes the *complex* form
    (``fldChar begin -> instrText -> separate -> result -> end``), but Word
    and LibreOffice both also emit the *simple* form
    (``<w:fldSimple w:instr=" PAGE "/>``), so a footer built by either tool
    has to be recognised too — otherwise a second page number gets stapled
    onto one that was already there.

    Field codes may also be split across several ``w:instrText`` runs, so the
    texts are joined before tokenising rather than tested one run at a time.
    """
    code = "".join(it.text or "" for it in p_el.iter(_q("instrText")))
    code += " " + " ".join(fs.get(_q("instr"), "") or ""
                           for fs in p_el.iter(_q("fldSimple")))
    return {tok for tok in code.upper().split() if tok in PAGE_FIELD_CODES}


def cmd_page_numbers(doc, args) -> dict:
    """Append `Page X of Y` to the footer — once, and without losing anything.

    Two behaviours are load-bearing here:

    * **Idempotent.** If the footer already carries a ``PAGE``/``NUMPAGES``
      field the document is reported ``changed: False`` and (for an in-place
      run) `main` skips `save()` entirely, so plan→apply→apply leaves the
      file untouched instead of stacking a second "Page 1 of 1".
    * **Append, never replace.** The existing runs of the footer paragraph are
      left in place; the new runs go after them. A footer reading
      ``CONFIDENTIAL`` becomes ``CONFIDENTIALPage 1 of 1``, not
      ``Page 1 of 1``.
    """
    footer = doc.sections[0].footer
    present = set()
    for para in footer.paragraphs:
        present |= _field_codes(para._p)
    if present:
        return {"footer_fields": list(PAGE_FIELD_CODES), "changed": False,
                "already_present": sorted(present)}
    para = footer.paragraphs[0] if footer.paragraphs \
        else footer.add_paragraph()
    para.add_run("Page ")
    _add_field(para, " PAGE ", "1")
    para.add_run(" of ")
    _add_field(para, " NUMPAGES ", "1")
    return {"footer_fields": list(PAGE_FIELD_CODES), "changed": True}


# ---------------------------------------------------------------------------
# char-format
# ---------------------------------------------------------------------------
# ``w:highlight/@w:val`` is ``ST_HighlightColor`` (wml.xsd) and holds exactly
# these **17** values.  This list is the authority for the command, *not*
# ``docx.enum.text.WD_COLOR_INDEX``: that enum has 18 members of which only 16
# are legal here (``AUTO`` writes the string ``"default"`` and ``INHERITED``
# writes nothing — neither is an ``ST_HighlightColor`` member), and it is
# **missing** ``darkMagenta``, which is the one value the feature was requested
# for.  ``WD_COLOR_INDEX.VIOLET`` writes ``darkMagenta`` under a name nobody
# asked for, so the enum is accepted as an *alias* source, never as the gate.
HIGHLIGHT_VALUES = (
    "black", "blue", "cyan", "darkBlue", "darkCyan", "darkGray", "darkGreen",
    "darkMagenta", "darkRed", "darkYellow", "green", "lightGray", "magenta",
    "none", "red", "white", "yellow",
)

# ``w:highlight`` absent = inherit from the style; ``w:val="none"`` = explicitly
# no highlight.  The command writes nothing for "inherit" and removes the
# element for "none", so neither case leaves a value Word has to interpret.
HIGHLIGHT_INHERIT = "inherit"

_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_COLOR_HEX = re.compile(r"^#?([0-9A-Fa-f]{6})$")


def _highlight_key(value: str) -> str:
    """Case/underscore/hyphen-insensitive key: ``DARK_MAGENTA`` → ``darkmagenta``."""
    return re.sub(r"[\s_\-]", "", value).lower()


def _build_highlight_aliases() -> dict:
    """Every accepted spelling → the canonical ``ST_HighlightColor`` value.

    Three spellings have to work, because three different things name the same
    colour:

    * the schema value itself — ``darkMagenta`` (what ``w:val`` stores);
    * the python-docx member that produces it — ``VIOLET`` / ``DARK_MAGENTA``
      style, i.e. ``WD_COLOR_INDEX`` names, mapped through ``.xml_value``;
    * ``dark_magenta``, which is what a caller writes when they know the colour
      but not which library spells it which way.

    The enum is read for its ``xml_value``, so a member whose value is not a
    legal ``ST_HighlightColor`` (``AUTO`` → ``"default"``) or that has no value
    at all (``INHERITED``) is skipped instead of becoming a trap.
    """
    aliases = {}
    legal = {v.lower(): v for v in HIGHLIGHT_VALUES}
    for canon in HIGHLIGHT_VALUES:
        aliases[_highlight_key(canon)] = canon
    aliases[_highlight_key(HIGHLIGHT_INHERIT)] = HIGHLIGHT_INHERIT
    for member in WD_COLOR_INDEX:
        xml_value = getattr(member, "xml_value", None)
        canon = legal.get((xml_value or "").lower())
        if canon is not None:
            aliases.setdefault(_highlight_key(member.name), canon)
    return aliases


HIGHLIGHT_ALIASES = _build_highlight_aliases()


def normalize_highlight(value: str) -> str:
    """Any accepted highlight spelling → the canonical ``ST_HighlightColor``.

    Raises :class:`CharFormatError` for anything else: a misspelt colour used to
    be written verbatim, which produced a document that validates and renders
    *no highlight at all*, with nothing to explain why.
    """
    key = _highlight_key(value or "")
    canon = HIGHLIGHT_ALIASES.get(key)
    if canon is None:
        raise CharFormatError(
            f"--highlight {value!r} is not a legal w:highlight value; "
            f"legal: {', '.join(HIGHLIGHT_VALUES)} "
            f"(any spelling, e.g. darkMagenta / DARK_MAGENTA / dark_magenta)")
    return canon


def normalize_color(value: str) -> str:
    """``--color`` → the ``RRGGBB`` written into ``w:color/@w:val``.

    ``auto`` is passed through because it is a legal ``ST_HexColorAuto`` value
    and means "let the style decide"; everything else must be six hex digits,
    with or without a leading ``#``.  A three-digit shorthand is refused rather
    than expanded: Word stores no shorthand, so accepting it would be inventing
    a value the caller never wrote.
    """
    raw = (value or "").strip()
    if raw.lower() == "auto":
        return "auto"
    m = _COLOR_HEX.match(raw)
    if not m:
        raise CharFormatError(
            f"--color {value!r} is not a colour: expected 6 hex digits "
            f"(FF0000 or #00a0ff) or 'auto'")
    return m.group(1).upper()


def _rpr_successors(tag: str) -> list:
    """The ``w:``-prefixed tags that must follow ``tag`` inside ``w:rPr``.

    A thin alias for :func:`docx_crossref.successors` on ``RPR_ORDER`` — the
    repo's single copy of the ``CT_RPr`` sequence — so this file cannot drift
    from ``docx_create``: that module builds the same element from the same
    list, and a second hand-typed copy is exactly how ``w:highlight`` ends up
    appended after ``w:u`` on one path and before it on another.

    Why this matters at all: ``w:rPr`` is ``xsd:sequence``.  A child at the
    wrong position still produces well-formed XML, still passes
    ``docx_validate`` (it only checks well-formedness plus a few structure
    rules), and is **silently dropped by Word** — the page shows no highlight
    and nothing reports an error.
    """
    from docx_crossref import successors
    return successors(tag)


def _put_child(rpr, tag: str, value=None):
    """Create-or-replace ``w:<tag>`` in ``w:rPr``, at its schema position.

    The old element is **removed and rebuilt**, not patched.  Two reasons, both
    observed:

    * a pre-existing element may sit at the wrong index (a hand-written
      ``rPr.append(...)``, or content written by another tool).  Patching it in
      place would keep the bad position, so the "fix" would change the XML and
      not the page;
    * ``w:color`` may carry ``w:themeColor`` / ``w:themeShade`` / ``w:themeTint``,
      and in ``CT_Color`` a theme attribute wins over the RGB twin at render
      time.  Rebuilding drops them, which is what makes ``--color`` able to
      override a themed style at all (same reasoning as
      ``docx_create._blacken_run_colour``).

    Returns the element now in the tree.
    """
    old = rpr.find(_q(tag))
    if old is not None:
        rpr.remove(old)
    el = OxmlElement(f"w:{tag}")
    if value is not None:
        el.set(_q("val"), value)
    rpr.insert_element_before(el, *_rpr_successors(tag))
    return el


def _splittable(r_el) -> bool:
    """True for a run this command may cut into pieces.

    Only ``w:rPr`` + exactly one ``w:t`` qualifies.  A run holding a ``w:tab``,
    ``w:br``, ``w:fldChar`` or ``w:drawing`` has a ``run.text`` that is a
    *rendering* of several children, so a character offset into that string does
    not name a cut point in the XML.  Cutting anyway would move a tab or a
    picture into the wrong fragment.
    """
    others = [c.tag for c in r_el if c.tag != _q("rPr")]
    return len(others) == 1 and others[0] == _q("t")


def _merge_ranges(ranges) -> list:
    """Sort and coalesce ``(start, end)`` pairs so no offset is cut twice."""
    merged = []
    for start, end in sorted(ranges):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _split_run_for(run, ranges) -> list:
    """Cut *run* so each range in *ranges* becomes its own run element.

    Returns the run elements that cover the requested ranges; the untouched
    fragments are returned to the tree around them.  Every fragment is a
    ``copy.deepcopy`` of the original ``w:r``, so the run's whole ``w:rPr`` —
    bold, size, font, subscript, anything — travels with the text; this is the
    same "clone the carrier" idea as ``_fill_cell_text``, applied at character
    granularity instead of cell granularity.

    A run that cannot be cut (:func:`_splittable`) is only acceptable when a
    range covers it whole; a partial range there is refused instead of being
    rounded up to the whole run, because "highlight a bit more text than you
    asked for" is indistinguishable from a bug at the caller's end.
    """
    text = run.text
    total = len(text)
    cuts = {0, total}
    for start, end in ranges:
        # A range may extend past the run (the caller clipped it) — clamp so the
        # cut set only contains real offsets.
        cuts.add(max(0, min(total, start)))
        cuts.add(max(0, min(total, end)))
    cuts = sorted(cuts)
    segments = []
    for start, end in zip(cuts, cuts[1:]):
        if start == end:
            continue
        wanted = any(a <= start and end <= b for a, b in ranges)
        segments.append((start, end, wanted))

    if len(segments) == 1 and segments[0][2]:
        return [run._r]
    if not _splittable(run._r):
        raise CharFormatError(
            f"cannot apply character formatting to part of a run: the run "
            f"holding {text!r} also carries tabs, breaks, fields or a drawing, "
            f"so it cannot be split without moving those children. Format a "
            f"wider anchor that covers the whole run, or edit the run directly.")

    r_el = run._r
    parent = r_el.getparent()
    made = []
    for start, end, wanted in segments:
        fragment = copy.deepcopy(r_el)
        t_el = fragment.find(_q("t"))
        t_el.text = text[start:end]
        # A fragment that starts or ends with a space is trimmed by Word unless
        # the run says otherwise; set it on every fragment because any of them
        # may be the one holding the edge.
        t_el.set(_XML_SPACE, "preserve")
        parent.insert(parent.index(r_el), fragment)
        if wanted:
            made.append(fragment)
    parent.remove(r_el)
    return made


def _apply_char_format(r_el, *, highlight, colour, bold, italic) -> None:
    """Write one run's character formatting, in schema order, idempotently."""
    rpr = r_el.get_or_add_rPr()
    if highlight == HIGHLIGHT_INHERIT:
        pass
    elif highlight == "none":
        old = rpr.find(_q("highlight"))
        if old is not None:
            rpr.remove(old)
    elif highlight is not None:
        _put_child(rpr, "highlight", highlight)
    if colour is not None:
        _put_child(rpr, "color", colour)
    # ``w:b`` / ``w:i`` with no ``w:val`` mean "on"; ``w:val="0"`` means "off"
    # and is *not* the same as absent, which inherits from the style.  Only the
    # Latin twin is written: python-docx's ``font.bold`` does the same, and a
    # ``w:bCs`` that disagrees with ``w:b`` is its own kind of surprise.
    if bold is not None:
        _put_child(rpr, "b", None if bold else "0")
    if italic is not None:
        _put_child(rpr, "i", None if italic else "0")


def cmd_char_format(doc, args) -> dict:
    """Set character formatting on the runs that carry ``--find``.

    Selector and scope, both deliberate:

    * **``--find`` only, no ``--index``.**  ``docx_read --runs`` numbers
      paragraphs across body *and* table cells *and* text boxes, while
      ``doc.paragraphs`` numbers body paragraphs only.  Offering ``--index``
      here would mean two different things called "index 7" in the same tool
      pair, and the caller would have no way to tell which one they got.  Text
      is also the more durable selector: any edit shifts positions, and a
      shifted position still "works" while highlighting the wrong line.
    * **Scope matches ``replace``**: body, tables (nested included), headers and
      footers.  Text boxes and tracked-change (``w:ins``) content are outside
      it, for the same reason ``set-cell`` is — ``Paragraph.runs`` does not
      reach them, so neither can a run-level writer.

    A run is only touched when it overlaps the anchor; a run that carries the
    anchor *and* other text is split, so the surrounding characters keep the
    formatting they had.  The anchor may span several consecutive runs.
    """
    highlight = None
    if args.highlight is not None:
        highlight = normalize_highlight(args.highlight)
    colour = normalize_color(args.color) if args.color is not None else None
    bold, italic = args.bold, args.italic
    if highlight is None and colour is None and bold is None and italic is None:
        raise CharFormatError(
            "char-format needs at least one format to set: --highlight, "
            "--color, --bold/--no-bold, --italic/--no-italic")
    if not args.find:
        raise CharFormatError("char-format --find must not be empty")

    applied = {}
    if highlight is not None:
        applied["highlight"] = highlight
    if colour is not None:
        applied["color"] = colour
    if bold is not None:
        applied["bold"] = bold
    if italic is not None:
        applied["italic"] = italic

    occurrences = 0
    touched_paragraphs = 0
    changed_runs = 0
    paragraphs_seen = 0
    # Dedupe by element *equality*, not by ``id()``.  lxml builds a proxy on
    # demand, so an ``id()`` taken here can be recycled by a later proxy for a
    # different node — measured: a header paragraph and a footer paragraph came
    # back with the same ``id()`` and the footer was skipped, so `--find` on
    # footer text reported "no run carries this text" for text that was there.
    # Element ``==`` compares the underlying node (see
    # ``docx_read._own_paragraphs``), which is the identity that matters.
    seen = []

    for para in iter_all_paragraphs(doc):
        # Merged table cells are yielded once per grid column they span; acting
        # twice on the same ``w:p`` is harmless (the second pass finds the runs
        # already cut), but it would inflate every count in the report.
        p_el = para._p
        if any(p_el == other for other in seen):
            continue
        seen.append(p_el)
        paragraphs_seen += 1

        runs = [r for r in para.runs if _writable_run(r)]
        spans = []
        pos = 0
        for run in runs:
            text = run.text
            spans.append((run, pos, pos + len(text)))
            pos += len(text)
        joined = "".join(run.text for run, _, _ in spans)
        if not joined:
            continue

        found = []
        at = joined.find(args.find)
        while at >= 0:
            found.append((at, at + len(args.find)))
            at = joined.find(args.find, at + len(args.find))
        if not found:
            continue

        occurrences += len(found)
        touched_paragraphs += 1

        per_run = {}
        for start, end in found:
            for run, r_start, r_end in spans:
                if r_end > start and r_start < end:
                    per_run.setdefault(run, []).append(
                        (max(start, r_start) - r_start,
                         min(end, r_end) - r_start))

        for run, _, _ in spans:
            if run not in per_run:
                continue
            for r_el in _split_run_for(run, _merge_ranges(per_run[run])):
                _apply_char_format(r_el, highlight=highlight, colour=colour,
                                   bold=bold, italic=italic)
                changed_runs += 1

    if not occurrences:
        raise CharFormatError(
            f"char-format --find {args.find!r}: no run in the document carries "
            f"this text. Searched {paragraphs_seen} paragraph(s) across body, "
            f"tables, headers and footers (text boxes and tracked-change "
            f"content are out of scope, as for every other run-level writer).")

    result = {"find": args.find, "occurrences": occurrences,
              "runs_changed": changed_runs,
              "paragraphs_touched": touched_paragraphs,
              "paragraphs_searched": paragraphs_seen, "applied": applied}
    # Flattened at top level as well, so a caller can read
    # ``result["highlight"]`` without knowing the ``applied`` wrapper — the
    # value reported is the **normalized** one that was actually written, which
    # is how a caller confirms ``DARK_MAGENTA`` landed as ``darkMagenta``.
    result.update(applied)
    return result


def main() -> int:
    ap = argparse.ArgumentParser(description="Edit a .docx file.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("path", help="input .docx")
        p.add_argument("-o", "--output",
                       help="output path (default: overwrite input)")

    p = sub.add_parser("replace", help="find-and-replace text")
    common(p)
    p.add_argument("--find", required=True)
    p.add_argument("--replace", required=True)
    p.add_argument("--body-only", action="store_true",
                   help="skip headers/footers")

    p = sub.add_parser("set-cell", help="set table cell text")
    common(p)
    p.add_argument("--table", type=int, required=True, help="table index")
    p.add_argument("--row", type=int, required=True)
    p.add_argument("--col", type=int, required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--replace-cell", action="store_true",
                   help="old behaviour: clear the whole cell and write one "
                        "paragraph with one run (drops all run-level "
                        "formatting). Default keeps what a flat string can keep.")

    p = sub.add_parser("insert", help="insert paragraph at body index")
    common(p)
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--text", required=True)
    p.add_argument("--style", default=None)

    p = sub.add_parser("delete", help="delete body paragraph by index")
    common(p)
    p.add_argument("--index", type=int, required=True)

    p = sub.add_parser("style", help="apply style to body paragraph")
    common(p)
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--style", required=True)

    p = sub.add_parser("normalize",
                       help="merge adjacent runs with identical formatting")
    common(p)

    p = sub.add_parser("toc", help="insert a TOC field (Word computes it)")
    common(p)
    p.add_argument("--index", type=int, default=0,
                   help="body paragraph index to insert before (default 0)")

    p = sub.add_parser("page-numbers",
                       help="add PAGE/NUMPAGES fields to the footer")
    common(p)

    p = sub.add_parser(
        "char-format",
        help="set run-level character formatting (highlight/colour/bold/italic)",
        description="Set character formatting on the runs that carry an "
                    "anchor text. Runs are split when the anchor is only part "
                    "of one, so the text around it keeps its own formatting.")
    common(p)
    _apply_char_format_args(p)

    args = ap.parse_args()
    # The input file is validated here — before `Document()` and before any
    # subcommand — because a broken package used to be a bare traceback with
    # exit 1 on **every** subcommand (measured 2026-11-11: 19 lines for a
    # missing/non-zip file, 23 for a directory, 31 for a ZIP with no
    # `[Content_Types].xml`). Exit 2 is the project-wide code for "you named
    # something that is not there / not usable as given"; exit 1 stays reserved
    # for "the document is real but has a problem", which is `docx_validate`'s
    # job. stdout stays empty: every success prints JSON, so a failure must not
    # look like an empty-but-successful JSON document. `doc.save()` is the only
    # writer and it is far below, so nothing is written on this path.
    try:
        _check_input(args.path)
        doc = _open_document(args.path)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if args.cmd == "replace":
        if args.body_only:
            n = 0
            for para in iter_all_paragraphs(doc, include_headers_footers=False):
                n += replace_in_paragraph(para, args.find, args.replace)
            result = {"replacements": n}
        else:
            result = cmd_replace(doc, args)
    else:
        try:
            if args.cmd == "set-cell":
                result = cmd_set_cell(doc, args)
            elif args.cmd == "insert":
                result = cmd_insert(doc, args)
            elif args.cmd == "delete":
                result = cmd_delete(doc, args)
            elif args.cmd == "normalize":
                result = cmd_normalize(doc)
            elif args.cmd == "toc":
                result = cmd_toc(doc, args)
            elif args.cmd == "page-numbers":
                result = cmd_page_numbers(doc, args)
            elif args.cmd == "char-format":
                result = cmd_char_format(doc, args)
            else:
                result = cmd_style(doc, args)
        except CharFormatError as exc:
            # A bad colour value or an anchor that matches nothing: the caller
            # named something that is not there, which is exit 2 everywhere
            # else in this CLI (D-024 / D-030). `doc.save()` is below, so the
            # input file is untouched.
            print(f"error: {exc}", file=sys.stderr)
            return 2
        except IndexOutOfRangeError as exc:
            # Same exit code argparse uses for a bad argument value: the
            # caller named a paragraph/table/row/column that does not exist.
            # Nothing has been saved at this point — `doc.save()` is below —
            # so the input file is byte-for-byte untouched.
            print(f"error: {exc}", file=sys.stderr)
            return 2

    out = args.output or args.path
    # A no-op (-page-numbers on a footer that already carries the fields) must
    # not rewrite the package: an in-place re-run then leaves mtime and SHA256
    # untouched, which is what "idempotent" has to mean for an operator who
    # runs the same plan twice. A distinct -o still gets a file, or the CLI
    # contract ("there is always an output path") would break.
    if result.get("changed", True) \
            or os.path.abspath(out) != os.path.abspath(args.path):
        doc.save(out)
    result.update({"ok": True, "output": out})
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
