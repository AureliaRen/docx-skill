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

Examples:
  docx_edit.py replace in.docx --find old --replace new -o out.docx
  docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text "42"
  docx_edit.py insert in.docx --index 3 --text "New para" --style Normal
  docx_edit.py delete in.docx --index 3
  docx_edit.py style in.docx --index 0 --style "Heading 1"
  docx_edit.py normalize in.docx -o out.docx
  docx_edit.py toc in.docx --index 1 -o out.docx
  docx_edit.py page-numbers in.docx -o out.docx

Field results (TOC entries, page numbers) are computed by Word or
LibreOffice when the document is opened, not by python-docx; until then
the fields show placeholder text.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import os
import sys

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


def cmd_set_cell(doc, args) -> dict:
    """Set one table cell's text, refusing any index that names nothing.

    All three dimensions are checked, in order, before a cell is touched; the
    message names the dimension that was wrong together with its own legal
    range, so ``--row 9`` never reports the column count.

    Column count comes from the row itself rather than ``table.columns``: a
    cell grid can be ragged, and ``table.cell(row, col)`` indexes the row.
    """
    tables = doc.tables
    _check_range("set-cell --table", args.table, len(tables))
    table = tables[args.table]
    rows = table.rows
    _check_range("set-cell --row", args.row, len(rows))
    _check_range("set-cell --col", args.col, len(rows[args.row].cells))
    table.cell(args.row, args.col).text = args.text
    return {"table": args.table, "row": args.row, "col": args.col}


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
            else:
                result = cmd_style(doc, args)
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
