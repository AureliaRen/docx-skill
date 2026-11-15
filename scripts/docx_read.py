#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Read a .docx: text, structure outline, styles, images, revision detection.

Usage:
  docx_read.py file.docx --text        # full text incl. tables + headers/footers
  docx_read.py file.docx --structure   # JSON outline (headings, tables, counts)
  docx_read.py file.docx --blocks      # ordered body stream (paragraphs+tables)
  docx_read.py file.docx --runs        # per-run inventory: sub/superscript, bold,
                                       # italic, size, font + mixed_format flag
  docx_read.py file.docx --styles      # JSON list of styles actually used
  docx_read.py file.docx --images DIR  # extract embedded images into DIR
  docx_read.py file.docx --revisions   # JSON: tracked changes / comments present?

Text output is JSON: {"body": [...], "tables": [[...rows]], "headers": [...],
"footers": [...]}. Body text is the accepted/as-is text (python-docx ignores
deleted-in-revision text and shows inserted text).
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import os
import sys
import zipfile

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def cell_text(cell) -> str:
    """Flatten a cell's paragraphs and nested tables in document order."""
    parts = []
    for block in cell.iter_inner_content():
        if hasattr(block, "rows"):
            for row in table_to_rows(block):
                parts.extend(text for text in row if text)
        elif block.text:
            parts.append(block.text)
    return "\n".join(parts)


def table_to_rows(table) -> list:
    return [[cell_text(cell) for cell in row.cells] for row in table.rows]


def append_part_text(target: list[str], part, seen: set[int]) -> None:
    """Append one header/footer variant without duplicating linked parts."""
    marker = id(part._element)
    if marker in seen:
        return
    seen.add(marker)
    target.extend(p.text for p in part.paragraphs)
    for table in part.tables:
        target.append(json.dumps(table_to_rows(table), ensure_ascii=False))


def extract_text(doc) -> dict:
    out = {"body": [p.text for p in doc.paragraphs],
           "tables": [table_to_rows(t) for t in doc.tables],
           "headers": [], "footers": []}
    seen_headers: set[int] = set()
    seen_footers: set[int] = set()
    for section in doc.sections:
        for part in (section.header, section.first_page_header,
                     section.even_page_header):
            append_part_text(out["headers"], part, seen_headers)
        for part in (section.footer, section.first_page_footer,
                     section.even_page_footer):
            append_part_text(out["footers"], part, seen_footers)
    return out


_NAME_LEVELS = {"章标题": 1, "节标题": 2, "条标题": 3,
                "一级标题": 1, "二级标题": 2, "三级标题": 3}


def heading_level(para):
    """Outline level of a paragraph, or None for body text.

    Resolution order: direct w:outlineLvl -> style-chain w:outlineLvl ->
    English "Heading N" -> CJK style-name table. Word writes level 9 for
    normal body text, so 9+ maps to None.
    """
    pPr = para._p.pPr
    if pPr is not None:
        el = pPr.find(qn("w:outlineLvl"))
        if el is not None:
            val = int(el.get(qn("w:val")))
            return None if val >= 9 else val + 1
    style, seen = para.style, set()
    while style is not None and id(style) not in seen:
        seen.add(id(style))
        spPr = style.element.find(qn("w:pPr"))
        if spPr is not None:
            el = spPr.find(qn("w:outlineLvl"))
            if el is not None:
                val = int(el.get(qn("w:val")))
                return None if val >= 9 else val + 1
        style = style.base_style
    name = para.style.name if para.style else ""
    if name.startswith("Heading"):
        try:
            return int(name.split()[-1])
        except ValueError:
            return 1
    return _NAME_LEVELS.get(name)


def extract_structure(doc) -> dict:
    outline = []
    for i, para in enumerate(doc.paragraphs):
        level = heading_level(para)
        if level:
            outline.append({"index": i, "level": level, "text": para.text})
    return {
        "outline": outline,
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "tables": [{"rows": len(t.rows), "cols": len(t.columns)}
                   for t in doc.tables],
        "section_count": len(doc.sections),
    }


def styles_used(doc) -> list:
    used = set()
    for para in doc.paragraphs:
        if para.style:
            used.add(para.style.name)
        for run in para.runs:
            if run.style:
                used.add(run.style.name)
    for table in doc.tables:
        if table.style:
            used.add(table.style.name)
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.style:
                        used.add(para.style.name)
    return sorted(used)


def extract_images(path: str, outdir: str) -> list:
    os.makedirs(outdir, exist_ok=True)
    written = []
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            # zip directory entries ("word/media/") are not files
            if name.startswith("word/media/") and not name.endswith("/"):
                target = os.path.join(outdir, os.path.basename(name))
                with open(target, "wb") as f:
                    f.write(zf.read(name))
                written.append(target)
    return written


def extract_blocks(doc) -> list:
    """Ordered body stream: paragraphs and tables interleaved, as authored.

    The flat body[]/tables[] split loses insertion order; this keeps it, and
    carries the per-paragraph style plus any embedded image relationship ids
    so a markdown round trip can rebuild headings, tables and figures.
    """
    blocks = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            para = Paragraph(child, doc)
            level = heading_level(para)
            rids = [e.get(qn("r:embed")) or e.get(qn("r:link"))
                    for e in child.iter(qn("a:blip"))]
            names = []
            for rid in rids:
                if not rid:
                    continue
                try:
                    part = doc.part.related_parts[rid]
                    names.append(os.path.basename(str(part.partname)))
                except KeyError:
                    pass
            blocks.append({
                "type": "heading" if level else "paragraph",
                "style": para.style.name if para.style else "",
                "level": level,
                "text": para.text,
                "images": names,
            })
        elif tag == "tbl":
            table = Table(child, doc)
            blocks.append({"type": "table", "style": table.style.name
                           if table.style else "",
                           "rows": table_to_rows(table)})
    return blocks


# ---------------------------------------------------------------- run-level
# `--blocks` stops at the paragraph: it reports the style and the image
# relationship ids, and nothing about the runs inside. That is deliberate -- it
# is the contract the markdown round trip consumes -- so the run-level view
# lives in its own mode instead of growing `--blocks`. A caller that is about to
# *write* into a paragraph needs to see the run split first: a cell whose first
# run carries `w:vertAlign="subscript"` looks like an ordinary "FS" from every
# existing mode, and a flat `set-cell` would then format the whole string as
# subscript (measured on the owner's own homework template).
_RPR = qn("w:rPr")


def _rpr_flag(rpr, tag: str):
    """Direct-formatting boolean as a **tri-state**.

    ``None``  = the element is absent (nothing stated at this level);
    ``True``  = present with no ``w:val`` or an affirmative one;
    ``False`` = present with ``w:val`` of ``0`` / ``false`` / ``off``.

    The tri-state matters: ``w:b`` present-but-false and ``w:b`` absent mean
    different things to a caller that is about to copy this run's formatting
    onto new text. A plain ``bool(get(...))`` reads both as "not bold" and
    throws that difference away.
    """
    if rpr is None:
        return None
    el = rpr.find(qn(f"w:{tag}"))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val is None:
        return True
    return val.strip().lower() not in ("0", "false", "off")


def _rpr_vert_align(rpr):
    """``w:vertAlign/@w:val`` -- ``subscript`` / ``superscript`` / ``baseline``.

    Returns ``None`` when the run states nothing. ``baseline`` is the value Word
    writes to *cancel* an inherited sub/superscript, so it is reported as the
    string rather than folded into ``None``: they are not interchangeable.
    """
    if rpr is None:
        return None
    el = rpr.find(qn("w:vertAlign"))
    if el is None:
        return None
    return el.get(qn("w:val"))


def _rpr_size_pt(rpr):
    """``w:sz/@w:val`` is half-points; report points. ``None`` when absent."""
    if rpr is None:
        return None
    el = rpr.find(qn("w:sz"))
    if el is None:
        return None
    raw = el.get(qn("w:val"))
    try:
        return int(raw) / 2
    except (TypeError, ValueError):
        return None


def _rpr_font(rpr):
    """The ``ascii`` font of ``w:rFonts`` (``None`` when absent)."""
    if rpr is None:
        return None
    el = rpr.find(qn("w:rFonts"))
    if el is None:
        return None
    return el.get(qn("w:ascii"))


def describe_run(run) -> dict:
    """One run as JSON-able fields: text, character formatting, field/drawing."""
    r_el = run._r
    rpr = r_el.find(_RPR)
    return {
        "text": run.text,
        "vert_align": _rpr_vert_align(rpr),
        "bold": _rpr_flag(rpr, "b"),
        "italic": _rpr_flag(rpr, "i"),
        "size_pt": _rpr_size_pt(rpr),
        "font": _rpr_font(rpr),
        # A field run carries the field's control characters, not content the
        # caller can overwrite with text; a drawing run carries a picture.
        # Both are called out so `--runs` never presents them as writable text.
        "is_field": (r_el.find(qn("w:fldChar")) is not None
                     or r_el.find(qn("w:instrText")) is not None),
        "is_drawing": (r_el.find(qn("w:drawing")) is not None
                       or r_el.find(qn("w:pict")) is not None),
    }


def _format_fingerprint(run_info: dict) -> tuple:
    """Everything that would be lost by overwriting a run with flat text."""
    return (run_info["vert_align"], run_info["bold"], run_info["italic"],
            run_info["size_pt"], run_info["font"])


def is_mixed_format(runs: list) -> bool:
    """True when a paragraph's plain-text runs disagree on formatting.

    "Plain-text run" = carries text and is neither a field nor a drawing: those
    two cannot be given new text in the first place, so counting them would flag
    paragraphs that are perfectly safe to overwrite.

    The whole point of the flag is the owner's second trap: a paragraph reading
    ``FS`` where ``F`` is a subscript run and ``S`` is not is *mixed*, and a
    flat-string write cannot express that split. One run, or several runs that
    agree, is not mixed -- flagging those would make the flag noise, and noise
    is what the caller will learn to ignore.
    """
    plain = [r for r in runs
             if r["text"] and not r["is_field"] and not r["is_drawing"]]
    if len(plain) < 2:
        return False
    return len({_format_fingerprint(r) for r in plain}) > 1


def _table_path_label(table_path) -> str:
    """``table[0].r1.c2`` chain, or ``body`` for a top-level paragraph."""
    if not table_path:
        return "body"
    return " > ".join(f"table[{t}].r{r}.c{c}" for t, r, c in table_path)


def _iter_paragraphs_annotated(container, table_path=(), seen_cells=None):
    """Yield ``(paragraph, label, table_path)`` in document order.

    Descends into table cells and nested tables, carrying the chain of
    ``(table, row, col)`` indices that leads to each cell, so the caller can go
    straight from "this cell is the trap" to the ``set-cell --table/--row/--col``
    that would touch it.

    Merged cells are visited once: ``row.cells`` repeats the same ``w:tc`` for
    every grid column it spans, and emitting it twice would double every run in
    it. Identity of the ``w:tc`` element is the dedupe key -- that is the same
    object ``table.cell(r, c)`` returns, so a caller that acts on the reported
    indices still lands on the cell that was described.
    """
    if seen_cells is None:
        seen_cells = set()
    for para in container.paragraphs:
        yield para, _table_path_label(table_path), table_path
    for t_index, table in enumerate(container.tables):
        for r_index, row in enumerate(table.rows):
            for c_index, cell in enumerate(row.cells):
                marker = id(cell._tc)
                if marker in seen_cells:
                    continue
                seen_cells.add(marker)
                yield from _iter_paragraphs_annotated(
                    cell, table_path + ((t_index, r_index, c_index),),
                    seen_cells)


def extract_runs(doc) -> dict:
    """Per-paragraph run inventory: the view a writer needs *before* writing.

    Scope: body paragraphs plus every table cell, nested tables included.
    Header/footer parts are **not** walked -- `--blocks` does not cover them
    either, and `set-cell` cannot reach them, so listing them here would report
    traps that no edit command can act on. Use `--text` for their text.
    """
    paragraphs = []
    for index, (para, label, table_path) in enumerate(
            _iter_paragraphs_annotated(doc)):
        runs = [describe_run(run) for run in para.runs]
        paragraphs.append({
            "index": index,
            "path": label,
            "table_path": [list(step) for step in table_path],
            "style": para.style.name if para.style else "",
            "text": para.text,
            "mixed_format": is_mixed_format(runs),
            "runs": runs,
        })
    return {"paragraphs": paragraphs}


def detect_revisions(path: str) -> dict:
    """Detect tracked changes and comments by scanning the raw XML parts."""
    markers = {"insertions": b"<w:ins ", "deletions": b"<w:del ",
               "format_changes": b"<w:rPrChange"}
    result = {k: False for k in markers}
    result["comments"] = False
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        result["comments"] = any(n.startswith("word/comments") for n in names)
        for name in names:
            if name.startswith("word/") and name.endswith(".xml"):
                data = zf.read(name)
                for key, marker in markers.items():
                    if marker in data:
                        result[key] = True
    result["has_tracked_changes"] = any(
        result[k] for k in ("insertions", "deletions", "format_changes"))
    return result


def _md_cell(text: str) -> str:
    return text.replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def to_markdown(doc, image_dir: str | None = None,
                base_dir: str | None = None) -> str:
    """Body stream -> Markdown plain text, in authored order.

    Headings come from the outline-level resolver, so CJK custom styles
    (章标题/节标题/条标题) map to #/##/###. Tables are emitted inline where
    they sit. Inline images become ![](relpath) when image_dir is given.
    """
    blocks = extract_blocks(doc)
    out: list[str] = []
    for b in blocks:
        if b["type"] == "heading":
            text = b["text"].strip()
            if not text:
                continue
            out.append("#" * min(b["level"] or 1, 6) + " " + text)
            out.append("")
        elif b["type"] == "table":
            rows = b["rows"]
            if not rows:
                continue
            width = max(len(r) for r in rows)
            for r_i, row in enumerate(rows):
                cells = list(row) + [""] * (width - len(row))
                out.append("| " + " | ".join(_md_cell(c) for c in cells) + " |")
                if r_i == 0:
                    out.append("|" + "|".join(["---"] * width) + "|")
            out.append("")
        else:
            text = b["text"].strip()
            for img in b.get("images", []):
                if image_dir and base_dir:
                    target = os.path.join(image_dir, img)
                    link = os.path.relpath(target, base_dir).replace("\\", "/")
                else:
                    link = img
                out.append(f"![{os.path.splitext(img)[0]}]({link})")
            if text:
                out.append(text)
            if text or b.get("images"):
                out.append("")
    return "\n".join(out).rstrip() + "\n"


class InputError(Exception):
    """The path named on the command line is not a readable .docx package.

    Raised *before* any output is produced, so `main` can refuse with exit 2,
    one line on stderr, and an empty stdout. That separation matters here
    because every successful mode writes **JSON to stdout**: relaying a
    failure through stdout as JSON would make every `json.loads(stdout)`
    caller read a failure as a success. Exit 2 is the project-wide code for
    "you named something that is not there / not usable as given"; exit 1 is
    reserved for "the document is real but has a problem", which is
    `docx_validate`'s job, not this one's.
    """


def _check_input(path: str) -> None:
    """Reject a path that cannot possibly be a Word package, before opening it.

    Three distinct failures, each with its own wording, because they call for
    three different fixes from the caller:

    * ``no such file``      → the path is wrong (typo, wrong cwd, bad quoting);
    * ``not a file``        → it exists but is a directory (a common shell glob);
    * ``not a ZIP package`` → it is a file but not a ZIP (usually a .doc/.rtf/
      .txt that somebody renamed to .docx).

    ``zipfile.is_zipfile`` is used rather than letting ``Document()`` raise:
    it reads only the tail of the file, so a 3 GB non-zip is refused cheaply.
    """
    if not os.path.exists(path):
        raise InputError(f"--path {path}: no such file")
    if not os.path.isfile(path):
        raise InputError(f"not a file: {path}")
    if not zipfile.is_zipfile(path):
        raise InputError(f"not a ZIP package: {path}")


def _open_document(path: str):
    """``Document(path)`` with the package errors turned into :class:`InputError`.

    Defence in depth against TOCTOU: `_check_input` runs first, but the file
    can be replaced, truncated or deleted between that check and this open —
    and ``is_zipfile`` only inspects the tail, so a package can pass the check
    and still lack ``word/document.xml``. Both cases surface as one usage
    error instead of a python-docx traceback.
    """
    try:
        return Document(path)
    except PackageNotFoundError:
        raise InputError(f"not a Word package: {path}") from None
    except zipfile.BadZipFile:
        raise InputError(f"not a ZIP package: {path}") from None
    except KeyError:
        # A ZIP with no ``[Content_Types].xml``: ``is_zipfile`` accepts it,
        # ``Document`` then dies on a bare KeyError while reading the package.
        raise InputError(f"not a Word package: {path}") from None


def main() -> int:
    ap = argparse.ArgumentParser(description="Read/inspect a .docx file.")
    ap.add_argument("path", help=".docx file to read")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--text", action="store_true", help="extract all text as JSON")
    g.add_argument("--structure", action="store_true", help="outline JSON")
    g.add_argument("--styles", action="store_true", help="styles used, JSON")
    g.add_argument("--images", metavar="DIR", help="extract images to DIR")
    g.add_argument("--revisions", action="store_true",
                   help="detect tracked changes / comments")
    g.add_argument("--blocks", action="store_true",
                   help="ordered body stream (paragraphs + tables interleaved)")
    g.add_argument("--runs", action="store_true",
                   help="per-paragraph run inventory: text, sub/superscript, "
                        "bold, italic, size, font, field/drawing, and "
                        "mixed_format (read this before writing into a cell)")
    g.add_argument("--md", metavar="FILE",
                   help="write a Markdown plain-text rendition to FILE")
    ap.add_argument("--image-dir", metavar="DIR",
                    help="with --md: extract images here and link them")
    args = ap.parse_args()

    try:
        _check_input(args.path)
        return _run(args)
    except InputError as exc:
        # stdout stays empty on every failure: success is the only thing that
        # prints, and it prints JSON. Nothing has been written at this point
        # either — `--images` / `--md` only create their targets below.
        print(f"error: {exc}", file=sys.stderr)
        return 2


def _run(args) -> int:
    """Dispatch one output mode. Assumes the input already passed validation."""
    if args.images:
        try:
            written = extract_images(args.path, args.images)
        except (zipfile.BadZipFile, OSError) as exc:
            # `_check_input` already accepted this path; if it fails now the
            # file changed underneath us. Same one-line usage error, and the
            # output directory is only created inside `extract_images`, so a
            # refused run leaves no directory behind.
            raise InputError(f"cannot read {args.path}: {exc}") from None
        print(json.dumps({"images": written}, ensure_ascii=False))
        return 0
    if args.revisions:
        try:
            revisions = detect_revisions(args.path)
        except (zipfile.BadZipFile, OSError) as exc:
            raise InputError(f"cannot read {args.path}: {exc}") from None
        print(json.dumps(revisions, ensure_ascii=False))
        return 0

    doc = _open_document(args.path)
    if args.md:
        if args.image_dir:
            extract_images(args.path, args.image_dir)
        md = to_markdown(doc, image_dir=args.image_dir,
                         base_dir=os.path.dirname(os.path.abspath(args.md)))
        with open(args.md, "w", encoding="utf-8") as f:
            f.write(md)
        blocks = extract_blocks(doc)
        print(json.dumps({"ok": True, "markdown": args.md, "chars": len(md),
                          "blocks": len(blocks),
                          "tables": sum(1 for b in blocks
                                        if b["type"] == "table"),
                          "headings": sum(1 for b in blocks
                                          if b["type"] == "heading"),
                          "images": sum(len(b.get("images", []))
                                        for b in blocks)},
                         ensure_ascii=False))
        return 0
    if args.blocks:
        out = {"blocks": extract_blocks(doc)}
    elif args.runs:
        out = extract_runs(doc)
    elif args.text:
        out = extract_text(doc)
    elif args.structure:
        out = extract_structure(doc)
    else:
        out = {"styles": styles_used(doc)}
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
