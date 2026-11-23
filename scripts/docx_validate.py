#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Health-check a .docx package and report issues as JSON.

Usage: docx_validate.py file.docx

Checks (health-check tier, NOT full XSD schema validation):
  - the file is a readable zip and python-docx can open it
  - required package parts exist ([Content_Types].xml, document.xml)
  - every relationship in every .rels file resolves to a part in the
    package (dangling image/hyperlink/etc. rels are reported; external
    targets such as hyperlinks are skipped)
  - r:embed / r:id references in document.xml resolve to relationships
  - embedded images are non-empty and start with known magic bytes
    (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/SVG); no PIL required
  - paragraph and run style ids referenced by the document exist in
    styles.xml
  - **structural sanity** of document.xml: a ``<w:drawing/>`` with neither
    ``wp:inline`` nor ``wp:anchor``, and a ``<w:p>`` / ``<w:tbl>`` sitting
    directly inside a ``<w:r>``.  Both make Word refuse the whole file while
    every reference in it still resolves, so the reference checks above cannot
    see them.  This is *not* XSD validation — it is only the two shapes that
    are cheap to detect and fatal in practice.
  - **restricted run-property values** in every ``word/*.xml`` part: the
    enumerated children of ``<w:rPr>`` whose ``w:val`` comes from a closed
    value list.  Two severities, **decided by what Word 16.0 actually does with
    each value rather than by the schema** (see the comment block above
    ``_EFFECT_VALUES`` for the measurement table):
    ``invalid-run-property`` (**error**) for the ones that make Word refuse to
    open the file — ``w:highlight`` / ``w:vertAlign`` / ``w:effect`` — and
    ``nonstandard-run-property`` (**warning**) for ``w:u`` / ``w:color``,
    which Word opens happily and merely re-interprets.  A value outside the
    schema but tolerated by Word must not become an error: this report is
    consumed downstream, and a false error rejects a document that opens.
    The rule cuts both ways, so the tolerances are recorded too: ``w:vertAlign
    w:val="normal"`` opens (schema-illegal, so it is accepted) while
    ``w:effect w:val="ants"`` is **refused** (schema-legal, so it is an error).

``word/document.xml`` (and ``word/styles.xml``) failing to parse is reported as
one ``bad-document-xml`` / ``bad-styles-xml`` issue instead of a traceback: the
package is a real .docx, so it is a document problem (exit 1, JSON on stdout),
not a usage error.

Output: {"ok": bool, "issues": [{"severity": "error"|"warning", ...}]}
Exit code 1 when any error-severity issue is found (warnings exit 0).
Exit code 2 when the path is not a readable .docx at all (one line on stderr).

Footer semantics are delegated to :mod:`docx_footer` (one implementation, two
entry points; ``docx_footer check`` is the detailed report).  Two codes come
from there:

* ``footer-textbox-overflow`` (**error**) — a floating text box holding the page
  line is narrower than the line needs, so the number is clipped.  This is
  content loss, not ugliness: a floating box does not re-flow, the glyphs are
  drawn past the box and cut by the margin.  See the severity note in
  :mod:`docx_footer` for why the estimate is allowed to raise an error at all.
* ``footer-no-page-field`` (**warning**) — ``<w:pgNumType>`` is set on a section
  that renders a non-empty footer, but that footer carries no PAGE field.  A
  warning rather than an error because it is a legal intermediate state (the
  counter is configured, the footer is filled later) and because the real corpus
  has four such documents that are simply *without* a footer, not *with a
  broken* one.

A footer part that cannot be read reports ``footer-check-skipped`` (warning)
instead of raising — this script must never trade its JSON report for a
traceback.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import posixpath
import re
import sys
import zipfile

from lxml import etree

# 输入校验的三件套由 `docx_read` 单一拥有（`InputError` / `_check_input`），
# 这里只 import、不复制判定 —— 与 `docx_edit.py` / `docx_mathify.py` 同形。
from docx_read import InputError, _check_input

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
# ``wp:`` — the WordprocessingDrawing namespace that owns ``wp:inline`` /
# ``wp:anchor``.  A ``<w:drawing/>`` with neither child is the shape Word
# refuses outright (it has no idea where or how big the object is).  The
# DrawingML ``a:`` namespace is deliberately *not* used: ``a:blip`` presence is
# not required for a legal drawing (charts, shapes, text boxes).
WP = ("http://schemas.openxmlformats.org/drawingml/2006/"
      "wordprocessingDrawing")

#: Children a ``<w:p>`` / ``<w:tbl>`` may never sit directly inside a run.
#: ``w:r`` holds run content (``w:t``, ``w:drawing``, ``w:br``, ``w:tab`` …);
#: a block-level element in there is XML that Word silently repairs or refuses,
#: and python-docx happily opens it — so no other check here would notice.
RUN_FORBIDDEN_CHILDREN = (f"{{{W}}}p", f"{{{W}}}tbl")

IMAGE_MAGIC = (
    b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff", b"GIF87a", b"GIF89a",
    b"BM", b"II*\x00", b"MM\x00*",
    b"\x01\x00\x00\x00",              # EMF
    b"\xd7\xcd\xc6\x9a", b"\x01\x00\x09\x00",  # WMF variants
    b"<?xml", b"<svg",
)


def _issue(issues, severity, code, detail):
    issues.append({"severity": severity, "code": code, "detail": detail})


# ---------------------------------------------------------------------------
# restricted run-property values
# ---------------------------------------------------------------------------
# ``<w:rPr>`` has children whose ``w:val`` comes from a **closed** value list
# (wml.xsd enumerations).  An out-of-list value is not "ugly formatting": Word
# 16.0 refuses to open the whole package with 「文件可能已经损坏」 (COM
# ``Documents.Open`` → ``com_error 25272``, ``wdmain11.chm``).  Every reference
# in such a file still resolves and python-docx still opens it, so nothing else
# in this script can see it.
#
# --- the governing rule ----------------------------------------------------
#
# **severity follows Word's measured behaviour, not the schema.**  A value the
# schema forbids but Word opens must not become an error: this report feeds
# ``docx_edit`` / ``docx_template`` / ``docx_mathify``, and a false error makes
# them reject a document that opens fine.  Conversely a schema-legal value that
# Word refuses must be an error — the schema's opinion changes nothing about a
# file the user cannot open.
#
# Measured has *two* dimensions, and the second one is not decoration:
#
#   (a) does ``Documents.Open`` succeed at all — decides error vs. not;
#   (b) does the written value survive a Word round trip (``SaveAs2`` then read
#       the part back) — decides whether the value is really a legal *spelling*
#       (kept verbatim), a name for the default (dropped), or content Word
#       silently replaces (a defect worth reporting even though the file opens).
#
# --- the measurements (Word 16.0, Windows, COM) ----------------------------
#
# One freshly built document per value: base ``_smoke/in.docx`` with a single
# run's ``rPr`` patched.  ``save`` is what the value looked like after Word
# re-saved the opened document.
#
#   <w:highlight w:val="darkMagenta">   opens   save darkMagenta
#   <w:highlight w:val="DARK_MAGENTA">  REFUSED (com 0x800A1529 / -2147352567)
#   <w:highlight w:val="Magenta">       REFUSED   <- the camel case is part of
#   <w:highlight w:val="purple">        REFUSED      the value: only the exact
#   <w:highlight w:val="FF0000">        REFUSED      `darkMagenta` spelling is
#   <w:highlight w:val="darkMagento">   REFUSED      in ST_HighlightColor
#   <w:highlight/>  (no w:val)          REFUSED
#   <w:vertAlign w:val="baseline">      opens   save (dropped)
#   <w:vertAlign w:val="subscript">     opens   save subscript  Subscript=-1
#   <w:vertAlign w:val="superscript">   opens   save superscript  Super=-1
#   <w:vertAlign w:val="normal">        opens   save (dropped)  Super=0 Sub=0
#   <w:vertAlign w:val="Normal">        REFUSED  <- capital N is not it either
#   <w:vertAlign w:val="SUPERSCRIPT">   REFUSED
#   <w:vertAlign w:val="Baseline">      REFUSED  <- case, again
#   <w:vertAlign w:val="bogus">         REFUSED
#   <w:vertAlign/>  (no w:val)          REFUSED
#   <w:effect w:val="blinkBackground">  opens   save blinkBackground
#   <w:effect w:val="lights">           opens   save lights
#   <w:effect w:val="sparkle">          opens   save none   <- re-interpreted
#   <w:effect w:val="none">             opens   save none
#   <w:effect w:val="ants">             REFUSED  <- schema-legal, Word refuses
#   <w:effect w:val="blink">            REFUSED  <- not in ST_TextEffect at all
#   <w:effect w:val="BlinkBackground">  REFUSED
#   <w:effect/>     (no w:val)          REFUSED
#
#   <w:u w:val="SINGLE"> / "tripleWave" opens   save (dropped) <- re-interpreted
#   <w:u/>           (no w:val)         opens   save (dropped) = single
#   <w:color w:val="#FF0000">           opens   save FF0000    <- strips the #
#   <w:color w:val="F00"> / "red" /
#            "ZZZZZZ" / "" / missing     opens   <- re-interpreted
#
# --- what that buys, per element -------------------------------------------
#
#   error   — ``w:highlight`` / ``w:vertAlign`` / ``w:effect``: an out-of-list
#             value makes Word refuse the file (dimension (a)).  ``error``
#             matches the ``empty-drawing`` / ``invalid-nesting`` precedent: the
#             same class of "the file does not open".
#   warning — ``w:u`` / ``w:color``: out-of-schema, Word opens and merely
#             re-interprets.  Still worth naming — a validator silent about a
#             hand-written ``w:color="red"`` teaches the wrong lesson — but it
#             must not change ``ok`` or the exit code.
#
# Two values moved *into* the legal sets on the strength of dimension (a) alone,
# because Word opening them is the fact the contract cares about:
#
#   * ``w:vertAlign w:val="normal"`` — **not** in ST_VerticalAlignRun, yet Word
#     opens it and treats it exactly like baseline (Superscript=0, Subscript=0),
#     then drops the element on save.  It is also ``docx_crossref.VERT_ALIGN``'s
#     spelling of "no offset at all" and therefore the default of
#     ``--vert-align``: reporting it as an error rejected the repo's own
#     documented default.  It is legal here for that reason, not because the
#     schema grew a member.  (P21; the comment block used to *say* it opens
#     while the rule table said error — comment and code disagreed.)
#   * ``w:effect w:val="ants"`` — the reverse case: it **is** in ST_TextEffect
#     and Word still refuses the package.  It is therefore *not* in the legal
#     set.  Under-reporting a value Word opens is the accepted trade; quietly
#     accepting a value Word refuses is not, so this one is an error by
#     measurement even though the schema disagrees.  (The old code listed
#     ``ants`` as legal and explained the gap in prose instead — the review was
#     right that "deliberately missed" has to look different from "legal".)
#
# --- one neighbouring check, measured, deliberately left as an error ---------
#
# ``missing-style`` (an undefined ``w:pStyle`` / ``w:rStyle``) also opens — but
# Word *substitutes* a different style id on save (measured: ``NoSuchStyleXyz``
# came back as ``aa``), i.e. the page is silently re-formatted.  That is
# dimension (b) failing, not a value Word merely re-spells, so it stays an
# error and is out of scope for the two fingerprinted sets above.
#
# --- where the value lists come from (none re-typed here) -------------------
#
#   * ``ST_HighlightColor`` — ``docx_edit.HIGHLIGHT_VALUES`` (the authority the
#     ``char-format`` command validates against; 17 values, and its comment
#     explains why ``WD_COLOR_INDEX`` is an alias source rather than the gate).
#     Note that ``w:val="none"`` is a member: it means "explicitly no
#     highlight", which is why ``none`` is legal here and *absent* ``w:val``
#     is not.
#   * ``ST_Underline`` — ``WD_UNDERLINE``'s ``xml_value``s from python-docx,
#     the same library the writer commands use.
#   * ``ST_VerticalAlignRun`` **plus** Word's tolerated ``normal`` —
#     ``docx_crossref.VERT_ALIGN``, the repo's single table for ``w:vertAlign``.
#     Both its keys and its non-``None`` values are ``w:val`` strings Word
#     opens; ``normal`` maps to ``None`` (no element) on the write side, which
#     is why taking only the values used to lose it.
#   * ``ST_TextEffect`` — not modelled anywhere in the repo, so it is spelled
#     out below, **minus** ``ants`` (measured refusal above).
#
# ``RPR_ORDER`` is imported from ``docx_crossref`` only to assert that every
# element named below is really an ``rPr`` child — the repo's single order table
# is also the single list of what an ``rPr`` may contain, so a typo such as
# ``w:higlight`` fails loudly at import time instead of silently checking a
# element that never exists.
#: ST_TextEffect minus the one value Word refuses to open (``ants``).  The
#: schema list is longer than this on purpose: see the measurement table above.
_EFFECT_VALUES = ("blinkBackground", "lights", "sparkle", "none")

_COLOR_STRICT = re.compile(r"^(?:auto|[0-9A-Fa-f]{6})$")


def _legal_highlight() -> tuple:
    """``ST_HighlightColor``, taken from `docx_edit`, never re-typed here."""
    from docx_edit import HIGHLIGHT_VALUES
    return tuple(HIGHLIGHT_VALUES)


def _legal_underline() -> tuple:
    """``ST_Underline`` from python-docx's ``WD_UNDERLINE.xml_value``s."""
    from docx.enum.text import WD_UNDERLINE
    vals = {getattr(m, "xml_value", None) for m in WD_UNDERLINE}
    return tuple(sorted(v for v in vals if v))


def _legal_vert_align() -> tuple:
    """``ST_VerticalAlignRun`` **as Word accepts it**: keys + values of ``VERT_ALIGN``.

    Both sides of that table are ``w:val`` spellings Word opens, and they are
    not the same set: ``baseline`` / ``superscript`` / ``subscript`` are the
    schema members (they appear on both sides), while ``normal`` appears only as
    a key — the write side maps it to ``None`` ("no element at all"), because
    baseline is the OOXML default.  A hand-written ``<w:vertAlign w:val="normal"/>``
    is nonetheless measured to open (Word reads it as baseline, then drops it on
    save), and ``normal`` is ``--vert-align``'s default value in
    :mod:`docx_crossref`, so taking only the table's *values* rejected the
    repo's own default spelling.  Measured, not assumed: see the block above
    ``_EFFECT_VALUES``.
    """
    from docx_crossref import VERT_ALIGN
    return tuple(sorted({v for v in VERT_ALIGN.values() if v} | set(VERT_ALIGN)))


def _restricted_run_rules():
    """``tag`` -> ``(severity, legal values, needs w:val)``.

    Built lazily so that a sibling module that will not import degrades to
    "this one check is skipped" instead of killing the whole report — the same
    contract the footer wiring below already has.
    """
    rules = {}
    try:
        rules["highlight"] = ("error", _legal_highlight(), True)
    except Exception:                              # noqa: BLE001 - by design
        pass
    try:
        rules["vertAlign"] = ("error", _legal_vert_align(), True)
    except Exception:                              # noqa: BLE001 - by design
        pass
    rules["effect"] = ("error", _EFFECT_VALUES, True)
    try:
        rules["u"] = ("warning", _legal_underline(), False)
    except Exception:                              # noqa: BLE001 - by design
        pass
    rules["color"] = ("warning", None, False)      # value checked by regex
    try:
        from docx_crossref import RPR_ORDER
    except Exception:                              # noqa: BLE001 - by design
        return rules
    for tag in list(rules):
        if tag not in RPR_ORDER:                   # pragma: no cover - typo guard
            del rules[tag]
    return rules


_RUN_PROPERTY_CODE = {
    "error": "invalid-run-property",
    "warning": "nonstandard-run-property",
}

#: ``word/*.xml`` parts whose ``<w:rPr>`` children reach Word's parser.
#: Measured: an illegal ``w:highlight`` in *any* of these makes ``Documents.Open``
#: fail — ``stylesWithEffects.xml`` and ``numbering.xml`` included.  Parts are
#: picked by name rather than by "every word/*.xml" so that ``word/comments.xml``,
#: ``word/footnotes.xml`` and the header/footer parts are covered too (they are
#: named ``word/header1.xml`` … and do carry ``rPr``), while theme/fontTable,
#: which cannot hold run properties, are skipped without a parse.
_RUN_PROPERTY_PARTS = re.compile(
    r"^word/(?:document|styles|stylesWithEffects|numbering|settings|"
    r"comments|commentsExtended|footnotes|endnotes|"
    r"header\d*|footer\d*|glossary/document)\.xml$")


def _check_run_property_values(zf, names, issues) -> None:
    """Report out-of-list ``w:val`` on restricted ``<w:rPr>`` children.

    Two decisions worth keeping:

    * **A missing ``w:val`` is not automatically a problem.**  ``<w:u/>`` means
      ``single`` (measured: Word opens it) and ``<w:color/>`` means ``auto``
      (also fine), so for those two the element is only checked when a value is
      present.  ``<w:highlight/>`` / ``<w:vertAlign/>`` / ``<w:effect/>`` have no
      schema default (each ``w:val`` is ``use="required"``) and Word refuses the
      file, so for them absence is the error.
    * **Only the element's own name is inspected.**  ``w:u`` carries ``w:color``
      too; that attribute is a *colour*, not the underline kind, so it is not
      looked up in ``ST_Underline``.
    """
    rules = _restricted_run_rules()
    if not rules:
        return
    for part in sorted(names):
        if not _RUN_PROPERTY_PARTS.match(part):
            continue
        try:
            root = etree.fromstring(zf.read(part))
        except etree.XMLSyntaxError:
            # Already reported as bad-document-xml / bad-styles-xml where it
            # matters; a second complaint about the same broken part is noise.
            continue
        val_attr = f"{{{W}}}val"
        for rpr in root.iter(f"{{{W}}}rPr"):
            for child in rpr:
                if not isinstance(child.tag, str):
                    continue                     # comment / PI node
                tag = child.tag
                if not tag.startswith(f"{{{W}}}"):
                    continue
                local = tag.split("}")[1]
                rule = rules.get(local)
                if rule is None:
                    continue
                severity, legal, needs_val = rule
                code = _RUN_PROPERTY_CODE[severity]
                value = child.get(val_attr)
                if value is None:
                    if needs_val:
                        _issue(issues, severity, code,
                               f"{part}: <w:{local}> has no w:val, which Word "
                               f"refuses to open")
                    continue
                if legal is None:                 # w:color
                    if not _COLOR_STRICT.match(value):
                        _issue(issues, severity, code,
                               f"{part}: <w:{local} w:val={value!r}> is not six "
                               f"hex digits or 'auto'; Word opens the file but "
                               f"re-interprets the colour")
                    continue
                if value not in legal:
                    _issue(issues, severity, code,
                           f"{part}: <w:{local} w:val={value!r}> is not a legal "
                           f"value; legal: {', '.join(legal)}")


def _rel_target(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _rels_owner(rels_name: str) -> str:
    """``word/_rels/document.xml.rels`` -> ``word/document.xml``.

    The part a .rels file describes is named by the file stem, *not* by its
    parent directory.  Keeping the two straight matters because a package has
    many .rels files sharing one directory (``document``, ``header1``,
    ``footer1``, ``comments``, ``footnotes``, ``endnotes``, ``fontTable`` …);
    keying them all by ``word`` makes them overwrite each other, and the winner
    then depends on set-iteration order (i.e. on ``PYTHONHASHSEED``).

    ``_rels/.rels`` describes the package itself and maps to ``""``.
    """
    rels_dir = posixpath.dirname(rels_name)
    stem = posixpath.basename(rels_name)[: -len(".rels")]
    if posixpath.basename(rels_dir) != "_rels":
        # Not a conforming location; fall back to "the directory above".
        return posixpath.normpath(posixpath.join(rels_dir, "..")).rstrip(".") \
            .strip("/")
    parent = posixpath.dirname(rels_dir)
    if not stem:  # ``_rels/.rels`` -> the package root
        return ""
    return posixpath.join(parent, stem) if parent else stem


def validate(path: str) -> dict:
    issues: list[dict] = []

    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        _issue(issues, "error", "not-a-zip", str(exc))
        return {"ok": False, "issues": issues}

    names = set(zf.namelist())
    bad = zf.testzip()
    if bad is not None:
        _issue(issues, "error", "corrupt-member", f"CRC check failed: {bad}")

    for required in ("[Content_Types].xml", "word/document.xml"):
        if required not in names:
            _issue(issues, "error", "missing-part",
                   f"required part absent: {required}")
    if issues and any(i["severity"] == "error" for i in issues):
        return {"ok": False, "issues": issues}

    # --- relationships resolve ------------------------------------------
    # Keyed by the **part that owns the .rels file**, not by its parent
    # directory.  ``word/_rels/document.xml.rels`` belongs to
    # ``word/document.xml``; ``word/_rels/header1.xml.rels`` belongs to
    # ``word/header1.xml``.  Collapsing both to ``"word"`` (the old behaviour)
    # put seven parts into one dictionary slot, so whichever .rels was read last
    # won — and since ``names`` is a *set*, that order depends on
    # ``PYTHONHASHSEED``.  The symptom was a false ``unresolved-reference`` on
    # perfectly valid files, decided at random per process: measured on
    # ``source-snapshot/test_unpack/full_test_input.docx``, seeds 0/3/4/5
    # reported two errors while seeds 1/2 reported ``ok:true``.
    rel_ids_by_source: dict[str, dict] = {}
    for rels_name in sorted(names):
        if not rels_name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(zf.read(rels_name))
        except etree.XMLSyntaxError as exc:
            _issue(issues, "error", "bad-rels-xml", f"{rels_name}: {exc}")
            continue
        source_part = _rels_owner(rels_name)
        ids = {}
        for rel in root.iter(f"{{{PR}}}Relationship"):
            rid, target = rel.get("Id"), rel.get("Target", "")
            mode = rel.get("TargetMode", "Internal")
            ids[rid] = target
            if mode == "External":
                continue
            # ``source_part`` is the *owning part* (a file path, e.g.
            # ``word/document.xml``); ``_rel_target`` resolves against its
            # directory.  The old ``source_part + "/x"`` hack only existed
            # because the old key was a bare directory.
            resolved = _rel_target(source_part, target)
            if resolved not in names:
                _issue(issues, "error", "dangling-rel",
                       f"{rels_name}: {rid} -> {target} (missing part)")
        # Two .rels files never own the same part, so this is a plain insert —
        # but merge rather than overwrite, so a duplicated member (legal in a
        # ZIP, only the first is read) cannot silently drop ids.
        rel_ids_by_source.setdefault(source_part or "_package", {}).update(ids)

    # --- r:id / r:embed references in document.xml -----------------------
    # An XMLSyntaxError here means the *package* is real but its main part is
    # not well-formed — a document problem, not a usage error.  It is folded
    # into one issue so the JSON contract on stdout survives and the caller
    # gets ``ok:false`` + exit 1 instead of nineteen lines of lxml internals.
    # Every later check reads ``doc_root``, so this returns immediately.
    try:
        doc_root = etree.fromstring(zf.read("word/document.xml"))
    except etree.XMLSyntaxError as exc:
        _issue(issues, "error", "bad-document-xml",
               f"word/document.xml: {exc}")
        return {"ok": False, "issues": issues}
    doc_rels = rel_ids_by_source.get("word/document.xml", {})
    for el in doc_root.iter():
        for attr in (f"{{{R}}}id", f"{{{R}}}embed", f"{{{R}}}link"):
            rid = el.get(attr)
            if rid and rid not in doc_rels:
                _issue(issues, "error", "unresolved-reference",
                       f"document.xml references {rid} with no relationship")

    # --- structural sanity of the body ------------------------------------
    # Everything above resolves *references*; none of it looks at the shape of
    # the markup.  Two shapes are worth naming because Word rejects them while
    # the file still passes every check above:
    #
    #   * ``<w:drawing/>`` with no ``wp:inline`` / ``wp:anchor`` — the classic
    #     after-effect of deleting an inline shape while a picture part is
    #     replaced: the run keeps an empty drawing, and Word refuses to open
    #     the document with "unreadable content".
    #   * a block element (``<w:p>`` / ``<w:tbl>``) directly inside a run.
    #
    # Scope discipline: only these two, only in document.xml, and only when the
    # shape is unambiguous.  A drawing with ``wp:anchor`` but no ``wp:inline`` is
    # legal (floating objects), and a drawing inside ``mc:AlternateContent``
    # still has one of the two children — so neither is touched.
    for drawing in doc_root.iter(f"{{{W}}}drawing"):
        if (drawing.find(f"{{{WP}}}inline") is None
                and drawing.find(f"{{{WP}}}anchor") is None):
            _issue(issues, "error", "empty-drawing",
                   "document.xml: <w:drawing> has neither wp:inline nor "
                   "wp:anchor; Word refuses the file")

    for run in doc_root.iter(f"{{{W}}}r"):
        for child_tag in RUN_FORBIDDEN_CHILDREN:
            for child in run.iterchildren(child_tag):
                local = child_tag.split("}")[1]
                _issue(issues, "error", "invalid-nesting",
                       f"document.xml: <w:{local}> nested directly inside "
                       f"<w:r>; that is not valid run content")

    # --- embedded images decode ------------------------------------------
    for name in [n for n in names
                 if n.startswith("word/media/") and not n.endswith("/")]:
        data = zf.read(name)
        if not data:
            _issue(issues, "error", "empty-image", name)
        elif not any(data.startswith(m) for m in IMAGE_MAGIC):
            _issue(issues, "warning", "unknown-image-format",
                   f"{name}: unrecognized magic bytes")

    # --- styles referenced exist ------------------------------------------
    defined = set()
    if "word/styles.xml" in names:
        # Same contract as document.xml above: a broken styles part is a
        # document problem (issue + exit 1), never a traceback.
        try:
            styles_root = etree.fromstring(zf.read("word/styles.xml"))
        except etree.XMLSyntaxError as exc:
            _issue(issues, "error", "bad-styles-xml",
                   f"word/styles.xml: {exc}")
            styles_root = None
        if styles_root is not None:
            defined = {s.get(f"{{{W}}}styleId")
                       for s in styles_root.iter(f"{{{W}}}style")}
    for tag, attr in ((f"{{{W}}}pStyle", f"{{{W}}}val"),
                      (f"{{{W}}}rStyle", f"{{{W}}}val"),
                      (f"{{{W}}}tblStyle", f"{{{W}}}val")):
        for el in doc_root.iter(tag):
            sid = el.get(attr)
            if sid and sid not in defined:
                _issue(issues, "error", "missing-style",
                       f"style id referenced but not defined: {sid}")

    # --- restricted run-property values -----------------------------------
    # Runs after the structural checks and before the footer wiring: it needs
    # the package's part list only, so it works even when document.xml parsed
    # but a header did not (those parts are skipped individually above).
    _check_run_property_values(zf, names, issues)

    # --- footer semantics -------------------------------------------------
    # Everything above asks "does this name resolve to a part" and "is this
    # markup shaped legally".  None of it asks whether the *content* of a
    # footer survives being rendered, and a footer pushed into a 4.56 pt wide
    # floating text box passes every check in this file: legal XML, resolving
    # rels, a real PAGE field.  Measured on the document that started P17:
    # ``{"ok": true, "issues": []}`` while the page number was clipped at two
    # digits.  So the whole footer semantics live in `docx_footer` and this is
    # only the wiring — one implementation, two entry points.
    #
    # "Cannot read it" must not become a traceback: `footer_issues` swallows
    # its own failures into a warning, so a broken footer part degrades to
    # ``footer-check-skipped`` instead of losing the report.  The import is
    # guarded for the same reason — a sibling module that will not import (an
    # incomplete install, a name that moved) must not turn every document into
    # a crash of *this* script.
    try:
        from docx_footer import footer_issues
    except Exception as exc:                       # noqa: BLE001 - by design
        _issue(issues, "warning", "footer-check-skipped",
               "docx_footer unavailable (%s: %s)" % (type(exc).__name__, exc))
    else:
        issues.extend(footer_issues(path))

    # --- python-docx can open it ------------------------------------------
    try:
        from docx import Document
        Document(path)
    except Exception as exc:
        _issue(issues, "error", "python-docx-open-failed", str(exc))

    ok = not any(i["severity"] == "error" for i in issues)
    return {"ok": ok, "issues": issues}


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Health-check a .docx (not XSD schema validation).")
    ap.add_argument("path", help="the .docx file to check")
    args = ap.parse_args()

    # ---- 双契约：输入错误退 2，内容问题仍退 1（P9-C）---------------------
    # 本脚本是全仓唯一一个「1 有正当语义」的 CLI：`ok:false` + `issues` 是它的
    # 正常产出，`1` 表示**文档是真的、但内容有 error**。所以这里只把「输入根本
    # 不是包」这三档提到 2，**内容那一档的 JSON 输出一个字都不动**。
    #
    # 旧行为（实测 2026-11-11）：三档全部落进 `validate()` 的 `zipfile.ZipFile`
    # except 分支，被收成一条 `not-a-zip` issue → **exit 1 + JSON**。于是
    # 「路径写错了」和「文档坏了」在调用方眼里是同一件事，唯一区别是 issue 的
    # code —— 而任何 `if rc: 报错` 的调用方都会把它当成文档问题去修文档。
    #
    # stdout 在拒绝路径上必须为空：成功与「内容有问题」两条路都会打 JSON，
    # 空的 stdout 才不会被 `json.loads` 读成一个空报告。
    try:
        _check_input(args.path)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    report = validate(args.path)
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
