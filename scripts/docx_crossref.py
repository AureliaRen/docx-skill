#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Cross-references and caption numbering: bookmarks, SEQ / REF fields.

A cross-reference in Word is two fields wired to a bookmark:

  ``SEQ 图 \\* ARABIC``     the *caption counter*. Each caption of the same
                            label shares one running counter; Word renumbers
                            them all when one is added or removed.
  ``REF fig_1 \\h``         the *pointer*. Displays the bookmarked text and,
                            with ``\\h``, behaves as a hyperlink.

Both are "stale" until Word recalculates fields (F9 / print), so every field
is written with a cached result: the document shows the right number even
before the user updates anything, and it stays correct when they do.

Clickable references (why the reader has to *click* them)
--------------------------------------------------------
``\\h`` is only a field *switch*: it tells Word to treat the recalculated
result as a hyperlink, and that hyperlink exists in Word's memory rather than
in the part we write.  Measured on Word (COM): a bare ``REF bm \\h`` field
leaves ``Document.Hyperlinks.Count`` at **0**, so a reader who opens the file
and clicks the number clicks nothing.  ``cmd_ref`` therefore also wraps the
whole field in a real hyperlink element::

    <w:hyperlink w:anchor="图_1" w:history="1">
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText> REF </w:instrText></w:r>
      <w:r><w:instrText>图_1 \\h </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>1</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:hyperlink>

Word keeps that shape across an edit/save round trip, and reports
``Hyperlinks.Count`` = 2 with ``Hyperlinks(i).SubAddress`` = the bookmark name
— the first *programmatic* proof that a cross-reference is clickable.  The
inner ``\\h`` stays as well: belt and braces, and it is what Word itself writes
for its UI "Insert as hyperlink" path.  ``--no-hyperlink`` writes the bare
field (the pre-v2 behaviour).

Superscript / normal size
-------------------------
Every run of the field — begin, instrText, separate, cached result, end —
carries ``<w:vertAlign w:val="superscript">`` under ``--vert-align
superscript`` (or the ``--superscript`` shorthand), ``subscript`` for the
lowered form, and nothing at all for ``normal``.  ``w:vertAlign`` has exactly
one legal slot in ``CT_RPr`` (an ``xsd:sequence``: after ``sz``/``u``/``shd``,
before ``rtl``/``lang``); written anywhere else Word silently drops it.

The field also carries ``\\* MERGEFORMAT`` by default, which is what makes the
raised/lowered look survive Word's own recalculation: without it Word
re-renders the result from the bookmark target and a superscript reference
turns into ordinary text on the next F9 / print / PDF export (measured on Word,
see MERGE_FORMAT).  ``--no-mergeformat`` omits it.

Cached results mirror Word
--------------------------
Every field is written with a cached result, so the document is readable
before Word recalculates anything.  The cached text is derived the way Word
derives the real one (see REF_FORMATS for the oracle behind each entry):

  figure number   ``REF 图_1 \\h``            -> the bookmarked number, "1"
  whole caption   ``REF cap \\h`` + wide bm   -> "图 1  系统架构示意"
  in-text [1]     ``REF ref_1 \\n``           -> the list number, "[1]"
  above / below   ``REF 图_1 \\p``            -> left empty: Word and the
  page number     ``PAGEREF 图_1 \\h``           locale decide (上方/下方, page)

Subcommands:
  caption    add "图 N  题注文本" to a par/table, with a SEQ field + bookmark
  ref        insert a REF field pointing at an existing bookmark
  bookmark   put a bare bookmark on a paragraph (to reference prose)
  renumber   recompute SEQ counters and refresh every cached REF result
  list       show bookmarks, SEQ counters, REF targets and hyperlink anchors

Examples:
  docx_crossref.py caption in.docx --index 42 --label 图 --text "系统架构示意"
  docx_crossref.py ref in.docx --index 70 --target fig_1 --prefix "如图" --suffix 所示
  docx_crossref.py ref in.docx --index 70 --target tab_2 --format above-below
  docx_crossref.py ref in.docx --index 12 --target ref_1 --format paragraph-number --superscript
  docx_crossref.py ref in.docx --index 12 --target fig_1 --no-hyperlink
  docx_crossref.py renumber in.docx -o out.docx
  docx_crossref.py list in.docx

Exit codes: 0 success, 1 expected failure (unknown bookmark), 2 usage error.
"""
from __future__ import annotations
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix as _encoding_fix  # noqa: F401

import argparse
import copy
import json
import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_utils import PPR_ORDER, insert_ordered, is_cjk, next_bm_id

XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Which Word field each cross-reference flavour really is.
#
# Every entry below was read off Word itself (COM oracle: ``Fields.Add`` +
# ``Code.Text`` / ``Result.Text``, plus the UI path
# ``Range.InsertCrossReference``); the raw output is in
# work/agents/xref/REPORT.md.  The mapping is not cosmetic:
#
#   * "page number" is a **PAGEREF** field.  ``REF <bm> \p`` is the
#     "above/below" switch and returns "above"/"below" (= 上方/下方), not a page.
#   * ``\n`` / ``\r`` / ``\w`` are *paragraph number* flavours: for a numbered
#     reference list item they render its list number **in its own format**
#     (``[1]`` for a ``[%1]`` list), which is exactly what an in-text citation
#     wants; they are not "the whole caption".
#   * "entire caption" is not a switch at all, it is bookmark *coverage*: a
#     plain ``REF <bm> \h`` returns whatever text the bookmark spans, so
#     ``caption --bookmark-span paragraph`` + ``ref --format full`` yields the
#     whole caption ("图 1  系统架构示意"), verified on Word.
REF_FORMATS = {
    "number": ("REF", ""),                  # the text the bookmark spans
    "default": ("REF", r" \h"),             # same, plus the hyperlink switch
    "above-below": ("REF", r" \p"),         # "above"/"below" (上方/下方)
    "page": ("PAGEREF", r" \h"),            # the page the bookmark sits on
    "paragraph-number": ("REF", r" \n"),    # list number, e.g. [1]
    "relative-context": ("REF", r" \r"),    # list number + numbering context
    "full-context": ("REF", r" \w"),        # list number, full context
    "full": ("REF", r" \h"),                # whole caption (needs a wide bm)
}
# Kept for callers that only want the switch text.
REF_SWITCHES = {k: v[1] for k, v in REF_FORMATS.items()}

# Flavour -> how the cached result is derived before Word recalculates:
#   "text"    the text the bookmark spans (narrow caption bm -> "1")
#   "number"  the paragraph's own list number with its format, e.g. "[1]"
#   "word"    Word/locale dependent ("above", a page number): leave empty
CACHED_KIND = {
    "number": "text", "default": "text", "full": "text",
    "paragraph-number": "number", "relative-context": "number",
    "full-context": "number",
    "above-below": "word", "page": "word",
}

# CLI/API name -> the ``w:val`` written into ``<w:vertAlign>``.  ``normal``
# means "no vertAlign element at all" (baseline is the default in OOXML, and
# an explicit baseline element is pointless noise).
VERT_ALIGN = {"normal": None, "baseline": "baseline",
              "superscript": "superscript", "subscript": "subscript"}

# CT_RPr children, in schema order (wml.xsd).  ``vertAlign`` sits between
# ``fitText`` and ``rtl``; a run property written out of order is dropped by
# Word without a warning.
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]

# ``\* MERGEFORMAT`` tells Word to keep the field result's formatting when it
# recalculates.  Without it a superscript reference silently becomes normal
# text the first time Word updates fields (F9, print, or a PDF export) — the
# runs we wrote still carry ``w:vertAlign``, but the *result* is re-rendered
# from the bookmark target's formatting.  Measured on Word (COM):
#
#   REF 图_1 \h                    -> Fields(k).Result.Font.Superscript  1 -> 0
#   REF 图_1 \h \* MERGEFORMAT     ->                                  1 -> 1
#
# and the exported PDF agrees (11pt baseline before, 7pt raised after).  The
# switch changes no result *text* for any flavour (verified for ``\h``, ``\p``,
# ``\n``, ``\w`` and PAGEREF pairs), so it is on by default.
MERGE_FORMAT = r" \* MERGEFORMAT"

# Bookmark names may not start with a digit and are capped at 40 chars.
_NAME_SAFE = re.compile(r"[^0-9A-Za-z_\u4e00-\u9fff]")

_INSTR = qn("w:instrText")
_FLDCHAR = qn("w:fldChar")
_T = qn("w:t")
_FIELD_NODES = (_INSTR, _FLDCHAR, _T)


def bm_name(label: str, n: int) -> str:
    """Stable, Word-legal bookmark name for a caption, e.g. ``fig_1``."""
    slug = _NAME_SAFE.sub("_", label).strip("_") or "item"
    return f"{slug}_{n}"[:40]


def with_vert_align(rpr, vert: str | None):
    """A copy of *rpr* carrying ``<w:vertAlign w:val="vert">`` in schema order.

    Any ``w:vertAlign`` already on *rpr* is dropped first, so *vert* ``None``
    means *normal size* rather than "whatever the run we inherited from
    happened to be".  Returns ``None`` when nothing is left to set.
    """
    out = copy.deepcopy(rpr) if rpr is not None else OxmlElement("w:rPr")
    old = out.find(qn("w:vertAlign"))
    if old is not None:
        out.remove(old)
    if vert:
        va = OxmlElement("w:vertAlign")
        insert_ordered(out, va, RPR_ORDER)
        va.set(qn("w:val"), vert)
    return out if len(out) else None


def _text_run(text: str, rpr=None):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement("w:t")
    t.set(XMLSPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def _instr_run(instr: str, rpr=None):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    it = OxmlElement("w:instrText")
    it.set(XMLSPACE, "preserve")
    it.text = instr
    r.append(it)
    return r


def _fld_run(kind: str, rpr=None):
    r = OxmlElement("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    f = OxmlElement("w:fldChar")
    f.set(qn("w:fldCharType"), kind)
    r.append(f)
    return r


def field_elements(instr: str, cached: str, rpr=None, vert: str | None = None) -> list:
    """A complete complex field with a cached result.

    ``begin / instrText / separate / result / end`` — the ``separate`` and
    result runs are what make the number visible before Word updates fields.

    *vert* applies ``<w:vertAlign>`` to **every** run of the field, which is
    what makes a superscript reference stay raised instead of inheriting the
    paragraph's baseline.
    """
    rpr = with_vert_align(rpr, vert)
    return [_fld_run("begin", rpr), _instr_run(f" {instr.strip()} ", rpr),
            _fld_run("separate", rpr), _text_run(cached, rpr),
            _fld_run("end", rpr)]


def hyperlink_wrap(anchor: str, children: list):
    """Wrap *children* in ``<w:hyperlink w:anchor="anchor" w:history="1">``.

    This is the element Word reads to build ``Hyperlinks(i)``; the anchor is an
    internal bookmark, so the link resolves inside the document (no
    relationship part needed).
    """
    h = OxmlElement("w:hyperlink")
    h.set(qn("w:anchor"), anchor)
    h.set(qn("w:history"), "1")
    for c in children:
        h.append(c)
    return h


def _para_rpr(para):
    """Run properties of the paragraph's first formatted run.

    Walks ``w:r`` descendants rather than ``para.runs``: after a reference has
    been inserted, the only run in the paragraph may live inside a
    ``w:hyperlink``, and python-docx's ``.runs`` skips those.
    """
    for r in para._p.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            return rpr
    pPr = para._p.pPr
    if pPr is not None:
        return pPr.find(qn("w:rPr"))
    return None


def _bookmark_pair(name: str, bid: int):
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    return start, end


def _wrap_whole_paragraph(p, name: str, bid: int):
    """Bracket the *whole* paragraph with a bookmark named *name*.

    ``bookmarkStart`` lands immediately after ``w:pPr`` (so it precedes the
    label run) and ``bookmarkEnd`` goes last.  That is the coverage Word's own
    caption machinery uses: ``REF <name>`` returns the entire caption text, and
    a table of figures' ``PAGEREF`` lands on it instead of dangling.
    """
    pPr = p.find(qn("w:pPr"))
    start, end = _bookmark_pair(name, bid)
    if pPr is not None:
        pPr.addnext(start)
    else:                                                    # pragma: no cover
        p.insert(0, start)
    p.append(end)
    return start, end


def all_bookmarks(doc) -> set:
    root = doc.element.body
    return {b.get(qn("w:name")) for b in root.iter(qn("w:bookmarkStart"))
            if b.get(qn("w:name"))}


# ---- caption bookmarks and paragraph style ---------------------------------
#
# Word's "Insert Table of Figures" (``TOC \c "图"``) does not read the caption
# text: it builds one ``PAGEREF`` per entry, and every ``PAGEREF`` points at a
# bookmark.  Word's own caption machinery names that family ``_Ref`` + a number
# from a document-wide counter, which is also what it puts round a caption when
# you insert a cross-reference to it.  A caption with only a hand-chosen
# bookmark name (``fig_1``) inserts into the table of figures but every entry
# is dead on click, because no ``_Ref…`` anchor exists to point at.
#
# Measured on the reference document (``work/src.docx``, a Word-saved file):
#   * 17 ``_Ref…`` bookmarks, ids 133..149, names ``_Ref3925`` … ``_Ref8179``
#     (four-digit Word counters, *not* consecutive);
#   * every one of them sits on a ``pStyle=27`` reference-list paragraph, not
#     on a caption — they are the citation targets behind ``REF _Ref… \r \h``;
#   * the 12 caption paragraphs (``pStyle=28``, style name 图表名) carry
#     ``_Toc…`` bookmarks, and those bookmarks start *after* the label + SEQ
#     field, so they span the caption prose only.
# So this module does not copy a caption ``_Ref`` shape out of the original —
# there is none to copy.  It produces the shape Word produces for a caption
# that a table of figures can link to: a ``_Ref<number>`` bookmark covering the
# whole caption, side by side with whatever bookmark the caller asked for.
_REF_NUM = re.compile(r"_Ref(\d+)\Z")
_REF_RESERVED_ATTR = "_docx_crossref_ref_reserved"


def _reserved_set(doc, attr: str) -> set:
    """The set of names already handed out but not yet written into *doc*.

    Kept on the document object so a caller can mint a whole batch before
    inserting any of them.  Without it, ``next_*_name`` reads only the bookmarks
    that are already in the XML and returns the same name twice -- which is
    exactly the duplicate this module now refuses to write
    (work/REVIEW-ADVERSARIAL-R6.md 应修-2).  A reloaded document simply starts
    with an empty set, which is correct: nothing was minted for it yet.
    """
    reserved = getattr(doc, attr, None)
    if not isinstance(reserved, set):
        reserved = set()
        try:
            setattr(doc, attr, reserved)
        except Exception:                                  # pragma: no cover
            reserved = set()
    return reserved


def _mint(doc, pattern, head: str, attr: str, pad: int = 0,
          reserve: bool = True) -> str:
    """Shared minting loop: next free ``head<number>``.

    ``pad`` zero-pads the number to that width (``0`` = no padding).  The TOC
    family pads to four digits so the generated anchors sort lexicographically;
    the ``_Ref`` family does not, because its names are already in the wild as
    ``_Ref1``/``_Ref2`` and changing the shape would break existing documents
    and their tests.

    ``reserve`` controls the batch guarantee.  ``True`` remembers the name so
    the next call cannot hand out the same one -- necessary for a caller that
    mints a batch before writing any of them.  ``False`` is the plain "next free
    name given what is in the document right now" reading, which is what a
    caller wants when it writes each name immediately: two successive calls then
    answer the same question and get the same answer.
    """
    used = all_bookmarks(doc)
    reserved = _reserved_set(doc, attr) if reserve else set()

    top = 0
    for name in used | reserved:
        m = pattern.match(name or "")
        if m:
            top = max(top, int(m.group(1)))

    n = top + 1
    while True:
        tail = ("%0*d" % (pad, n)) if pad else str(n)
        # Keep the whole name inside the cap even for an absurdly long prefix,
        # so a caller-supplied prefix cannot smuggle in a 41-character name.
        room = BOOKMARK_MAX_LEN - len(tail)
        candidate = (head[:room] if len(head) > room else head) + tail
        if candidate not in used and candidate not in reserved:
            break
        n += 1
    if reserve:
        reserved.add(candidate)
    return candidate


def minted_ref_names(doc, count: int) -> list:
    """``count`` distinct ``_Ref`` names for a batch insert.

    Use this instead of calling :func:`next_ref_name` in a loop when the names
    are collected first and written afterwards: same underlying counter, with
    the batch guarantee switched on so no two names can repeat.
    """
    return [_mint(doc, _REF_NUM, "_Ref", _REF_RESERVED_ATTR) for _ in range(count)]


def next_ref_name(doc) -> str:
    """The next free ``_Ref<number>`` name, unique against what *doc* holds.

    The counter continues from the highest ``_Ref`` already present, so a
    document that already carries Word-generated ones never collides with them,
    and repeated runs keep marching upwards instead of reusing a name.

    Deliberately *not* reserving: a name is only taken once it is in the
    document, so asking twice without writing anything gets the same answer.
    Call :func:`minted_ref_names` for the batch case.
    """
    return _mint(doc, _REF_NUM, "_Ref", _REF_RESERVED_ATTR, reserve=False)


# ---- bookmark-name guards ---------------------------------------------------
#
# Measured on Word 16.0 and LibreOffice 26.8.2.2 (``work/T3-REPORT.md`` §3.1-§3.4):
# all three defects below are **silent**, and none of them is visible in the XML
# we write — the damage only appears after Word or LO saves the file, so the
# check has to happen at generation time:
#
#   duplicates (§3.1)  Word *merges* them: ``ShowHidden`` reports 9 bookmarks for
#                      10 starts and the surviving Start is the later one, so a
#                      table-of-contents entry jumps to the wrong chapter.  LO
#                      *renames* the second one to ``<name>_副本_1``, so every
#                      ``PAGEREF`` written against the original name goes dead.
#   spaces (§3.2)      both applications replace a space with ``_`` on save, so
#                      ``REF _Ref图 1 \h`` is read as "REF _Ref图" plus a literal
#                      1 and the field breaks.
#   length (§3.3)      both truncate the name to 40 characters.
#
# The guard therefore lives here, next to the code that mints names.
BOOKMARK_MAX_LEN = 40

# Legal characters: **everything except the control characters Word itself
# rejects**.  An allow-list of ``[A-Za-z0-9_]`` + CJK was tried first and was
# wrong -- it rejected names Word accepts verbatim and resolves correctly.
# Measured (Word 16.0, open + save + read back, then ``Fields.Update()`` on a
# ``REF`` pointing at each): 21 of 22 sampled names survived byte-identical and
# resolved, including ``图(1)``, ``图（1）``, ``图1：说明``, ``图[1]``, ``图【1】``,
# ``第〇章``, full-width ``Ａ``, ``图·1``, ``图—1``, ``图+1``, ``图%1``,
# ``图@1``, ``图#1``, ``图*1``, ``图♀1``, ``图°1``.  Only two classes really are
# rewritten, and those are the two checks below:
#   * whitespace is silently turned into ``_`` (T3 §3.2);
#   * anything past 40 characters is silently truncated (T3 §3.3).
# So the character test is a **black list** of control characters, which is what
# ``ST_String`` actually forbids.  See work/REVIEW-ADVERSARIAL-R6.md 阻塞-1.
_NAME_FORBIDDEN = frozenset(chr(c) for c in range(0x00, 0x20)) | {"\x7f"}
_TOC_RESERVED_ATTR = "_docx_crossref_toc_reserved"


def _bookmark_char_ok(ch: str) -> bool:
    """True for characters Word keeps in a bookmark name.

    Only control characters are rejected: everything printable (ASCII, CJK,
    punctuation, full-width forms) survives a Word and a LibreOffice round trip
    byte for byte, so treating it as illegal would block working names.
    """
    return ch not in _NAME_FORBIDDEN


def validate_bookmark_name(name, *, max_len: int = BOOKMARK_MAX_LEN) -> list[str]:
    """Return the problems that make *name* unsafe as a bookmark name.

    An empty list means the name is safe.  Every problem reported here is one
    that Word and/or LibreOffice fix up **silently** (``work/T3-REPORT.md``
    §3.1-§3.3), which is why it has to be caught before the file is written:
    the XML is accepted either way, and the broken ``REF``/``PAGEREF`` fields
    only show up in the reader's hands.
    """
    if not isinstance(name, str):
        return [f"name must be a string, got {type(name).__name__}"]
    if not name:
        return ["name is empty"]

    problems: list[str] = []

    bad_space = sorted({c for c in name if c.isspace()})
    if bad_space:
        shown = ", ".join(f"U+{ord(c):04X}" for c in bad_space)
        problems.append(
            f"name contains whitespace ({shown}); Word and LibreOffice both "
            f"silently rewrite it to '_' on save (T3 §3.2)")

    if len(name) > max_len:
        problems.append(
            f"name is {len(name)} characters, over the {max_len} limit; Word "
            f"and LibreOffice both silently truncate to {max_len} (T3 §3.3)")

    bad_chars = sorted({c for c in name
                        if not c.isspace() and not _bookmark_char_ok(c)})
    if bad_chars:
        shown = ", ".join(
            f"{c!r} (U+{ord(c):04X})" if c.isprintable() else f"U+{ord(c):04X}"
            for c in bad_chars)
        problems.append(
            f"name contains control characters, which ST_String forbids: "
            f"{shown}")

    return problems


def validate_bookmark_uniqueness(doc) -> dict:
    """Count the bookmark names in *doc* and report any that repeat.

    Duplicates are the one defect no XML audit can see: Word merges them on
    open (the later Start wins, so a TOC entry jumps to the wrong chapter) and
    LibreOffice renames the second one ``<name>_副本_1`` (so every field written
    against the original name dies).  Both are silent — ``work/T3-REPORT.md``
    §3.1 — so ``ok`` is False the moment any name repeats.

    Returns ``{"duplicates": [...], "count": n, "unique": m, "ok": bool}``:
    *count* counts every named ``bookmarkStart``, *unique* the distinct names.
    """
    seen: dict = {}
    for b in doc.element.body.iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name"))
        if name:
            seen[name] = seen.get(name, 0) + 1
    duplicates = sorted(n for n, k in seen.items() if k > 1)
    return {"duplicates": duplicates,
            "count": sum(seen.values()),
            "unique": len(seen),
            "ok": not duplicates}


def next_toc_name(doc, prefix: str = "_Toc") -> str:
    """A fresh ``_Toc0001``-shaped bookmark name, padded to four digits.

    The shape is what keeps a TOC anchor legal without further thinking: no
    whitespace, ASCII only, and short enough for the 40-character cap that Word
    and LibreOffice both impose (``work/T3-REPORT.md`` §3.2-§3.3) — and the
    four-digit zero padding gives 9999 anchors that sort lexicographically,
    which is handy when a TOC field lists them.

    Names already present in *doc* are skipped, so a re-run never collides.
    Names this function *handed out* but that have not been written into the
    document yet are remembered on the document object and skipped too, so a
    loop that mints a batch of names before inserting any of them cannot emit
    the same name twice.
    """
    used = all_bookmarks(doc)

    reserved = getattr(doc, _TOC_RESERVED_ATTR, None)
    if not isinstance(reserved, set):                      # first call, or a
        reserved = set()                                   # reloaded document
        try:
            setattr(doc, _TOC_RESERVED_ATTR, reserved)
        except Exception:                                  # pragma: no cover
            reserved = set()

    # A caller-supplied prefix has to be sanitised, not just checked: Word
    # rewrites whitespace to ``_`` on save, so a name containing a space would
    # change identity between the file we write and the file the reader opens,
    # and any field written against the original string would die.  Control
    # characters are replaced for the same reason (``ST_String`` forbids them).
    head = "".join(
        "_" if (c.isspace() or not _bookmark_char_ok(c)) else c
        for c in (prefix or ""))
    if not head:
        head = "_Toc"

    pattern = re.compile(re.escape(head) + r"(\d+)\Z")
    return _mint(doc, pattern, head, _TOC_RESERVED_ATTR, pad=4)


def set_paragraph_style(p, style_id: str):
    """Set ``<w:pStyle w:val="style_id"/>`` as the first child of ``w:pPr``.

    ``CT_PPr`` is an ``xsd:sequence`` and ``pStyle`` is its first member; a
    ``pStyle`` written anywhere else is dropped by Word without a warning — the
    paragraph then silently keeps its old style, and "Insert Table of Figures"
    (which filters by paragraph style) cannot see the caption at all.

    *style_id* must be a ``w:styleId``, not a style *name*: the reference
    document calls its caption style ``28`` while the name behind it is 图表名,
    and the number differs per document — so it is a parameter, never a
    constant.
    """
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    el = OxmlElement("w:pStyle")
    el.set(qn("w:val"), str(style_id))
    return insert_ordered(pPr, el, PPR_ORDER)


def bookmark_paragraphs(doc) -> dict:
    """``{bookmark name: the <w:p> it starts in}``."""
    out: dict = {}
    for p in doc.element.body.iter(qn("w:p")):
        for b in p.iter(qn("w:bookmarkStart")):
            name = b.get(qn("w:name"))
            if name and name not in out:
                out[name] = p
    return out


def paragraph_text(p) -> str:
    """Visible text of a paragraph element (nested runs included)."""
    return "".join(t.text or "" for t in p.iter(_T))


def bookmark_span_text(p, name: str):
    """The text covered by bookmark *name* inside *p*, or None when absent."""
    bid = None
    parts: list = []
    for e in p.iter():
        if e.tag == qn("w:bookmarkStart") and e.get(qn("w:name")) == name:
            bid = e.get(qn("w:id"))
            parts = []
            continue
        if bid is None:
            continue
        if e.tag == _T:
            parts.append(e.text or "")
        elif e.tag == qn("w:bookmarkEnd") and e.get(qn("w:id")) == bid:
            return "".join(parts)
    return None


def bookmarked_text(doc, name: str):
    """The text a plain ``REF <name>`` displays: what the bookmark spans.

    Narrow caption bookmarks (``cmd_caption``'s default) span the ``SEQ``
    field, so this is the number; ``--bookmark-span paragraph`` makes it the
    whole caption, which is how Word's own "entire caption" reference works.
    """
    p = bookmark_paragraphs(doc).get(name)
    if p is None:
        return None
    return bookmark_span_text(p, name)


def _is_wide_caption(p, name: str) -> bool:
    """True when the bookmark spans a whole caption, not just its SEQ field.

    Word's own ``_Ref…`` caption bookmarks cover label + number + caption text;
    ours used to cover the SEQ field alone.  The distinction decides whether
    ``REF <name> \\h`` shows "1" or "图 1  系统架构示意", and it is read off the
    markup — a SEQ field inside the span *plus* text covered outside it.
    """
    bid = None
    in_seq = False
    saw_seq = False
    outside_text = False
    for e in p.iter():
        if e.tag == qn("w:bookmarkStart") and e.get(qn("w:name")) == name:
            bid = e.get(qn("w:id"))
            continue
        if bid is None:
            continue
        if e.tag == qn("w:bookmarkEnd") and e.get(qn("w:id")) == bid:
            break
        if e.tag == _INSTR and re.match(r"\s*SEQ\s+", e.text or ""):
            saw_seq = True
            in_seq = True
            continue
        if e.tag == _FLDCHAR and in_seq \
                and e.get(qn("w:fldCharType")) == "end":
            in_seq = False
            continue
        if e.tag == _T and not in_seq:
            outside_text = True
    return saw_seq and outside_text


def content_text_result(doc, name: str) -> str:
    """What Word displays for ``REF <name>`` / ``REF <name> \\h``.

    * a bookmark spanning a whole caption -> the caption text;
    * otherwise the number the bookmark name declares (``图_1`` -> "1"), which
      is this project's naming convention for captions and reference items;
    * otherwise the text the bookmark spans.

    Used both when the field is written and by ``renumber``, so the cached
    result and a later Word recalculation agree.
    """
    p = bookmark_paragraphs(doc).get(name)
    if p is not None and _is_wide_caption(p, name):
        txt = (bookmark_span_text(p, name) or "").strip()
        if txt:
            return txt
    m = re.search(r"_(\d+)$", name)
    if m:
        return m.group(1)
    if p is None:
        return ""
    return (bookmark_span_text(p, name) or "").strip()


def _numpr_of(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    nid = numPr.find(qn("w:numId"))
    lvl = numPr.find(qn("w:ilvl"))
    return (int(nid.get(qn("w:val"))) if nid is not None else 0,
            int(lvl.get(qn("w:val"))) if lvl is not None else 0)


def _format_lvl_text(lvl_text: str, n: int) -> str:
    out = lvl_text
    for i in range(9, 0, -1):
        out = out.replace("%" + str(i), str(n) if i == 1 else "1")
    return out


def list_number_for(doc, p) -> str:
    """The paragraph's rendered list number (``[1]``), or "" when unnumbered.

    Mirrors Word's ``ListFormat.ListString`` for the flat single-level lists
    this skill writes: the ``lvlText`` of the level with ``%1`` replaced by the
    item's ordinal.  Cross-checked against Word: a ``[%1]`` list yields ``[1]``
    for the first item, which is what ``REF <bm> \\n`` shows for it.
    """
    cur = _numpr_of(p)
    if cur is None:
        return ""
    num_id, ilvl = cur
    try:
        from docx_numbering import numbering_root
        root = numbering_root(doc)
    except Exception:                                        # noqa: BLE001
        return ""
    if root is None:
        return ""
    abstract_id = None
    for e in root.findall(qn("w:num")):
        if e.get(qn("w:numId")) == str(num_id):
            ref = e.find(qn("w:abstractNumId"))
            abstract_id = None if ref is None else ref.get(qn("w:val"))
            break
    if abstract_id is None:
        return ""
    lvl_text = "%1"
    for e in root.findall(qn("w:abstractNum")):
        if e.get(qn("w:abstractNumId")) != abstract_id:
            continue
        for lvl in e.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) != str(ilvl):
                continue
            lt = lvl.find(qn("w:lvlText"))
            if lt is not None and lt.get(qn("w:val")):
                lvl_text = lt.get(qn("w:val"))
        break
    n = 0
    for other in doc.element.body.iter(qn("w:p")):
        if _same_element(other, p):
            n += 1
            break
        if _numpr_of(other) == (num_id, ilvl):
            n += 1
    return _format_lvl_text(lvl_text, n)


def _same_element(a, b) -> bool:
    if a is None or b is None:
        return False
    if a is b:
        return True
    try:
        ta, tb = a.getroottree(), b.getroottree()
        return ta is tb and ta.getpath(a) == tb.getpath(b)
    except AttributeError:                                   # pragma: no cover
        return False


def seq_counters(doc) -> dict:
    """Count ``SEQ <label>`` occurrences per label, in document order."""
    counts: dict = {}
    for it in doc.element.body.iter(_INSTR):
        m = re.match(r"\s*SEQ\s+(\S+)", it.text or "")
        if m:
            label = m.group(1).strip('"')
            counts[label] = counts.get(label, 0) + 1
    return counts


# ---- field-introspection helpers -------------------------------------------

def seq_instr_in(p, label: str) -> list:
    """``w:instrText`` elements of the ``SEQ <label>`` fields inside *p*."""
    hits = []
    for it in p.iter(_INSTR):
        m = re.match(r"\s*SEQ\s+(\S+)", it.text or "")
        if m and m.group(1).strip('"') == label:
            hits.append(it)
    return hits


def seq_ordinal(doc, para, label: str) -> int:
    """1-based position of *para*'s own ``SEQ <label>`` among all of them."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        for it in p.iter(_INSTR):
            m = re.match(r"\s*SEQ\s+(\S+)", it.text or "")
            if not m or m.group(1).strip('"') != label:
                continue
            n += 1
            if _same_element(p, para._p):
                return n
    return n + 1


def _field_nodes(p) -> list:
    """``instrText`` / ``fldChar`` / ``t`` descendants of *p*, document order.

    ``p.iter()`` descends into ``w:hyperlink`` (and any other field wrapper),
    which is what makes renumber work on references written in the real
    hyperlink shape: their ``instrText`` is a *grandchild* of the paragraph,
    and a ``list(p)`` walk silently skipped every one of them.
    """
    return [e for e in p.iter() if e.tag in _FIELD_NODES]


def _field_span(p, instr_el):
    """``(begin_run, end_run)`` of the complex field owning *instr_el*."""
    nodes = [e for e in p.iter() if e.tag in (_FLDCHAR, _INSTR)]
    try:
        i = nodes.index(instr_el)
    except ValueError:                                       # pragma: no cover
        return None, None
    begin = end = None
    for j in range(i - 1, -1, -1):
        if nodes[j].tag == _FLDCHAR and nodes[j].get(qn("w:fldCharType")) == "begin":
            begin = nodes[j]
            break
    for j in range(i + 1, len(nodes)):
        if nodes[j].tag == _FLDCHAR and nodes[j].get(qn("w:fldCharType")) == "end":
            end = nodes[j]
            break
    return begin, end


def _enclosing_run(p, el):
    """The ``w:r`` ancestor of *el*."""
    for r in p.iter(qn("w:r")):
        if _same_element(r, el) or any(_same_element(c, el) for c in r.iter()):
            return r
    return None


def _advance_cached(nodes, i: int, new_text: str) -> bool:
    """Rewrite the cached result that follows node *i*; True when it changed."""
    for j in range(i + 1, len(nodes)):
        e = nodes[j]
        if e.tag == _T:
            if (e.text or "") != new_text:
                e.text = new_text
                return True
            return False
        if e.tag == _FLDCHAR and e.get(qn("w:fldCharType")) == "end":
            return False
    return False


def cached_result(doc, name: str, fmt: str) -> str:
    """The text a reader sees before Word recalculates the field.

    Always mirrors what Word itself computes for that flavour (oracle: see
    REF_FORMATS), so the preview and the recalculated result agree:

    * content-text flavours -> the number the bookmark declares, or the whole
      caption when the bookmark spans one;
    * paragraph-number flavours -> the item's list number with its format;
    * above/below and page number -> empty (Word and the locale decide).
    """
    kind = CACHED_KIND.get(fmt, "text")
    if kind == "text":
        return content_text_result(doc, name)
    if kind == "number":
        p = bookmark_paragraphs(doc).get(name)
        return list_number_for(doc, p) if p is not None else ""
    return ""


def _cached_for(doc, args, fmt: str) -> str:
    if getattr(args, "cached", None) is not None:
        return str(args.cached)
    return cached_result(doc, args.target, fmt)


def _vert_from_args(args) -> str | None:
    if getattr(args, "superscript", False):
        return "superscript"
    name = getattr(args, "vert_align", None) or "normal"
    if name not in VERT_ALIGN:
        raise SystemExit(f"unknown --vert-align: {name} "
                         f"(known: {sorted(VERT_ALIGN)})")
    return VERT_ALIGN[name]


# ---- subcommands ------------------------------------------------------------

def cmd_caption(doc, args) -> dict:
    """Number a paragraph as a caption, or renumber an existing one in place.

    Never throws content away: ``w:pPr`` and the original run properties are
    kept, existing text survives as the caption text when ``--text`` is not
    given, and a paragraph that already carries a ``SEQ`` for the same label is
    updated in place so hand-applied formatting (bold label, coloured text,
    hyperlinks) is not silently lost.
    """
    target = doc.paragraphs[args.index]
    p = target._p
    label = args.label
    rpr = _para_rpr(target)
    text = args.text
    drop_text = bool(getattr(args, "drop_text", False))

    existing = seq_instr_in(p, label)
    if existing:
        return _recaption(doc, target, args, label, existing[0], text)

    counters = seq_counters(doc)
    n = counters.get(label, 0) + 1
    name = args.bookmark or bm_name(label, n)
    if name in all_bookmarks(doc):
        # re-captioning the same paragraph must not duplicate a bookmark
        name = f"{name}_{n}"[:40]

    old_text = paragraph_text(p).strip()
    preserved = None
    if not text and old_text and not drop_text:
        text = old_text                      # keep the prose, reported below
        preserved = old_text

    span = getattr(args, "bookmark_span", None) or "number"
    if span not in ("number", "paragraph"):
        raise SystemExit(f"unknown --bookmark-span: {span} "
                         "(known: ['number', 'paragraph'])")

    want_ref = bool(getattr(args, "ref_bookmark", False))
    pstyle = getattr(args, "pstyle", None)

    target.clear()                           # keeps <w:pPr>, drops content
    if pstyle:
        # Written after clear(): clear() keeps the pPr element, but a style set
        # on a paragraph whose pPr is then rebuilt would be silently lost.
        set_paragraph_style(p, str(pstyle))
    bid = next_bm_id(doc.element.body)
    start, end = _bookmark_pair(name, bid)

    field = field_elements(f"SEQ {label} \\* ARABIC", str(n), rpr)
    if span == "paragraph":
        # Word's own caption bookmark shape: the whole caption is bookmarked,
        # so ``REF <name>`` returns the caption text, not just the number.
        p.append(start)
        p.append(_text_run(f"{label} ", rpr))
        for el in field:
            p.append(el)
        if text:
            p.append(_text_run(f"  {text}", rpr))
        p.append(end)
    else:
        p.append(_text_run(f"{label} ", rpr))
        p.append(start)
        for el in field:
            p.append(el)
        p.append(end)
        if text:
            p.append(_text_run(f"  {text}", rpr))
    ref_name = None
    if want_ref:
        # Outermost pair, so _Ref<number> covers label + SEQ + caption text and
        # both the table of figures and a plain REF see the whole caption.
        ref_name = next_ref_name(doc)
        _wrap_whole_paragraph(p, ref_name, next_bm_id(doc.element.body))
    return {"ok": True, "index": args.index, "number": n,
            "bookmark": name, "label": label, "recaption": False,
            "bookmark_span": span, "preserved_text": preserved,
            "ref_bookmark": ref_name, "pstyle": pstyle}


def _recaption(doc, target, args, label: str, instr_el, text: str) -> dict:
    """Refresh the number of a caption that already exists in this paragraph.

    Nothing is cleared, so every run the user formatted by hand (bold label,
    coloured text, a hyperlink inside the caption) survives untouched; only the
    cached ``SEQ`` result — and, when *text* is given, the trailing caption
    text — is rewritten.
    """
    p = target._p
    rpr = _para_rpr(target)
    n = seq_ordinal(doc, target, label)
    _advance_cached(_field_nodes(p), _field_nodes(p).index(instr_el), str(n))

    pstyle = getattr(args, "pstyle", None)
    if pstyle:
        set_paragraph_style(p, str(pstyle))

    # The caller's own bookmark, not a ``_Ref…`` one: taking whichever comes
    # first would report the table-of-figures anchor as if it were the caption's
    # name and then "reuse" it for a reference the caller never asked for.
    name = None
    ref_present = None
    for b in p.iter(qn("w:bookmarkStart")):
        bname = b.get(qn("w:name"))
        if not bname:
            continue
        if _REF_NUM.match(bname):
            ref_present = bname
        elif name is None:
            name = bname
    created = False
    if name is None:
        name = args.bookmark or bm_name(label, n)
        if name in all_bookmarks(doc):
            name = f"{name}_{n}"[:40]
        begin, end = _field_span(p, instr_el)
        bid = next_bm_id(doc.element.body)
        start, end_el = _bookmark_pair(name, bid)
        if begin is not None:
            _enclosing_run(p, begin).addprevious(start)
        else:                                                # pragma: no cover
            p.append(start)
        if end is not None:
            _enclosing_run(p, end).addnext(end_el)
        else:                                                # pragma: no cover
            p.append(end_el)
        created = True

    if text:
        # Replace only what follows the SEQ field; keep the bookmark end so the
        # caption stays referencable, and keep the field itself.
        _, end = _field_span(p, instr_el)
        anchor = _enclosing_run(p, end) if end is not None else None
        if anchor is not None:
            kids = list(p)
            at = kids.index(anchor)
            for child in kids[at + 1:]:
                if child.tag in (qn("w:bookmarkEnd"), _FLDCHAR):
                    continue
                p.remove(child)
            new_run = _text_run(f"  {text}", rpr)
            tail = _wrapping_bookmark_end(p)
            if tail is not None:
                # Keep the caption text inside the _Ref span: appending after
                # the closing bookmarkEnd would leave it outside the anchor.
                tail.addprevious(new_run)
            else:
                p.append(new_run)

    ref_name = ref_present
    if bool(getattr(args, "ref_bookmark", False)) and ref_present is None:
        ref_name = next_ref_name(doc)
        _wrap_whole_paragraph(p, ref_name, next_bm_id(doc.element.body))

    return {"ok": True, "index": args.index, "number": n,
            "bookmark": name, "label": label, "recaption": True,
            "bookmark_created": created, "ref_bookmark": ref_name,
            "pstyle": pstyle}


def _wrapping_bookmark_end(p):
    """The trailing ``w:bookmarkEnd`` of a whole-paragraph ``_Ref`` bookmark.

    Returns the element when the paragraph's last child is the end of a
    ``_Ref…`` bookmark, i.e. the paragraph is wrapped the way a table-of-figures
    anchor needs.  ``None`` otherwise, so callers fall back to plain appends.
    """
    kids = list(p)
    if not kids or kids[-1].tag != qn("w:bookmarkEnd"):
        return None
    bid = kids[-1].get(qn("w:id"))
    for b in p.iter(qn("w:bookmarkStart")):
        if b.get(qn("w:id")) == bid and _REF_NUM.match(b.get(qn("w:name")) or ""):
            return kids[-1]
    return None


def cmd_ref(doc, args) -> dict:
    """Append a cross-reference field to a paragraph.

    The field is wrapped in a real ``<w:hyperlink>`` so the number is clickable
    without Word ever recalculating anything, and every run of the field gets
    the requested ``<w:vertAlign>``.
    """
    names = all_bookmarks(doc)
    # A forward reference is the normal shape of a citation: the body says
    # "[1]" and the reference list that owns ``ref_1`` sits further down, so
    # the bookmark does not exist yet while the body is being written.
    # Bookmarks are document-wide and Word resolves them regardless of order,
    # so refusing to write one here would make the common case impossible.
    # ``allow_missing`` skips the check; the caller is expected to verify the
    # anchors once the whole document exists (``docx_create`` does).
    if args.target not in names and not getattr(args, "allow_missing", False):
        raise SystemExit(f"no such bookmark: {args.target} "
                         f"(known: {sorted(names)[:8]})")
    fmt = getattr(args, "format", None) or "default"
    field, switch = REF_FORMATS.get(fmt, ("REF", r" \h"))
    vert = _vert_from_args(args)
    use_hyperlink = not getattr(args, "no_hyperlink", False)

    target = doc.paragraphs[args.index]
    rpr = _para_rpr(target)
    cached = _cached_for(doc, args, fmt)
    merge = not getattr(args, "no_mergeformat", False)
    if merge:
        switch = switch + MERGE_FORMAT

    p = target._p
    if args.prefix:
        p.append(_text_run(args.prefix, rpr))
    els = field_elements(f"{field} {args.target}{switch}", str(cached), rpr, vert)
    if use_hyperlink:
        p.append(hyperlink_wrap(args.target, els))
    else:
        for el in els:
            p.append(el)
    if args.suffix:
        p.append(_text_run(args.suffix, rpr))
    return {"ok": True, "index": args.index, "target": args.target,
            "cached": str(cached), "format": fmt, "field": field,
            "switch": switch, "merge_format": merge,
            "vert_align": vert or "normal",
            "hyperlink": use_hyperlink, "anchor": args.target if use_hyperlink else None}


def check_crossref_anchors(doc) -> dict:
    """Verify that every ``REF``/``PAGEREF``/hyperlink anchor names a bookmark.

    A forward reference is normal while a document is being built, but a
    *dangling* one is a silent defect: the cached result is all the reader ever
    sees and the click goes nowhere. Call this once the whole document exists.

    Returns ``{"resolved": n, "targets": n, "missing": [...], "known": [...],
    "invalid_names": [...], "duplicate_names": [...]}``.

    ``invalid_names`` and ``duplicate_names`` are **hard failures**: both
    describe defects Word and LibreOffice repair silently on save, so the
    document looks fine in the XML yet opens with broken fields.  A caller
    should treat a non-empty value of either as a build error, exactly as it
    treats a non-empty ``missing`` — a duplicate name makes Word merge the two
    bookmarks (a table-of-contents entry then jumps to the wrong chapter) and
    makes LibreOffice rename the second one ``_副本_1`` (every ``PAGEREF``
    against the original name dies); a name with a space gets its space turned
    into ``_`` by both, and anything over 40 characters is truncated by both.
    See ``work/T3-REPORT.md`` §3.1-§3.3.

    ``invalid_names`` entries are ``{"name": ..., "problems": [...]}``, one per
    offending bookmark *occurrence* (so the same bad name appearing twice is
    reported twice), in document order.
    """
    body = doc.element.body
    names = all_bookmarks(doc)
    targets: list[str] = []
    for it in body.iter(qn("w:instrText")):
        m = re.match(r"\s*(?:REF|PAGEREF)\s+([^\s\\]+)", it.text or "")
        if m:
            targets.append(m.group(1))
    for h in body.iter(qn("w:hyperlink")):
        anchor = h.get(qn("w:anchor"))
        if anchor:
            targets.append(anchor)

    invalid_names: list[dict] = []
    for b in body.iter(qn("w:bookmarkStart")):
        name = b.get(qn("w:name"))
        if not name:
            continue
        problems = validate_bookmark_name(name)
        if problems:
            invalid_names.append({"name": name, "problems": problems})

    missing = sorted({t for t in targets if t not in names})
    return {"resolved": sum(1 for t in targets if t in names),
            "targets": len(targets), "missing": missing,
            "known": sorted(names),
            "invalid_names": invalid_names,
            "duplicate_names": validate_bookmark_uniqueness(doc)["duplicates"]}


def cmd_bookmark(doc, args) -> dict:
    target = doc.paragraphs[args.index]
    names = all_bookmarks(doc)
    name = args.name
    if name in names:
        raise SystemExit(f"bookmark already exists: {name}")
    bid = next_bm_id(doc.element.body)
    start, end = _bookmark_pair(name, bid)
    p = target._p
    ppr = p.pPr
    if ppr is not None:
        ppr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)
    return {"ok": True, "index": args.index, "bookmark": name}


def cmd_renumber(doc, args) -> dict:
    """Recompute SEQ counters and refresh every cached REF result.

    Word does this on F9; doing it at write time means the file is correct the
    moment it is opened, for readers that never update fields.

    The walk is over ``p.iter()`` (recursive) rather than the paragraph's
    direct children: a reference written with the clickable-hyperlink shape has
    its ``instrText`` one level deeper, and a flat walk misses all of them.

    Two passes.  SEQ first, so a cached reference can copy the *final* number.
    Every rewrite mirrors what Word computes for that flavour (oracle):
    content-text fields get the bookmarked text, paragraph-number fields get
    the item's list number, and the flavours Word resolves from layout
    (``\\p`` above/below, PAGEREF page numbers) are left alone rather than
    filled with a wrong-looking guess.
    """
    body = doc.element.body
    bm_paras = bookmark_paragraphs(doc)
    counters: dict = {}
    seq_fixed = 0
    ref_fixed = 0
    ref_skipped = 0

    for p in body.iter(qn("w:p")):
        nodes = _field_nodes(p)
        for i, e in enumerate(nodes):
            if e.tag != _INSTR:
                continue
            m = re.match(r"SEQ\s+(\S+)", (e.text or "").strip())
            if not m:
                continue
            label = m.group(1).strip('"')
            counters[label] = counters.get(label, 0) + 1
            if _advance_cached(nodes, i, str(counters[label])):
                seq_fixed += 1

    for p in body.iter(qn("w:p")):
        nodes = _field_nodes(p)
        for i, e in enumerate(nodes):
            if e.tag != _INSTR:
                continue
            instr = (e.text or "").strip()
            m = re.match(r"(REF|PAGEREF)\s+(\S+)", instr)
            if not m:
                continue
            field, name = m.group(1), m.group(2)
            if field == "PAGEREF" or re.search(r"\\p(\s|$)", instr):
                ref_skipped += 1
                continue
            if re.search(r"\\[nrw](\s|$)", instr):
                bm = bm_paras.get(name)
                new_text = list_number_for(doc, bm) if bm is not None else ""
            else:
                new_text = content_text_result(doc, name)
            if not new_text:
                ref_skipped += 1
                continue
            if _advance_cached(nodes, i, new_text):
                ref_fixed += 1

    return {"ok": True, "seq_updated": seq_fixed, "ref_updated": ref_fixed,
            "ref_skipped": ref_skipped, "counters": counters}


def cmd_list(doc) -> dict:
    body = doc.element.body
    seqs = []
    refs = []
    for it in body.iter(_INSTR):
        txt = (it.text or "").strip()
        if re.match(r"SEQ\s+", txt):
            seqs.append(txt)
        elif re.match(r"(REF|PAGEREF)\s+", txt):
            refs.append(txt)
    anchors = [h.get(qn("w:anchor")) for h in body.iter(qn("w:hyperlink"))
               if h.get(qn("w:anchor"))]
    return {"bookmarks": sorted(all_bookmarks(doc)),
            "seq_fields": seqs, "ref_fields": refs,
            "counters": seq_counters(doc),
            "hyperlinks": len(anchors), "hyperlink_anchors": anchors}


# ---- CLI --------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Cross-reference and caption numbering for .docx.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("path", help="input .docx")
        p.add_argument("-o", "--output", help="output path (default: overwrite)")

    p = sub.add_parser("caption", help="turn a paragraph into a numbered caption")
    common(p)
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--label", default="图",
                   help="caption label; each label has its own counter")
    p.add_argument("--text", default="", help="caption text after the number")
    p.add_argument("--bookmark", help="bookmark name (default <label>_<n>)")
    p.add_argument("--bookmark-span", default="number",
                   choices=["number", "paragraph"],
                   help="number: bookmark the SEQ counter only, so REF returns "
                        "'1'; paragraph: bookmark the whole caption, Word's own "
                        "_Ref… shape, so REF returns the full caption text")
    p.add_argument("--drop-text", action="store_true",
                   help="discard the paragraph's existing text instead of "
                        "keeping it as the caption text")
    p.add_argument("--ref-bookmark", action="store_true",
                   help="also bracket the whole caption with a _Ref<number> "
                        "bookmark, the anchor family Word's 'Insert Table of "
                        "Figures' builds its PAGEREF links from; the "
                        "--bookmark name is kept alongside it")
    p.add_argument("--pstyle",
                   help="w:styleId to hang on the caption paragraph "
                        "(e.g. the caption style of the target template); "
                        "written as the first child of w:pPr, which the "
                        "schema requires")

    p = sub.add_parser("ref", help="insert a REF cross-reference")
    common(p)
    p.add_argument("--index", type=int, required=True,
                   help="paragraph to append the reference to")
    p.add_argument("--target", required=True, help="bookmark name")
    p.add_argument("--prefix", default="", help="text before the number")
    p.add_argument("--suffix", default="", help="text after the number")
    p.add_argument("--format", default="default", choices=sorted(REF_FORMATS),
                   help=r"number | full (whole caption, needs "
                        r"--bookmark-span paragraph) | above-below (\p) | "
                        r"page (PAGEREF) | paragraph-number (\n, the [1] the "
                        r"reference list shows) | relative-context (\r) | "
                        r"full-context (\w) | default (\h hyperlink)")
    p.add_argument("--cached", help="result shown before fields update "
                   "(default mirrors what Word computes for --format)")
    p.add_argument("--vert-align", default="normal", choices=sorted(VERT_ALIGN),
                   help="superscript/subscript raise or lower the reference; "
                        "normal keeps it at baseline")
    p.add_argument("--superscript", action="store_true",
                   help="shorthand for --vert-align superscript")
    p.add_argument("--no-hyperlink", action="store_true",
                   help="write a bare field instead of wrapping it in "
                        "<w:hyperlink w:anchor=...>")
    p.add_argument("--no-mergeformat", action="store_true",
                   help=r"omit the \* MERGEFORMAT switch; without it Word "
                        r"drops the superscript from the result on the next "
                        r"field update")
    p.add_argument("--allow-missing", action="store_true",
                   help="write the field even though the target bookmark does "
                        "not exist yet (forward reference to a reference list "
                        "that comes later); verify the anchors afterwards")

    p = sub.add_parser("bookmark", help="put a bookmark on a paragraph")
    common(p)
    p.add_argument("--index", type=int, required=True)
    p.add_argument("--name", required=True)

    p = sub.add_parser("renumber",
                       help="recompute SEQ counters and cached REF results")
    common(p)

    p = sub.add_parser("list", help="show bookmarks, SEQ and REF fields")
    common(p)

    args = ap.parse_args()
    doc = Document(args.path)
    result = {
        "caption": lambda: cmd_caption(doc, args),
        "ref": lambda: cmd_ref(doc, args),
        "bookmark": lambda: cmd_bookmark(doc, args),
        "renumber": lambda: cmd_renumber(doc, args),
        "list": lambda: cmd_list(doc),
    }[args.cmd]()

    if args.cmd != "list":
        out = args.output or args.path
        doc.save(out)
        result["output"] = out
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
