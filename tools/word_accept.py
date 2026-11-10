#!/usr/bin/env python3
# MIT License. Part of the docx skill (patched): unified acceptance probe.
"""Unified acceptance probe for .docx — XML layer + Word COM layer.

Why this file exists
--------------------
``docx_validate`` returning ``ok: true`` does not mean Word can open the file,
and ``python-docx`` reading a document does not mean Word renders its numbering,
its fields or its cross-references.  The only trustworthy acceptance evidence is
Word itself.  Every agent in the project therefore calls *this* file instead of
writing its own COM snippet, so the evidence has one shape and one meaning.

Two layers, on purpose
----------------------
``probe_xml(docx)``
    Pure stdlib (``zipfile`` + ``xml.etree``).  No Office, no ``pywin32``, no
    ``lxml``, nothing from ``scripts/``.  It runs on CI and it is the *contract*
    layer: its keys and shapes never change, and it is what the caller still
    gets when Word is unavailable.  Everything that can be decided from the
    bytes — structure, bookmarks, field codes, hyperlink wrappers, ``numPr``,
    ``lvlText`` — is decided here.

``probe_word(docx)``
    Word COM via ``pywin32``.  Adds only what Word knows: page count, section
    count, page geometry, rendered list strings, hyperlink *targets*, field
    *results*, PDF export and per-page PNG.

``main()`` merges the two into one JSON object on stdout.  The top-level
contract keys are **always present and populated**: when Word did not run they
fall back to the XML value, and ``sources`` records which layer each number
came from.  A consumer never has to guard for a missing key.

Page numbers: displayed vs physical vs computed
-----------------------------------------------
Two numbers exist for every page and they are not interchangeable:

* ``WD_ACTIVE_END_PAGE`` (3) — the **physical** sheet ordinal, 1-based from the
  start of the document, ignoring every ``w:pgNumType`` restart and format.
* ``WD_ACTIVE_END_ADJUSTED_PAGE`` (1) — the number **printed** on the page, i.e.
  what ``w:pgNumType`` adjusted the counter to.  Turn it into a glyph with
  :func:`page_glyph`, which delegates to ``scripts/docx_section.cached_page``
  so the counter alphabets live in exactly one place.

Measured on ``work/subjects/math/paper.docx`` (5 sections, the last one
``lowerRoman start=1``): physical page 7 carries the printed number 2, physical
page 8 carries ``i``, and the exported PDF's footer band reads ``- i -`` on that
page.  Reporting "page 7" and "i" as the same thing is wrong in both directions,
so every page-number statement carries its source.

``computed`` is a third, weaker source and the rule around it is absolute:
**computed 不得当作交付页码**.  :func:`estimated_paragraph_pages` fills
``paragraph_pages`` when Word is absent or was not asked, but those rows come
from ``w:pgNumType`` plus the section's paragraph span alone — they know which
section a paragraph belongs to, never where it falls on a page, and they inherit
the XML index space, which drifts from Word's by one paragraph per table row.
Every such row is stamped ``exact: false`` / ``source: "computed"`` and the
report repeats it in ``warnings``.

Contract
--------
::

    python tools/word_accept.py <in.docx> [--pdf OUT.pdf] [--png DIR]
        [--json OUT.json] [--expect-sections N] [--expect-bookmarks N]
        [--expect-params N] [--page-probe I[,J,...]]

    --expect-params N        asserted ``paragraphs`` (alias of
                             ``--expect-paragraphs``)
    --word auto|on|off       default ``auto`` = use Word when it is there
    --page-probe 1,9,43      Word paragraph indices (1-based) whose page numbers
                             to read; two COM round trips each, hence opt-in

Exit codes: 0 success or skipped, 1 ``ok: false``, 2 usage error.
Everything a human should read (summary, errors) goes to stderr, so stdout is
always exactly one JSON document.

Hard-won facts encoded here
---------------------------
1. Numbering must be read from ``Paragraphs(i).Range.ListFormat.ListString``;
   ``Range().Text`` does not contain the generated number.  To avoid one COM
   round trip per paragraph, the Word layer uses the native
   ``Document.ListParagraphs`` collection, whose order matches the document
   order of the ``w:numPr`` paragraphs found by ``probe_xml``.  The two lists are
   zipped by *ordinal*, never by paragraph index.
2. ``XML paragraph index != Word paragraph index``: Word adds paragraph marks
   for tables beyond the ones the XML holds.  Measured as ``columns + 1`` per
   row (one mark per cell plus the row-end mark), **not** a flat ``+1``:
   1x1 table ``+2``, 2x2 ``+6``, 3x4 ``+15``
   (work/REVIEW-ADVERSARIAL-R6.md §7.1).  On ``f5.docx``: 261 ``w:p`` in XML
   against 287 ``Word.Paragraphs`` with 26 ``w:tr``.  So never translate an XML
   index into ``Paragraphs(i)`` by arithmetic — least of all by row count; carry
   both and label them.
3. ``\\h`` on a REF field does **not** create a ``w:hyperlink`` element.  A real
   wrapper — ``<w:hyperlink w:anchor="..." w:history="1">`` around the whole
   field — is what makes ``Document.Hyperlinks.Count >= 1`` and
   ``Hyperlinks(i).SubAddress`` non-empty.  Those two numbers are the
   programmatic "it is clickable" signal, and ``probe_xml`` reports the same
   wrapper count so CI can assert it without Word.
4. ``gencache.EnsureDispatch("Word.Application")`` breaks on the second call in
   one process (``RPC server unavailable`` then ``can not automate the makepy
   process``).  Use plain ``win32.Dispatch`` and build the app at most once per
   process; ``release_app()`` quits it once, at the end.
"""
from __future__ import annotations

import argparse
import atexit
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

TOOL = "word_accept"
VERSION = "1.0"

# --------------------------------------------------------------- namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_WP = ("http://schemas.openxmlformats.org/drawingml/2006/"
         "wordprocessingDrawing")
NS_M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

DOCUMENT_PART = "word/document.xml"
NUMBERING_PART = "word/numbering.xml"

TEXT_CAP = 80          # characters kept from any paragraph/field text
DEFAULT_LIMIT = 200    # max entries in any list-valued output key
# A per-character COM walk is slow (one round trip per character), so the
# script audit samples at most this many characters from the top of the body.
# Every document this tool has been pointed at has its script in the first
# pages; the cap is reported alongside the counts so a truncated scan is never
# mistaken for "there is no superscript here".
SCRIPT_SCAN_CAP = 20000
# Candidate runs handed to the Word oracle.  One COM round trip each, so this
# is a budget rather than a limit: a document with more claims than this is
# reported as truncated instead of silently under-counted.
SCRIPT_PROBE_BUDGET = 400

# Word constants
WD_STAT_PAGES = 2
# ``wdActiveEndAdjustedPageNumber`` — the number *printed* on the page, i.e. the
# value the page counter has been adjusted to by ``w:pgNumType``.  A page shown
# as ``iii`` reports 3 here, not 3-as-in-third-sheet.
WD_ACTIVE_END_ADJUSTED_PAGE = 1
# ``wdActiveEndPageNumber`` — the physical sheet ordinal, 1-based from the start
# of the document, ignoring every ``w:pgNumType`` restart and format.
WD_ACTIVE_END_PAGE = 3
WD_EXPORT_FORMAT_PDF = 17
# Footer index: 1 primary, 2 first page, 3 even pages
WD_FOOTER_INDEXES = (1, 2, 3)


def q(local: str) -> str:
    """WordprocessingML tag."""
    return "{%s}%s" % (NS_W, local)


def qa(local: str) -> str:
    """DrawingML tag."""
    return "{%s}%s" % (NS_A, local)


def qwp(local: str) -> str:
    """wordprocessingDrawing tag."""
    return "{%s}%s" % (NS_WP, local)


def qm(local: str) -> str:
    """Office Math (OMML) tag — ``m:sSup``, ``m:sSub``, ``m:oMath`` and friends."""
    return "{%s}%s" % (NS_M, local)


def _fix_stdio() -> None:
    """Reconfigure stdout/stderr to UTF-8 on Windows.

    A standalone copy of the stdio fix from ``scripts/_bootstrap.py``: this tool
    must keep
    working when ``scripts/`` is mid-edit, so it imports nothing from there.
    """
    if __import__("os").name != "nt":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError, ValueError):
            pass


_fix_stdio()


# ===========================================================================
# XML layer
# ===========================================================================
class DocxProbeError(Exception):
    """The bytes are not usable as a .docx document part."""


def _cap(text, n: int = TEXT_CAP) -> str:
    text = "" if text is None else str(text)
    text = " ".join(text.split())
    return text[:n]


def _para_text(p) -> str:
    """Visible text of a paragraph, with tabs and breaks preserved.

    ``<w:tab/>`` is an *element*, not a character, so a naive concatenation of
    ``w:t`` texts silently drops it.  That matters: a table-of-contents entry is
    ``"摘要" + <w:tab/> + "i"``, and without the tab the trailing page number
    cannot be told apart from the entry's own last word.
    """
    parts = []
    for el in p.iter():
        tag = el.tag
        if tag == q("t"):
            parts.append(el.text or "")
        elif tag == q("tab"):
            parts.append("\t")
        elif tag in (q("br"), q("cr")):
            parts.append("\n")
    return "".join(parts)


def _read_part(zf: zipfile.ZipFile, name: str):
    try:
        return zf.read(name)
    except KeyError as exc:
        raise DocxProbeError("part missing: %s" % name) from exc


def _style_names(zf: zipfile.ZipFile):
    """``{styleId: style name}`` from ``word/styles.xml``.

    Necessary because ``<w:pStyle w:val="...">`` holds the *styleId*, not the
    name.  In the 示例竞赛 source the TOC paragraphs carry ``val="9"``/``"12"``/
    ``"6"``, whose names are ``toc 1``/``toc 2``/``toc 3``; the Chinese body
    styles are ``章标题``/``节标题``/``条标题``/``正文`` behind numeric
    ids.  Anything that pattern-matches on style *names* without resolving
    first sees numbers and matches nothing — the same class of bug as the
    upstream 6.5 % heading recall.
    """
    out = {}
    if "word/styles.xml" not in zf.namelist():
        return out
    try:
        root = ET.fromstring(zf.read("word/styles.xml"))
    except ET.ParseError:
        return out
    for st in root.findall(q("style")):
        sid = st.get(q("styleId"))
        if not sid:
            continue
        nm = st.find(q("name"))
        out[sid] = (nm.get(q("val")) if nm is not None else None) or sid
    return out


def _para_style(p, style_names):
    """``(styleId, resolved name)`` for a paragraph; ``(None, None)`` if none."""
    ppr = p.find(q("pPr"))
    if ppr is None:
        return None, None
    st = ppr.find(q("pStyle"))
    if st is None:
        return None, None
    sid = st.get(q("val"))
    return sid, style_names.get(sid, sid)


def _numbering_formats(zf: zipfile.ZipFile):
    """``{numId: {ilvl: lvlText}}`` plus the abstract-level text map.

    ``numId`` is what a ``<w:numPr>`` names; ``abstractNum`` holds the actual
    ``lvlText``.  Both must be joined through ``<w:num><w:abstractNumId>``.
    """
    num_to_abs = {}
    abs_lvl = {}
    if NUMBERING_PART not in zf.namelist():
        return num_to_abs, abs_lvl
    try:
        root = ET.fromstring(zf.read(NUMBERING_PART))
    except ET.ParseError:
        return num_to_abs, abs_lvl
    for an in root.findall(q("abstractNum")):
        aid = an.get(q("abstractNumId"))
        lvls = {}
        for lvl in an.findall(q("lvl")):
            ilvl = lvl.get(q("ilvl"), "0")
            txt = lvl.find(q("lvlText"))
            fmt = lvl.find(q("numFmt"))
            lvls[ilvl] = {
                "lvlText": txt.get(q("val")) if txt is not None else None,
                "numFmt": fmt.get(q("val")) if fmt is not None else None,
                "start": (lvl.find(q("start")).get(q("val"))
                          if lvl.find(q("start")) is not None else None),
            }
        abs_lvl[aid] = lvls
    for num in root.findall(q("num")):
        nid = num.get(q("numId"))
        aid_el = num.find(q("abstractNumId"))
        if nid is not None and aid_el is not None:
            num_to_abs[nid] = aid_el.get(q("val"))
    return num_to_abs, abs_lvl


def _field_entries(root, limit: int):
    """All complex + simple fields in document order.

    Walks *every* descendant, so a field wrapped in ``<w:hyperlink>`` is still
    found — that is exactly the case §7.2 of the project brief introduces.
    Returns ``(count, [{"code", "result"}])``.
    """
    entries = []
    stack = []   # open complex fields
    count = 0
    for el in root.iter():
        tag = el.tag
        if tag == q("fldChar"):
            kind = el.get(q("fldCharType"))
            if kind == "begin":
                count += 1
                stack.append({"code": [], "result": [], "sep": False})
            elif kind == "separate":
                if stack:
                    stack[-1]["sep"] = True
            elif kind == "end":
                if stack:
                    f = stack.pop()
                    entries.append({"code": _cap("".join(f["code"]).strip()),
                                    "result": _cap("".join(f["result"]))})
        elif tag == q("fldSimple"):
            count += 1
            entries.append({
                "code": _cap((el.get(q("instr")) or "").strip()),
                "result": _cap("".join(t.text or "" for t in el.iter(q("t")))),
            })
        elif tag == q("instrText"):
            if stack:
                stack[-1]["code"].append(el.text or "")
        elif tag == q("t") and stack:
            bucket = stack[-1]["result"] if stack[-1]["sep"] else stack[-1]["code"]
            bucket.append(el.text or "")
    while stack:   # unterminated field: still report what was seen
        f = stack.pop()
        entries.append({"code": _cap("".join(f["code"]).strip()),
                        "result": _cap("".join(f["result"]))})
    return count, entries[:limit]


# The trailing page number of a TOC line.  Three families, because a document
# may number its sections with any of them:
#
#   digits        ``15``
#   Roman         ``iv``, ``xii``   (a subset of the letters below)
#   letters       ``a``, ``b``, ``aa``  — an appendix numbered ``lowerLetter``
#
# The letter branch has to exist: without it every entry of a Latin-numbered
# appendix was reported as "no page number at all", i.e. a correct TOC was
# labelled broken 8 times over and its tokens were dropped from the sequence
# (work/REVIEW-ADVERSARIAL-R6.md 阻塞-2).  A single letter is *also* a valid
# Roman numeral (``i``, ``v``, ``x``, ``l``, ``c``, ``d``, ``m``), so the family
# is decided later, by ``_token_number`` with the section's own ``fmt`` and by
# ``_family`` for the run splitting — not here, where nothing else is known.
#
# The separator matters as much as the value.  A TOC line is *title + tab +
# number*; when the number has not been filled in yet the line ends with the
# tab, and a pattern that also accepted a plain space would read the last word
# of the title as the page number -- ``节 0<TAB>`` reported page ``0``, and
# ``第一章 概述<TAB>`` reported nothing but ``第 5 章 概述`` would report ``5``.
# So the tab form is tried first and a space/dot form is only a fallback for
# producers that write no tab at all.
_TOC_NUM_AFTER_TAB_RE = re.compile(r"(?:[0-9]{1,4}|[A-Za-z]{1,6})")
_TOC_NUM_LOOSE_RE = re.compile(r"[. ]([0-9]{1,4}|[A-Za-z]{1,6})\s*$")
_TOC_STYLE_RE = re.compile(r"^(toc|目录|contents)", re.IGNORECASE)
# A caption paragraph style, in the many names it is spelled with: Word's
# built-in "caption", the demo template's "图表名", and the Chinese "题注".
_CAPTION_STYLE_RE = re.compile(r"(caption|图表名|题注|图名|表名)", re.IGNORECASE)


def _toc_token(text: str):
    """Trailing page number of a table-of-contents entry, or ``None``.

    A tab is the entry/number separator in every producer seen so far, and its
    presence is authoritative: if the line has a tab, the page number is what
    follows the **last** tab and nothing before it counts.  Only a line with no
    tab at all falls back to the looser space/dot reading, so a title that ends
    in a digit cannot masquerade as a page number.

    The tab test is made on the *raw* text: stripping first would remove a
    trailing tab -- exactly the marker that says "the number has not been filled
    in yet" -- and hand a title like ``附录 A<TAB>`` to the loose pattern, which
    then reported ``A`` as the page number.
    """
    raw = (text or "").replace("\u00a0", " ")
    if not raw.strip():
        return None
    if "\t" in raw:
        tail = raw.rsplit("\t", 1)[1].strip()
        if tail and _TOC_NUM_AFTER_TAB_RE.fullmatch(tail):
            return tail
        return None
    m = _TOC_NUM_LOOSE_RE.search(raw.rstrip())
    return m.group(1) if m else None


# Word's built-in TOC paragraph styles are named ``toc 1``, ``toc 2``, … (and
# ``TOC1``/``Toc 2`` in documents written by other producers).  Only these carry
# a *level*; ``table of figures`` and the Chinese ``目录`` style are TOC-ish but
# not levelled, so they are audited separately (``_TOC_STYLE_RE``).
_TOC_LEVEL_RE = re.compile(r"toc\s*([0-9])", re.IGNORECASE)

_ROMAN_VALUE = {"i": 1, "v": 5, "x": 10, "l": 50, "c": 100, "d": 500, "m": 1000}

# Paragraph-page probing is opt-in because it costs COM round trips: a probe
# that asks for 300 paragraphs is 900 RPCs, which is how the automation channel
# starts answering ``被呼叫方拒绝接收呼叫``.  Callers pass the handful of
# paragraphs they actually hand to a reader (TOC targets, section heads).
PAGE_PROBE_BUDGET = 60

# ``scripts/docx_section`` is imported lazily and never fatally: the tool must
# keep working when that module is mid-edit, missing, or fails to import (its
# own dependencies are python-docx / lxml, neither of which the XML layer
# otherwise needs).
_SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
_SECTION_MOD = {"mod": None, "tried": False}


def _section_module():
    """``scripts/docx_section`` as a module, or ``None``.  Imported once."""
    if _SECTION_MOD["tried"]:
        return _SECTION_MOD["mod"]
    _SECTION_MOD["tried"] = True
    mod = None
    try:
        here = str(_SCRIPTS_DIR)
        if here not in sys.path:
            sys.path.insert(0, here)
        import docx_section as mod      # local import by design; see above
    except Exception:
        mod = None
    _SECTION_MOD["mod"] = mod
    return mod


def page_glyph(fmt, n):
    """The glyph a reader sees for page number ``n`` under ``w:pgNumType/@fmt``.

    ``3`` + ``lowerRoman`` -> ``"iii"``; ``None``/``decimal`` -> ``"3"``.  The
    counter alphabets live in exactly one place — ``docx_section.cached_page`` —
    so this delegates instead of growing a second copy; if that module cannot be
    imported the number is passed through as text, which is lossy but honest.
    """
    try:
        num = int(n)
    except (TypeError, ValueError):
        return "" if n is None else str(n)
    mod = _section_module()
    if mod is not None:
        try:
            return str(mod.cached_page(fmt, num))
        except Exception:
            pass
    return str(num)


def _token_number(tok, fmt=None):
    """``"iii"`` -> 3, ``"12"`` -> 12, anything unparseable -> ``None``.

    ``fmt`` is the enclosing section's ``w:pgNumType/@fmt``.  It matters: the
    same token means different numbers under different counters.  ``"c"`` is 3
    under ``lowerLetter`` (a, b, c) and 100 under ``lowerRoman`` (c = 100);
    ``"i"`` is 9 under ``lowerLetter`` and 1 under ``lowerRoman``.  Guessing
    Roman for every alphabetic token silently misnumbers a Latin appendix and
    makes an increasing TOC look like it goes backwards -- see
    work/REVIEW-ADVERSARIAL-R6.md 阻塞-2.

    Only the counter families whose alphabet can collide are disambiguated;
    anything else falls back to the old digit-then-Roman reading.
    """
    if tok is None:
        return None
    t = str(tok).strip()
    if not t:
        return None
    if t.isdigit():
        return int(t)

    low = t.lower()

    # A section that says "letters" means letters, full stop.  ``lowerLetter``
    # continues a, b, ..., z, aa, ab -- so a multi-character token is only
    # valid there when every character is a Latin letter, which is also true of
    # ``iii``; the caller's fmt is the tie-breaker and is trusted.
    if fmt in ("lowerLetter", "upperLetter"):
        if low.isalpha() and low.isascii():
            total = 0
            for ch in low:
                total = total * 26 + (ord(ch) - ord("a") + 1)
            return total
        return None

    if fmt in ("decimalFullWidth", "decimalEnclosedCircle", "decimalEnclosedCircleChinese",
               "ideographDigital", "chineseCounting", "chineseCountingThousand",
               "chineseLegalSimplified", "japaneseCounting", "taiwaneseCounting"):
        # These all render a digit string in another glyph set; the trailing
        # token was matched as digits already when the document is well formed.
        # A non-digit token means the reader sees something we cannot invert.
        return None

    if low == "0":
        return 0
    if not all(ch in _ROMAN_VALUE for ch in low):
        return None
    total = 0
    prev = 0
    for ch in reversed(low):
        val = _ROMAN_VALUE[ch]
        if val < prev:
            total -= val
        else:
            total += val
            prev = val
    return total


def _toc_audit(paras, style_names, limit: int = DEFAULT_LIMIT,
               section_of=None, section_specs=None) -> dict:
    """What the table of contents looks like from the bytes alone.

    Two sets of paragraphs matter, and they are counted separately because they
    are not the same set:

    * **levelled entries** — resolved style name matches ``^toc\\s*\\d``
      (``toc 1``, ``toc 2``, ``TOC1``).  ``entries`` / ``styles`` / ``levels``
      describe these.  Resolution is required: the XML stores
      ``w:pStyle w:val="9"``, so matching the raw value finds nothing.
    * **any TOC-styled paragraph** — resolved name matches :data:`_TOC_STYLE_RE`
      (``toc…``, ``目录…``, ``contents…``, and hand-rolled names such as
      ``TocLine`` which carry no level).  ``page_tokens`` /
      ``page_parse_failures`` / ``monotonic`` are computed over these, i.e. the
      *same* set ``page_numbers`` is built from, so the two blocks can never
      disagree about which paragraphs the document calls a TOC.

    ``monotonic`` answers the question a reader asks first: do the numbers ever
    go backwards?  Three things make a naive reading wrong, all measured:

    * the token has to be read under the **counter of its own section**
      (``_token_number``'s ``fmt``); ``a, b, c`` in a Latin appendix is 1, 2, 3,
      not 1, 2, 100;
    * a document legitimately **restarts** its counter when the format changes
      (front matter ``i, ii, iii`` then body ``1, 2, ...``).  Comparing across
      such a boundary reports a correct thesis TOC as broken, so the comparison
      is made **within one section** only.  ``monotonic_across_sections`` keeps
      the blunt whole-document answer visible for callers that want it.
    * **case is part of the counter family**: ``A, B`` is ``upperLetter`` and
      ``a, b`` is ``lowerLetter``, so a case-only change splits runs the same
      way an alphabet change does.  The step across it is still *observed* in
      ``monotonic`` (``A, B, a, b`` really does go 2 -> 1) but the verdict is
      downgraded to ``unknown``: from the bytes, a restarted counter in the
      other case and a mis-cased token look the same.

    ``page_parse_failures`` lists the entries whose trailing page number could
    not be read at all — an entry with no number is a broken TOC even when it
    renders.  A **TOC title line** (style ``TOC Heading``, text ``目录``) is not
    an entry and is excluded from that count; otherwise every titled TOC would
    report itself broken.

    ``source`` is ``"toc"`` when the document really has TOC-styled paragraphs,
    ``"computed"`` when it has none — that is a structural statement ("no table
    of contents here"), not a measurement.
    """
    entries = 0
    styles = []
    levels = set()
    broad = 0
    tokens = []
    failures = []
    # ``[(token, section_index), ...]`` in document order.  The pair is kept
    # together because the counter that decides how to read a token is the one
    # its own section declares -- see ``_token_number``.
    token_sections = []
    for idx, p in enumerate(paras):
        _sid, name = _para_style(p, style_names)
        if not name:
            continue
        flat = name.replace("\u3000", " ")
        m = _TOC_LEVEL_RE.match(flat)
        if m:
            entries += 1
            if name not in styles:
                styles.append(name)
            levels.add(int(m.group(1)))
        if not _TOC_STYLE_RE.match(name):
            continue
        # ``TOC Heading`` is the *title* of the list, not a line of it: it has no
        # level number and never carries a page number.  Counting it would make
        # ``page_parse_failures`` permanently non-empty for any titled TOC.
        if m is None:
            continue
        broad += 1
        text = _para_text(p)
        tok = _toc_token(text)
        if tok is None:
            failures.append({"index": idx, "style": name, "text": _cap(text)})
        else:
            tokens.append(tok)
            sec = None
            if section_of is not None:
                try:
                    sec = section_of(idx)
                except Exception:
                    sec = None
            token_sections.append((tok, sec))

    def _fmt_of(section_index):
        """The counter family in force for a section, or ``None``."""
        if not section_specs or section_index is None:
            return None
        for spec in section_specs:
            if spec.get("index") == section_index:
                return spec.get("fmt")
        return None

    nums = [_token_number(tok, _fmt_of(sec)) for tok, sec in token_sections]
    parseable = [n for n in nums if n is not None]
    whole_document_monotonic = all(
        b >= a for a, b in zip(parseable, parseable[1:]))

    by_section = {}
    for tok, sec in token_sections:
        by_section.setdefault(sec, []).append(tok)

    # Monotonicity has to answer "do the numbers go backwards *within one
    # numbering run*", and that run is not the section a TOC line physically
    # sits in -- every entry of a table of contents lives in the TOC's own
    # section, at the front of the document.  What varies is the counter the
    # *target* heading uses.  So the sequence is split wherever the glyph
    # family changes (``ii, iii`` then ``1, 15``), and each run is checked on
    # its own: a document that restarts its counter for the body is correct,
    # not broken (work/REVIEW-ADVERSARIAL-R6.md 阻塞-2).
    #
    # Every family that is compared MUST be handed the counter it is read
    # under.  A ``letter`` token parsed with ``fmt=None`` falls through to the
    # Roman reading, so plain letters other than i/v/x/l/c/d/m come back as
    # ``None``, the run stays empty, and ``all()`` over an empty sequence is
    # ``True`` -- which reported a descending appendix (``c, a, b``) as
    # monotonic.  Silently passing a real regression is worse than the noisy
    # false alarm this split was introduced to remove, so the mapping below is
    # exhaustive and the fallback is a *refusal to judge*, not a green light
    # (work/REVIEW-ADVERSARIAL-R6.md 第三轮 阻塞-④).
    #
    # Case is part of the family: Word's letter and Roman counters each come
    # in a lower and an UPPER case (``lowerLetter``/``upperLetter``,
    # ``lowerRoman``/``upperRoman``), and a case change between two counters
    # is a legitimate restart, so ``upperLetter -> lowerLetter`` must split
    # runs just like ``roman -> decimal`` does (work/P3-REPORT-B.md §1.5,
    # adopted for R7-5a).  The bare ``letter``/``roman`` keys stay because a
    # run of single-letter tokens is *inferred* to be Latin in pass 2 below.
    _FAMILY_FMT = {"roman": "lowerRoman", "roman_upper": "upperRoman",
                   "roman_lower": "lowerRoman",
                   "letter": "lowerLetter", "letter_upper": "upperLetter",
                   "letter_lower": "lowerLetter", "decimal": "decimal"}

    def _family(tok):
        t = str(tok).strip()
        if not t:
            return None
        low = t.lower()
        if low.isdigit():
            return "decimal"
        if not low.isalpha():
            # Punctuation or a mix: no counter family this tool can read.
            return None
        # The case of the token is the case of the counter that wrote it
        # (``A, B`` is ``upperLetter``; ``a, b`` is ``lowerLetter``).  A token
        # that mixes cases inside itself (``Aa``) is an artefact no counter
        # emits and is typed by its lowercase form.  A single letter that
        # could be Roman stays ``single`` -- case-free -- because its
        # alphabet is not settled until the run speaks.
        case = "upper" if t.isupper() else "lower"
        if len(low) > 1 and all(ch in _ROMAN_VALUE for ch in low):
            return "roman_" + case
        if len(low) > 1:
            return "letter_" + case
        # A single letter is ambiguous only if it could be a Roman numeral.
        # Seven letters can (i v x l c d m); every other letter -- a, b, e, f,
        # k, ... -- is unambiguously a Latin sequence number, and saying so
        # keeps an appendix numbered ``a, b, c`` from being reported as
        # unreadable.
        return "single" if low in _ROMAN_VALUE else "letter_" + case

    def _glyph_case(tok):
        """``"upper"`` / ``"lower"`` for an ASCII letter token, else ``None``.

        Decimal tokens carry no case, so a step between a letter and a
        decimal is never a *case* change; a token that mixes cases inside
        itself does not speak for either counter.
        """
        t = str(tok).strip()
        if not (t.isascii() and t.isalpha()):
            return None
        if t.isupper():
            return "upper"
        if t.islower():
            return "lower"
        return None

    # Two passes, because a lone letter cannot be typed until its neighbours
    # speak.  ``i`` opens either ``i, ii, iii`` (Roman) or ``i, j, k`` (Latin);
    # ``c`` is 100 in one counter and 3 in the other.  Pass 1 splits the stream
    # into runs at every *explicit* family change; a run's explicit members
    # decide what its single letters meant.  Pass 2 compares numbers inside
    # each run.  A run made only of singles never gets typed, so it is counted
    # as unread rather than silently assumed -- that distinction is what keeps
    # the verdict from going green on material nobody could read.
    families = [_family(tok) for tok in tokens]

    runs = []
    current = None
    for fam in families:
        if fam is None:
            current = None
            runs.append([None])
            continue
        if fam == "single":
            if current is None:
                current = []
                runs.append(current)
            current.append("single")
            continue
        explicit = {f for f in current or () if f != "single"}
        if current is None or (explicit and fam not in explicit):
            current = []
            runs.append(current)
        current.append(fam)

    unparsed = sum(1 for fam in families if fam is None)

    monotonic = True
    # A step down that coincides with a case change (``II -> i``, ``B -> a``)
    # is *observed* -- it drives ``monotonic`` False exactly like any other
    # backwards step -- but it is not *asserted*: from the bytes alone, a
    # counter restarted in the other case (``upperLetter`` front matter, then
    # ``lowerLetter`` body) cannot be told apart from a mis-cased token, so
    # the verdict is downgraded below for the same reason a declared
    # ``w:start`` restart is.  The observation runs over adjacent *readable*
    # tokens across runs too, not just within a run: the splitter hands a
    # lone ``i`` (family ``single``) to the previous run, where it still reads
    # as a drop under that run's counter -- the token pair, not the run edge,
    # is what carries the case change.  The loops below therefore no longer
    # break at the first backwards step: the case observation and the
    # ``unparsed`` count must see the whole stream.
    case_restart = False
    pos = 0
    prev_num = None
    prev_case = None
    for run in runs:
        if run == [None]:
            pos += 1
            continue
        kinds = {f for f in run if f != "single"}
        if len(kinds) == 1:
            fam = next(iter(kinds))
            settled = True
        else:
            # A run of nothing but single letters.  It still has to be
            # *judged*: skipping it is how a descending appendix (``c, a, b``,
            # ``d, c, b, a``) reported itself monotonic.  Latin is the right
            # default here -- a Roman sequence is written largest-first, so
            # reading these as Latin can only under-report a jump, never invent
            # one, and any real decrease shows up either way.  The run is
            # advisory (`unknown`) because the counter was inferred, not read.
            fam = "letter"
            settled = False
            unparsed += len(run)
        run_nums = []
        for _ in run:
            tok = tokens[pos]
            num = _token_number(tok, _FAMILY_FMT[fam])
            pos += 1
            if num is None:
                if settled:
                    unparsed += 1
                continue
            tok_case = _glyph_case(tok)
            if (prev_num is not None and num < prev_num
                    and prev_case and tok_case and prev_case != tok_case):
                monotonic = False
                case_restart = True
            prev_num = num
            prev_case = tok_case
            run_nums.append(num)
            if any(b < a for a, b in zip(run_nums, run_nums[1:])):
                monotonic = False

    # A green ``monotonic`` over a half-read TOC is not a clean verdict, so the
    # count is published as an integer and the boolean is *downgraded* when any
    # token could not be read.
    #
    # There is a second reason to refuse a verdict: a document that restarts its
    # counter mid-sequence.  math's TOC reads ``1, 1, 1, 3, 1, 1, 1`` -- section
    # 3 ends on page 3 and section 4 declares ``w:start="1"``, so the drop from 3
    # to 1 is correct.  From the bytes alone that is indistinguishable from a
    # genuinely broken TOC, and both readings use the same decimal family, so no
    # amount of glyph inspection separates them.  When any section declares an
    # explicit ``start``, a backwards step therefore downgrades the verdict to
    # ``unknown`` instead of asserting a regression -- an audit that cries wolf
    # on every restarted document is worse than one that admits it cannot tell.
    # Only a ``start`` on a section *after the first* restarts a counter.  The
    # first section's ``start="1"`` is where numbering begins, not a restart, so
    # counting it would downgrade every ordinary document to ``unknown``.
    #
    # There is a third reason to refuse a verdict, structurally identical to
    # the second: a case change inside one alphabet.  ``A, B, a, b`` (front
    # matter on ``upperLetter``, body restarted on ``lowerLetter``) reads
    # ``1, 2, 1, 2`` and the ``2 -> 1`` step is real in ``monotonic`` -- but
    # the bytes cannot say whether the second half is a restarted counter or
    # a mis-cased continuation, so ``case_restart`` downgrades the verdict
    # exactly the way a declared restart does.  A case change *without* a
    # step down (``a, b, E, F``) asserts nothing scary and stays ``true``; a
    # step down *without* a case change (``c, a, b``) is still ``false``.
    restarts = [s for s in (section_specs or [])
                if s.get("start") and (s.get("index") or 1) > 1]
    if monotonic and unparsed:
        verdict = "unknown"
    elif monotonic:
        verdict = "true"
    elif restarts or case_restart:
        verdict = "unknown"
    else:
        verdict = "false"

    return {
        "entries": entries,
        "entries_toc_style": broad,
        "styles": sorted(styles),
        "levels": sorted(levels),
        "page_tokens": tokens[:limit],
        "page_tokens_total": len(tokens),
        "page_parse_failures": failures[:limit],
        "monotonic": monotonic,
        "monotonic_verdict": verdict,
        "monotonic_across_sections": whole_document_monotonic,
        "monotonic_partial": unparsed > 0,
        "unparsed_tokens": unparsed,
        "page_number_restarts": [
            {"section": s.get("index"), "start": s.get("start"),
             "fmt": s.get("fmt")} for s in restarts],
        "sections": sorted(s for s in by_section if s is not None),
        "source": "toc" if broad else "computed",
    }


def _sectpr_pgnum(sectPr) -> dict:
    """``{"fmt": str|None, "start": str|None}`` from ``<w:sectPr>/<w:pgNumType>``.

    ``start`` is kept as the raw string the XML holds (``"1"``) rather than an
    int, so "absent" and "explicitly 1" stay distinguishable — that difference is
    the whole reason a section restarts numbering or continues from the previous
    one.  ``fmt`` is the Word counter name (``lowerRoman``, ``decimal``,
    ``upperLetter``, ``decimalZero``, …) that :func:`page_glyph` understands.

    **A ``sectPr`` may carry more than one** ``pgNumType`` and Word then obeys
    only the *last* one.  Measured on Word 16 with a section holding
    ``lowerRoman start=1`` followed by ``upperRoman start=5``: the page reads
    ``V``, i.e. the earlier element is dead.  That is exactly what happens when
    ``python-docx``'s ``add_section()`` clones a ``sectPr`` and the caller then
    appends a second ``pgNumType`` instead of replacing the first.  Reading the
    last element keeps this layer in agreement with what the reader sees; the
    duplicate count is reported separately by :func:`_section_specs` so the
    defect is still visible.
    """
    if sectPr is None:
        return {"fmt": None, "start": None, "count": 0}
    els = sectPr.findall(q("pgNumType"))
    if not els:
        return {"fmt": None, "start": None, "count": 0}
    el = els[-1]
    return {"fmt": el.get(q("fmt")), "start": el.get(q("start")),
            "count": len(els)}


def _section_specs(body, all_paras) -> list:
    """Per-section page-number spec plus the paragraph span each section covers.

    A ``<w:sectPr>`` is attached to the **last paragraph of the section it
    closes** (python-docx's ``add_section()`` puts it there too); only the final
    section's one hangs off the end of ``<w:body>``.  So a section runs from the
    paragraph after the previous ``<w:sectPr>`` through the paragraph that
    carries its own — inclusive at both ends, because the closing paragraph is
    still body text of that section.  That span is what makes "which section is
    paragraph N in" answerable without Word.

    Two index spaces are reported because this file uses both, and they are
    **named apart** so a caller cannot silently mix them
    (work/REVIEW-ADVERSARIAL-R6.md 应修-4):

    * ``first_para_index`` / ``last_para_index`` — 0-based ordinals among the
      body's **direct** ``<w:p>`` children (tables are not counted; a table's
      paragraphs live inside ``<w:tbl>``).  Same space as
      ``paragraphs_body_direct``.
    * ``first_para_index_xml`` / ``last_para_index_xml`` — 0-based positions in
      ``root.iter(w:p)``, the same space as ``numpr_paragraphs[].index_xml - 1``
      and ``_vertalign_claims[].para_index``.  Same space as
      ``paragraphs_xml``.

    ``index_space`` on each dict says which pair the *span* is in, so a caller
    never has to guess: it is ``"direct"`` (the pair the ``w:sectPr`` walk is
    itself defined on — a section is opened and closed by body-direct
    paragraphs) and both pairs are always present.  A consumer that wants the
    XML space asks for it explicitly instead of re-deriving it here:
    :func:`estimated_paragraph_pages` takes ``index_space="xml"`` and stamps the
    same value on its rows.

    A middle section whose closing paragraph carries the ``sectPr`` is closed by
    *that* paragraph.  The final section runs to the end of the body, so it is
    closed explicitly at ``len-1`` — otherwise a document whose last section has
    no trailing ``<w:sectPr>`` child element reports a span that stops short and
    every lookup past it returns ``None`` (应修-5).

    Neither equals a Word ``Paragraphs`` index exactly: Word inserts one extra
    paragraph per table row (see the module docstring, fact 2), so the Word layer
    carries its own index and never translates by arithmetic.
    """
    flat = {}
    for i, p in enumerate(all_paras):
        flat[id(p)] = i
    direct_flat = [flat.get(id(child)) for child in body if child.tag == q("p")]
    specs = []
    para_ord = 0
    start_ord = 0
    start_flat = 0

    def _emit(pgnum, first_ord, first_flat, last_ord, last_flat):
        specs.append({
            "index": len(specs) + 1,
            "first_para_index": first_ord,
            "first_para_index_xml": first_flat,
            # The paragraph carrying the closing ``w:sectPr`` — the section's own
            # last paragraph, which is what makes the range inclusive on both
            # ends and lets "which section is paragraph N in" be answered by a
            # simple ``first <= N <= last`` test.
            "last_para_index": last_ord,
            "last_para_index_xml": last_flat,
            "index_space": "direct",
            "fmt": pgnum["fmt"],
            "start": pgnum["start"],
            # >1 means the producer appended a ``pgNumType`` instead of editing
            # the one it cloned; Word obeys the last of them and silently drops
            # the others, so the count is reported rather than "cleaned up".
            "pgNumType_count": pgnum["count"],
        })

    seen_sectpr = False
    for child in body:
        if child.tag == q("p"):
            flat_pos = direct_flat[para_ord] if para_ord < len(direct_flat) \
                else None
            ppr = child.find(q("pPr"))
            sp = ppr.find(q("sectPr")) if ppr is not None else None
            if sp is not None:
                last_flat = flat_pos if flat_pos is not None else len(all_paras)
                _emit(_sectpr_pgnum(sp), start_ord, start_flat, para_ord,
                      last_flat)
                start_ord = para_ord + 1
                start_flat = last_flat + 1
                seen_sectpr = True
            para_ord += 1
        elif child.tag == q("sectPr"):
            _emit(_sectpr_pgnum(child), start_ord, start_flat, para_ord - 1,
                  max(len(all_paras) - 1, start_flat))
            seen_sectpr = True

    # The last section of a document has no closing paragraph and its
    # ``sectPr`` may be the body-level child handled above -- but a document can
    # also end with neither (python-docx always writes the body-level one, a
    # hand-built part may not).  Either way the final section reaches the end of
    # the body: leaving it short makes every lookup past its end return
    # ``section=None``, which downstream reads as "this document has no such
    # section" (应修-5).
    if para_ord > start_ord and (not specs or specs[-1]["last_para_index"] < para_ord - 1):
        _emit({"fmt": None, "start": None, "count": 0},
              start_ord, start_flat, para_ord - 1,
              max(len(all_paras) - 1, start_flat))
    return specs


def probe_xml(docx) -> dict:
    """Word-free structural probe.  Never raises; failures land in ``errors``.

    Keys are frozen for the life of the project — Wave 2 agents assert on them.
    """
    path = Path(docx)
    rep = {
        "xml_ok": False,
        "xml_error": None,
        "part": DOCUMENT_PART,
        "parts": 0,
        "media": 0,
        "paragraphs_xml": 0,
        "paragraphs_body_direct": 0,
        "paragraphs_in_tables": 0,
        "paragraphs_in_textboxes": 0,
        "table_rows": 0,
        "tables_xml": 0,
        "sections_xml": 0,
        "section_specs": [],
        "page_number_formats": [],
        "page_number_restarts": [],
        "bookmarks_xml": 0,
        "bookmarks_internal_xml": 0,
        "bookmark_names_xml": [],
        "drawings_xml": 0,
        "inline_drawings_xml": 0,
        "anchored_drawings_xml": 0,
        "images_xml": 0,
        "fields_xml": 0,
        "field_list_xml": [],
        "hyperlinks_xml": 0,
        "anchor_hyperlinks_xml": 0,
        "hyperlink_anchors": [],
        "external_hyperlinks_xml": 0,
        "headers_xml": 0,
        "footers_xml": 0,
        "numpr_paragraphs": [],
        "number_formats": {},
        "style_usage": [],
        "toc_styles": [],
        "page_numbers": [],
        "page_numbers_source": "none",
        "xml_warnings": [],
    }
    try:
        size = path.stat().st_size
    except OSError as exc:
        rep["xml_error"] = "file not readable: %s" % exc
        return rep
    if size == 0:
        rep["xml_error"] = "file is empty"
        return rep

    try:
        zf = zipfile.ZipFile(str(path))
    except (zipfile.BadZipFile, OSError) as exc:
        rep["xml_error"] = "not a zip container: %s" % exc
        return rep

    with zf:
        names = zf.namelist()
        rep["parts"] = len(names)
        rep["media"] = sum(1 for n in names if n.startswith("word/media/"))
        rep["headers_xml"] = sum(1 for n in names
                                 if re.match(r"word/header\d+\.xml$", n))
        rep["footers_xml"] = sum(1 for n in names
                                 if re.match(r"word/footer\d+\.xml$", n))
        try:
            raw = _read_part(zf, DOCUMENT_PART)
        except DocxProbeError as exc:
            rep["xml_error"] = str(exc)
            return rep
        num_to_abs, abs_lvl = _numbering_formats(zf)
        style_names = _style_names(zf)

    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        rep["xml_error"] = "document.xml is not well-formed XML: %s" % exc
        return rep

    body = root.find(q("body"))
    if body is None:
        rep["xml_error"] = "document.xml has no <w:body>"
        return rep

    parents = {id(child): parent for parent in root.iter() for child in parent}

    def has_ancestor(el, tag) -> bool:
        cur = parents.get(id(el))
        while cur is not None:
            if cur.tag == tag:
                return True
            cur = parents.get(id(cur))
        return False

    paras = list(root.iter(q("p")))
    rep["paragraphs_xml"] = len(paras)
    rep["paragraphs_body_direct"] = sum(1 for c in body if c.tag == q("p"))
    rep["paragraphs_in_tables"] = sum(1 for p in paras if has_ancestor(p, q("tc")))
    rep["paragraphs_in_textboxes"] = sum(
        1 for p in paras if has_ancestor(p, q("txbxContent")))
    rep["tables_xml"] = sum(1 for t in root.iter(q("tbl"))
                            if has_ancestor(t, q("body")))
    rep["table_rows"] = sum(1 for _ in root.iter(q("tr")))
    rep["sections_xml"] = sum(1 for _ in root.iter(q("sectPr")))

    # --- per-section page numbering ----------------------------------------
    # Page numbering is a *section* property, never a document property, so a
    # page number opinion is worthless without knowing which section said what.
    # ``section_specs`` is the Word-free answer: one row per section, in
    # document order, carrying the raw ``w:pgNumType`` and the paragraph where
    # the section starts.  ``page_number_formats`` / ``page_number_restarts`` are
    # flat conveniences over the same rows — a reader asking "is this document
    # numbered i, ii, 1, 2?" needs the fmt list and, separately, the sections
    # that explicitly restart (``start`` present), because an absent ``start``
    # means "continue from the previous section", not "restart at 1".
    rep["section_specs"] = _section_specs(body, paras)
    fmts = []
    restarts = []
    dup_pgnum = []
    for spec in rep["section_specs"]:
        if spec["fmt"] and spec["fmt"] not in fmts:
            fmts.append(spec["fmt"])
        if spec["start"] is not None:
            restarts.append({"section": spec["index"],
                             "start": spec["start"], "fmt": spec["fmt"]})
        if spec.get("pgNumType_count", 0) > 1:
            dup_pgnum.append({"section": spec["index"],
                              "count": spec["pgNumType_count"]})
    rep["page_number_formats"] = fmts
    rep["page_number_restarts"] = restarts
    if dup_pgnum:
        rep["xml_warnings"].append(
            "section(s) %s carry more than one w:pgNumType; Word obeys only the "
            "last one and silently ignores the rest"
            % ([d["section"] for d in dup_pgnum],))

    # --- bookmarks ---------------------------------------------------------
    names = []
    internal = 0
    for bm in root.iter(q("bookmarkStart")):
        nm = bm.get(q("name")) or ""
        if nm.startswith("_"):
            internal += 1
        elif nm:
            names.append(nm)
    rep["bookmarks_xml"] = len(names) + internal
    rep["bookmarks_internal_xml"] = internal
    rep["bookmark_names_xml"] = names

    # --- duplicate bookmark names: a defect XML alone cannot see ------------
    # Word *silently merges* two bookmarks that share a name: both
    # ``bookmarkStart`` elements stay in the XML, but Word keeps only one, so a
    # ``PAGEREF``/hyperlink aimed at the second one lands on the first.  The
    # document looks right and every ``w:name`` count looks right; only the
    # click goes to the wrong place.  Measured: two ``_TocDup`` in the XML at
    # character 0 and 10, Word reports one bookmark at Start=10.
    all_bm = [bm.get(q("name")) or "" for bm in root.iter(q("bookmarkStart"))]
    seen = {}
    for nm in all_bm:
        if nm:
            seen[nm] = seen.get(nm, 0) + 1
    dups = {nm: cnt for nm, cnt in seen.items() if cnt > 1}
    rep["bookmark_duplicates"] = dups
    rep["bookmark_starts_total"] = len([n for n in all_bm if n])
    rep["bookmark_names_unique_xml"] = len(seen)
    if dups:
        rep["xml_warnings"].append(
            "duplicate bookmark name(s) %s: Word silently merges these, so a "
            "reference to the later one jumps to the earlier one's position"
            % (sorted(dups)[:6],))

    # --- caption bookkeeping (table/figure of contents depends on these) ----
    # Word's "Insert Table of Figures" filters by *paragraph style* and links
    # each entry with a PAGEREF to the ``_Ref…`` bookmark it expects round
    # every caption.  A caption that carries a SEQ field but neither of those
    # inserts into the list and then cannot be clicked, so count all three.
    ref_names = [n for n in names if n.startswith("_Ref")] or [
        (bm.get(q("name")) or "") for bm in root.iter(q("bookmarkStart"))
        if (bm.get(q("name")) or "").startswith("_Ref")]
    toc_anchor_names = [bm.get(q("name")) or "" for bm in root.iter(q("bookmarkStart"))
                        if (bm.get(q("name")) or "").startswith("_Toc")]
    seq_labels = {}
    for it in root.iter(q("instrText")):
        m = re.match(r"\s*SEQ\s+(\S+)", it.text or "")
        if m:
            seq_labels[m.group(1)] = seq_labels.get(m.group(1), 0) + 1
    caption_paras = 0
    for p in paras:
        _sid, sname = _para_style(p, style_names)
        if sname and _CAPTION_STYLE_RE.match(sname.replace(" ", "")):
            caption_paras += 1
    rep["caption_audit"] = {
        "seq_total": sum(seq_labels.values()),
        "seq_labels": seq_labels,
        "ref_bookmarks": len(set(ref_names)),
        "toc_bookmarks": len(set(toc_anchor_names)),
        "caption_styled_paragraphs": caption_paras,
        "toc_of_figures_fields": sum(
            1 for it in root.iter(q("instrText"))
            if re.match(r"\s*TOC\b", it.text or "") and
            re.search(r'\\c\s|\s"图"|\s"表"', it.text or "")),
    }
    # What actually decides whether a table of figures works, measured on Word
    # 16 by generating captions and letting Word insert the list itself:
    #
    #   * the list is filtered by **SEQ label** (``TOC \c "图"``), not by style;
    #   * Word builds its own ``_Toc…`` bookmark per entry and points every
    #     ``PAGEREF``/hyperlink at *that*, never at a pre-written ``_Ref…``;
    #   * so ``_Ref`` is not required, and a missing ``pStyle`` breaks only the
    #     caption's *appearance*, not its discoverability.
    #
    # The defect that really made the six papers' lists dead was ``SEQ 表``
    # being absent — with no SEQ field there is no entry to find at all.  Hence
    # ``entries_possible`` keys on the SEQ counters, and the style / _Ref flags
    # are reported as quality signals rather than as hard requirements.
    rep["caption_audit"]["entry_source_ok"] = bool(
        rep["caption_audit"]["seq_labels"])
    rep["caption_audit"]["cosmetic_ok"] = bool(caption_paras)
    rep["caption_audit"]["legacy_ref_anchor_present"] = bool(
        rep["caption_audit"]["ref_bookmarks"])

    # --- script position: real text superscript vs OMML formula script ------
    # The operator's acceptance test is "select the superscript, press
    # Ctrl+Shift+=, it must return to normal size".  Only a *text* run with
    # ``w:vertAlign`` behaves that way; an OMML ``m:sSup`` is a formula object
    # whose raised glyph is layout, so the shortcut stacks a second superscript
    # onto already-shrunken type and the character gets smaller (reported).
    sup_runs = sub_runs = 0
    sup_chars = sub_chars = 0
    sup_field_only = sub_field_only = 0
    for r in root.iter(q("r")):
        rpr = r.find(q("rPr"))
        if rpr is None:
            continue
        va = rpr.find(q("vertAlign"))
        if va is None:
            continue
        val = va.get(q("val"))
        if val not in ("superscript", "subscript"):
            continue
        # A cross-reference or SEQ field is split over five runs (begin,
        # instrText, separate, result, end) that all inherit the vertical
        # alignment, but only the *result* run carries visible characters.
        # Counting runs therefore over-reports by ~5x: one superscript ``[1]``
        # reads as five.  Characters are the honest unit.
        text = "".join(t.text or "" for t in r.iter(q("t")))
        n = len(text)
        if val == "superscript":
            sup_runs += 1
            if n:
                sup_chars += n
            else:
                sup_field_only += 1
        else:
            sub_runs += 1
            if n:
                sub_chars += n
            else:
                sub_field_only += 1
    omml_sup = sum(1 for _ in root.iter(qm("sSup")))
    omml_sub = sum(1 for _ in root.iter(qm("sSub")))
    total_script = sup_chars + sub_chars + omml_sup + omml_sub

    # Candidate runs for the Word oracle.  Word's character offsets count a
    # paragraph mark per paragraph (and one per tab / break / inline image),
    # while the XML walk only sees ``w:t`` text, so an absolute offset computed
    # here would drift by one character per preceding paragraph — enough to
    # land one character short of a two-character superscript.  Instead each
    # claim records the paragraph it lives in; the probe scans that paragraph's
    # range in Word and counts the scripted characters there, which needs no
    # offset agreement at all.
    claims = []
    for pidx, p in enumerate(root.iter(q("p"))):
        for r in p.iter(q("r")):
            rpr = r.find(q("rPr"))
            if rpr is None:
                continue
            va = rpr.find(q("vertAlign"))
            if va is None:
                continue
            val = va.get(q("val"))
            if val not in ("superscript", "subscript"):
                continue
            text = "".join(t.text or "" for t in r.iter(q("t")))
            if not text:
                continue
            claims.append({"val": val, "text": text, "para_index": pidx})
    rep["_vertalign_claims"] = claims[:DEFAULT_LIMIT]
    rep["vertalign_claims_total"] = len(claims)
    rep["superscript_audit"] = {
        # Runs vs characters: a field fragment (begin/instr/separate/result/end)
        # carries the alignment but no text, so ``*_runs`` counts ~5x the visible
        # script while ``*_chars`` counts what a reader actually sees.  The
        # ratio uses characters; runs are kept for schema-level debugging.
        "vertalign_superscript_runs": sup_runs,
        "vertalign_subscript_runs": sub_runs,
        "vertalign_superscript_chars": sup_chars,
        "vertalign_subscript_chars": sub_chars,
        "vertalign_field_fragment_runs": sup_field_only + sub_field_only,
        "omml_sup_objects": omml_sup,
        "omml_sub_objects": omml_sub,
        "text_script_total": sup_chars + sub_chars,
        "omml_script_total": omml_sup + omml_sub,
        "standard_ratio": (round((sup_chars + sub_chars) / total_script, 4)
                           if total_script else None),
        "word_checked": False,
        "word_superscript_true": None,
        "word_superscript_omml": None,
    }

    # --- drawings / images -------------------------------------------------
    reps_draw = 0
    inline = 0
    anchored = 0
    for dr in root.iter(q("drawing")):
        reps_draw += 1
        if dr.find(qwp("anchor")) is not None:
            anchored += 1
        elif dr.find(qwp("inline")) is not None:
            inline += 1
    rep["drawings_xml"] = reps_draw
    rep["inline_drawings_xml"] = inline
    rep["anchored_drawings_xml"] = anchored
    rep["images_xml"] = sum(1 for _ in root.iter(qa("blip")))

    # --- fields ------------------------------------------------------------
    fcount, fentries = _field_entries(root, DEFAULT_LIMIT)
    rep["fields_xml"] = fcount
    rep["field_list_xml"] = fentries

    # --- hyperlink wrappers (the clickability proxy) ------------------------
    wrappers = list(root.iter(q("hyperlink")))
    anchors = []
    external = 0
    for hl in wrappers:
        anchor = hl.get(q("anchor"))
        rid = hl.get("{%s}id" % NS_R)
        text = "".join(t.text or "" for t in hl.iter(q("t")))
        if anchor:
            anchors.append({"anchor": anchor, "text": _cap(text),
                            "wrapper": True})
        elif rid:
            external += 1
    rep["hyperlinks_xml"] = len(wrappers)
    rep["anchor_hyperlinks_xml"] = len(anchors)
    rep["external_hyperlinks_xml"] = external
    rep["hyperlink_anchors"] = anchors

    # --- automatic numbering ------------------------------------------------
    numpr = []
    formats = {}
    for idx, p in enumerate(paras, 1):
        ppr = p.find(q("pPr"))
        if ppr is None:
            continue
        np_ = ppr.find(q("numPr"))
        if np_ is None:
            continue
        ilvl_el = np_.find(q("ilvl"))
        numid_el = np_.find(q("numId"))
        ilvl = ilvl_el.get(q("val")) if ilvl_el is not None else "0"
        num_id = numid_el.get(q("val")) if numid_el is not None else None
        text_val = None
        if num_id is not None:
            aid = num_to_abs.get(num_id)
            lvl = (abs_lvl.get(aid) or {}).get(str(ilvl)) if aid else None
            if lvl:
                text_val = lvl.get("lvlText")
                formats.setdefault(str(num_id), {})[str(ilvl)] = text_val
        numpr.append({
            "index_xml": idx,
            "text": _cap(_para_text(p)),
            "num_id": num_id,
            "ilvl": str(ilvl),
            "lvl_text": text_val,
        })
    rep["numpr_paragraphs"] = numpr[:DEFAULT_LIMIT]
    rep["number_formats"] = formats

    # --- page numbers from the TOC, and the style census --------------------
    toc_pages = []
    usage = {}
    for p in paras:
        sid, name = _para_style(p, style_names)
        key = (sid, name) if sid else (None, "(no pStyle)")
        usage[key] = usage.get(key, 0) + 1
        if not name or not _TOC_STYLE_RE.match(name):
            continue
        tok = _toc_token(_para_text(p))
        if tok:
            toc_pages.append(tok)
    if toc_pages:
        rep["page_numbers"] = toc_pages[:DEFAULT_LIMIT]
        rep["page_numbers_source"] = "toc"
    rep["style_usage"] = [
        {"style_id": sid, "name": name, "count": cnt}
        for (sid, name), cnt in sorted(usage.items(),
                                       key=lambda kv: (-kv[1], str(kv[0][0])))
    ][:DEFAULT_LIMIT]
    rep["toc_styles"] = sorted({nm for (_sid, nm) in usage
                                if nm and _TOC_STYLE_RE.match(nm)})

    # --- table-of-contents audit ------------------------------------------
    # The TOC is the document's own claim about page numbers; auditing it here
    # keeps the claim and the evidence in one report.  ``toc_audit.page_tokens``
    # is the same sequence as ``page_numbers`` but deliberately *not* gated on a
    # successful token parse, so a TOC that renders without numbers shows up as
    # ``entries > 0`` with ``page_parse_failures`` full instead of looking empty.
    #
    # The section map is handed in so a token can be read under the counter its
    # own section declares: ``c`` is 3 in a Latin appendix and 100 as a Roman
    # numeral, and a counter that restarts at a format change is not a
    # regression (work/REVIEW-ADVERSARIAL-R6.md 阻塞-2).
    _specs = rep.get("section_specs") or []

    def _section_of_para(idx, _specs=_specs):
        """Which section the paragraph at ``root.iter(w:p)`` position *idx* is in.

        ``_toc_audit`` walks ``root.iter(w:p)``, so *idx* lives in the **xml**
        space -- and the span to test it against is ``*_xml``, not the
        body-direct one.  Mixing the two silently answers ``1`` for every
        paragraph once a table or a paragraph inside ``<w:tbl>`` shifts the two
        ordinal spaces apart (work/REVIEW-ADVERSARIAL-R6.md 应修-4).
        """
        for spec in _specs:
            first = spec.get("first_para_index_xml")
            last = spec.get("last_para_index_xml")
            if first is None or last is None:
                continue
            if first <= idx <= last:
                return spec.get("index")
        return None

    rep["toc_audit"] = _toc_audit(paras, style_names, section_of=_section_of_para,
                                  section_specs=_specs)
    if rep["page_number_restarts"]:
        rep["xml_warnings"].append(
            "page numbering restarts in section(s) %s; a page number read from "
            "the TOC or a footer is the *displayed* number, not the physical "
            "page — keep both when comparing"
            % ([r["section"] for r in rep["page_number_restarts"]],))

    if rep["paragraphs_in_textboxes"]:
        rep["xml_warnings"].append(
            "document has %d paragraph(s) inside text boxes; Word's "
            "Paragraphs.Count may include them differently than the XML walk"
            % rep["paragraphs_in_textboxes"])

    rep["xml_ok"] = True
    return rep


# ===========================================================================
# Word COM layer
# ===========================================================================
_STATE = {"app": None, "win32": None, "avail": None, "atexit": False}


def _import_win32():
    """Import pywin32 lazily so the XML layer needs neither it nor Office."""
    if _STATE["win32"] is None:
        import win32com.client as win32   # noqa: WPS433 (deliberately lazy)
        _STATE["win32"] = win32
    return _STATE["win32"]


def word_available() -> bool:
    if _STATE["avail"] is None:
        try:
            _import_win32()
            _STATE["avail"] = True
        except Exception:
            _STATE["avail"] = False
    return bool(_STATE["avail"])


def reset_word_cache() -> None:
    """Drop cached availability/app state (tests and long-lived processes)."""
    _STATE["win32"] = None
    _STATE["avail"] = None


def get_app():
    """The one Word application this process will ever create.

    ``win32.Dispatch`` — *not* ``gencache.EnsureDispatch``: the makepy cache
    breaks on a second call in the same process with "RPC server unavailable"
    followed by "can not automate the makepy process".  Cached at module level,
    quit exactly once from ``release_app``.
    """
    if _STATE["app"] is None:
        win32 = _import_win32()
        app = win32.Dispatch("Word.Application")
        try:
            app.Visible = False        # never steal the user's screen
            app.DisplayAlerts = 0      # 0 = wdAlertsNone, no repair dialogs
        except Exception:
            pass
        _STATE["app"] = app
        if not _STATE["atexit"]:
            atexit.register(release_app)
            _STATE["atexit"] = True
    return _STATE["app"]


def release_app() -> None:
    """Quit Word if this process started it.  Safe to call repeatedly."""
    app = _STATE["app"]
    _STATE["app"] = None
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass


def _mm(points) -> float:
    try:
        return round(float(points) * 25.4 / 72.0, 2)
    except (TypeError, ValueError):
        return None


def _word_list_paragraphs(doc, mode: str, limit: int):
    """Rendered list strings, one COM call per *list* paragraph only.

    ``mode="numpr"`` uses Word's native ``Document.ListParagraphs`` collection,
    which is what makes the 16-reference case cost 16 round trips instead of one
    per paragraph in a 287-paragraph document.
    ``mode="all"`` is the slow, exhaustive scan kept as an escape hatch.
    """
    out = []
    errors = []
    if mode == "all":
        total = doc.Paragraphs.Count
        for i in range(1, total + 1):
            pr = doc.Paragraphs(i).Range
            s = pr.ListFormat.ListString
            if s:
                out.append({"index": i, "list_string": str(s),
                            "text": _cap(pr.Text.replace("\r", "").replace("\x07", ""))})
                if len(out) >= limit:
                    break
        return out, errors
    try:
        coll = doc.ListParagraphs
        total = coll.Count
    except Exception as exc:
        errors.append("Document.ListParagraphs unavailable: %s" % exc)
        return out, errors
    for i in range(1, total + 1):
        if len(out) >= limit:
            break
        try:
            pr = coll(i).Range
            fmt = pr.ListFormat
            try:
                global_index = doc.Range(0, pr.Start).Paragraphs.Count + 1
            except Exception:
                global_index = None
            out.append({
                "index": global_index,
                "list_string": str(fmt.ListString),
                "list_type": int(fmt.ListType),
                "ilvl": int(fmt.ListLevelNumber),
                "text": _cap(pr.Text.replace("\r", "").replace("\x07", "")),
            })
        except Exception as exc:
            errors.append("ListParagraphs(%d) failed: %s" % (i, exc))
    return out, errors


def _pdf_from_word(doc, pdf_path):
    """PDF export; escalate from ExportAsFixedFormat to SaveAs2 on failure."""
    errors = []
    warnings = []
    try:
        doc.ExportAsFixedFormat(str(pdf_path), WD_EXPORT_FORMAT_PDF)
    except Exception as exc:
        warnings.append("ExportAsFixedFormat failed (%s); trying SaveAs2" % exc)
        try:
            doc.SaveAs2(str(pdf_path), FileFormat=WD_EXPORT_FORMAT_PDF)
        except Exception as exc2:
            errors.append("PDF export failed: %s" % exc2)
    if Path(pdf_path).exists():
        return str(pdf_path), errors, warnings
    return None, errors, warnings


def _render_pngs(pdf_path, png_dir, dpi: int, limit: int = 400):
    """Per-page PNG via PyMuPDF.  Returns ``(paths, pages, footers, errors)``."""
    paths = []
    footers = []
    errors = []
    pages = None
    try:
        import fitz
    except Exception as exc:
        return paths, pages, footers, ["PyMuPDF unavailable: %s" % exc]
    out_dir = Path(png_dir)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
        with fitz.open(str(pdf_path)) as pdf:
            pages = pdf.page_count
            for i, page in enumerate(pdf, 1):
                if i > limit:
                    break
                target = out_dir / ("page-%03d.png" % i)
                page.get_pixmap(dpi=dpi).save(str(target))
                paths.append(str(target))
                band = fitz.Rect(0, page.rect.height * 0.88,
                                 page.rect.width, page.rect.height)
                footers.append(" ".join(page.get_text("text", clip=band).split()))
    except Exception as exc:
        errors.append("PNG rendering failed: %s" % exc)
    return paths, pages, footers, errors


def probe_word(docx, pdf_path=None, png_dir=None, dpi: int = 110,
               update_fields: bool = True, numbering_scan: str = "numpr",
               limit: int = DEFAULT_LIMIT, vertalign_claims=None,
               page_probe_indices: list[int] | None = None,
               section_formats=None) -> dict:
    """Word COM probe.  Never raises: every failure becomes ``skipped``/``errors``.

    Two page-number facts are kept apart on purpose, because conflating them is
    how a report ends up telling a student their paper's page 7 is "i":

    * ``WD_ACTIVE_END_ADJUSTED_PAGE`` (1) is the **displayed** number — the one
      the reader sees in the footer, already adjusted by ``w:pgNumType``.  A page
      printed ``iii`` reports ``3`` here.  Turn it into the glyph with
      :func:`page_glyph` using that section's ``fmt``.
    * ``WD_ACTIVE_END_PAGE`` (3) is the **physical** sheet ordinal, counted from
      page 1 of the document and ignoring every restart.

    ``page_probe_indices`` turns on per-paragraph probing of both.  It is opt-in
    because each paragraph costs two COM round trips; leave it ``None`` and
    ``out["paragraph_pages"]`` stays empty with ``paragraph_pages_source``
    ``"none"``.  This probe **never invents** a page number: when Word is absent
    or nothing was asked for, the caller can consult
    :func:`estimated_paragraph_pages`, whose output is labelled ``computed`` and
    is a structural estimate from ``w:pgNumType`` — it **must not be delivered
    as a page number to a reader**, because it knows where a section starts but
    not where a paragraph falls inside it.

    ``section_formats`` is the per-section ``w:pgNumType/@fmt`` list (1-based
    section order, from ``probe_xml()["section_specs"]``).  Word does not expose
    ``pgNumType`` through its object model, so the format has to come from the
    XML layer; passing it is what makes ``displayed_text`` a glyph rather than a
    bare number.
    """
    out = {
        "word_ok": False,
        "skipped": False,
        "skipped_reason": None,
        "errors": [],
        "warnings": [],
        "pages": None, "paragraphs": None, "sections": None, "bookmarks": None,
        "bookmarks_total": None, "inline_shapes": None, "floating_shapes": None,
        "fields": None, "hyperlinks": None, "tables": None,
        "bookmark_names": [], "bookmark_names_all": [],
        "hyperlink_targets": [], "field_list": [],
        "page_numbers": [], "page_numbers_source": "none",
        "section_details": [],
        "paragraph_pages": [], "paragraph_pages_source": "none",
        "pdf": None, "pngs": [], "pdf_pages": None, "pdf_page_footers": [],
        "word_version": None,
    }
    if not word_available():
        out["skipped"] = True
        out["skipped_reason"] = ("Word COM unavailable (pywin32 import failed "
                                 "or Office is not installed)")
        out["errors"].append(out["skipped_reason"])
        return out

    doc = None
    try:
        app = get_app()
        try:
            out["word_version"] = str(app.Version)
        except Exception:
            pass
        doc = app.Documents.Open(str(Path(docx).resolve()), ReadOnly=True,
                                 AddToRecentFiles=False)
        if doc is None:
            raise RuntimeError("Word returned no document object")
        try:
            doc.Visible = False
        except Exception:
            pass

        if update_fields:
            try:
                doc.Fields.Update()
            except Exception as exc:
                out["warnings"].append("Fields.Update failed: %s" % exc)

        # --- statistics ----------------------------------------------------
        try:
            out["pages"] = int(doc.ComputeStatistics(WD_STAT_PAGES))
        except Exception as exc:
            out["warnings"].append("ComputeStatistics(pages) failed: %s" % exc)
        out["paragraphs"] = int(doc.Paragraphs.Count)
        out["sections"] = int(doc.Sections.Count)
        out["tables"] = int(doc.Tables.Count)
        out["fields"] = int(doc.Fields.Count)
        out["hyperlinks"] = int(doc.Hyperlinks.Count)
        try:
            out["inline_shapes"] = int(doc.InlineShapes.Count)
            out["floating_shapes"] = int(doc.Shapes.Count)
        except Exception as exc:
            out["warnings"].append("shape count failed: %s" % exc)
        out["bookmarks_total"] = int(doc.Bookmarks.Count)

        # --- bookmarks -----------------------------------------------------
        names_all, names_user = [], []
        for i in range(1, min(out["bookmarks_total"], limit) + 1):
            try:
                nm = str(doc.Bookmarks(i).Name)
            except Exception:
                continue
            names_all.append(nm)
            if not nm.startswith("_"):
                names_user.append(nm)
        out["bookmarks"] = out["bookmarks_total"]
        out["bookmark_names_all"] = names_all
        out["bookmark_names"] = names_user
        if out["bookmarks_total"] > limit:
            out["warnings"].append("bookmarks truncated at limit=%d" % limit)

        # --- script position, measured the way Word measures it ------------
        # The distinguishing signal, taken on Word 16 against a document with
        # both kinds of script (the operator's report and this project's OMML
        # pipeline):
        #
        #   OMML formula superscript -> Font.Superscript 0, Size shrunk by the
        #                              formula engine, OMaths.Count >= 1
        #   real text superscript    -> Font.Superscript -1, Size unchanged,
        #                              OMaths.Count 0   (Ctrl+Shift+= toggles)
        #
        # So "real superscript" is: Superscript == -1 AND the character is not
        # inside an OMath range.  Size is recorded rather than compared, because
        # Word reports the *rendered* size for formula glyphs.
        try:
            out["omaths"] = int(doc.OMaths.Count)
        except Exception:
            out["omaths"] = None
        omml_starts = []
        omml_ends = []
        if out["omaths"]:
            for i in range(1, min(int(out["omaths"]), limit) + 1):
                try:
                    rng = doc.OMaths(i).Range
                    omml_starts.append(int(rng.Start))
                    omml_ends.append(int(rng.End))
                except Exception:
                    continue

        def _in_omath(pos: int) -> bool:
            for a, b in zip(omml_starts, omml_ends):
                if a <= pos < b:
                    return True
            return False

        # Field *code* text (``REF ref_1 \n``, ``HYPERLINK \l …``) inherits the
        # field's formatting and reports Superscript -1, so a naive per-character
        # walk "finds" dozens of real superscripts spelling out field instructions.
        # Those ranges are not body content, so they are excluded.
        code_ranges = []
        try:
            for i in range(1, int(doc.Fields.Count) + 1):
                try:
                    cr = doc.Fields(i).Code
                    code_ranges.append((int(cr.Start), int(cr.End)))
                except Exception:
                    continue
        except Exception:
            pass

        def _in_field_code(pos: int) -> bool:
            for a, b in code_ranges:
                if a <= pos < b:
                    return True
            return False

        true_sup = true_sub = omml_script = 0
        field_code_chars = 0
        size_sample = []
        scanned = 0
        try:
            total_chars = int(doc.Content.End)
        except Exception:
            total_chars = 0
        # A per-character COM round trip costs ~1 ms, so walking a 4 000
        # character body is 4 000 RPCs — long enough that Word's automation
        # channel starts answering ``被呼叫方拒绝接收呼叫`` (measured: a clean
        # document produced 20 such errors and a verdict of "none" while
        # ``OMaths.Count`` still read 50).  So Word is used as an *oracle on
        # candidates*: the XML layer already lists every run that claims a
        # ``w:vertAlign``, and only those characters are sampled here.
        #
        # ``Range.Font.Superscript`` on a *multi-character* range returns
        # ``9999999`` (``wdUndefined``) when the characters disagree, so each
        # character inside a candidate is read on its own — a whole-range read
        # cannot tell "all superscript" from "superscript starts here", and
        # would report a real superscript as undefined.
        probed = 0
        max_probe = SCRIPT_PROBE_BUDGET
        for claim in (vertalign_claims or []):
            if probed >= max_probe:
                out["warnings"].append(
                    "script probe truncated at %d candidates" % max_probe)
                break
            pidx = claim.get("para_index")
            if pidx is None:
                continue
            try:
                prng = doc.Paragraphs(int(pidx) + 1).Range
                pstart, pend = int(prng.Start), int(prng.End)
            except Exception:
                continue
            sup = sub = 0
            size = None
            first = True
            for pos in range(pstart, pend):
                try:
                    rng = doc.Range(pos, pos + 1)
                    one_sup = int(rng.Font.Superscript)
                    one_sub = int(rng.Font.Subscript)
                except Exception:
                    continue
                if first:
                    try:
                        size = float(rng.Font.Size)
                    except Exception:
                        size = None
                    first = False
                if one_sup == -1:
                    sup += 1
                if one_sub == -1:
                    sub += 1
            probed += 1
            scanned += sup + sub
            inside = _in_omath(pstart)
            if inside:
                omml_script += sup + sub
            else:
                true_sup += sup
                true_sub += sub
            if len(size_sample) < 8:
                size_sample.append({"char": _cap(claim.get("text"), 6),
                                    "xml_claim": claim.get("val"),
                                    "paragraph": pidx,
                                    "word_superscript_chars": sup,
                                    "word_subscript_chars": sub,
                                    "word_size": size,
                                    "in_omath": inside,
                                    "agree": bool(
                                        (claim.get("val") == "superscript"
                                         and sup > 0)
                                        or (claim.get("val") == "subscript"
                                            and sub > 0))})
        audit = out.setdefault("superscript_audit", {})
        audit.update({
            "word_checked": True,
            "word_scanned_chars": scanned,
            "word_field_code_chars": field_code_chars,
            "word_scan_cap": SCRIPT_PROBE_BUDGET,
            "word_probed_candidates": probed,
            "word_superscript_true": true_sup,
            "word_subscript_true": true_sub,
            "word_superscript_omml": omml_script,
            "word_text_script_total": true_sup + true_sub,
            "word_size_sample": size_sample,
            "word_verdict": ("real-text-script" if (true_sup + true_sub)
                             else ("omml-only" if omml_script else "none")),
        })

        # --- hyperlinks: the clickability signal ---------------------------
        targets = []
        for i in range(1, min(out["hyperlinks"], limit) + 1):
            try:
                h = doc.Hyperlinks(i)
                text = h.TextToDisplay
                if not text:
                    # Internal anchor links report no TextToDisplay; the visible
                    # text is the cached field result, e.g. "1".
                    text = h.Range.Text
                # Superscript is wdToggle (9999999) or True (-1) when on, 0 when
                # off.  The owner asked for both variants explicitly, so the
                # acceptor reports which one each reference uses.
                try:
                    sup = int(h.Range.Font.Superscript)
                except Exception:
                    sup = None
                targets.append({
                    "sub_address": str(h.SubAddress or ""),
                    "text": _cap(str(text).replace("\r", "").replace("\x07", "")),
                    "address": str(h.Address or ""),
                    "type": int(h.Type) if h.Type is not None else None,
                    "superscript": sup,
                    "superscript_on": (None if sup is None else sup != 0),
                })
            except Exception as exc:
                out["errors"].append("Hyperlinks(%d) failed: %s" % (i, exc))
        out["hyperlink_targets"] = targets

        # --- fields --------------------------------------------------------
        flist = []
        pagerefs = []
        for i in range(1, min(out["fields"], limit) + 1):
            try:
                f = doc.Fields(i)
                code = _cap(f.Code.Text, 120)
                result = _cap(f.Result.Text)
            except Exception as exc:
                out["errors"].append("Fields(%d) failed: %s" % (i, exc))
                continue
            flist.append({"code": code, "result": result})
            if "PAGEREF" in code.upper():
                pagerefs.append(result)
        out["field_list"] = flist
        if pagerefs:
            out["page_numbers"] = pagerefs
            out["page_numbers_source"] = "pageref"
        if out["fields"] > limit:
            out["warnings"].append("fields truncated at limit=%d" % limit)

        # --- numbering -----------------------------------------------------
        lists, lerr = _word_list_paragraphs(doc, numbering_scan, limit)
        out["list_paragraphs"] = lists
        out["errors"].extend(lerr)

        # --- per-section geometry and page numbers -------------------------
        for si in range(1, out["sections"] + 1):
            sec = doc.Sections(si)
            info = {"index": si}
            try:
                ps = sec.PageSetup
                info.update({
                    "page_width_mm": _mm(ps.PageWidth),
                    "page_height_mm": _mm(ps.PageHeight),
                    "top_margin_mm": _mm(ps.TopMargin),
                    "bottom_margin_mm": _mm(ps.BottomMargin),
                    "left_margin_mm": _mm(ps.LeftMargin),
                    "right_margin_mm": _mm(ps.RightMargin),
                    "orientation": int(ps.Orientation),
                })
            except Exception as exc:
                info["error"] = "PageSetup failed: %s" % exc
            try:
                info["headers"] = int(sec.Headers.Count)
                info["footers"] = int(sec.Footers.Count)
            except Exception:
                pass
            try:
                info["last_page"] = int(
                    doc.Range(sec.Range.End - 1, sec.Range.End)
                    .Information(WD_ACTIVE_END_PAGE))
            except Exception:
                pass
            # Same page, read the other way: the number the footer shows.  On a
            # front-matter section numbered ``i, ii, iii`` this is 3 while
            # ``last_page`` above is 7 — the two are not interchangeable and the
            # pair is what lets a caller prove which one it has.
            try:
                disp = int(doc.Range(sec.Range.End - 1, sec.Range.End)
                           .Information(WD_ACTIVE_END_ADJUSTED_PAGE))
                info["last_page_displayed_number"] = disp
            except Exception:
                disp = None
            sec_fmt = None
            if section_formats:
                sec_fmt = section_formats[si - 1] if si - 1 < len(section_formats) \
                    else None
            info["page_format"] = sec_fmt
            if disp is not None:
                info["last_page_displayed_text"] = page_glyph(sec_fmt, disp)
            try:
                info["start_page"] = int(
                    doc.Range(sec.Range.Start, sec.Range.Start + 1)
                    .Information(WD_ACTIVE_END_PAGE))
            except Exception:
                pass
            footer_page = None
            for fi in WD_FOOTER_INDEXES:
                try:
                    ftr = sec.Footers(fi)
                    if not ftr.Exists:
                        continue
                    for fl in ftr.Range.Fields:
                        if "PAGE" in str(fl.Code.Text).upper():
                            footer_page = _cap(fl.Result.Text)
                            break
                except Exception:
                    continue
                if footer_page:
                    break
            if footer_page:
                info["footer_page_number"] = footer_page
            out["section_details"].append(info)

        # --- per-paragraph page numbers (opt-in) ---------------------------
        # Both readings are taken for the same paragraph and both are kept:
        # ``physical`` (WD_ACTIVE_END_PAGE, the sheet ordinal) and
        # ``displayed_number``/``displayed_text`` (WD_ACTIVE_END_ADJUSTED_PAGE,
        # the number in the footer, rendered through the section's pgNumType
        # fmt).  They disagree exactly where a section restarts or changes
        # format — which is the case this probe exists for.  Indices are Word
        # ``Paragraphs`` indices (1-based), not XML ones; ``section`` is the
        # 1-based section number the paragraph sits in, from Word's own range
        # comparison, so nothing is translated by arithmetic.
        wanted = list(page_probe_indices or [])
        if wanted:
            if len(wanted) > PAGE_PROBE_BUDGET:
                out["warnings"].append(
                    "page_probe_indices truncated to %d of %d entries"
                    % (PAGE_PROBE_BUDGET, len(wanted)))
                wanted = wanted[:PAGE_PROBE_BUDGET]
            sec_bounds = []
            for sd in out["section_details"]:
                try:
                    srange = doc.Sections(sd["index"]).Range
                    sec_bounds.append((sd["index"], int(srange.Start),
                                       int(srange.End)))
                except Exception:
                    continue

            def _section_of(pos):
                for sidx, a, b in sec_bounds:
                    if a <= pos < b:
                        return sidx
                return None

            rows = []
            for raw_idx in wanted:
                try:
                    idx = int(raw_idx)
                except (TypeError, ValueError):
                    out["warnings"].append(
                        "page_probe_indices: %r is not an integer" % (raw_idx,))
                    continue
                if idx < 1 or idx > int(out["paragraphs"] or 0):
                    out["warnings"].append(
                        "page_probe_indices: paragraph %d out of range 1..%s"
                        % (idx, out["paragraphs"]))
                    rows.append({"index": idx, "physical": None,
                                 "displayed_number": None,
                                 "displayed_text": None, "section": None,
                                 "error": "out of range"})
                    continue
                row = {"index": idx, "physical": None, "displayed_number": None,
                       "displayed_text": None, "section": None}
                try:
                    prng = doc.Paragraphs(idx).Range
                except Exception as exc:
                    row["error"] = "Paragraphs(%d) failed: %s" % (idx, exc)
                    rows.append(row)
                    continue
                try:
                    row["physical"] = int(
                        prng.Information(WD_ACTIVE_END_PAGE))
                except Exception as exc:
                    row["error"] = "Information(page) failed: %s" % exc
                sidx = _section_of(int(prng.Start))
                row["section"] = sidx
                fmt = None
                if sidx is not None and section_formats \
                        and sidx - 1 < len(section_formats):
                    fmt = section_formats[sidx - 1]
                row["page_format"] = fmt
                try:
                    disp = int(prng.Information(WD_ACTIVE_END_ADJUSTED_PAGE))
                except Exception as exc:
                    row.setdefault("error", "Information(adjusted page) "
                                             "failed: %s" % exc)
                    disp = None
                if disp is not None:
                    row["displayed_number"] = disp
                    row["displayed_text"] = page_glyph(fmt, disp)
                rows.append(row)
            out["paragraph_pages"] = rows
            # A row is not evidence: every out-of-range index still produces
            # one, with ``displayed_number = None``.  Labelling the whole batch
            # "word" on the strength of ``rows`` being non-empty would let a
            # caller trust a page number that was never measured -- the same
            # failure the computed/word split exists to prevent, just inverted
            # (work/REVIEW-ADVERSARIAL-R6.md 应修-3).
            measured = sum(1 for r in rows if r.get("displayed_number") is not None)
            if measured == 0:
                out["paragraph_pages_source"] = "none"
            elif measured < len(rows):
                out["paragraph_pages_source"] = "word-partial"
            else:
                out["paragraph_pages_source"] = "word"

        # --- PDF + PNG -----------------------------------------------------
        if pdf_path:
            pdf_path = Path(pdf_path)
            pdf_path.parent.mkdir(parents=True, exist_ok=True)
            got, perr, pwarn = _pdf_from_word(doc, pdf_path)
            out["errors"].extend(perr)
            out["warnings"].extend(pwarn)
            out["pdf"] = got
            if got and png_dir:
                pngs, pages, footers, rerr = _render_pngs(got, png_dir, dpi)
                out["pngs"] = pngs
                out["pdf_pages"] = pages
                out["pdf_page_footers"] = footers
                out["errors"].extend(rerr)
        out["word_ok"] = True
    except Exception as exc:
        msg = "%s: %s" % (type(exc).__name__, exc)
        out["errors"].append("Word could not open/probe the document: %s" % msg)
        out["word_ok"] = False
    finally:
        if doc is not None:
            try:
                doc.Close(False)      # never save
            except Exception:
                pass
    return out


# ===========================================================================
# merge / expectations / reporting
# ===========================================================================
def estimated_paragraph_pages(section_specs, section_details=None,
                              indices=None, index_space="xml") -> dict:
    """Structural page estimate used when Word is absent.  **Not a page number.**

    Compute what can honestly be computed from the bytes: which section each
    requested paragraph belongs to (``section_specs[].first_para_index_xml``), that
    section's counter format, and — when ``section_details`` came from a earlier
    Word run — where the section's *last* page was.  What cannot be computed is
    where inside the section a paragraph falls, because that needs layout.

    So the returned ``displayed_number``/``displayed_text`` are derived from the
    section's ``pgNumType`` **start**, not from the paragraph's position: they
    describe the section's first page number, and the row is marked
    ``exact: false`` with ``basis: "section-start"``.  The point of the block is
    that a caller can see *which section and which alphabet* a paragraph lives
    under without Word, and that every such row is stamped ``computed``.

    ``computed`` rows must never be delivered as a document's page numbers.  Use
    them to explain a structure, to sanity-check a Word reading, or to keep a
    report populated when Office is missing — never as "the page this paragraph
    is on".

    **The section mapping inherits the XML index space, not Word's.**  Word
    inserts extra paragraph marks for tables — measured as ``columns + 1`` per
    row, not a flat ``+1`` (work/REVIEW-ADVERSARIAL-R6.md §7.1) — so the two
    spaces drift and neither can be translated into ``Paragraphs(n)`` by
    arithmetic.  On ``work/subjects/math/paper.docx`` (one table, 130 ``w:p``,
    102 of them body-direct, 5 sections) the 0-based ordinal 95 gets *different*
    answers from the two spaces: body-direct puts it in section 5, the XML space
    in section 4 — and Word's ``Paragraphs(96)`` was measured in section 4
    (work/REVIEW-ADVERSARIAL-R6.md).  That is why the XML space is the default
    here: it is the space the rest of the report speaks, and on this document it
    is the one that agrees with Word.  Agreement is not a guarantee — the drift
    is structural — so :func:`build_report` still warns whenever the document
    has table rows, and every row still says ``exact: false``.  With Word
    present this whole function is bypassed: the Word layer resolves the section
    from the paragraph's own range, which has no drift.

    **Which span is compared is a parameter, and the label always says which.**
    ``_section_specs`` emits *two* pairs (``first/last_para_index`` over the
    body's direct ``<w:p>`` children and ``first/last_para_index_xml`` over
    ``root.iter(w:p)``), and mixing them silently maps nearly every paragraph to
    section 1 (work/REVIEW-ADVERSARIAL-R6.md 应修-4).  ``index_space`` picks the
    pair *and* is copied onto every row, so the label cannot drift from the
    arithmetic again — the previous revision compared the body-direct pair while
    stamping ``"xml"`` on the rows, which handed a caller section ``None`` for
    every paragraph living inside ``<w:tbl>`` (measured: a 1-column/3-row table
    made XML ordinals 4-5 unresolvable and mapped ordinal 3 to section 2 instead
    of 1; work/P3-REPORT-B.md).

    ``index_space="xml"`` is the default because it is the space the rest of the
    report already publishes (``paragraphs_xml``, ``numpr_paragraphs[].index_xml``,
    ``_vertalign_claims[].para_index``, and the ``_section_of_para`` lookup that
    ``toc_audit`` uses).  ``index_space="direct"`` is for a caller who wants
    "which body paragraph" and never looks inside tables.  Note that
    ``section_specs[].index_space`` stays ``"direct"``: it records the pair the
    ``w:sectPr`` walk is *defined* on, not the pair a consumer should use.

    Returns::

        {"source": "computed",
         "warning": "...computed 不得当作交付页码...",
         "index_space": "xml",
         "rows": [{"index": i, "section": sidx, "page_format": fmt,
                   "physical": None, "displayed_number": n,
                   "displayed_text": glyph, "exact": False,
                   "basis": "section-start", "index_space": "xml"}]}
    """
    specs = list(section_specs or [])
    spaces = {"xml": ("first_para_index_xml", "last_para_index_xml"),
              "direct": ("first_para_index", "last_para_index")}
    if index_space not in spaces:
        raise ValueError("index_space must be one of %s, got %r"
                         % (sorted(spaces), index_space))
    first_key, last_key = spaces[index_space]
    bounds = []
    for pos, spec in enumerate(specs):
        start = spec.get(first_key)
        end = spec.get(last_key)
        nxt = (specs[pos + 1].get(first_key)
               if pos + 1 < len(specs) else None)
        bounds.append((spec.get("index", pos + 1), start, end, nxt,
                       spec.get("fmt"), spec.get("start")))
    last_page = {}
    for sd in (section_details or []):
        if sd.get("index") is not None:
            last_page[sd["index"]] = sd.get("last_page")

    rows = []
    for raw_idx in (indices or []):
        try:
            idx = int(raw_idx)
        except (TypeError, ValueError):
            continue
        # Same index convention as ``probe_word``: Word ``Paragraphs`` indices
        # are 1-based, the section spans are 0-based ordinals in the space
        # ``index_space`` names.  Converting in one place keeps the two layers
        # comparable — the same ``--page-probe`` list can be read by either and
        # mean the same paragraph.
        ordinal = idx - 1
        hit = None
        for sidx, start, end, nxt, fmt, st in bounds:
            if start is None:
                continue
            # Prefer the inclusive span; fall back to [start, next start) when a
            # spec has no closing paragraph (a malformed document, or a section
            # whose sectPr the producer dropped).
            if end is not None:
                inside = start <= ordinal <= end
            else:
                inside = ordinal >= start and (nxt is None or ordinal < nxt)
            if inside:
                hit = (sidx, fmt, st)
                break
        if hit is None:
            rows.append({"index": idx, "section": None, "page_format": None,
                         "physical": None, "displayed_number": None,
                         "displayed_text": None, "exact": False,
                         "basis": "unknown-section",
                         "index_space": index_space})
            continue
        sidx, fmt, st = hit
        num = None
        try:
            num = int(st) if st not in (None, "") else 1
        except (TypeError, ValueError):
            num = None
        rows.append({
            "index": idx,
            "section": sidx,
            "page_format": fmt,
            "physical": None,
            "displayed_number": num,
            "displayed_text": page_glyph(fmt, num) if num is not None else None,
            "section_last_page": last_page.get(sidx),
            "exact": False,
            "basis": "section-start",
            "index_space": index_space,
        })
    return {
        "source": "computed",
        "index_space": index_space,
        "warning": ("computed: 结构性估算（按节起始页码推），不是排版结果；"
                    "computed 不得当作交付页码"),
        "rows": rows,
    }


def _merge_numbering(xml_entries, word_lists, limit):
    """Zip ``numPr`` paragraphs with Word's list paragraphs **by ordinal**.

    Paragraph indices cannot be translated between the two layers (Word counts
    one extra paragraph per table row), but both sequences are in document
    order, so pairing them positionally is exact as long as the counts agree.
    Disagreements are reported, never silently hidden.
    """
    out = []
    pairs = min(len(xml_entries), len(word_lists))
    mismatched = 0
    for i in range(pairs):
        x = xml_entries[i]
        w = word_lists[i]
        xt = (x.get("text") or "")[:12]
        wt = (w.get("text") or "")[:12]
        match = bool(xt) and bool(wt) and xt == wt
        if not match:
            mismatched += 1
        out.append({
            "index": w.get("index"),
            "index_xml": x.get("index_xml"),
            "index_source": "word" if w.get("index") is not None else "xml",
            "list_string": w.get("list_string"),
            "list_type": w.get("list_type"),
            "ilvl": x.get("ilvl"),
            "num_id": x.get("num_id"),
            "lvl_text": x.get("lvl_text"),
            "rendered": bool(w.get("list_string")),
            "text_match": match,
            "text": w.get("text") or x.get("text"),
        })
    for x in xml_entries[pairs:]:
        out.append({
            "index": None, "index_xml": x.get("index_xml"),
            "index_source": "xml", "list_string": None, "list_type": None,
            "ilvl": x.get("ilvl"), "num_id": x.get("num_id"),
            "lvl_text": x.get("lvl_text"), "rendered": False,
            "text_match": False, "text": x.get("text"),
            "note": "numPr in XML but Word reports no list paragraph here",
        })
    for w in word_lists[pairs:]:
        out.append({
            "index": w.get("index"), "index_xml": None, "index_source": "word",
            "list_string": w.get("list_string"), "list_type": w.get("list_type"),
            "ilvl": w.get("ilvl"), "num_id": None, "lvl_text": None,
            "rendered": bool(w.get("list_string")), "text_match": False,
            "text": w.get("text"),
            "note": "Word list paragraph with no matching numPr in XML",
        })
    return out[:limit], mismatched


EXPECT_ALIASES = {
    "sections": "sections",
    "bookmarks": "bookmarks",
    "params": "paragraphs",
    "paragraphs": "paragraphs",
    "pages": "pages",
    "tables": "tables",
    "fields": "fields",
    "hyperlinks": "hyperlinks",
    "inline-shapes": "inline_shapes",
}


def build_report(args) -> dict:
    """Run the layers the arguments ask for and return the merged report."""
    docx = Path(args.docx)
    xml = probe_xml(docx)

    word = None
    skipped = False
    want_word = args.word != "off"
    # Do not hand a structurally broken file to Word: that is how repair
    # dialogs and half-hour hangs happen.
    if want_word and args.word == "auto" and not xml["xml_ok"]:
        want_word = False
        skipped = True
    if want_word:
        word = probe_word(docx, pdf_path=args.pdf, png_dir=args.png,
                          dpi=args.dpi, update_fields=not args.no_update_fields,
                          numbering_scan=args.numbering_scan,
                          limit=args.limit_list,
                          vertalign_claims=xml.get("_vertalign_claims"),
                          page_probe_indices=getattr(args, "page_probe_indices",
                                                     None),
                          section_formats=[s["fmt"] for s in
                                           xml.get("section_specs") or []])
        if word["skipped"]:
            want_word = False
            skipped = True
    else:
        skipped = True

    errors = []          # everything the caller should read, skip reasons included
    hard_errors = []     # only what may turn ok into false
    warnings = list(xml.get("xml_warnings") or [])
    if not xml["xml_ok"]:
        msg = "XML layer: %s" % xml["xml_error"]
        errors.append(msg)
        hard_errors.append(msg)
    if word is not None:
        errors.extend(word["errors"])
        warnings.extend(word["warnings"])
        if word.get("word_ok"):
            # Word ran and still reported problems: those are real failures.
            hard_errors.extend(word["errors"])
    if args.word == "on" and (word is None or word.get("skipped")):
        msg = "--word on was requested but the Word COM layer did not run"
        errors.append(msg)
        hard_errors.append(msg)
    if word is not None:
        errors.extend(word["errors"])
        warnings.extend(word["warnings"])

    # --- top-level contract values ------------------------------------------
    sources = {}
    bk_names = list(xml["bookmark_names_xml"])
    if word is not None and word.get("bookmark_names"):
        for nm in word["bookmark_names"]:
            if nm not in bk_names:
                bk_names.append(nm)

    hl_targets = ([{"sub_address": t["sub_address"], "text": t["text"],
                    "address": t.get("address", ""), "type": t.get("type"),
                    "superscript": t.get("superscript"),
                    "superscript_on": t.get("superscript_on")}
                   for t in word["hyperlink_targets"]]
                  if word is not None and word.get("hyperlink_targets")
                  else [{"sub_address": a["anchor"], "text": a["text"],
                         "address": "", "type": 0, "superscript": None,
                         "superscript_on": None}
                        for a in xml["hyperlink_anchors"]])

    # The project's "is the cross-reference clickable?" verdict, in one place
    # (BRIEF-COMMON §7.3): a real <w:hyperlink w:anchor> wrapper is what makes
    # Hyperlinks.Count non-zero — "\h" alone does not.
    hl_count = ((word or {}).get("hyperlinks") if (word or {}).get("hyperlinks")
                is not None else xml["hyperlinks_xml"])
    internals = sum(1 for t in hl_targets if t.get("sub_address"))
    externals = sum(1 for t in hl_targets if t.get("address"))
    xref_summary = {
        "hyperlinks": hl_count,
        "targets": len(hl_targets),
        "internal_anchors": internals,
        "external": externals,
        "superscript": sum(1 for t in hl_targets
                           if t.get("superscript_on") is True),
        "normal_size": sum(1 for t in hl_targets
                           if t.get("superscript_on") is False),
        "clickable": (bool(hl_targets)
                      and all(t.get("sub_address") or t.get("address")
                              for t in hl_targets)),
        "wrapper_count_xml": xml["anchor_hyperlinks_xml"],
        "ref_fields": sum(1 for f in ((word or {}).get("field_list")
                                      or xml["field_list_xml"])
                          if "REF" in (f.get("code") or "").upper()),
    }

    page_numbers = list(xml["page_numbers"])
    page_src = xml["page_numbers_source"]
    if word is not None and word.get("page_numbers"):
        page_numbers = list(word["page_numbers"])
        page_src = word["page_numbers_source"]

    # --- per-paragraph page numbers: Word evidence, else a labelled estimate -
    # "Never fabricate a page number" is the rule here.  With Word and a request
    # list there are real readings; without either, the caller gets a
    # ``computed`` structure that says which section a paragraph is in and what
    # alphabet that section prints with — and says, in the same object, that it
    # is not a layout measurement.
    probe_idx = list(getattr(args, "page_probe_indices", None) or [])
    paragraph_pages = []
    paragraph_pages_source = "none"
    if word is not None and word.get("paragraph_pages"):
        paragraph_pages = list(word["paragraph_pages"])
        paragraph_pages_source = word.get("paragraph_pages_source") or "word"
    elif probe_idx:
        estimate = estimated_paragraph_pages(
            xml["section_specs"], (word or {}).get("section_details"),
            probe_idx[:args.limit_list])
        paragraph_pages = [
            dict(row, source="computed") for row in estimate["rows"]]
        paragraph_pages_source = estimate["source"]
        warnings.append(
            "paragraph_pages are computed structural estimates (section start "
            "page only), not measured page numbers — computed 不得当作交付页码")
        if xml.get("table_rows"):
            warnings.append(
                "paragraph_pages indexes are XML ordinals, but this document has "
                "%d table row(s): Word counts one extra paragraph per row, so an "
                "index here is not Paragraphs(n) in Word — run with Word for "
                "measured values; paragraphs inside a table do resolve to their "
                "section in this index space, only the Word paragraph ordinal "
                "drifts" % xml["table_rows"])

    numbering_render = []
    mismatch = 0
    if word is not None and word.get("word_ok"):
        numbering_render, mismatch = _merge_numbering(
            xml["numpr_paragraphs"], word.get("list_paragraphs") or [],
            args.limit_list)
    else:
        for x in xml["numpr_paragraphs"][:args.limit_list]:
            numbering_render.append({
                "index": x["index_xml"], "index_xml": x["index_xml"],
                "index_source": "xml", "list_string": None, "list_type": None,
                "ilvl": x["ilvl"], "num_id": x["num_id"],
                "lvl_text": x["lvl_text"], "rendered": None,
                "text_match": None, "text": x["text"],
            })

    if mismatch:
        warnings.append(
            "%d of %d numPr/list pairs did not match by text; ordinal pairing "
            "may be off — check numbering_render[].text_match"
            % (mismatch, len(numbering_render)))

    def pick(key, xml_key):
        if word is not None and word.get(key) is not None:
            sources[key] = "word"
            return word[key]
        sources[key] = "xml"
        return xml[xml_key]

    # pages and section geometry are Word-only facts, never XML fallbacks
    sources["pages"] = "word" if (word or {}).get("pages") is not None else "none"
    sources["paragraph_pages"] = paragraph_pages_source
    sources["page_numbers"] = ("word" if (word or {}).get("page_numbers")
                               else (xml["page_numbers_source"] if page_numbers
                                     else "none"))

    report = {
        "tool": TOOL,
        "version": VERSION,
        "docx": str(docx),
        "ok": None,
        "openable": None,
        "skipped": skipped,
        "word_available": bool(word_available()),
        "pages": (word or {}).get("pages"),
        "paragraphs": pick("paragraphs", "paragraphs_xml"),
        "sections": pick("sections", "sections_xml"),
        "bookmarks": pick("bookmarks", "bookmarks_xml"),
        "inline_shapes": pick("inline_shapes", "inline_drawings_xml"),
        "fields": pick("fields", "fields_xml"),
        "hyperlinks": pick("hyperlinks", "hyperlinks_xml"),
        "tables": pick("tables", "tables_xml"),
        "bookmark_names": bk_names[:args.limit_list],
        "hyperlink_targets": hl_targets[:args.limit_list],
        "field_list": ((word or {}).get("field_list")
                       or xml["field_list_xml"])[:args.limit_list],
        "numbering_render": numbering_render,
        "numbering_summary": {
            "xml_numpr": len(xml["numpr_paragraphs"]),
            "word_list_paragraphs": len((word or {}).get("list_paragraphs") or []),
            "rendered": sum(1 for e in numbering_render if e.get("rendered")),
            "scan": args.numbering_scan if word is not None else "none",
            "text_mismatch": mismatch,
        },
        "page_numbers": page_numbers[:args.limit_list],
        "page_numbers_source": page_src,
        "page_number_formats": list(xml["page_number_formats"]),
        "page_number_restarts": list(xml["page_number_restarts"]),
        "section_specs": list(xml["section_specs"]),
        "paragraph_pages": paragraph_pages,
        "paragraph_pages_source": paragraph_pages_source,
        "toc_audit": dict(xml.get("toc_audit") or {}),
        "xref_summary": xref_summary,
        "pdf": (word or {}).get("pdf"),
        "pngs": (word or {}).get("pngs") or [],
        "pdf_pages": (word or {}).get("pdf_pages"),
        "pdf_page_footers": (word or {}).get("pdf_page_footers") or [],
        "section_details": (word or {}).get("section_details") or [],
        "errors": errors,
        "hard_errors": hard_errors,
        "warnings": warnings,
        "sources": sources,
        "xml": xml,
        "word": ({"word_ok": word["word_ok"], "skipped": word["skipped"],
                  "skipped_reason": word["skipped_reason"],
                  "word_version": word.get("word_version"),
                  "bookmarks_total": word.get("bookmarks_total"),
                  "floating_shapes": word.get("floating_shapes"),
                  "bookmark_names_all": word.get("bookmark_names_all"),
                  "omaths": word.get("omaths"),
                  "superscript_audit": word.get("superscript_audit"),
                  "list_paragraphs": word.get("list_paragraphs")}
                 if word is not None else None),
    }

    # --- openability --------------------------------------------------------
    if word is not None and word.get("word_ok"):
        report["openable"] = True
    elif not xml["xml_ok"]:
        report["openable"] = False
    else:
        report["openable"] = None     # Word never ran: unknown, not "no"

    # --- expectations -------------------------------------------------------
    expectations = []
    hard_fail = False
    seen_attrs = set()
    for name, attr in sorted(EXPECT_ALIASES.items()):
        if attr in seen_attrs:          # --expect-params / --expect-paragraphs
            continue
        want = getattr(args, "expect_" + name.replace("-", "_"), None)
        if want is None:
            continue
        seen_attrs.add(attr)
        actual = report.get(attr)
        passed = (actual == want)
        source = sources.get(attr, "none")
        expectations.append({"name": attr, "expected": want,
                             "actual": actual, "pass": passed,
                             "source": source})
        if passed:
            continue
        if actual is None:
            # An assertion we cannot measure is not a failure.  Say so, so the
            # caller knows to re-run with Word rather than to fix the document.
            expectations[-1]["note"] = "unverifiable without Word"
            warnings.append("expect-%s=%s not checkable without Word"
                            % (name, want))
        else:
            hard_fail = True
            msg = ("expect-%s: wanted %s, got %s (from %s)"
                   % (name, want, actual, source))
            errors.append(msg)
            hard_errors.append(msg)     # keeps hard_errors == the reason ok is false
    if args.expect_openable is not None:
        want = (args.expect_openable == "yes")
        actual = report["openable"]
        passed = (actual == want)
        expectations.append({"name": "openable", "expected": want,
                             "actual": actual, "pass": passed,
                             "source": "word" if actual is not None else "none"})
        if not passed:
            if actual is None:
                warnings.append("expect-openable not checkable without Word")
            else:
                hard_fail = True
                msg = "expect-openable: wanted %s, got %s" % (want, actual)
                errors.append(msg)
                hard_errors.append(msg)
    report["expectations"] = expectations

    # --- ok -----------------------------------------------------------------
    # ``errors`` includes the reason the Word layer was skipped; that is
    # information, not a failure.  Only ``hard_errors`` and a violated
    # expectation can turn ok into false, and a run whose only gap is "Word was
    # not available" stays null so the caller decides.
    if hard_errors or hard_fail:
        report["ok"] = False
    elif report["openable"] is True:
        report["ok"] = True
    else:
        report["ok"] = None           # skipped/unverifiable: caller decides
    return report


def render_summary(report: dict) -> str:
    """Human-readable stderr summary.  Never printed to stdout."""
    lines = []
    lines.append("[word_accept] %s" % report["docx"])
    lines.append("  ok=%s openable=%s skipped=%s word_available=%s"
                 % (report["ok"], report["openable"], report["skipped"],
                    report["word_available"]))
    pairs = [("pages", "pages"), ("paragraphs", "paragraphs"),
             ("sections", "sections"), ("bookmarks", "bookmarks"),
             ("inline_shapes", "inline_shapes"), ("fields", "fields"),
             ("hyperlinks", "hyperlinks"), ("tables", "tables")]
    for key, _ in pairs:
        lines.append("  %-14s %-6s (%s)"
                     % (key, report.get(key), report["sources"].get(key, "n/a")))
    ns = report["numbering_summary"]
    lines.append("  numbering      xml=%d word=%d rendered=%d mismatched=%d"
                 % (ns["xml_numpr"], ns["word_list_paragraphs"],
                    ns["rendered"], ns["text_mismatch"]))
    rendered = [e for e in report["numbering_render"] if e.get("list_string")]
    for e in rendered[:8]:
        lines.append("    #%s %-8s %s"
                     % (e.get("index"), e.get("list_string"),
                        (e.get("text") or "")[:28]))
    if report["hyperlink_targets"]:
        for t in report["hyperlink_targets"][:8]:
            lines.append("    link -> %-16s %s"
                         % (t.get("sub_address"), (t.get("text") or "")[:24]))
    xs = report["xref_summary"]
    lines.append("  xref           hyperlinks=%d anchors=%d ext=%d "
                 "superscript=%d normal=%d clickable=%s"
                 % (xs["hyperlinks"], xs["internal_anchors"], xs["external"],
                    xs["superscript"], xs["normal_size"], xs["clickable"]))
    lines.append("  page_numbers   %s (%s)"
                 % (report["page_numbers"][:12], report["page_numbers_source"]))
    fmts = report.get("page_number_formats") or []
    restarts = report.get("page_number_restarts") or []
    if fmts or restarts:
        lines.append("  page format    %s%s"
                     % (",".join(fmts) or "(inherit)",
                        ("  restarts at %s" % (
                            [(r["section"], r["start"]) for r in restarts],))
                        if restarts else ""))
    pp = report.get("paragraph_pages") or []
    if pp:
        lines.append("  para pages     %d rows (%s)"
                     % (len(pp), report.get("paragraph_pages_source")))
        for row in pp[:8]:
            lines.append("    para %-4s sec=%-3s physical=%-4s shown=%s"
                         % (row.get("index"), row.get("section"),
                            row.get("physical"), row.get("displayed_text")))
    ta = report.get("toc_audit") or {}
    if ta.get("entries"):
        lines.append("  toc_audit      entries=%s levels=%s tokens=%d "
                     "monotonic=%s source=%s"
                     % (ta.get("entries"), ta.get("levels"),
                        len(ta.get("page_tokens") or []),
                        ta.get("monotonic"), ta.get("source")))
    su = report["xml"].get("style_usage") or []
    if su:
        lines.append("  styles         %d kinds: %s"
                     % (len(su), ", ".join(
                         "%s=%d" % (s["name"], s["count"]) for s in su[:6])))
    if report["section_details"]:
        for s in report["section_details"][:4]:
            lines.append("  section %s     %sx%s mm  L%s R%s T%s B%s"
                         % (s.get("index"), s.get("page_width_mm"),
                            s.get("page_height_mm"), s.get("left_margin_mm"),
                            s.get("right_margin_mm"), s.get("top_margin_mm"),
                            s.get("bottom_margin_mm")))
    if report["pdf"]:
        lines.append("  pdf            %s (%s pages, %d png)"
                     % (report["pdf"], report["pdf_pages"], len(report["pngs"])))
    for e in report["errors"]:
        lines.append("  ERROR   %s" % e)
    for w in report["warnings"]:
        lines.append("  warn    %s" % w)
    return "\n".join(lines)


# ===========================================================================
# CLI
# ===========================================================================
def _page_probe_list(text):
    """``"1,9,43"`` -> ``[1, 9, 43]``; blank -> ``None``.

    Duplicates are dropped (a caller assembling the list from several sources
    tends to repeat), order is preserved (document order is what makes the
    output readable), and anything non-numeric is a usage error rather than a
    silently skipped entry.
    """
    seen = set()
    out = []
    for part in str(text).replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            n = int(part)
        except ValueError:
            raise argparse.ArgumentTypeError(
                "expected comma-separated paragraph indices, got %r" % part)
        if n < 1:
            raise argparse.ArgumentTypeError(
                "paragraph indices are 1-based, got %d" % n)
        if n in seen:
            continue
        seen.add(n)
        out.append(n)
    return out or None


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="word_accept.py",
        description="Unified .docx acceptance probe: XML structure + Word COM "
                    "evidence in one JSON object on stdout.")
    p.add_argument("docx", help="path to the .docx under test")
    p.add_argument("--pdf", metavar="OUT.pdf",
                   help="export a PDF via Word (ExportAsFixedFormat 17)")
    p.add_argument("--png", metavar="DIR",
                   help="render every PDF page to DIR/page-NNN.png (needs --pdf)")
    p.add_argument("--json", metavar="OUT.json",
                   help="also write the report to this file (UTF-8, no BOM)")
    p.add_argument("--word", choices=("auto", "on", "off"), default="auto",
                   help="auto (default): use Word when present; on: require it; "
                        "off: XML layer only")
    p.add_argument("--dpi", type=int, default=110, help="PNG resolution")
    p.add_argument("--numbering-scan", choices=("numpr", "all"), default="numpr",
                   help="numpr (default) uses Document.ListParagraphs; all "
                        "scans every paragraph (slow)")
    p.add_argument("--no-update-fields", action="store_true",
                   help="skip Fields.Update() (leaves REF results stale)")
    p.add_argument("--limit-list", type=int, default=DEFAULT_LIMIT,
                   help="cap for every list-valued output key")
    p.add_argument("--page-probe", dest="page_probe_indices", default=None,
                   type=_page_probe_list, metavar="I[,J,...]",
                   help="Word paragraph indices (1-based) to read page numbers "
                        "for; comma separated, e.g. 1,9,43.  Each one costs two "
                        "COM round trips, so this is opt-in.  Without Word the "
                        "report carries a 'computed' structural estimate that "
                        "must not be delivered as a page number")
    p.add_argument("--quiet", action="store_true", help="suppress the stderr summary")

    p.add_argument("--expect-sections", type=int, default=None,
                   metavar="N", help="assert sections == N")
    p.add_argument("--expect-bookmarks", type=int, default=None,
                   metavar="N", help="assert bookmarks == N")
    p.add_argument("--expect-params", type=int, default=None,
                   metavar="N", help="assert paragraphs == N (segment count)")
    p.add_argument("--expect-paragraphs", type=int, default=None,
                   metavar="N", help="alias of --expect-params")
    p.add_argument("--expect-pages", type=int, default=None, metavar="N")
    p.add_argument("--expect-tables", type=int, default=None, metavar="N")
    p.add_argument("--expect-fields", type=int, default=None, metavar="N")
    p.add_argument("--expect-hyperlinks", type=int, default=None, metavar="N")
    p.add_argument("--expect-inline-shapes", type=int, default=None, metavar="N")
    p.add_argument("--expect-openable", choices=("yes", "no"), default=None)
    return p


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        report = build_report(args)
    finally:
        release_app()

    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.json:
        out = Path(args.json)
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(text + "\n", encoding="utf-8")
    sys.stdout.write(text + "\n")
    sys.stdout.flush()
    if not args.quiet:
        sys.stderr.write(render_summary(report) + "\n")
    if report["ok"] is False:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
