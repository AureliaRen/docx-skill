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

Contract
--------
::

    python tools/word_accept.py <in.docx> [--pdf OUT.pdf] [--png DIR]
        [--json OUT.json] [--expect-sections N] [--expect-bookmarks N]
        [--expect-params N]

    --expect-params N        asserted ``paragraphs`` (alias of
                             ``--expect-paragraphs``)
    --word auto|on|off       default ``auto`` = use Word when it is there

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
2. ``XML paragraph index != Word paragraph index``: Word counts one extra
   paragraph per table row (the row-end mark).  Measured on ``f5.docx``: 261
   ``w:p`` in XML, 287 ``Word.Paragraphs``, 26 ``w:tr``.  So never translate an
   XML index into ``Paragraphs(i)`` by arithmetic; carry both and label them.
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

    A standalone copy of ``scripts/encoding_fix.py``: this tool must keep
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


_TOC_NUM_RE = re.compile(r"[\t. ]([0-9]{1,4}|[ivxlcdm]{1,6})\s*$")
_TOC_STYLE_RE = re.compile(r"^(toc|目录|contents)", re.IGNORECASE)
# A caption paragraph style, in the many names it is spelled with: Word's
# built-in "caption", the demo template's "图表名", and the Chinese "题注".
_CAPTION_STYLE_RE = re.compile(r"(caption|图表名|题注|图名|表名)", re.IGNORECASE)


def _toc_token(text: str):
    """Trailing page number of a table-of-contents entry, or ``None``."""
    t = (text or "").replace("\u00a0", " ").rstrip()
    if not t:
        return None
    m = _TOC_NUM_RE.search(t)
    return m.group(1) if m else None


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
               limit: int = DEFAULT_LIMIT, vertalign_claims=None) -> dict:
    """Word COM probe.  Never raises: every failure becomes ``skipped``/``errors``."""
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
                          vertalign_claims=xml.get("_vertalign_claims"))
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
