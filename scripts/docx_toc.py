#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Build a *static* table of contents: real paragraphs, real ``<w:tab/>``.

A TOC field (``{ TOC \\o "1-3" }``) only becomes text when Word or LibreOffice
recalculates it.  Anything that reads the .docx without an Office engine —
``docx_read.py --text``, a text extractor, a diff, a reviewer opening it in
something else — sees the placeholder sentence instead of the contents.  So a
document that has to *read* correctly everywhere needs the TOC written out as
ordinary paragraphs:

    第一章　静电场<w:tab/>1

with the page number flushed right by a right-aligned dot-leader tab stop, and
the lines in Word's own built-in ``toc 1`` / ``toc 2`` / ``toc 3`` styles so
"Update field" still recognises them later.

Public API (frozen — other modules call exactly this):

    TOC_STYLE_NAME(level) -> "toc 1"
    collect_outline(doc)  -> [{"level", "text", "para_index"}, ...]
    build_static_toc(doc, *, levels, pages, title, insert_at,
                     indent_chars, tab_pos_mm) -> dict

Heading detection is *not* reimplemented here: ``docx_read.heading_level()``
is the single judge of what a heading is (direct ``w:outlineLvl``, style-chain
``w:outlineLvl``, ``Heading N``, and the CJK style-name table).  A style that
T1 teaches ``heading_level`` about therefore shows up here for free.

Rebuild is idempotent: paragraphs already carrying a ``toc N`` style (plus a
title paragraph this module bookmarked) are removed first, so regenerating
after editing the body never accumulates a second TOC.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import re

from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.shared import Twips

import docx_create
import docx_crossref
import docx_read

# Input validation is **not** reimplemented here.  ``docx_read`` owns the single
# definition (``InputError`` / ``_check_input`` / ``_open_document``) so the
# CLIs cannot drift apart -- it already knows the fourth failure (a ZIP with no
# ``[Content_Types].xml``, which ``is_zipfile`` happily accepts and ``Document``
# then dies on with a bare ``KeyError``).  Same shape as ``docx_edit``, which
# imports the very same three names; both halves live in ``scripts/``, which
# ``_bootstrap`` already put on ``sys.path``.
from docx_read import InputError, _check_input, _open_document

# Word's built-in TOC line styles are ``toc 1`` … ``toc 9``; their style ids are
# ``TOC1`` … ``TOC9``.  python-docx knows neither (the default template keeps
# them latent), so the style is created here and the id is pinned to what Word
# expects — otherwise a round-trip through Word turns it into a *custom* style
# and "Update field" no longer touches the lines.
TOC_STYLE_NAME_PREFIX = "toc "
TOC_TITLE_STYLE = "TOC Heading"
TOC_TITLE_BOOKMARK = "_DocxTocTitle"
DEFAULT_LEVELS = (1, 2, 3)
DEFAULT_INDENT_CHARS = (0, 2, 4)

# What the tab stop falls back to when nothing better is known: A4 with 20 mm
# side margins, i.e. 170 mm, is a maths paper; everything else in the house
# style is 160 mm (21 cm text block is rare).  160 matches the historical
# hand-written smoke spec, which is why it is the default and not 170.
DEFAULT_TAB_POS_MM = 160.0

# A paragraph styled ``toc 1`` is this module's own output, never a heading.
_SELF_STYLE_RE = re.compile(r"^toc[\s_-]*(\d+)$", re.IGNORECASE)


# ---- style names -------------------------------------------------------------

def TOC_STYLE_NAME(level) -> str:
    """``1`` -> ``"toc 1"`` (Word's built-in entry style for outline level 1)."""
    try:
        number = int(level)
    except (TypeError, ValueError):
        raise ValueError("TOC level must be an int 1-9, got %r" % (level,))
    if not 1 <= number <= 9:
        raise ValueError("TOC level must be 1-9, got %r" % (level,))
    return TOC_STYLE_NAME_PREFIX + str(number)


def _level_of_style_name(name) -> "int | None":
    match = _SELF_STYLE_RE.match(str(name or "").strip())
    if not match:
        return None
    return int(match.group(1))


def is_toc_paragraph(para) -> bool:
    """True when *para* is one of the entry lines this module writes."""
    try:
        name = para.style.name
    except (AttributeError, KeyError, ValueError):
        name = ""
    return _level_of_style_name(name) is not None


def is_toc_title(para) -> bool:
    """True for a TOC heading line — the built-in style or our bookmarked one."""
    try:
        name = (para.style.name or "").strip().lower()
    except (AttributeError, KeyError, ValueError):
        name = ""
    if name == TOC_TITLE_STYLE.lower():
        return True
    return TOC_TITLE_BOOKMARK in _bookmark_names(para)


def _bookmark_names(para) -> set:
    names = set()
    for start in para._p.iter(qn("w:bookmarkStart")):
        name = start.get(qn("w:name"))
        if name:
            names.add(name)
    return names


# ---- outline -----------------------------------------------------------------

def collect_outline(doc) -> list:
    """Every heading in *doc*, in document order.

    Each entry is ``{"level": int, "text": str, "para_index": int}`` where
    ``para_index`` indexes ``doc.paragraphs``.  The judgement of "is this a
    heading" is ``docx_read.heading_level()`` and nothing else.

    This module's own output is skipped: a TOC line styled ``toc 1`` sits at
    outline level 1 in no sense, but a template that gave the style an
    ``outlineLvl`` would otherwise make the TOC list itself.
    """
    outline = []
    for index, para in enumerate(doc.paragraphs):
        if is_toc_paragraph(para) or is_toc_title(para):
            continue
        level = docx_read.heading_level(para)
        if not level:
            continue
        outline.append({"level": int(level), "text": para.text,
                        "para_index": index})
    return outline


# ---- style plumbing ----------------------------------------------------------

def ensure_toc_style(doc, level):
    """Create (once) Word's built-in ``toc N`` paragraph style; return its name.

    The style is deliberately bare — based on ``Normal``, no font, no size — so
    a template that redefines body text also redefines the TOC lines.  Only the
    style *id* is pinned, because that is what makes Word treat it as the
    built-in rather than one more custom style.
    """
    name = TOC_STYLE_NAME(level)
    try:
        doc.styles[name]
        return name
    except KeyError:
        pass
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    try:
        style.base_style = doc.styles["Normal"]
    except KeyError:
        pass
    element = style.element
    element.set(qn("w:styleId"), "TOC%d" % int(level))
    custom = qn("w:customStyle")
    if element.get(custom) is not None:
        del element.attrib[custom]
    # Word shows the TOC styles in the "Update field" family; uiPriority keeps
    # them out of the way in the style gallery rather than at the top.
    return name


def _assign_style(para, name) -> str:
    """Assign a style by name, falling back to the paragraph's current one."""
    try:
        para.style = name
        return para.style.name
    except (KeyError, ValueError, AttributeError):
        return para.style.name if para.style is not None else ""


# ---- paragraph bits ----------------------------------------------------------

def _write_char_indent(para, chars, char_pt=None) -> "float | None":
    """Left indent expressed in *characters*, the way Word stores a TOC indent.

    ``w:ind`` keeps two twins: ``w:leftChars`` (what the ruler shows, and what
    survives a font-size change) and ``w:left`` in twips (what everyone else
    reads).  Writing both consistently is what stops Word from "repairing" the
    indent on the next save; ``docx_create`` does exactly this for
    ``firstLineChars``.
    """
    if chars is None:
        return None
    chars = float(chars)
    point_size = float(char_pt) if char_pt is not None else \
        float(docx_create.DEFAULT_CHAR_PT)
    para.paragraph_format.left_indent = Twips(
        int(round(chars * point_size * 20)))
    ind = para._p.get_or_add_pPr().find(qn("w:ind"))
    if ind is not None:
        ind.set(qn("w:leftChars"), str(int(round(chars * 100))))
    return chars


def _indent_for(level, indent_chars):
    """``indent_chars[level - 1]``, tolerant of short tuples and ``None``."""
    if not indent_chars:
        return 0
    index = int(level) - 1
    if index < 0:
        index = 0
    if index >= len(indent_chars):
        return indent_chars[-1]
    value = indent_chars[index]
    return 0 if value is None else value


def _page_text(pages, index) -> str:
    """The literal page number for entry *index*; ``""`` when unresolved.

    ``pages=None`` is the normal case: the number is filled in later by a
    paginator that has actually laid the document out, and Word would otherwise
    show a number that is right only for the first save.  A ``pages`` sequence
    is honoured verbatim — ``"i"``, ``"ii"``, ``"A-1"``, anything; the point of
    a static TOC is that the caller's string is what lands in the file.
    """
    if pages is None:
        return ""
    try:
        value = pages[index]
    except (IndexError, KeyError, TypeError):
        return ""
    if value is None:
        return ""
    return str(value)


# ---- tab stop position -------------------------------------------------------

def resolve_tab_pos_mm(doc, section, tab_pos_mm=None) -> "tuple[float, str]":
    """Right tab position in mm, plus which source produced it.

    Priority, highest first:

    1. the explicit ``tab_pos_mm`` argument — the caller measured the template;
    2. ``docx_create.content_width_mm(section)`` — the real text-block width,
       which is the only value that is right for a non-A4 page or asymmetric
       margins.  That helper is owned by T1, so it may not exist yet: an import
       that lacks it must not break TOC generation;
    3. ``DEFAULT_TAB_POS_MM`` (160 mm).

    Returns ``(mm, source)`` where source is ``"tab_pos_mm"``,
    ``"content_width_mm"`` or ``"fallback"``.
    """
    if tab_pos_mm is not None:
        return float(tab_pos_mm), "tab_pos_mm"

    resolver = getattr(docx_create, "content_width_mm", None)
    if callable(resolver):
        # The helper's signature is T1's call, not ours: it may take the
        # section, the document, or neither, and it may still be a stub that
        # raises.  Try each shape and move on when one does not fit — an
        # exception from Any call is only ever a reason to try the next one,
        # never to abandon the search or to fail the build.
        candidates = [((section,), {})] if section is not None else []
        candidates.append(((doc,), {}))
        for args, kwargs in candidates:
            try:
                value = resolver(*args, **kwargs)
            except Exception:                 # noqa: BLE001 - try the next shape
                continue
            if value is None:
                continue
            try:
                mm = float(value)
            except (TypeError, ValueError):
                continue
            if mm > 0:
                return mm, "content_width_mm"
        # Last shape: a zero-argument helper.
        try:
            value = resolver()
        except Exception:                     # noqa: BLE001 - nothing left to try
            return float(DEFAULT_TAB_POS_MM), "fallback"
        try:
            mm = float(value)
        except (TypeError, ValueError):
            mm = 0.0
        if mm > 0:
            return mm, "content_width_mm"
    return float(DEFAULT_TAB_POS_MM), "fallback"


# ---- removal (idempotence) ---------------------------------------------------
#
# A style name is **not** proof of ownership.  ``toc 1`` is a built-in Word
# style, so a document that arrives with a table of contents Word already
# generated looks exactly like one of ours -- deleting it silently would wipe
# the user's own entries with nothing to show for it (measured:
# work/REVIEW-ADVERSARIAL-R6.md 阻塞-4).  So every paragraph this module writes
# carries its own marker, and only marked paragraphs are ever removed.
#
# The marker is a bookmark on the paragraph (``_DocxTocEntry0001``) rather than
# ``w:vanish`` or a custom style, for two reasons: a bookmark is invisible in
# print and in the body text a reader edits, and it survives a LibreOffice
# round trip (T3 §3.4 measured ``_Toc…`` names doing exactly that).
ENTRY_MARK_PREFIX = "_DocxTocEntry"
TITLE_MARK_PREFIX = "_DocxTocTitleTag"


def _next_entry_mark(doc) -> str:
    """A fresh ``_DocxTocEntryNNNN`` name, unique across *doc*.

    Delegates to ``docx_crossref.next_toc_name`` so the reserved-set logic (two
    minted names can never collide even before either is written) is shared
    rather than reimplemented.  The name satisfies all three bookmark rules T3
    established: unique, no spaces, well under 40 characters.
    """
    return docx_crossref.next_toc_name(doc, prefix=ENTRY_MARK_PREFIX)


def _mark(para, name: str) -> None:
    docx_create._bookmark_para(para, name)


def is_our_toc_paragraph(para) -> bool:
    """True when *para* was written by :func:`build_static_toc`.

    Both halves must hold: the TOC style **and** one of our own markers.  A
    paragraph a user brought along with the right style has no marker and is
    left alone.
    """
    if not is_toc_paragraph(para):
        return False
    return any(str(n).startswith(ENTRY_MARK_PREFIX)
               for n in _bookmark_names(para))


def is_our_toc_title(para) -> bool:
    """True when *para* is a title paragraph this module wrote.

    Ownership needs a marker, not just a name.  ``TOC_TITLE_BOOKMARK`` is a
    fixed constant and ``is_toc_title`` accepts any paragraph carrying it, so a
    document that happens to hold a paragraph with that bookmark — hand-edited,
    or written by another tool — would be deleted as if it were ours
    (measured: work/REVIEW-ADVERSARIAL-R6.md 阻塞-②).

    The marker is the **run tag** this module stamps on every title it writes:
    a ``_DocxTocTitle<NNNN>`` bookmark minted from the same counter the entries
    use, so it is unique per document and cannot be guessed by a user.  The
    fixed ``_DocxTocTitle`` bookmark is still written for compatibility with
    documents produced before this tag existed; it is no longer sufficient on
    its own.
    """
    names = _bookmark_names(para)
    if not any(str(n).startswith(TITLE_MARK_PREFIX) for n in names):
        return False
    return is_toc_title(para)


def drop_existing_toc(doc, *, adopt_unmarked: bool = False) -> dict:
    """Remove this module's previous output; report what went and where.

    Only paragraphs carrying this module's marker are removed (see the comment
    above ``ENTRY_MARK_PREFIX``).  ``adopt_unmarked=True`` additionally treats
    an unmarked TOC-styled paragraph as ours -- the escape hatch for rebuilding
    a document whose original TOC was generated by an older version of this
    module, or by Word itself and which the caller deliberately wants replaced.
    It is off by default because deleting a user's entries is unrecoverable.

    Returns ``{"entries": n, "titles": n, "adopted": n,
    "anchor": <w:p element or None>}``.  ``anchor`` is the paragraph that
    preceded the first removed line, so a rebuild lands where the old TOC was
    instead of at the end of the document.
    """
    dropped_entries = 0
    dropped_titles = 0
    adopted = 0
    anchor = None
    previous = None
    removed_any = False
    for para in list(doc.paragraphs):
        ours_entry = is_our_toc_paragraph(para)
        ours_title = is_our_toc_title(para)
        take_entry = ours_entry or (adopt_unmarked and is_toc_paragraph(para))
        take_title = ours_title
        if not (take_entry or take_title):
            previous = para._p
            continue
        if not removed_any:
            anchor = previous
            removed_any = True
        if take_entry:
            dropped_entries += 1
            if not ours_entry:
                adopted += 1
        if take_title:
            dropped_titles += 1
        element = para._p
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)
    return {"entries": dropped_entries, "titles": dropped_titles,
            "adopted": adopted, "anchor": anchor}


# ---- building ----------------------------------------------------------------

class InsertIndexError(ValueError):
    """``--insert-at`` cannot be honoured honestly.

    Raised *before* the document is touched, so ``main`` can refuse with exit 2
    and nothing is written -- the file is exactly as it was.  Mirrors
    ``docx_edit.InsertIndexError``: "put the TOC before paragraph i" and "put the
    paragraph before index i" need the same answer for the same question.
    """


def _check_insert_index(insert_at, n) -> None:
    """Refuse an ``--insert-at`` that does not name a real insertion point.

    ``insert_at`` is a 0-based position in ``doc.paragraphs`` (body only -- table
    cell paragraphs are not counted), and must name a paragraph that exists: the
    TOC is inserted *before* it, so it lands at exactly that index.  Anything
    else is a caller bug, and clamping it silently is worse than failing, because
    the caller is told a position it will never get:

    * ``--insert-at 999`` on a 3-paragraph body used to fall through to an append
      and report ``inserted_at: null`` with exit 0, having written the file;
    * ``--insert-at -1`` used to hit Python's negative indexing, insert before the
      *last* paragraph (landing at N-1) and echo the argument back.

    An empty body (N=0) is the one degenerate case with a single valid insertion
    point: 0, which appends the TOC -- there is no paragraph to go before.
    """
    index = int(insert_at)
    if n == 0:
        if index != 0:
            raise InsertIndexError(
                f"--insert-at {index} is out of range: the body has no "
                f"paragraphs, so the only valid index is [0, 0] (current N=0)")
        return
    if not 0 <= index <= n - 1:
        raise InsertIndexError(
            f"--insert-at {index} is out of range: legal range is "
            f"[0, {n - 1}], current N={n}")


def _check_insert_index_shrunk(insert_at, n) -> None:
    """Second check, against the body left *after* the previous TOC was removed.

    :func:`_check_insert_index` judges ``--insert-at`` against the body the
    caller named, which is the only honest reading of the argument.  But
    rebuilding a TOC first removes the old one, and that removal shortens the
    body.  An index that was legal a moment ago can now point past the end.

    Falling back to an append at that point is *silent misplacement*, not a
    convenience: the TOC ends up somewhere the caller never asked for, in a
    document that exits 0.  So the request is refused instead, and the message
    says which of the two checks failed -- otherwise the caller would see an
    index that looked legal against the document on disk and be told it is out
    of range, with no way to reconcile the two.
    """
    index = int(insert_at)
    if n == 0:
        # Same degenerate case the first check allows through: with no
        # paragraphs left there is no "before which" to honour, and 0 is the
        # single index that means "append".  Refusing it would break rebuilding
        # a TOC-only document -- a document whose *only* paragraph was that TOC
        # is back to N=0 once it is dropped.
        if index != 0:
            raise InsertIndexError(
                f"--insert-at {index} is out of range: once the previous TOC "
                f"was removed the body had no paragraphs left, so the only "
                f"valid index is [0, 0] (current N=0)")
        return
    if index < n:
        return
    raise InsertIndexError(
        f"--insert-at {index} is out of range after the previous TOC was "
        f"removed: legal range is now [0, {n - 1}], current N={n} (the index "
        f"was legal against the body before removal; pass an index within the "
        f"shortened body, or omit --insert-at to land where the old TOC was)")


def _target_section(doc, section_index=None, insert_at=None):
    """Which section's geometry the tab stop should follow.

    Returns ``(section, index)``.  Resolution order: an explicit
    ``section_index``; else the section that owns the paragraph at
    ``insert_at``; else the first section.  A tab stop taken from the wrong
    section misplaces every dotted leader by the margin difference (measured:
    165.89 mm vs 175.90 mm on a two-section sample, i.e. 10 mm of drift -- see
    work/REVIEW-ADVERSARIAL-R6.md 应修-6).

    A section's ``sectPr`` is attached to the **last paragraph of that section**
    and only the final section's hangs off ``<w:body>``.  So the section owning
    paragraph *n* is "how many section-closing paragraphs come **strictly
    before** *n*": a paragraph that carries ``w:pPr/w:sectPr`` is the last one of
    its own section, not the first of the next, so it must not be counted for
    its own index -- the TOC inserted *before* it lands inside that same section.
    """
    sections = list(doc.sections)
    if not sections:
        return None, None

    if section_index is not None:
        # An index that names no section cannot be answered honestly.  Falling
        # back to section 0 used to hand back a width nobody asked for and
        # misplaced every dotted leader by the margin difference (measured: 40 mm
        # on a three-section sample), in a document that exited 0 -- silent
        # misplacement, so the request is refused instead.  Negative indices are
        # the usual "count from the end" spelling and stay legal down to -N,
        # which is the behaviour every other index in this skill already has.
        try:
            want = int(section_index)
        except (TypeError, ValueError):
            raise InsertIndexError(
                f"--section-index {section_index!r} is not an integer: legal "
                f"range is [0, {len(sections) - 1}] or "
                f"[-{len(sections)}, -1], current N={len(sections)}") from None
        if not -len(sections) <= want < len(sections):
            raise InsertIndexError(
                f"--section-index {want} is out of range: legal range is "
                f"[0, {len(sections) - 1}] or [-{len(sections)}, -1], current "
                f"N={len(sections)}")
        return sections[want], want % len(sections)

    if insert_at is not None:
        paras = doc.paragraphs
        try:
            at = int(insert_at)
        except (TypeError, ValueError):
            at = 0
        # ``build_static_toc`` has already refused an ``insert_at`` that names no
        # real paragraph, so ``at`` is in range whenever this is reached with a
        # caller-supplied index.  The guard stays because this helper is called
        # with the same argument from two places, and falling back to section 0
        # is the documented behaviour for "there is no paragraph to ask" -- only
        # now it can no longer happen for a *bad* index, which used to make the
        # tab stop follow the wrong section's geometry in silence.
        if 0 <= at < len(paras):
            closing = 0
            for i, para in enumerate(paras):
                if i >= at:
                    break
                pPr = para._p.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                    closing += 1
            idx = min(max(closing, 0), len(sections) - 1)
            return sections[idx], idx

    return sections[0], 0


def _title_text(title):
    """Accept ``"目录"`` or ``{"text": "目录", "style": "MyToc"}}``."""
    if isinstance(title, dict):
        text = title.get("text", "")
        return ("" if text is None else str(text)), title.get("style")
    return ("" if title is None else str(title)), None


def _add_title_paragraph(doc, title):
    text, wanted = _title_text(title)
    para = doc.add_paragraph()
    style = wanted or TOC_TITLE_STYLE
    assigned = _assign_style(para, style)
    if assigned != style:
        # "TOC Heading" is a latent style in the default template; when a
        # stripped template has dropped it too, bold Normal is the closest
        # thing that still reads as a heading.
        _assign_style(para, "Normal")
        bold_text = text
        run = para.add_run(bold_text)
        run.bold = True
    else:
        para.add_run(text)
    para.paragraph_format.keep_with_next = True
    # Two bookmarks: the constant one a TOC field convention expects, and a
    # unique tag that proves this paragraph is ours.  The tag is what
    # ``is_our_toc_title`` keys on, so a user paragraph that merely reuses the
    # constant name is never mistaken for one of ours.
    docx_create._bookmark_para(para, TOC_TITLE_BOOKMARK)
    _mark(para, docx_crossref.next_toc_name(doc, prefix=TITLE_MARK_PREFIX))
    return para


def _add_entry_paragraph(doc, style_name, text, page_text, tab_pos_mm,
                         indent_chars):
    para = doc.add_paragraph()
    _assign_style(para, style_name)
    # ``add_runs`` is the module that owns "a literal \\t means a real
    # <w:tab/>" (python-docx's add_run would put it in <w:t>, where Word draws
    # nothing at all).  Reusing it keeps that rule in one place.
    docx_create.add_runs(para, {"text": text + "\t" + page_text})
    docx_create._apply_tab_stops(
        para, [{"pos_mm": tab_pos_mm, "align": "right", "leader": "dot"}])
    if indent_chars is not None:
        _write_char_indent(para, indent_chars)
    return para


def _body_insert_reference(doc, anchor, insert_at):
    """Element the new paragraphs must be inserted *before* (``None`` = append).

    ``insert_at`` has already been checked twice by the time this runs: once
    against the body the caller *named*, and again against the body that is
    left after the previous TOC was dropped (see
    :func:`_check_insert_index_shrunk`).  So an index that survives both
    checks is honoured as-is, and the ``None`` this returns for the
    ``insert_at`` branch is no longer a "bad index means append" fallback.
    """
    if insert_at is not None:
        paragraphs = doc.paragraphs
        index = int(insert_at)
        # ``index == 0`` on an empty body is the one "append" spelling: both
        # checks let it through, and ``_place`` with ``None`` puts the new
        # paragraphs where the body ends, which is exactly what appending means.
        if index < len(paragraphs):
            return paragraphs[index]._p
        return None
    if anchor is not None:
        return anchor.getnext()
    body = doc.element.body
    for child in body:
        if child.tag != qn("w:sectPr"):
            return child
    return None


def _place(doc, elements, reference) -> None:
    """Move freshly appended paragraphs to *reference* (kept in order)."""
    body = doc.element.body
    if reference is not None:
        for element in elements:
            reference.addprevious(element)
        return
    sectPr = body.find(qn("w:sectPr"))
    for element in elements:
        if sectPr is not None:
            sectPr.addprevious(element)
        else:
            body.append(element)


def _wanted_levels(levels):
    """``(1, 2)`` -> ``{1, 2}``; ``None`` -> ``None``, meaning "no filter".

    One place decides what a ``levels`` argument means, because two callers now
    need the same answer: :func:`build_static_toc`, and the CLI's
    ``--pages`` count check — which must count exactly the entries the build
    would write, or it would reject a correct invocation.
    """
    if levels is None:
        return None
    try:
        return {int(level) for level in levels}
    except (TypeError, ValueError):
        raise ValueError("levels must be an iterable of ints, got %r"
                         % (levels,))


def select_outline(doc, levels=DEFAULT_LEVELS) -> list:
    """The outline entries :func:`build_static_toc` would write, in order.

    Same filter, same order: ``collect_outline`` minus the levels nobody asked
    for.  Exposed (not private) because counting entries before a build is a
    legitimate caller need — the CLI uses it to check ``--pages`` before it
    touches the document.
    """
    wanted = _wanted_levels(levels)
    return [entry for entry in collect_outline(doc)
            if wanted is None or entry["level"] in wanted]


def build_static_toc(doc, *, levels=DEFAULT_LEVELS, pages=None, title=None,
                     insert_at=None, indent_chars=DEFAULT_INDENT_CHARS,
                     tab_pos_mm=None, section_index=None,
                     adopt_unmarked=False) -> dict:
    """Write a static table of contents into *doc*; return a summary dict.

    * ``levels`` — outline levels to include, e.g. ``(1, 2)``.
    * ``pages`` — ``None`` writes an empty page slot for a later paginator;
      a sequence is written verbatim, in outline order.  The API tolerates a
      short or long sequence (missing values become blank slots, surplus values
      are unused); the CLI does not — ``--pages`` there must carry exactly one
      value per entry, because a count mismatch is a typo, not an intent.
    * ``title`` — ``None`` writes no heading line; a string (or
      ``{"text": …}``) writes one, in ``TOC Heading`` when available.
    * ``insert_at`` — index into ``doc.paragraphs`` (after the old TOC has been
      removed) to insert before; ``None`` rebuilds where the old one was, or
      appends.
    * ``indent_chars`` — per-level left indent in characters, ``(0, 2, 4)``.
    * ``tab_pos_mm`` — right tab stop position; by default the text-block width
      of the section the entries land in (``section_index``, else the section
      of the ``insert_at`` paragraph, else the first).  Getting this wrong
      misplaces every dotted leader, which is the whole reason it is computed.
    * ``section_index`` — which section's geometry the tab stop follows.  Only
      needed when the TOC goes into a section other than the first and the
      caller wants the width without guessing.
    * ``adopt_unmarked`` — see :func:`drop_existing_toc`.  Off by default: an
      unmarked TOC-styled paragraph may well be the user's own.

    Every entry paragraph carries a ``_DocxTocEntry…`` bookmark, so a later
    rebuild removes exactly what this function wrote (see
    :func:`drop_existing_toc`).

    Returns ``{"entries", "levels", "styles", "tab_pos_mm", "tab_pos_source",
    "dropped_entries", "dropped_titles", "adopted", "section_index",
    "inserted_at"}``.
    """
    wanted = _wanted_levels(levels)

    # Judged against the body the caller described, and judged *first*: the
    # build below appends title and entry paragraphs to ``doc``, so a check any
    # later would have to undo writes to report the same thing.  See
    # :func:`_check_insert_index`.
    if insert_at is not None:
        _check_insert_index(insert_at, len(doc.paragraphs))

    fallback = drop_existing_toc(doc, adopt_unmarked=adopt_unmarked)
    anchor = fallback["anchor"]

    # The index was checked against the body the caller *named*; dropping the
    # previous TOC just shortened that body, so check again before anything is
    # appended.  Appending an index that no longer exists is silent
    # misplacement -- measured: a 9-paragraph document whose TOC sat at the top
    # accepted `--insert-at 8`, removed the old TOC (N back to 6), then appended
    # the rebuilt one at the bottom and exited 0, i.e. the TOC silently jumped
    # from the front to the end.  The caller has to hear about that instead.
    if insert_at is not None:
        _check_insert_index_shrunk(insert_at, len(doc.paragraphs))

    outline = [entry for entry in collect_outline(doc)
               if wanted is None or entry["level"] in wanted]

    section, section_used = _target_section(doc, section_index, insert_at)
    position, source = resolve_tab_pos_mm(doc, section, tab_pos_mm)

    styles = []
    for entry in outline:
        name = ensure_toc_style(doc, entry["level"])
        if name not in styles:
            styles.append(name)

    made = []
    if title is not None:
        made.append(_add_title_paragraph(doc, title))

    for index, entry in enumerate(outline):
        style_name = TOC_STYLE_NAME(entry["level"])
        text = str(entry["text"] or "").replace("\t", " ").rstrip()
        para = _add_entry_paragraph(
            doc, style_name, text, _page_text(pages, index), position,
            _indent_for(entry["level"], indent_chars))
        _mark(para, _next_entry_mark(doc))
        made.append(para)

    reference = _body_insert_reference(doc, anchor, insert_at)
    _place(doc, [para._p for para in made], reference)

    inserted_at = None
    if made:
        first = made[0]._p
        for index, para in enumerate(doc.paragraphs):
            if para._p is first:
                inserted_at = index
                break

    return {
        "entries": len(outline),
        "levels": sorted({entry["level"] for entry in outline}),
        "styles": styles,
        "tab_pos_mm": position,
        "tab_pos_source": source,
        "dropped_entries": fallback["entries"],
        "dropped_titles": fallback["titles"],
        "adopted": fallback["adopted"],
        "section_index": section_used,
        "inserted_at": inserted_at,
    }


# ---- CLI ---------------------------------------------------------------------
#
# The API above is what other modules call, but "write a TOC into this file" is
# also a shell-sized task and deserves a shell-sized entry point: without one,
# the only way in is a one-liner that has to import the module by hand.

def _levels_arg(text):
    """"1,2,3" -> (1, 2, 3); also accepts "1-3"."""
    out = []
    for part in str(text).replace(" ", "").split(","):
        if not part:
            continue
        if "-" in part:
            lo, _, hi = part.partition("-")
            out.extend(range(int(lo), int(hi) + 1))
        else:
            out.append(int(part))
    if not out:
        raise ValueError("no levels given")
    return tuple(out)


def _pages_arg(text):
    """"1,3,7" -> ["1","3","7"], preserving anything non-numeric verbatim.

    Page numbers are *glyphs*, not integers: a ``lowerRoman`` section prints
    ``i``, and a document may mix ``1`` and ``i`` in one TOC.  So the values are
    kept as the strings the caller wrote.
    """
    return [p.strip() for p in str(text).split(",") if p.strip()]


def main(argv=None) -> int:
    import argparse
    import json
    import sys

    ap = argparse.ArgumentParser(
        description="Build a static (readable-without-Word) table of contents.")
    ap.add_argument("docx", help="input .docx")
    ap.add_argument("-o", "--output",
                    help="output .docx (default: overwrite the input)")
    ap.add_argument("--levels", default="1,2,3",
                    help="outline levels to include, e.g. 1,2 or 1-3 "
                         "(default 1,2,3)")
    ap.add_argument("--pages",
                    help="page numbers for the entries, in order, comma "
                         "separated, e.g. 1,3,7 or i,ii,1 — exactly one per "
                         "entry (a different count is a usage error, exit 2). "
                         "Omit to leave the slots empty for a later pass")
    ap.add_argument("--title", help="heading text for the TOC itself, "
                                    "e.g. 目录")
    ap.add_argument("--insert-at", type=int, default=None,
                    help="insert before this paragraph index (default: where "
                         "the previous TOC was, else the end)")
    ap.add_argument("--section-index", type=int, default=None,
                    help="which section's text width sets the right tab stop")
    ap.add_argument("--indent-chars", default="0,2,4",
                    help="per-level left indent in characters (default 0,2,4)")
    ap.add_argument("--adopt-unmarked", action="store_true",
                    help="also replace an unmarked TOC (one this module did "
                         "not write) — off by default, because that may be the "
                         "user's own")
    ap.add_argument("--dry-run", action="store_true",
                    help="report what would be written, save nothing")
    args = ap.parse_args(argv)

    levels = _levels_arg(args.levels)
    pages = _pages_arg(args.pages) if args.pages else None

    # The path is judged before the package is opened, and long before
    # ``doc.save()`` at the bottom, so a refusal writes nothing and leaves the
    # file exactly as it was.  Without this guard a mistyped path died inside
    # ``Document()``: nineteen lines of traceback on stderr and exit 1, which a
    # caller has to read as Python internals to learn that the path is wrong.
    # Exit 2 is the project-wide code for "you named something that is not there
    # / not usable as given"; exit 1 stays reserved for "the document is real but
    # has a problem" (``docx_validate``'s job).  stdout stays empty: every
    # success prints JSON, so a failure must not look like an empty-but-
    # successful JSON document.
    try:
        _check_input(args.docx)
        doc = _open_document(args.docx)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    # ``--pages`` is positional-by-ordinal: the *n*-th value lands on the
    # *n*-th entry.  A count mismatch therefore is not a style question — the
    # surplus values are dropped and the missing ones come out as blank slots,
    # both silently, and a TOC with a blank page number in the middle of a
    # delivered file is indistinguishable from one that was never paginated.
    # Counting here (before anything is written) turns that into a usage error;
    # the API stays tolerant, because a caller passing a partial list on purpose
    # is a different thing from a shell user mistyping a comma.
    if pages is not None:
        entries = select_outline(doc, levels)
        if len(pages) != len(entries):
            print("error: --pages has %d value(s) but the TOC has %d entr%s; "
                  "give exactly one page number per entry, or omit --pages to "
                  "leave the slots empty"
                  % (len(pages), len(entries), "y" if len(entries) == 1 else "ies"),
                  file=sys.stderr)
            return 2

    try:
        report = build_static_toc(
            doc,
            levels=levels,
            pages=pages,
            title=args.title,
            insert_at=args.insert_at,
            indent_chars=_levels_arg(args.indent_chars),
            section_index=args.section_index,
            adopt_unmarked=args.adopt_unmarked,
        )
    except InsertIndexError as exc:
        # Same exit code argparse uses for a bad argument value: the caller
        # passed an --insert-at that cannot be honoured.  Nothing has been saved
        # at this point, so the file on disk is byte-for-byte what it was.
        print(f"error: {exc}", file=sys.stderr)
        return 2
    if not args.dry_run:
        doc.save(args.output or args.docx)
    report["output"] = None if args.dry_run else (args.output or args.docx)
    report["dry_run"] = bool(args.dry_run)
    print(json.dumps(report, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
