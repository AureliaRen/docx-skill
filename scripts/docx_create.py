#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Create a .docx document from a JSON spec.

Usage: docx_create.py spec.json output.docx
Run with --help for the spec format summary.

Spec (JSON object):
{
  "page": {"width_mm": 210, "height_mm": 297,
           "margins_mm": {"top": 25, "bottom": 25, "left": 20, "right": 20}},
  "header": "text shown in page header",
  "footer": "text shown in page footer",
  "footer_page_numbers": true,
  "styles": [{"name": "MyStyle", "base": "Normal", "font": "Arial",
              "size_pt": 12, "bold": true, "color": "1F4E79",
              "outline_level": 2}],
  "blocks": [
    {"type": "heading", "text": "Title", "level": 1},
    {"type": "paragraph", "style": "MyStyle", "runs": [
        {"text": "plain "}, {"text": "bold", "bold": true},
        {"text": " italic", "italic": true},
        {"text": " under", "underline": true}]},
    {"type": "paragraph", "text": "shortcut: single plain run"},
    {"type": "bullet_list", "items": ["a", "b"]},
    {"type": "numbered_list", "items": ["one", "two"],
     "num_text": "[%1]", "bookmark_prefix": "ref_"},
    {"type": "table", "header": ["Col1", "Col2"],
     "rows": [["1", "2"]], "style": "Light Grid Accent 1",
     "header_bold": true},
    {"type": "image", "path": "pic.png", "width_mm": 60, "alignment": "center"},
    {"type": "page_break"},
    {"type": "toc"},
    {"type": "caption", "label": "图", "text": "系统架构"},
    {"type": "crossref", "target": "图_1", "prefix": "如图",
     "superscript": true, "hyperlink": true},
    {"type": "section_break", "start_type": "new_page",
     "page": {"width_mm": 210, "height_mm": 297},
     "header": "第一章", "footer_page_numbers": {"fmt": "lowerRoman",
     "start": 1, "template": "page"}, "different_first_page": false,
     "cols": 1}
  ]
}

Paragraph format keys — accepted by every block that produces paragraphs
(``heading``, ``paragraph``, ``bullet_list``, ``numbered_list``, ``caption``,
``crossref``, ``toc``, ``image``, ``page_break``):

  ``indent_left_mm`` / ``indent_right_mm``    left/right indent
  ``first_line_indent_chars``                 first-line indent in *characters*
                                              (Word's ``w:firstLineChars``; the
                                              twips twin is computed from the
                                              effective font size, or from
                                              ``char_width_pt``)
  ``line_spacing``                            multiple, e.g. 1.5 (WD_LINE_SPACING.MULTIPLE)
  ``space_before_pt`` / ``space_after_pt``
  ``alignment``                               left | center | right | justify
  ``keep_with_next`` / ``page_break_before``  booleans
  ``bookmark``                                wrap the paragraph in a bookmark
                                              (multi-paragraph blocks bookmark
                                              the first paragraph)

``numbered_list`` additionally takes ``bookmark_prefix`` — item *k* gets the
bookmark ``<prefix><k>``, which is what a body reference points at.

``styles[]`` entries take ``outline_level`` (1-based, ``1``..``9``) — it writes
``<w:pPr><w:outlineLvl w:val="level-1"/>`` on the *style*, which is what makes a
template's own ``卷标题`` behave like ``Heading 1`` for the navigation pane and
the TOC field.  ``outline_level: 0`` writes Word's body-text marker ``9``
(explicitly *not* collected), and omitting the key writes nothing.
``docx_read.heading_level`` reads the same number back.

``ensure_toc_styles(doc)`` defines Word's built-in ``toc 1/2/3`` (based on
``Normal``, 1.5-line spacing, per-level character indent); ``content_width_mm(
section)`` is the text-block width, which is where a right-aligned TOC tab stop
belongs.

``table`` takes ``keep_together`` (bool, default false).  Word has no
table-level "keep with next", so it is written in two layers: ``w:cantSplit``
inside every row's ``w:trPr`` (a row never breaks across pages) and
``keep_with_next`` on every paragraph of every cell except the last row (the
table holds itself together).  ``keep_together`` is a request to Word, not a
guarantee: a table longer than a page still spans pages, because ``cantSplit``
only forbids splitting *a row*.

A ``table_caption`` block followed **immediately** by a ``table`` block is
linked automatically, with no key to write: the caption paragraph gets
``keep_with_next`` and the table's first row gets ``w:cantSplit``.  Anything
between them (even an empty ``paragraph``) breaks the link, because a caption
two blocks away from its table is not what caused the split.

``crossref`` passes ``superscript`` (bool), ``vert_align``
(``normal``/``superscript``/``subscript``) and ``hyperlink`` (bool, default
true) through to ``docx_crossref.cmd_ref``.

``image`` takes either ``width_mm`` (absolute) or ``width_frac`` (a share of
the **current section's** text width, ``0 < f <= 1``; A4 with 26 mm side
margins is 158 mm, so ``width_frac: 1.0`` is a full-width figure).  Giving both
is an error rather than a silent winner.

``section_break`` closes the running section and starts a new one; keys are
``start_type``, ``page``, ``header``, ``header_align``, ``footer``,
``footer_page_numbers``, ``different_first_page``, ``cols``,
``no_footer`` (blank out an inherited footer), ``reuse_last_paragraph`` and
``apply_to``.  ``apply_to`` names which side of the break the page setup
describes: ``next`` (default) is the section that begins here, ``previous``
is the section the break closes.  ``next`` is the default because the final
section of a document is followed by no break, so it can only be described
from the break above it.  Each section's page number format and restart point
live in its own ``<w:pgNumType>``, so cover / ``i, ii`` / restart-at-``1`` is
expressible; see ``docx_section.py``.  ``footer_page_numbers`` accepts ``fmt``,
``start``, ``template`` (``page``/``page_of``/``zh``/``dash``/``custom``), ``total``,
``align``, ``prefix``, ``suffix``.  ``zh`` writes ``第 X 页　共 Y 页`` with a
full-width space (U+3000) between ``页`` and ``共``; ``custom`` interpolates
``{page}`` and ``{total}`` inside ``prefix``/``suffix``, so a non-English layout
does not need a template of its own.

Extras: `"footer_page_numbers": true` at the top level adds a
"Page X of Y" footer built from PAGE/NUMPAGES fields to the first section,
and a `toc` block inserts a Table of Contents field. Field results are
computed by Word/LibreOffice when the file is opened, not by python-docx;
every field is written with a cached result so the file reads correctly
before anyone recalculates.

Exit codes — ``0`` the document was written (JSON on stdout); ``1`` it was
built but is not deliverable (a dangling cross-reference target, an invalid
or duplicate bookmark name); ``2`` the *spec itself* was rejected **before
anything was written**.  Code ``2`` prints one readable line to stderr and
never a traceback: it names the offending key and the nearest legitimate
spelling, because a misspelt key used to be dropped in silence — the spec
then built a document with the content missing and still exited ``0``.
A genuine run-time failure of the builder keeps its exit ``1`` and its
traceback, so a bug can still be located.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import contextlib
import difflib
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING,
                            WD_TAB_ALIGNMENT, WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor, Twips

from fix_zoom import fix_docx
from docx_section import (
    add_page_number_field, add_section_break, apply_page_to_section,
    configure_section, set_hdrftr_text,
)
from docx_crossref import check_crossref_anchors
from docx_utils import next_bm_id, insert_ordered

# Keys every paragraph-producing block understands.
PARA_FORMAT_KEYS = (
    "indent_left_mm", "indent_right_mm", "first_line_indent_chars",
    "line_spacing", "space_before_pt", "space_after_pt", "alignment",
    "keep_with_next", "page_break_before", "bookmark", "tab_stops",
    "bottom_border",
)

# An exam answer line is a paragraph whose *bottom border* is the rule.  There
# is no other way to express it: an empty paragraph with no border renders as
# blank white space, which is what a student sees as "nowhere to write".
#   {"type": "paragraph", "text": "", "bottom_border": {"style": "single",
#    "size": 8, "space": 1, "color": "000000"}}
_BORDER_STYLES = {
    "single", "double", "dashed", "dotted", "thick", "wave", "none",
}
_BORDER_SIDES = ("top", "left", "bottom", "right")

# A Chinese paper's table-of-contents is prose + a tab + a page number, with dot
# leaders and the number flush right.  That is a tab stop, not a paragraph
# property, so it needs its own vocabulary:
#   {"pos_mm": 155, "align": "right", "leader": "dot"}
_TAB_ALIGNS = {
    "left": WD_TAB_ALIGNMENT.LEFT,
    "center": WD_TAB_ALIGNMENT.CENTER,
    "right": WD_TAB_ALIGNMENT.RIGHT,
    "decimal": WD_TAB_ALIGNMENT.DECIMAL,
}
_TAB_LEADERS = {
    "none": WD_TAB_LEADER.SPACES,
    "space": WD_TAB_LEADER.SPACES,
    "dot": WD_TAB_LEADER.DOTS,
    "dots": WD_TAB_LEADER.DOTS,
    "dash": WD_TAB_LEADER.DASHES,
    "dashes": WD_TAB_LEADER.DASHES,
    "line": WD_TAB_LEADER.LINES,
    "lines": WD_TAB_LEADER.LINES,
    "heavy": WD_TAB_LEADER.HEAVY,
    "middle_dot": WD_TAB_LEADER.MIDDLE_DOT,
}

_ALIGNMENTS = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# Word measures a "character" of indent at the run's own point size, so a
# 2-character indent on 12pt text is 480 twips. Fall back to 12pt — the
# Chinese body-copy default — when nothing else is known.
DEFAULT_CHAR_PT = 12.0


def apply_page(doc, page: dict) -> None:
    """Backwards-compatible entry point: page setup for the first section."""
    apply_page_to_section(doc.sections[0], page)


def content_width_mm(section) -> float:
    """Text-block ("type area") width of *section*, in millimetres.

    A right-aligned tab stop for a TOC page number has to sit exactly on the
    right edge of the text block.  That edge is ``page width - left - right``,
    and **not** a house number: the maths volume is A4 with 20 mm margins
    (170 mm) while the other five are A4 with 30/20 (160 mm).  Hard-coding
    either value misplaces every dotted leader on the other page geometry.

    Margins are stored in twips, so a 20 mm margin reads back as 20.0015 mm;
    the result is rounded to 0.01 mm because nobody — and no layout engine —
    distinguishes 170.0036 from 170.00.
    """
    width = section.page_width
    if width is None:
        raise ValueError("section has no page width; set page_size/page width")
    total = float(width.mm)
    for margin in (section.left_margin, section.right_margin):
        if margin is not None:
            total -= float(margin.mm)
    return round(total, 2)


def _set_style_fonts(style, latin: str, east_asia: str | None) -> None:
    """Bind latin / CJK font slots in one ``w:rFonts``.

    ``style.font.name`` only writes ``ascii`` and ``hAnsi``, so CJK runs in a
    style that asks for 黑体 still fall back to the theme's East Asian font
    (Word reports ``Font.NameFarEast == '+中文正文'``).  Setting ``eastAsia``
    explicitly is what makes the heading actually render 黑体.
    """
    rpr = style.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        from docx.oxml import OxmlElement as _El
        rf = _El("w:rFonts")
        # CT_RPr is an xsd:sequence — rFonts may only sit after rStyle and
        # before everything from b onwards.
        rpr.insert_element_before(
            rf, "w:b", "w:bCs", "w:i", "w:iCs", "w:noProof", "w:snapToGrid",
            "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
            "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
            "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
            "w:eastAsianLayout", "w:specVanish", "w:oMath")
    # A built-in style (``Heading 1`` and friends) binds its fonts through the
    # *theme* attributes -- ``w:asciiTheme="majorHAnsi"`` etc.  In ``CT_Fonts``
    # a theme attribute takes precedence over its explicit twin, so writing
    # ``w:ascii="黑体"`` next to a surviving ``w:asciiTheme`` leaves Word
    # rendering the theme face: the XML would show the font the spec asked for
    # while the page shows another one.  Naming a face therefore means taking
    # the theme binding out of the way, which is part of honouring ``font``
    # rather than something done to the keys the spec did not mention.
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if rf.get(qn("w:" + attr)) is not None:
            del rf.attrib[qn("w:" + attr)]
    for attr, value in (("ascii", latin), ("hAnsi", latin),
                        ("eastAsia", east_asia or latin),
                        ("cs", latin)):
        rf.set(qn("w:" + attr), value)


# A paragraph's outline level is what Word's "Heading 1..9" styles carry and
# what the navigation pane / TOC field collect.  A template's own ``卷标题``
# would otherwise be plain text to Word: the document looks right and the TOC
# comes out empty, with nothing in the XML to show why.  ``w:outlineLvl`` is
# **0-based** (built-in Heading 1 stores ``val="0"``) while a spec author counts
# from 1, so the mapping happens once, here.
OUTLINE_BODY_VAL = "9"          # Word's own "not an outline level" marker


def _outline_lvl_val(outline_level) -> str:
    """``outline_level`` spec value → ``w:outlineLvl/@w:val``.

    ``1``..``9`` are 1-based outline levels (level 1 → ``"0"``).  ``0`` is the
    explicit *body text* level: Word writes ``9`` for "Body Text" and
    ``docx_read.heading_level`` maps 9+ back to ``None``, so ``0`` is how a
    spec says "this style must never be collected", which is different from
    omitting the key.
    """
    try:
        level = int(outline_level)
    except (TypeError, ValueError):
        raise ValueError("outline_level must be an int 0..9, got %r"
                         % (outline_level,))
    if isinstance(outline_level, float) and outline_level != level:
        raise ValueError("outline_level must be a whole number, got %r"
                         % (outline_level,))
    if isinstance(outline_level, str):
        raise ValueError("outline_level must be numeric, got %r"
                         % (outline_level,))
    if level == 0:
        return OUTLINE_BODY_VAL
    if 1 <= level <= 9:
        return str(level - 1)
    raise ValueError("outline_level out of range: %r (0 = body, 1..9 = level)"
                     % (outline_level,))


def _apply_style_outline_level(style, outline_level) -> str:
    """Write ``<w:pPr><w:outlineLvl w:val="…"/>`` on a *style*, schema-ordered.

    ``CT_PPr`` is an ``xsd:sequence``: ``outlineLvl`` sits after
    ``jc``/``textDirection``/``textAlignment``/``textboxTightWrap`` and before
    ``divId``.  Appending it at the end of ``w:pPr`` — the obvious thing to do
    — puts it after ``rPr``/``sectPr``, where **Word silently drops it**, and
    the style then looks correct in the XML while behaving as body text.  The
    successor chain below is the same one :func:`_apply_bottom_border` uses for
    the predecessor side of that sequence.
    """
    val = _outline_lvl_val(outline_level)
    pPr = style.element.get_or_add_pPr()
    existing = pPr.find(qn("w:outlineLvl"))
    if existing is not None:
        existing.set(qn("w:val"), val)
        return val
    el = OxmlElement("w:outlineLvl")
    el.set(qn("w:val"), val)
    pPr.insert_element_before(el, "w:divId", "w:cnfStyle", "w:rPr",
                              "w:sectPr", "w:pPrChange")
    return val


def _get_or_add_style(doc, name):
    """The style called *name*, adding a paragraph style only when it is absent.

    ``Styles.add_style`` raises ``ValueError`` on a name that already exists,
    which made ``styles[]`` unable to reach any built-in style at all: asking
    for ``{"name": "Heading 1", "size_pt": 16}`` was an error rather than an
    edit, so the only way to restyle a heading was a copy under a new name.
    Returns ``(style, created)``.
    """
    try:
        return doc.styles[name], False
    except KeyError:
        return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH), True


def add_styles(doc, styles: list) -> None:
    """Apply ``styles[]`` entries; an existing style is modified *in place*.

    Only the keys a spec actually gives are written.  Everything else on the
    style is left exactly as it was — not reset, not cleared, not re-based.
    That is what makes the two features compose: the heading styles are
    neutralised to black at build time (see
    :func:`_neutralise_heading_styles`), and a spec that only says
    ``{"name": "Heading 1", "size_pt": 16}`` then changes the size while the
    colour stays black.  A "reset then apply" implementation would send it back
    to the theme's blue, which is the defect this ordering exists to avoid.
    """
    for s in styles:
        style, _created = _get_or_add_style(doc, s["name"])
        # ``base`` is the one key that *is* an error when it points nowhere:
        # silently inventing a base style would make the document look almost
        # right, and python-docx's own ``KeyError`` already carries a message
        # ``_describe`` can reword.
        if s.get("base"):
            style.base_style = doc.styles[s["base"]]
        font = style.font
        if s.get("font"):
            font.name = s["font"]
            # ``font_east_asia`` lets a spec pick a different CJK face
            # (e.g. latin Times New Roman + eastAsia 宋体); default is the
            # same face for both, which is what a Chinese exam paper wants.
            _set_style_fonts(style, s["font"], s.get("font_east_asia"))
        if s.get("size_pt"):
            font.size = Pt(s["size_pt"])
        if s.get("bold") is not None:
            font.bold = s["bold"]
        if s.get("italic") is not None:
            font.italic = s["italic"]
        if s.get("color"):
            # ``ColorFormat.rgb``'s setter drops ``w:themeColor`` first, so the
            # explicit value is the one Word renders.  Assigning the element's
            # ``w:val`` directly would leave the theme attribute winning.
            font.color.rgb = RGBColor.from_string(s["color"])
        # ``0`` is meaningful (explicit body level), so this is ``is not None``
        # rather than a truthiness test.
        if s.get("outline_level") is not None:
            _apply_style_outline_level(style, s["outline_level"])


# Word ships ``toc 1``..``toc 9`` as *built-in* styles.  A document created by
# python-docx does not contain them until something asks, and Word's "Update
# field" then formats the generated lines with whatever latent definition it
# has — which on a Chinese template is not the 1.5-line, indented look the
# volume wants.  Defining them up front puts the look in the file itself.
TOC_STYLE_PREFIX = "toc "
TOC_UI_PRIORITY = 39            # same value Word uses for the built-in family
TOC_LINE_TWIPS = 360            # 1.5 lines, ``w:spacing/@w:line``
TOC_LINE_RULE = "auto"


def _toc_indent_for(level: int, indent_chars) -> float:
    """``indent_chars[level - 1]``, tolerant of a short tuple and of ``None``."""
    if not indent_chars:
        return 0.0
    index = max(int(level) - 1, 0)
    if index >= len(indent_chars):
        index = len(indent_chars) - 1
    value = indent_chars[index]
    return 0.0 if value is None else float(value)


def ensure_toc_styles(doc, levels=(1, 2, 3),
                      indent_chars=(0, 2, 4)) -> dict:
    """Create — or complete — Word's built-in ``toc 1/2/3`` paragraph styles.

    ``docx_toc`` only needs the *id* so Word recognises the family; a printed
    volume needs the whole definition: based on ``Normal`` (so the template's
    body font and size flow through), ``next = Normal`` (pressing Enter at the
    end of a TOC line starts body text, not another TOC line), Word's
    ``uiPriority`` 39 plus ``unhideWhenUsed``/``qFormat`` so the styles behave
    like the built-ins in the gallery, and a per-level left indent written as
    **both** ``w:left`` (twips) and ``w:leftChars`` (what the ruler shows and
    what survives a font-size change).

    The indents follow the reference volume: level 1 flush, level 2 = 2
    characters, level 3 = 4 characters — i.e. ``toc 2`` gets
    ``leftChars="200"`` and ``toc 3`` ``leftChars="400"``.

    Idempotent by construction: ``Styles.add_style`` raises on a name that
    already exists, so an existing style is *updated in place* instead.  That
    also means a document that already carries Word's own ``toc 1`` (a
    template, or a second call) keeps its identity — no duplicate style, no
    second ``w:style`` with the same name, which Word reports as a broken file.

    Returns a report dict (``levels``/``styles``/``style_ids``/
    ``indent_chars``/``created``/``updated``).
    """
    report = {"levels": [], "styles": [], "style_ids": [],
              "indent_chars": [], "created": [], "updated": []}
    normal = None
    try:
        normal = doc.styles["Normal"]
    except KeyError:
        pass

    for level in levels:
        level = int(level)
        if level < 1:
            raise ValueError("toc level must be >= 1, got %r" % (level,))
        name = TOC_STYLE_PREFIX + str(level)
        chars = _toc_indent_for(level, indent_chars)

        try:
            style = doc.styles[name]
            created = False
        except KeyError:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            created = True

        element = style.element
        # `toc1`/`toc2`/`toc3` and no `w:customStyle`: python-docx mints
        # ``customStyle="1"`` for every style it adds, and Word then shows the
        # TOC lines as a custom style rather than as its own.
        element.set(qn("w:styleId"), "toc%d" % level)
        custom = qn("w:customStyle")
        if element.get(custom) is not None:
            del element.attrib[custom]

        if normal is not None:
            style.base_style = normal
            style.next_paragraph_style = normal

        element.get_or_add_uiPriority().val = TOC_UI_PRIORITY
        element.get_or_add_unhideWhenUsed()
        element.get_or_add_qFormat()

        pf = style.paragraph_format
        pf.line_spacing = TOC_LINE_TWIPS / 240.0
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        # Word measures a "character" at the run's own point size; use the
        # document-wide body default so the twips twin and the Chars twin stay
        # consistent (a mismatch is what makes Word "repair" the indent).
        twips = int(round(chars * DEFAULT_CHAR_PT * 20))
        pf.left_indent = Twips(twips)
        pPr = element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            ind.set(qn("w:leftChars"), str(int(round(chars * 100))))

        report["levels"].append(level)
        report["styles"].append(name)
        report["style_ids"].append("toc%d" % level)
        report["indent_chars"].append(chars)
        (report["created"] if created else report["updated"]).append(name)
    return report


# Word's built-in ``caption`` style ships with ``<w:color w:val="4F81BD"
# w:themeColor="accent1"/>`` and bold — it is designed for a blue-accented
# Word template, not for a printed exam paper or a thesis.  Because the caption
# style is what "Insert Table of Figures" then formats, every caption came out
# bold blue (seen in the rendered sample).  A caption paragraph in this project
# is meant to be plain black, so the built-in style is neutralised once, at
# document build time, and a spec can still override it via ``styles``.
_CAPTION_STYLE_IDS = ("Caption", "caption", "图表名", "题注")
# 10.5 pt == 五号, written in half-points the way ``w:sz`` wants it.
CAPTION_HALF_POINTS = 21


def _rpr_order():
    """The ``CT_RPr`` child order, shared with the cross-reference fields."""
    from docx_crossref import RPR_ORDER
    return RPR_ORDER


def _neutralise_caption_style(doc) -> bool:
    """Strip the theme colour and bold from the built-in caption style.

    Returns True when a style was actually touched, so the caller can report
    it.  Only the *style* is changed — an explicit ``styles`` entry in the spec
    still wins, because ``add_styles`` runs afterwards.
    """
    touched = False
    for name in _CAPTION_STYLE_IDS:
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        rpr = style.element.find(qn("w:rPr"))
        if rpr is None:
            continue
        colour = rpr.find(qn("w:color"))
        if colour is not None:
            # ``themeColor``/``themeShade`` would re-apply the blue even with an
            # explicit val, so the whole element is replaced by a plain black.
            rpr.remove(colour)
            new = OxmlElement("w:color")
            new.set(qn("w:val"), "000000")
            insert_ordered(rpr, new, _rpr_order())
            touched = True
        for tag in ("w:b", "w:bCs"):
            b = rpr.find(qn(tag))
            if b is not None:
                rpr.remove(b)
                touched = True
        # 9 pt is Word's built-in caption size; a caption in these papers is
        # 五号 (10.5 pt), one step below the 12 pt body.
        for tag in ("w:sz", "w:szCs"):
            sz = rpr.find(qn(tag))
            if sz is not None and sz.get(qn("w:val")) in ("18", "20"):
                sz.set(qn("w:val"), str(CAPTION_HALF_POINTS))
                touched = True
    return touched


# ``Heading 1``..``Heading 9`` / ``Title`` / ``Subtitle`` ship in Word's own
# template as **theme-coloured**: Heading 1 is ``w:val="365F91"
# w:themeColor="accent1" w:themeShade="BF"``, Title is ``text2``, Subtitle is
# ``accent1`` and italic.  That is a design choice for a screen document and the
# wrong one for a printed Chinese assignment: every heading comes out blue, and
# the theme attribute wins over the RGB twin at render time, so an author cannot
# correct it by editing the visible hex value.  The paper format wants black
# headings, so the built-ins are neutralised once at build time — before
# ``add_styles``, which is what makes a spec able to ask for blue back.
HEADING_COLOUR = "000000"
# English UI names, the names a Chinese Word template carries, and the two
# one-off styles.  Both spellings matter: a template localised to Chinese names
# its headings ``标题 1`` while the styleId may still be ``Heading1``, and the
# reverse also occurs, so neither spelling alone covers both.
HEADING_STYLE_NAMES = tuple(
    ["Heading %d" % n for n in range(1, 10)]
    + ["标题 %d" % n for n in range(1, 10)]
    + ["Title", "标题", "Subtitle", "副标题"])
# Matched as styleIds too, for styles whose *name* was localised away from the
# English one while the id kept Word's spelling.
HEADING_STYLE_IDS = tuple(
    ["Heading%d" % n for n in range(1, 10)] + ["Title", "Subtitle"])
# Every built-in paragraph style is ``w:link``-ed to a character twin, and on a
# Chinese template the link target's *name* is ``标题 1 字符``.  Leaving those
# alone is not "not touching what the spec did not ask for": Word lists them in
# the style gallery in blue, and selecting a few words inside a heading and
# applying the character style turns them blue again.  The paragraph style being
# black is not enough when the linked style it points at is not.
HEADING_CHAR_STYLE_NAMES = tuple(
    ["Heading %d Char" % n for n in range(1, 10)]
    + ["标题 %d 字符" % n for n in range(1, 10)]
    + ["Title Char", "标题 字符", "Subtitle Char", "副标题 字符"])
HEADING_CHAR_STYLE_IDS = tuple(
    ["Heading%dChar" % n for n in range(1, 10)] + ["TitleChar", "SubtitleChar"])
# ``Subtitle`` is additionally italic, which a Chinese assignment does not want
# for a subtitle line (P9).
_SUBTITLE_KEYS = ("subtitle", "副标题", "subtitle char", "副标题 字符")
_HEADING_NEUTRALISE_ATTRS = ("w:themeColor", "w:themeShade", "w:themeTint")
# The six sides ``CT_PBdr`` allows.  ``Title`` carries a ``w:bottom`` rule in the
# theme accent colour, so a black title still sits on a blue underline — the
# exact symptom the original report described.
BORDER_SIDES = ("w:top", "w:left", "w:bottom", "w:right", "w:between", "w:bar")


def _has_theme_colour(colour) -> bool:
    """True when ``w:color`` resolves through the theme rather than an RGB value."""
    if any(colour.get(qn(attr)) is not None
           for attr in _HEADING_NEUTRALISE_ATTRS):
        return True
    val = colour.get(qn("w:val"))
    return val is None or val.lower() == "auto"


def _blacken_run_colour(rpr, report: dict, label: str) -> bool:
    """Force ``rpr/w:color`` to explicit black.  Shared by both style kinds.

    The element is replaced, not patched: ``themeColor`` (with ``themeShade`` /
    ``themeTint``) is what Word renders, so leaving it beside a new ``@w:val``
    would change the XML and not the page.  A style that is *already* an explicit
    colour is left byte-identical — blackening is a default, not a sweep.
    """
    colour = rpr.find(qn("w:color"))
    if colour is None:
        report["no_colour"].append(label)
        return False
    if not _has_theme_colour(colour):
        report["explicit"].append(label)
        return False
    rpr.remove(colour)
    new = OxmlElement("w:color")
    new.set(qn("w:val"), HEADING_COLOUR)
    insert_ordered(rpr, new, _rpr_order())
    return True


def _strip_italic(rpr) -> bool:
    """Remove ``w:i``/``w:iCs``; True when either was there."""
    removed = False
    for tag in ("w:i", "w:iCs"):
        el = rpr.find(qn(tag))
        if el is not None:
            rpr.remove(el)
            removed = True
    return removed


def _blacken_borders(pPr) -> bool:
    """Recolour every ``w:pBdr`` side to black, leaving thickness/spacing alone.

    ``Title`` ships ``<w:bottom w:val="single" w:sz="8" w:space="4"
    w:color="4F81BD" w:themeColor="accent1"/>``.  Only the colour attributes are
    touched: ``w:sz`` (8 = 1 pt) and ``w:space`` are the design, not the symptom,
    and a "make the title black" change that also thickened the rule would be
    fixing something nobody reported.  ``w:color`` is normalised too, because the
    hex sitting next to a theme colour is a guess Word ignores — and
    ``w:val="auto"`` (or a border with no ``w:val``) would render in the
    document's automatic colour, a second route back to the theme blue.
    """
    changed = False
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        return False
    for side in BORDER_SIDES:
        edge = pBdr.find(qn(side))
        if edge is None:
            continue
        for attr in _HEADING_NEUTRALISE_ATTRS:
            if edge.get(qn(attr)) is not None:
                del edge.attrib[qn(attr)]
                changed = True
        if edge.get(qn("w:color")) != HEADING_COLOUR:
            edge.set(qn("w:color"), HEADING_COLOUR)
            changed = True
    return changed


def _neutralise_style_colour(style, report: dict, label: str,
                             drop_italic: bool = False,
                             borders: bool = False) -> bool:
    """Neutralise one style: run colour, optional italic strip, optional borders.

    Both the paragraph styles and their ``*Char`` twins go through this one
    function — the two kinds differ only in *where* their ``w:rPr`` lives and in
    whether a ``w:pPr`` exists at all, and duplicating the logic per kind is how
    the character variants ended up blue in the first place.
    """
    changed = False
    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        # No run properties at all: nothing is coloured here, and adding a
        # colour would *install* an opinion the style did not have.
        report["missing"].append(label)
    else:
        changed |= _blacken_run_colour(rpr, report, label)
        if drop_italic and _strip_italic(rpr):
            changed = True
    if borders:
        pPr = style.element.find(qn("w:pPr"))
        if pPr is not None and _blacken_borders(pPr):
            report["borders"].append(label)
            changed = True
    return changed


def _neutralise_heading_styles(doc) -> dict:
    """Make the built-in heading / title / subtitle styles print black.

    Same shape as :func:`_neutralise_caption_style`, widened in three ways that
    each plug a hole the first version left open:

    * the **character** twins (``Heading 1 Char`` …) are included, because Word
      links each paragraph style to one (``w:link``) and a heading whose linked
      character style is still blue turns blue the moment a few words inside it
      are given that style;
    * ``Title``'s ``w:pBdr`` rule is recoloured, since a black title over a blue
      underline still reads as "the title is blue";
    * ``Subtitle`` loses its italic (P9) — and so does ``Subtitle Char``.

    Returns a report dict (``styles`` / ``touched`` / ``explicit`` /
    ``no_colour`` / ``missing`` / ``borders``) rather than the caption function's
    bare ``True``: the caption case had exactly one question ("was anything
    touched?"), while here the caller needs to know *which* styles were found and
    which of those already carried an explicit colour — a Chinese template
    supplies only the localised spellings, and that silence has to be reportable
    rather than invisible.

    Runs **before** ``add_styles``, so a spec's explicit ``color`` still wins:
    that is the opt-out channel, not a competing default.
    """
    report = {"styles": [], "touched": [], "explicit": [], "no_colour": [],
              "missing": [], "borders": []}
    wanted_names = {n.casefold()
                    for n in HEADING_STYLE_NAMES + HEADING_CHAR_STYLE_NAMES}
    wanted_ids = {i.casefold()
                  for i in HEADING_STYLE_IDS + HEADING_CHAR_STYLE_IDS}
    seen = set()
    for style in doc.styles:
        name = str(style.name or "")
        style_id = str(style.style_id or "")
        if name.casefold() not in wanted_names \
                and style_id.casefold() not in wanted_ids:
            continue
        label = name or style_id
        if label in seen:
            continue
        seen.add(label)
        report["styles"].append(label)
        is_subtitle = name.casefold() in _SUBTITLE_KEYS or \
            style_id.casefold() in ("subtitle", "subtitlechar")
        if _neutralise_style_colour(style, report, label,
                                    drop_italic=is_subtitle,
                                    borders=True):
            report["touched"].append(label)
    return report


_VERT_ALIASES = {
    "superscript": "superscript", "sup": "superscript", "script": "superscript",
    "上标": "superscript", "上角标": "superscript",
    "subscript": "subscript", "sub": "subscript",
    "下标": "subscript", "下角标": "subscript",
    "none": None, "normal": None, "baseline": None, "": None,
}


def _run_vert_align(run_spec: dict):
    """Normalise a run's script request to a ``w:vertAlign`` value or ``None``.

    Both ``vert_align`` and ``superscript``/``subscript`` booleans are honoured,
    because a hand-written spec tends to use whichever it saw first.  An
    explicit ``vert_align`` wins over the booleans.
    """
    raw = run_spec.get("vert_align")
    if raw is not None:
        key = str(raw).strip().lower()
        if key not in _VERT_ALIASES:
            raise ValueError(
                "unknown vert_align: %r (known: superscript, subscript, none)"
                % (raw,))
        return _VERT_ALIASES[key]
    if run_spec.get("superscript"):
        return "superscript"
    if run_spec.get("subscript"):
        return "subscript"
    return None


def _set_run_vert_align(run, vert: str) -> None:
    """Write ``<w:rPr><w:vertAlign w:val="vert"/>`` on *run* in schema order.

    ``CT_RPr`` is a *sequence*: a ``vertAlign`` emitted after ``sz``/``color``
    but before ``rtl`` is fine, one emitted at the end of the element is
    silently discarded by Word.  ``RPR_ORDER`` (shared with the cross-reference
    fields, which Word itself validated) supplies that position, and
    ``insert_ordered`` replaces an existing ``vertAlign`` rather than adding a
    duplicate.
    """
    from docx_crossref import RPR_ORDER
    rpr = run._r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        run._r.insert(0, rpr)
    va = OxmlElement("w:vertAlign")
    va.set(qn("w:val"), vert)
    insert_ordered(rpr, va, RPR_ORDER)


def add_runs(para, block: dict) -> None:
    runs = block.get("runs")
    if runs is None:
        runs = [{"text": block.get("text", "")}]
    for r in runs:
        # A literal ``\t`` in the text means a real ``<w:tab/>``: python-docx's
        # ``add_run`` would put it in ``<w:t>``, which Word renders as nothing.
        # TOC lines ("摘要<TAB>i") depend on this actually being a tab.
        text = r.get("text", "")
        parts = text.split("\t")
        run = None
        made = []
        for k, part in enumerate(parts):
            if k:
                (run if run is not None else para.add_run()).add_tab()
            if part or not parts[1:]:
                run = para.add_run(part)
                made.append(run)
        if run is not None:
            if r.get("bold"):
                run.bold = True
            if r.get("italic"):
                run.italic = True
            if r.get("underline"):
                run.underline = True
        # Character-level script position.  ``vert_align`` is the same switch
        # the cross-reference fields use, so a spec has one spelling for it:
        # ``superscript`` / ``subscript`` write a real ``<w:rPr><w:vertAlign>``
        # — the *text* script Word itself produces with Ctrl+Shift+= — which
        # is what makes "select it, press Ctrl+Shift+=, it returns to normal
        # size" true.  An OMML ``m:sSup`` object cannot do that: its raised
        # glyph is a formula-engine layout, and pressing the shortcut on it
        # stacks a second superscript on already-shrunken type.
        vert = _run_vert_align(r)
        if vert:
            for one in made:
                _set_run_vert_align(one, vert)


# ---- paragraph-level formatting ---------------------------------------------

def _effective_pt(para, block: dict) -> float:
    """Point size to use when converting a character indent into twips.

    Explicit override first, then the paragraph's own runs, then its style.
    """
    for key in ("char_width_pt", "first_line_indent_font_pt"):
        value = block.get(key)
        if value:
            return float(value)
    for run in para.runs:
        size = run.font.size
        if size is not None:
            return size.pt
    try:
        size = para.style.font.size
        if size is not None:
            return size.pt
    except (AttributeError, KeyError, ValueError):
        pass
    return DEFAULT_CHAR_PT


def _bookmark_para(para, name: str) -> str:
    """Wrap *para* in a bookmark: start after ``pPr``, end after everything.

    ``w:bookmarkStart`` is a child of ``w:p`` (not of ``w:pPr``), and it has to
    sit after the property block so the paragraph keeps its formatting while
    being fully enclosed by the bookmark range.
    """
    p = para._p
    bid = next_bm_id(p.getroottree().getroot())
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bid))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bid))
    pPr = p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(start)
    else:
        p.insert(0, start)
    p.append(end)
    return name


def _apply_para_format(para, block: dict, skip_bookmark: bool = False) -> dict:
    """Apply the shared paragraph-format vocabulary to *para*.

    Returns the subset that was actually applied, which also makes the helper
    testable without reaching into XML for every key.
    """
    pf = para.paragraph_format
    out: dict = {}

    if block.get("indent_left_mm") is not None:
        pf.left_indent = Mm(block["indent_left_mm"])
        out["indent_left_mm"] = block["indent_left_mm"]
    if block.get("indent_right_mm") is not None:
        pf.right_indent = Mm(block["indent_right_mm"])
        out["indent_right_mm"] = block["indent_right_mm"]

    if block.get("first_line_indent_chars") is not None:
        chars = float(block["first_line_indent_chars"])
        point_size = _effective_pt(para, block)
        # w:firstLine is twips (1/20 pt); Word writes both attributes and the
        # Chars twin wins, but keeping them consistent stops Word from
        # "repairing" the indent the next time the file is saved.
        pf.first_line_indent = Twips(int(round(chars * point_size * 20)))
        ind = para._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
        out["first_line_indent_chars"] = chars

    if block.get("line_spacing") is not None:
        pf.line_spacing = float(block["line_spacing"])
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        out["line_spacing"] = float(block["line_spacing"])

    if block.get("space_before_pt") is not None:
        pf.space_before = Pt(block["space_before_pt"])
        out["space_before_pt"] = block["space_before_pt"]
    if block.get("space_after_pt") is not None:
        pf.space_after = Pt(block["space_after_pt"])
        out["space_after_pt"] = block["space_after_pt"]

    if block.get("alignment") is not None:
        key = str(block["alignment"]).lower()
        if key not in _ALIGNMENTS:
            raise ValueError("unknown alignment: %r (choose %s)"
                             % (block["alignment"],
                                ", ".join(sorted(set(_ALIGNMENTS)))))
        para.alignment = _ALIGNMENTS[key]
        out["alignment"] = key

    if block.get("keep_with_next") is not None:
        pf.keep_with_next = bool(block["keep_with_next"])
        out["keep_with_next"] = bool(block["keep_with_next"])
    if block.get("page_break_before") is not None:
        pf.page_break_before = bool(block["page_break_before"])
        out["page_break_before"] = bool(block["page_break_before"])

    if block.get("bookmark") and not skip_bookmark:
        _bookmark_para(para, block["bookmark"])
        out["bookmark"] = block["bookmark"]

    stops = block.get("tab_stops")
    if stops:
        out["tab_stops"] = _apply_tab_stops(para, stops)

    border = block.get("bottom_border")
    if border:
        out["bottom_border"] = _apply_bottom_border(para, border)

    return out


def _apply_bottom_border(para, spec) -> dict:
    """Give *para* a bottom rule — how an exam answer line is actually drawn.

    ``spec`` is ``True``/a dict/``False``:

    * ``True`` → solid single 0.5pt black line;
    * ``False`` → remove an inherited bottom border;
    * dict → ``style`` (single/double/dashed/…), ``size`` (eighths of a point,
      Word's ``w:sz``), ``space`` (pt between text and line), ``color``
      (RRGGBB). Any side in ``sides`` may be given instead of just the bottom.

    Word stores this in ``w:pPr/w:pBdr``, which is an ``xsd:sequence`` —
    ``top`` must precede ``left``, ``bottom``, ``right``, ``between``, ``bar``.
    """
    pPr = para._p.get_or_add_pPr()
    existing = pPr.find(qn("w:pBdr"))
    if existing is not None:
        pPr.remove(existing)
    if spec is False:
        return {"removed": True}

    if spec is True:
        spec = {}
    if not isinstance(spec, dict):
        raise ValueError("bottom_border must be a bool or a dict, got %r"
                         % (spec,))

    style = str(spec.get("style", "single")).lower()
    if style not in _BORDER_STYLES:
        raise ValueError("unknown border style %r (choose %s)"
                         % (spec.get("style"), ", ".join(sorted(_BORDER_STYLES))))
    size = int(spec.get("size", 8))          # 8 eighths = 1pt; 6 ≈ 0.75pt
    space = int(round(float(spec.get("space", 1)) * 1))   # already in pt units
    color = str(spec.get("color", "000000")).lstrip("#").upper()

    sides = spec.get("sides") or ["bottom"]
    if isinstance(sides, str):
        sides = [sides]

    pBdr = OxmlElement("w:pBdr")
    applied = {}
    for side in _BORDER_SIDES:
        if side not in sides:
            continue
        el = OxmlElement("w:" + side)
        el.set(qn("w:val"), style)
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), str(space))
        el.set(qn("w:color"), color)
        pBdr.append(el)
        applied[side] = {"style": style, "size": size, "space": space,
                         "color": color}

    if not applied:
        return {"removed": True}
    # CT_PPr is an xsd:sequence: pBdr sits after numPr/suppressLineNumbers and
    # before shd/tabs/spacing/ind/jc.
    pPr.insert_element_before(
        pBdr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
        "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
        "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
        "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
        "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
        "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
        "w:rPr", "w:sectPr", "w:pPrChange")
    return applied


def _apply_tab_stops(para, stops) -> list:
    """Write ``w:tabs`` tab stops; a TOC line needs this to show dotted leaders.

    Each stop is ``{"pos_mm": 155, "align": "right", "leader": "dot"}``.
    ``pos_mm`` may also be given as ``pos_cm``; ``align`` defaults to ``left``
    and ``leader`` to ``none``.  Word keeps the stops in ascending ``w:pos``
    order, so they are sorted before writing.
    """
    from docx.enum.text import WD_TAB_ALIGNMENT as A, WD_TAB_LEADER as LD
    written = []
    resolved = []
    for stop in stops:
        if isinstance(stop, dict):
            mm = stop.get("pos_mm")
            if mm is None and stop.get("pos_cm") is not None:
                mm = float(stop["pos_cm"]) * 10.0
            if mm is None:
                raise ValueError("tab stop needs pos_mm (or pos_cm): %r" % (stop,))
            al = str(stop.get("align", "left")).lower()
            ld = str(stop.get("leader", "none")).lower()
            if al not in _TAB_ALIGNS:
                raise ValueError("unknown tab align %r (choose %s)"
                                 % (al, ", ".join(sorted(_TAB_ALIGNS))))
            if ld not in _TAB_LEADERS:
                raise ValueError("unknown tab leader %r (choose %s)"
                                 % (ld, ", ".join(sorted(_TAB_LEADERS))))
            resolved.append((float(mm), _TAB_ALIGNS[al], _TAB_LEADERS[ld],
                             al, ld))
        else:                                    # bare number = left tab at mm
            resolved.append((float(stop), A.LEFT, LD.SPACES, "left", "none"))

    pf = para.paragraph_format
    for mm, align, leader, al, ld in sorted(resolved, key=lambda r: r[0]):
        pf.tab_stops.add_tab_stop(Mm(mm), align, leader)
        written.append({"pos_mm": mm, "align": al, "leader": ld})
    return written


def _item_text(item) -> str:
    if isinstance(item, dict):
        return str(item.get("text", ""))
    return str(item)


class _Ns:
    """Stand-in for a CLI argparse namespace."""

    def __init__(self, **kw):
        self.__dict__.update(kw)


def _caption_args(index: int, block: dict, default_label: str = "图") -> _Ns:
    """Assemble ``cmd_caption``'s namespace for a caption block.

    Every knob is read with ``getattr`` on the *receiving* side, so a spec may
    ask for the newer switches (``ref_bookmark``, ``pstyle``) even while
    ``docx_crossref`` is still being extended — an unknown switch is simply
    not passed on.

    The paragraph style is deliberately *not* forced to ``Caption``: a spec
    that already names one (``正文``, ``SCaption``, ``EX图表名``) has chosen the
    template's own caption look, and overriding it would silently drop the
    size and centring that template defines.  ``Caption`` is only the fallback
    for a caption that names no style at all.
    """
    return _Ns(index=index,
               label=block.get("label") or default_label,
               text=block.get("text", ""),
               bookmark=block.get("bookmark"),
               # Word's own shape: a ``_Ref…`` bookmark covering the whole
               # caption is what makes the table of figures clickable.  It is
               # an *extra* bookmark; ``bookmark_span`` below still governs the
               # caller-named one, so a REF keeps returning just the number.
               ref_bookmark=bool(block.get("ref_bookmark", True)),
               # The caption paragraph style, so "Insert Table of Figures"
               # (which filters by style) can see the entry at all.
               pstyle=block.get("pstyle", block.get("style") or "Caption"),
               bookmark_span=block.get("bookmark_span", "number"),
               drop_text=bool(block.get("drop_text", False)))


def _with_caption_defaults(block: dict, doc) -> dict:
    """Fill in the paragraph alignment a caption needs when unnamed by style.

    Word's built-in ``Caption`` style is left-aligned; a caption in these papers
    is centred.  A spec that names its own ``style`` already inherits that
    template's look, so nothing is added here.  The *size* is set on the
    ``Caption`` style itself (see ``_neutralise_caption_style``), because font
    size is not a paragraph-format key and per-block run formatting would have
    to be threaded through ``cmd_caption`` for no benefit.
    """
    if block.get("style") or block.get("pstyle"):
        return block
    out = dict(block)
    out.setdefault("alignment", "center")
    return out


def _crossref_field_args(block: dict, index: int, prefix: str) -> _Ns:
    """Assemble ``cmd_ref``'s namespace, including the newer display switches.

    ``vert_align`` is the single source of truth; ``superscript`` is derived
    from it when not given, and both are always present so ``cmd_ref`` may read
    either one with ``getattr`` while it is being extended in parallel.
    """
    vert = block.get("vert_align")
    sup = block.get("superscript")
    if vert is None:
        vert = "superscript" if sup else "normal"
    if sup is None:
        sup = vert == "superscript"
    return _Ns(index=index,
               target=block["target"],
               prefix=prefix,
               suffix=block.get("suffix", ""),
               format=block.get("format", "default"),
               cached=block.get("cached"),
               superscript=bool(sup),
               vert_align=vert,
               hyperlink=bool(block.get("hyperlink", True)),
               # A spec is written top-to-bottom, so a body reference to the
               # reference list is necessarily forward: the bookmark appears
               # later in the document. Bookmarks are document-wide, so the
               # anchor is checked after every block is built instead.
               allow_missing=True)


def _break_target(doc, new_section, block: dict):
    """Which section a ``section_break`` block's page setup describes.

    ``apply_to: "next"`` (the default) configures the section that starts at
    the break — the one whose content follows.  ``apply_to: "previous"``
    configures the section the break closes.  "next" is the default because
    the last section of a document is followed by no break at all, so it can
    only be reached from the break above it.
    """
    where = str(block.get("apply_to", "next")).strip().lower()
    if where in ("next", "after", "following", ""):
        return new_section
    if where in ("previous", "prev", "before", "closed"):
        secs = list(doc.sections)
        if len(secs) < 2:
            raise ValueError("apply_to='previous' needs a section to close")
        return secs[-2]
    raise ValueError("unknown apply_to: %r (choose next or previous)"
                     % (block["apply_to"],))


def _row_paragraphs(row) -> list:
    """Every paragraph of every cell in *row*, in cell order.

    A row has no ``keepNext`` of its own — ``keepNext`` is a *paragraph*
    property, so "this row must not end the page" can only be said on all of
    the row's paragraphs at once.
    """
    out = []
    for cell in row.cells:
        out.extend(cell.paragraphs)
    return out


def _cant_split(row) -> None:
    """``<w:trPr><w:cantSplit/></w:trPr>``: the row never breaks across pages.

    ``w:cantSplit`` is the only table-row property Word honours here; there is
    no table-level equivalent.  ``trPr`` is an ``xsd:sequence``, so the element
    goes in through :func:`docx_utils.insert_ordered` — appending it after a
    ``trHeight`` would make Word drop it while the XML still looked fine.
    """
    trPr = row._tr.get_or_add_trPr()
    insert_ordered(trPr, OxmlElement("w:cantSplit"))


def _keep_rows_together(rows) -> None:
    """Make a block of rows hold together: no row splits, no row is left alone.

    Two separate channels, because Word spells them in two places:

    * every row gets ``w:cantSplit`` (the row's content stays on one page);
    * every row *but the last* gets ``keep_with_next`` on each of its cell
      paragraphs (Word keeps that paragraph on the same page as the next one,
      which is what glues row *n* to row *n+1*).

    Writing ``w:keepNext`` into ``w:trPr`` instead would be silently ignored by
    Word — ``keepNext`` is not among ``trPr``'s children — while every XML
    checker still reported the document valid.
    """
    if not rows:
        return
    for row in rows:
        _cant_split(row)
    for row in rows[:-1]:
        for para in _row_paragraphs(row):
            para.paragraph_format.keep_with_next = True


def _caption_glues_to_next_table(blocks: list, index: int) -> bool:
    """True when ``blocks[index]`` is a ``table_caption`` directly before a table.

    "Directly before" means **no block in between**: a caption that has to
    share a page with its table is useless if an empty paragraph sits between
    them, and a caption two blocks away is not what caused the split.  Kept as
    a function (rather than an inline index comparison at the call site) so the
    adjacency rule is stated once and can be tested on its own.
    """
    if not isinstance(blocks, list) or not 0 <= index < len(blocks):
        return False
    block = blocks[index]
    if not isinstance(block, dict) or block.get("type") != "table_caption":
        return False
    if index + 1 >= len(blocks):
        return False
    nxt = blocks[index + 1]
    return isinstance(nxt, dict) and nxt.get("type") == "table"


def add_block(doc, block: dict, blocks: list | None = None,
              index: int | None = None) -> list:
    """Add one block; return the body paragraphs it created or attached to.

    Every paragraph-producing block then runs through
    :func:`_apply_para_format`, so the format vocabulary is uniform instead of
    being re-implemented per block type.

    *blocks* / *index* are optional and used only by ``table_caption``, which
    has to look at what comes *next* in the spec to know whether it belongs to
    a table.  A caller that builds blocks one at a time (as the tests do) may
    omit them and simply gets no automatic caption-to-table link.
    """
    btype = block["type"]
    made: list = []
    own_bookmark = False        # the block placed block["bookmark"] itself

    if btype == "heading":
        made = [doc.add_heading(block.get("text", ""),
                                level=block.get("level", 1))]
    elif btype == "paragraph":
        para = doc.add_paragraph(style=block.get("style"))
        add_runs(para, block)
        made = [para]
    elif btype == "bullet_list":
        for item in block.get("items", []):
            made.append(doc.add_paragraph(_item_text(item),
                                          style=block.get("style",
                                                          "List Bullet")))
    elif btype == "numbered_list":
        # A numbered list needs a *definition* to point at. Word's "List
        # Number" style carries no numPr, so paragraphs styled with it render
        # with no numbers at all; attach a real numId instead.
        from docx_numbering import cmd_define, set_numpr
        nid = block.get("num_id")
        if nid is None:
            nid = cmd_define(doc, _Ns(
                text=block.get("num_text", "%1."),
                fmt=block.get("num_fmt", "decimal"),
                start=int(block.get("start", 1)), jc="left",
                indent_left=block.get("indent_left", 425),
                hanging=block.get("hanging", 425)))["numId"]
        prefix = block.get("bookmark_prefix")
        if prefix:
            own_bookmark = True         # per-item bookmarks take over
        for k, item in enumerate(block.get("items", []), start=1):
            para = doc.add_paragraph(_item_text(item),
                                     style=block.get("style"))
            set_numpr(para, nid, block.get("ilvl", 0))
            if prefix:
                _bookmark_para(para, "%s%d" % (prefix, k))
            made.append(para)
    elif btype == "caption":
        # "图 N  题注" carrying a SEQ counter + bookmark, so later paragraphs
        # can point at it with a REF field.
        #
        # A caption is also *the* thing Word's "Insert Table of Figures"
        # collects, and that command builds ``PAGEREF`` links out of the
        # ``_Toc…`` bookmarks Word itself adds round each caption.  So the
        # ``_Ref`` bookmark and a caption paragraph style are on by default
        # here and can only be turned off explicitly.
        from docx_crossref import cmd_caption
        doc.add_paragraph("")
        cmd_caption(doc, _caption_args(
            len(doc.paragraphs) - 1, _with_caption_defaults(block, doc)))
        made = [doc.paragraphs[-1]]
        own_bookmark = True
    elif btype == "table_caption":
        # A table's caption is a *separate paragraph* above the table, so the
        # block exists on its own rather than hanging off ``table``: a spec can
        # then decide per table whether the caption goes above or below.
        from docx_crossref import cmd_caption
        doc.add_paragraph("")
        cmd_caption(doc, _caption_args(
            len(doc.paragraphs) - 1,
            _with_caption_defaults(block, doc), default_label="表"))
        made = [doc.paragraphs[-1]]
        own_bookmark = True
        # A caption that starts a page alone above its table is the defect this
        # adjacency check removes.  The *table* side (first row ``cantSplit``)
        # is handled by the ``table`` branch, which knows it is the table this
        # caption belongs to.
    elif btype == "crossref":
        # Appends a REF field to the previous paragraph, or to a new one.
        from docx_crossref import cmd_ref
        fresh = block.get("own_paragraph") or not doc.paragraphs
        if fresh:
            doc.add_paragraph(block.get("prefix", ""))
        cmd_ref(doc, _crossref_field_args(
            block, len(doc.paragraphs) - 1,
            "" if fresh else block.get("prefix", "")))
        made = [doc.paragraphs[-1]]
    elif btype == "table":
        header = block.get("header", [])
        rows = block.get("rows", [])
        ncols = len(header) if header else (len(rows[0]) if rows else 1)
        # A ``table_caption`` block immediately above this one belongs to this
        # table, and the link is *automatic* (D-029): the caption paragraph is
        # asked to stay with the next paragraph, and the table's first row is
        # told not to split.  The paragraph is read here, before anything is
        # appended, because that is the last moment ``doc.paragraphs[-1]`` is
        # still the caption rather than the table's own content.
        caption_para = None
        if blocks is not None and index is not None \
                and _caption_glues_to_next_table(blocks, index - 1):
            caption_para = doc.paragraphs[-1] if doc.paragraphs else None
        # A table caption conventionally sits *above* the table, and Word's
        # "Insert Table of Figures" only picks up a caption that is its own
        # paragraph — so it is emitted before the table, not inside it.
        if block.get("caption"):
            from docx_crossref import cmd_caption
            doc.add_paragraph("")
            cmd_caption(doc, _caption_args(
                len(doc.paragraphs) - 1,
                {"label": block.get("caption_label", "表"),
                 "text": block["caption"],
                 "bookmark": block.get("caption_bookmark")},
                default_label="表"))
        table = doc.add_table(rows=0, cols=ncols)
        table.style = block.get("style", "Table Grid")
        if header:
            cells = table.add_row().cells
            for i, text in enumerate(header):
                cells[i].text = str(text)
                if block.get("header_bold", True):
                    for para in cells[i].paragraphs:
                        for run in para.runs:
                            run.bold = True
        for row in rows:
            cells = table.add_row().cells
            for i, text in enumerate(row):
                cells[i].text = str(text)
        if caption_para is not None:
            caption_para.paragraph_format.keep_with_next = True
            if table.rows:
                _cant_split(table.rows[0])
        if block.get("keep_together"):
            # The ``header`` row participates: it is part of the table, so it
            # must not be the row that splits.
            _keep_rows_together(table.rows)
        made = []
    elif btype == "image":
        # ``is not None``, not truthiness: ``width_frac: 0.0`` is a *number the
        # spec wrote*, and a truthiness test would treat it as "no width given"
        # and silently fall back to the picture's intrinsic size.  Same reason
        # ``add_styles`` reads ``outline_level`` this way.
        width_mm = block.get("width_mm")
        width_frac = block.get("width_frac")
        if width_frac is not None:
            # A fraction is a share of *this section's* text width, so the
            # figure follows the page geometry instead of a hard-coded 158 mm.
            # Content widths differ per volume (170 mm vs 160 mm in this
            # project), which is exactly why the fraction is resolved here.
            width = Mm(content_width_mm(doc.sections[-1]) * float(width_frac))
        elif width_mm is not None:
            width = Mm(width_mm)
        else:
            width = None
        doc.add_picture(block["path"], width=width)
        made = [doc.paragraphs[-1]]
    elif btype == "page_break":
        para = doc.add_paragraph()
        para.add_run().add_break(WD_BREAK.PAGE)
        made = [para]
    elif btype == "toc":
        from docx_edit import _add_field
        para = doc.add_paragraph()
        # `instr` lets the caller ask for a table-of-figures
        # (TOC \h \c "图") instead of the outline TOC.
        _add_field(para, block.get("instr") or r' TOC \o "1-3" \h \z \u ',
                   "Table of contents - open in Word/LibreOffice and "
                   "update fields to populate.")
        made = [para]
    elif btype == "section_break":
        section, _moved = add_section_break(
            doc, block.get("start_type", "new_page"),
            reuse_last=block.get("reuse_last_paragraph", True))
        # A break sits between two sections, so the block's page setup has to
        # name which one it describes.  The default is the section that
        # *begins* here, because the final section of a document has no break
        # after it and could otherwise never be configured.
        configure_section(_break_target(doc, section, block), block)
    else:
        raise ValueError(f"unknown block type: {btype}")

    for i, para in enumerate(made):
        # ``bookmark`` describes one range; a multi-paragraph block gets it on
        # the paragraph it starts with, and only that one.
        _apply_para_format(para, block,
                           skip_bookmark=own_bookmark or i > 0)
    return made


# ---- spec validation --------------------------------------------------------
#
# Everything above reads a block with ``block.get(...)``, so the only possible
# outcome of a misspelt key used to be silence: ``{"blockz": [...]}`` built a
# valid, empty document, printed ``"blocks": 0`` and exited 0, and the author
# had no way to see that the content never went in.  The vocabulary the builder
# actually reads is therefore declared here and checked *before* a single
# paragraph exists.
#
# Strict on purpose: an unknown key is an error, not a warning.  Making the
# author fix a typo is far cheaper than shipping a document that is quietly
# missing a chapter.  Values are not clamped either — an out-of-range value is
# reported with the range it would have been clamped into.

# ``SpecError`` is the one exception class that ``main`` turns into exit 2; any
# other exception keeps exit 1 and its traceback, because that is a builder bug
# rather than something the spec author can fix.
class SpecError(Exception):
    """A spec that asks for something the builder cannot honour."""


TOP_LEVEL_KEYS = ("blocks", "footer", "footer_page_numbers", "header", "page",
                  "styles")
PAGE_KEYS = ("height_mm", "margins_mm", "orient", "width_mm")
MARGIN_KEYS = ("bottom", "footer", "gutter", "header", "left", "right", "top")
STYLE_KEYS = ("base", "bold", "color", "font", "font_east_asia", "italic",
              "name", "outline_level", "size_pt")
RUN_KEYS = ("bold", "italic", "subscript", "superscript", "text", "underline",
            "vert_align")
TAB_STOP_KEYS = ("align", "leader", "pos_cm", "pos_mm")
BOTTOM_BORDER_KEYS = ("color", "sides", "size", "space", "style")
# ``add_page_number_field``'s keyword arguments — a spec may set them at the top
# level or inside a ``section_break``; an unknown one there is a ``TypeError``
# at ``**kwargs`` expansion time, i.e. another bare traceback.
PAGE_NUMBER_KEYS = ("align", "bold", "fmt", "font", "page_type", "prefix",
                    "size_pt", "start", "suffix", "template", "total")
# ``char_width_pt`` / ``first_line_indent_font_pt`` are inputs to the character
# → twips conversion rather than paragraph *outputs*, so they are not in
# ``PARA_FORMAT_KEYS``; a spec sets them on the same blocks all the same.
CHAR_WIDTH_KEYS = ("char_width_pt", "first_line_indent_font_pt")

# The one shape a spec author most often gets wrong, spelled out so it can be
# quoted verbatim in an error message.  ``margins_mm`` is an *object* with four
# required and three optional members; writing it as a list of numbers in the
# obvious visual order (top, right, bottom, left) is a natural mistake, and
# "must be a JSON object" alone does not say which names to use instead.
MARGINS_SHAPE_EXAMPLE = ('{"top": 25, "bottom": 25, "left": 20, "right": 20}')


def _with_para_format(*keys) -> set:
    """A block's own keys plus the shared paragraph-format vocabulary."""
    return set(keys) | set(PARA_FORMAT_KEYS) | set(CHAR_WIDTH_KEYS)


BLOCK_KEYS = {
    "heading": _with_para_format("text", "level"),
    "paragraph": _with_para_format("text", "runs", "style"),
    "bullet_list": _with_para_format("items", "style"),
    "numbered_list": _with_para_format("items", "num_id", "num_text",
                                       "num_fmt", "start", "indent_left",
                                       "hanging", "bookmark_prefix", "ilvl",
                                       "style"),
    "caption": _with_para_format("label", "text", "style", "pstyle",
                                 "bookmark_span", "ref_bookmark", "drop_text"),
    "table_caption": _with_para_format("label", "text", "style", "pstyle",
                                       "bookmark_span", "ref_bookmark",
                                       "drop_text"),
    "crossref": _with_para_format("target", "prefix", "suffix", "format",
                                  "cached", "superscript", "vert_align",
                                  "hyperlink", "own_paragraph"),
    # ``table`` produces no paragraph of its own, and its ``style`` is the
    # *table* style, so the paragraph vocabulary does not apply to it.
    "table": {"header", "rows", "style", "header_bold", "caption",
              "caption_label", "caption_bookmark", "keep_together"},
    "image": _with_para_format("path", "width_mm", "width_frac"),
    "page_break": _with_para_format(),
    "toc": _with_para_format("instr"),
    # ``section_break`` has its own ``header``/``footer`` keys, so a paragraph
    # ``alignment`` here would be a silent no-op — exactly what this check
    # exists to remove.
    "section_break": {"start_type", "page", "header", "header_align",
                      "header_font", "header_size_pt", "header_bold",
                      "footer", "footer_align", "footer_font",
                      "footer_size_pt", "footer_bold", "footer_page_numbers",
                      "different_first_page", "cols", "no_footer",
                      "reuse_last_paragraph", "apply_to"},
}


def _suggestions(key, allowed) -> list:
    """Nearest legitimate spellings of *key*, best first (may be empty)."""
    return difflib.get_close_matches(str(key), sorted(allowed), n=2,
                                     cutoff=0.5)


def _shape_hint(where: str) -> str:
    """A concrete example object for the place that expected one.

    "must be a JSON object, got list" says what is wrong but not what would be
    right, so an author handed it still has to go and read the source.  The
    examples below are the *smallest* objects that work at each nesting level,
    which is what makes the message actionable rather than merely accurate.

    ``margins_mm`` is the case this exists for: a spec author who wrote
    ``"margins_mm": [25, 25, 20, 20]`` has the right numbers in the right
    order, and "must be an object" alone leaves them to guess four key names.
    """
    if "margins_mm" in where:
        return (MARGINS_SHAPE_EXAMPLE
                + " (header / footer / gutter are optional)")
    if where == "spec" or where.startswith("spec."):
        return '{"blocks": [...], "page": {...}}'
    if "runs[" in where:
        return '{"text": "bold bit", "bold": true}'
    if "tab_stops[" in where:
        return '{"pos_mm": 155, "align": "right", "leader": "dot"}'
    if "bottom_border" in where:
        return '{"style": "single", "size": 8, "space": 1, "color": "000000"}'
    if where.startswith("styles[") or where == "styles":
        return '{"name": "Heading 1", "size_pt": 16}'
    if "items[" in where:
        return '{"text": "first item"}'
    return '{"type": "paragraph", "text": "..."}'


def _check_keys(where: str, obj, allowed, problems: list) -> bool:
    """Flag every key of *obj* that is not in *allowed*; True for a dict.

    One problem line per object rather than per key: a spec with two typos in
    one block reads better as one sentence than as two.

    The "got list" branch names the shape that would have worked.  A spec author
    who wrote ``"margins_mm": [25, 25, 20, 20]`` has the right numbers in the
    right order — telling them "must be an object" alone leaves them to guess
    the key names, and ``margins_mm`` has six of them.
    """
    if not isinstance(obj, dict):
        problems.append("%s must be a JSON object, got %s; expected e.g. %s"
                        % (where, type(obj).__name__, _shape_hint(where)))
        return False
    unknown = [k for k in obj if k not in allowed]
    if unknown:
        named = []
        for key in unknown:
            near = _suggestions(key, allowed)
            named.append('"%s"%s' % (key, ' (did you mean "%s"?)' % near[0]
                                     if near else ""))
        message = ("%s: unknown key %s; valid keys here: %s"
                   % (where, ", ".join(named), ", ".join(sorted(allowed))))
        # ``valid keys here`` names the members but not how they nest, which is
        # the part a list-shaped ``margins_mm`` got wrong in the first place.
        # The sentence above is kept verbatim; the shape is appended, so the
        # existing wording and its tests still hold.
        if "margins_mm" in where:
            message += "; expected shape: %s" % _shape_hint(where)
        problems.append(message)
    return True


def _check_page(where: str, page, problems: list) -> None:
    if page is None:
        return
    if _check_keys(where, page, PAGE_KEYS, problems):
        margins = page.get("margins_mm")
        if margins is not None:
            _check_keys(where + ".margins_mm", margins, MARGIN_KEYS, problems)


def _check_page_numbers(where: str, fpn, problems: list) -> None:
    # ``true`` is the documented shorthand for "the default Page X of Y".
    if fpn is None or isinstance(fpn, bool):
        return
    _check_keys(where, fpn, PAGE_NUMBER_KEYS, problems)


def _check_runs(where: str, runs, problems: list) -> None:
    if runs is None:
        return
    if not isinstance(runs, list):
        problems.append('%s: "runs" must be a list of run objects' % where)
        return
    for k, run in enumerate(runs):
        _check_keys("%s.runs[%d]" % (where, k), run, RUN_KEYS, problems)


def _check_items(where: str, items, problems: list) -> None:
    if items is None:
        return
    if not isinstance(items, list):
        problems.append('%s: "items" must be a list' % where)
        return
    for k, item in enumerate(items):
        if isinstance(item, dict):
            _check_keys("%s.items[%d]" % (where, k), item, ("text",),
                        problems)


def _check_level(where: str, block: dict, problems: list) -> None:
    """``heading``'s level: python-docx takes 0..9; report the range, no clamp."""
    if "level" not in block:
        return
    level = block["level"]
    whole = (isinstance(level, (int, float)) and not isinstance(level, bool)
             and float(level).is_integer())
    if not whole:
        problems.append('%s: "level" must be a whole number 0..9, got %r'
                        % (where, level))
    elif not 0 <= level <= 9:
        problems.append('%s: "level" out of range: %r (allowed 0..9: 0 writes '
                        "Word's Title style, 1..9 write Heading 1..9); the "
                        "value is not clamped" % (where, level))


def _check_table(where: str, block: dict, problems: list) -> None:
    """A row longer than the table is wide loses cells — silently, before."""
    header = block.get("header") or []
    rows = block.get("rows") or []
    if header and not isinstance(header, list):
        problems.append('%s: "header" must be a list of cell values' % where)
        header = []
    if not isinstance(rows, list):
        problems.append('%s: "rows" must be a list of rows' % where)
        return
    if header:
        ncols = len(header)
    elif rows and isinstance(rows[0], list):
        ncols = len(rows[0])
    else:
        ncols = 1
    for k, row in enumerate(rows):
        if not isinstance(row, list):
            problems.append("%s: rows[%d] must be a list of cell values"
                            % (where, k))
        elif len(row) > ncols:
            problems.append("%s: rows[%d] has %d cells but the table has %d "
                            "column(s)%s; the extra cells would be dropped"
                            % (where, k, len(row), ncols,
                               ' (from "header")' if header else ""))


def _check_keep_together(where: str, block: dict, problems: list) -> None:
    """``table``'s ``keep_together`` is a bool; a truthy string is a typo.

    ``"keep_together": "yes"`` would work by accident (any non-empty string is
    truthy) and then break the day someone writes ``"false"`` — which is also
    truthy.  Accepting only a real bool keeps the key honest, and ``bool`` is
    excluded from the number branch for the same reason ``_check_level`` does.
    """
    if "keep_together" not in block:
        return
    value = block["keep_together"]
    if not isinstance(value, bool):
        problems.append('%s: "keep_together" must be true or false, got %r'
                        % (where, value))


def _check_width_frac(where: str, block: dict, problems: list) -> None:
    """``width_frac`` is a share of the text width: ``0 < f <= 1``, not clamped.

    Two separate errors on purpose.  Giving both keys is a *contradiction* the
    builder cannot resolve without silently picking one, so it is reported as
    such rather than as a range problem.  A fraction outside ``(0, 1]`` is
    reported with the range it would have been clamped into — the repository
    does not clamp, because a figure silently half the requested size is worse
    than a rejected spec.
    """
    frac = block.get("width_frac")
    if frac is None:
        return
    if block.get("width_mm") is not None:
        problems.append('%s: "width_frac" and "width_mm" are mutually '
                        "exclusive; give one or the other (width_frac is a "
                        "share of the text width, width_mm is absolute)"
                        % where)
        return
    if isinstance(frac, bool) or not isinstance(frac, (int, float)):
        problems.append('%s: "width_frac" must be a number 0 < f <= 1, got %r'
                        % (where, frac))
        return
    if not 0 < frac <= 1:
        problems.append('%s: "width_frac" out of range: %r (allowed 0 < f <= 1: '
                        "1.0 is the full text width, 0.5 is half of it); the "
                        "value is not clamped" % (where, frac))


def _check_image(where: str, block: dict, problems: list) -> None:
    if not block.get("path"):
        problems.append('%s: image blocks need "path"' % where)
    _check_width_frac(where, block, problems)


def _check_block(index: int, block, problems: list) -> None:
    where = "blocks[%d]" % index
    if not isinstance(block, dict):
        problems.append("%s must be a JSON object, got %s; expected e.g. %s"
                        % (where, type(block).__name__, _shape_hint(where)))
        return
    btype = block.get("type")
    if btype is None:
        problems.append('%s: needs "type"; choose one of: %s'
                        % (where, ", ".join(sorted(BLOCK_KEYS))))
        return
    if not isinstance(btype, str) or btype not in BLOCK_KEYS:
        near = _suggestions(btype, BLOCK_KEYS)
        problems.append('%s: unknown block type %r%s; choose one of: %s'
                        % (where, btype,
                           ' (did you mean "%s"?)' % near[0] if near else "",
                           ", ".join(sorted(BLOCK_KEYS))))
        return

    label = '%s (type "%s")' % (where, btype)
    if not _check_keys(label, block, BLOCK_KEYS[btype] | {"type"}, problems):
        return

    if btype == "heading":
        _check_level(where, block, problems)
    elif btype == "table":
        _check_table(where, block, problems)
        _check_keep_together(where, block, problems)
    elif btype == "image":
        _check_image(where, block, problems)
    elif btype == "crossref":
        if not block.get("target"):
            problems.append('%s: crossref blocks need "target" (a bookmark '
                            "name)" % where)
    elif btype == "section_break":
        _check_page(where + ".page", block.get("page"), problems)
        _check_page_numbers(where + ".footer_page_numbers",
                            block.get("footer_page_numbers"), problems)

    _check_runs(where, block.get("runs"), problems)
    _check_items(where, block.get("items"), problems)

    stops = block.get("tab_stops")
    if isinstance(stops, list):
        for k, stop in enumerate(stops):
            if isinstance(stop, dict):
                _check_keys("%s.tab_stops[%d]" % (where, k), stop,
                            TAB_STOP_KEYS, problems)
    border = block.get("bottom_border")
    if isinstance(border, dict):
        _check_keys(where + ".bottom_border", border, BOTTOM_BORDER_KEYS,
                    problems)


def validate_spec(spec: dict) -> None:
    """Raise :class:`SpecError` listing everything wrong with *spec* at once.

    Reporting all problems in one pass means a spec with three typos needs one
    edit-and-rerun cycle, not three.
    """
    problems: list = []
    if not _check_keys("spec", spec, TOP_LEVEL_KEYS, problems):
        raise SpecError("; ".join(problems))

    _check_page("spec.page", spec.get("page"), problems)
    _check_page_numbers("spec.footer_page_numbers",
                        spec.get("footer_page_numbers"), problems)

    styles = spec.get("styles")
    if styles is not None:
        if not isinstance(styles, list):
            problems.append('spec: "styles" must be a list of style objects')
        else:
            for i, style in enumerate(styles):
                where = "styles[%d]" % i
                if _check_keys(where, style, STYLE_KEYS, problems):
                    if not style.get("name"):
                        problems.append('%s: needs "name"' % where)

    blocks = spec.get("blocks")
    if blocks is not None:
        if not isinstance(blocks, list):
            problems.append('spec: "blocks" must be a list of block objects')
        else:
            for i, block in enumerate(blocks):
                _check_block(i, block, problems)

    if problems:
        raise SpecError("; ".join(problems))


def _describe(exc) -> str:
    """One readable clause describing an exception raised while building."""
    if isinstance(exc, SystemExit):
        return str(exc.code)
    if isinstance(exc, KeyError):
        # ``KeyError`` arrives from two places with different shapes: a dict
        # lookup of a key the spec forgot (``'path'``) and python-docx's own
        # "no style with name 'x'" prose.  Only the first is worth rewording.
        text = exc.args[0] if exc.args else ""
        text = text if isinstance(text, str) else repr(text)
        return text if " " in text else "missing required key %r" % (text,)
    text = str(exc)
    return text or type(exc).__name__


# Exceptions the builder raises *because the spec asked for something
# impossible* — an unknown alignment, an unknown page-number format, a row
# wider than the table, a missing "path".  They are the author's to fix, so
# they are reported like any other spec problem instead of as a traceback.
# Anything outside this tuple keeps exit 1 and its traceback: that is a bug in
# the builder, and the traceback is what makes it locatable.
_BUILD_ERRORS = (KeyError, IndexError, TypeError, ValueError, SystemExit)


def _input_error(message: str) -> int:
    """Reject the spec: JSON on stdout, one ``error:`` line on stderr, exit 2.

    Exit 2 promises that *nothing was written*, so a caller can retry after
    fixing the spec without wondering whether half a document exists.

    The stderr line carries the repository-wide ``error:`` prefix (D-024) rather
    than the script name, so a caller can tell a refusal from a warning by the
    first token alone, exactly as the other CLIs allow.  The JSON on stdout is
    **kept**: ``docx_create`` is one of the scripts SKILL.md lists as
    "JSON on stdout", and its refusal payload is the machine-readable half of
    the same answer (``ok: false`` plus the reason).  Emptying stdout here would
    make the two channels disagree about whether the run produced a result.
    """
    print(json.dumps({"ok": False, "output": None, "error": message},
                     ensure_ascii=False))
    print("error: %s" % message, file=sys.stderr)
    return 2


def _load_spec(path: str) -> dict:
    """Read the spec file; every "this file cannot be used" becomes SpecError."""
    try:
        with open(path, encoding="utf-8") as f:
            text = f.read()
    except OSError as exc:
        raise SpecError("cannot read spec %s: %s" % (path, exc.strerror or exc))
    except UnicodeDecodeError as exc:
        raise SpecError("spec %s is not UTF-8 text: %s" % (path, exc))
    try:
        spec = json.loads(text)
    except json.JSONDecodeError as exc:
        raise SpecError("spec %s is not valid JSON: %s (line %d, column %d)"
                        % (path, exc.msg, exc.lineno, exc.colno))
    if not isinstance(spec, dict):
        raise SpecError("spec %s must be a JSON object, got %s"
                        % (path, type(spec).__name__))
    return spec


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Create a .docx from a JSON spec.",
        epilog="See the module docstring (top of this file) for the spec format.")
    ap.add_argument("spec", help="path to JSON spec file")
    ap.add_argument("output", help="path of .docx to write")
    args = ap.parse_args()

    # ---- everything up to the save is guarded: a spec problem exits 2 with
    # one line, and only a genuine builder bug keeps its traceback (exit 1).
    try:
        spec = _load_spec(args.spec)
        validate_spec(spec)
    except SpecError as exc:
        return _input_error(str(exc))

    doc = Document()
    try:
        if spec.get("page"):
            apply_page(doc, spec["page"])
        # Before any spec style: the built-in caption style is bold blue (a Word
        # template default), and a caption is what "Insert Table of Figures"
        # formats, so it would turn every caption in a printed paper blue.
        caption_neutralised = _neutralise_caption_style(doc)
        # Same reason, one family up: the built-in heading/title styles are
        # theme-coloured (blue, and Subtitle also italic), which a printed
        # Chinese assignment does not want.  Before ``add_styles``, so a spec
        # that names a colour still gets it — that is the opt-out.
        headings_neutralised = _neutralise_heading_styles(doc)
        if spec.get("styles"):
            add_styles(doc, spec["styles"])
        if spec.get("header"):
            set_hdrftr_text(doc.sections[0], spec["header"], kind="header")
        if spec.get("footer"):
            set_hdrftr_text(doc.sections[0], spec["footer"], kind="footer")
        blocks = spec.get("blocks", [])
        for i, block in enumerate(blocks):
            # The index is passed so a ``table_caption`` can see whether the
            # very next block is its table — that adjacency is what decides the
            # automatic keep-with-table link, and it cannot be recovered once
            # the blocks have been built.
            add_block(doc, block, blocks=blocks, index=i)
        if spec.get("footer_page_numbers"):
            fpn = spec["footer_page_numbers"]
            add_page_number_field(
                doc.sections[0],
                **(dict(fpn) if isinstance(fpn, dict) else {}))
        dangling = check_crossref_anchors(doc)
    except SpecError as exc:
        return _input_error(str(exc))
    except _BUILD_ERRORS as exc:
        return _input_error("%s: %s" % (type(exc).__name__, _describe(exc)))
    if dangling["missing"]:
        # Writing a REF at a non-existent bookmark produces a field whose
        # cached text is all the reader ever sees and whose click goes
        # nowhere, so this is an error, not a warning.
        print(json.dumps({"ok": False, "output": None,
                          "error": "dangling cross-reference targets",
                          "missing": dangling["missing"],
                          "known_bookmarks": dangling["known"][:12]},
                         ensure_ascii=False))
        return 1
    if dangling["invalid_names"] or dangling["duplicate_names"]:
        # Word merges duplicate bookmarks -- the survivor is the *later* one, so
        # a table-of-contents entry jumps to the wrong chapter while the layout
        # still looks right.  LibreOffice instead renames the second one
        # (``_副本_1``), which kills every PAGEREF that used the original name.
        # Both applications also rewrite a space in a bookmark name to ``_`` and
        # truncate past 40 characters, identically.  All of these are silent:
        # the XML layer looks fine and no application reports anything, so the
        # only place to catch them is here, before the file is written.
        # Measured evidence: work/T3-REPORT.md sections 3.1-3.3.
        print(json.dumps({"ok": False, "output": None,
                          "error": "invalid or duplicate bookmark names",
                          "invalid_names": dangling["invalid_names"],
                          "duplicate_names": dangling["duplicate_names"]},
                         ensure_ascii=False))
        return 1
    try:
        doc.save(args.output)
    except OSError as exc:
        # A path the caller cannot write is an input problem, not a builder bug.
        return _input_error("cannot write %s: %s"
                            % (args.output, exc.strerror or exc))
    # R-11: python-docx writes <w:zoom> without the schema-required w:percent;
    # patch it right after save so pipeline output is valid without a manual
    # second pass (see fix_zoom.py). fix_docx prints CLI-style diagnostics to
    # stdout; this script must keep stdout JSON-only (SKILL.md: "Machine-
    # readable scripts write JSON to stdout; diagnostics go to stderr"), so
    # reroute its stdout to stderr for the duration of the call.
    #
    # The return code is *not* discarded: fix_docx() reports 2 for an input it
    # cannot read or a package it cannot patch, and a build that hands back a
    # document Word will refuse is not a success.  Returning 0 there would tell
    # the caller "ok" for a file that never got its w:percent -- the same class
    # of defect as echoing an --index that was never applied.
    with contextlib.redirect_stdout(sys.stderr):
        zoom_code = fix_docx(Path(args.output))
    if zoom_code:
        return zoom_code
    print(json.dumps({"ok": True, "output": args.output,
                      "blocks": len(spec.get("blocks", [])),
                      "sections": len(doc.sections),
                      "crossrefs": dangling["resolved"]}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
