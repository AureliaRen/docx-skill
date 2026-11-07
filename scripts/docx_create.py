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
              "size_pt": 12, "bold": true, "color": "1F4E79"}],
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

``crossref`` passes ``superscript`` (bool), ``vert_align``
(``normal``/``superscript``/``subscript``) and ``hyperlink`` (bool, default
true) through to ``docx_crossref.cmd_ref``.

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
``start``, ``template`` (``page``/``page_of``/``dash``/``custom``), ``total``,
``align``, ``prefix``, ``suffix``.

Extras: `"footer_page_numbers": true` at the top level adds a
"Page X of Y" footer built from PAGE/NUMPAGES fields to the first section,
and a `toc` block inserts a Table of Contents field. Field results are
computed by Word/LibreOffice when the file is opened, not by python-docx;
every field is written with a cached result so the file reads correctly
before anyone recalculates.
"""
from __future__ import annotations
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix as _encoding_fix  # noqa: F401

import argparse
import json
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING,
                            WD_TAB_ALIGNMENT, WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor, Twips

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
    for attr, value in (("ascii", latin), ("hAnsi", latin),
                        ("eastAsia", east_asia or latin),
                        ("cs", latin)):
        rf.set(qn("w:" + attr), value)


def add_styles(doc, styles: list) -> None:
    for s in styles:
        style = doc.styles.add_style(s["name"], WD_STYLE_TYPE.PARAGRAPH)
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
            font.color.rgb = RGBColor.from_string(s["color"])


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


def add_block(doc, block: dict) -> list:
    """Add one block; return the body paragraphs it created or attached to.

    Every paragraph-producing block then runs through
    :func:`_apply_para_format`, so the format vocabulary is uniform instead of
    being re-implemented per block type.
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
    elif btype == "image":
        width = Mm(block["width_mm"]) if block.get("width_mm") else None
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


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Create a .docx from a JSON spec.",
        epilog="See the module docstring (top of this file) for the spec format.")
    ap.add_argument("spec", help="path to JSON spec file")
    ap.add_argument("output", help="path of .docx to write")
    args = ap.parse_args()

    with open(args.spec, encoding="utf-8") as f:
        spec = json.load(f)

    doc = Document()
    if spec.get("page"):
        apply_page(doc, spec["page"])
    # Before any spec style: the built-in caption style is bold blue (a Word
    # template default), and a caption is what "Insert Table of Figures"
    # formats, so it would turn every caption in a printed paper blue.
    caption_neutralised = _neutralise_caption_style(doc)
    if spec.get("styles"):
        add_styles(doc, spec["styles"])
    if spec.get("header"):
        set_hdrftr_text(doc.sections[0], spec["header"], kind="header")
    if spec.get("footer"):
        set_hdrftr_text(doc.sections[0], spec["footer"], kind="footer")
    for block in spec.get("blocks", []):
        add_block(doc, block)
    if spec.get("footer_page_numbers"):
        fpn = spec["footer_page_numbers"]
        add_page_number_field(doc.sections[0],
                              **(dict(fpn) if isinstance(fpn, dict) else {}))
    dangling = check_crossref_anchors(doc)
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
    doc.save(args.output)
    print(json.dumps({"ok": True, "output": args.output,
                      "blocks": len(spec.get("blocks", [])),
                      "sections": len(doc.sections),
                      "crossrefs": dangling["resolved"]}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
