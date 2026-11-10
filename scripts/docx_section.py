#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Section-level page setup: breaks, page numbering, headers and footers.

Page numbering in Word is a *section* property, never a document property:

  ``<w:pgNumType w:fmt="lowerRoman" w:start="1"/>``

sitting inside one section's ``<w:sectPr>`` restarts the counter and changes
its format for that section and for every later section that does not override
it.  That single element is the mechanism behind the ordinary Chinese report
layout — cover with no number at all, front matter numbered ``i, ii, iii``,
body restarting at ``1``.  One ``PAGE`` field in one footer cannot express it.

Three structural traps this module exists to handle, all of them verified
against python-docx 1.2.0 and Word 16:

1. ``w:sectPr`` is a member of ``CT_SectPr``'s ``xsd:sequence``.  Appended in
   the wrong place (after ``w:cols``, say) Word silently discards the whole
   section setting — no error, just a section that ignores ``w:pgNumType``.
   Every write here goes through ``docx_utils.fix_child_order`` with
   :data:`SECTPR_ORDER`.

2. ``Section.footer`` *inherits* the previous section's footer part when the
   section carries no ``w:footerReference`` of its own, and ``Document.add_section()``
   clones the running ``<w:sectPr>``.  A brand-new section therefore starts out
   sharing the old section's footer part, and ``_get_or_add_definition()`` walks
   all the way back to the first section before it creates anything.  Writing a
   PAGE field without unlinking first edits somebody else's footer.
   :func:`own_hdrftr` forces a private part.

3. The section that ends at a break owns the ``<w:sectPr>`` embedded in the
   *last paragraph of that section*; only the final section's ``<w:sectPr>``
   lives at the end of ``<w:body>``.  :func:`add_section_break` keeps that
   ownership straight (and parks the embedded property on the last real
   paragraph instead of on a fresh empty one).

Subcommands:
  list     每节的页面设置 / pgNumType / 页眉页脚引用
  break    追加一节（并可选配置它）
  pagenum  设置 ``<w:pgNumType>``（fmt / start）
  pagenums 在该节**自己的**页脚里写 PAGE 域
  header   写该节自己的页眉文本
  footer   写该节自己的页脚文本
  cols     设置分栏

Examples:
  docx_section.py list in.docx
  docx_section.py break in.docx --start-type new_page -o out.docx
  docx_section.py pagenum out.docx --section 1 --fmt lowerRoman --start 1
  docx_section.py pagenums out.docx --section 2 --template page_of --align center

Exit codes: 0 success, 1 expected failure (bad section index / format),
2 usage error.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import sys

from docx import Document
from docx.enum.section import (
    WD_HEADER_FOOTER, WD_ORIENTATION, WD_SECTION_START,
)
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm

from docx_utils import fix_child_order, insert_ordered

XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

# CT_SectPr is an xsd:sequence (source: wml.xsd).  ``w:pgNumType`` sits between
# ``w:lnNumType`` and ``w:cols``; anywhere else and Word drops it.
SECTPR_ORDER = [
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type",
    "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType",
    "cols", "formProt", "vAlign", "noEndnote", "titlePg", "textDirection",
    "bidi", "rtlGutter", "docGrid", "printerSettings", "sectPrChange",
]

# ST_NumberFormat values that make sense for a page counter.
PAGE_NUM_FMT = {
    "decimal", "decimalZero", "lowerRoman", "upperRoman", "lowerLetter",
    "upperLetter", "chineseCounting", "chineseLegalSimplified",
    "ideographDigital",
}

START_TYPES = {
    "continuous": WD_SECTION_START.CONTINUOUS,
    "new_column": WD_SECTION_START.NEW_COLUMN,
    "new_page": WD_SECTION_START.NEW_PAGE,
    "even_page": WD_SECTION_START.EVEN_PAGE,
    "odd_page": WD_SECTION_START.ODD_PAGE,
}
# ``newPage`` is what the XML says; accept Word's own spelling too.
START_ALIASES = {"nextpage": "new_page"}

_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_ORIENT = {"portrait": WD_ORIENTATION.PORTRAIT,
           "landscape": WD_ORIENTATION.LANDSCAPE}

# which Section attribute implements each header/footer slot
_HF_SLOT = {
    "default": ("footer", "header"),
    "first": ("first_page_footer", "first_page_header"),
    "even": ("even_page_footer", "even_page_header"),
}
_HF_TYPE = {None: "default", "": "default", "default": "default",
            "first": "first", "first_page": "first",
            "even": "even", "even_page": "even"}


# ---- page-counter rendering (for the cached field result) -------------------

_ROMAN = ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
          (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
          (5, "v"), (4, "iv"), (1, "i"))


def roman(n: int, upper: bool = False) -> str:
    """``3`` -> ``'iii'``.  Values below 1 are passed through as text."""
    n = int(n)
    if n < 1:
        return str(n)
    out = []
    for val, sym in _ROMAN:
        while n >= val:
            out.append(sym)
            n -= val
    txt = "".join(out)
    return txt.upper() if upper else txt


def letters(n: int, upper: bool = False) -> str:
    """``1`` -> ``'a'``, ``27`` -> ``'aa'`` (Word's lowerLetter sequence)."""
    n = int(n)
    if n < 1:
        return str(n)
    out = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out.append(chr(ord("a") + rem))
    txt = "".join(reversed(out))
    return txt.upper() if upper else txt


def _fullwidth_digits(n: int) -> str:
    """``3`` -> ``'３'`` (``decimalFullWidth``)."""
    return "".join(chr(0xFF10 + int(d)) for d in str(int(n)))


_CN_DIGITS = "〇一二三四五六七八九"
_CN_UNITS = ("", "十", "百", "千")
_CN_BIG = ("", "万", "亿")


def _chinese_number(n: int) -> str:
    """``3`` -> ``'三'``, ``12`` -> ``'十二'`` (``chineseCounting``/ideograph)."""
    n = int(n)
    if n < 0:
        return str(n)
    if n == 0:
        return _CN_DIGITS[0]
    if n < 10:
        return _CN_DIGITS[n]
    groups = []
    rest = n
    while rest > 0:
        groups.append(rest % 10000)
        rest //= 10000
    out = []
    for gi in range(len(groups) - 1, -1, -1):
        g = groups[gi]
        if g == 0:
            continue
        part = []
        zero_pending = False
        for pos in range(3, -1, -1):
            d = (g // (10 ** pos)) % 10
            if d == 0:
                if part:
                    zero_pending = True
                continue
            if zero_pending:
                part.append(_CN_DIGITS[0])
                zero_pending = False
            part.append(_CN_DIGITS[d] + _CN_UNITS[pos])
        chunk = "".join(part)
        # 十一 is written without the leading 一; 一十一 is wrong in a page number.
        if gi == 0 and len(groups) == 1 and chunk.startswith("一十"):
            chunk = chunk[1:]
        out.append(chunk + _CN_BIG[gi])
    return "".join(out)


_EN_ONES = ("", "one", "two", "three", "four", "five", "six", "seven", "eight",
            "nine", "ten", "eleven", "twelve", "thirteen", "fourteen",
            "fifteen", "sixteen", "seventeen", "eighteen", "nineteen")
_EN_TENS = ("", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy",
            "eighty", "ninety")
_EN_ORDINAL_IRREGULAR = {
    "one": "first", "two": "second", "three": "third", "five": "fifth",
    "eight": "eighth", "nine": "ninth", "twelve": "twelfth",
}


def _english_ordinal(n: int) -> str:
    """``3`` -> ``'third'`` (``ordinalText``), the way Word renders it."""
    n = int(n)
    if n <= 0:
        return str(n)
    if n < 20:
        word = _EN_ONES[n]
    elif n < 100:
        tens, ones = divmod(n, 10)
        word = _EN_TENS[tens] + ("-" + _EN_ONES[ones] if ones else "")
    elif n < 1000:
        hundreds, rest = divmod(n, 100)
        word = _EN_ONES[hundreds] + " hundred"
        if rest:
            word += " and " + _english_ordinal(rest)
            return word
    else:
        # Page numbers past 999 are not worth a full English generator; falling
        # back to the numeral is honest and never matches a real Word glyph, so
        # the caller can tell the difference.
        return str(n)
    tail = word.rsplit("-", 1)[-1]
    ordinal_tail = _EN_ORDINAL_IRREGULAR.get(
        tail, tail[:-1] + "ieth" if tail.endswith("y") else tail + "th")
    if "-" in word:
        head = word.rsplit("-", 1)[0]
        return head + "-" + ordinal_tail
    return ordinal_tail


def cached_page(fmt: str | None, start=None) -> str:
    """The text the PAGE field shows before Word recalculates it.

    Writing the right glyph here means the file renders correctly even for a
    reader that never updates fields (and makes PDF export verifiable without
    forcing an update first).

    Every branch below is a **measured** Word 16.0 rendering, not a guess: a
    document per format, a ``PAGE`` field in the footer, ``Fields.Update()``,
    then read back the rendered text (work/REVIEW-ADVERSARIAL-R6.md 阻塞-3).
    Formats Word can render but this function does not model would otherwise
    fall through to ``str(n)`` -- silently claiming an Arabic numeral for a page
    that shows ``三`` or nothing at all -- so the fall-through is kept
    deliberately narrow and the two cases where the true glyph is *not* the
    number are handled explicitly:

    * ``none`` renders an **empty string**; there is no page number on the page;
    * ``numberInDash`` wraps it as ``- 3 -``.
    """
    n = 1 if start in (None, "") else int(start)
    if fmt == "none":
        return ""
    if fmt == "decimalFullWidth":
        return _fullwidth_digits(n)
    if fmt in ("chineseCounting", "chineseCountingThousand",
               "ideographDigital"):
        return _chinese_number(n)
    if fmt == "numberInDash":
        return "- %d -" % n
    if fmt == "decimalEnclosedCircle":
        # ①..⑳ exist as single code points; past that Word keeps enclosing, and
        # the composed form is what it renders.
        return chr(0x2460 + n - 1) if 1 <= n <= 20 else str(n)
    if fmt == "ordinalText":
        return _english_ordinal(n)
    if fmt == "lowerRoman":
        return roman(n)
    if fmt == "upperRoman":
        return roman(n, True)
    if fmt == "lowerLetter":
        return letters(n)
    if fmt == "upperLetter":
        return letters(n, True)
    if fmt == "decimalZero":
        return "%02d" % n
    return str(n)


def _norm_start_type(name):
    if name is None:
        return WD_SECTION_START.NEW_PAGE
    key = str(name).strip().lower()
    key = START_ALIASES.get(key, key)
    if key not in START_TYPES:
        raise SystemExit("unknown start_type %r; choose from %s"
                         % (name, ", ".join(sorted(START_TYPES))))
    return START_TYPES[key]


_TYPE_XML_NAMES = {"continuous": "continuous", "nextColumn": "new_column",
                   "nextPage": "new_page", "evenPage": "even_page",
                   "oddPage": "odd_page"}


def _start_type_name(sectPr) -> str:
    el = sectPr.find(qn("w:type"))
    if el is None:
        return "new_page"          # CT_SectType defaults to nextPage
    val = el.get(qn("w:val")) or "nextPage"
    return _TYPE_XML_NAMES.get(val, val)


# ---- section access ---------------------------------------------------------

def sections_of(doc) -> list:
    return list(doc.sections)


def resolve_index(doc, index: int) -> int:
    """Normalise a possibly-negative section index, or fail loudly."""
    secs = sections_of(doc)
    if not secs:
        raise SystemExit("document has no sections")
    i = int(index)
    if not -len(secs) <= i < len(secs):
        raise SystemExit("no such section %d (document has %d)"
                         % (index, len(secs)))
    return i % len(secs)


def get_section(doc, index: int):
    """Section by index; negative indexes count from the end."""
    return sections_of(doc)[resolve_index(doc, index)]


# ---- page size / margins ----------------------------------------------------

def apply_page_to_section(section, page: dict) -> dict:
    """Set paper size, orientation and margins on *section*.

    Unlike ``docx_create.apply_page`` this takes a section, so the same dict
    can configure any section of a multi-section document.  Omitted keys are
    left alone, which matters because ``add_section()`` clones the running
    section — a section that says nothing inherits the previous one.
    """
    page = page or {}
    if page.get("width_mm"):
        section.page_width = Mm(page["width_mm"])
    if page.get("height_mm"):
        section.page_height = Mm(page["height_mm"])
    if page.get("orient"):
        key = str(page["orient"]).lower()
        if key not in _ORIENT:
            raise SystemExit("unknown orient %r; choose portrait or landscape"
                             % (page["orient"],))
        section.orientation = _ORIENT[key]
    m = page.get("margins_mm") or {}
    for side in ("top", "bottom", "left", "right"):
        if side in m:
            setattr(section, "%s_margin" % side, Mm(m[side]))
    for side in ("header", "footer", "gutter"):
        if side in m:
            setattr(section, "%s_distance" % side, Mm(m[side]))
    return {"width_mm": page.get("width_mm"),
            "height_mm": page.get("height_mm"),
            "margins_mm": dict(m)}


# ---- page numbering ---------------------------------------------------------

def set_pg_num_type(section, fmt: str | None = None, start=None) -> dict:
    """Write (or update) ``<w:pgNumType>`` in this section's ``<w:sectPr>``.

    ``fmt`` is an ``ST_NumberFormat`` value such as ``lowerRoman``; ``start``
    restarts the counter at that page.  Both are optional; passing neither
    still guarantees the element exists and is schema-ordered.
    """
    if fmt is not None and fmt not in PAGE_NUM_FMT:
        raise SystemExit("unknown page number fmt %r; choose from %s"
                         % (fmt, ", ".join(sorted(PAGE_NUM_FMT))))
    sectPr = section._sectPr
    el = sectPr.find(qn("w:pgNumType"))
    if el is None:
        el = OxmlElement("w:pgNumType")
        insert_ordered(sectPr, el, SECTPR_ORDER)
    if fmt is not None:
        el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), str(int(start)))
    fix_child_order(sectPr, SECTPR_ORDER)
    return {"fmt": el.get(qn("w:fmt")), "start": el.get(qn("w:start"))}


def clear_pg_num_type(section) -> bool:
    """Drop ``<w:pgNumType>`` so the section continues the previous counter."""
    sectPr = section._sectPr
    el = sectPr.find(qn("w:pgNumType"))
    if el is None:
        return False
    sectPr.remove(el)
    return True


def get_pg_num_type(section) -> dict | None:
    el = section._sectPr.find(qn("w:pgNumType"))
    if el is None:
        return None
    return {"fmt": el.get(qn("w:fmt")), "start": el.get(qn("w:start"))}


def set_columns(section, num: int, space_twips=425) -> dict:
    """``<w:cols w:num="N" w:space="..."/>`` on this section.

    ``w:space`` defaults to 425 twips, the value Word itself writes — a
    ``w:cols`` with no spacing renders columns glued together.
    """
    n = int(num)
    if n < 1:
        raise SystemExit("--num must be >= 1")
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        insert_ordered(sectPr, cols, SECTPR_ORDER)
    if n > 1:
        cols.set(qn("w:num"), str(n))
    elif cols.get(qn("w:num")) is not None:
        del cols.attrib[qn("w:num")]
    if space_twips is not None:
        cols.set(qn("w:space"), str(int(space_twips)))
    fix_child_order(sectPr, SECTPR_ORDER)
    return {"num": n, "space": cols.get(qn("w:space"))}


# ---- headers / footers ------------------------------------------------------

def _ref_id(sectPr, kind: str, page_type: str):
    """rId of this section's own ``w:headerReference``/``w:footerReference``."""
    if sectPr is None:
        return None
    tag = "w:footerReference" if kind == "footer" else "w:headerReference"
    for el in sectPr.findall(qn(tag)):
        if (el.get(qn("w:type")) or "default") == page_type:
            return el.get(RID)
    return None


def own_hdrftr(section, kind: str = "footer", page_type: str = "default"):
    """Return *section*'s **private** header or footer proxy.

    Three things make this necessary rather than just calling ``section.footer``:

    * a section with no ``w:footerReference`` inherits the previous section's
      part, and python-docx resolves that lazily by walking back to the first
      section — so writing text could edit an earlier section's footer;
    * ``Document.add_section()`` clones the running ``<w:sectPr>``, so a brand
      new section starts out holding the *same* rId the previous section holds;
    * a definition this section already owns must be preserved — calling this
      twice (footer text, then a PAGE field) has to append, not wipe.

    The first case is detected by comparing the reference with the previous
    section's: equal rIds mean the pair came from one clone. Assigning
    ``is_linked_to_previous = True`` then drops only this section's reference,
    and ``Part.drop_rel()`` refuses to delete a relationship that is still
    referenced elsewhere, so a shared part survives for its other owner.
    """
    pt = _HF_TYPE.get(page_type)
    if pt is None:
        raise SystemExit("unknown page_type %r; choose default, first, even"
                         % (page_type,))
    if kind not in ("header", "footer"):
        raise SystemExit("kind must be header or footer")
    attr = _HF_SLOT[pt][0 if kind == "footer" else 1]
    hf = getattr(section, attr)
    sectPr = section._sectPr
    r_id = _ref_id(sectPr, kind, pt)
    if r_id is not None and r_id == _ref_id(sectPr.preceding_sectPr, kind, pt):
        hf.is_linked_to_previous = True          # this part is not ours
    if hf.is_linked_to_previous:
        hf.is_linked_to_previous = False         # give us our own blank part
    return getattr(section, attr)


def _clear_para(para) -> None:
    """Strip every child of a paragraph except its ``<w:pPr>``."""
    p = para._p
    for child in list(p):
        if child.tag != qn("w:pPr"):
            p.remove(child)


def _rpr(font=None, size_pt=None, bold=None):
    """A ``<w:rPr>`` for the field runs, or None when nothing was asked for.

    ``EG_RPrBase`` is an ``xsd:choice`` with ``maxOccurs="unbounded"``, so run
    properties may appear in any order — no sequencing work needed here.
    """
    rpr = OxmlElement("w:rPr")
    if font:
        fonts = OxmlElement("w:rFonts")
        for a in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn(a), font)
        rpr.append(fonts)
    if bold:
        for tag in ("b", "bCs"):
            e = OxmlElement(tag)
            e.set(qn("val"), "1")
            rpr.append(e)
    if size_pt:
        for tag in ("sz", "szCs"):
            e = OxmlElement(tag)
            e.set(qn("val"), str(int(round(float(size_pt) * 2))))
            rpr.append(e)
    return rpr if len(rpr) else None


def _field_runs(instr: str, cached: str, rpr=None) -> list:
    """A complex field: begin / instrText / separate / result / end."""
    from copy import deepcopy

    def _wrap(child):
        r = OxmlElement("w:r")
        if rpr is not None:
            r.append(deepcopy(rpr))
        r.append(child)
        return r

    def _fld(kind):
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        return _wrap(f)

    it = OxmlElement("w:instrText")
    it.set(XMLSPACE, "preserve")
    it.text = " %s " % instr.strip()
    t = OxmlElement("w:t")
    t.set(XMLSPACE, "preserve")
    t.text = str(cached)
    return [_fld("begin"), _wrap(it), _fld("separate"), _wrap(t), _fld("end")]


def set_hdrftr_text(section, text: str, kind: str = "footer",
                    page_type: str = "default",
                    align: str | None = None, font=None, size_pt=None,
                    bold=None) -> dict:
    """Replace a header/footer with *text* (``\\n`` splits paragraphs)."""
    hf = own_hdrftr(section, kind, page_type)
    paras = list(hf.paragraphs)
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        para = paras[i] if i < len(paras) else hf.add_paragraph()
        _clear_para(para)
        if align:
            para.alignment = _align(align)
        run = para.add_run(line)
        if font:
            run.font.name = font
        if size_pt:
            run.font.size = _pt(size_pt)
        if bold is not None:
            run.font.bold = bool(bold)
    for extra in paras[len(lines):]:
        extra._p.getparent().remove(extra._p)
    return {"kind": kind, "page_type": page_type, "lines": len(lines)}


def _align(name):
    if name is None:
        return None
    key = str(name).lower()
    if key not in _ALIGN:
        raise SystemExit("unknown alignment %r; choose %s"
                         % (name, ", ".join(sorted(set(_ALIGN)))))
    return _ALIGN[key]


def _pt(value):
    from docx.shared import Pt
    return Pt(value)


def add_page_number_field(section, template="page_of", fmt=None, start=None,
                          align="center", total=None, prefix=None,
                          suffix=None, font=None, size_pt=None, bold=None,
                          page_type: str = "default") -> dict:
    """Put a PAGE field in *section*'s **own** footer.

    ``template``: ``page`` (just the counter), ``page_of`` (``Page X of Y``),
    ``dash`` (``- X -``), or ``custom`` (use ``prefix``/``suffix`` verbatim).
    ``total``: ``numpages`` (whole document, the default) or ``sectionpages``.

    The counter's *format* and *restart point* come from this section's
    ``w:pgNumType``, not from the field, so ``fmt``/``start`` are applied to
    the section first and the field's cached result is rendered to match.
    """
    if fmt is not None or start is not None:
        set_pg_num_type(section, fmt=fmt, start=start)
    eff = get_pg_num_type(section) or {}
    eff_fmt, eff_start = eff.get("fmt"), eff.get("start")
    cached = cached_page(eff_fmt, eff_start)

    tpl = "page_of" if template is True else (template or "page_of")
    total_field = "SECTIONPAGES" if str(total).lower() == "sectionpages" \
        else "NUMPAGES"

    hf = own_hdrftr(section, "footer", page_type)
    paras = list(hf.paragraphs)
    para = paras[0] if paras else hf.add_paragraph()
    if align is not None:
        para.alignment = _align(align)
    rpr = _rpr(font, size_pt, bold)

    if tpl == "page":
        plan = [("f", "PAGE", cached)]
    elif tpl == "dash":
        plan = [("t", "- ", None), ("f", "PAGE", cached), ("t", " -", None)]
    elif tpl == "custom":
        plan = [("t", prefix or "", None), ("f", "PAGE", cached),
                ("t", suffix or "", None)]
    elif tpl == "page_of":
        plan = [("t", prefix if prefix is not None else "Page ", None),
                ("f", "PAGE", cached),
                ("t", suffix if suffix is not None else " of ", None),
                ("f", total_field, "1")]
    else:
        raise SystemExit("unknown template %r; choose page, page_of, dash, "
                         "custom" % (template,))

    for kind, value, cached_val in plan:
        if kind == "t":
            if not value:
                continue
            run = para.add_run(value)
            if font:
                run.font.name = font
            if size_pt:
                run.font.size = _pt(size_pt)
            if bold is not None:
                run.font.bold = bool(bold)
        else:
            for el in _field_runs(value, cached_val, rpr):
                para._p.append(el)

    return {"page_numbers": True, "template": tpl, "fmt": eff_fmt,
            "start": eff_start, "cached": cached, "total_field": total_field,
            "page_type": page_type}


# ---- breaks -----------------------------------------------------------------

def add_section_break(doc, start_type="new_page", reuse_last: bool = True):
    """Append a section and return it.

    ``Document.add_section()`` leaves an extra empty paragraph behind: it
    clones the running ``<w:sectPr>`` (that clone becomes the *old* section's
    property, which is how the old section keeps its own numbering) and hangs
    it on a fresh paragraph — and that paragraph carries no run text, so it
    renders as a stray blank line at the end of the closed section.

    Word itself stores the property on the last paragraph that actually has
    content, so ``reuse_last`` moves it there and deletes the empty carrier.
    It only does so when the carrier is provably the last element before the
    body-level ``<w:sectPr>`` and the element before *that* is a paragraph —
    otherwise a table at the end of the section would be silently pushed into
    the next section.
    """
    section = doc.add_section(_norm_start_type(start_type))
    # ``Document.add_section()`` clones the running ``<w:sectPr>`` wholesale,
    # ``w:pgNumType`` included — so the new section silently inherits the
    # previous section's *restart point* and its counter begins again at that
    # number.  A section that was meant to keep counting (the usual case, e.g.
    # body pages continuing past a roman-numbered TOC) then restarts at 1.
    # Drop the inherited element: the caller re-adds one explicitly via
    # ``configure_section`` when it really wants a format or a restart.
    _drop_inherited_pg_num_type(section)
    moved = False
    if reuse_last:
        moved = _relocate_sectpr(doc, section)
    return section, moved


def _drop_inherited_pg_num_type(section) -> bool:
    """Strip the ``w:pgNumType`` that ``add_section()`` copied from the parent.

    Only clears it when the clone actually carries one, so an explicit
    ``pgNumType`` written later by the caller is never touched.
    """
    sectPr = section._sectPr
    el = sectPr.find(qn("w:pgNumType"))
    if el is None:
        return False
    sectPr.remove(el)
    return True


def _relocate_sectpr(doc, section) -> bool:
    body = doc.element.body
    children = list(body)
    if len(children) < 3:
        return False
    if children[-1].tag != qn("w:sectPr"):
        return False
    if not _is_body_sectPr(doc, section._sectPr):
        return False        # <w:body> holds exactly one sectPr; it is this one
    carrier = children[-2]
    if carrier.tag != qn("w:p"):
        return False
    pPr = carrier.find(qn("w:pPr"))
    if pPr is None:
        return False
    clone = pPr.find(qn("w:sectPr"))
    if clone is None:
        return False
    if [c for c in carrier if c.tag != qn("w:pPr")]:
        return False                      # carrier holds real content
    prev = children[-3]
    if prev.tag != qn("w:p"):
        return False
    prev_pPr = prev.find(qn("w:pPr"))
    if prev_pPr is not None and prev_pPr.find(qn("w:sectPr")) is not None:
        return False
    pPr.remove(clone)
    insert_ordered(prev.get_or_add_pPr(), clone, "pPr")
    body.remove(carrier)
    return True


def configure_section(section, spec: dict) -> dict:
    """Apply a ``section_break`` block's payload to *section*.

    Recognised keys: ``page``, ``cols``, ``different_first_page``, ``header``,
    ``header_align``, ``footer``, ``footer_page_numbers``.  Anything absent is
    left inheriting from the previous section, which is what makes a
    three-section cover/TOC/body document cheap to describe.
    """
    spec = spec or {}
    out: dict = {}
    if spec.get("page"):
        out["page"] = apply_page_to_section(section, spec["page"])
    if spec.get("cols"):
        out["cols"] = set_columns(section, spec["cols"])
    if spec.get("different_first_page") is not None:
        section.different_first_page_header_footer = \
            bool(spec["different_first_page"])
        out["different_first_page"] = bool(spec["different_first_page"])
    if spec.get("header") is not None:
        out["header"] = set_hdrftr_text(
            section, spec["header"], kind="header",
            align=spec.get("header_align"),
            font=spec.get("header_font"), size_pt=spec.get("header_size_pt"),
            bold=spec.get("header_bold"))
    elif spec.get("header_align"):
        hf = own_hdrftr(section, "header")
        hf.paragraphs[0].alignment = _align(spec["header_align"])
        out["header"] = {"kind": "header", "lines": 0}
    if spec.get("footer") is not None:
        out["footer"] = set_hdrftr_text(
            section, spec["footer"], kind="footer", align=spec.get("footer_align"),
            font=spec.get("footer_font"), size_pt=spec.get("footer_size_pt"),
            bold=spec.get("footer_bold"))
    fpn = spec.get("footer_page_numbers")
    if fpn:
        kw = dict(fpn) if isinstance(fpn, dict) else {}
        kw.setdefault("align", spec.get("footer_align"))
        out["page_numbers"] = add_page_number_field(section, **kw)
    elif spec.get("no_footer"):
        # explicit "this section has no footer of its own" — needed when the
        # previous section defined one and inheritance would show it here.
        hf = own_hdrftr(section, "footer")
        _clear_para(hf.paragraphs[0])
        out["footer"] = {"kind": "footer", "lines": 0, "cleared": True}
    return out


# ---- reporting --------------------------------------------------------------

RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _hf_refs(sectPr) -> list:
    out = []
    for tag, kind in (("w:headerReference", "header"),
                      ("w:footerReference", "footer")):
        for el in sectPr.findall(qn(tag)):
            out.append({"kind": kind,
                        "type": el.get(qn("w:type")) or "default",
                        "r_id": el.get(RID)})
    return out


def _hf_preview(doc, ref: dict) -> str:
    """First N characters of the referenced header/footer part, or ''."""
    r_id = ref.get("r_id")
    if not r_id:
        return ""
    try:
        part = doc.part.related_parts[r_id]
    except KeyError:
        return ""
    el = part.element
    texts = [t.text or "" for t in el.iter(qn("w:t"))]
    fields = [i.text.strip() for i in el.iter(qn("w:instrText"))]
    body = "".join(texts)
    if fields:
        body = (body + " " + " ".join("[%s]" % f for f in fields)).strip()
    return body[:60]


def section_info(doc, index: int, section) -> dict:
    sectPr = section._sectPr
    pg = get_pg_num_type(section)
    cols = sectPr.find(qn("w:cols"))
    num = cols.get(qn("w:num")) if cols is not None else None
    refs = _hf_refs(sectPr)
    for r in refs:
        r["preview"] = _hf_preview(doc, r)
    return {
        "index": index,
        "start_type": _start_type_name(sectPr),
        "page_width_mm": _mm(section.page_width),
        "page_height_mm": _mm(section.page_height),
        "margins_mm": {k: _mm(getattr(section, "%s_margin" % k))
                       for k in ("top", "bottom", "left", "right")},
        "page_number": pg,
        "cols": int(num) if num else 1,
        "different_first_page": sectPr.find(qn("w:titlePg")) is not None,
        "sectPr_location": "body" if _is_body_sectPr(doc, sectPr)
        else "paragraph",
        "header_footer": refs,
    }


def _is_body_sectPr(doc, sectPr) -> bool:
    """True when this sectPr is the body-level one (i.e. the final section).

    Every other section's property sits in the ``<w:pPr>`` of the paragraph it
    terminates.  Tested via the parent tag rather than element identity,
    because lxml hands out wrapper objects whose ``is`` comparison is not a
    contract.
    """
    parent = sectPr.getparent()
    return parent is not None and parent.tag == qn("w:body")


def _mm(value) -> float | None:
    if value is None:
        return None
    return round(value.mm, 2)


def cmd_list(doc) -> dict:
    secs = sections_of(doc)
    return {"ok": True, "sections": len(secs),
            "section_detail": [section_info(doc, i, s)
                               for i, s in enumerate(secs)]}


# ---- CLI --------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Section-level page setup for .docx.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("path", help="input .docx")
        p.add_argument("-o", "--output",
                       help="output path (default: overwrite input)")

    def pick(p):
        p.add_argument("--section", type=int, default=-1,
                       help="section index (default -1 = last)")

    p = sub.add_parser("list", help="show every section's setup")
    common(p)

    p = sub.add_parser("break", help="append a section")
    common(p)
    p.add_argument("--start-type", default="new_page",
                   choices=sorted(START_TYPES))
    p.add_argument("--no-reuse-last", action="store_true",
                   help="keep python-docx's empty carrier paragraph")

    p = sub.add_parser("pagenum", help="set <w:pgNumType> on a section")
    common(p)
    pick(p)
    p.add_argument("--fmt", choices=sorted(PAGE_NUM_FMT),
                   help="decimal | lowerRoman | upperRoman | ...")
    p.add_argument("--start", type=int, help="restart the counter here")
    p.add_argument("--clear", action="store_true",
                   help="remove pgNumType (continue the previous counter)")

    p = sub.add_parser("pagenums", help="write PAGE fields in a section footer")
    common(p)
    pick(p)
    p.add_argument("--template", default="page_of",
                   choices=["page", "page_of", "dash", "custom"])
    p.add_argument("--fmt", choices=sorted(PAGE_NUM_FMT))
    p.add_argument("--start", type=int)
    p.add_argument("--align", default="center",
                   choices=sorted(set(_ALIGN)))
    p.add_argument("--total", default="numpages",
                   choices=["numpages", "sectionpages"])
    p.add_argument("--prefix")
    p.add_argument("--suffix")
    p.add_argument("--page-type", default="default",
                   choices=["default", "first", "even"])

    p = sub.add_parser("header", help="write a section's own header text")
    common(p)
    pick(p)
    p.add_argument("--text", required=True)
    p.add_argument("--align", choices=sorted(set(_ALIGN)))
    p.add_argument("--page-type", default="default",
                   choices=["default", "first", "even"])

    p = sub.add_parser("footer", help="write a section's own footer text")
    common(p)
    pick(p)
    p.add_argument("--text", required=True)
    p.add_argument("--align", choices=sorted(set(_ALIGN)))
    p.add_argument("--page-type", default="default",
                   choices=["default", "first", "even"])

    p = sub.add_parser("cols", help="set the column count")
    common(p)
    pick(p)
    p.add_argument("--num", type=int, required=True)
    p.add_argument("--space", type=int, default=425, help="w:space in twips")

    args = ap.parse_args()
    doc = Document(args.path)

    if args.cmd == "list":
        print(json.dumps(cmd_list(doc), ensure_ascii=False))
        return 0

    if args.cmd == "break":
        section, moved = add_section_break(
            doc, args.start_type, reuse_last=not args.no_reuse_last)
        result = {"ok": True, "sections": len(sections_of(doc)),
                  "start_type": args.start_type, "reused_last_paragraph": moved}
    else:
        idx = resolve_index(doc, args.section)
        section = sections_of(doc)[idx]
        if args.cmd == "pagenum":
            if args.clear:
                result = {"ok": True, "cleared": clear_pg_num_type(section)}
            else:
                result = set_pg_num_type(section, fmt=args.fmt, start=args.start)
                result["ok"] = True
        elif args.cmd == "pagenums":
            result = add_page_number_field(
                section, template=args.template, fmt=args.fmt, start=args.start,
                align=args.align, total=args.total, prefix=args.prefix,
                suffix=args.suffix, page_type=args.page_type)
            result["ok"] = True
        elif args.cmd == "cols":
            result = set_columns(section, args.num, args.space)
            result["ok"] = True
        else:
            result = set_hdrftr_text(
                section, args.text,
                kind="header" if args.cmd == "header" else "footer",
                page_type=args.page_type, align=args.align)
            result["ok"] = True
        result["section"] = idx

    out = args.output or args.path
    doc.save(out)
    result["output"] = out
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
