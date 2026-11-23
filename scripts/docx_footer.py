#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Footer health: what shape a footer is, how wide its box is, and whether the
page number actually fits — as JSON on stdout.

本文件存在的理由
----------------

`docx_validate` 回答的是「这个包是不是良构的」，而一份**页脚被塞进 4.56 pt 宽
浮动文本框**的文档，它的每一项检查都过：XML 合法、部件解析得到、`PAGE` 也是真
域。可页码一旦进位到两位就被裁掉，而没有任何一个检查说得出这件事（实测：主人
的真实排版作业里，`docx_validate` 对这类文档返 ``{"ok": true, "issues": []}``）。

`docx_section list` 同样说不出：它只暴露 ``header_footer[].preview`` 一个拍平
的字符串，例如 ``第 1 页　共 1 页 [PAGE] [NUMPAGES]`` —— 这串东西在「整行排在
版心里」和「整行排在一个 4.56 pt 的框里」两种情况下**完全一样**。

所以这里另开一条只读通道：读页脚部件的原始 XML，回答三件事

1. ``shape`` —— 这个页脚是普通段落（``plain``）、内容整个落在文本框里
   （``floating-textbox``），还是框内框外都有东西（``mixed``）；
2. ``box_width_pt`` —— 框到底多宽（``wp:extent/@cx`` 的 EMU ÷ 12700，或 VML
   ``style="width:4.56pt"``），``4.56`` 这种数如实报，不四舍五入成好看的数；
3. ``page_number_width_pt`` —— 按页脚里**实际**的字体与字号，把 ``第 1 页`` /
   ``第 12 页`` / ``第 123 页`` 这几种形态渲染出来各要多少 pt，再与框宽比。

宽度是**估算**，不是字体度量
----------------------------

本脚本不读 TTF/OTF 的 ``hmtx`` 表，也不调 Word / LibreOffice 排版。估算模型明
白写在这里，免得把估算值当实测值引用：

* 全角字符（``unicodedata.east_asian_width`` 为 ``W``/``F``，含汉字、全角空格
  U+3000、全角标点）= **1.0 em**；
* ASCII 数字与字母 = **0.5 em**（Times New Roman / 宋体的数字都是半角 0.5 em，
  这正是这几种形态里最要紧的一档）；
* ASCII 空格 = **0.25 em**；其余半角标点 = **0.3 em**。

乘上该段文字自己的字号（``w:sz`` 语义：半磅，÷2）得到 pt。误差方向是**低估**：
真实字体里数字多为整 0.5 em，所以「估算不超框」偏乐观 —— 一旦估算值已经超框，
实际一定超。这也是本脚本只在证据充分时给 ``error`` 的原因。

每个输出里都带 ``estimate_model``，引用这个数时要一起引它。

退出码（D-030 口径）
--------------------

* ``0`` —— 检查跑完，没有 error 级问题（warning 不改变退出码）；
* ``1`` —— 有 error 级问题。这是 ``check`` 的**回答**，不是工具被误用；
* ``2`` —— 用法错误：路径缺失 / 是目录 / 不是 Word 包 / ``--section`` 越界。
  此时**零写盘、stdout 为空、stderr 恰好一行 ``error: …``**。

``check`` 是只读命令：本文件从头到尾没有一处写文件（连临时文件都没有）。

Examples:
  python scripts/docx_footer.py check in.docx
  python scripts/docx_footer.py check in.docx --section 2
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import posixpath
import sys
import unicodedata
import zipfile

from lxml import etree

from docx_read import InputError, _check_input

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
V = "urn:schemas-microsoft-com:vml"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

W_SECTPR = "{%s}sectPr" % W
W_FOOTERREF = "{%s}footerReference" % W
W_PGNUMTYPE = "{%s}pgNumType" % W

#: 1 pt = 12700 EMU (English Metric Unit).  ``wp:extent/@cx`` is in EMU, so a
#: 4.56 pt box is ``cx="57912"``.  Named rather than a bare 12700 in the
#: arithmetic, so the conversion is reviewable.
EMU_PER_PT = 12700.0

#: ``NUMPAGES`` and ``SECTIONPAGES`` both render a page count, so both count as
#: "a total field" when sizing the line.
TOTAL_FIELDS = ("NUMPAGES", "SECTIONPAGES")

#: ``wp:`` children that describe wrapping.  Read off ``wp:anchor`` rather than
#: the ``wrapType`` attribute because Word 2010+ writes the element form.
WRAP_ELEMENTS = ("wrapNone", "wrapSquare", "wrapTight", "wrapThrough",
                 "wrapTopAndBottom")

ESTIMATE_MODEL = ("full-width chars 1.0 em, ASCII digits/letters 0.5 em, "
                  "space 0.25 em, other printable half-width 0.3 em; "
                  "each segment multiplied by its own size (w:sz / 2). "
                  "An estimate, not a font-metric read.")


# ---- small XML helpers -------------------------------------------------------

def _w(tag: str) -> str:
    return "{%s}%s" % (W, tag)


def _wp(tag: str) -> str:
    return "{%s}%s" % (WP, tag)


def _attr(el, tag: str, ns: str = W):
    return el.get("{%s}%s" % (ns, tag))


def _unqualified(el, tag: str):
    """Geometry attributes are **unqualified** in OOXML: ``wp:extent/@cx``,
    ``wps:bodyPr/@lIns``, ``wp:anchor/@wrapType``, ``v:shape/@style``.

    Reading them with a namespace prefix (the natural mistake, since the
    *element* is namespaced) returns ``None`` for every real document and the
    box width silently degrades to "unknown" — which then never overflows.
    This helper exists so that bug can only be made once.
    """
    return el.get(tag) if el is not None else None


def _local(el) -> str:
    try:
        return etree.QName(el).localname
    except ValueError:                 # comments / processing instructions
        return ""


def _inside(el, ancestor_local: str) -> bool:
    """True when *el* has an ancestor whose local name is *ancestor_local*.

    ``w:sectPr`` appears twice in a document that tracks section changes: once
    as the live property, once inside ``w:sectPrChange`` as the *previous*
    value.  Counting the second one would invent a section that Word does not
    have, and every index after it would be off by one.
    """
    node = el.getparent()
    while node is not None:
        if _local(node) == ancestor_local:
            return True
        node = node.getparent()
    return False


# ---- fields ------------------------------------------------------------------

def field_keyword(instr: str) -> str:
    """``" PAGE \\\\* MERGEFORMAT "`` -> ``"PAGE"``; empty instruction -> ``""``.

    A field instruction is free text as far as the file format goes: Word writes
    ``PAGE``, `` NUMPAGES  \\* MERGEFORMAT ``, ``PAGEREF _Toc1 \\h``.  The first
    token is the field's keyword and the only part that decides what renders.
    """
    text = (instr or "").strip()
    return text.split()[0].upper() if text else ""


def field_keywords(root) -> list:
    """Every field instruction in *root*, in document order, as keywords.

    两种形态都要认，因为 Word 两种都会写:

    * ``<w:fldSimple w:instr=" PAGE "/>`` —— 一条元素，结果在它自己的子 run 里；
    * complex 形态 —— ``w:fldChar begin`` / ``w:instrText`` / ``w:fldChar
      separate`` / 结果 run / ``w:fldChar end``，且**可以嵌套**（一个 ``IF`` 域
      里套一个 ``PAGE``）。

    嵌套这一层就是这里用栈、不用正则的原因：``w:instrText`` 元素按文档顺序出
    现，只有栈知道当前这段指令属于哪一层域 —— 正则会把两条指令首尾相连成一个
    读不出来的关键字。
    """
    out: list = []
    stack: list = []
    for el in root.iter():
        tag = el.tag
        if tag == _w("fldSimple"):
            out.append(field_keyword(_attr(el, "instr")))
        elif tag == _w("fldChar"):
            kind = _attr(el, "fldCharType")
            if kind == "begin":
                stack.append({"instr": [], "reported": False})
            elif kind == "separate":
                if stack and not stack[-1]["reported"]:
                    out.append(field_keyword("".join(stack[-1]["instr"])))
                    stack[-1]["reported"] = True
            elif kind == "end":
                if stack:
                    top = stack.pop()
                    # A field with no ``separate`` still carries an instruction;
                    # Word tolerates that and renders the instruction's result,
                    # so it counts as a field here too.
                    if not top["reported"]:
                        out.append(field_keyword("".join(top["instr"])))
        elif tag == _w("instrText"):
            if stack:
                stack[-1]["instr"].append(el.text or "")
    return out


def has_field(root, name: str) -> bool:
    """True when *root* holds a field whose keyword is *name* (case-insensitive)."""
    want = name.upper()
    return any(k == want for k in field_keywords(root))


# ---- width estimation --------------------------------------------------------

def char_em(ch: str) -> float:
    """Advance width of one character in em.  See :data:`ESTIMATE_MODEL`."""
    if ch in (" ", "\u00a0"):
        return 0.25
    if ch == "\t":
        return 1.0
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return 1.0
    if ch.isdigit() or ch.isalpha():
        return 0.5
    if ch.isprintable():
        return 0.3
    return 0.0


def text_width_pt(text: str, size_pt: float) -> float:
    """Rendered width of *text* at *size_pt*, rounded to 0.01 pt."""
    return round(sum(char_em(c) for c in (text or "")) * float(size_pt), 2)


def _segments_text(segments) -> str:
    """Human-readable form: fields render as ``{PAGE}`` / ``{NUMPAGES}``."""
    return "".join("{%s}" % v if k == "field" else v for k, v, _s in segments)


def variant_widths(segments) -> dict:
    """Width of the line with 1 / 2 / 3 digits in **every** numeric field.

    Keys are the digit string substituted (``"1"``, ``"12"``, ``"123"``), so a
    report reads as「第 12 页　共 12 页 needs 63.0 pt」.  All numeric fields take
    the same digit count on purpose: a total can never have fewer digits than
    the page it counts to, so that is the pessimistic-but-reachable case.  The
    single-digit column is the one that always has to fit on page 1.
    """
    out = {}
    for digits in ("1", "12", "123"):
        total = 0.0
        for kind, value, size in segments:
            if kind == "field":
                total += text_width_pt(digits, size)
            else:
                total += text_width_pt(value, size)
        out[digits] = round(total, 2)
    return out


def _has_number(segments) -> bool:
    return any(k == "field" and (v == "PAGE" or v in TOTAL_FIELDS)
               for k, v, _s in segments)


# ---- styles: effective font size --------------------------------------------

class StyleCtx:
    """Resolve ``w:sz`` through the run → paragraph mark → ``pStyle`` → ``basedOn``
    → ``docDefaults`` chain.

    Sizing the page line needs the *effective* size of the run carrying the
    number, and in practice it is written in one of three places: the run's own
    ``w:rPr``, the paragraph mark's ``w:pPr/w:rPr``, or the paragraph style
    (which may inherit through ``basedOn``, e.g. ``Footer`` → ``Normal`` →
    ``docDefaults``).  Reading only the first would report a 9 pt footer as the
    document default and make every width wrong by the same factor.
    """

    def __init__(self, styles_root=None):
        self.size_by_id: dict = {}
        self.based_on: dict = {}
        self.default_para_style = None
        self.default_pt: float = 11.0       # Word's own fallback is 11 pt
        if styles_root is None:
            return
        doc_el = styles_root.find(_w("docDefaults"))
        if doc_el is not None:
            sz = doc_el.find("%s/%s/%s" % (_w("rPrDefault"), _w("rPr"), _w("sz")))
            if sz is not None:
                val = _attr(sz, "val")
                if val:
                    try:
                        self.default_pt = float(val) / 2.0
                    except ValueError:
                        pass
        for style in styles_root.iter(_w("style")):
            sid = _attr(style, "styleId")
            if not sid:
                continue
            sz = style.find("%s/%s" % (_w("rPr"), _w("sz")))
            if sz is not None:
                val = _attr(sz, "val")
                if val:
                    try:
                        self.size_by_id[sid] = float(val) / 2.0
                    except ValueError:
                        pass
            based = style.find(_w("basedOn"))
            if based is not None:
                self.based_on[sid] = _attr(based, "val")
            if (_attr(style, "type") == "paragraph"
                    and _attr(style, "default") in ("1", "true")):
                self.default_para_style = sid

    def size_for_style(self, sid):
        seen = set()
        while sid and sid not in seen:
            seen.add(sid)
            if sid in self.size_by_id:
                return self.size_by_id[sid]
            sid = self.based_on.get(sid)
        return None

    def size_for_paragraph(self, p_el) -> float:
        """Effective size for a paragraph, before any run-level override."""
        ppr = p_el.find(_w("pPr"))
        if ppr is not None:
            sz = ppr.find("%s/%s" % (_w("rPr"), _w("sz")))
            if sz is not None and _attr(sz, "val"):
                try:
                    return float(_attr(sz, "val")) / 2.0
                except ValueError:
                    pass
            style = ppr.find(_w("pStyle"))
            sid = _attr(style, "val") if style is not None else None
            got = self.size_for_style(sid or self.default_para_style)
            if got:
                return got
        got = self.size_for_style(self.default_para_style)
        return got if got else self.default_pt

    def run_size(self, r_el, fallback: float) -> float:
        rpr = r_el.find(_w("rPr"))
        if rpr is not None:
            for tag in ("sz", "szCs"):
                el = rpr.find(_w(tag))
                if el is not None and _attr(el, "val"):
                    try:
                        return float(_attr(el, "val")) / 2.0
                    except ValueError:
                        pass
        return fallback


# ---- footer content → segments ----------------------------------------------

def footer_segments(root, ctx: StyleCtx, skip_boxes: bool = True) -> list:
    """Flatten a footer into ``[(kind, value, size_pt), …]``.

    ``kind`` is ``"text"`` (value is literal text) or ``"field"`` (value is the
    field keyword).  Cached field *results* are dropped: a footer says ``PAGE``
    and Word renders whatever the counter is, so the ``1`` sitting between
    ``separate`` and ``end`` is a stale copy, not content — measuring it would
    report a 4 pt wide "1" instead of the real line.

    ``skip_boxes`` keeps ``w:drawing`` / ``w:pict`` subtrees out of the result: a
    text box's paragraph is measured separately (against its own width), and
    counting it here as well would report the same line twice.
    """
    segs: list = []
    _walk(root, ctx, segs, {"stack": []}, ctx.size_for_paragraph(root), skip_boxes)
    return segs


def _walk(node, ctx, segs, state, para_pt, skip_boxes) -> None:
    for ch in node:
        tag = ch.tag
        if not isinstance(tag, str):
            continue
        if skip_boxes and tag in (_w("drawing"), _w("pict")):
            continue
        if tag == _w("p"):
            _walk(ch, ctx, segs, state, ctx.size_for_paragraph(ch), skip_boxes)
        elif tag == _w("fldSimple"):
            kw = field_keyword(_attr(ch, "instr"))
            size = _first_run_size(ch, ctx, para_pt)
            segs.append(("field", kw, size) if kw
                        else ("text", _plain_text(ch), size))
        elif tag == _w("r"):
            _walk_run(ch, segs, state, ctx.run_size(ch, para_pt))
        else:
            _walk(ch, ctx, segs, state, para_pt, skip_boxes)


def _walk_run(r_el, segs, state, size) -> None:
    stack = state["stack"]
    for sub in r_el:
        tag = sub.tag
        if not isinstance(tag, str):
            continue
        if tag == _w("fldChar"):
            kind = _attr(sub, "fldCharType")
            if kind == "begin":
                stack.append({"instr": [], "result": False, "emitted": False,
                              "cached": []})
            elif kind == "separate":
                if stack:
                    top = stack[-1]
                    top["result"] = True
                    kw = field_keyword("".join(top["instr"]))
                    if kw:
                        # Emitting at ``separate`` puts the token exactly where
                        # Word starts rendering the result, so a footer mixing
                        # text and fields keeps its running order.
                        segs.append(("field", kw, size))
                        top["emitted"] = True
            elif kind == "end":
                if stack:
                    top = stack.pop()
                    if not top["emitted"] and top["cached"]:
                        # A field whose keyword we could not read (an empty
                        # instruction, a broken chain): measure its cached text,
                        # which is all there is.
                        segs.append(("text", "".join(top["cached"]), size))
        elif tag == _w("instrText"):
            if stack:
                stack[-1]["instr"].append(sub.text or "")
        elif tag == _w("t"):
            if stack and stack[-1]["result"]:
                if not stack[-1]["emitted"]:
                    stack[-1]["cached"].append(sub.text or "")
            else:
                segs.append(("text", sub.text or "", size))
        elif tag == _w("tab"):
            segs.append(("text", "\t", size))


def _first_run_size(node, ctx: StyleCtx, fallback: float) -> float:
    for r in node.iter(_w("r")):
        return ctx.run_size(r, fallback)
    return fallback


def _plain_text(node) -> str:
    return "".join(t.text or "" for t in node.iter(_w("t")))


# ---- text boxes --------------------------------------------------------------

def _pt_from_measure(value):
    """``"4.56pt"`` / ``"0.16cm"`` / ``"57912"`` (bare EMU) -> points."""
    if not value:
        return None
    text = str(value).strip().lower()
    num = ""
    for ch in text:
        if ch.isdigit() or ch in ".+-":
            num += ch
        else:
            break
    if not num:
        return None
    try:
        amount = float(num)
    except ValueError:
        return None
    unit = text[len(num):].strip()
    for prefix, factor in (("pt", 1.0), ("in", 72.0), ("cm", 72.0 / 2.54),
                           ("mm", 72.0 / 25.4), ("px", 0.75),
                           ("emu", 1.0 / EMU_PER_PT)):
        if unit.startswith(prefix):
            return round(amount * factor, 2)
    if not unit:                       # no unit at all: OOXML geometry is EMU
        return round(amount / EMU_PER_PT, 2)
    return None


def _style_props(style) -> dict:
    out = {}
    for chunk in (style or "").split(";"):
        if ":" not in chunk:
            continue
        key, _, val = chunk.partition(":")
        out[key.strip().lower()] = val.strip()
    return out


def _anchor_wrap(anchor):
    """``none`` / ``square`` / ``tight`` / ``through`` / ``topAndBottom``.

    Two spellings exist in the wild and both are read:

    * the DrawingML child elements ``<wp:wrapNone/>`` … (what Word 2010+ writes);
    * the ``wrapType`` attribute on ``<wp:anchor>`` (what earlier writers and
      several WPS / LibreOffice versions write).

    Whichever is present describes how the box floats *over the page*, which is
    what decides whether a clipped line gets re-flowed (it does not) or is simply
    lost (it is).
    """
    for child in anchor.iter():
        name = _local(child)
        if name in WRAP_ELEMENTS:
            rest = name[len("wrap"):]
            return rest[0].lower() + rest[1:]
    return _unqualified(anchor, "wrapType")


def _inset(body, name: str) -> float:
    if body is None:
        return 0.0
    raw = _unqualified(body, name)
    if not raw:
        return 0.0
    try:
        return round(float(raw) / EMU_PER_PT, 2)
    except ValueError:
        return 0.0


def textboxes(root, ctx: StyleCtx) -> list:
    """Every text box in a footer, with geometry, body properties and content.

    Three container shapes are recognised, because all three reach Word:

    * DrawingML floating — ``w:drawing`` → ``wp:anchor`` → … → ``wps:txbx``;
    * DrawingML inline — ``w:drawing`` → ``wp:inline`` → … → ``wps:txbx``;
    * VML — ``w:pict`` → ``v:shape`` → ``v:textbox`` (the classic
      ``mso-fit-shape-to-text`` box WPS writes by default).

    A box with no ``wp:extent`` and no VML ``width`` reports
    ``box_width_pt: null`` — an unknown width is not a 0 pt width, and treating
    it as one would turn every unmeasured box into a false overflow.
    """
    boxes: list = []
    for drawing in root.iter(_w("drawing")):
        host = drawing.find(_wp("anchor"))
        floating = host is not None
        if host is None:
            host = drawing.find(_wp("inline"))
        if host is None:
            continue
        for txbx in host.iter("{%s}txbx" % WPS):
            content = txbx.find(_w("txbxContent"))
            wsp = txbx.getparent()
            body = wsp.find("{%s}bodyPr" % WPS) if wsp is not None else None
            extent = host.find(_wp("extent"))
            width = None
            if extent is not None:
                width = _pt_from_measure(_unqualified(extent, "cx"))
            boxes.append(_box_record(content, ctx, floating, width, body,
                                     anchor=host if floating else None))
    for shape in root.iter("{%s}shape" % V):
        txbx = shape.find("{%s}textbox" % V)
        if txbx is None:
            continue
        content = txbx.find(_w("txbxContent"))
        props = _style_props(_unqualified(shape, "style"))
        width = _pt_from_measure(props.get("width"))
        boxes.append(_box_record(content, ctx, True, width, None, vml=props))
    return boxes


def _box_record(content, ctx: StyleCtx, floating: bool, width, body,
                anchor=None, vml=None) -> dict:
    segs = footer_segments(content, ctx, skip_boxes=False) if content is not None \
        else []
    widths = variant_widths(segs)
    ins_l, ins_r = _inset(body, "lIns"), _inset(body, "rIns")
    usable = None if width is None else round(max(width - ins_l - ins_r, 0.0), 2)
    wrap = _anchor_wrap(anchor) if anchor is not None else None
    if wrap is None and body is not None:
        # ``wps:bodyPr/@wrap`` is the *text* wrap inside the shape — a weaker
        # signal, used only when the anchor said nothing.
        wrap = _attr(body, "wrap", WPS)
    if wrap is None and vml:
        wrap = vml.get("mso-wrap-style")
    autofit = bool(body is not None
                   and body.find("{%s}spAutoFit" % A) is not None)
    if not autofit and vml:
        autofit = str(vml.get("mso-fit-shape-to-text", "")).lower() in ("t", "true")
    at = _first_overflow(widths, usable)
    has_num = _has_number(segs)
    return {
        "floating": floating,
        "spAutoFit": autofit,
        "wrap": wrap,
        "box_width_pt": width,
        "usable_width_pt": usable,
        "inset_left_pt": ins_l,
        "inset_right_pt": ins_r,
        "has_page_field": any(k == "field" and v == "PAGE" for k, v, _s in segs),
        "has_numpages_field": any(k == "field" and v in TOTAL_FIELDS
                                  for k, v, _s in segs),
        "text": _segments_text(segs),
        "font_size_pt": _first_field_size(segs),
        # Widths and the overflow verdict must come from the **same** computation.
        # ``overflow`` is decided from ``widths`` for every box, so gating the
        # widths on ``has_num`` left the two out of sync for a plain-text box:
        # ``overflow`` was True while ``page_number_widths_pt`` was None, and
        # anything reading ``widths[digits]`` raised TypeError (measured: 13-line
        # traceback, empty stdout, and `docx_validate` downgrading it to
        # ``{"ok": true}`` + one warning).  The widths are therefore published
        # whenever *either* a numeric field exists or the box overflows.
        "page_number_width_pt": widths["12"] if (has_num or at is not None) else None,
        "page_number_widths_pt": dict(widths) if (has_num or at is not None) else None,
        "overflow": at is not None,
        "overflow_at_digits": at,
        "_segments": [list(s) for s in segs],
    }


def _first_field_size(segments):
    for kind, _value, size in segments:
        if kind == "field":
            return round(size, 2)
    if segments:
        return round(segments[0][2], 2)
    return None


def _first_overflow(widths: dict, usable):
    """The smallest digit count whose rendered width exceeds *usable*, or None."""
    if usable is None:
        return None
    for digits in ("1", "12", "123"):
        if widths[digits] > usable:
            return digits
    return None


# ---- package walking ---------------------------------------------------------

def _rels_map(zf, part: str) -> dict:
    """``{rId: target part name}`` for the .rels that belongs to *part*."""
    folder, base = posixpath.split(part)
    rels_name = posixpath.join(folder, "_rels", base + ".rels")
    try:
        data = zf.read(rels_name)
    except KeyError:
        return {}
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return {}
    out = {}
    for rel in root.iter("{%s}Relationship" % PR):
        rid, target = rel.get("Id"), rel.get("Target") or ""
        if (rel.get("TargetMode") or "Internal") == "External":
            continue
        out[rid] = posixpath.normpath(posixpath.join(folder, target)).lstrip("/")
    return out


def section_props(doc_root) -> list:
    """Every live ``w:sectPr`` in document order — i.e. Word's section order."""
    return [sp for sp in doc_root.iter(W_SECTPR) if not _inside(sp, "sectPrChange")]


def footer_refs(sectPr) -> list:
    """``[{"type": …, "r_id": …}]`` for this section's own footer references."""
    return [{"type": _attr(el, "type") or "default",
             "r_id": _attr(el, "id", R)}
            for el in sectPr.findall(W_FOOTERREF)]


def pg_num_type(sectPr):
    """``{"fmt": …, "start": …}`` or None — the section's ``<w:pgNumType>``."""
    el = sectPr.find(W_PGNUMTYPE)
    if el is None:
        return None
    return {"fmt": _attr(el, "fmt"), "start": _attr(el, "start")}


# ---- one footer part ---------------------------------------------------------

def analyze_footer(part_name: str, root, ctx: StyleCtx) -> dict:
    """Everything this module can say about one footer part.

    ``shape`` has three values, and the definition is deliberately about *where
    the footer's content lives* rather than about which markup appears:

    * ``plain`` — no text box markup at all (``w:pict`` / ``wps:txbx`` /
      ``v:shape`` / ``mc:AlternateContent`` around a box): the footer is a normal
      paragraph, so its line is laid out in the text column and re-flowed;
    * ``floating-textbox`` — every text box is anchored/floating and the footer
      holds no direct paragraph text: the content lives inside the box;
    * ``mixed`` — a text box **and** direct paragraph text, or a mix of floating
      and inline boxes.

    The distinction matters because a ``floating-textbox`` footer is where
    content gets clipped instead of re-flowed.
    """
    outside = footer_segments(root, ctx, skip_boxes=True)
    boxes = textboxes(root, ctx)
    keywords = field_keywords(root)
    has_page = "PAGE" in keywords
    has_total = any(k in TOTAL_FIELDS for k in keywords)

    # The line that has to fit in the box is the line *inside* the box, when the
    # page field lives there.  Measuring the flattened footer would blend the
    # outside text into a width that never has to fit anywhere.
    carrier = next((b for b in boxes if b["has_page_field"]), None)
    if carrier is not None:
        widths = carrier["page_number_widths_pt"] or variant_widths([])
        text = carrier["text"]
    else:
        widths = variant_widths(outside) if has_page else variant_widths([])
        text = _segments_text(outside)

    overflowing = [b for b in boxes if b["overflow"]]
    at = min((b["overflow_at_digits"] for b in overflowing
              if b["overflow_at_digits"]), key=len, default=None)

    outside_text = _segments_text(outside)
    has_inline_box = any(not b["floating"] for b in boxes)
    if not boxes:
        shape = "plain"
    elif not outside_text.strip() and not has_inline_box:
        shape = "floating-textbox"
    else:
        # A box *and* direct paragraph text, or a floating box next to an inline
        # one: two layout systems in one footer, which is exactly the state that
        # makes "which of these lines can be clipped" ambiguous to a reader.
        shape = "mixed"

    widths_of_boxes = [b["box_width_pt"] for b in boxes
                       if b["box_width_pt"] is not None]
    first = overflowing[0] if overflowing else (boxes[0] if boxes else None)
    return {
        "part": part_name,
        "shape": shape,
        "spAutoFit": bool(any(b["spAutoFit"] for b in boxes)) if boxes else None,
        "wrap": first["wrap"] if first else None,
        "box_width_pt": min(widths_of_boxes) if widths_of_boxes else None,
        "usable_width_pt": first["usable_width_pt"] if first else None,
        "textboxes": boxes,
        "has_page_field": has_page,
        "has_numpages_field": has_total,
        "fields": keywords,
        "text": text,
        "font_size_pt": _first_field_size(outside) if outside else
        (carrier["font_size_pt"] if carrier else None),
        "page_number_width_pt": widths["12"] if has_page else None,
        "page_number_widths_pt": dict(widths) if has_page else None,
        "overflow": bool(overflowing),
        "overflow_at_digits": at,
    }


# ---- the whole package -------------------------------------------------------

def _require_word_package(path: str) -> None:
    """``_check_input`` 只保证「是个 ZIP」；这里再保证「是个 Word 包」。

    ZIP 是 Word 包的必要条件而不是充分条件：一个只有
    ``[Content_Types].xml`` 的 ZIP 通过了 ``is_zipfile``，随后读
    ``word/document.xml`` 会抛 ``KeyError`` —— 那是**路径给错了**，不是文档有问
    题，所以必须退 2（一行 stderr、空 stdout），不能变成 traceback。
    ``docx_read._open_document`` 走的是同一条判定（它把 ``KeyError`` 收成
    ``InputError``），这里不重开一遍 python-docx，只查那一个部件在不在。
    """
    with zipfile.ZipFile(path) as zf:
        if "word/document.xml" not in zf.namelist():
            raise InputError("not a Word package: %s" % path)


def check_path(path: str) -> dict:
    """Read *path* and report every section's footers.  Read-only: never writes.

    页脚部件缺失、不是良构 XML 时**不抛异常**：那种情况是「读不懂」，不是「命令
    用错」，所以收成一条 ``footer-part-unreadable`` 的 warning，让调用方
    （`docx_validate`）照样拿到一份可解释的报告。只有整个包打不开才会抛 —— 而
    那条路在 CLI 里已被 ``_check_input`` 挡在 exit 2。
    """
    issues: list = []
    sections: list = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        styles_root = None
        if "word/styles.xml" in names:
            try:
                styles_root = etree.fromstring(zf.read("word/styles.xml"))
            except etree.XMLSyntaxError:
                styles_root = None
        ctx = StyleCtx(styles_root)

        try:
            doc_root = etree.fromstring(zf.read("word/document.xml"))
        except KeyError:
            issues.append({"severity": "error",
                           "code": "footer-missing-document-xml",
                           "detail": ("word/document.xml: part absent — this is "
                                      "not a Word package")})
            return {"ok": False, "path": path, "sections": [], "issues": issues,
                    "estimate_model": ESTIMATE_MODEL}
        except etree.XMLSyntaxError as exc:
            issues.append({"severity": "error", "code": "footer-bad-document-xml",
                           "detail": "word/document.xml: %s" % exc})
            return {"ok": False, "path": path, "sections": [], "issues": issues,
                    "estimate_model": ESTIMATE_MODEL}

        rels = _rels_map(zf, "word/document.xml")
        cache: dict = {}
        reported: set = set()
        inherited_default = (None, None)     # (part name, owning section index)

        for index, sectPr in enumerate(section_props(doc_root)):
            own = footer_refs(sectPr)
            ftrs: list = []
            for ref in own:
                part_name = rels.get(ref["r_id"])
                if part_name is None:
                    foot = _unreadable("", ref["type"], False, None,
                                       "relationship %s resolves to no part"
                                       % ref["r_id"])
                else:
                    foot = _analyze_part(zf, names, part_name, ctx, cache,
                                         issues, reported)
                foot.update({"type": ref["type"], "inherited": False,
                             "inherited_from_section": None})
                ftrs.append(foot)
                if ref["type"] == "default":
                    inherited_default = (part_name or "", index)

            if not any(r["type"] == "default" for r in own) \
                    and inherited_default[0]:
                part_name, owner = inherited_default
                foot = _analyze_part(zf, names, part_name, ctx, cache, issues,
                                     reported)
                foot.update({"type": "default", "inherited": True,
                             "inherited_from_section": owner})
                ftrs.append(foot)

            sections.append({
                "index": index,
                "pg_num_type": pg_num_type(sectPr),
                "footer_kind": _footer_kind(own, inherited_default),
                "footers": ftrs,
            })

    # One issue per (code, part, section-that-renders-it), not per part.
    #
    # The distinction matters because a footer part is *shared*: section 1 with
    # no footer of its own renders section 0's, so a clipped page number is
    # clipped in both.  De-duplicating by part alone reported the defect against
    # section 0 only, and ``--section 1`` then answered "ok" about a page whose
    # number is still cut — exactly the false negative this package exists to
    # remove.  ``seen`` therefore keys on the section too.
    seen: set = set()
    for sec in sections:
        for foot in sec["footers"]:
            _check_overflow(sec, foot, issues, seen)
            _check_page_field(sec, foot, issues, seen)

    return {
        "ok": not any(i["severity"] == "error" for i in issues),
        "path": path,
        "sections": sections,
        "issues": issues,
        "estimate_model": ESTIMATE_MODEL,
    }


def _analyze_part(zf, names, part_name, ctx, cache, issues, reported) -> dict:
    if part_name in cache:
        return dict(cache[part_name])
    if part_name not in names:
        foot = _unreadable(part_name, "default", False, None,
                           "part absent from the package")
        cache[part_name] = foot
        return dict(foot)
    try:
        root = etree.fromstring(zf.read(part_name))
    except etree.XMLSyntaxError as exc:
        key = ("footer-part-unreadable", part_name)
        if key not in reported:
            reported.add(key)
            issues.append({"severity": "warning",
                           "code": "footer-part-unreadable",
                           "detail": "%s: not well-formed XML (%s)"
                                     % (part_name, exc)})
        foot = _unreadable(part_name, "default", False, None, "not well-formed XML")
        cache[part_name] = foot
        return dict(foot)
    foot = analyze_footer(part_name, root, ctx)
    cache[part_name] = foot
    return dict(foot)


def _unreadable(part_name, ttype, inherited, owner, why) -> dict:
    return {"part": part_name or "", "shape": "unknown", "type": ttype,
            "inherited": inherited, "inherited_from_section": owner,
            "spAutoFit": None, "wrap": None, "box_width_pt": None,
            "usable_width_pt": None, "textboxes": [], "has_page_field": False,
            "has_numpages_field": False, "fields": [], "text": "",
            "font_size_pt": None, "page_number_width_pt": None,
            "page_number_widths_pt": None, "overflow": False,
            "overflow_at_digits": None, "unreadable": why}


def _footer_kind(own, inherited_default) -> str:
    if any(r["type"] == "default" for r in own):
        return "own"
    return "inherited" if inherited_default[0] else "none"


def _check_overflow(sec, foot, issues, reported) -> None:
    """``footer-textbox-overflow`` — a box narrower than the number it must hold.

    severity **error**：这不是「可能难看」，是内容**永久丢失**。浮动框不参与
    重排（``wrap="none"`` 时连推挤都不会发生），超出框宽的字画到框外后被页边裁
    掉，读者拿到的是 ``第 1`` 而不是 ``第 12 页``，而文档里没有任何一处会报错。
    """
    if not foot["overflow"]:
        return
    part = foot["part"]
    key = ("footer-textbox-overflow", part, sec["index"])
    if key in reported:
        return
    reported.add(key)
    box = next(b for b in foot["textboxes"] if b["overflow"])
    digits = box["overflow_at_digits"]
    # Defence in depth: ``overflow`` and the widths are produced together in
    # ``_box_record``, so a missing widths dict here means the two drifted apart
    # again.  Report the overflow without the width figure rather than raising --
    # a traceback costs the caller the whole report, and the box width alone is
    # still enough to act on.
    widths = box.get("page_number_widths_pt") or {}
    needs = widths.get(digits)
    needs_text = _fmt(needs) if needs is not None else "(width unknown)"
    issues.append({
        "severity": "error",
        "code": "footer-textbox-overflow",
        "section": sec["index"],
        "detail": ("section %d %s footer %s: %s pt wide floating text box "
                   "cannot hold the line — %r needs %s pt at %s digit(s) "
                   "(box %s pt, usable %s pt, wrap=%s, spAutoFit=%s)"
                   % (sec["index"], part, foot.get("type") or "default",
                      _fmt(box["box_width_pt"]), box["text"],
                      needs_text, digits,
                      _fmt(box["box_width_pt"]), _fmt(box["usable_width_pt"]),
                      box["wrap"], box["spAutoFit"])),
    })


def _check_page_field(sec, foot, issues, reported) -> None:
    """``footer-no-page-field`` — numbering configured, number nowhere shown.

    severity **warning**, not error, for two reasons measured on the real corpus:

    * 这个形态存在合法用法 —— 先把 ``pgNumType`` 设成罗马数字、页脚留给后续工
      序去填的文档并不因此坏掉；
    * 真实语料里「``pgNumType`` 在、页脚没 PAGE 域」有 4 份，全都是**整篇没有页
      脚**（``footer_kind: none``）—— 那是「没做页脚」，不是「页脚里漏了号」。

    所以命中条件收窄到「这一节解析得到页脚、页脚有内容、里面有别的域或文字，但
    没有 PAGE」；空页脚是被有意清掉的，不报。

    另外还要求 ``pgNumType`` **真的配置了东西**（``w:fmt`` 或 ``w:start`` 至少
    一个）。裸 ``<w:pgNumType/>`` 什么也没设 —— 它只是「这个元素存在」，而
    ``docx_section set_pg_num_type`` 在不传 fmt/start 时正是这样写的（它的契约
    是「保证元素存在且顺序合法」）。把裸元素算成「编号设了却不显示」是纯误报：
    实测一份「裸 pgNumType + 只有文字的页脚」的文档会因此被判 warning。
    """
    if sec["pg_num_type"] is None or foot.get("unreadable"):
        return
    if not any(sec["pg_num_type"].get(k) for k in ("fmt", "start")):
        return
    if foot["has_page_field"] or not foot["part"]:
        return
    if not (foot["text"].strip() or foot["fields"]):
        return
    part = foot["part"]
    key = ("footer-no-page-field", part, sec["index"])
    if key in reported:
        return
    reported.add(key)
    pgn = sec["pg_num_type"]
    issues.append({
        "severity": "warning",
        "code": "footer-no-page-field",
        "section": sec["index"],
        "detail": ("section %d sets <w:pgNumType start=%r fmt=%r> but its %s "
                   "footer %s has no PAGE field (fields=%r, text=%r) — the "
                   "counter is configured and never shown"
                   % (sec["index"], pgn.get("start"), pgn.get("fmt"),
                      "inherited" if foot.get("inherited") else "own",
                      part, foot["fields"], foot["text"])),
    })


def _fmt(value) -> str:
    if value is None:
        return "unknown"
    return ("%.2f" % float(value)).rstrip("0").rstrip(".")


def footer_issues(path: str) -> list:
    """Issues only — the shape `docx_validate` consumes.

    绝不抛异常：``docx_validate`` 的定位是「读得懂就报，读不懂也不能崩」，所以
    「页脚这一层读不动」也收成一条 warning 交给调用方，让报告里留痕而不是变成
    traceback。
    """
    try:
        return check_path(path)["issues"]
    except Exception as exc:                       # noqa: BLE001 - by design
        return [{"severity": "warning", "code": "footer-check-skipped",
                 "detail": "%s: %s" % (type(exc).__name__, exc)}]


# ---- CLI --------------------------------------------------------------------

def _select(report: dict, index: int) -> dict:
    secs = report["sections"]
    if not secs:
        raise IndexError("document has no sections")
    if not -len(secs) <= index < len(secs):
        raise IndexError("no such section %d (document has %d)"
                         % (index, len(secs)))
    keep = index % len(secs)
    out = dict(report)
    out["sections"] = [s for s in secs if s["index"] == keep]
    # ``issues`` is a document-level list; keep the entries that belong to the
    # section the caller asked for, plus the ones with no section attribution at
    # all (a footer part that will not parse is a fact about the whole package,
    # not about one section).  Matched on the structured ``section`` key, never
    # by parsing the human-readable detail string — that would break the moment
    # the wording changed.
    out["issues"] = [i for i in report["issues"]
                     if i.get("section") in (keep, None)]
    out["ok"] = not any(i["severity"] == "error" for i in out["issues"])
    return out


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Footer shape / width / page-field check for .docx "
                    "(read-only, JSON on stdout).")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("check", help="report every section's footers")
    p.add_argument("path", help="the .docx file to check")
    p.add_argument("--section", type=int, default=None,
                   help="only this section index (0-based; negative counts from "
                        "the end)")
    p.add_argument("--json", action="store_true",
                   help="explicit JSON (JSON is already the only output form; "
                        "the flag keeps calls looking like the other scripts)")
    args = ap.parse_args()

    # 用法错误先于任何读取：路径三档（缺失 / 是目录 / 非 ZIP）由 `docx_read` 的
    # 三件套单点拥有，这里只把它转成 exit 2 + 一行 stderr + 空 stdout。
    try:
        _check_input(args.path)
        _require_word_package(args.path)
    except InputError as exc:
        print("error: %s" % exc, file=sys.stderr)
        return 2

    report = check_path(args.path)
    if args.section is not None:
        try:
            report = _select(report, args.section)
        except IndexError as exc:
            # 越界是**用法错误**：用户点了一节不存在的节，不是文档有问题。本项目
            # exit 2 的定义正是「下标在合法范围之外」（见 SKILL.md 的 Exit codes
            # 行：an `--index` / `--insert-at` outside the legal range）。
            print("error: %s" % exc, file=sys.stderr)
            return 2

    print(json.dumps(report, ensure_ascii=False))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())