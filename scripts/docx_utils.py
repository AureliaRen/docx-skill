"""Shared OOXML helpers for docx-editing scripts.  One-liner imports everywhere."""
from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Additional namespaces used by drawing/image/relationship scripts
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Full namespace set for root element fix (ElementTree loses non-w namespaces)
NS = {k: v for k, v in {
    "w": W, "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "r": R,
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": WP, "a": A, "pic": PIC,
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "o": "urn:schemas-microsoft-com:office:office",
    "m": M, "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpsCustomData": "http://www.wps.cn/officeDocument/2013/wpsCustomData",
}.items()}

ROOT_START = ('<?xml version="1.0" encoding="UTF-8"?><w:document ' +
    ' '.join(f'xmlns:{k}="{v}"' for k, v in NS.items()) +
    ' mc:Ignorable="w14 w15 wp14">')


# ---- tag helpers -----------------------------------------------------------

def wtag(name: str) -> str:
    """``{W}name``"""
    return f"{{{W}}}{name}"


def mtag(name: str) -> str:
    """``{M}name``"""
    return f"{{{M}}}{name}"


# ---- ElementTree boilerplate -----------------------------------------------

def regw():
    ET.register_namespace("w", W)


def regm():
    ET.register_namespace("m", M)


def register_all() -> None:
    """Register all known OOXML namespace prefixes with ElementTree.

    Call this before parsing or writing any XML in scripts that touch
    drawing (wp/a/pic), relationships (r), math (m), or other non-w
    namespaces.  Prevents ElementTree from generating ``ns1:``, ``ns2:``
    prefixes that break namespace-dependent tools.
    """
    for prefix, uri in NS.items():
        ET.register_namespace(prefix, uri)


def parse(path: str | Path) -> ET.ElementTree:
    return ET.parse(str(path))


def write(tree: ET.ElementTree, path: str | Path) -> None:
    """Serialize *tree* with correct OOXML prefixes and a valid root tag.

    ElementTree drops declarations for namespaces no element happens to use,
    while leaving attributes such as ``mc:Ignorable="w14 wp14"`` intact. Word
    rejects a part whose ``mc:Ignorable`` names an undeclared prefix, so every
    touched part is repaired after serialization.
    """
    register_all()
    tree.write(str(path), xml_declaration=True, encoding="UTF-8")
    fix_root_any(Path(path))


def fix_root_any(path: Path) -> None:
    """Repair the root element of any OOXML part (document, styles, header...).

    ``fix_root`` only knows how to rewrite ``<w:document>``; this generalisation
    works for ``<w:styles>``, ``<w:ftr>``, ``<w:hdr>`` and friends, which the
    template/comment/revision scripts also serialize.  It restores every
    declaration in ``NS``, keeps any extra ones the part already had, and
    rewrites ``mc:Ignorable`` to list only prefixes that are now declared.
    """
    raw = path.read_text(encoding="utf-8")
    start = raw.find("<", raw.find("?>") + 2 if "?>" in raw else 0)
    if start == -1:
        return
    end = raw.find(">", start)
    if end == -1:
        return
    root_tag = raw[start:end + 1]
    if root_tag.endswith("/>") or "xmlns" in root_tag and "<w:" not in root_tag:
        pass
    existing_ns = dict(_XMLNS_RE.findall(root_tag))

    merged = dict(existing_ns)
    for prefix, uri in NS.items():
        merged[prefix] = uri

    ignorable = " ".join(k for k in merged if k not in ("w", "r", "mc"))

    # Drop every xmlns declaration we are about to re-emit, then append the
    # full set, so nothing is declared twice.
    stripped = _XMLNS_RE.sub("", root_tag)
    stripped = _re.sub(r'\s+', " ", stripped)
    stripped = _re.sub(r'\s+mc:Ignorable="[^"]*"', "", stripped)
    stripped = _re.sub(r'\s+>', ">", stripped)
    head = stripped[:-1].rstrip()
    if head.endswith("/"):
        head = head[:-1].rstrip()
    new_root = (head + " "
                + " ".join(f'xmlns:{k}="{v}"' for k, v in merged.items())
                + f' mc:Ignorable="{ignorable}">')

    body = raw[end + 1:]
    path.write_text(new_root + body, encoding="utf-8")


# ---- shared utilities ------------------------------------------------------

def insert_near(body: ET.Element, marker: str | None, position: str,
                *paras: ET.Element) -> bool:
    """Insert *paras* near the paragraph whose text contains *marker*.

    Searches inside ``<w:body>`` and all ``<w:tc>`` elements. Falls back to
    inserting before ``<w:sectPr>``.

    Returns True when the marker was found and the paragraphs landed next to it,
    False when the fallback was used. Callers that were told to anchor on a
    marker should treat False as an error rather than silently accepting a
    formula that ended up at the end of the document.
    """
    if marker:
        containers = [body] + list(body.iter(wtag("tc")))
        for container in containers:
            children = list(container)
            for i, c in enumerate(children):
                if c.tag != wtag("p"):
                    continue
                if marker in "".join(t.text or "" for t in c.iter(wtag("t"))):
                    at = i + 1 if position == "after" else i
                    for j, p in enumerate(paras):
                        container.insert(at + j, p)
                    return True
    # fallback
    children = list(body)
    idx = next((i for i, c in enumerate(children) if c.tag == wtag("sectPr")), len(children))
    for j, p in enumerate(paras):
        body.insert(idx + j, p)
    return False


def next_bm_id(root: ET.Element) -> int:
    """Smallest unused bookmark id."""
    used = set()
    for bm in root.iter(wtag("bookmarkStart")):
        try:
            used.add(int(bm.get(wtag("id"), "0")))
        except ValueError:
            pass
    n = 0
    while n in used:
        n += 1
    return n


# ---- schema-ordered child insertion ----------------------------------------

# OOXML complex types are xsd:sequence — children MUST appear in this exact
# order or the whole document fails XSD validation ("element not expected").
# Source: schemas/ISO-IEC29500-4_2016/wml.xsd, verified 2026-10-27.
TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
               "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
               "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
               "tblCellMar", "tblLook", "tblCaption", "tblDescription"]

TRPR_ORDER = ["cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore",
              "wAfter", "cantSplit", "trHeight", "tblHeader", "tblCellSpacing",
              "jc", "hidden"]

TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
              "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
              "hideMark", "headers"]

PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
             "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
             "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
             "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
             "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl",
             "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"]

_ORDER_MAP = {
    "tblPr": TBLPR_ORDER, "trPr": TRPR_ORDER, "tcPr": TCPR_ORDER,
    "pPr": PPR_ORDER,
}


def insert_ordered(parent: ET.Element, child: ET.Element,
                   order: list[str] | str | None = None) -> ET.Element:
    """Insert *child* into *parent* at the position the OOXML schema requires.

    ``order`` may be an explicit list, a parent tag name (looked up in
    ``_ORDER_MAP``), or None (inferred from *parent*'s local tag).

    Replaces an existing same-name child instead of adding a duplicate —
    xsd:sequence allows each child at most once in these complex types, and a
    duplicate surfaces as "element not expected" during validation.

    Returns the element actually present in the tree (existing or new).
    """
    if order is None or isinstance(order, str):
        parent_tag = (order if isinstance(order, str)
                      else parent.tag.split("}")[-1])
        order = _ORDER_MAP.get(parent_tag)
    if not order:
        parent.append(child)
        return child

    name = child.tag.split("}")[-1]
    if name not in order:
        parent.append(child)
        return child

    existing = parent.find(child.tag)
    if existing is not None:
        # keep the schema-unique element but let the caller's attributes win
        for k, v in child.attrib.items():
            existing.set(k, v)
        return existing

    want = order.index(name)
    insert_at = len(parent)
    for i, sibling in enumerate(parent):
        stag = sibling.tag.split("}")[-1]
        if stag in order and order.index(stag) > want:
            insert_at = i
            break
    parent.insert(insert_at, child)
    return child


def set_w_val(parent: ET.Element, name: str, val: str,
              order: list[str] | str | None = None) -> ET.Element:
    """Create-or-update ``<w:name w:val="val"/>`` inside *parent*, schema-ordered."""
    el = parent.find(wtag(name))
    if el is None:
        el = ET.Element(wtag(name))
        insert_ordered(parent, el, order)
    el.set(wtag("val"), str(val))
    return el


def fix_child_order(parent: ET.Element, order: list[str] | str | None = None) -> int:
    """Re-sort an existing element's children into schema order.

    Returns the number of children moved.  Use this on elements written by
    hand-appended XML (python-docx's ``append`` puts things at the end).
    """
    if order is None or isinstance(order, str):
        parent_tag = (order if isinstance(order, str)
                      else parent.tag.split("}")[-1])
        order = _ORDER_MAP.get(parent_tag)
    if not order:
        return 0
    kids = list(parent)
    ranked = [k for k in kids if k.tag.split("}")[-1] in order]
    if len(ranked) < 2:
        return 0
    wanted = sorted(ranked, key=lambda k: order.index(k.tag.split("}")[-1]))
    if wanted == ranked:
        return 0
    # position of the first ordered child currently in the tree
    kids_now = list(parent)
    start = min(i for i, k in enumerate(kids_now) if k in ranked)
    for k in ranked:
        parent.remove(k)
    # re-insert the block starting at the same index (ElementTree is index-based)
    for offset, k in enumerate(wanted):
        parent.insert(start + offset, k)
    return len(wanted)


def cm_to_dxa(cm: float) -> int:
    """Centimetres → twentieths of a point (w:w / w:dw unit)."""
    return int(round(cm / 2.54 * 1440))


# ---- run / paragraph builders ----------------------------------------------

def mkrun(text: str, *,
          font: str | dict | None = None,
          size: int | None = None,
          bold: bool = False,
          color: str | None = None,
          vert: str | None = None) -> ET.Element:
    """Create ``<w:r><w:rPr>…</w:rPr><w:t>text</w:t></w:r>``."""
    r = ET.Element(wtag("r"))
    rpr = ET.Element(wtag("rPr"))
    if font:
        rf = ET.Element(wtag("rFonts"))
        if isinstance(font, str):
            for a in ("ascii", "hAnsi", "eastAsia", "cs"):
                rf.set(wtag(a), font)
        else:
            for a, v in font.items():
                rf.set(wtag(a), v)
        rpr.append(rf)
    if size:
        for tagn in ("sz", "szCs"):
            e = ET.Element(wtag(tagn)); e.set(wtag("val"), str(size)); rpr.append(e)
    if bold:
        for tagn in ("b", "bCs"):
            rpr.append(ET.Element(wtag(tagn)))
    if color:
        c = ET.Element(wtag("color")); c.set(wtag("val"), color.lstrip("#")); rpr.append(c)
    if vert:
        v = ET.Element(wtag("vertAlign")); v.set(wtag("val"), vert); rpr.append(v)
    if len(rpr) > 0:
        r.append(rpr)
    t = ET.Element(wtag("t")); t.text = text
    if text and (text[0] == " " or text[-1] == " "):
        t.set(XML_SPACE, "preserve")
    r.append(t)
    return r


def mkpara(*children: ET.Element,
           align: str | None = None,
           spacing: dict | None = None,
           indent: dict | None = None) -> ET.Element:
    """Create ``<w:p>`` with optional ``<w:pPr>``."""
    p = ET.Element(wtag("p"))
    ppr = ET.Element(wtag("pPr"))
    if align:
        jc = ET.Element(wtag("jc")); jc.set(wtag("val"), align); ppr.append(jc)
    if spacing:
        sp = ET.Element(wtag("spacing"))
        for k, v in spacing.items():
            sp.set(wtag(k), str(v))
        ppr.append(sp)
    if indent:
        ind = ET.Element(wtag("ind"))
        for k, v in indent.items():
            ind.set(wtag(k), str(v))
        ppr.append(ind)
    if len(ppr) > 0:
        p.insert(0, ppr)
    for c in children:
        p.append(c)
    return p


def mkbookmark(name: str, bid: int) -> tuple[ET.Element, ET.Element]:
    """Return ``(bookmarkStart, bookmarkEnd)`` pair."""
    s = ET.Element(wtag("bookmarkStart"))
    s.set(wtag("id"), str(bid)); s.set(wtag("name"), name)
    e = ET.Element(wtag("bookmarkEnd"))
    e.set(wtag("id"), str(bid))
    return s, e


# ---- namespace fix ---------------------------------------------------------

# Pre-compiled regex for fix_root (avoids re-parsing on every call)
import re as _re
_XMLNS_RE = _re.compile(r'xmlns:(\w+)="([^"]*)"')
_NSN_COLON_RE = _re.compile(r'\b(ns\d+):')


def _check_ns_in_body(body: str) -> bool:
    """Fast check: does *body* contain any ``nsN:`` auto-generated prefix?"""
    # Look for "ns" followed by digit(s) followed by ":" — a quick substring
    # scan is 10-50× cheaper than a full regex findall.
    idx = body.find("ns")
    while idx != -1:
        j = idx + 2
        while j < len(body) and body[j].isdigit():
            j += 1
        if j > idx + 2 and j < len(body) and body[j] == ":":
            return True
        idx = body.find("ns", j)
    return False


def fix_root(path: Path) -> None:
    """Ensure the root element declares all known namespaces and ``mc:Ignorable``.

    ``write()`` already calls ``register_all()``, so ElementTree uses correct
    prefixes.  However it may still drop declarations for namespaces that
    aren't used in the tree.  This function restores the full set from ``NS``,
    preserving any extra declarations found in the current root (e.g. WPS
    custom namespaces).  Call this after ``write()`` on ``document.xml``.

    Performance: the expensive body-wide regex scan for auto-generated
    ``nsN:`` prefixes is only run when a fast substring check first detects
    their presence (which is rare when ``register_all()`` is called before
    ``write()``).
    """
    raw = path.read_text(encoding="utf-8")
    start = raw.find("<w:document")
    end = raw.find(">", start)
    if start == -1 or end == -1:
        return

    # --- 1. Fix root tag: merge existing + NS declarations ------------------
    existing_ns = dict(_XMLNS_RE.findall(raw[start:end]))

    merged = {}
    for prefix, uri in existing_ns.items():
        merged[prefix] = uri
    for prefix, uri in NS.items():
        merged[prefix] = uri   # our known URIs take precedence

    ignorable = " ".join(k for k in merged if k != "w")

    new_root = ('<?xml version="1.0" encoding="UTF-8"?><w:document ' +
        ' '.join(f'xmlns:{k}="{v}"' for k, v in merged.items()) +
        f' mc:Ignorable="{ignorable}">')

    body = raw[end + 1:]

    # --- 2. Fix auto-generated nsN: prefixes in child elements ---------------
    # Fast-path: substring scan before expensive regex
    if _check_ns_in_body(body):
        uri_to_prefix = {uri: prefix for prefix, uri in merged.items()}
        # Collect unique nsN: prefixes that need fixing
        prefixes_to_fix = set()
        for m in _NSN_COLON_RE.finditer(body):
            old_prefix = m.group(1)
            old_uri = existing_ns.get(old_prefix)
            if old_uri and old_uri in uri_to_prefix:
                prefixes_to_fix.add(old_prefix)
        # Apply replacements — only in XML tag/attribute context to avoid
        # corrupting text content that happens to contain nsN: literally.
        # Patterns: <ns1:tag, </ns1:tag,  ns1:attr=
        _TAG_PREFIX_RE = _re.compile(r'(</?)(ns\d+)(:\w+)')
        _ATTR_PREFIX_RE = _re.compile(r'\s(ns\d+)(:\w+=)')
        for old_prefix in prefixes_to_fix:
            new_prefix = uri_to_prefix[existing_ns[old_prefix]]
            body = _TAG_PREFIX_RE.sub(
                rf'\g<1>{new_prefix}\g<3>', body)
            body = _ATTR_PREFIX_RE.sub(
                rf' {new_prefix}\g<2>', body)

    path.write_text(new_root + body, encoding="utf-8")


# ---- character classification -----------------------------------------------

def is_cjk(ch: str) -> bool:
    """Return True if *ch* is a CJK (Chinese/Japanese/Korean) character."""
    cp = ord(ch)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or   # CJK Unified
            0x3040 <= cp <= 0x309F or 0x30A0 <= cp <= 0x30FF or   # Hiragana/Katakana
            0xAC00 <= cp <= 0xD7AF or 0xF900 <= cp <= 0xFAFF or   # Hangul / Compat
            0x20000 <= cp <= 0x2A6DF or 0x2F800 <= cp <= 0x2FA1F)  # Ext B / Compat Suppl


# ---- default fonts ---------------------------------------------------------

FONT_TNR = {"ascii": "Times New Roman", "hAnsi": "Times New Roman",
            "eastAsia": "微软雅黑", "cs": "Times New Roman"}

FONT_HEITI = {"ascii": "Times New Roman", "hAnsi": "Times New Roman",
              "eastAsia": "黑体", "cs": "Times New Roman"}

FONT_SONG = {"ascii": "Times New Roman", "hAnsi": "Times New Roman",
             "eastAsia": "宋体", "cs": "Times New Roman"}


if __name__ == "__main__":
    print("docx_utils is a library module, import it:")
    print("    from docx_utils import wtag, parse, write, insert_ordered, FONT_TNR")
    raise SystemExit(0)
