#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Manage Word auto-numbering (``word/numbering.xml``) in a .docx.

Word numbering is a two-level indirection that nothing else in this toolkit
exposed:

  ``<w:abstractNum>``  the *definition* — level format, ``lvlText`` such as
                       ``[%1]``, start value, indent
  ``<w:num>``          a *counter instance* pointing at a definition; the
                       ``numId`` is what a paragraph actually names
  ``<w:numPr>``        on the paragraph: ``ilvl`` + ``numId``

Restarting a list means pointing some paragraphs at a different ``<w:num>``;
changing the bracket style means editing ``lvlText`` on the abstractNum. The
built-in ``numbered_list`` block in ``docx_create`` only ever applies the
``List Number`` *style*, which carries no ``numPr`` — so documents built that
way have zero real numbering. This script writes the real thing.

Subcommands:
  list     show abstractNum / num definitions and paragraph usage
  define   append a new definition + counter, print the numId to use
  apply    attach numPr to paragraphs (by index, range, text or style)
  detach   strip numPr from paragraphs
  restart  give a paragraph its own counter so numbering restarts there

Examples:
  docx_numbering.py list in.docx
  docx_numbering.py define in.docx --text "[%1]" --fmt decimal -o out.docx
  docx_numbering.py apply in.docx --index 155:171 --num-id 1 -o out.docx
  docx_numbering.py apply in.docx --style 正文 --num-id 1
  docx_numbering.py restart in.docx --index 7 --num-id 1 -o out.docx

Exit codes: 0 success, 1 expected failure (no such numId, empty selection),
2 usage error.
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import random
import sys

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_read import InputError, _check_input, _open_document
from docx_utils import insert_ordered

# Formats Word accepts for <w:numFmt w:val="...">.
VALID_FMT = {
    "decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman",
    "upperRoman", "ordinal", "cardinalText", "ordinalText", "hex",
    "chicago", "ideographDigital", "japaneseCounting", "chineseCounting",
    "chineseLegalSimplified", "none", "bullet",
}


# ---- part access ------------------------------------------------------------

def numbering_root(doc, create: bool = False):
    """Return the ``<w:numbering>`` root, optionally creating the part."""
    try:
        return doc.part.numbering_part.element
    except (KeyError, NotImplementedError, AttributeError):
        if not create:
            return None

    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    root = parse_xml(f"<w:numbering {nsdecls('w')}/>")
    part = Part(PackURI("/word/numbering.xml"),
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.numbering+xml",
                b"", doc.part.package)
    part._element = root
    doc.part.relate_to(part, RT.NUMBERING)
    return root


def _existing_ids(root) -> tuple[set[int], set[int]]:
    abs_ids = {int(e.get(qn("w:abstractNumId")))
               for e in root.findall(qn("w:abstractNum"))}
    num_ids = {int(e.get(qn("w:numId"))) for e in root.findall(qn("w:num"))}
    return abs_ids, num_ids


def _new_nsid() -> str:
    """Word uses an 8-hex-digit nsid to keep formats distinct."""
    return f"{random.randint(0, 0xFFFFFFFF):08X}"


# ---- paragraphs -------------------------------------------------------------

def iter_paragraphs(doc):
    """Body paragraphs plus those inside tables, in document order."""
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


def parse_selection(spec: str) -> tuple[int, int]:
    """``"155:171"`` -> (155, 171); ``"7"`` -> (7, 7). End is exclusive."""
    if ":" in spec:
        a, b = spec.split(":", 1)
        return int(a), int(b)
    n = int(spec)
    return n, n + 1


def select(doc, args) -> list:
    """Resolve the paragraph selection from the CLI arguments.

    Different selector *kinds* combine with AND — "the paragraphs that use
    this style **and** contain this text" — because that is what lets a caller
    separate two runs of prose that share one style (a body style used for
    both prose and the reference list, say). Repeated ``--index`` values stay
    OR'd, since they are just several ranges of the same kind.
    """
    paras = list(iter_paragraphs(doc))

    index_hit: set[int] | None = None
    if args.index:
        index_hit = set()
        for spec in args.index:
            lo, hi = parse_selection(spec)
            for i in range(max(0, lo), min(len(paras), hi)):
                index_hit.add(i)

    picked: list = []
    for i, p in enumerate(paras):
        if index_hit is not None and i not in index_hit:
            continue
        if args.style and not (p.style is not None
                               and p.style.name == args.style):
            continue
        if args.match and args.match not in p.text:
            continue
        picked.append(p)
    return picked


def set_numpr(para, num_id: int, ilvl: int) -> None:
    """Attach (or replace) ``<w:numPr>`` on a paragraph.

    ``ilvl`` must precede ``numId`` inside numPr — CT_NumPr is xsd:sequence.
    """
    pPr = para._p.get_or_add_pPr()
    old = pPr.find(qn("w:numPr"))
    if old is not None:
        pPr.remove(old)
    numPr = OxmlElement("w:numPr")
    lvl = OxmlElement("w:ilvl")
    lvl.set(qn("w:val"), str(ilvl))
    nid = OxmlElement("w:numId")
    nid.set(qn("w:val"), str(num_id))
    numPr.append(lvl)
    numPr.append(nid)
    insert_ordered(pPr, numPr, "pPr")


def clear_numpr(para) -> bool:
    pPr = para._p.pPr
    if pPr is None:
        return False
    old = pPr.find(qn("w:numPr"))
    if old is None:
        return False
    pPr.remove(old)
    return True


def current_numpr(para):
    pPr = para._p.pPr
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    nid = numPr.find(qn("w:numId"))
    lvl = numPr.find(qn("w:ilvl"))
    return (int(nid.get(qn("w:val"))) if nid is not None else 0,
            int(lvl.get(qn("w:val"))) if lvl is not None else 0)


# ---- subcommands ------------------------------------------------------------

def cmd_list(doc) -> dict:
    root = numbering_root(doc)
    if root is None:
        return {"abstract_nums": [], "nums": [], "usage": {}}

    abs_defs = {}
    for e in root.findall(qn("w:abstractNum")):
        aid = e.get(qn("w:abstractNumId"))
        lvl0 = e.find(qn("w:lvl"))
        text = fmt = start = None
        if lvl0 is not None:
            lt = lvl0.find(qn("w:lvlText"))
            nf = lvl0.find(qn("w:numFmt"))
            st = lvl0.find(qn("w:start"))
            text = None if lt is None else lt.get(qn("w:val"))
            fmt = None if nf is None else nf.get(qn("w:val"))
            start = None if st is None else st.get(qn("w:val"))
        abs_defs[aid] = {"abstractNumId": aid, "lvlText": text,
                         "numFmt": fmt, "start": start,
                         "levels": len(e.findall(qn("w:lvl")))}

    nums = []
    for e in root.findall(qn("w:num")):
        aid_el = e.find(qn("w:abstractNumId"))
        nums.append({"numId": e.get(qn("w:numId")),
                     "abstractNumId": None if aid_el is None
                     else aid_el.get(qn("w:val"))})

    usage: dict[str, int] = {}
    for p in iter_paragraphs(doc):
        cur = current_numpr(p)
        if cur:
            usage[str(cur[0])] = usage.get(str(cur[0]), 0) + 1

    return {"abstract_nums": list(abs_defs.values()), "nums": nums,
            "usage": usage}


def cmd_define(doc, args) -> dict:
    if args.fmt not in VALID_FMT:
        raise SystemExit(f"unknown --fmt {args.fmt!r}; choose from "
                         f"{', '.join(sorted(VALID_FMT))}")
    root = numbering_root(doc, create=True)
    abs_ids, num_ids = _existing_ids(root)

    aid = 0
    while aid in abs_ids:
        aid += 1
    nid = 1
    while nid in num_ids:
        nid += 1

    nsid = _new_nsid()
    lvl_text = args.text
    if "%1" not in lvl_text and not lvl_text.endswith("%"):
        # A definition with no placeholder never shows a number.
        lvl_text = lvl_text + "%1"

    abs_el = OxmlElement("w:abstractNum")
    abs_el.set(qn("w:abstractNumId"), str(aid))
    for tag, val in (("w:nsid", nsid),
                     ("w:multiLevelType", "singleLevel"),
                     ("w:tmpl", nsid)):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        abs_el.append(e)

    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), "0")
    lvl.set(qn("w:tentative"), "0")
    for tag, val in (("w:start", str(args.start)),
                     ("w:numFmt", args.fmt),
                     ("w:lvlText", lvl_text),
                     ("w:lvlJc", args.jc)):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        lvl.append(e)

    pPr = OxmlElement("w:pPr")
    if args.indent_left or args.hanging:
        ind = OxmlElement("w:ind")
        if args.indent_left:
            ind.set(qn("w:left"), str(args.indent_left))
        if args.hanging:
            ind.set(qn("w:hanging"), str(args.hanging))
        pPr.append(ind)
    if len(pPr):
        lvl.append(pPr)

    abs_el.append(lvl)

    # CT_Numbering is xsd:sequence: every abstractNum precedes every num.
    first_num = root.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abs_el)
    else:
        root.append(abs_el)

    num_el = OxmlElement("w:num")
    num_el.set(qn("w:numId"), str(nid))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), str(aid))
    num_el.append(ref)
    root.append(num_el)

    return {"ok": True, "numId": nid, "abstractNumId": aid,
            "lvlText": lvl_text, "numFmt": args.fmt}


def cmd_apply(doc, args) -> dict:
    root = numbering_root(doc)
    if root is None:
        raise SystemExit("document has no numbering part; run `define` first")
    _, num_ids = _existing_ids(root)
    if args.num_id not in num_ids:
        raise SystemExit(f"no such numId: {args.num_id} "
                         f"(known: {sorted(num_ids)})")

    paras = select(doc, args)
    if not paras:
        raise SystemExit("selection matched no paragraphs")
    for p in paras:
        set_numpr(p, args.num_id, args.ilvl)
    return {"ok": True, "applied": len(paras), "numId": args.num_id,
            "ilvl": args.ilvl,
            "first_text": paras[0].text[:40] if paras else ""}


def cmd_detach(doc, args) -> dict:
    paras = select(doc, args)
    n = sum(1 for p in paras if clear_numpr(p))
    return {"ok": True, "detached": n, "selected": len(paras)}


def cmd_restart(doc, args) -> dict:
    """Clone the definition behind *num_id* into a fresh counter.

    Word continues a list when two paragraphs share a numId. A new numId on
    the same abstractNum starts a new counter, so the next item begins again
    at ``--start-value``.
    """
    root = numbering_root(doc)
    if root is None:
        raise SystemExit("document has no numbering part")
    _, num_ids = _existing_ids(root)

    src = None
    for e in root.findall(qn("w:num")):
        if int(e.get(qn("w:numId"))) == args.num_id:
            src = e
            break
    if src is None:
        raise SystemExit(f"no such numId: {args.num_id}")
    aid_el = src.find(qn("w:abstractNumId"))
    aid = aid_el.get(qn("w:val"))

    nid = 1
    while nid in num_ids:
        nid += 1
    num_el = OxmlElement("w:num")
    num_el.set(qn("w:numId"), str(nid))
    ref = OxmlElement("w:abstractNumId")
    ref.set(qn("w:val"), aid)
    num_el.append(ref)
    if args.start_value is not None:
        ov = OxmlElement("w:lvlOverride")
        ov.set(qn("w:ilvl"), "0")
        so = OxmlElement("w:startOverride")
        so.set(qn("w:val"), str(args.start_value))
        ov.append(so)
        num_el.append(ov)
    root.append(num_el)

    paras = select(doc, args)
    if not paras:
        raise SystemExit("selection matched no paragraphs")
    for p in paras:
        set_numpr(p, nid, args.ilvl)
    return {"ok": True, "newNumId": nid, "abstractNumId": aid,
            "applied": len(paras)}


# ---- CLI --------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Manage Word auto-numbering in a .docx.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("path", help="input .docx")
        p.add_argument("-o", "--output",
                       help="output path (default: overwrite input)")

    def sel(p):
        p.add_argument("--index", action="append",
                       help="body paragraph index or range, e.g. 155:171 "
                            "(repeatable)")
        p.add_argument("--style", help="every paragraph using this style name")
        p.add_argument("--match", help="every paragraph containing this text")

    p = sub.add_parser("list", help="show numbering definitions and usage")
    common(p)

    p = sub.add_parser("define", help="append a definition and counter")
    common(p)
    p.add_argument("--text", default="%1.",
                   help="lvlText; %%1 is the counter, e.g. '[%%1]' (default '%%1.')")
    p.add_argument("--fmt", default="decimal", help="numFmt (default decimal)")
    p.add_argument("--start", type=int, default=1)
    p.add_argument("--jc", default="left", help="lvlJc (default left)")
    p.add_argument("--indent-left", type=int, default=0,
                   help="w:ind/@w:left in twips")
    p.add_argument("--hanging", type=int, default=0,
                   help="w:ind/@w:hanging in twips")

    p = sub.add_parser("apply", help="attach numPr to paragraphs")
    common(p)
    sel(p)
    p.add_argument("--num-id", type=int, required=True)
    p.add_argument("--ilvl", type=int, default=0)

    p = sub.add_parser("detach", help="strip numPr from paragraphs")
    common(p)
    sel(p)

    p = sub.add_parser("restart", help="restart numbering at a paragraph")
    common(p)
    sel(p)
    p.add_argument("--num-id", type=int, required=True)
    p.add_argument("--ilvl", type=int, default=0)
    p.add_argument("--start-value", type=int, default=1)

    args = ap.parse_args()
    if args.cmd != "list" and not any(
            getattr(args, k, None) for k in ("index", "style", "match")) \
            and args.cmd in ("apply", "detach", "restart"):
        ap.error("one of --index / --style / --match is required")

    # The input file is validated here — before `Document()` and before any
    # sub-command — because a broken package used to be a bare traceback with
    # exit 1 (measured 2026-11-11: 19 lines for a missing/non-zip file,
    # `IsADirectoryError` for a directory). Exit 2 is the project-wide code for
    # "you named something that is not there / not usable as given"; exit 1
    # stays reserved for "the document is real but the selection matched
    # nothing" (see the module docstring's exit-code table). stdout stays
    # empty: `list` prints JSON on success, so a failure must not look like an
    # empty document with no numbering. `doc.save()` is far below, so a refused
    # run writes nothing.
    try:
        _check_input(args.path)
        doc = _open_document(args.path)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    result = {
        "list": lambda: cmd_list(doc),
        "define": lambda: cmd_define(doc, args),
        "apply": lambda: cmd_apply(doc, args),
        "detach": lambda: cmd_detach(doc, args),
        "restart": lambda: cmd_restart(doc, args),
    }[args.cmd]()

    if args.cmd != "list":
        out = args.output or args.path
        doc.save(out)
        result.update({"output": out})
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
