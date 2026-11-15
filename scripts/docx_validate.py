#!/usr/bin/env python3
# MIT License. Part of the Hermes docx skill.
"""Health-check a .docx package and report issues as JSON.

Usage: docx_validate.py file.docx

Checks (health-check tier, NOT full XSD schema validation):
  - the file is a readable zip and python-docx can open it
  - required package parts exist ([Content_Types].xml, document.xml)
  - every relationship in every .rels file resolves to a part in the
    package (dangling image/hyperlink/etc. rels are reported; external
    targets such as hyperlinks are skipped)
  - r:embed / r:id references in document.xml resolve to relationships
  - embedded images are non-empty and start with known magic bytes
    (PNG/JPEG/GIF/BMP/TIFF/EMF/WMF/SVG); no PIL required
  - paragraph and run style ids referenced by the document exist in
    styles.xml
  - **structural sanity** of document.xml: a ``<w:drawing/>`` with neither
    ``wp:inline`` nor ``wp:anchor``, and a ``<w:p>`` / ``<w:tbl>`` sitting
    directly inside a ``<w:r>``.  Both make Word refuse the whole file while
    every reference in it still resolves, so the reference checks above cannot
    see them.  This is *not* XSD validation — it is only the two shapes that
    are cheap to detect and fatal in practice.

``word/document.xml`` (and ``word/styles.xml``) failing to parse is reported as
one ``bad-document-xml`` / ``bad-styles-xml`` issue instead of a traceback: the
package is a real .docx, so it is a document problem (exit 1, JSON on stdout),
not a usage error.

Output: {"ok": bool, "issues": [{"severity": "error"|"warning", ...}]}
Exit code 1 when any error-severity issue is found (warnings exit 0).
Exit code 2 when the path is not a readable .docx at all (one line on stderr).
"""
from __future__ import annotations
import _bootstrap  # noqa: F401

import argparse
import json
import posixpath
import sys
import zipfile

from lxml import etree

# 输入校验的三件套由 `docx_read` 单一拥有（`InputError` / `_check_input`），
# 这里只 import、不复制判定 —— 与 `docx_edit.py` / `docx_mathify.py` 同形。
from docx_read import InputError, _check_input

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
# ``wp:`` — the WordprocessingDrawing namespace that owns ``wp:inline`` /
# ``wp:anchor``.  A ``<w:drawing/>`` with neither child is the shape Word
# refuses outright (it has no idea where or how big the object is).  The
# DrawingML ``a:`` namespace is deliberately *not* used: ``a:blip`` presence is
# not required for a legal drawing (charts, shapes, text boxes).
WP = ("http://schemas.openxmlformats.org/drawingml/2006/"
      "wordprocessingDrawing")

#: Children a ``<w:p>`` / ``<w:tbl>`` may never sit directly inside a run.
#: ``w:r`` holds run content (``w:t``, ``w:drawing``, ``w:br``, ``w:tab`` …);
#: a block-level element in there is XML that Word silently repairs or refuses,
#: and python-docx happily opens it — so no other check here would notice.
RUN_FORBIDDEN_CHILDREN = (f"{{{W}}}p", f"{{{W}}}tbl")

IMAGE_MAGIC = (
    b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff", b"GIF87a", b"GIF89a",
    b"BM", b"II*\x00", b"MM\x00*",
    b"\x01\x00\x00\x00",              # EMF
    b"\xd7\xcd\xc6\x9a", b"\x01\x00\x09\x00",  # WMF variants
    b"<?xml", b"<svg",
)


def _issue(issues, severity, code, detail):
    issues.append({"severity": severity, "code": code, "detail": detail})


def _rel_target(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def _rels_owner(rels_name: str) -> str:
    """``word/_rels/document.xml.rels`` -> ``word/document.xml``.

    The part a .rels file describes is named by the file stem, *not* by its
    parent directory.  Keeping the two straight matters because a package has
    many .rels files sharing one directory (``document``, ``header1``,
    ``footer1``, ``comments``, ``footnotes``, ``endnotes``, ``fontTable`` …);
    keying them all by ``word`` makes them overwrite each other, and the winner
    then depends on set-iteration order (i.e. on ``PYTHONHASHSEED``).

    ``_rels/.rels`` describes the package itself and maps to ``""``.
    """
    rels_dir = posixpath.dirname(rels_name)
    stem = posixpath.basename(rels_name)[: -len(".rels")]
    if posixpath.basename(rels_dir) != "_rels":
        # Not a conforming location; fall back to "the directory above".
        return posixpath.normpath(posixpath.join(rels_dir, "..")).rstrip(".") \
            .strip("/")
    parent = posixpath.dirname(rels_dir)
    if not stem:  # ``_rels/.rels`` -> the package root
        return ""
    return posixpath.join(parent, stem) if parent else stem


def validate(path: str) -> dict:
    issues: list[dict] = []

    try:
        zf = zipfile.ZipFile(path)
    except (OSError, zipfile.BadZipFile) as exc:
        _issue(issues, "error", "not-a-zip", str(exc))
        return {"ok": False, "issues": issues}

    names = set(zf.namelist())
    bad = zf.testzip()
    if bad is not None:
        _issue(issues, "error", "corrupt-member", f"CRC check failed: {bad}")

    for required in ("[Content_Types].xml", "word/document.xml"):
        if required not in names:
            _issue(issues, "error", "missing-part",
                   f"required part absent: {required}")
    if issues and any(i["severity"] == "error" for i in issues):
        return {"ok": False, "issues": issues}

    # --- relationships resolve ------------------------------------------
    # Keyed by the **part that owns the .rels file**, not by its parent
    # directory.  ``word/_rels/document.xml.rels`` belongs to
    # ``word/document.xml``; ``word/_rels/header1.xml.rels`` belongs to
    # ``word/header1.xml``.  Collapsing both to ``"word"`` (the old behaviour)
    # put seven parts into one dictionary slot, so whichever .rels was read last
    # won — and since ``names`` is a *set*, that order depends on
    # ``PYTHONHASHSEED``.  The symptom was a false ``unresolved-reference`` on
    # perfectly valid files, decided at random per process: measured on
    # ``source-snapshot/test_unpack/full_test_input.docx``, seeds 0/3/4/5
    # reported two errors while seeds 1/2 reported ``ok:true``.
    rel_ids_by_source: dict[str, dict] = {}
    for rels_name in sorted(names):
        if not rels_name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(zf.read(rels_name))
        except etree.XMLSyntaxError as exc:
            _issue(issues, "error", "bad-rels-xml", f"{rels_name}: {exc}")
            continue
        source_part = _rels_owner(rels_name)
        ids = {}
        for rel in root.iter(f"{{{PR}}}Relationship"):
            rid, target = rel.get("Id"), rel.get("Target", "")
            mode = rel.get("TargetMode", "Internal")
            ids[rid] = target
            if mode == "External":
                continue
            # ``source_part`` is the *owning part* (a file path, e.g.
            # ``word/document.xml``); ``_rel_target`` resolves against its
            # directory.  The old ``source_part + "/x"`` hack only existed
            # because the old key was a bare directory.
            resolved = _rel_target(source_part, target)
            if resolved not in names:
                _issue(issues, "error", "dangling-rel",
                       f"{rels_name}: {rid} -> {target} (missing part)")
        # Two .rels files never own the same part, so this is a plain insert —
        # but merge rather than overwrite, so a duplicated member (legal in a
        # ZIP, only the first is read) cannot silently drop ids.
        rel_ids_by_source.setdefault(source_part or "_package", {}).update(ids)

    # --- r:id / r:embed references in document.xml -----------------------
    # An XMLSyntaxError here means the *package* is real but its main part is
    # not well-formed — a document problem, not a usage error.  It is folded
    # into one issue so the JSON contract on stdout survives and the caller
    # gets ``ok:false`` + exit 1 instead of nineteen lines of lxml internals.
    # Every later check reads ``doc_root``, so this returns immediately.
    try:
        doc_root = etree.fromstring(zf.read("word/document.xml"))
    except etree.XMLSyntaxError as exc:
        _issue(issues, "error", "bad-document-xml",
               f"word/document.xml: {exc}")
        return {"ok": False, "issues": issues}
    doc_rels = rel_ids_by_source.get("word/document.xml", {})
    for el in doc_root.iter():
        for attr in (f"{{{R}}}id", f"{{{R}}}embed", f"{{{R}}}link"):
            rid = el.get(attr)
            if rid and rid not in doc_rels:
                _issue(issues, "error", "unresolved-reference",
                       f"document.xml references {rid} with no relationship")

    # --- structural sanity of the body ------------------------------------
    # Everything above resolves *references*; none of it looks at the shape of
    # the markup.  Two shapes are worth naming because Word rejects them while
    # the file still passes every check above:
    #
    #   * ``<w:drawing/>`` with no ``wp:inline`` / ``wp:anchor`` — the classic
    #     after-effect of deleting an inline shape while a picture part is
    #     replaced: the run keeps an empty drawing, and Word refuses to open
    #     the document with "unreadable content".
    #   * a block element (``<w:p>`` / ``<w:tbl>``) directly inside a run.
    #
    # Scope discipline: only these two, only in document.xml, and only when the
    # shape is unambiguous.  A drawing with ``wp:anchor`` but no ``wp:inline`` is
    # legal (floating objects), and a drawing inside ``mc:AlternateContent``
    # still has one of the two children — so neither is touched.
    for drawing in doc_root.iter(f"{{{W}}}drawing"):
        if (drawing.find(f"{{{WP}}}inline") is None
                and drawing.find(f"{{{WP}}}anchor") is None):
            _issue(issues, "error", "empty-drawing",
                   "document.xml: <w:drawing> has neither wp:inline nor "
                   "wp:anchor; Word refuses the file")

    for run in doc_root.iter(f"{{{W}}}r"):
        for child_tag in RUN_FORBIDDEN_CHILDREN:
            for child in run.iterchildren(child_tag):
                local = child_tag.split("}")[1]
                _issue(issues, "error", "invalid-nesting",
                       f"document.xml: <w:{local}> nested directly inside "
                       f"<w:r>; that is not valid run content")

    # --- embedded images decode ------------------------------------------
    for name in [n for n in names
                 if n.startswith("word/media/") and not n.endswith("/")]:
        data = zf.read(name)
        if not data:
            _issue(issues, "error", "empty-image", name)
        elif not any(data.startswith(m) for m in IMAGE_MAGIC):
            _issue(issues, "warning", "unknown-image-format",
                   f"{name}: unrecognized magic bytes")

    # --- styles referenced exist ------------------------------------------
    defined = set()
    if "word/styles.xml" in names:
        # Same contract as document.xml above: a broken styles part is a
        # document problem (issue + exit 1), never a traceback.
        try:
            styles_root = etree.fromstring(zf.read("word/styles.xml"))
        except etree.XMLSyntaxError as exc:
            _issue(issues, "error", "bad-styles-xml",
                   f"word/styles.xml: {exc}")
            styles_root = None
        if styles_root is not None:
            defined = {s.get(f"{{{W}}}styleId")
                       for s in styles_root.iter(f"{{{W}}}style")}
    for tag, attr in ((f"{{{W}}}pStyle", f"{{{W}}}val"),
                      (f"{{{W}}}rStyle", f"{{{W}}}val"),
                      (f"{{{W}}}tblStyle", f"{{{W}}}val")):
        for el in doc_root.iter(tag):
            sid = el.get(attr)
            if sid and sid not in defined:
                _issue(issues, "error", "missing-style",
                       f"style id referenced but not defined: {sid}")

    # --- python-docx can open it ------------------------------------------
    try:
        from docx import Document
        Document(path)
    except Exception as exc:
        _issue(issues, "error", "python-docx-open-failed", str(exc))

    ok = not any(i["severity"] == "error" for i in issues)
    return {"ok": ok, "issues": issues}


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Health-check a .docx (not XSD schema validation).")
    ap.add_argument("path", help="the .docx file to check")
    args = ap.parse_args()

    # ---- 双契约：输入错误退 2，内容问题仍退 1（P9-C）---------------------
    # 本脚本是全仓唯一一个「1 有正当语义」的 CLI：`ok:false` + `issues` 是它的
    # 正常产出，`1` 表示**文档是真的、但内容有 error**。所以这里只把「输入根本
    # 不是包」这三档提到 2，**内容那一档的 JSON 输出一个字都不动**。
    #
    # 旧行为（实测 2026-11-11）：三档全部落进 `validate()` 的 `zipfile.ZipFile`
    # except 分支，被收成一条 `not-a-zip` issue → **exit 1 + JSON**。于是
    # 「路径写错了」和「文档坏了」在调用方眼里是同一件事，唯一区别是 issue 的
    # code —— 而任何 `if rc: 报错` 的调用方都会把它当成文档问题去修文档。
    #
    # stdout 在拒绝路径上必须为空：成功与「内容有问题」两条路都会打 JSON，
    # 空的 stdout 才不会被 `json.loads` 读成一个空报告。
    try:
        _check_input(args.path)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    report = validate(args.path)
    print(json.dumps(report, ensure_ascii=False))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
