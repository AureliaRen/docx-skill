#!/usr/bin/env python3
# MIT License. Part of the docx skill.
"""Add the schema-required w:percent to <w:zoom> in a packed .docx.

python-docx writes ``<w:zoom w:val="bestFit"/>`` when it creates settings.xml.
The schema (CT_Zoom) declares ``w:percent`` as ``use="required"`` and ``w:val``
as optional, so strict validation rejects the file with:

    Element '{...}zoom': The attribute '{...}percent' is required but missing.

This script adds ``w:percent="100"`` in place, keeping any existing w:val.

Usage:
    python fix_zoom.py file.docx                 # repair in place
    python fix_zoom.py file.docx -o out.docx     # write elsewhere
    python fix_zoom.py file.docx --percent 120   # non-default zoom

解决什么问题
-----------
`docx_create.py` / `docx_template.py` 生成的文档里，settings.xml 的 `<w:zoom>`
只有 `w:val="bestFit"`。CT_Zoom 把 `w:percent` 声明成 `use="required"`，所以
严格 XSD 校验器（以及 Word 在部分修复路径下）会直接判文档损坏。本脚本补上这个
必需属性，已有 `w:val` 原样保留。

接口
----
``fix_settings_xml(xml: str, percent: int) -> tuple[str, int]``
    纯函数：拿 settings.xml 的文本，返回 (补好的文本, 改了几处)。`tests/`
    直接 import 它做断言，所以名字与签名不许改。

踩过的坑
--------
- 裸正则 ``r"<w:zoom\\s*/>"`` **抓不到** ``<w:zoom w:val="bestFit"/>``——带属性的
  形式一个都匹配不上，脚本"成功"了但文档依旧非法。必须是
  ``r"<w:zoom\\b([^>]*?)(/?)>"``，属性段和自闭合斜杠都兜住。
- 已经有 `w:percent` 的不能再加：重复属性会让 XML 直接不合法。命中即原样返回。
- zip 重写必须复用原 ``ZipInfo``（``zout.writestr(info, blob)``），条目顺序与
  压缩方式才和原包一致；按字典随手写会把 `[Content_Types].xml` 挪位。
- 没有 ``word/settings.xml`` 的包是合法的（settings 是可选部件），报
  "nothing to do" 并以 0 退出，不是错误。
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import _bootstrap  # noqa: F401 - flat-layout bootstrap + minidom / UTF-8 fallbacks

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# <w:zoom ...> in either self-closing or paired form.
# NOTE: matches attribute-bearing tags too — a bare r"<w:zoom\s*/>" silently
# skips <w:zoom w:val="bestFit"/> and leaves the document invalid.
ZOOM_RE = re.compile(r"<w:zoom\b([^>]*?)(/?)>", re.S)
PERCENT_ATTR_RE = re.compile(r'\bw:percent\s*=\s*"[^"]*"')


def fix_settings_xml(xml: str, percent: int) -> tuple[str, int]:
    """Return (patched_xml, number_of_zoom_elements_touched)."""
    touched = 0

    def repl(m: re.Match) -> str:
        nonlocal touched
        attrs, selfclose = m.group(1), m.group(2)
        if PERCENT_ATTR_RE.search(attrs):
            return m.group(0)                    # already valid
        touched += 1
        attrs = attrs.rstrip()
        if not attrs.startswith(" "):
            attrs = " " + attrs if attrs else ""
        return f'<w:zoom{attrs} w:percent="{percent}"{selfclose}>'

    return ZOOM_RE.sub(repl, xml), touched


def fix_docx(src: Path, dst: Path | None = None, percent: int = 100) -> int:
    """Programmatic entry (R-11): patch <w:zoom> in *src* (default: in place).

    Same behaviour and exit codes as the CLI; extracted so docx_create can
    call it right after ``doc.save()`` without going through argv. The pure
    ``fix_settings_xml`` is untouched (tests import it by name).
    """
    if dst is None:
        dst = src
    try:
        zin = zipfile.ZipFile(src)
    except (OSError, zipfile.BadZipFile) as exc:
        # An unreadable input path is a usage error (`2`), not "the document
        # says something I cannot act on" (`1`): the caller named a file that
        # is not there, or not a package.  Nothing has been written at this
        # point, so the refusal is total.
        print(f"error: {src} is not a readable zip: {exc}", file=sys.stderr)
        return 2
    with zin:
        names = zin.namelist()
        if "word/settings.xml" not in names:
            print("no word/settings.xml — nothing to do")
            return 0
        blobs = {n: zin.read(n) for n in names}
        infos = {i.filename: i for i in zin.infolist()}
    xml = blobs["word/settings.xml"].decode("utf-8")
    patched, touched = fix_settings_xml(xml, percent)
    if touched == 0:
        print("zoom element already valid — nothing to do")
        if dst != src:
            shutil.copyfile(src, dst)
        return 0

    blobs["word/settings.xml"] = patched.encode("utf-8")

    tmp = dst.with_suffix(dst.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in blobs:
            info = infos[name]
            zout.writestr(info, blobs[name])
    tmp.replace(dst)
    print(f'patched {touched} <w:zoom> element(s) with w:percent="{percent}" -> {dst}')
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Add required w:percent to <w:zoom>")
    ap.add_argument("docx")
    ap.add_argument("-o", "--output", default=None,
                    help="output path (default: patch in place)")
    ap.add_argument("--percent", type=int, default=100,
                    help="zoom percentage to write (default 100)")
    args = ap.parse_args()

    src = Path(args.docx)
    if not src.is_file():
        # Usage error, not "the document says something I cannot act on": the
        # caller named a path that is not a file, and nothing has been opened.
        why = "is a directory, not a file" if src.is_dir() else "not found"
        print(f"error: {src} {why}", file=sys.stderr)
        return 2
    dst = Path(args.output) if args.output else src
    return fix_docx(src, dst, args.percent)


if __name__ == "__main__":
    sys.exit(main())
