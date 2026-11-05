r"""Insert native Office Math (OMML) into an unpacked .docx from LaTeX.

Engine
------
LaTeX --(latex2mathml)--> MathML --(Microsoft ``MML2OMML.XSL`` via lxml)--> OMML
（转换逻辑在 `math_omml.py`，与 `docx_mathify.py` 共用）

``MML2OMML.XSL`` 随 Microsoft Office 安装，是纯 XSLT 1.0，由 libxslt 直接应用 ——
**不需要启动 Office、不需要 pandoc、不需要 LaTeX 发行版**。

Requires
--------
    python -m pip install latex2mathml      # pure Python, zero deps, offline
    # MML2OMML.XSL 自动探测；可用 $env:MML2OMML_XSL 覆盖

Usage
-----
    python insert_formula.py unpacked/ "E=mc^2" --after "marker"
    python insert_formula.py unpacked/ "\frac{-b}{2a}" --display --after "marker"

批量把整篇文档里的线性公式换掉 → 用 `docx_mathify.py`，本脚本只管单条插入。

三条实测出来的硬规则
--------------------
1. 行间公式走 ``m:oMathPara``（脚本内部显式 ``display="block"``），否则退化成行内。
2. 行内公式独占整段时 Word 会当行间公式渲染 → 补一个 U+200B run 顶住
   （pandoc issue #11674 的官方修法）。
3. ``\quad \, \; \: \!`` 与普通空格都会被 XSLT 丢弃 → 默认改写成 ``\ ``（NBSP）。

History (2026-10-23)
--------------------
重写。旧版手写 OMML 构造器产出 schema 非法结构（``<m:oMath>`` 被嵌进
``m:num``/``m:den``/``m:e``/``m:sub``/``m:sup``/``m:deg``、``m:rad`` 缺 ``m:e``、
``_nary()`` 丢 ``naryPr`` 导致 ``\sum`` 渲染成孤立 ``∫``），Word 会判整篇 docx
「文件可能已经损坏」（错误码 25272）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import os as _os, sys as _sys; _sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix
from docx_utils import (
    ET,
    fix_root,
    insert_near,
    mkrun,
    parse,
    regm,
    register_all,
    write,
    wtag,
)
from math_omml import M_NS, latex_to_omml_lxml


def _omml_stdlib(latex: str, display: bool, raw_spacing: bool = False) -> ET.Element:
    """lxml 元素 → stdlib ElementTree 元素（本脚本在 stdlib ET 树上改 document.xml）。"""
    from lxml import etree
    el = latex_to_omml_lxml(latex, display=display, raw_spacing=raw_spacing)
    return ET.fromstring(etree.tostring(el, encoding="unicode"))


def build_para(latex: str, display: bool, raw_spacing: bool = False) -> ET.Element:
    """构造承载公式的 ``<w:p>``。"""
    omml = _omml_stdlib(latex, display, raw_spacing)
    para = ET.Element(wtag("p"))
    if display:
        ppr = ET.SubElement(para, wtag("pPr"))
        ET.SubElement(ppr, wtag("jc"), {wtag("val"): "center"})
        para.append(omml)                      # m:oMathPara
    else:
        para.append(omml)                      # m:oMath
        para.append(mkrun("\u200b"))           # 规则 2：防止被提升为行间公式
    return para


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Insert native Office Math (OMML) into an unpacked .docx "
                    "from LaTeX (latex2mathml + Microsoft MML2OMML.XSL)."
    )
    p.add_argument("unpacked_dir")
    p.add_argument("formula")
    p.add_argument("--display", action="store_true",
                   help="行间公式（产出 m:oMathPara，居中）")
    p.add_argument("--after", metavar="MARKER", help="插在该文字段落之后")
    p.add_argument("--before", metavar="MARKER", help="插在该文字段落之前")
    p.add_argument("--raw-spacing", action="store_true",
                   help=r"不改写 \quad \, \; \: \!（注意：这些间距会被 XSLT 丢掉）")
    p.add_argument("--dump", action="store_true", help="只打印生成的 OMML，不写文件")
    args = p.parse_args(argv)

    register_all()
    regm()
    ET.register_namespace("m", M_NS)

    try:
        para = build_para(args.formula, args.display, args.raw_spacing)
    except (RuntimeError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    if args.dump:
        print(ET.tostring(para, encoding="unicode"))
        return 0

    doc_xml = Path(args.unpacked_dir) / "word" / "document.xml"
    if not doc_xml.exists():
        print(f"error: {doc_xml} not found", file=sys.stderr)
        return 1

    tree = parse(doc_xml)
    root = tree.getroot()
    body = root.find(wtag("body"))
    if body is None:
        print("error: <w:body> not found", file=sys.stderr)
        return 1

    if args.after and args.before:
        print("error: --after and --before are mutually exclusive", file=sys.stderr)
        return 1

    marker = args.after or args.before
    position = "after" if args.after else "before"
    anchored = insert_near(body, marker, position, para)
    if marker and not anchored:
        # Do not silently drop the formula at the end of the document: the
        # caller asked for a specific anchor, and a misplaced equation is harder
        # to notice than a failed command.
        print(f"error: marker not found: {marker!r}", file=sys.stderr)
        return 1

    write(tree, doc_xml)
    fix_root(doc_xml)
    print(f"Inserted {'display' if args.display else 'inline'} formula: {args.formula}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
