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

import _bootstrap  # noqa: F401
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
from fix_phasor_dot import fix_phasor_dots

# 上一次 _omml_stdlib() 修掉的相量点数（CLI 用来打印，库调用方也可直接读）。
_LAST_PHASOR_FIXES = 0
# 上一次相量点修复抛出的异常描述（None = 没出错）。
_LAST_PHASOR_ERROR: str | None = None


def _omml_stdlib(latex: str, display: bool, raw_spacing: bool = False) -> ET.Element:
    r"""lxml 元素 → stdlib ElementTree 元素（本脚本在 stdlib ET 树上改 document.xml）。

    转换完**立刻**修相量点：``latex2mathml`` 把 ``\dot{X}`` 转成 ``m:limUpp``，
    Word 里是「大字 + 飘着的点」，正确形态是 ``m:acc``（``\u0307`` 组合点）。
    不挂这一刀，``\dot{I}`` 就会以丑态落进交付物。

    异常语义：``latex_to_omml_lxml`` 失败仍向上抛（调用方按既有逻辑报错退出）；
    ``fix_phasor_dots`` 失败**只在本函数内**捕获——公式已经转好了，相量点没修
    只影响观感，没有理由让整条公式作废。失败信息留在 ``_LAST_PHASOR_ERROR``，
    由 CLI 打到 stderr。

    统计出口：修了几处写进模块级 ``_LAST_PHASOR_FIXES``（CLI 打印「修正相量点
    N 处」，库调用方也可直接读该变量）。
    """
    global _LAST_PHASOR_FIXES, _LAST_PHASOR_ERROR
    from lxml import etree
    el = latex_to_omml_lxml(latex, display=display, raw_spacing=raw_spacing)
    _LAST_PHASOR_FIXES = 0
    _LAST_PHASOR_ERROR = None
    try:
        _LAST_PHASOR_FIXES = fix_phasor_dots(el)
    except Exception as exc:            # 修相量点出事不该让整条公式丢掉
        _LAST_PHASOR_ERROR = f"{type(exc).__name__}: {str(exc)[:160]}"[:200]
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
        if _LAST_PHASOR_FIXES:
            print(f"# 修正相量点 {_LAST_PHASOR_FIXES} 处（m:limUpp → m:acc）",
                  file=sys.stderr)
        if _LAST_PHASOR_ERROR:
            print(f"# warning: 相量点修复失败：{_LAST_PHASOR_ERROR}", file=sys.stderr)
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
    # 「修正相量点 N 处」只在真修了才追加，没有相量点的公式（常态）输出与旧版
    # 完全一致，不给下游解析多一个字段。
    note = f"  修正相量点 {_LAST_PHASOR_FIXES} 处" if _LAST_PHASOR_FIXES else ""
    print(f"Inserted {'display' if args.display else 'inline'} formula: "
          f"{args.formula}{note}")
    if _LAST_PHASOR_ERROR:
        print(f"warning: 相量点修复失败：{_LAST_PHASOR_ERROR}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
