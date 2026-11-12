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

退出码契约（P6-FIX-C2 起明确，此前是隐含行为；P9-D 修订 ``1``/``2`` 的边界）
----------------------------------------------------------------------------
* **0** —— 成功（写入分支或 ``--dump`` 打印）。
* **2** —— **用法错误 / 输入不可用**：``--dump`` 与 ``--after``/``--before`` 同时给、
  ``--after``/``--before`` 给了空串、``--after`` 与 ``--before`` 同时给、
  LaTeX 转换环境缺失，以及 **``unpacked_dir`` 不可用**（目录不存在 / 里面没有
  ``word/document.xml`` / 那个 ``word/document.xml`` 不是良构 XML）。
  全部**在任何写盘之前**判定，文档一个字节都不动。
  ⚠️ P9-D 之前后三档（``word/document.xml`` 不存在、XML 不良构）报 **1**，
  其中 XML 不良构那条还是裸 traceback。它们是「你给的东西不是一个可用的解包
  .docx 目录」，与 `docx_edit` / `docx_read` / `docx_validate` 等兄弟 CLI 的
  输入错误同义，所以统一提到 2（全仓口径：输入不对 = 2）。
* **1** —— **文档是真实的、但内容有问题**：``<w:body>`` 缺失、marker 在文档里
  找不到。这两档都要求 ``word/document.xml`` 已经被成功解析，属于「合并/插入
  做不下去」，不再是「输入给错了」。

不静默说谎的两条硬约束（P6-FIX-C2，依据 P5-AUDIT-QUALITY §5.7）
--------------------------------------------------------------
1. **不许静默丢弃参数**：``--dump`` 只打印不写文件，锚点对它毫无意义 ——
   同时给就明确报错退 2，**不允许**像旧版那样先 ``return 0`` 把 ``--after`` 吞掉。
2. **不许静默降级落点**：带锚点时公式的落点只报**实测坐标**（``at body[i]`` /
   ``at tc[i]``，由插入后在树上反查该元素的下标得到），不报入参；
   marker 找不到时报错里给**实际可搜范围**（正文段落 0-based ``[0, N-1]`` + 单元格数）。

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


def _locate(body: ET.Element, para: ET.Element) -> tuple[str, int]:
    """回显**实际**落点：插入后到树上反查该段落的下标，不猜、不回显入参。

    ``insert_near()`` 的容器列表是 ``[body] + list(body.iter(w:tc))``，落点可能是
    ``<w:body>`` 的直接子元素，也可能在某个 ``<w:tc>`` 里。这里按**对象身份**
    （``is``）在同样的容器序列里找一遍：找到的下标就是真位置，找不到返回 ``("?", -1)``
    （宁可报"不知道"，也不编一个数字）。

    下标口径 = **容器自身直接子元素的下标**（含 ``w:sectPr``、``w:tbl``），
    即 ``body[i]`` 可与 ``<w:body>`` 的子元素序列直接对齐核对。
    """
    containers = [("body", body)] + [("tc", c) for c in body.iter(wtag("tc"))]
    for name, container in containers:
        for i, child in enumerate(list(container)):
            if child is para:
                return name, i
    return "?", -1


def _usage_error(msg: str) -> int:
    """用法错误统一退 2（与 argparse 自身的参数错误同码）。"""
    print(f"error: {msg}", file=sys.stderr)
    return 2


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Insert native Office Math (OMML) into an unpacked .docx "
                    "from LaTeX (latex2mathml + Microsoft MML2OMML.XSL)."
    )
    p.add_argument("unpacked_dir")
    p.add_argument("formula")
    p.add_argument("--display", action="store_true",
                   help="行间公式（产出 m:oMathPara，居中）")
    p.add_argument("--after", metavar="MARKER",
                   help="插在该文字段落之后（与 --dump 互斥）")
    p.add_argument("--before", metavar="MARKER",
                   help="插在该文字段落之前（与 --after / --dump 互斥）")
    p.add_argument("--raw-spacing", action="store_true",
                   help=r"不改写 \quad \, \; \: \!（注意：这些间距会被 XSLT 丢掉）")
    p.add_argument("--dump", action="store_true",
                   help="只打印生成的 OMML，不写文件（与 --after/--before 互斥）")
    args = p.parse_args(argv)

    # ---- 用法级校验：全部先于 register_all/build_para/任何读写 ----
    # 旧版把 --after 的互斥检查放在 --dump 的 return 0 **之后**，于是
    # `--after X --dump` 会静默把 --after 吞掉（P5-AUDIT-QUALITY §5.7）。
    # 现在这类矛盾参数一律明确报错退 2，且文档一个字节都不动。
    if args.dump and (args.after is not None or args.before is not None):
        return _usage_error(
            "--dump cannot be combined with --after/--before: --dump only prints "
            "the OMML and never writes the document, so the anchor would be "
            "silently ignored. Drop --dump to insert, or drop the anchor to dump.")
    if args.after is not None and args.before is not None:
        return _usage_error("--after and --before are mutually exclusive")
    for _flag, _value in (("--after", args.after), ("--before", args.before)):
        if _value is not None and not _value.strip():
            return _usage_error(
                f"{_flag} was given an empty marker; it would silently append the "
                "formula at the end of the document instead of anchoring it. "
                "Omit the flag to append at the end, or pass a non-empty marker.")

    register_all()
    regm()
    ET.register_namespace("m", M_NS)

    # ---- 输入校验（P9-D）：`unpacked_dir` 不可用一律退 2，先于任何写盘 ----
    # 本脚本的输入**不是** .docx 包而是**解包后的目录**，所以不套 `_check_input`
    # （那个函数判的是「文件是不是 zip 包」，对着目录用会得出相反结论）。
    # 三档坏输入：目录不存在 / 目录里没有 word/document.xml / 那个 XML 不良构。
    # 前两档旧版已经是 `error:` 前缀 + 1 行，只差返回码（实测 rc=1）；
    # 第三档旧版是 16 行裸 traceback（`ET.ParseError`），这里收成同样的一行。
    #
    # ★ 为什么排在 `build_para()` **之前**（且只对写入分支生效）
    #   ① 语义：「你给的目录不对」不该被「本机没装 Office」抢先报出来。排在后面
    #      时，无 XSL 的机器上 `build_para()` 先炸成 exit 2 + ``MML2OMML_XSL``
    #      报错，输入问题永远看不到 —— 与 P6-FIX-C2 把用法校验提到 build_para
    #      之前同因。
    #   ② 可测：排在后面时这三档的用例在无 XSL 的机器上会以「同一个 rc=2、不同
    #      的原因」**假绿**，只能被迫挂 ``@needs_omml``；提前之后它们真的不依赖
    #      XSL，坏输入用例得以裸跑 pass。
    # ``--dump`` 不读文档（语义是「只打印」），所以整块跳过：它从不要求目录存在。
    tree = None
    body = None
    doc_xml = Path(args.unpacked_dir) / "word" / "document.xml"
    if not args.dump:
        if not doc_xml.exists():
            print(f"error: {doc_xml} not found", file=sys.stderr)
            return 2
        try:
            tree = parse(doc_xml)
        except Exception as exc:        # ET.ParseError / OSError（读不了）
            # 只取首行：ElementTree / lxml 的多行报错会把上下文一起打出来，
            # 这里刻意保持「stderr 恰好一行」的 CLI 契约。
            first = (str(exc).splitlines() or [""])[0]
            print(f"error: {doc_xml} is not well-formed XML: "
                  f"{type(exc).__name__}: {first}", file=sys.stderr)
            return 2
        body = tree.getroot().find(wtag("body"))
        if body is None:
            # <w:body> 缺失是「文档真实但内容不对」，仍是 exit 1，不是用法错误。
            print("error: <w:body> not found", file=sys.stderr)
            return 1

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

    marker = args.after if args.after is not None else args.before
    position = "after" if args.after is not None else "before"
    # 可搜范围必须在 insert_near **之前**数：insert_near 的 fallback 会先把公式段
    # 追加进 body 再返回 False，事后数出来的段数会多 1（自己就成了假报错）。
    n_body = sum(1 for c in list(body) if c.tag == wtag("p"))
    n_tc = len(list(body.iter(wtag("tc"))))
    anchored = insert_near(body, marker, position, para)
    where = _locate(body, para)          # ← 插入**之后**反查，落点是真值
    if marker is not None and not anchored:
        # Do not silently drop the formula at the end of the document: the
        # caller asked for a specific anchor, and a misplaced equation is harder
        # to notice than a failed command. 报错里给实际可搜范围，别让调用方猜。
        rng = f"0-based [0, {n_body - 1}]" if n_body else "document has none"
        print(f"error: marker not found: {marker!r} — searched {n_body} body "
              f"paragraph(s) ({rng}) and {n_tc} table cell(s); no paragraph "
              "text contains this marker", file=sys.stderr)
        return 1

    write(tree, doc_xml)
    fix_root(doc_xml)
    # 「修正相量点 N 处」只在真修了才追加，没有相量点的公式（常态）输出与旧版
    # 完全一致，不给下游解析多一个字段。
    note = f"  修正相量点 {_LAST_PHASOR_FIXES} 处" if _LAST_PHASOR_FIXES else ""
    # 落点回显（P6-FIX-C2）：只在**给了锚点**时追加，且报的是插入后反查到的实测
    # 下标（`at body[3]` / `at tc[1]`），不是入参。无锚点的追加路径 stdout 与
    # 旧版逐字节一致，不给既有下游解析加字段。
    at_note = f"  at {where[0]}[{where[1]}]" if marker is not None else ""
    print(f"Inserted {'display' if args.display else 'inline'} formula: "
          f"{args.formula}{note}{at_note}")
    if _LAST_PHASOR_ERROR:
        print(f"warning: 相量点修复失败：{_LAST_PHASOR_ERROR}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
