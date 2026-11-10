# -*- coding: utf-8 -*-
r"""fix_fonts.py —— docx 字体统一（v2 样式层 + v3 手动大标题 合并版，就地改写 .docx）

解决什么问题
------------
中英混排 docx 的字体乱象：正文该宋体/标题该微软雅黑却五花八门，公式（OMML）
的西文字体被 eastAsia 设置连带污染。本工具一次跑完三层：

1. **样式层**（v2）：``Normal`` → eastAsia=宋体；``Heading*``/``Title`` → eastAsia=微软雅黑；
   ``docDefaults`` → 宋体。样式改了，正文跟着样式走的部分自动对齐。
2. **run 层**（v2）：所有 ``w:r`` 按所在段落是否 heading 分派 eastAsia
   （heading/Title → 雅黑，否则宋体）；ascii/hAnsi/cs 一律 Times New Roman。
   **公式 run（``m:oMath`` 内）只改 ascii/hAnsi/cs=TNR、不动 eastAsia**——
   公式的中文/字形由 OMML 自己管理，eastAsia 若被覆盖会污染公式观感。
3. **手动大标题**（v3）：``w:sz >= --heading-threshold``（默认 14pt，内部 ×2 成半磅）
   **且段落无 ``w:pStyle``** 的 run → eastAsia=微软雅黑。这是给"没套标题样式、
   只是手动把字号放大"的独立标题段兜底；「无 pStyle」是判据核心——套了样式的
   段落归样式层管，免得双重分派。

接口
----
::

    python fix_fonts.py <docx> [<docx2> ...] [--heading-threshold 14]

就地改写（临时文件 + shutil.move，先写后挪，中断不损原文件）。可一次传多个文件。

两条铁律（硬约束 #42；违反任意一条 Word 就报「文件可能已损坏」）
--------------------------------------------------------------
1. ``dedupe_rpr()``：``w:rPr`` 内**同名子元素只留最后一个**——OOXML 对 rPr 的
   子元素是 xsd:sequence 且多数 maxOccurs=1，重复即 schema 非法；实测 Word/WPS
   打开含重复 ``w:rFonts`` 的文件直接弹修复提示。
2. ``set_fonts()``：删旧 ``w:rFonts`` 后必须 ``rpr.insert(0, rf)``——
   **``w:rFonts`` 必须是 ``w:rPr`` 的第一个子元素**（CT_RParaProps 的序列首位）。
   ``append`` 放到尾部 Word 判非法。v2/v3 各踩一次才通。

踩过的坑
--------
- zip 就地改写必须"全部读入内存 → 写临时文件 → shutil.move"；边读边写同一
  zip 会把自己截断。
- 阈值单位：界面给 pt，XML 里 ``w:sz`` 是**半磅**（14pt = 28）；不做 ×2 换算
  就会把 7pt 当 14pt，全文档误判。
- ``m:oMath`` 归属判定靠 id() 集合一次收齐：``doc.iter`` 嵌套调用时同一元素
  会进多次集合，用 id() 去重正合适；判定要在遍历 run **之前**建好集合，
  遍历中再收集会漏。
- lxml 序列化时 ``standalone=True`` 是 Word 原文件的形态，缺了也不致命，
  但保持一致让 diff 最小。
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import _bootstrap  # noqa: F401 - flat-layout bootstrap + minidom / UTF-8 fallbacks

from lxml import etree

# D-008: reuse the remote root-element repair so the bare-zipfile output gets
# the same full xmlns declaration set + rewritten mc:Ignorable that
# docx_utils.write() applies automatically to every part it serializes.
from docx_utils import fix_root_any

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


# ---- 铁律 1：rPr 内同名子元素只留最后一个 ----
def dedupe_rpr(rpr) -> None:
    by_tag: dict[str, list] = {}
    for child in list(rpr):
        by_tag.setdefault(child.tag, []).append(child)
    for tag, els in by_tag.items():
        for e in els[:-1]:
            rpr.remove(e)


# ---- 铁律 2：rFonts 必须是 rPr 第一个子元素（insert(0)，不是 append）----
def set_fonts(rpr, ea: str) -> None:
    for old in rpr.findall(q("rFonts")):
        rpr.remove(old)
    rf = etree.Element(q("rFonts"))
    rf.set(q("ascii"), "Times New Roman")
    rf.set(q("hAnsi"), "Times New Roman")
    rf.set(q("eastAsia"), ea)
    rf.set(q("cs"), "Times New Roman")
    rpr.insert(0, rf)


# ---- D-008：根元素声明收尾（import 复用 docx_utils.fix_root_any）----
def _fix_root_bytes(data: bytes) -> bytes:
    """Apply fix_root_any() to a serialized XML part via a temp file.

    fix_root_any() 吃磁盘路径（读文本→改根元素→写回），本脚本是
    lxml→blobs→zip 内存管线，故用临时文件中转一次：把序列化结果落盘、
    修复根元素 xmlns 声明集 + mc:Ignorable，再读回 bytes 进 zip。
    """
    fd, name = tempfile.mkstemp(suffix=".xml", prefix="fontfix_root_")
    os.close(fd)
    tmp = Path(name)
    try:
        tmp.write_bytes(data)
        fix_root_any(tmp)
        return tmp.read_bytes()
    finally:
        tmp.unlink(missing_ok=True)


def _fix_docx(path: Path, threshold_half_pt: int) -> None:
    src = path
    tmp = src.with_suffix(".fontfix.tmp.docx")

    with zipfile.ZipFile(src) as zin:
        names = zin.namelist()
        blobs = {n: zin.read(n) for n in names}

    doc = etree.fromstring(blobs["word/document.xml"])
    styles = etree.fromstring(blobs["word/styles.xml"])

    # ---- 样式层（v2）----
    for st in styles.iter(q("style")):
        sid = st.get(q("styleId")) or ""
        if sid == "Normal":
            ea = "宋体"
        elif sid.lower().startswith("heading") or sid == "Title":
            ea = "微软雅黑"
        else:
            continue
        rpr = st.find(q("rPr"))
        if rpr is None:
            rpr = etree.SubElement(st, q("rPr"))
        dedupe_rpr(rpr)
        set_fonts(rpr, ea)

    for dd in styles.iter(q("docDefaults")):
        rpr = dd.find(f"{q('rPrDefault')}/{q('rPr')}")
        if rpr is None:
            continue
        dedupe_rpr(rpr)
        set_fonts(rpr, "宋体")

    # ---- 公式归属：先建 id 集合，再遍历 run ----
    math_ids: set[int] = set()
    for om in doc.iter(f"{{{M}}}oMath"):
        for el in om.iter():
            math_ids.add(id(el))

    def para_is_heading(p) -> bool:
        ppr = p.find(q("pPr"))
        if ppr is None:
            return False
        ps = ppr.find(q("pStyle"))
        if ps is None:
            return False
        val = (ps.get(q("val")) or "").lower()
        return val.startswith("heading") or val == "title"

    # ---- run 层（v2）+ 手动大标题（v3）----
    n_body = n_head = n_math = n_manual = 0
    for p in doc.iter(q("p")):
        head = para_is_heading(p)
        ea = "微软雅黑" if head else "宋体"
        ppr = p.find(q("pPr"))
        has_style = ppr is not None and ppr.find(q("pStyle")) is not None
        for r in p.findall(q("r")):
            if id(r) in math_ids:
                # 公式 run：只管 ascii/hAnsi/cs=TNR，eastAsia 不动
                rpr = r.find(q("rPr"))
                if rpr is not None:
                    for rf in rpr.findall(q("rFonts")):
                        rf.set(q("ascii"), "Times New Roman")
                        rf.set(q("hAnsi"), "Times New Roman")
                        rf.set(q("cs"), "Times New Roman")
                        if rf.get(q("eastAsia")) is None:
                            rf.set(q("eastAsia"), "Times New Roman")
                    n_math += 1
                continue
            rpr = r.find(q("rPr"))
            if rpr is None:
                rpr = etree.Element(q("rPr"))
                r.insert(0, rpr)
            # v3：手动放大（sz >= 阈值）且段落无 pStyle → 当标题
            is_manual = False
            sz = rpr.find(q("sz"))
            if sz is not None:
                try:
                    if int(sz.get(q("val"))) >= threshold_half_pt and not has_style:
                        is_manual = True
                except (TypeError, ValueError):
                    pass
            dedupe_rpr(rpr)
            set_fonts(rpr, "微软雅黑" if (head or is_manual) else "宋体")
            if head:
                n_head += 1
            elif is_manual:
                n_manual += 1
            else:
                n_body += 1

    blobs["word/document.xml"] = _fix_root_bytes(etree.tostring(
        doc, xml_declaration=True, encoding="UTF-8", standalone=True))
    blobs["word/styles.xml"] = _fix_root_bytes(etree.tostring(
        styles, xml_declaration=True, encoding="UTF-8", standalone=True))

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm in names:
            zout.writestr(nm, blobs[nm])

    shutil.move(str(tmp), str(src))
    print(f"  正文 {n_body}  标题 {n_head}  手动大标题 {n_manual}  公式 {n_math}   [{src.name}]")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Unify fonts in a .docx (styles + runs + manual headings)")
    ap.add_argument("docx", nargs="+", help=".docx file(s) to fix in place")
    ap.add_argument("--heading-threshold", type=float, default=14.0,
                    help="minimum font size in pt for a styleless run to be "
                         "treated as a manual heading (default: 14)")
    args = ap.parse_args()
    threshold_half_pt = int(round(args.heading_threshold * 2))  # pt → 半磅
    for arg in args.docx:
        p = Path(arg)
        if not p.exists():
            print(f"error: {p} not found", file=sys.stderr)
            return 1
        _fix_docx(p, threshold_half_pt)
    return 0


if __name__ == "__main__":
    sys.exit(main())
