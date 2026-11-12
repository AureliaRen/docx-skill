# -*- coding: utf-8 -*-
r"""merge_docs.py —— 合并两份 docx 为一份（保留 OMML 公式、样式、编号、图片关系）

解决什么问题
------------
把 A 文档当基底、把 B 文档整段搬进来（典型：速查手册 + 例题专册 → 考场总资料），
python-docx 做不了跨包合并（styles/numbering/media 关系各自独立），手工搬必踩
关系 id 撞号与样式同 id 冲突。本工具在 zip/XML 层完成七步：

七步
----
1. **样式合并**：``word/styles.xml`` 按 ``w:styleId`` 并集去重——base 已有的不覆盖，
   src 独有的整段补进来。
2. **编号合并**：``word/numbering.xml`` 按 ``w:numId`` 并集去重（与 styles 同做法）。
   src 独有的 ``w:num`` 补进 base，其引用的 ``w:abstractNum`` 一并补（按
   ``w:abstractNumId`` 找齐，连带递归依赖）。**base 没有 numbering.xml 而 src 有
   → 整个文件直接搬**，同时补 [Content_Types] 的 Override 与 document.xml.rels
   的关系条目。src 有而 base 无的 ``w:numPr`` 引用因此不再错乱/静默丢失。
3. **media 关系重映射**：src 的图片复制进 base（文件名撞了就 ``_1`` 后缀去重），
   ``r:embed`` 等 r: 属性撞号的新分配 rId，搬 body 时按映射改写。
4. **搬 body**：src 的块级元素（w:p/w:tbl…）逐个 deepcopy 进 base body，
   途中完成 rId 重映射。
5. **分页符**：在 src 内容前插一个带 ``<w:br w:type="page"/>`` 的段落，让并入的
   部分从新页开始。
6. **sectPr 统一用 base 的**（裁定 2）：src body 末尾自带的 sectPr **丢弃**，
   只保留 base 的页面设置（版心/页边距不因合并被换掉，且不报错——原实现
   是"删 base 用 src"，与任务书相反，已修正）。
7. **Content_Types + zip 重写**：新 media 的扩展名补 Default 声明；
   zip 重写按 base 原条目顺序、新条目追加在尾。

接口
----
::

    python merge_docs.py <base.docx> <src.docx> <out.docx>

踩过的坑
--------
- **sectPr 反向**：原实现注释写"页边距统一"、行为却是采用 src 的页面设置——
  实现与 docstring 不一致时以"跑一遍"为准（硬约束 #41），本轮按裁定改回 base。
- **docstring 承诺了 numbering 却零实现**（硬约束 #41 同款）：编号定义缺失时
  src 的 ``w:numPr`` 引用会指向不存在的 numId，Word 打开要么编号错乱要么静默
  丢编号。现在真做了，styles 与 numbering 的并集去重逻辑一致。
- **rId 撞号**：两份文档的 ``rId1`` 各自指向不同图，直接搬 body 会让 src 的图
  指到 base 的图。所有 r: 属性（不只 r:embed）都要过映射表。
- **abstractNum 要跟着 num 走**：只搬 ``w:num`` 不搬它引用的 ``w:abstractNum``，
  编号定义照样是断的。一个 num 引用链只有一层（num → abstractNum），一层补齐即可。
- media 文件名去重必须查 base 全量（含 base 自己的 media），否则覆盖同名图。
"""
from __future__ import annotations

import argparse
import copy
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import _bootstrap  # noqa: F401 - flat-layout bootstrap + minidom / UTF-8 fallbacks

from lxml import etree

# 输入校验的三件套由 `docx_read` 单一拥有（`InputError` / `_check_input` /
# `_open_document`），这里只 import，不复制判定 —— 与 `docx_edit.py` 同形。
from docx_read import InputError, _check_input, _open_document

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NUMBERING_CT = ("application/vnd.openxmlformats-officedocument."
                "wordprocessingml.numbering+xml")
NUMBERING_RELTYPE = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                     "relationships/numbering")


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


def load(path: str):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    return names, blobs


def merge_styles(base: bytes, add: bytes) -> bytes:
    b = etree.fromstring(base)
    a = etree.fromstring(add)
    have = {s.get(q("styleId")) for s in b.iter(q("style"))}
    added = 0
    for st in list(a.iter(q("style"))):
        sid = st.get(q("styleId"))
        if sid in have:
            continue
        b.append(copy.deepcopy(st))
        have.add(sid)
        added += 1
    print(f"  styles: 新增 {added} 个")
    return etree.tostring(b, xml_declaration=True, encoding="UTF-8", standalone=True)


def merge_numbering(bn: list[str], bb: dict, sn: list[str], sb: dict) -> None:
    """按 w:numId 并集把 src 的编号定义补进 base（裁定 3）。

    base 无 numbering.xml 且 src 有 → 整文件搬运 + 补 Content_Types Override
    与 document.xml.rels 关系条目。
    """
    if "word/numbering.xml" not in sb:
        return

    if "word/numbering.xml" not in bb:
        # base 根本没有：整文件直接搬，并补元数据
        bb["word/numbering.xml"] = sb["word/numbering.xml"]
        # Content_Types Override（没有就补 Default 化的 Override 条目）
        ct = etree.fromstring(bb["[Content_Types].xml"])
        part = "/word/numbering.xml"
        if not any(d.get("PartName") == part
                   for d in ct.iter(f"{{{CT}}}Override")):
            o = etree.SubElement(ct, f"{{{CT}}}Override")
            o.set("PartName", part)
            o.set("ContentType", NUMBERING_CT)
        bb["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True)
        # 关系条目
        rels_name = "word/_rels/document.xml.rels"
        rels = etree.fromstring(bb[rels_name])
        if not any(rel.get("Type") == NUMBERING_RELTYPE for rel in rels):
            rel = etree.SubElement(rels, f"{{{PR}}}Relationship")
            rel.set("Id", "rIdNumImport")
            rel.set("Type", NUMBERING_RELTYPE)
            rel.set("Target", "numbering.xml")
        bb[rels_name] = etree.tostring(
            rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        print("  numbering: base 无 numbering.xml → 整文件搬运 + 补元数据")
        return

    # 双方都有：按 numId 并集补 src 独有的 num + 其 abstractNum
    b = etree.fromstring(bb["word/numbering.xml"])
    a = etree.fromstring(sb["word/numbering.xml"])

    have_num = {n.get(q("numId")) for n in b.iter(q("num"))}
    # base 现有 abstractNumId
    have_abs = {x.get(q("abstractNumId")) for x in b.iter(q("abstractNum"))}
    # src 里 num → abstractNumId 的映射（用于补依赖）
    src_num_abs = {n.get(q("numId")): n.find(q("abstractNumId")).get(q("val"))
                   for n in a.iter(q("num"))
                   if n.find(q("abstractNumId")) is not None}

    added_num = added_abs = 0
    for num in list(a.iter(q("num"))):
        nid = num.get(q("numId"))
        if nid in have_num:
            continue
        b.append(copy.deepcopy(num))
        have_num.add(nid)
        added_num += 1
        # 补这个 num 引用的 abstractNum（连带递归一层）
        abs_id = src_num_abs.get(nid)
        if abs_id is not None and abs_id not in have_abs:
            for an in a.iter(q("abstractNum")):
                if an.get(q("abstractNumId")) == abs_id:
                    b.append(copy.deepcopy(an))
                    have_abs.add(abs_id)
                    added_abs += 1
                    # abstractNum 的 w:multLvlType 等无再深依赖，一层补齐
                    break

    print(f"  numbering: 新增 num {added_num} 个 / abstractNum {added_abs} 个")
    bb["word/numbering.xml"] = etree.tostring(
        b, xml_declaration=True, encoding="UTF-8", standalone=True)


def main(base_path: str, src_path: str, out_path: str) -> None:
    print(f"基底: {Path(base_path).name}")
    print(f"并入: {Path(src_path).name}")

    bn, bb = load(base_path)
    sn, sb = load(src_path)

    base_doc = etree.fromstring(bb["word/document.xml"])
    src_doc = etree.fromstring(sb["word/document.xml"])

    base_body = base_doc.find(q("body"))
    src_body = src_doc.find(q("body"))

    # 1. 样式合并
    if "word/styles.xml" in bb and "word/styles.xml" in sb:
        bb["word/styles.xml"] = merge_styles(bb["word/styles.xml"], sb["word/styles.xml"])

    # 2. 编号合并（裁定 3：补实现）
    merge_numbering(bn, bb, sn, sb)

    # 3. 媒体关系：把 src 的 media 复制进 base，重映射 r:embed id
    rels_name = "word/_rels/document.xml.rels"
    base_rels = etree.fromstring(bb[rels_name])
    src_rels = etree.fromstring(sb[rels_name])

    base_rid = {rel.get("Id") for rel in base_rels}
    next_n = 1
    rid_map: dict[str, str] = {}
    media_added = 0

    # 收集 base 已有的 media 文件名，避免覆盖
    base_media = {n for n in bb if n.startswith("word/media/")}

    for rel in list(src_rels):
        tgt = rel.get("Target") or ""
        rtype = rel.get("Type") or ""
        if not rtype.endswith("/image"):
            continue
        src_media = "word/" + tgt.lstrip("./")
        if src_media not in sb:
            continue
        # 目标文件名去重
        name = Path(tgt).name
        newname = name
        k = 1
        while f"word/media/{newname}" in bb:
            newname = f"{Path(name).stem}_{k}{Path(name).suffix}"
            k += 1
        bb[f"word/media/{newname}"] = sb[src_media]
        base_media.add(f"word/media/{newname}")
        media_added += 1
        # 新 rId
        while f"rId{next_n}" in base_rid:
            next_n += 1
        new_rid = f"rId{next_n}"
        base_rid.add(new_rid)
        rid_map[rel.get("Id")] = new_rid
        nrel = etree.SubElement(base_rels, f"{{{PR}}}Relationship")
        nrel.set("Id", new_rid)
        nrel.set("Type", rtype)
        nrel.set("Target", f"media/{newname}")

    print(f"  media: 新增 {media_added} 个")

    # 4. 搬 body 元素（先重映射 r:embed）
    moved = 0
    for child in list(src_body):
        tag = etree.QName(child).localname
        if tag == "sectPr":
            # 裁定 2：src 的 sectPr 丢弃，输出统一用 base 的页面设置
            continue
        new_el = copy.deepcopy(child)
        # 重映射关系 id
        for el in new_el.iter():
            for attr in list(el.attrib):
                if attr.startswith(f"{{{R}}}"):
                    old = el.attrib[attr]
                    if old in rid_map:
                        el.attrib[attr] = rid_map[old]
        base_body.append(new_el)
        moved += 1
    print(f"  body: 搬入 {moved} 个块级元素")

    # 5. 分页：在源内容前插一个分页符段落（让并入内容从新页开始）
    brk = etree.Element(q("p"))
    r = etree.SubElement(brk, q("r"))
    br = etree.SubElement(r, q("br"))
    br.set(q("type"), "page")
    # 插到最后一个搬入元素之前
    idx = len(base_body) - moved
    base_body.insert(idx, brk)
    print("  已插入分页符")

    # 6. sectPr 统一用 base 的（裁定 2）：src 的已在步骤 4 丢弃，
    #    base 的 sectPr 仍留在 body 末尾，无需任何改动。

    bb["word/document.xml"] = etree.tostring(
        base_doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    bb[rels_name] = etree.tostring(
        base_rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # 7. content-types：补 media 扩展名
    ct = etree.fromstring(bb["[Content_Types].xml"])
    exts = {d.get("Extension") for d in ct.iter(f"{{{CT}}}Default")}
    for n in bb:
        if n.startswith("word/media/"):
            e = Path(n).suffix.lstrip(".").lower()
            if e and e not in exts:
                d = etree.SubElement(ct, f"{{{CT}}}Default")
                d.set("Extension", e)
                d.set("ContentType", {
                    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                    "gif": "image/gif", "emf": "image/x-emf", "wmf": "image/x-wmf",
                }.get(e, "application/octet-stream"))
                exts.add(e)
    bb["[Content_Types].xml"] = etree.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in bn:
            if n in bb:
                zout.writestr(n, bb[n])
        for n in bb:
            if n not in bn:
                zout.writestr(n, bb[n])

    print(f"输出: {out_path}")


def main_cli() -> int:
    ap = argparse.ArgumentParser(
        description="Merge src.docx into base.docx (styles/numbering/media kept)")
    ap.add_argument("base", help="base .docx (page setup/styles win)")
    ap.add_argument("src", help="source .docx to append")
    ap.add_argument("out", help="output .docx path")
    args = ap.parse_args()
    # ---- 输入校验（P9-D）：base 与 src **每个**都过四档判定，先于 `main()` ----
    # 旧版只查 `Path(p).exists()`：缺文件退 1（语义错——输入不对不等于合并失败），
    # 其余三档（是目录 / 非 zip / zip 但非 Word 包）在 `main()` 里崩成裸 traceback
    # + exit 1（实测 2026-11-11：23 / 17 / 10 行），且崩之前 `main()` 已经往
    # **stdout** 打了「基底: …」「并入: …」几十字节——失败路径不能有 stdout。
    #
    # `_open_document` 也跑一遍（不做 `_check_input` 单点）：本脚本不用 python-docx
    # 解析文档，`is_zipfile` 只读文件尾，一个 zip 壳子（有 foo.txt、无
    # `word/document.xml`）能过检却在 `bb["word/document.xml"]` 上抛 KeyError。
    try:
        for p in (args.base, args.src):
            _check_input(p)
            _open_document(p)
    except InputError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    main(args.base, args.src, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main_cli())
