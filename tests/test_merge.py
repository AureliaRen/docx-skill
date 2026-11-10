# -*- coding: utf-8 -*-
"""``merge_docs.py`` 回归：段落/图片守恒、sectPr 归 base、numbering 并集。

本轮的语义修复（裁定 2 / 裁定 3）必须有专门用例，否则下次重构又会静默改回去：

* **sectPr 统一用 base**：src body 末尾自带的 sectPr 丢弃。否则两份页面设置
  不同的文档合并后，版心会静默变成 src 的（原实现注释写"页边距统一"、行为却是
  采用 src 的，实现与 docstring 相反）。
* **numbering 按 ``w:numId`` 并集补**：src 独有的 ``w:num`` 补进 base，其引用的
  ``w:abstractNum`` 一并补齐一层。否则 src 的 ``w:numPr`` 指向不存在的 numId，
  Word 打开要么编号错乱要么静默丢编号。
* **base 无 numbering.xml 而 src 有 → 整文件搬运**，同时补
  ``[Content_Types].xml`` 的 Override 与 ``document.xml.rels`` 的关系条目——
  只搬文件不补元数据，包直接非法。

段落数守恒的公式是 ``merged == base_p * 2 + 1``：2×原段（base 自己 + src 整段搬入）
加 1 个分页符段。图片数守恒只在 src 无图时成立（self-merge 会把图复制一份并
``_1`` 去重，那是正确行为，见最后一条用例）。
"""
from __future__ import annotations

import os
import struct
import subprocess
import sys
import zipfile
import zlib
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from lxml import etree  # noqa: E402

docx = pytest.importorskip("docx", reason="需要 python-docx 造 fixture")

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------- helpers
def run_merge(base: Path, src: Path, out: Path) -> subprocess.CompletedProcess:
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "merge_docs.py"),
         str(base), str(src), str(out)],
        capture_output=True, env=env, check=False,
    )


def run_validate(path: Path) -> subprocess.CompletedProcess:
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_validate.py"), str(path)],
        capture_output=True, env=env, check=False,
    )


def make_png(path: Path) -> None:
    """stdlib 造一张 2x2 红 PNG（照 hermes e2e 的做法，不依赖 Pillow）。"""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", 2, 2, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * 2 for _ in range(2))
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                     + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def body_of(path: Path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    return root.find(q("body"))


def counts(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(q("body"))
    return {
        "p": len(body.findall(q("p"))),
        "sectPr": len(body.findall(q("sectPr"))),
        "media": sorted(n for n in names if n.startswith("word/media/")),
        "names": names,
    }


def sect_signature(path: Path):
    """sectPr 里 pgSz + pgMar 的属性快照（逐属性比较，不只看元素个数）。"""
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    body = root.find(q("body"))
    sect = body.find(q("sectPr"))
    if sect is None:
        return None
    pg = sect.find(q("pgSz"))
    pm = sect.find(q("pgMar"))
    return (dict(pg.attrib) if pg is not None else None,
            dict(pm.attrib) if pm is not None else None)


def make_base(path: Path, *, with_image: bool = False, tmp: Path | None = None) -> None:
    """A4 常规页边距的基底文档。"""
    d = docx.Document()
    d.add_paragraph("base paragraph one")
    d.add_paragraph("base paragraph two")
    s = d.sections[0]
    s.page_width = docx.shared.Mm(210)
    s.page_height = docx.shared.Mm(297)
    s.top_margin = docx.shared.Mm(25)
    s.bottom_margin = docx.shared.Mm(25)
    s.left_margin = docx.shared.Mm(20)
    s.right_margin = docx.shared.Mm(20)
    if with_image:
        assert tmp is not None
        png = tmp / "pic.png"
        make_png(png)
        d.add_picture(str(png), width=docx.shared.Mm(20))
    d.save(str(path))


def make_src(path: Path) -> None:
    """窄边距 + 不同纸张的源文档（跟 base 的页面设置刻意不同）。"""
    d = docx.Document()
    d.add_paragraph("src paragraph one")
    s = d.sections[0]
    s.page_width = docx.shared.Mm(148)
    s.page_height = docx.shared.Mm(210)
    s.top_margin = docx.shared.Mm(10)
    s.bottom_margin = docx.shared.Mm(10)
    s.left_margin = docx.shared.Mm(8)
    s.right_margin = docx.shared.Mm(8)
    d.save(str(path))


def repack(src: Path, dst: Path, mutate) -> None:
    """读全包 → 交给 mutate(names, blobs) 改 → 按原顺序写回。"""
    with zipfile.ZipFile(src) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    mutate(names, blobs)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, blobs[n])


def add_custom_numbering(names, blobs) -> None:
    """给 src 加 numId=42 → abstractNumId=99，并在正文里引用它。"""
    root = etree.fromstring(blobs["word/numbering.xml"])

    an = etree.SubElement(root, q("abstractNum"))
    an.set(q("abstractNumId"), "99")
    lvl = etree.SubElement(an, q("lvl"))
    lvl.set(q("ilvl"), "0")
    for tag, val in (("start", "1"), ("numFmt", "decimal"), ("lvlText", "%1.")):
        el = etree.SubElement(lvl, q(tag))
        el.set(q("val"), val)

    num = etree.SubElement(root, q("num"))
    num.set(q("numId"), "42")
    ai = etree.SubElement(num, q("abstractNumId"))
    ai.set(q("val"), "99")

    blobs["word/numbering.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)


def drop_numbering(names, blobs) -> None:
    """把 base 变成「没有 numbering.xml」的包：部件 + CT Override + rels 一起摘。"""
    del blobs["word/numbering.xml"]
    names.remove("word/numbering.xml")

    ct = etree.fromstring(blobs["[Content_Types].xml"])
    for o in list(ct.iter(f"{{{CT}}}Override")):
        if o.get("PartName") == "/word/numbering.xml":
            ct.remove(o)
    blobs["[Content_Types].xml"] = etree.tostring(
        ct, xml_declaration=True, encoding="UTF-8", standalone=True)

    rels = etree.fromstring(blobs["word/_rels/document.xml.rels"])
    for rel in list(rels):
        if (rel.get("Type") or "").endswith("/numbering"):
            rels.remove(rel)
    blobs["word/_rels/document.xml.rels"] = etree.tostring(
        rels, xml_declaration=True, encoding="UTF-8", standalone=True)


def numbering_ids(path: Path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/numbering.xml"))
    num_ids = [n.get(q("numId")) for n in root.iter(q("num"))]
    abs_ids = [a.get(q("abstractNumId")) for a in root.iter(q("abstractNum"))]
    return num_ids, abs_ids


# ---------------------------------------------------------------- 1. 段落守恒
def test_paragraph_count_conserved_self_merge(tmp_path: Path):
    """base 与自身合并 → ``merged == base_p * 2 + 1``（1 = 分页符段）。"""
    base = tmp_path / "base.docx"
    make_base(base)
    out = tmp_path / "out.docx"

    bp = counts(base)["p"]

    proc = run_merge(base, base, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    mp = counts(out)["p"]
    assert mp == bp * 2 + 1, f"段落数 {mp} != {bp}*2+1"


def test_page_break_paragraph_inserted(tmp_path: Path):
    """src 内容前必须有一个带 ``w:br w:type="page"`` 的段落。"""
    base, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base)
    make_src(src)
    out = tmp_path / "out.docx"

    assert run_merge(base, src, out).returncode == 0

    body = body_of(out)
    brs = [br for br in body.iter(q("br"))
           if br.get(q("type")) == "page"]
    assert len(brs) == 1, f"分页符数量 {len(brs)}，期望 1"


# ---------------------------------------------------------------- 2. 图片守恒
def test_media_count_conserved_when_src_has_no_image(tmp_path: Path):
    """base 含 1 张图、src 无图 → 合并后 media 仍是 1 个。"""
    base, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base, with_image=True, tmp=tmp_path)
    make_src(src)
    out = tmp_path / "out.docx"

    base_media = counts(base)["media"]
    assert len(base_media) == 1, f"fixture 前置假设失效：base media={base_media}"

    proc = run_merge(base, src, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    assert counts(out)["media"] == base_media, "图片数不守恒"


def test_image_rid_remapped_on_name_collision(tmp_path: Path):
    """两份都有同名 image1.png → 第二份改名 ``image1_1.png``，两条 rels 各自有效。

    rId 撞号是这个脚本的核心坑：两份文档的 rId1 各指向不同图，直接搬 body 会让
    src 的图指到 base 的图。所以这里同时断言「文件名去重」与「rels 解析」。
    """
    a = tmp_path / "a.docx"
    b = tmp_path / "b.docx"
    make_base(a, with_image=True, tmp=tmp_path)
    make_base(b, with_image=True, tmp=tmp_path)
    out = tmp_path / "out.docx"

    proc = run_merge(a, b, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    media = counts(out)["media"]
    assert len(media) == 2, f"撞名图没去重成两份：{media}"
    assert "word/media/image1.png" in media
    assert "word/media/image1_1.png" in media

    with zipfile.ZipFile(out) as z:
        rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
        rids = [r.get("Id") for r in rels.iter(f"{{{PR}}}Relationship")]
    assert len(rids) == len(set(rids)), f"rels 里有重复 rId：{rids}"


# ---------------------------------------------------------------- 3. sectPr 归 base
def test_sectpr_is_base_not_src(tmp_path: Path):
    """base A4/常规 vs src A6/窄边距 → 输出只有 1 份 sectPr 且逐属性等于 base。"""
    base, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base)
    make_src(src)
    out = tmp_path / "out.docx"

    base_sig = sect_signature(base)
    src_sig = sect_signature(src)
    assert base_sig != src_sig, "fixture 前置假设失效：base 与 src 页面设置相同"

    proc = run_merge(base, src, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    c = counts(out)
    assert c["sectPr"] == 1, f"输出有 {c['sectPr']} 份 sectPr，期望 1"

    out_sig = sect_signature(out)
    assert out_sig == base_sig, (
        "合并后页面设置不等于 base —— 裁定 2 被改回了「用 src」")
    assert out_sig != src_sig, "合并后页面设置等于 src（原 bug 行为）"


def test_sectpr_pgmar_attributes_match_base_exactly(tmp_path: Path):
    """逐属性比对 pgMar（不只比 pgSz），确认没有被 src 的值混进来。"""
    base, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base)
    make_src(src)
    out = tmp_path / "out.docx"
    assert run_merge(base, src, out).returncode == 0

    base_pgsz, base_pgmar = sect_signature(base)
    out_pgsz, out_pgmar = sect_signature(out)

    assert out_pgsz == base_pgsz
    for attr in ("top", "bottom", "left", "right"):
        assert out_pgmar.get(q(attr)) == base_pgmar.get(q(attr)), (
            f"pgMar/{attr} 不等于 base")


# ---------------------------------------------------------------- 4. numbering
def test_numbering_union_adds_src_num_and_abstract(tmp_path: Path):
    """src 独有 numId=42 → 输出 numIds 含 42，且 abstractNumId=99 依赖补齐。"""
    base, src_plain = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base)
    make_src(src_plain)

    src = tmp_path / "src42.docx"
    repack(src_plain, src, add_custom_numbering)

    out = tmp_path / "out.docx"
    proc = run_merge(base, src, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    num_ids, abs_ids = numbering_ids(out)
    assert "42" in num_ids, f"src 的自定义编号没并进来：numIds={num_ids}"
    assert "99" in abs_ids, f"abstractNum 依赖没补齐：abstractNumIds={abs_ids}"

    # base 原有的也没丢
    base_nums, _ = numbering_ids(base)
    assert set(base_nums) <= set(num_ids), "base 原有编号被覆盖了"


def test_numbering_carried_whole_when_base_has_none(tmp_path: Path):
    """base 无 numbering.xml → 整文件搬运 + CT Override + rels 关系。"""
    base_plain, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base_plain)
    make_src(src)

    base = tmp_path / "base_nonum.docx"
    repack(base_plain, base, drop_numbering)
    assert "word/numbering.xml" not in counts(base)["names"]

    out = tmp_path / "out.docx"
    proc = run_merge(base, src, out)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        ct = z.read("[Content_Types].xml").decode("utf-8")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")

    assert "word/numbering.xml" in names, "numbering 部件没搬进来"
    assert "numbering+xml" in ct, "Content_Types 缺 numbering Override"
    assert "relationships/numbering" in rels, "document.xml.rels 缺 numbering 关系"

    # 元数据补全之后包必须仍然合法
    proc = run_validate(out)
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")


# ---------------------------------------------------------------- 5. 产物可用
def test_merged_output_opens_and_validates(tmp_path: Path):
    """合并产物：python-docx 能开 + docx_validate ok:true。"""
    base, src = tmp_path / "base.docx", tmp_path / "src.docx"
    make_base(base, with_image=True, tmp=tmp_path)
    make_src(src)
    out = tmp_path / "out.docx"
    assert run_merge(base, src, out).returncode == 0

    docx.Document(str(out))     # 不抛即通过

    proc = run_validate(out)
    assert proc.returncode == 0, proc.stdout.decode("utf-8", "replace")
    assert '"ok": true' in proc.stdout.decode("utf-8").replace("'", '"')

    # 内容真的两份都在
    txt = "".join(t.text or "" for t in body_of(out).iter(q("t")))
    assert "base paragraph one" in txt
    assert "src paragraph one" in txt
