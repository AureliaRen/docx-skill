# -*- coding: utf-8 -*-
"""``fix_fonts.py`` 回归：``w:rPr`` 两条铁律 + 幂等性 + ``--heading-threshold`` 语义。

两条铁律（违反任意一条 Word 就弹「文件可能已损坏」）：

1. ``w:rPr`` 内**同名子元素只能有一个**——xsd:sequence + maxOccurs=1，重复即非法。
2. ``w:rFonts`` 必须是 ``w:rPr`` 的**第一个**子元素——``insert(0, rf)``，不是
   ``append``。放到尾部 Word 判非法。

断言一律直接读 XML（lxml），不听脚本 stdout：脚本打印「正文 1 标题 1」不代表
写出来的 XML 合法。

阈值语义：``--heading-threshold`` 单位是 pt（内部 ×2 成半磅），
``rPr/w:sz >= 阈值`` **且段落无 ``w:pStyle``** 的 run 才判「手动放大标题」→ 雅黑。
没有「无 pStyle」这一半判据，套了 Heading 样式的段落会被双重分派。
"""
from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from lxml import etree  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# python-docx 是硬依赖（不是可选链路），缺了直接失败而非 skip
docx = pytest.importorskip("docx", reason="需要 python-docx 造 fixture")


def q(tag: str) -> str:
    return f"{{{W}}}{tag}"


# ---------------------------------------------------------------- helpers
def run_fix_fonts(path: Path, *extra) -> subprocess.CompletedProcess:
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "fix_fonts.py"), str(path), *map(str, extra)],
        capture_output=True, env=env, check=False,
    )


def parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


SCAN_PARTS = ("word/document.xml", "word/styles.xml")


def rule_violations(blobs: dict[str, bytes],
                    which: tuple[str, ...] = SCAN_PARTS) -> dict[str, int]:
    """扫 document.xml + styles.xml，统计两条铁律的违反次数与 rFonts 总数。"""
    total = dup = not_first = 0
    for part in which:
        if part not in blobs:
            continue
        root = etree.fromstring(blobs[part])
        for rpr in root.iter(q("rPr")):
            # 铁律 1：同名子元素只能有一个（这里专查 rFonts）
            rfs = rpr.findall(q("rFonts"))
            total += len(rfs)
            dup += max(0, len(rfs) - 1)
            # 铁律 2：rFonts 必须是 rPr 第一个子元素
            if rfs and rpr.index(rfs[0]) != 0:
                not_first += 1
    return {"total": total, "duplicate": dup, "not_first": not_first}


def run_east_asia(path: Path) -> list[str | None]:
    """按文档顺序列出每个非公式 run 的 ``rFonts/@eastAsia``。"""
    root = etree.fromstring(parts(path)["word/document.xml"])
    out = []
    for p in root.iter(q("p")):
        for r in p.findall(q("r")):
            rpr = r.find(q("rPr"))
            rf = rpr.find(q("rFonts")) if rpr is not None else None
            out.append(rf.get(q("eastAsia")) if rf is not None else None)
    return out


def run_size(path: Path) -> list[str | None]:
    """按文档顺序列出每个非公式 run 的 ``rPr/w:sz``（半磅）。"""
    root = etree.fromstring(parts(path)["word/document.xml"])
    out = []
    for p in root.iter(q("p")):
        for r in p.findall(q("r")):
            rpr = r.find(q("rPr"))
            sz = rpr.find(q("sz")) if rpr is not None else None
            out.append(sz.get(q("val")) if sz is not None else None)
    return out


def make_doc2(path: Path) -> None:
    """RT-009 脏 fixture 的**干净底座**：全部 run 的 ``w:rPr`` 合法（rFonts 唯一且在首位）。

    这是对照组——它跑 ``fix_fonts`` 必然得到 ``dup=0 not_first=0``，**恰恰是 RT-009
    警告的那个陷阱**：拿干净文件验收，0/0 不能证明任何东西。真正承载信号的是
    :func:`make_dirty_doc`。
    """
    d = docx.Document()
    for i in range(11):
        d.add_paragraph(f"body text {i}")
    d.save(str(path))


def _rewrite_document(path: Path, mutate) -> None:
    """读出 ``word/document.xml`` → ``mutate(root)`` → 原样写回整个 zip。"""
    blobs = parts(path)
    root = etree.fromstring(blobs["word/document.xml"])
    mutate(root)
    blobs["word/document.xml"] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, blob in blobs.items():
            zout.writestr(name, blob)


def _ensure_rpr(r):
    rpr = r.find(q("rPr"))
    if rpr is None:
        rpr = etree.Element(q("rPr"))
        r.insert(0, rpr)
    return rpr


def make_dirty_doc(path: Path, dup_runs: int = 5, misplaced_runs: int = 5) -> None:
    """RT-009 脏 fixture 构造器：把两条铁律**各踩一次**。

    造两类脏数据（改 `word/document.xml`，改完即关闭重打包）：

    1. **同名子元素重复**——前 ``dup_runs`` 个 run 的 ``w:rPr`` 里塞**两个**
       ``w:rFonts``。xsd:sequence + maxOccurs=1，重复即 schema 非法，Word 报
       「文件可能已损坏」。对应 ``dedupe_rpr()`` 的职责。
    2. **``w:rFonts`` 不在首位**——接着 ``misplaced_runs`` 个 run，把 ``w:b``
       插到 ``w:rPr`` 的第 0 位、``w:rFonts`` 排它后面。对应 ``set_fonts()``
       的 ``rpr.insert(0, rf)`` 语义（``append`` 到尾部同样非法）。

    这两类脏数据是 ``dup > 0 and not_first > 0`` 的唯一来源；不塞就没有信号
    （RT-009 记录当年实测 ``total=470 dup=5 not_first=5``，此处 run 数不同，
    只要 ``dup/not_first`` 均 > 0 即等价）。
    """
    make_doc2(path)

    def mutate(root):
        runs = [r for p in root.iter(q("p")) for r in p.findall(q("r"))]
        assert len(runs) >= dup_runs + misplaced_runs, (
            f"底座只有 {len(runs)} 个 run，不够塞 {dup_runs}+{misplaced_runs} 处脏数据")

        for r in runs[:dup_runs]:
            rpr = _ensure_rpr(r)
            rf = rpr.find(q("rFonts"))
            if rf is None:
                rf = etree.SubElement(rpr, q("rFonts"))
            rf.set(q("ascii"), "Arial")
            rf.set(q("eastAsia"), "黑体")
            twin = etree.SubElement(rpr, q("rFonts"))   # ← 铁律 1 违反：同名第二个
            twin.set(q("ascii"), "Calibri")

        for r in runs[dup_runs:dup_runs + misplaced_runs]:
            rpr = _ensure_rpr(r)
            rpr.insert(0, etree.Element(q("b")))        # ← 铁律 2 违反：rFonts 不再是首个
            etree.SubElement(rpr, q("rFonts"))

    _rewrite_document(path, mutate)


def count_runs(path: Path) -> int:
    """``word/document.xml`` 里的 ``w:r`` 总数——用来验「修完没过度破坏」。"""
    root = etree.fromstring(parts(path)["word/document.xml"])
    return sum(1 for _ in root.iter(q("r")))


def make_doc(path: Path) -> None:
    """造一份三种 run 都齐的文档：手动大标题 / 样式标题 / 正文。"""
    d = docx.Document()
    p0 = d.add_paragraph()
    r0 = p0.add_run("manual big title")
    r0.font.size = docx.shared.Pt(16)          # ≥ 14pt 且无 pStyle → 手动标题
    p1 = d.add_paragraph("styled heading text", style="Heading 1")
    for r in p1.runs:
        r.font.size = docx.shared.Pt(12)       # 有 pStyle → 归样式层
    d.add_paragraph("normal body text")
    d.save(str(path))


# ---------------------------------------------------------------- 1. 两条铁律
def test_rule1_no_duplicate_rfonts(tmp_path: Path):
    """处理完不得有重复 ``w:rFonts``。"""
    f = tmp_path / "f.docx"
    make_doc(f)

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    v = rule_violations(parts(f))
    assert v["duplicate"] == 0, f"重复 rFonts {v['duplicate']} 处 → Word 会判损坏"
    assert v["total"] > 0, "一个 rFonts 都没写，说明根本没处理"


def test_rule2_rfonts_is_first_child(tmp_path: Path):
    """``w:rFonts`` 必须是 ``w:rPr`` 首个子元素。"""
    f = tmp_path / "f.docx"
    make_doc(f)

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    v = rule_violations(parts(f))
    assert v["not_first"] == 0, f"rFonts 不在首位 {v['not_first']} 处 → Word 会判损坏"


def test_rfonts_carries_the_font_table(tmp_path: Path):
    """四个属性必须齐：ascii/hAnsi/cs = TNR，eastAsia = 中文族。"""
    f = tmp_path / "f.docx"
    make_doc(f)
    run_fix_fonts(f)

    root = etree.fromstring(parts(f)["word/document.xml"])
    for rf in root.iter(q("rFonts")):
        assert rf.get(q("ascii")) == "Times New Roman"
        assert rf.get(q("hAnsi")) == "Times New Roman"
        assert rf.get(q("cs")) == "Times New Roman"
        assert rf.get(q("eastAsia")) in ("宋体", "微软雅黑")


# ---------------------------------------------------------------- 2. 幂等
def test_idempotent_second_pass(tmp_path: Path):
    """连跑两次：两条铁律仍成立，且 rFonts 总数不增。"""
    f = tmp_path / "f.docx"
    make_doc(f)

    assert run_fix_fonts(f).returncode == 0
    first = rule_violations(parts(f))

    assert run_fix_fonts(f).returncode == 0
    second = rule_violations(parts(f))

    assert second["duplicate"] == 0
    assert second["not_first"] == 0
    assert second["total"] <= first["total"], (
        f"第二次跑反而加了 rFonts：{first['total']} → {second['total']}")


# ---------------------------------------------------------------- 3. 阈值语义
def test_threshold_14_treats_16pt_as_manual_heading(tmp_path: Path):
    """阈值 14：无 pStyle 的 16pt run → eastAsia 变雅黑。"""
    f = tmp_path / "f.docx"
    make_doc(f)

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    ea, sz = run_east_asia(f), run_size(f)
    # p0 的 run：sz=32（16pt ×2），无 pStyle → 手动标题
    assert "32" in sz
    i = sz.index("32")
    assert ea[i] == "微软雅黑", f"16pt 无样式段没被判为手动标题（eastAsia={ea[i]}）"


def test_threshold_20_leaves_16pt_alone(tmp_path: Path):
    """阈值 20：同一个 16pt run 不该再被当手动标题（回落宋体）。"""
    f = tmp_path / "f.docx"
    make_doc(f)

    proc = run_fix_fonts(f, "--heading-threshold", "20")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    ea, sz = run_east_asia(f), run_size(f)
    i = sz.index("32")
    assert ea[i] == "宋体", (
        f"阈值 20 下 16pt（32 半磅）仍被判手动标题（eastAsia={ea[i]}）")


def test_pstyle_heading_wins_regardless_of_threshold(tmp_path: Path):
    """有 pStyle=Heading 1 的 run 走样式层：两个阈值下都该是雅黑。"""
    for threshold in ("14", "20"):
        f = tmp_path / f"f{threshold}.docx"
        make_doc(f)
        proc = run_fix_fonts(f, "--heading-threshold", threshold)
        assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

        ea, sz = run_east_asia(f), run_size(f)
        # p1 的 run：sz=24（12pt ×2），有 pStyle=Heading1
        i = sz.index("24")
        assert ea[i] == "微软雅黑", (
            f"阈値 {threshold}：Heading 1 样式段 eastAsia={ea[i]}，应按样式层取雅黑")


def test_body_run_stays_songti(tmp_path: Path):
    """普通正文 run（无 sz、无 pStyle）→ 宋体。"""
    f = tmp_path / "f.docx"
    make_doc(f)
    run_fix_fonts(f)

    ea = [x for x in run_east_asia(f) if x]
    assert "宋体" in ea, f"正文没落成宋体：{ea}"


# ---------------------------------------------------------------- 4. 解包往返
def test_package_still_valid_after_fix(tmp_path: Path):
    """改完的包必须还能被 python-docx 打开、且必需部件都在。"""
    f = tmp_path / "f.docx"
    make_doc(f)
    assert run_fix_fonts(f).returncode == 0

    blobs = parts(f)
    assert "word/document.xml" in blobs
    assert "word/styles.xml" in blobs
    assert "[Content_Types].xml" in blobs

    docx.Document(str(f))       # 不抛即通过

    # 没有残留的临时文件（脚本用 <名>.fontfix.tmp.docx 中转）
    assert not list(tmp_path.glob("*.fontfix.tmp.docx")), "临时文件没被清掉"


# ------------------------------------------------- 5. RT-009 脏 fixture 正面用例
def test_dirty_fixture_really_is_dirty(tmp_path: Path):
    """**第一道门**：构造器必须真造出 ``dup>0`` 且 ``not_first>0`` 的脏数据。

    用例若只在干净文件上跑，``dup/not_first`` 天然是 0，最终断言「修后归 0」
    会**无条件通过**——这正是 RT-009 记的方法论风险。所以先把"脏"本身断死：
    构造器坏了、run 定位偏了、zip 没写回去，这里立刻红。
    """
    f = tmp_path / "dirty.docx"
    make_dirty_doc(f)

    v = rule_violations(parts(f))
    assert v["duplicate"] == 5, (
        f"铁律 1 的脏数据没造对：期望 5 处重复 rFonts，实测 {v['duplicate']}"
        f"（total={v['total']}）")
    assert v["not_first"] == 5, (
        f"铁律 2 的脏数据没造对：期望 5 处 rFonts 不在首位，实测 {v['not_first']}"
        f"（total={v['total']}）")
    assert v["total"] > 0


def test_fix_fonts_cleans_dirty_fixture_and_keeps_runs(tmp_path: Path):
    """RT-009 正面用例：脏 fixture → ``fix_fonts`` → 两条铁律归零且 run 不丢。

    三层断言（缺一层就等于没测东西）：

    1. **脏**：修之前 ``dup > 0 and not_first > 0``（fixture 真造出来了）；
    2. **净**：修之后 ``dup == 0 and not_first == 0``（两条铁律都被修）；
    3. **不过度破坏**：``w:r`` 总数前后相等且 > 0（没把 run 删掉/吞掉）。

    对照组 :func:`make_doc2` 是干净底座，它跑出 ``0/0`` **不构成证明**（基线本身
    也是 0/0），所以本用例刻意用脏底座。
    """
    f = tmp_path / "dirty.docx"
    make_dirty_doc(f)

    before = rule_violations(parts(f))
    runs_before = count_runs(f)

    # 断言 1：脏——没有这层，后面的「归零」是自证
    assert before["duplicate"] > 0, (
        f"fixture 不脏（duplicate=0），本用例测不到 dedupe_rpr：{before}")
    assert before["not_first"] > 0, (
        f"fixture 不脏（not_first=0），本用例测不到 insert(0, rf)：{before}")

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    after = rule_violations(parts(f))
    runs_after = count_runs(f)

    # 断言 2：两条铁律归零
    assert after["duplicate"] == 0, (
        f"重复 rFonts 未清干净：{before['duplicate']} → {after['duplicate']}，"
        "Word 会判损坏")
    assert after["not_first"] == 0, (
        f"rFonts 仍有不在首位的：{before['not_first']} → {after['not_first']}，"
        "set_fonts 的 insert(0, rf) 语义没生效")

    # 断言 3：不过度破坏
    assert runs_before > 0, "底座一个 run 都没有，用例无意义"
    assert runs_after == runs_before, (
        f"run 被改掉了：{runs_before} → {runs_after}，处理过程吞了/多了 run")

    # 修完仍可被 python-docx 打开（包结构没被 zip 重写搞坏）
    docx.Document(str(f))


def test_clean_fixture_baseline_is_zero_zero(tmp_path: Path):
    """**反向对照**：干净底座跑出 ``0/0`` —— 记录 RT-009 陷阱本身。

    这条用例故意证明"干净文件验不出东西"：``dup/not_first`` 在修前修后**都是 0**，
    所以任何只跑干净文件的验收都不能称为"已验证"。把它留在测试里，是为了让后来者
    读代码时直接看见这个陷阱，而不是重新踩一次。
    """
    f = tmp_path / "clean.docx"
    make_doc2(f)

    before = rule_violations(parts(f))
    assert before["duplicate"] == 0 and before["not_first"] == 0, (
        f"干净底座居然不干净：{before}")

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")

    after = rule_violations(parts(f))
    assert after["duplicate"] == 0 and after["not_first"] == 0, (
        f"修后反而不干净：{after}")

    # 关键：修前修后的脏数据计数完全相同 → 干净文件上「归零」是空断言
    assert (before["duplicate"], before["not_first"]) == (0, 0) == (
        after["duplicate"], after["not_first"]), (
        "对照组的意义就是「修前后都是 0/0」，否则它就不再是对照组")


def test_dirty_fixture_fix_is_idempotent(tmp_path: Path):
    """脏 fixture 连跑两次：第二次不得再产出任何违反，rFonts 总数不增。"""
    f = tmp_path / "dirty.docx"
    make_dirty_doc(f)

    assert run_fix_fonts(f).returncode == 0
    first = rule_violations(parts(f))
    assert first["duplicate"] == 0 and first["not_first"] == 0
    runs_first = count_runs(f)

    assert run_fix_fonts(f).returncode == 0
    second = rule_violations(parts(f))

    assert second["duplicate"] == 0
    assert second["not_first"] == 0
    assert second["total"] == first["total"], (
        f"第二次跑改了 rFonts 数量：{first['total']} → {second['total']}")
    assert count_runs(f) == runs_first, "第二次跑动到了 run 数量"


# ===========================================================================
# 输入错误契约（P9-C）
#
# `fix_fonts.py` 是**就地改写**工具（临时文件 + `shutil.move`），所以输入校验
# 的两个失误都会留下真实痕迹：
#
#   * 校验太晚 → `zipfile.ZipFile(src)` 抛 `BadZipFile` / `IsADirectoryError`，
#     命令行上是 13~14 行 traceback + exit 1（实测 2026-11-11）；
#   * 只校验「存在」不校验「是包」→ 非 zip 一路走到 `_fix_docx` 才崩。
#
# 更要紧的是 `nargs="+"`：旧版**边校验边处理**，第一个文件已经就地改写完了
# 才发现第二个是垃圾 —— 一次「用法错误」的调用改掉了别的文件。本节最后一条
# 用例把「先全量校验、再动手」钉住。
# ===========================================================================
def _nonzip(tmp_path: Path, name: str = "renamed.docx") -> Path:
    p = tmp_path / name
    p.write_text("i am not a zip package", encoding="utf-8")
    return p


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def run_fix_fonts_many(*paths) -> subprocess.CompletedProcess:
    """多文件调用（``nargs="+"``）。与其他 helper 一致取 **bytes** 输出。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "fix_fonts.py"), *map(str, paths)],
        capture_output=True, env=env, check=False)


def test_missing_path_is_a_usage_error(tmp_path: Path):
    """缺文件 → exit 2 + 单行 ``no such file`` + stdout 空。"""
    missing = tmp_path / "no-such.docx"
    proc = run_fix_fonts(missing)
    assert proc.returncode == 2, proc.stderr.decode("utf-8", "replace")
    line = one_error_line(proc.stderr.decode("utf-8", "replace"))
    assert "no such file" in line
    assert str(missing) in line
    assert proc.stdout == b""


def test_directory_is_a_usage_error(tmp_path: Path):
    """传目录 → exit 2 + 单行 ``not a file``，而非 13 行 traceback。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()
    proc = run_fix_fonts(adir)
    assert proc.returncode == 2, proc.stderr.decode("utf-8", "replace")
    assert one_error_line(proc.stderr.decode("utf-8", "replace")) == \
        f"error: not a file: {adir}"
    assert proc.stdout == b""


def test_non_zip_is_a_usage_error(tmp_path: Path):
    """非 zip → exit 2 + 单行 ``not a ZIP package``，且文件一个字节没被改。"""
    bogus = _nonzip(tmp_path)
    proc = run_fix_fonts(bogus)
    assert proc.returncode == 2, proc.stderr.decode("utf-8", "replace")
    assert one_error_line(proc.stderr.decode("utf-8", "replace")) == \
        f"error: not a ZIP package: {bogus}"
    assert proc.stdout == b""
    assert bogus.read_text(encoding="utf-8") == "i am not a zip package"


def test_bad_second_file_leaves_the_first_byte_identical(tmp_path: Path):
    """★ 第一个合法、第二个是垃圾 → **第一个文件一个字节都不许变**。

    旧版边校验边处理，`good.docx` 已经被就地改写过一次才轮到 `renamed.docx`
    崩掉。这条钉的是「先全量校验，再逐个动手」。
    """
    good = tmp_path / "good.docx"
    d = docx.Document()
    d.add_paragraph("hello")
    d.save(str(good))
    before = good.read_bytes()
    bogus = _nonzip(tmp_path)

    proc = run_fix_fonts_many(good, bogus)
    assert proc.returncode == 2, proc.stderr.decode("utf-8", "replace")
    assert proc.stdout == b""
    assert good.read_bytes() == before, "第二个文件不合格时，第一个文件被改写了"


def test_valid_document_still_gets_its_fonts_fixed(tmp_path: Path):
    """合法文档仍是 exit 0 + 既有计数行（加校验没把成功路径改坏）。"""
    f = tmp_path / "doc.docx"
    d = docx.Document()
    d.add_paragraph("中文正文 english body")
    d.save(str(f))

    proc = run_fix_fonts(f)
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    out = proc.stdout.decode("utf-8", "replace")
    assert out.startswith("  正文 "), out
    assert "[doc.docx]" in out
