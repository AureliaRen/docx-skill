# -*- coding: utf-8 -*-
r"""``insert_formula.py`` 回归（RT-002）+ 相量点计数/错误通道回归（RT-001 修法 6）。

两组共 6 条：

**RT-002 —— `insert_formula.py` 此前全仓零测试覆盖**

1. ``--after "marker"`` 命中 → 公式段落在 marker 段落的**正下一个位置**；
2. marker 找不到 → ``exit 1``，且 ``word/document.xml`` **一个字节都没变**；
3. ``--after`` 与 ``--before`` 同时给 → 非零退出，且文档未被修改。

**RT-001 —— 相量点修复的计数与失败通道（B1 刚落的实现）**

4. 插 ``\dot{...}`` 后产物里 ``m:limUpp`` 为 0、存在 ``m:acc``（修复真发生）；
5. ``stats["phasor_fixes"]`` 在三样例上分别是 1 / 1 / 2；
6. ``fix_phasor_dots`` 抛异常 → 进 ``phasor_errors``、**不进** ``failed``、公式仍然可用。

## ⚠️ 全部 6 条都挂 ``@needs_omml`` —— 第 1/2/3 条也必须挂（别再误判）

原因不只是第 4/5/6 条要断言 OMML 产物。**``insert_formula.main()`` 的
``build_para()``（L125–L129）在所有 marker 校验之前执行**：

```
L125-129  try: para = build_para(...)              ← 这里就要做 XSLT 转换，需要 XSL
          except (RuntimeError, FileNotFoundError): return 2   ← 无 XSL 直接从这里返回
L152      if args.after and args.before: ... return 1          ← 互斥检查在后面
L158-164  insert_near(...) / marker not found: return 1        ← marker 检查也在后面
```

无 XSL 时 ``build_para`` 抛 ``FileNotFoundError`` → L129 ``return 2`` → **后面两个
防护根本执行不到**。于是第 1/2/3 条测的行为在无 Office 的机器上**不可达**，
实测表现为 ``AssertionError: … rc=2 … MML2OMML_XSL … 不是文件``。

所以这三条**不是「不依赖 OMML」，而是「依赖 OMML 才能走到被测逻辑」**——在测试
意义上与依赖 OMML 等价，必须 skip 而不是 fail。首轮判成「1/2/3 不依赖 OMML、
云端照跑」是错的：那是把「**不检查** OMML 产物」当成了「**不依赖** OMML 环境」，
漏掉了「被测逻辑的执行前置条件」这一层。云端 CI 上这个误判会直接红。

## 为什么用子进程 + XML 级断言

`insert_formula` 是 CLI，要被锁住的是**退出码**与**写进 document.xml 的结构**，不是
它的 stdout 措辞。stdout 只做辅助断言（有就断，用来兜住「跑到 write 分支了」）。
位置断言读 `<w:body>` 的**直接子元素序列**，因为「文件里有 oMath」这种断言对
「追加到了文末」的回归完全免疫 —— 那正是 RT-002 要防的失效形态。

## 两条实测出来的坑（写测试时踩过，别重踩）

**坑 1：测相量点必须用会走 OMML 的片段。**
`docx_mathify` 对 ``$...$`` 片段有 ``MIXED_SUPERSCRIPT`` 分流，纯上下标会降级成
文字角标（``w:rPr/w:vertAlign``），产出 **0 个** ``<m:oMath>``，此时测不到
``m:acc``。实测 ``$\\dot{U}_{AB}$`` 与 ``$\\dot{I}_{L}+\\dot{I}_{C}$`` 走 OMML；
而 ``$\\dot{I}$`` 单点式**经 mathify 的 span 检测根本没被认成公式**
（``oMath=0``、``phasor_fixes=0``）——所以第 5 条不走 mathify，直接测
``_omml_with_phasor()`` 这个计数发生地（它收的是已确认要转 LaTeX 的片段）。

**坑 2：``--dry-run`` 下 ``phasor_fixes`` 恒为 0。**
`mathify(dry_run=True)` 不跑 `rebuild()`，与 ``plain_no_layout`` 同源。
**本文件不使用任何 dry-run 场景**去测相量点计数，否则会得到「恒 0 的假失败」。
第 6 条虽然有 ``out=None``，但那是**不产出文件**、不是 dry-run：`rebuild()` 照跑，
计数照累。

Run:  python -m pytest tests/test_insert_formula.py -q
"""
from __future__ import annotations

import hashlib
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

import docx_mathify as M  # noqa: E402
from docx_pack import unpack  # noqa: E402

# OMML 可用性判定收在 `tests/_omml.py` 一处（全仓唯一实现，见那个文件的 docstring）。
from _omml import needs_omml  # noqa: E402

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MNS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# python-docx 是硬依赖（造 fixture 用），缺了直接失败而非 skip
docx = pytest.importorskip("docx", reason="需要 python-docx 造 fixture")


def w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def m(tag: str) -> str:
    return f"{{{MNS}}}{tag}"


# ---------------------------------------------------------------- helpers
def run_insert(unpacked: Path, *args) -> subprocess.CompletedProcess:
    """以子进程跑 ``insert_formula.py``（CLI 契约要按真进程验）。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "insert_formula.py"), str(unpacked),
         *map(str, args)],
        capture_output=True, env=env, check=False,
    )


def make_unpacked(tmp_path: Path, paragraphs: list[str]) -> Path:
    """造一个 docx（每项一段）→ 解包成目录，返回解包目录。"""
    d = docx.Document()
    for text in paragraphs:
        d.add_paragraph(text)
    src = tmp_path / "in.docx"
    d.save(str(src))

    up = tmp_path / "unpacked"
    unpack(src, up)
    assert (up / "word" / "document.xml").is_file(), "解包没产出 word/document.xml"
    return up


def doc_xml(unpacked: Path) -> Path:
    return unpacked / "word" / "document.xml"


def doc_sha(unpacked: Path) -> str:
    return hashlib.sha256(doc_xml(unpacked).read_bytes()).hexdigest()


def xml_root(unpacked: Path):
    return etree.fromstring(doc_xml(unpacked).read_bytes())


def body_paragraphs(unpacked: Path) -> list:
    """``<w:body>`` 的**直接** ``<w:p>`` 子元素，按文档顺序。

    刻意不用 ``iter()``：``iter`` 会把表格单元格里的段落也捞进来，而本文件断言的是
    「公式段落落在 body 的第几个位置」—— 只认直接子元素才是那个位置。
    """
    body = xml_root(unpacked).find(w("body"))
    assert body is not None, "<w:body> 不见了"
    return [c for c in body if c.tag == w("p")]


def para_text(p) -> str:
    return "".join(t.text or "" for t in p.iter(w("t")))


def count(el, ns_tag: str) -> int:
    return len(list(el.iter(ns_tag)))


def omath_in(el) -> int:
    return count(el, m("oMath"))


def index_of_paragraph_containing(paras: list, needle: str) -> int:
    for i, p in enumerate(paras):
        if needle in para_text(p):
            return i
    raise AssertionError(
        f"fixture 里找不到含 {needle!r} 的段落；实际段落 = "
        f"{[para_text(p) for p in paras]}")


# ===========================================================================
# RT-002 · 1) --after 命中 → 公式插在 marker 段落之后
# ===========================================================================
@needs_omml      # ← 见模块 docstring：build_para() 在所有 marker 校验之前跑
def test_after_marker_inserts_paragraph_right_after_marker(tmp_path: Path):
    r"""``--after "marker here"``：公式段落必须紧贴 marker 段落的下一位。

    三段 fixture：``alpha one`` / ``marker here`` / ``omega three``。

    断言全部落在 **body 直接子元素序列**上：

    * 恰好多出一段，且该段含 ``m:oMath``（公式真进了 body，不是进了别处）；
    * 该段下标 == marker 段下标 + 1（**紧跟**，不是"文档里某处有"）；
    * marker 段自身不含 oMath（没有把公式塞进 marker 段内部）;
    * 原来的第 3 段仍在公式段之后（顺序没被打乱）。

    「文件里有 oMath 就算过」这类断言对「静默追加到文末」的回归**完全免疫**：
    ``insert_near`` 的 fallback 正是插在 ``w:sectPr`` 之前（= 文末），照样能数到 1 个
    oMath。所以位置必须按下标断。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    before = body_paragraphs(up)
    assert len(before) == 3, f"fixture 段落数不对：{[para_text(p) for p in before]}"

    proc = run_insert(up, r"\frac{a}{b}", "--after", "marker here")
    assert proc.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (proc.returncode,
                                     proc.stderr.decode("utf-8", "replace")))
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "Inserted inline formula" in stdout, f"没走到写入分支：stdout={stdout!r}"

    after = body_paragraphs(up)
    texts = [para_text(p) for p in after]
    assert len(after) == 4, f"应多出 1 段公式段，实际段落 = {texts}"

    marker_i = index_of_paragraph_containing(after, "marker here")
    formula_i = marker_i + 1

    assert omath_in(after[formula_i]) >= 1, (
        f"marker 段的下一段不是公式段：段落序 = {texts}")
    assert omath_in(after[marker_i]) == 0, (
        f"公式被塞进了 marker 段内部：段落序 = {texts}")

    total_omath = omath_in(xml_root(up))
    assert total_omath == 1, f"oMath 数量不是 1（说明插了多处或没插）：{total_omath}"

    # 顺序没乱：首段仍是 alpha，末段仍是 omega（公式段夹在中间）
    assert texts[0] == "alpha one", f"文档开头被改动：{texts}"
    assert texts[-1] == "omega three", (
        f"公式落到了文末（fallback 到 sectPr 之前）而不是 marker 之后：{texts}")


# ===========================================================================
# RT-002 · 2) marker 找不到 → exit 1，且文档一个字节都没变
# ===========================================================================
@needs_omml      # ← 见模块 docstring：build_para() 在所有 marker 校验之前跑
def test_marker_not_found_exits_1_and_leaves_document_untouched(tmp_path: Path):
    """给一个文档里不存在的 marker → 必须 ``exit 1``，且**不写文件**。

    只断退出码是不够的：把「marker 未命中」悄悄降级成「追加到文末并 exit 0」，
    退出码断言会红——但如果实现改成「报错并 exit 1，同时仍写了一次文件」，
    退出码断言依然绿。所以还要断 ``document.xml`` 的 SHA256 与段落数**都没变**。

    这条在源码里的锚点是：``write(tree, doc_xml)``（L166）位于
    ``if marker and not anchored`` 的 ``return 1`` **之后** —— 失败路径根本不落盘。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))

    proc = run_insert(up, "E=mc^2", "--after", "no-such-marker")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode == 1, (
        "marker 未命中必须 exit 1（静默追加到文末是最坏结果）："
        "rc=%s stdout=%r stderr=%r"
        % (proc.returncode, proc.stdout.decode("utf-8", "replace"), stderr))
    assert "marker" in stderr.lower(), f"报错没说明是 marker 的问题：{stderr!r}"

    assert doc_sha(up) == sha_before, (
        "失败路径仍然改写了 document.xml（半成品落盘），"
        f"改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, "失败路径仍增删了段落"
    assert omath_in(xml_root(up)) == 0, "失败路径仍把公式写进了文档"


# ===========================================================================
# RT-002 · 3) --after 与 --before 同时给 → 报错退出，文档不变
# ===========================================================================
@needs_omml      # ← 见模块 docstring：build_para() 在所有 marker 校验之前跑
def test_after_and_before_together_is_rejected(tmp_path: Path):
    """两个互斥锚点同时给 → 非零退出，且文档未被修改。

    退出码断 ``!= 0`` 而不是死钉某个具体值：这里要锁的语义是"**互相矛盾的
    锚点不许被当成成功**"，至于它是 1（实测值）还是 2（argparse 风格的用法错误）
    属于实现口味，钉死具体数字只会给未来无害的调整制造假报警。
    实测当前实现返回 **1**（``insert_formula.py`` L152–L154）。

    同时断文档未变：互斥检查在 ``insert_near`` **之前**，但把它挪到 ``write()``
    之后就会变成"报错但已经写坏"，这条断言正是防那个。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))

    proc = run_insert(up, "E=mc^2", "--after", "marker here",
                      "--before", "marker here")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode != 0, (
        "互斥锚点被当成功了：rc=0，公式可能被塞到了某个来路不明的位置；"
        "stdout=%r" % proc.stdout.decode("utf-8", "replace"))
    assert "mutually exclusive" in stderr.lower() or "after" in stderr.lower(), (
        f"报错没指向互斥参数：{stderr!r}")

    assert doc_sha(up) == sha_before, (
        f"报错的同时改了文档：改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, "报错的同时增删了段落"
    assert omath_in(xml_root(up)) == 0, "报错的同时把公式写进了文档"


# ===========================================================================
# RT-001 · 4) 修复发生：limUpp 清零 + 出现 m:acc
# ===========================================================================
@needs_omml
def test_insert_formula_writes_accent_instead_of_limupp(tmp_path: Path):
    r"""插 ``\dot{I}`` → 产物 ``m:limUpp`` 为 0、存在 ``m:acc``。

    这是 RT-001「功能本身没坏，只是不说」的**正向证据**：`latex2mathml` 会把
    ``\dot{I}`` 转成 ``m:limUpp``（Word 里是"大字 + 飘着的点"），
    ``fix_phasor_dots`` 必须把它换成 ``m:acc``。断言直接数 XML 元素——数字符串
    ``"<m:acc"`` 会把 ``<m:accPr>`` 一起数进去，得到 2 这种迷惑数字。

    为什么样例必须是 ``\dot{I}`` 这种真有 ``\dot`` 的片段：没有 ``\dot`` 就没有
    ``m:limUpp``，断言会退化成"0 == 0 且 acc 在别处碰巧大于 0"，测空气。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    proc = run_insert(up, r"\dot{I}", "--after", "marker here")
    assert proc.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (proc.returncode,
                                     proc.stderr.decode("utf-8", "replace")))

    root = xml_root(up)
    limupp = count(root, m("limUpp"))
    acc = count(root, m("acc"))

    assert omath_in(root) == 1, "公式没进文档（oMath 数为 0 或 >1）"
    assert limupp == 0, (
        f"{limupp} 处 m:limUpp 残留 → Word 里相量点会是浮着的大点，修复没生效")
    assert acc >= 1, (
        "没有 m:acc → 相量点没被改成组合点；"
        f"当前 acc={acc} limUpp={limupp}")

    # 辅助断言：CLI 的计数后缀（RT-001 修法 5）。主证据是上面的 XML 计数。
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "修正相量点 1 处" in stdout, (
        f"stdout 没报修了几处（RT-001 修法 5 的回归）：{stdout!r}")


# ===========================================================================
# RT-001 · 5) 计数正确：1 / 1 / 2
# ===========================================================================
@needs_omml
def test_phasor_fix_count_matches_samples():
    r"""``stats["phasor_fixes"]`` 在三样例上必须是 **1 / 1 / 2**。

    直接测计数发生地 ``docx_mathify._omml_with_phasor()``：它收的是**已经确定要转
    OMML 的 LaTeX 片段**，正是 RT-001 里被点名的那个包装器。绕开 mathify 的 span
    检测是刻意的——``$\\dot{I}$`` 这种单点式经 mathify 的 ``find_math_spans``
    根本不会被认成公式（实测 ``oMath=0``、``fixes=0``），拿它当样例只会测到
    「公式压根没进文档」，与相量点计数无关（见模块 docstring 的坑 1）。

    同时断 ``limUpp`` 归零，把"计数涨了但没真修"这种回归也挡在门外：只断计数的话，
    把 ``fix_phasor_dots`` 换成 `return 2` 也能过。

    三个样例放在**同一条用例**里循环（而不是 ``parametrize``）：一条用例对应返工单
    里「计数正确」这一个语义，且三个样例一起跑才看得出「1 / 1 / 2 的相对关系」
    （比如恒返回 2 的实现会被第一个样例抓住）。为了让失败信息一次说完三处，
    断言前先聚齐全部 mismatch 再一次报出，不做「第一个错就停」。
    """
    samples = [
        (r"\dot{I}", 1),
        (r"\dot{U}_{AB}", 1),
        (r"\dot{I}_{L}+\dot{I}_{C}", 2),
    ]

    observed = []
    problems = []

    for tex, expected in samples:
        stats = {"phasor_fixes": 0, "phasor_errors": []}
        om = M._omml_with_phasor(tex, stats)

        got = stats["phasor_fixes"]
        limupp = count(om, m("limUpp"))
        acc = count(om, m("acc"))
        observed.append((tex, expected, got, limupp, acc))

        if om is None:
            problems.append(f"{tex!r}：转换返回 None")
            continue
        if got != expected:
            problems.append(f"{tex!r}：计数错，期望 {expected}，实际 {got}")
        if stats["phasor_errors"]:
            problems.append(
                f"{tex!r}：正常路径不该有相量点错误，实际 {stats['phasor_errors']}")
        if limupp != 0:
            problems.append(f"{tex!r}：仍有 {limupp} 处 m:limUpp 残留，修复没真发生")
        if acc != expected:
            problems.append(
                f"{tex!r}：m:acc 数 {acc} 与修复处数 {expected} 不符")

    assert not problems, (
        "相量点计数/修复不符（样例 = latex, 期望计数, 实际计数, limUpp, acc）：\n"
        + "\n".join(f"  {row}" for row in observed)
        + "\n问题：\n" + "\n".join(f"  - {p}" for p in problems))


# ===========================================================================
# RT-001 · 6) 异常进 phasor_errors，而非 failed
# ===========================================================================
@needs_omml
def test_phasor_exception_lands_in_phasor_errors_not_failed(
        tmp_path: Path, monkeypatch):
    r"""``fix_phasor_dots`` 抛异常 → ① ``phasor_errors`` 1 条且带异常类型名
    ② ``failed`` 为空 ③ 仍返回可用的 OMML。

    这是 RT-001 的**语义分野**：``failed`` 的语义是"公式没转成"，而相量点没修好时
    公式本身是好的。混在一起会让用户按错误方向排查，也会让一处 ``\dot`` 出问题就
    把整段公式回退成纯文本。

    三层断言缺一不可：
    * 只断 ``phasor_errors`` 有 1 条 → 挡不住"同时也记进 failed"；
    * 只断 ``failed == []`` → 挡不住"异常被整个吞掉、两条通道都没记"；
    * 不断"仍返回可用 OMML" → 挡不住"记完就 return None / 回退纯文本"。

    样例用 ``\\dot{U}_{AB}``：实测它能走 OMML（``\\dot{I}`` 经 mathify 不产生
    oMath，见模块 docstring 坑 1）。
    """
    def boom(_el):
        raise RuntimeError("synthetic phasor failure")

    monkeypatch.setattr(M, "fix_phasor_dots", boom)

    # ---- A) 直调包装器：返回值必须仍是 oMath 元素
    unit_stats = {"phasor_fixes": 0, "phasor_errors": [], "failed": []}
    om = M._omml_with_phasor(r"\dot{U}_{AB}", unit_stats)

    assert om is not None, "相量点修复失败时把整个公式丢掉了（返回 None）"
    assert om.tag == m("oMath"), f"返回的不是 oMath 元素：{om.tag}"

    errs = unit_stats["phasor_errors"]
    assert len(errs) == 1, f"phasor_errors 应有 1 条，实际 {len(errs)}：{errs}"
    assert "RuntimeError" in errs[0]["error"], (
        f"错误串没带异常类型名，无法区分根因：{errs[0]['error']!r}")
    assert errs[0]["latex"] == r"\dot{U}_{AB}", (
        f"错误串没带公式片段，无法定位：{errs[0]['latex']!r}")
    assert unit_stats["failed"] == [], (
        f"相量点失败被误记成转换失败：{unit_stats['failed']}")
    assert unit_stats["phasor_fixes"] == 0, "失败时计数被虚增了"

    # ---- B) 端到端：mathify 的 stats 两条通道照样分得开
    d = docx.Document()
    d.add_paragraph(r"相量 $\dot{U}_{AB}$ 结束")
    src = tmp_path / "in.docx"
    d.save(str(src))
    out = tmp_path / "out.docx"

    stats = M.mathify(src, out, None, None)

    # 直接下标取键 —— 键不在初始化字典里就 KeyError（锁住 RT-001 修法 1）
    assert stats["phasor_fixes"] == 0, (
        f"失败路径不该有修复计数：{stats['phasor_fixes']}")
    assert len(stats["phasor_errors"]) == 1, (
        f"端到端 phasor_errors 应有 1 条，实际：{stats['phasor_errors']}")
    assert stats["failed"] == [], (
        f"端到端把相量点失败记成了转换失败：{stats['failed']}")

    # 公式仍在文档里（没回退成纯文本），且是未修版本（limUpp 还在）
    with zipfile.ZipFile(out) as zf:
        raw = zf.read("word/document.xml")
    root = etree.fromstring(raw)
    assert omath_in(root) >= 1, "相量点失败导致公式没进文档（被回退成纯文本了）"
    assert count(root, m("limUpp")) >= 1, (
        "既然修失败，limUpp 应原样保留；它为 0 说明 monkeypatch 没生效，"
        "本用例实际没测到异常路径")
