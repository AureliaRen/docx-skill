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

## P5-A1 补测（7～12 条）：三个开关 + `--before` 正向 + 表格单元格 +

全部挂 `@needs_omml`，理由同上（`build_para()` 先于一切 marker 校验）。六条各锁一个
**此前零覆盖**的开关或路径，断言口径全部来自实测（下面每条 docstring 里写了实测
数字），不臆造期望值：

7. ``--before`` 正向：公式段落落在 marker 段的**前一位**（此前只有「与 --after 互斥」
   的反向用例，没有正向位置断言）；
8. ``--display``：产物是 ``m:oMathPara``（块级）而非裸 ``m:oMath``，且**不带**行内
   兜底的 U+200B run；
9. ``--dump``：``document.xml`` 的 SHA256 与段落数**都不变**（真只读），stdout 是
   可解析的 ``<w:p>`` 且含恰好 1 个 ``m:oMath``；
10. ``--raw-spacing``：与默认路径的 OMML **结构差异**（默认把 ``\\quad`` 改写成 NBSP
    run，raw 路径该 run 被 XSLT 丢掉）——默认 ``m:r`` 数比 raw 多 1；
11. 表格单元格内 marker：``--after`` 命中 ``w:tc`` 里的文字时**真插进那个单元格**
    （tc 内多出一段带 oMath 的 ``w:p``），body 直接子元素序列不变；
12. ``--after`` + ``--dump`` 同时给：**明确报错退 2**（用法错误），文档零改动。
    ⚠️ **本条在 P6-FIX-C2 被反转**：它此前锁的是「``--dump`` 先 return 0、静默把
    ``--after`` 吞掉」，而那正是 P5-AUDIT-QUALITY §5.7 的缺陷本体。锁住缺陷的
    用例不是回归保护，是缺陷的保鲜膜 —— 报名见 `docs\reports\P6-FIX-C2.md`。

## P6-FIX-C2 补测（13～17 条）：静默说谎类缺陷

13. 空 marker（``--after ""`` / ``--before " "``）→ exit 2（旧版 rc=0 且**静默追加到
    文末**），文档零改动；
14. marker 找不到的报错里给出**实际可搜范围**（``0-based [0, N-1]`` + 单元格数），
    且 N 必须取**插入前**快照（``insert_near`` 的 fallback 会先追加再返回 False，
    事后数会多 1 —— 写这条用例时实测抓到的）；
15. 落点回显 ``at body[i]`` 的下标**就是**该段落在 ``<w:body>`` 直接子元素序列里的
    实测下标（旧版完全没有回显 → 红），并与「marker 下标 + 1」交叉核对；
16. 边界：``--before`` 首段 → ``at body[0]``；``--after`` 末段（下标 N-1）→
    ``at body[N]``；
17. 容器：表格单元格内 marker → ``at tc[2]``（**不是** body 下标；下标口径含
    ``w:tcPr``）。

Run:  python -m pytest tests/test_insert_formula.py -q
"""
from __future__ import annotations

import hashlib
import io
import os
import re
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


def freeze_docx(path: Path) -> Path:
    r"""把 **.docx 包**里 zip 条目的时间戳钉成 ``1980-01-01``，就地重写该包。

    为什么要钉（P5-A2 发现、P5-A1 §八 移植）：

    ``python-docx`` 保存走 ``zipfile`` 默认的 DOS 时间戳，**粒度 2 秒**。同一 2 秒
    刻度内构建的两个内容不同的包，若只比**包**的 SHA256，会因为时间戳相同而**碰巧
    一致**——「包 SHA 不变 → 内容没变」这条断言就会随机放行真实回归。

    ## 本文件为什么要它（以及它在本文件里**不是**主角）

    本文件所有「改前/改后」断言比的是**解包后**的 ``word/document.xml``
    （``doc_sha()``），不是 .docx 包。实测（P5-A1 §八 实验 A/B）：

    * 不同时刻构建的同内容包 → 包 SHA **不同**（时间戳在变）、
      ``word/document.xml`` SHA **完全相同**（该文件里没有任何时间戳字段，
      实测不含 ``dcterms``/``created``/``modified``/年份串）；
    * 同一刻度内构建 → 包 SHA 相同、``document.xml`` SHA 也相同。

    所以 **DOS 时间戳够不到 ``document.xml`` 这一层**，本文件的 ``doc_sha()``
    本来就是确定性内容指纹。移植 ``freeze_docx()`` 的价值是**消除包层的时间漂移**，
    让 fixture 可复现、dump/比对不受时刻影响；它不是 ``doc_sha()`` 判别力的来源。

    ## 真正让 ``doc_sha()`` 有判别力的东西

    ``docx_utils.write()`` 会重排 XML（``tree.write`` + ``fix_root_any`` 改根标签），
    即使语义没变也会改字节——实测「原样 ``a29b80155ad6`` → 无改动 ``parse``+``write``
    一次 ``a3009f94ce6fc65f``」。所以「保留 ``return 1`` 但先 ``write()``」这类回归
    必被抓住。**这是内容变更检测，与 2 秒刻度无关**；P5-A1 §八 的实验 C 连跑 3 次
    全 RED 就是这个证据（三次耗时 2.40/2.24/2.48s，横跨多个刻度窗口，结果相同）。
    """
    src = Path(path)
    raw = src.read_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(raw)) as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(out, "w") as zout:
            for info in infos:
                new = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
                new.compress_type = info.compress_type
                new.external_attr = info.external_attr
                zout.writestr(new, zin.read(info.filename))
    src.write_bytes(out.getvalue())
    return src


def make_unpacked(tmp_path: Path, paragraphs: list[str]) -> Path:
    """造一个 docx（每项一段）→ 解包成目录，返回解包目录。

    打包后先 ``freeze_docx()`` 再解包：让 fixture 的 zip 时间戳固定，包层字节可复现。
    解包出来的 ``word/document.xml`` 本身不含任何时间戳（见 ``freeze_docx`` 的
    docstring），所以这条 freeze 不影响 ``doc_sha()`` 的判别力。
    """
    d = docx.Document()
    for text in paragraphs:
        d.add_paragraph(text)
    src = tmp_path / "in.docx"
    d.save(str(src))
    freeze_docx(src)

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


# ------------------------------------------------- P6-FIX-C2 新增的小工具
def body_children(unpacked: Path) -> list:
    """``<w:body>`` 的**全部**直接子元素（含 ``w:tbl`` / ``w:sectPr``）。

    ``body_paragraphs()`` 只留 ``w:p``，下标记的是"第几段"；而 CLI 回显的
    ``at body[i]`` 报的是**容器直接子元素的下标**（含 ``w:sectPr``）。两者在
    ``--after 末段`` 这种边界上会差 1（公式段落在 ``sectPr`` 之前，段序里是末段、
    子元素序列里下标还得数上 ``sectPr`` 前面的东西）。要核真值就必须用子元素序列，
    不能拿段序凑。
    """
    root = xml_root(unpacked)
    body = root.find(w("body"))
    assert body is not None, "<w:body> 不见了"
    return list(body)


ECHO_RE = re.compile(r"at (body|tc)\[(\d+)\]")


def echo_of(stdout: str) -> tuple[str, int]:
    """从成功 stdout 里取落点回显 ``at body[i]`` / ``at tc[i]``。"""
    found = ECHO_RE.search(stdout)
    assert found, (
        "成功输出里没有落点回显（`at body[i]` / `at tc[i]`）；"
        f"stdout={stdout!r}")
    return found.group(1), int(found.group(2))


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
    freeze_docx(src)          # zip 时间戳钉死（见 freeze_docx docstring）
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


# ===========================================================================
# P5-A1 · 7) --before 正向：公式段落在 marker 段的**前一位**
# ===========================================================================
@needs_omml
def test_before_marker_inserts_paragraph_right_before_marker(tmp_path: Path):
    r"""``--before "marker here"``：公式段落必须紧贴 marker 段的**前一位**。

    此前 ``--before`` 只出现在「与 ``--after`` 互斥」的反向用例里（第 3 条），
    **正向位置一次都没验过**——于是「``--before`` 被实现成 ``--after``」这种
    最典型的回归，现有 6 条全绿也发现不了。

    实测（fixture 三段 ``alpha one`` / ``marker here`` / ``omega three``）：
    ``insert_near(position="before")`` 取 ``at = i``，产物段落序为

    ```
    ['alpha one', '\u200b', 'marker here', 'omega three']
    ```

    ``\u200b`` 是插进去的公式段（行内路径必带 U+200B 兜底 run，见源码规则 2）。
    所以下标关系是 ``公式段 == marker 段 - 1``，而不是 ``+ 1``——这正是本条要钉住
    的方向；改成 ``--after`` 会让公式跑到 marker 后面，``formula_i == marker_i - 1``
    立刻不成立。

    首末段必须不变：``--before`` 走 fallback 时插在 ``w:sectPr`` 之前（= 文末），
    末段断言正是拦这个。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    before = body_paragraphs(up)
    assert len(before) == 3, f"fixture 段落数不对：{[para_text(p) for p in before]}"

    proc = run_insert(up, r"\frac{a}{b}", "--before", "marker here")
    assert proc.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (proc.returncode,
                                     proc.stderr.decode("utf-8", "replace")))
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "Inserted inline formula" in stdout, f"没走到写入分支：stdout={stdout!r}"

    after = body_paragraphs(up)
    texts = [para_text(p) for p in after]
    assert len(after) == 4, f"应多出 1 段公式段，实际段落 = {texts}"

    marker_i = index_of_paragraph_containing(after, "marker here")
    formula_i = marker_i - 1
    assert formula_i >= 0, f"marker 落在第 0 段，--before 不可能有前一位：{texts}"

    assert omath_in(after[formula_i]) == 1, (
        f"marker 段的**前一段**不是公式段（--before 方向反了？）："
        f"段落序 = {texts}，各段 oMath 数 = {[omath_in(p) for p in after]}")
    assert omath_in(after[marker_i]) == 0, (
        f"公式被塞进了 marker 段内部：段落序 = {texts}")

    total = omath_in(xml_root(up))
    assert total == 1, f"oMath 数量不是 1（插了多处或没插）：{total}"

    assert texts[0] == "alpha one", f"文档开头被改动：{texts}"
    assert texts[-1] == "omega three", (
        f"公式落到了文末（fallback 到 sectPr 之前）而不是 marker 之前：{texts}")


# ===========================================================================
# P5-A1 · 8) --display：块级 m:oMathPara，不带行内 U+200B 兜底
# ===========================================================================
@needs_omml
def test_display_switch_emits_block_omathpara_without_zwsp(tmp_path: Path):
    r"""``--display``：公式段落里是 ``m:oMathPara``（块级），**没有** U+200B run。

    源码两条分支（``build_para`` L95–L101）本身就分叉：

    * ``display=True``  → ``w:pPr``（``jc=center``）+ ``m:oMathPara``，**不补** U+200B；
    * ``display=False`` → 裸 ``m:oMath`` + ``mkrun("\u200b")``（防止独占整段时被 Word
      提升成行间公式）。

    实测 ``--display`` 产出的 ``<w:p>`` 直接子元素 == ``['pPr', 'oMathPara']``，
    全文档 ``oMathPara == 1``、``oMath == 1``、段落文字 == ``''``（无 U+200B）。

    三条断言各拦一类回归：
    * 只断 ``oMathPara == 1`` → 挡不住「oMathPara 存在但外面还套了个 oMath」；
    * 只断 ``oMath == 1`` → 挡不住「开关被忽略，退回行内路径」；
    * 不断「无 U+200B」→ 挡不住「display 分支顺手也补了兜底 run」，那会让行间公式
      里多一个零宽字符（Word 里看不见，但把段落判成「有正文内容」）。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    proc = run_insert(up, r"\frac{a}{b}", "--display", "--after", "marker here")
    assert proc.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (proc.returncode,
                                     proc.stderr.decode("utf-8", "replace")))
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "Inserted display formula" in stdout, (
        f"stdout 没标 display 分支（提示词措辞回归）：{stdout!r}")

    root = xml_root(up)
    paras = body_paragraphs(up)
    texts = [para_text(p) for p in paras]
    assert len(paras) == 4, f"应多出 1 段公式段，实际段落 = {texts}"

    marker_i = index_of_paragraph_containing(paras, "marker here")
    formula = paras[marker_i + 1]
    assert omath_in(formula) == 1, (
        f"marker 下一段没有公式：各段 oMath 数 = {[omath_in(p) for p in paras]}")

    # ← 承重断言 1：块级容器，且只在公式段里
    assert count(formula, m("oMathPara")) == 1, (
        "--display 没产出 m:oMathPara（退回行内了？）；该段子元素 = "
        f"{[etree.QName(c).localname for c in formula]}")
    assert count(root, m("oMathPara")) == 1, (
        f"全文档 oMathPara 数不是 1：{count(root, m('oMathPara'))}")
    assert count(root, m("oMath")) == 1, (
        f"全文档 oMath 数不是 1：{count(root, m('oMath'))}")

    # ← 承重断言 2：oMathPara 是公式段的直接子元素（不是套在 oMath 里）
    child_tags = [etree.QName(c).localname for c in formula]
    assert "oMathPara" in child_tags, (
        f"m:oMathPara 不在 <w:p> 直接子元素里：{child_tags}")

    # ← 承重断言 3：display 分支不补 U+200B 兜底 run
    assert "\u200b" not in texts[marker_i + 1], (
        f"块级公式段里混进了 U+200B 兜底 run（display 分支不该补）："
        f"{texts[marker_i + 1]!r}")


# ===========================================================================
# P5-A1 · 9) --dump：真只读 + stdout 是可解析的 <w:p>
# ===========================================================================
@needs_omml
def test_dump_switch_writes_nothing_and_prints_parsable_paragraph(tmp_path: Path):
    r"""``--dump``（单独使用）：``document.xml`` 的 SHA256 与段落数**都不变**，
    stdout 可解析。

    ``main()` L131–L138 在 ``--dump`` 分支里 **return 0**，位于
    ``doc_xml.exists()`` 检查、``parse()``、``insert_near()``、``write()`` **之前** ——
    所以它连 ``word/document.xml`` 是否存在都不看，纯打印。

    实测：单独 ``--dump``（不给任何锚点），rc=0、``document.xml`` 的 SHA256 与段落数
    均不变、stdout 以 ``<w:p `` 开头且含 1 个 ``m:oMath``。

    刻意**不给** ``--after``：这条只测「dump 本身只读」这一个语义；锚点与 dump 的
    组合归第 12 条。两件事分开，反向自证时才能各打各的（一条改坏只红一条）。

    「只断 rc==0 + stdout 有 oMath」不够：把 ``--dump`` 实现成「先写文件再打印」
    照样能过。所以必须同时断 **文件没动**（SHA256 + 段落数）与 **stdout 是可解析
    的 XML 且恰好 1 个 oMath**（不是一句提示语里含 "oMath" 字样）。

    同时断 ``Inserted`` 不在 stdout：那是写入分支的措辞，出现即证明它穿透了 dump 的
    ``return 0``。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))

    proc = run_insert(up, r"\frac{a}{b}", "--dump")
    assert proc.returncode == 0, (
        "rc=%s stderr=%r" % (proc.returncode,
                            proc.stderr.decode("utf-8", "replace")))

    # ← 承重断言 1：文件一个字节都没动
    assert doc_sha(up) == sha_before, (
        "--dump 仍改写了 document.xml（只读语义被破坏）："
        f"改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, (
        f"--dump 增删了段落：改前 {paras_before} 段 → 改后 "
        f"{len(body_paragraphs(up))} 段")

    # ← 承重断言 2：stdout 是可直接解析的 <w:p>，含恰好 1 个 oMath
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "<w:p" in stdout, f"stdout 不是 XML（没打印段落）：{stdout[:200]!r}"
    dumped = etree.fromstring(stdout.encode("utf-8"))
    assert etree.QName(dumped).localname == "p", (
        f"stdout 根元素不是 w:p：{etree.QName(dumped).localname}")
    assert omath_in(dumped) == 1, (
        f"dump 出来的段落里 oMath 数不是 1：{omath_in(dumped)}")
    assert count(dumped, m("limUpp")) == 0, "dump 出来的公式带着未修的 limUpp"


# ===========================================================================
# P5-A1 · 10) --raw-spacing：与默认路径的 OMML 结构差异
# ===========================================================================
@needs_omml
def test_raw_spacing_keeps_latex_spacing_instead_of_nbsp_run(tmp_path: Path):
    r"""``--raw-spacing``：不做 ``\quad → \ `` 改写，产物 **少一个 m:r**。

    默认路径经 ``math_omml.keep_spacing()``：
    ``re.sub(r"\\(?:qquad|quad|,|;|:|!)", r"\\ ", latex)`` —— ``\quad`` 被改写成
    控制空格 ``\ ``，产物里落成一个 **NBSP run**（``m:rPr/m:sty@val="p"`` +
    ``m:t`` = ``\u00a0``）。``--raw-spacing`` 跳过这一步，``\quad`` 原样进
    ``latex2mathml``，XSLT 把这个间距命令丢掉。

    实测 ``a\quad b`` 两条路径的 ``m:r`` 数与文本：

    | 路径 | m:r 数 | m:t 序列 |
    |---|---|---|
    | 默认（无开关） | 3 | ``['a', '\u00a0', 'b']`` |
    | ``--raw-spacing`` | 2 | ``['a', 'b']`` |

    所以断言口径是「精确差额」（``default_runs == raw_runs + 1`` + NBSP 只在默认
    路径出现），不是模糊的「两者不同」——后者对「开关把两条路径都跑了一遍」这类
    回归免疫。

    两组都必须先断 ``oMath == 1``：raw 路径若直接转换失败/降级成纯文字，``m:r`` 数
    差异会来自「公式压根没生成」，差异断言就成了假证据。
    """
    latex = r"a\quad b"

    def omath_texts(unpacked: Path) -> tuple[int, list[str]]:
        root = xml_root(unpacked)
        oms = list(root.iter(m("oMath")))
        assert len(oms) == 1, f"oMath 数不是 1：{len(oms)}（公式没生成或生成多处）"
        om = oms[0]
        runs = list(om.iter(m("r")))
        return len(runs), ["".join(t.text or "" for t in r.iter(m("t")))
                           for r in runs]

    # 两条路径各造一份 fixture：同一个解包目录跑两次，第二遍会读到已被改写过的
    # document.xml（多出一个 oMath），`omath == 1` 前置断言会先炸。
    (tmp_path / "def").mkdir()
    (tmp_path / "raw").mkdir()
    up_default = make_unpacked(tmp_path / "def", ["alpha one", "marker here"])
    up_raw = make_unpacked(tmp_path / "raw", ["alpha one", "marker here"])

    p_def = run_insert(up_default, latex, "--after", "marker here")
    p_raw = run_insert(up_raw, latex, "--after", "marker here", "--raw-spacing")
    for label, proc in (("默认", p_def), ("--raw-spacing", p_raw)):
        assert proc.returncode == 0, (
            f"{label} 路径失败：rc=%s stderr=%r"
            % (proc.returncode, proc.stderr.decode("utf-8", "replace")))

    n_def, texts_def = omath_texts(up_default)
    n_raw, texts_raw = omath_texts(up_raw)

    assert n_def == n_raw + 1, (
        "默认路径应比 --raw-spacing 多一个 run（\\quad → NBSP）：\n"
        f"  默认 {n_def} 个 m:r，文本 {texts_def}\n"
        f"  raw  {n_raw} 个 m:r，文本 {texts_raw}")
    assert "\u00a0" in texts_def, (
        f"默认路径没把 \\quad 改写成 NBSP run：{texts_def}")
    assert "\u00a0" not in texts_raw, (
        f"--raw-spacing 仍做了 \\quad → NBSP 改写（开关没生效）：{texts_raw}")
    assert texts_raw == ["a", "b"], (
        f"--raw-spacing 的 m:t 序列不是 ['a', 'b']（间距应被 XSLT 丢掉）：{texts_raw}")


# ===========================================================================
# P5-A1 · 11) 表格单元格内 marker：--after 真插进那个 w:tc
# ===========================================================================
@needs_omml
def test_after_marker_inside_table_cell_inserts_into_that_cell(tmp_path: Path):
    r"""``--after`` 命中表格单元格里的文字 → 公式段落**插进该单元格**。

    ``docx_utils.insert_near`` 的容器列表是 ``[body] + list(body.iter(w:tc))``，
    即**刻意**支持单元格内锚点。但现有 6 条里没有任何一条带表格，这条路径长期
    零覆盖。

    实测（body = ``p(alpha one)`` / ``tbl`` / ``p(omega three)``，表格 1×1 单元格
    文字 ``cell marker inside``）：rc=0，``w:tc`` 子元素由
    ``['tcPr', 'p']`` 变成 ``['tcPr', 'p', 'p']``，新段落含 1 个 ``m:oMath``，
    **body 直接子元素序列不变**（仍是 ``['p','tbl','p','sectPr']``，
    公式没有漏到 body 层）。

    这条同时钉住「不许 fallback」：若 ``insert_near`` 不再搜 ``w:tc``，marker 找不到
    时主线会 ``return 1``（`insert_formula` L159–L164 把 anchored=False 当错误），
    而**不是**静默追加到文末 —— rc 断言与「body 段落数不变」两条一起兜住这个。
    """
    d = docx.Document()
    d.add_paragraph("alpha one")
    table = d.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "cell marker inside"
    d.add_paragraph("omega three")
    src = tmp_path / "tbl.docx"
    d.save(str(src))
    freeze_docx(src)          # zip 时间戳钉死（见 freeze_docx docstring）

    up = tmp_path / "unpacked_tbl"
    unpack(src, up)
    assert doc_xml(up).is_file(), "解包没产出 word/document.xml"

    body_before = [etree.QName(c).localname for c in xml_root(up).find(w("body"))]
    assert body_before == ["p", "tbl", "p", "sectPr"], (
        f"fixture 结构不对（造表格失败）：{body_before}")

    proc = run_insert(up, "E=mc^2", "--after", "cell marker inside")
    assert proc.returncode == 0, (
        "单元格内 marker 未命中（insert_near 没搜 w:tc？）：rc=%s stderr=%r"
        % (proc.returncode, proc.stderr.decode("utf-8", "replace")))

    root = xml_root(up)
    body = root.find(w("body"))
    paras = [c for c in body if c.tag == w("p")]
    assert len(paras) == 2, (
        "公式段落漏到了 body 层（应插进单元格）：body 段落数 = "
        f"{[para_text(p) for p in paras]}")

    cells = list(body.iter(w("tc")))
    assert len(cells) == 1, f"表格单元格数不是 1：{len(cells)}"
    tc = cells[0]
    tc_paras = [c for c in tc if c.tag == w("p")]
    assert len(tc_paras) == 2, (
        "单元格内没多出公式段：tc 子元素 = "
        f"{[etree.QName(c).localname for c in tc]}")

    cell_texts = [para_text(p) for p in tc_paras]
    assert cell_texts[0] == "cell marker inside", (
        f"marker 段被挪动了（应保持第一段）：{cell_texts}")
    assert omath_in(tc_paras[1]) == 1, (
        f"单元格第二段不是公式段（--after 方向错？）：{cell_texts}，"
        f"各段 oMath 数 = {[omath_in(p) for p in tc_paras]}")
    assert omath_in(root) == 1, f"全文档 oMath 数不是 1：{omath_in(root)}"


# ===========================================================================
# P5-A1 · 12) 【P6-FIX-C2 反转】--after 与 --dump 同时给：用法错误 exit 2
# ===========================================================================
def test_after_with_dump_is_rejected_as_usage_error(tmp_path: Path):
    r"""``--after`` + ``--dump`` → **exit 2 + 明确报错**，文档零改动。

    ## 这条用例此前锁的是缺陷本体（之所以反转，别当成随手改）

    P5-A1 版把它写成「``--dump`` 先 ``return 0``、``--after`` 被完全忽略，rc=0」，
    并称之为「``--dump`` 的只读优先级高于锚点参数」这一既有语义。但 P5-AUDIT-QUALITY
    §5.7 的判定是：这是**静默丢弃参数**——调用方给了锚点、拿到 rc=0，会以为公式插在
    了锚点旁边，实际连文档都没打开。锁住缺陷的用例会把它变成"特性"，比没有用例更糟。

    修法选「明确互斥退 2」而不是「两个都生效」：``--dump`` 的语义是**只打印不写文件**，
    ``--after`` 的语义是**决定写在哪儿**——在不写文件的前提下，锚点没有任何可生效的
    落点，所以"两个都生效"不存在，只剩"明确报错"这一条诚实的路（任务书也允许二选一）。

    ## 为什么这条**不再**挂 ``@needs_omml``

    用法校验现在排在 ``build_para()`` **之前**（旧版是 ``--dump`` 的 ``return 0``
    抢在最前）：矛盾参数在需要 XSL 之前就退 2，所以这条用例在没装 Office 的机器上
    也能跑，不再 skip。这也是修法的一个副产品——用法错误不该依赖转换环境。

    断言四层（缺一层都挡不住一类退化）：
    * rc == 2（``!= 0`` 不够：exit 1 是运行期错误的码，用法错误必须可区分）；
    * stderr 点出「锚点会被忽略」而不是笼统 usage；
    * ``document.xml`` SHA + 段落数不变（报错的同时落盘是最坏结果）；
    * stdout 既没有 OMML（``<w:p``）也没有 ``Inserted``（没写文件也没打印）。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))

    proc = run_insert(up, "E=mc^2", "--after", "marker here", "--dump")
    stdout = proc.stdout.decode("utf-8", "replace")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode == 2, (
        "--after + --dump 必须退 2（用法错误）：rc=%s stdout=%r stderr=%r"
        % (proc.returncode, stdout[:200], stderr[:300]))
    assert "cannot be combined" in stderr, (
        f"报错没说明两个参数互斥：{stderr!r}")
    assert "--after" in stderr and "--dump" in stderr, (
        f"报错没指名是哪两个参数：{stderr!r}")
    assert "silently ignored" in stderr, (
        f"报错没说清「给了会怎样」（锚点被静默忽略）：{stderr!r}")

    assert doc_sha(up) == sha_before, (
        "报错的同时改写了 document.xml："
        f"改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, "报错的同时增删了段落"
    assert omath_in(xml_root(up)) == 0, "报错的同时把公式写进了文档"

    assert "<w:p" not in stdout, (
        f"互斥时不该还打印 OMML（说明 --dump 分支先跑了）：{stdout[:200]!r}")
    assert "Inserted" not in stdout, f"stdout 出现了写入分支的措辞：{stdout[:200]!r}"

# ===========================================================================
# P6-FIX-C2 · 13) 空 marker：旧版 rc=0 且静默追加到文末 → 现在用法错误 exit 2
# ===========================================================================
def test_empty_marker_is_rejected_instead_of_appending_at_end(tmp_path: Path):
    r"""``--after ""`` / ``--before " "`` → **exit 2**，文档零改动。

    ## 缺陷形态（旧版实测，P6-FIX-C2 §缺陷亲验 S3）

    ``--after ""`` 在旧版里走到的是「**没有锚点**」这条分支，而不是「锚点为空」：

    ```
    marker = args.after or args.before      # "" 是 falsy → marker == ""
    anchored = insert_near(body, "", ...)   # `if marker:` 也是 falsy → 直接 fallback
    if marker and not anchored:             # marker 为 "" → 这个防护根本不进
        ...
    ```

    于是**三个守卫全部被空串绕过**：公式被追加到文末、rc=0、stdout 打印
    ``Inserted inline formula: E=mc^2``。调用方传了锚点、拿到成功，公式却在文末
    ——正是任务书点名的「静默说谎」。空串不是"没要求"，它是"要求了一个空串"。

    ## 断言口径

    两种写法都测（``--after ""`` 与 ``--before " "``，后者连空白-only 也覆盖）：
    必须 rc=2、stderr 出现 ``empty marker``、``document.xml`` SHA 与段落数不变。
    段数断言是承重的：旧版这条正好多出 1 段（公式段），只断 rc 的话，
    把错误码改对但仍先追加再报错的实现能混过去。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))

    for flag, value in (("--after", ""), ("--before", " ")):
        proc = run_insert(up, "E=mc^2", flag, value)
        stderr = proc.stderr.decode("utf-8", "replace")
        assert proc.returncode == 2, (
            f"{flag} {value!r} 必须退 2（用法错误），否则就是静默追加到文末："
            "rc=%s stdout=%r stderr=%r"
            % (proc.returncode, proc.stdout.decode("utf-8", "replace")[:200],
               stderr[:300]))
        assert "empty marker" in stderr, (
            f"{flag} {value!r} 的报错没点明是空 marker：{stderr!r}")
        assert "silently append" in stderr, (
            f"{flag} {value!r} 的报错没说清后果（会被静默追加到文末）：{stderr!r}")

    assert doc_sha(up) == sha_before, (
        "空 marker 的失败路径仍改写了 document.xml："
        f"改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, (
        f"空 marker 仍往文档里塞了段落：改前 {paras_before} 段 → 改后 "
        f"{[para_text(p) for p in body_paragraphs(up)]}")
    assert omath_in(xml_root(up)) == 0, "空 marker 仍把公式写进了文档"


# ===========================================================================
# P6-FIX-C2 · 14) marker 找不到：报错必须给出实际可搜范围
# ===========================================================================
@needs_omml      # build_para() 先于 marker 校验，见模块 docstring
def test_marker_not_found_reports_actual_search_range(tmp_path: Path):
    r"""marker 未命中的报错里必须写清**实际范围** ``0-based [0, N-1]``。

    「marker not found: 'x'」这种只报事实、不报范围的消息，让人没法判断是自己
    拼错了、还是文档根本不是那份（3 段还是 300 段，处理方式完全不同）。这一条
    把范围披露钉进契约——也正是「越界索引要报实际范围 [0, N-1]」这条要求在
    insert_formula 上的对应物（本脚本没有 ``--index``，见 P6-FIX-C2 报告 §1.1）。

    ## 这里的 N 必须是**插入前**的段数（写这条用例时踩到的真坑）

    ``insert_near()`` 在 marker 未命中时的实现是：**先把公式段追加到文末、再返回
    False**。旧实现若在 ``insert_near()`` **之后**数段数，会数出 4 段（3 段正文 +
    刚落进去的公式段），报出 ``[0, 3]`` —— 一个比真相大 1 的假范围。实测抓到这个
    形态后才改成插入前取快照。所以这里断死 ``3 body paragraph(s)`` 与 ``[0, 2]``：
    报大了就说明范围是在插入之后数的。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    sha_before = doc_sha(up)
    paras_before = len(body_paragraphs(up))
    assert paras_before == 3, f"fixture 段落数不对：{paras_before}"

    proc = run_insert(up, "E=mc^2", "--after", "no-such-marker")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode == 1, (
        "marker 未命中仍是 exit 1（运行期错误，与用法错误的 2 区分开）："
        "rc=%s stderr=%r" % (proc.returncode, stderr))
    assert "[0, 2]" in stderr, (
        f"报错没给出实际可搜范围 [0, 2]：{stderr!r}")
    assert "3 body paragraph" in stderr, (
        f"报错没给出段数（且必须是插入前的 3）：{stderr!r}")
    assert "table cell" in stderr, (
        f"报错没说明还搜过单元格：{stderr!r}")

    assert doc_sha(up) == sha_before, (
        "marker 未命中的路径仍改写了 document.xml："
        f"改前 {sha_before[:12]} → 改后 {doc_sha(up)[:12]}")
    assert len(body_paragraphs(up)) == paras_before, "失败路径仍增删了段落"


# ===========================================================================
# P6-FIX-C2 · 15) 落点回显 = 插入后反查到的实测下标（不是入参、不是猜测）
# ===========================================================================
@needs_omml
def test_position_echo_reports_actual_body_index(tmp_path: Path):
    r"""成功输出里的 ``at body[i]`` 必须**等于**公式段在 ``<w:body>`` 里的真实下标。

    旧版什么都不回显（``Inserted inline formula: E=mc^2`` 就完了），调用方无从校对
    公式落在哪儿——这正是 ``docx_edit insert`` 那个「``inserted_at`` 回显入参 8、
    实际落在第 4 位」缺陷的**同族**（P5-AUDIT-QUALITY §2.5）。

    新回显由**插入后到树上反查**得到（``_locate()`` 按对象身份在下标上找），所以它
    天然是真值。但要防两种退化：
    * 回显改成"入参回显"（比如把 marker 段的下标当结果）→ 与实测下标不符就红；
    * 回显永远报 0 / 报固定值 → 与实测下标不符就红。

    三条断言：① stdout 里有且只有一处回显；② 该下标对应的 body 子元素就是含 oMath
    的段落；③ 该下标 == marker 段下标 + 1（交叉核对方向没反、也没偏差）。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    proc = run_insert(up, "E=mc^2", "--after", "marker here")
    assert proc.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (proc.returncode,
                                     proc.stderr.decode("utf-8", "replace")))
    stdout = proc.stdout.decode("utf-8", "replace")

    hits = ECHO_RE.findall(stdout)
    assert len(hits) == 1, f"落点回显应恰好一处，实际 {len(hits)} 处：{stdout!r}"

    kind, idx = echo_of(stdout)
    assert kind == "body", f"body 层的落点被报成了 {kind}：{stdout!r}"

    kids = body_children(up)
    assert 0 <= idx < len(kids), (
        f"回显下标 {idx} 越界（body 直接子元素 {len(kids)} 个）：{stdout!r}")
    target = kids[idx]
    assert etree.QName(target).localname == "p", (
        f"回显的 body[{idx}] 不是段落而是 {etree.QName(target).localname}：{stdout!r}")
    assert omath_in(target) == 1, (
        f"回显的 body[{idx}] 里没有 oMath —— 回显的位置不是公式真正落点："
        f"各子元素 = {[etree.QName(c).localname for c in kids]}")

    marker_i = [i for i, c in enumerate(kids)
                if c.tag == w("p") and "marker here" in para_text(c)]
    assert marker_i, "fixture 里找不到 marker 段"
    assert idx == marker_i[0] + 1, (
        f"回显 body[{idx}] 与「marker 下标 + 1 = {marker_i[0] + 1}」不符：{stdout!r}")


# ===========================================================================
# P6-FIX-C2 · 16) 边界：首段（0）与末段（N-1）
# ===========================================================================
@needs_omml
def test_anchor_boundaries_first_and_last_paragraph(tmp_path: Path):
    r"""两端边界各验一次：``--before`` 首段 → ``body[0]``；``--after`` 末段 → ``body[N]``。

    没有 ``--index`` 的脚本，锚点越界的极端形态就是**贴着文档两端**：第一个段落
    之前（下标 0，前面没有任何位置可退）和最后一个段落之后（下标 N，后面只剩
    ``w:sectPr``）。这两处最容易出「差一位」，而差一位恰恰是 ``--index`` 类缺陷的
    典型症状（``crossref`` 0-based 与 ``docx_toc`` 1-based 混用就是这么来的）。

    实测（三段 fixture，body 子元素 = ``[p, p, p, sectPr]``）：

    * ``--before "alpha one"`` → 回显 ``at body[0]``，公式段成为**第一段**；
    * ``--after "omega three"``（末段下标 2 = N-1）→ 回显 ``at body[3]``，即 N，
      公式段落在 ``sectPr`` 之前、成为**最后一段正文**。

    第二条同时锁死「下标口径 = body 直接子元素下标」：若把 ``sectPr`` 排除在外去数
    （只数段落），末段之后会报 ``body[3]`` 与 ``body[2]`` 之争——这里断死 3。
    """
    # ---- 边界 0：--before 首段
    d1 = tmp_path / "first"
    d1.mkdir()
    up1 = make_unpacked(d1, ["alpha one", "marker here", "omega three"])

    p1 = run_insert(up1, "E=mc^2", "--before", "alpha one")
    assert p1.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (p1.returncode,
                                     p1.stderr.decode("utf-8", "replace")))
    kind, idx = echo_of(p1.stdout.decode("utf-8", "replace"))
    assert (kind, idx) == ("body", 0), (
        f"--before 首段应报 at body[0]，实际 {kind}[{idx}]")
    kids1 = body_children(up1)
    assert omath_in(kids1[0]) == 1, (
        "公式没落在 body[0]："
        f"{[etree.QName(c).localname for c in kids1]}")

    # ---- 边界 N-1：--after 末段
    d2 = tmp_path / "last"
    d2.mkdir()
    up2 = make_unpacked(d2, ["alpha one", "marker here", "omega three"])
    assert len(body_children(up2)) == 4, (
        f"fixture body 子元素数不对：{[etree.QName(c).localname for c in body_children(up2)]}")

    p2 = run_insert(up2, "E=mc^2", "--after", "omega three")
    assert p2.returncode == 0, (
        "插入失败：rc=%s stderr=%r" % (p2.returncode,
                                     p2.stderr.decode("utf-8", "replace")))
    kind, idx = echo_of(p2.stdout.decode("utf-8", "replace"))
    assert (kind, idx) == ("body", 3), (
        f"--after 末段（N-1=2）应报 at body[3]（= N），实际 {kind}[{idx}]")
    kids2 = body_children(up2)
    assert [etree.QName(c).localname for c in kids2] == ["p", "p", "p", "p", "sectPr"], (
        f"公式段没落在 sectPr 之前：{[etree.QName(c).localname for c in kids2]}")
    assert omath_in(kids2[3]) == 1, "body[3] 不是公式段（回显与实测不符）"


# ===========================================================================
# P6-FIX-C2 · 17) 表格单元格内的落点必须报 tc[i]，不许报成 body 下标
# ===========================================================================
@needs_omml
def test_position_echo_inside_table_cell_reports_tc_index(tmp_path: Path):
    r"""marker 在单元格里 → 回显 ``at tc[i]``，且 i 是该 ``w:tc`` 直接子元素的下标。

    这是「回显不许撒谎」最硬的一条：``insert_near()`` 会到 ``w:tc`` 里查找锚点，
    公式段落因此落在**单元格内部**、body 层的段落数不变。若回显偷懒写死成
    ``body[...]``，用户按那个下标去 body 里找公式会**找不到**——报一个必然对不上
    的坐标，比不报更糟。

    实测（body = ``p(alpha one)`` / ``tbl`` / ``p(omega three)``，单元格文字
    ``cell marker inside``）：``tc`` 的直接子元素由 ``['tcPr', 'p']`` 变成
    ``['tcPr', 'p', 'p']``，回显 ``at tc[2]`` —— 下标口径含 ``w:tcPr``，所以是 2
    而不是 1。这个数字是刻意断死的：它同时说明「下标 = 容器直接子元素下标」这条
    口径在 tc 上也成立（不是"第几段"）。

    同时断 body 段落数不变：回显说 tc、公式却漏到 body 层，是另一种"坐标撒谎"。
    """
    d = docx.Document()
    d.add_paragraph("alpha one")
    table = d.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "cell marker inside"
    d.add_paragraph("omega three")
    src = tmp_path / "tbl.docx"
    d.save(str(src))
    freeze_docx(src)

    up = tmp_path / "unpacked_tbl"
    unpack(src, up)

    proc = run_insert(up, "E=mc^2", "--after", "cell marker inside")
    assert proc.returncode == 0, (
        "单元格内 marker 未命中：rc=%s stderr=%r"
        % (proc.returncode, proc.stderr.decode("utf-8", "replace")))

    stdout = proc.stdout.decode("utf-8", "replace")
    kind, idx = echo_of(stdout)
    assert kind == "tc", (
        f"落点在单元格里，回显却报 {kind}[{idx}] —— 拿这个坐标去 body 里找不到公式："
        f"{stdout!r}")

    body = xml_root(up).find(w("body"))
    cells = list(body.iter(w("tc")))
    assert len(cells) == 1, f"表格单元格数不是 1：{len(cells)}"
    tc_children = list(cells[0])
    assert idx == 2, (
        f"tc 直接子元素下标应为 2（tcPr / marker 段 / 公式段），实际报 {idx}："
        f"{[etree.QName(c).localname for c in tc_children]}")
    assert etree.QName(tc_children[idx]).localname == "p", (
        f"tc[{idx}] 不是段落：{[etree.QName(c).localname for c in tc_children]}")
    assert omath_in(tc_children[idx]) == 1, (
        f"tc[{idx}] 里没有 oMath（回显的单元格下标不是真落点）："
        f"{[etree.QName(c).localname for c in tc_children]}")

    body_paras = [c for c in body if c.tag == w("p")]
    assert len(body_paras) == 2, (
        "公式段落漏到了 body 层（回显说 tc、实际在 body）："
        f"{[para_text(p) for p in body_paras]}")

# ===========================================================================
# P9-D · 输入守卫：`unpacked_dir` 的三档坏输入 → exit 2；`<w:body>` 缺失仍是 1
# ===========================================================================
# ⚠️ 本节**不挂** ``@needs_omml``，与模块 docstring 里「第 1/2/3 条必须挂」不矛盾
#    —— 那说的是 **marker 类**用例（`insert_near` 在 `build_para()` 之后），本节的
#    断言对象是**输入目录**，而 P9-D 已把目录校验整块挪到 `build_para()` **之前**。
#    实测（`MML2OMML_XSL` 指到不存在路径后重跑）本节 4 条全 **pass**、不是 skip：
#    目录不对时脚本在需要 XSL 之前就退 2，报的是 document.xml，不是 MML2OMML.XSL。
#
# 唯一要挂的是**合法对照**（最后一条）：它得真跑公式链，没有 XSL 就无从插公式。
def _bad_unpacked(tmp_path: Path, kind: str) -> tuple[Path, str]:
    """造一档坏 ``unpacked_dir``，返回 ``(目录, stderr 里应出现的措辞)``。"""
    if kind == "missing":
        # 刻意不创建：连目录本身都没有
        return tmp_path / "no-such-unpacked", "not found"
    d = tmp_path / f"bad-{kind}"
    d.mkdir()
    if kind == "no-document-xml":
        # 目录在、但里面没有 word/document.xml（只有无关文件）
        (d / "readme.txt").write_text("nothing here", encoding="utf-8")
        return d, "not found"
    # 目录在、document.xml 在，但不是良构 XML（标签不闭合）
    (d / "word").mkdir()
    (d / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body></w:document>',
        encoding="utf-8")
    return d, "not well-formed XML"


@pytest.mark.parametrize("kind",
                         ["missing", "no-document-xml", "malformed-document-xml"])
def test_bad_unpacked_dir_is_exit_2_single_line_no_traceback(kind, tmp_path: Path):
    r"""三档坏 ``unpacked_dir`` → **exit 2** + stderr **恰好一行** + stdout 空 + 无 traceback。

    本脚本的输入形态是**解包目录**而不是 .docx 包，所以不套 ``_check_input``
    （那个函数判「文件是不是 zip」，对着目录用会得出相反结论）。三档：

    | 档 | 旧版实测（2026-11-11） | 现在 |
    |---|---|---|
    | 目录不存在 | rc=1 + 1 行 ``error: …\word\document.xml not found`` | rc=**2** + 1 行 |
    | 目录在但无 document.xml | rc=1 + 同上 1 行 | rc=**2** + 1 行 |
    | document.xml 不良构 | rc=1 + **16 行裸 traceback** | rc=**2** + 1 行 |

    前三档的 argv 都是完整的 ``insert_formula.py DIR FORMULA``：只给 DIR 会以
    ``the following arguments are required: formula`` 退 2，那是 argparse 假绿
    —— 判别器是「stderr 里出现那个坏目录路径 + 该档特定措辞」。

    为什么必须同时断 stdout 为空与单行：旧版前两档的 rc 只差一个数字，第三个档
    却是 traceback。只断 ``rc != 0`` 的用例对 traceback 完全免疫（那正是 P9 终审
    抓到的形态）。
    """
    d, wording = _bad_unpacked(tmp_path, kind)
    d_sha_before = None
    docx_xml = d / "word" / "document.xml"
    if docx_xml.is_file():
        d_sha_before = hashlib.sha256(docx_xml.read_bytes()).hexdigest()

    proc = run_insert(d, "E=mc^2")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode == 2, (
        f"[{kind}] 坏目录必须退 2：rc={proc.returncode}\n"
        f"stdout={proc.stdout!r}\nstderr={stderr!r}")
    assert "Traceback (most recent call last)" not in stderr, (
        f"[{kind}] 仍然是裸 traceback（契约要求单行 error）：{stderr!r}")

    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, (
        f"[{kind}] stderr 必须恰好一行，实际 {len(lines)} 行：{stderr!r}")
    assert lines[0].startswith("error: "), (
        f"[{kind}] stderr 不是 `error: ` 前缀：{lines[0]!r}")
    assert wording in lines[0], (
        f"[{kind}] 报错没指出坏在哪（期望 {wording!r}）：{lines[0]!r}")
    assert str(d) in lines[0], (
        f"[{kind}] 报错没提那个坏目录 —— argparse 假绿的判别器：{lines[0]!r}")

    assert proc.stdout == b"", (
        f"[{kind}] 失败路径不许有 stdout（成功时打的是 Inserted …）：{proc.stdout!r}")

    if d_sha_before is not None:
        assert hashlib.sha256(docx_xml.read_bytes()).hexdigest() == d_sha_before, (
            f"[{kind}] 失败路径改写了 document.xml（零写盘被破坏）")


def test_missing_body_is_still_exit_1(tmp_path: Path):
    r"""**不许顺手改掉的那条**：XML 良构但缺 ``<w:body>`` → 仍是 **exit 1**。

    ``<w:body>`` 缺失是「文档真实存在、但内容不是一个可插入的文档」，属运行期
    错误。把它也提到 2 就把「输入给错了」和「文档结构不对」混成一档 —— 本单只该
    动输入档（目录不存在 / 无 document.xml / XML 不良构），不动这一档。

    不需要 XSL：校验与 ``<w:body>`` 查找都在 ``build_para()`` 之前。所以裸跑，不挂
    ``@needs_omml``。
    """
    d = tmp_path / "no-body"
    (d / "word").mkdir(parents=True)
    (d / "word" / "document.xml").write_text(
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"/>',
        encoding="utf-8")

    proc = run_insert(d, "E=mc^2")
    stderr = proc.stderr.decode("utf-8", "replace")

    assert proc.returncode == 1, (
        f"缺 <w:body> 仍是 exit 1（运行期语义）：rc={proc.returncode} "
        f"stderr={stderr!r}")
    assert "<w:body> not found" in stderr, f"报错没点明缺 body：{stderr!r}"
    assert proc.stdout == b"", "失败路径不许有 stdout"


@needs_omml
def test_valid_unpacked_dir_still_inserts_formula(tmp_path: Path):
    r"""**合法对照**：真解包目录 → exit 0 + 公式真的落进文档（防「一律报错」假修）。

    唯一挂 ``@needs_omml`` 的一条：它要真跑 LaTeX → MathML → OMML 转换链，
    没有 ``MML2OMML.XSL`` 就没有公式可插。

    断言三层：rc=0、stdout 是 ``Inserted inline formula``、``document.xml`` 里
    真的多出 1 个 ``m:oMath``。只断 rc 会被「rc=0 但什么都没插」混过去。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    assert omath_in(xml_root(up)) == 0, "fixture 前置假设失效：本文档本来就有 oMath"
    paras_before = len(body_paragraphs(up))

    proc = run_insert(up, "E=mc^2", "--after", "marker here")
    assert proc.returncode == 0, (
        "合法输入必须照常插入：rc=%s stderr=%r"
        % (proc.returncode, proc.stderr.decode("utf-8", "replace")))
    stdout = proc.stdout.decode("utf-8", "replace")
    assert "Inserted inline formula" in stdout, f"没走到写入分支：{stdout!r}"

    assert len(body_paragraphs(up)) == paras_before + 1, (
        f"段落数没 +1（公式没插进去）：{paras_before} → "
        f"{len(body_paragraphs(up))}")
    assert omath_in(xml_root(up)) == 1, (
        f"document.xml 里的 oMath 数不是 1：{omath_in(xml_root(up))}")
