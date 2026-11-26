# -*- coding: utf-8 -*-
r"""``docx_validate`` 的 **OMML 文本残留 LaTeX 命令** 检查（P27-B）。

## 本文件存在的理由

``latex2mathml`` 对**不认识的控制序列**不抛异常，而是把命令名当**字面量**
吐进 ``<mi>`` → ``MML2OMML.XSL`` 原样搬进 ``<m:t>``。于是交付物的公式里直接
印着 ``\eqno`` / ``\textwidth``，而整条链路四个环节全部沉默：转换不抛、
``docx_mathify`` 的 ``stats["failed"]`` 只数异常、退出码只看那个计数、
``assert_legal`` 只查结构不看文本。三份这样的产物全部拿到
``{"ok": true, "issues": []}`` + exit 0（取证见
``docs/reports/P27-PROBE-CONTROLSEQ.md``）。

本文件钉的是**事后体检**这一半：``docx_validate`` 新增 ``omml_latex_residue``。
治因那一半（转换时就拒收）在 ``math_omml.py``，由 P27-A 负责 —— 所以本文件
的夹具**一律手写 OMML**，不经过转换链，A 的行为怎么变都不影响这里的读数。

## 三条不容许松动的性质

1. **只报不拒收**：severity 是 ``warning``，``ok`` 仍 ``True``、CLI 仍 exit 0。
   这份报告喂给下游做 pass/fail，Word 能打开这类文档（只是画出反斜杠），
   体检不许把能开的文件判死 —— 与 ``omml_empty_slot`` 同一条理由。
2. **不误报**：``<w:t>`` 里的反斜杠是**合法正文**（``C:\Users\me`` 这种路径），
   绝不能报。所以扫描范围只能是 ``m:`` 命名空间里的 ``<m:t>``。
3. **无残留不出键**：不含残留的文档里连 ``omml_latex_residue`` 这个字符串都
   不该出现 —— 不是空列表、不是 ``count: 0``。给所有文档挂一条常驻噪音，
   等于教人忽略这份报告。

## 夹具分两类，门控也分两类

**绝大多数用例是纯 XML 层**：``python-docx`` 造包 + ``zipfile`` 手术 + 手写
``<m:t>``，不碰 ``MML2OMML.XSL`` / Office COM / LibreOffice，CI runner 没有
Office 也照常 pass。手写的是「不由本机转换链产生、但真实存在于历史产物里」
的那种文档 —— 这也让本文件不受 ``math_omml`` 改动影响。

**只有判据 7 那组（25 条泄漏覆盖率 + 合法公式不误报）挂 ``needs_omml``**：
它必须跑真实转换链，才能证明新判据看得见转换器真正产出的残留形态。
在没有 Office 的 runner 上 ``latex_to_omml_lxml`` 会降级成纯文字，那时没有
``m:t`` 可看，断言会假红 —— 所以这组 skip，其余照跑。

## 判据：整段锚定，套在**空白切出来的 token** 上

``LATEX_RESIDUE_RE = ^\\[A-Za-z]+\\*?$``，匹配对象是 ``m:t`` 文本
``.split()`` 后的每个 token。这一条同时满足两种残留形态：

- **数学模式**：命令独占一段 ``m:t>``，token 就是它本身；
- **文本模式**（``\text{…}``）：命令与正文共享一段，但转换器用 NBSP 分词，
  ``split()`` 也切 NBSP，所以命令仍独占一个 token。

同时 ``C:\Users\me`` / ``\Users\me`` / ``a\\b`` 这些路径内容都**不命中**
（单 token、或 token 不是纯命令）—— 那是本单第一版用「任意位置匹配」时误报的
东西。三条护栏用例常驻钉住，缺一条就漏一层。

已知边界见判据 8 那组：紧贴反斜杠（``a\zzz``）、标点紧跟（``\zzz,``）、
以及单独一个路径词（``见 \Users 目录`` 会报）。三条都实测、都写成用例。

Run:  python -m pytest tests/test_validate_omml_residue.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_validate.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

from _omml import needs_omml  # noqa: E402

CODE = "omml_latex_residue"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"


# ---- helpers -----------------------------------------------------------------

def run_validate(path) -> subprocess.CompletedProcess:
    """真子进程：退出码是 CLI 契约的一部分，in-process 测不到。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def rewrite_part(src: Path, dst: Path, part: str, transform) -> None:
    """Copy *src* to *dst*, replacing ``part`` with ``transform(bytes) -> bytes``."""
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                data = transform(data)
            zout.writestr(item, data)


def docx_with(tmp_path: Path, name: str, body_xml: str = "") -> Path:
    """A one-paragraph docx with *body_xml* appended to ``w:body``.

    Same construction as ``tests/test_validate_structure.py``: string surgery on a
    real python-docx file, not ``OxmlElement``, because the point is to produce
    markup that the writer APIs never create — exactly what an older converter
    left behind.  ``xmlns:m`` is already declared on the root element by
    python-docx, so the injected ``m:`` prefix needs no declaration of its own.
    """
    base = tmp_path / "_base.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(base))
    target = tmp_path / name
    rewrite_part(
        base, target, "word/document.xml",
        lambda data: data.replace(
            b"</w:body>",
            body_xml.encode("utf-8") + b"</w:body>", 1))
    base.unlink()
    return target


def math_para(*texts: str) -> str:
    """A paragraph holding one math run per ``*texts*`` entry.

    Each text goes into its **own** ``<m:t>``, which is what the converter
    actually produces (one command, one run) and what ``count`` is defined on.
    """
    runs = "".join(f"<m:r><m:t>{t}</m:t></m:r>" for t in texts)
    return f"<w:p><m:oMath>{runs}</m:oMath></w:p>"


def residue(rep: dict) -> list:
    return [i for i in rep["issues"] if i["code"] == CODE]


def validate_inproc(path) -> dict:
    import docx_validate
    return docx_validate.validate(str(path))


# ===========================================================================
# 判据 1：命中 → warning + count + samples，且 ok 仍 True、CLI 仍 exit 0
# ===========================================================================
def test_residue_is_reported_as_warning(tmp_path: Path):
    r"""``<m:t>\eqno</m:t>`` → ``omml_latex_residue``，warning，count 1。

    这条同时钉四件事，任何一件错就红：键出现、severity 是 warning（不是
    error）、``count`` 数的是命中的 ``<m:t>`` 个数、``ok`` 仍是 True。
    只测「有没有这个键」会让一个判成 error 的实现也过。
    """
    src = docx_with(tmp_path, "hit.docx",
                    math_para(r"\eqno"))

    rep = validate_inproc(src)
    hits = residue(rep)
    assert len(hits) == 1, rep["issues"]
    hit = hits[0]
    assert hit["severity"] == "warning", "能打开的文件不许判 error"
    assert hit["count"] == 1
    assert hit["samples"] == [r"\eqno"]
    assert "document.xml" in hit["detail"]
    assert rep["ok"] is True, "有残留不等于文档不合法"


def test_cli_exit_code_stays_zero(tmp_path: Path):
    """CLI 层：``ok:true`` + **exit 0** + stderr 空。

    退出码契约 D-024 只有 2/1/3 三档，这条检查一档都不许动。在 ``validate()``
    里断言 ``ok is True`` 还不够 —— 真正被下游消费的是退出码。
    """
    src = docx_with(tmp_path, "hit_cli.docx", math_para(r"\eqno"))

    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == "", "warning 不该往 stderr 写东西"
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert [i["code"] for i in rep["issues"]] == [CODE]


# ===========================================================================
# 判据 2：多命令 → count 逐 run 计、samples 全覆盖
# ===========================================================================
def test_three_commands_three_runs(tmp_path: Path):
    r"""``\eqno`` + ``\nonumber`` + ``\textwidth``（三个独立 ``m:t``）→ count 3。

    三条就是取证报告里那三份坏产物同时出现在一个文档里的形状。
    """
    src = docx_with(tmp_path, "three.docx",
                    math_para(r"\eqno", r"\nonumber", r"\textwidth"))

    hits = residue(validate_inproc(src))
    assert len(hits) == 1
    assert hits[0]["count"] == 3
    assert hits[0]["samples"] == [r"\eqno", r"\nonumber", r"\textwidth"]


def test_samples_are_truncated_to_three(tmp_path: Path):
    """5 个不同命令 → ``count == 5`` 而 ``len(samples) == 3``。

    ``count`` 与 ``samples`` 是两个不同的问题：「坏了几处」要对得上，
    「坏的是哪些」只要给个开头 —— 否则一条 40 命令的公式会把报告淹掉。
    """
    cmds = [r"\eqno", r"\nonumber", r"\textwidth", r"\hfill", r"\zzz"]
    src = docx_with(tmp_path, "five.docx", math_para(*cmds))

    hit = residue(validate_inproc(src))[0]
    assert hit["count"] == 5
    assert len(hit["samples"]) == 3
    assert hit["samples"] == cmds[:3], "samples 按文档序取前三条"


def test_same_command_twice_counts_runs_not_names(tmp_path: Path):
    r"""同一个 ``\zzz`` 出现在两个 ``m:t`` → count 2，samples 只有一条。

    去重去的是**样本**，不是**计数**：两处坏了就是两处，报 1 会低估，
    报 2 条样本又是重复噪音。
    """
    src = docx_with(tmp_path, "dup.docx",
                    math_para(r"\zzz") + math_para(r"\zzz"))

    hit = residue(validate_inproc(src))[0]
    assert hit["count"] == 2, "命中的 m:t 个数是 2"
    assert hit["samples"] == [r"\zzz"], "同名命令只留一条样本"


# ===========================================================================
# 判据 3：不误报（本文件最重要的一组）
# ===========================================================================
def test_normal_formula_is_clean(tmp_path: Path):
    """正常公式（``x`` / ``−`` / ``max`` 一类纯文本 ``m:t``）→ 不出键。

    夹具里的 ``m:t`` 内容是真实转换会产生的那些：数字、运算符、函数名。
    一个「只要文档里有 ``m:`` 就报」的实现会在这里红。
    """
    body = (math_para("x") + math_para("\u2212") + math_para("max")
            + math_para("\u2202", "y"))
    src = docx_with(tmp_path, "clean_math.docx", body)

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]
    assert rep["ok"] is True


def test_backslash_in_body_text_is_not_reported(tmp_path: Path):
    r"""``<w:t>C:\Users\me</w:t>`` → **不报**。这条是本文件的护栏。

    正文里的反斜杠是合法内容（Windows 路径、正则示例、转义写法），
    ``m:`` 之外的东西一概不进扫描。把 ``w:t`` 一起扫的实现在这里必红 ——
    而那种误报会把每一份带路径的文档判上一条 warning。
    """
    body = (r'<w:p><w:r><w:t xml:space="preserve">C:\Users\me\report.docx</w:t>'
            r'</w:r></w:p>'
            r'<w:p><w:r><w:t xml:space="preserve">\eqno \textwidth \zzz</w:t>'
            r'</w:r></w:p>')
    src = docx_with(tmp_path, "body_backslash.docx", body)

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]
    assert rep["ok"] is True


def test_plain_document_has_no_key(tmp_path: Path):
    """纯文字 docx（无公式）→ 报告里连这个键都不出现。

    断言的是 ``CODE not in json.dumps(rep)``，不是「列表为空」：后者放过
    ``count: 0`` 那种常驻噪音，而正是那种噪音让人学会忽略报告。
    """
    src = docx_with(tmp_path, "plain.docx")

    rep = validate_inproc(src)
    assert rep["ok"] is True
    assert rep["issues"] == []
    assert CODE not in json.dumps(rep), "无残留的文档不许出现这个键"


def test_document_without_math_elements_is_clean(tmp_path: Path):
    """良构但整个文档没有任何 ``m:`` 元素 → 正常返回，不崩、不出键。

    与上一条的区别在**前提**：这里刻意塞了带公式感的东西（``w:drawing``
    之外的普通 run），确认扫描器不会因为「找不到 ``m:t``」而做奇怪的事。
    """
    body = '<w:p><w:r><w:t>x = 1</w:t></w:r></w:p>'
    src = docx_with(tmp_path, "no_math_ns.docx", body)

    rep = validate_inproc(src)
    assert rep["ok"] is True
    assert residue(rep) == []
    assert CODE not in json.dumps(rep)


# ===========================================================================
# 判据 4：与既有 omml_empty_slot 共存
# ===========================================================================
def test_coexists_with_empty_slot(tmp_path: Path):
    """空槽 + 残留同文 → 两条 issue 都在（顺序不作要求）。

    两条检查紧邻，最容易犯的错是提前 ``return`` 或共用变量把后一条吃掉。
    """
    body = ('<w:p><m:oMath><m:d><m:e/><m:e><m:r><m:t>x</m:t></m:r></m:e>'
            '</m:d></m:oMath></w:p>'
            + math_para(r"\eqno"))
    src = docx_with(tmp_path, "both.docx", body)

    rep = validate_inproc(src)
    codes = [i["code"] for i in rep["issues"]]
    assert "omml_empty_slot" in codes, rep["issues"]
    assert CODE in codes, rep["issues"]
    assert rep["ok"] is True, "两条都是 warning，不该判死"


# ===========================================================================
# 判据 5：已知边界（如实钉住，不假装它不存在）
# ===========================================================================
def test_lone_backslash_does_not_match(tmp_path: Path):
    r"""孤立的 ``\``（``\backslash`` 的产物）**不**报 —— 判据要命令名。

    单个反斜杠不是控制序列。把它也算进去会连「公式里就是个反斜杠字符」
    这种正确内容一起报了，而那种内容 Word 画出来就是想要的样子。
    """
    src = docx_with(tmp_path, "lone.docx", math_para("\\"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]


def test_windows_path_inside_math_text_is_not_reported(tmp_path: Path):
    r"""``m:t`` 文本是 ``C:\Users\me`` → **不报**（旧判据会误报，已修）。

    ``\text{C:\Users\me}`` 转成 OMML 后，那一段 ``<m:t>`` 的文本就是
    ``C:\Users\me``。旧判据 ``\\[A-Za-z]+``（任意位置命中）会把 ``\Users``
    当成命令名报出来；新判据要求**整段就是**一条命令，``C:\Users\me`` 开头是
    ``C:``，不满足 ⇒ 不报。

    这条是本单最重要的一格：它同时证明「收紧判据没有把覆盖弄丢」——
    ``test_residue_is_reported_as_warning`` 那批仍然绿。
    """
    src = docx_with(tmp_path, "path_in_math.docx",
                    math_para("C:\\Users\\me"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]
    assert rep["ok"] is True


def test_backslash_leading_text_is_not_reported(tmp_path: Path):
    r"""``m:t`` 文本是 ``\Users\me``（``\text{\Users\me}`` 的产物）→ **不报**。

    这一条比上面那条更严：``C:\Users\me`` 开头是 ``C``，只靠「开头锚定」就能
    挡住；而 ``\Users\me`` **开头就是反斜杠**，只有「整段锚定」（末尾也要锚）
    才挡得住 —— 它挡住的是 ``\Users`` 这个路径片段被当成命令名。

    缺了这条，判据从 ``^\\[A-Za-z]+\\*?$`` 放宽成 ``^\\[A-Za-z]+`` 会静悄悄
    通过全部用例，而 ``\text{\Users\me}`` 这类写法就会开始误报。
    """
    src = docx_with(tmp_path, "leading_backslash.docx",
                    math_para("\\Users\\me"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]


def test_double_backslash_text_is_not_reported(tmp_path: Path):
    r"""``m:t`` 文本是 ``a\\b``（``\text{a\\b}`` 的产物）→ **不报**。

    ``a\\b`` 开头是 ``a``，不是反斜杠 ⇒ 整段锚定天然不命中。这一条覆盖
    "正文里写了转义反斜杠" 的那类内容：它在公式里是合法文本，不是残留。
    """
    src = docx_with(tmp_path, "escaped_text.docx",
                    math_para("a\\\\b"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]


# ===========================================================================
# 判据 6：接线 —— 新检查必须真的在 validate() 里跑到
# ===========================================================================
def test_scan_is_reachable_from_validate():
    """``validate()`` 的源码里必须有这次扫描（防「写了函数没接上」）。

    所有用例都是走 ``validate()`` 的，所以这条严格来说是冗余的；留着是因为
    实现刻意**不抽成独立函数**（就地扫，与相邻的 ``omml_empty_slot`` 同形），
    一旦有人把它抽走却忘了调用，上面那些用例会全部变绿 —— 没有输入就没有
    断言，``residue() == []`` 依然成立。
    """
    import inspect

    import docx_validate

    src = inspect.getsource(docx_validate.validate)
    assert "omml_latex_residue" in src, "扫描没接进 validate()"
    assert "LATEX_RESIDUE_RE" in src, "判据正则没被用到"


def test_does_not_import_math_omml():
    """``docx_validate`` 的模块对象里不许出现 ``math_omml`` —— 独立实现是设计要求。

    ``math_omml`` 顶层拉 ``latex2mathml``，运行时还要 Office 的
    ``MML2OMML.XSL``；体检脚本必须在一台转换依赖不全的机器上照常可用。
    import 一行就会把两个重依赖挂在健康检查上。

    断言的是**模块对象没有这个属性**，不是「源码里没有这个词」：两种写法都
    在禁止之列（``import math_omml`` 与 ``from math_omml import x``），而源码
    里这个词还会出现在注释里（相邻那节就在解释为什么不 import 它）。属性检查
    把这几种 import 形态一次盖住，且不误伤注释。
    """
    import docx_validate

    assert getattr(docx_validate, "math_omml", None) is None, \
        "docx_validate 导入了 math_omml，独立实现被破坏"
    for name in dir(docx_validate):
        obj = getattr(docx_validate, name, None)
        mod = sys.modules.get(getattr(obj, "__module__", ""), None)
        assert getattr(mod, "__name__", "") != "math_omml" or name.startswith("__"), \
            f"docx_validate.{name} 来自 math_omml"

# ===========================================================================
# 判据 6b：判据本身 —— 必须是整段锚定，不是任意位置命中
# ===========================================================================
def test_pattern_is_whole_run_anchored():
    r"""判据必须是**整段锚定**：``^\\[A-Za-z]+\\*?$``，不是任意位置命中。

    为什么单独钉这一条：从 loose 收紧到 anchored 是**唯一的误报防线**。
    改回 ``re.compile(r"\\[A-Za-z]+")`` 之后，上面那些残留用例仍然全绿
    （残留本来就是整段），只有 ``test_windows_path_inside_math_text_is_not_reported``
    会红 —— 只看「命中」类用例拦不住这个回归，所以直接对正则本身断言。
    """
    import docx_validate

    rx = docx_validate.LATEX_RESIDUE_RE
    # 必须命中：残留的真实形态（含带星号的变体）
    for text in (r"\eqno", r"\textwidth", r"\vspace*", r"\zzz", r"\a"):
        assert rx.match(text), f"{text!r} 该命中却没命中"
    # 不许命中：正文路径、转义反斜杠、UNC 路径 —— 开头都不是控制序列
    for text in (r"C:\Users\me", r"a\\b", r"\\server\share",
                 r"路径 D:\E\BUAA", r"\Users\me", r"\tmp\x"):
        assert not rx.match(text), f"{text!r} 不该命中却命中了"


# ===========================================================================
# 判据 7：覆盖率 —— 25 条泄漏矩阵必须条条中（收紧判据不许漏）
# ===========================================================================

#: 取证报告 §② 的 25 条泄漏，按原式。每条走**真实转换链**再回读 ``m:t``，
#: 断言的是「本检查能看见它」，不是「我们记得它会泄漏」。
LEAK_MATRIX = [
    r"E=mc^{2}\eqno{3-1}", r"x\label{eq:one}", r"a=b\nonumber",
    r"x=y\hfill", r"U=IR\vspace{1em}", r"w=0.5\textwidth",
    r"a\zzz b", r"a=b\notag", r"x=\eqref{eq:one}",
    r"\intertext{and}x", r"\shortintertext{and}x", r"x\cline{1-2}",
    r"\multicolumn{2}{c}{x}", r"a\hline b", r"a\cr b", r"a\newline b",
    r"a\allowbreak b", r"a\pagebreak b", r"a\smallskip b",
    r"a\medskip b", r"a\bigskip b", r"a\vspace*{1em}b", r"a\qed",
    r"\settowidth{\x}{y}", r"\renewcommand{\x}{1}",
]


@needs_omml
@pytest.mark.parametrize("latex", LEAK_MATRIX)
def test_every_leak_shape_is_caught(latex, monkeypatch):
    r"""25 条泄漏逐条：转换产物里必有至少一段 ``m:t`` 满足新判据。

    这是「收紧判据没把覆盖弄丢」的证据 —— 判据从 loose 收到整段锚定，这 25 条
    必须一条不落地仍然被看见。

    **为什么要 monkeypatch 掉前置守卫**：P27-A 已在 ``math_omml`` 里加了层 1
    硬拒，这 25 条现在会直接抛 ``RuntimeError``、根本产不出 OMML。但本检查服务
    的对象正是**已经存在**的坏文档 —— 守卫上线之前生成的、别的工具生成的。
    所以这里绕过守卫去拿转换产物，看的是「当年那批文档里长什么样」。绕不过去
    就没有样本可断言，这条用例会变成同义反复。

    挂 ``needs_omml`` 是因为 CI runner 没有 Office，``latex_to_omml_lxml`` 会
    降级成纯文字（那时没有 ``m:t`` 可看，断言会假红）。
    """
    import docx_validate
    import math_omml

    # 两层守卫都要让开：层 1 拦在转换前，层 2 拦在产物出来之后。
    monkeypatch.setattr(math_omml, "assert_no_unsupported_commands",
                        lambda *_a, **_k: None)
    monkeypatch.setattr(math_omml, "assert_no_latex_residue",
                        lambda *_a, **_k: None)
    el = math_omml.latex_to_omml_lxml(latex)
    mts = [(t.text or "") for t in el.iter("{%s}t" % math_omml.M_NS)]
    hits = [t for t in mts if docx_validate.LATEX_RESIDUE_RE.match(t)]
    assert hits, f"{latex!r} 转换后没有 m:t 满足判据；m:t = {mts!r}"


@needs_omml
def test_legit_formulas_are_not_caught():
    r"""合法公式逐条不命中（含公式内 Windows 路径与转义反斜杠）。

    与上一条成对：只测泄漏的话，判据放宽成 ``.`` 也全绿。
    """
    import docx_validate
    import math_omml

    legit = [
        r"\frac{a}{b}", r"x^{2}", r"\max\{x,y\}\le z", r"\sin x",
        r"\alpha\beta", r"\mathrm{d}x", r"\mathcal{L}", r"\in",
        r"\rightarrow", r"\sqrt{2}", r"\pm x", r"\sin^{2} x",
        r"\text{C:\Users\me}", r"\text{a\\b}", r"\backslash",
        r"\text{见附录}", r"\text{see Fig. 3}", r"\int_0^1 x\,dx",
    ]
    for latex in legit:
        el = math_omml.latex_to_omml_lxml(latex)
        mts = [(t.text or "") for t in el.iter("{%s}t" % math_omml.M_NS)]
        hits = [t for t in mts if docx_validate.LATEX_RESIDUE_RE.match(t)]
        assert hits == [], f"{latex!r} 被误报成残留：{hits!r}"


# ===========================================================================
# 判据 8：文本模式内的残留**必须报**（旧版这里是「已知缺口」，已修）
# ===========================================================================
def test_text_mode_residue_is_reported(tmp_path: Path):
    r"""``m:t`` 文本是 ``a\xa0\\zzz\xa0b``（``\text{a \zzz b}`` 的产物）→ **必须报**。

    这是本判据存在的理由。命令落在 ``\text{}`` 里时，它与前后正文**共享一段**
    ``m:t>``，文本形如 ``a`` + NBSP + ``\zzz`` + NBSP + ``b``：

    - 「整段就是命令」看不见它（整段不是命令）；
    - 「文本里有反斜杠」能看见，但会把 ``C:\Users\me`` 一起误报。

    按空白切词后两者兼得：``split()`` 也切 NBSP，``\zzz`` 独占一个 token。
    """
    # NBSP 用 \xa0 显式写出：手写夹具要复现转换器的真实分隔符，不是普通空格。
    body = math_para("a\xa0\\zzz\xa0b")
    src = docx_with(tmp_path, "text_mode.docx", body)

    rep = validate_inproc(src)
    hits = residue(rep)
    assert len(hits) == 1, rep["issues"]
    assert hits[0]["count"] == 1
    assert hits[0]["samples"] == [r"\zzz"], hits[0]
    assert rep["ok"] is True


@needs_omml
def test_text_mode_residue_from_real_conversion():
    r"""真实转换链：``\text{a \zzz b}`` 的产物必须被判据看见。

    上一条用手写夹具钉判据，这一条钉「转换器真的这么产出」—— 手写的 NBSP 万一
    与真实分隔符不同（比如真实产物是多个 NBSP 或混用），手写那条会变成
    自说自话。绕过两层守卫的理由同判据 7。
    """
    import docx_validate
    import math_omml

    o1 = math_omml.assert_no_unsupported_commands
    o2 = math_omml.assert_no_latex_residue
    math_omml.assert_no_unsupported_commands = lambda *a, **k: None
    math_omml.assert_no_latex_residue = lambda *a, **k: None
    try:
        el = math_omml.latex_to_omml_lxml(r"\text{a \zzz b}")
    finally:
        math_omml.assert_no_unsupported_commands = o1
        math_omml.assert_no_latex_residue = o2

    mts = [(t.text or "") for t in el.iter("{%s}t" % math_omml.M_NS)]
    assert any("\\zzz" in t for t in mts), f"前提失效，m:t = {mts!r}"
    assert any(docx_validate.LATEX_RESIDUE_RE.match(tok)
               for t in mts for tok in t.split()), (
        f"真实产物没被看见：m:t = {mts!r}")


# ---- 两条已知边界（接受，钉住，不当 bug） ---------------------------------
def test_known_boundary_glued_backslash_is_not_reported(tmp_path: Path):
    r"""**已知边界**：``m:t`` = ``a\zzz\xa0b``（反斜杠紧贴前面的字母）→ 不报。

    ``\text{a\zzz b}`` 的产物：反斜杠前没有空白，``a\zzz`` 是**一个** token，
    整条不是控制序列 ⇒ 不命中。

    **接受这个漏**：要抓它就得回到「任意位置匹配」，那会让 ``C:\Users\me``、
    ``a\\b``、``见 \Users 目录`` 全部误报 —— 用一个系统性误报换一种极罕见的
    写法（正文里写 ``a\zzz`` 不加空格）。层 1 仍按命令名拒收它认识的命令。
    """
    src = docx_with(tmp_path, "glued.docx", math_para("a\\zzz\xa0b"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]


def test_known_boundary_punctuation_attached_is_not_reported(tmp_path: Path):
    r"""**已知边界**：``m:t`` = ``a\xa0\\zzz,\xa0b``（命令后紧跟逗号）→ 不报。

    ``\text{a \zzz, b}`` 的产物：``\zzz,`` 是**一个** token（标点没被切开）。
    同理 ``(\zzz)`` 与 ``\zzz{x}``（花括号形态无论如何都粘成 ``\zzz{x``）。

    **接受这个漏**。试过「按空白 + 标点切分」，实测无收益：

    - 修不了花括号形态 —— ``\zzz{x`` 怎么切都还带着 ``{x``；
    - 反而开始切路径内容，把 ``a\\b`` 这类本来干净的东西切开。

    所以判据停在「空白切词」，那正是转换器真实产出的分隔方式（文本模式用 NBSP
    分词）。要抓标点形态只能回到「任意位置匹配」，代价是路径全误报。
    """
    src = docx_with(tmp_path, "punct.docx", math_para("a\xa0\\zzz,\xa0b"))

    rep = validate_inproc(src)
    assert residue(rep) == [], rep["issues"]


def test_known_boundary_lone_path_word_is_reported(tmp_path: Path):
    r"""**已知边界**：``m:t`` = ``见\xa0\Users\xa0目录`` → **会报**（误报）。

    ``\text{见 \Users 目录}`` 的产物：``\Users`` 独占一个 token，形态上与残留
    命令**不可区分** —— 这是覆盖文本模式必然要付的价，换成「整段锚定」就
    看不见文本模式残留（本单第一版的结论）。

    对既有行为的影响要说清：``C:\Users\me`` 这类**带盘符/带反斜杠**的路径
    仍然干净（单 token、开头不是反斜杠）—— 新增的误报面只有「单独一个
    反斜杠单词」，不是所有路径。误报只是一条 warning，不改 ``ok`` 与退出码。
    """
    src = docx_with(tmp_path, "lone_path_word.docx",
                    math_para("见\xa0\\Users\xa0目录"))

    rep = validate_inproc(src)
    hits = residue(rep)
    assert len(hits) == 1, rep["issues"]
    assert hits[0]["severity"] == "warning", "即便误报也只是一条 warning"
    assert rep["ok"] is True

