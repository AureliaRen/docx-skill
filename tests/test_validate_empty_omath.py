# -*- coding: utf-8 -*-
r"""``docx_validate`` 的 **整条公式空着** 检查（P29-B）。

## 本文件存在的理由

``latex2mathml`` 把 ``\rule{10pt}{2pt}`` / ``\hspace{1em}`` 这一族映射成
``<mspace>``，而 ``MML2OMML.XSL`` 没有 ``mspace`` 的模板 —— 节点被丢掉，整条
公式只剩一个**零子元素**的 ``<m:oMath/>``。Word 打开这种文档，在公式该出现
的位置什么都不画（或画个空占位）。空组 ``{}``、定义了却没用的 ``\def`` 同理。

体检侧此前看不见它：相邻两条 OMML 检查只看 ``<m:d>`` 空槽与 ``<m:t>`` 文本残留，
都不问「这条公式是不是空的」。实测 ``work/p28c/docx/out-rule-bare.docx``
（P28-C 探针留下的现成包）拿到 ``{"ok": true, "issues": []}`` + exit 0。

## 与相邻两条检查的分工（三条互不吞并）

- ``omml_empty_slot``：**定界组里空一个槽**（``<m:d>`` 里一个 ``<m:e/>`` 空着）。
  公式本身有内容，只是其中一格画成空方块。
- ``omml_empty_omath``（本文件）：**整条公式空着**（``<m:oMath>`` 零子元素）。
  两者不是同一件事：前者是「里面坏一格」，后者是「整条没有了」。
- ``omml_latex_residue``：公式里印着反斜杠命令名（命令变成了字面量文字）。

转换侧（``math_omml``，P29-A）在公式成为 OMML **之前**拒收这些控制序列 —— 那
治因；本检查服务的是**已经存在**的文档：旧版 skill 产的、别的工具产的、或者
拒收上线之前的产物。两者互补，语义不同。

## 四条不容松动的性质

1. **只报不拒收**：severity 是 ``warning``，``ok`` 仍 ``True``、CLI 仍 exit 0。
   这份报告喂给下游做 pass/fail，Word 能打开这类文档（只是画不出来），体检不许
   把能开的文件判死 —— 与 ``omml_empty_slot`` 同一条理由。
2. **``count`` 是真数出来的**：三条空公式就得是 3，不是常量 1。调用方拿它算
   「有几条空公式」，从详情散文里取不出来。
3. **干净文档不出键**：有正常公式、没有空 ``m:oMath`` 的文档里连这个字符串都
   不该出现 —— 不是空列表、不是 ``count: 0`` 常驻噪音。
4. **与 ``omml_empty_slot`` 互不吞并**：同一份文档里两种毛病同时存在时，两条
   issue 各自独立出现。两条扫描紧邻，最容易犯的错是提前 ``return`` 或共用变量。

## 为什么新建文件，而不是追加进 ``tests/test_validate_omml_residue.py``

那个文件的 docstring 把「三条不许松动的性质」全部写在 ``omml_latex_residue``
语境里，夹具与判据都围绕 ``<m:t>`` 文本建。把第二个 code 塞进去会让「一个文件
钉一个判据」的读法失效，也要改造 ``residue()`` 这类按 code 取 issue 的辅助函数。
本检查的夹具形态（空元素 vs 文本）与判据（``len(el) == 0`` vs 正则）都不同，
独立成文件更省事。**造包方式照抄**那个文件的 ``docx_with`` / ``rewrite_part`` /
``run_validate``（同一个 ``python-docx`` 造包 + ``zipfile`` 手术思路，不自己发明
一套）。

## 门控：**没有** ``needs_omml``

本文件的夹具一律手写 OMML，只跑 ``docx_validate``（纯 ``zipfile`` + ``lxml``），
不经 ``MML2OMML.XSL`` / Office COM / LibreOffice —— 没有 Office 的 CI runner
上照常 pass，所以**不需要** ``@needs_omml``，加了反而会在 runner 上白白 skip。

Run:  python -m pytest tests/test_validate_empty_omath.py -q
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
SCRIPT = SCRIPTS / "docx_validate.py"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

CODE = "omml_empty_omath"
SIBLING = "omml_empty_slot"


# ---- helpers (与 tests/test_validate_omml_residue.py 同一套造包手法) ---------

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

    Same construction as ``tests/test_validate_omml_residue.py``: string surgery
    on a real python-docx file, not ``OxmlElement``, because the point is to
    produce markup that the writer APIs never create — exactly what an older
    converter left behind.  ``xmlns:m`` is already declared on the root element
    by python-docx, so the injected ``m:`` prefix needs no declaration.
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


def empty_math_para() -> str:
    """self-closing ``<m:oMath/>`` (zero children) — what the chain leaves behind."""
    return "<w:p><m:oMath/></w:p>"


def live_math_para(*texts: str) -> str:
    """A formula with content: one ``<m:t>`` per *texts* entry (real converter shape)."""
    runs = "".join(f"<m:r><m:t>{t}</m:t></m:r>" for t in texts)
    return f"<w:p><m:oMath>{runs}</m:oMath></w:p>"


def empty_slot_para() -> str:
    """A formula that HAS content, with one blank slot in a delimiter group."""
    return ("<w:p><m:oMath><m:d><m:e/><m:e><m:r><m:t>x</m:t></m:r></m:e>"
            "</m:d></m:oMath></w:p>")


def hits(rep: dict) -> list:
    return [i for i in rep["issues"] if i["code"] == CODE]


def validate_inproc(path) -> dict:
    import docx_validate
    return docx_validate.validate(str(path))


# ===========================================================================
# 判据 1：命中 -> warning + count 1，ok 仍 True；CLI 仍 exit 0
# ===========================================================================
def test_empty_omath_is_reported_as_warning(tmp_path: Path):
    """一条空 ``m:oMath`` -> ``omml_empty_omath``，warning，count 1，ok True。

    一条用例钉四件事：键出现、severity 是 warning（不是 error）、count 数得对、
    ``ok`` 不动。只测「有没有这个键」会让一个判成 error 的实现也过。
    """
    src = docx_with(tmp_path, "one.docx", empty_math_para())

    rep = validate_inproc(src)
    found = hits(rep)
    assert len(found) == 1, rep["issues"]
    hit = found[0]
    assert hit["severity"] == "warning", "能打开的文件不许判 error"
    assert hit["count"] == 1
    assert "document.xml" in hit["detail"]
    assert "<m:oMath>" in hit["detail"]
    assert rep["ok"] is True, "公式空着不等于文档不合法"


def test_cli_exit_code_stays_zero(tmp_path: Path):
    """CLI 层：``ok:true`` + **exit 0** + stderr 空。

    在 ``validate()`` 里断言 ``ok is True`` 还不够 —— 被下游真正消费的是退出码
    （D-024 的 2/1/3 三档，这条检查一档都不许动）。
    """
    src = docx_with(tmp_path, "one_cli.docx", empty_math_para())

    proc = run_validate(src)
    assert proc.returncode == 0, (proc.returncode, proc.stderr, proc.stdout)
    assert proc.stderr == "", "warning 不该往 stderr 写东西"
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert [i["code"] for i in rep["issues"]] == [CODE]


# ===========================================================================
# 判据 2：count 是真数出来的，不是常量 1
# ===========================================================================
def test_three_empty_omaths_count_three(tmp_path: Path):
    """三条空公式 -> ``count == 3``。

    把三份坏产物合到一份文档里的形状。写死 1 的实现在这里红 —— 而「有几条空
    公式」正是调用方要的数。
    """
    src = docx_with(tmp_path, "three.docx", empty_math_para() * 3)

    found = hits(validate_inproc(src))
    assert len(found) == 1, "整篇只出一条 issue，不是一条公式一条"
    assert found[0]["count"] == 3


def test_count_ignores_whitespace_only_text(tmp_path: Path):
    """``<m:oMath>   </m:oMath>``（只有空白文本、零子元素）同样算空 -> count 1。

    判据是 ``len(el) == 0``（子**元素**个数），空白文本节点不是子元素。这条钉住
    判据用的是哪一个「空」，别改成 ``el.text`` 之类的判断。
    """
    src = docx_with(tmp_path, "ws.docx",
                    "<w:p><m:oMath>   </m:oMath></w:p>")

    found = hits(validate_inproc(src))
    assert len(found) == 1, found
    assert found[0]["count"] == 1


# ===========================================================================
# 判据 3：不误报（本文件最重要的一组）
# ===========================================================================
def test_normal_formula_is_clean(tmp_path: Path):
    """有正常公式（``m:r``/``m:t`` 内容）-> 完全不出这个 code。

    断言的是 ``CODE not in json.dumps(rep)``，不是「列表为空」：后者会放过
    ``count: 0`` 那种常驻噪音。一个「只要文档里有 ``m:`` 就报」的实现这里必红。
    """
    body = (live_math_para("x") + live_math_para("−")
            + live_math_para("max") + live_math_para("∂", "y"))
    src = docx_with(tmp_path, "clean_math.docx", body)

    rep = validate_inproc(src)
    assert hits(rep) == [], rep["issues"]
    assert CODE not in json.dumps(rep), "没有空公式的文档不许出现这个键"
    assert rep["ok"] is True


def test_plain_document_has_no_key(tmp_path: Path):
    """纯文字 docx（连 ``m:`` 都没有）-> 报告里连这个键都不出现。

    与上一条的区别在**前提**：这里根本没有公式，确认扫描器不会因为「找不到
    ``m:oMath``」而做奇怪的事。
    """
    src = docx_with(tmp_path, "plain.docx")

    rep = validate_inproc(src)
    assert rep["ok"] is True
    assert rep["issues"] == []
    assert CODE not in json.dumps(rep), "无空公式的文档不许出现这个键"


def test_empty_slot_alone_does_not_trigger_empty_omath(tmp_path: Path):
    """只有 ``omml_empty_slot``（公式有内容、一格空着）-> **不**出 ``omml_empty_omath``。

    这条是两个 code 必须分开的防线。把「``<m:d>`` 里空 ``<m:e/>``」也算成「整条
    公式空着」的实现，会把一条本来就有内容的公式报成两条毛病；反过来，只查
    ``m:d`` 的实现看不见 ``<m:oMath/>``（这就是本单要补的洞）。
    """
    src = docx_with(tmp_path, "slot_only.docx", empty_slot_para())

    rep = validate_inproc(src)
    codes = [i["code"] for i in rep["issues"]]
    assert SIBLING in codes, rep["issues"]
    assert hits(rep) == [], "有内容的公式不是空公式"
    assert rep["ok"] is True


# ===========================================================================
# 判据 4：与既有 omml_empty_slot 共存（互不吞并）
# ===========================================================================
def test_coexists_with_empty_slot(tmp_path: Path):
    """空槽 + 空公式同文 -> 两条 issue 各自独立出现（顺序不作要求）。

    两条扫描紧邻，最容易犯的错是提前 ``return`` 或共用变量把后一条吃掉。
    """
    body = empty_slot_para() + empty_math_para()
    src = docx_with(tmp_path, "both.docx", body)

    rep = validate_inproc(src)
    codes = [i["code"] for i in rep["issues"]]
    assert SIBLING in codes, rep["issues"]
    assert CODE in codes, rep["issues"]
    assert rep["ok"] is True, "两条都是 warning，不该判死"
    by_code = {i["code"]: i for i in rep["issues"]}
    assert by_code[SIBLING]["count"] == 1, by_code[SIBLING]
    assert by_code[CODE]["count"] == 1, by_code[CODE]


# ===========================================================================
# 判据 5：``m:oMathPara`` 包裹时**不漏检**（已知边界里最要紧的一条）
# ===========================================================================
def test_omathpara_wrapped_empty_omath_is_still_found(tmp_path: Path):
    """``<m:oMathPara><m:oMath/></m:oMathPara>`` -> 仍然命中。

    转换为显示公式时常用 ``m:oMathPara`` 包一层（P28-C 的探针三份都是这个形状，
    实测 ``out-rule-bare.docx`` 里 ``oMathPara=1`` 而 ``oMath`` 子元素 0）。
    ``iter`` 是递归的，所以包一层不影响；这条把它钉住，防止有人把扫描改成
    ``doc_root.findall(...)``（只查直接子元素）后静默漏掉全部显示公式。
    """
    body = "<w:p><m:oMathPara><m:oMath/></m:oMathPara></w:p>"
    src = docx_with(tmp_path, "para_wrapped.docx", body)

    found = hits(validate_inproc(src))
    assert len(found) == 1, found
    assert found[0]["count"] == 1


# ===========================================================================
# 判据 6：接线 —— 新检查必须真的在 validate() 里跑到
# ===========================================================================
def test_scan_is_reachable_from_validate():
    """``validate()`` 的源码里必须有这次扫描（防「写了函数没接上」）。

    严格说冗余（所有用例都走 ``validate()``），留着是因为实现刻意**不抽成独立
    函数**（就地扫，与相邻的 ``omml_empty_slot`` 同形）：一旦有人把它抽走却忘了
    调用，上面那些「不出键」的用例会全部变绿 —— 没有输入就没有断言。
    """
    import inspect

    import docx_validate

    src = inspect.getsource(docx_validate.validate)
    assert CODE in src, "扫描没接进 validate()"

