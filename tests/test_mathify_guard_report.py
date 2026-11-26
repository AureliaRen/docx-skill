# -*- coding: utf-8 -*-
r"""``docx_mathify`` 的失败分类上报（P27-C）。

## 这个文件钉的是什么

``stats["failed"]`` 以前把两类完全不同的失败混成一个数字：

* **守卫拒收**（``kind="latex_guard"``）—— 源文本里写了 OMML 表达不了的控制
  序列（\\eqno / \\textwidth / 未知命令）。**用户该改源文本**，
  换机器、重装 Office、重跑都不会好。
* **转换器报错**（``kind="convert"``）—— 缺 ``MML2OMML.XSL`` / \\latex2mathml
  语法错 / XSLT 失败 / ``assert_legal`` 判结构非法。**环境或工具问题**。

两者都进 ``failed``（退出码 3 的判据不变，D-024 一个码都不加），但汇总行、
stderr、``--report`` 都要能分开说。否则用户看到「失败 3」不知道是改文档还是查环境。

## 分类判据为什么不用正则异常文案

文案是给人读的，改一次措辞就会把分类打歪。本单改用 ``stats["latex_guard_hits"]``
（P27-A 建的）：调用前快照长度、异常后比较，变长了就是本次命中守卫。
``test_classification_does_not_depend_on_error_wording`` 专门钉这一点 —— 把异常
文案整段换掉，分类必须不变。

## 三条硬性质

1. **两类必须能分开**：源文本写错 → 全是 guard；XSL 坏 → 全是 convert；
   两者混合 → 各自计数正确。
2. **零失败时输出逐字不变**：汇总行在有失败时才追加「（拒收 N，转换 M）」，
   没有失败时那行与 P27-C 之前一字不差 —— 既有测试（`test_paper_chemistry`
   的 `失败：0`、`test_formula` 的「失败」子串）才不会被打断。
3. **老 dict 不炸**：没有 ``kind`` 键的失败记录按 ``convert`` 归类（「不认识的
   失败」按环境问题算更保守，至少不会把环境问题报成「你写错了」）。

Run:  python -m pytest tests/test_mathify_guard_report.py -q
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

import docx_mathify  # noqa: E402
from _omml import needs_omml  # noqa: E402

GUARD = docx_mathify.GUARD_KIND
CONVERT = docx_mathify.CONVERT_KIND

MISSING_XSL = str(SKILL / "work" / "p27c" / "no_such_dir" / "MML2OMML.XSL")


def mkdoc(path, formulas) -> Path:
    """造一份 docx：每条 formula 一个段落（$...$ 让 find_math_spans 能认出来）。"""
    doc = Document()
    for f in formulas:
        doc.add_paragraph(f)
    doc.save(str(path))
    return path


def run_mathify(*argv, env_extra=None):
    env = dict(os.environ)
    if env_extra:
        env.update(env_extra)
    return subprocess.run([sys.executable, "-X", "utf8", str(SCRIPTS / "docx_mathify.py"),
                           *[str(a) for a in argv]],
                          capture_output=True, text=True, cwd=str(SKILL), env=env,
                          encoding="utf-8", errors="replace")


def summary_line(stdout: str) -> str:
    """汇总行：以「段落 」开头的那一行。"""
    for ln in (stdout or "").splitlines():
        if ln.startswith("段落 "):
            return ln
    raise AssertionError("stdout 里没有汇总行：%r" % (stdout[:400],))


def fail_lines(stdout: str):
    return [ln.strip() for ln in (stdout or "").splitlines()
            if ln.strip().startswith("FAIL[")]


# ===========================================================================
# C-1 分类：unit 级，直接调内部函数（不依赖 Office / XSL）
# ===========================================================================
def test_guard_kind_when_hits_grew():
    r"""调用期间 \\latex_guard_hits 长了一条 → 判成 guard。"""
    stats = {"failed": [], "latex_guard_hits": []}
    before = docx_mathify._guard_marks(stats)
    stats["latex_guard_hits"].append({"latex": "x", "layer": "pre",
                                     "commands": [chr(92) + "eqno"]})
    assert docx_mathify._classify_failure(stats, before) == GUARD


def test_convert_kind_when_hits_unchanged():
    r"""调用期间命中数没变 → 判成 convert（哪怕 hits 里本来就有旧记录）。"""
    stats = {"failed": [], "latex_guard_hits": [
        {"latex": "old", "layer": "pre", "commands": ["x"]}]}
    before = docx_mathify._guard_marks(stats)
    assert docx_mathify._classify_failure(stats, before) == CONVERT


def test_guard_marks_tolerates_missing_key():
    """没有 latex_guard_hits 键（老 stats / 手搓 dict）时按 0 算，不 KeyError。"""
    assert docx_mathify._guard_marks({}) == 0
    assert docx_mathify._classify_failure({}, 0) == CONVERT


def test_failed_by_kind_splits_and_preserves_order():
    """两类各自保序；老记录（无 kind）归 convert。"""
    stats = {"failed": [
        {"latex": "g1", "error": "e", "kind": GUARD},
        {"latex": "c1", "error": "e", "kind": CONVERT},
        {"latex": "g2", "error": "e", "kind": GUARD},
        {"latex": "legacy", "error": "e"},          # 没有 kind → convert
    ]}
    guard, convert = docx_mathify._failed_by_kind(stats)
    assert [f["latex"] for f in guard] == ["g1", "g2"]
    assert [f["latex"] for f in convert] == ["c1", "legacy"]


def test_failed_by_kind_on_empty_and_missing():
    assert docx_mathify._failed_by_kind({"failed": []}) == ([], [])
    assert docx_mathify._failed_by_kind({}) == ([], [])


# ===========================================================================
# C-1 分类：从真实异常走一遍（守卫生效时，_om_or_text 必须记 guard）
# ===========================================================================
@needs_omml
def test_om_or_text_records_guard_kind_on_rejection():
    r"""走 \\eqno：\\stats["failed"] 里那条的 kind 必须是 guard。"""
    stats = {"failed": [], "latex_guard_hits": []}
    runs = docx_mathify._om_or_text(r"E=mc^2 \\eqno{(3-1)}", "FALLBACK", stats)
    assert runs, "守卫拒收应回退成文字 run，不该凭空消失"
    assert len(stats["failed"]) == 1
    rec = stats["failed"][0]
    assert rec["kind"] == GUARD, rec
    assert "OMML 表达不了" in rec["error"]


@needs_omml
def test_om_or_text_records_convert_kind_on_syntax_error():
    r"""走真语法错（\\frac 缺参数）：kind 必须是 convert。"""
    stats = {"failed": [], "latex_guard_hits": []}
    docx_mathify._om_or_text(chr(92) + "frac", "FALLBACK", stats)
    assert len(stats["failed"]) == 1
    rec = stats["failed"][0]
    assert rec["kind"] == CONVERT, rec
    assert stats.get("latex_guard_hits") == [], "转换器报错不该伪造守卫命中"


def test_om_or_text_records_convert_kind_on_missing_xsl(monkeypatch):
    """缺 XSL（环境问题）也必须是 convert，绝不能被算成 guard。"""
    stats = {"failed": [], "latex_guard_hits": []}
    monkeypatch.setenv("MML2OMML_XSL", MISSING_XSL)
    import math_omml
    monkeypatch.setattr(math_omml, "_XSLT", None)   # 破掉缓存，强制重新探测
    docx_mathify._om_or_text(chr(92) + "frac{a}{b}", "FALLBACK", stats)
    assert len(stats["failed"]) == 1
    assert stats["failed"][0]["kind"] == CONVERT, stats["failed"]


@needs_omml
def test_classification_does_not_depend_on_error_wording(monkeypatch):
    r"""把守卫文案整段换掉，分类必须不变 —— 证明判据不是正则异常文案。

    这是本单最容易腐化的一点：哪天有人图省事改回「error 里含某某子串就是
    guard」，文案一改分类就歪，而这种歪法**不会让任何别的测试变红**。
    """
    import math_omml

    def fake_guard(latex, stats=None):
        if stats is not None:
            stats["latex_guard"] = stats.get("latex_guard", 0) + 1
            stats.setdefault("latex_guard_hits", []).append(
                {"latex": latex, "layer": "pre", "commands": [chr(92) + "eqno"]})
        raise RuntimeError("完全不同的一句话，且不含任何可被正则匹配的关键词")

    monkeypatch.setattr(math_omml, "assert_no_unsupported_commands", fake_guard)
    stats = {"failed": [], "latex_guard_hits": []}
    docx_mathify._om_or_text("x", "FALLBACK", stats)
    assert stats["failed"][0]["kind"] == GUARD
    assert "OMML 表达不了" not in stats["failed"][0]["error"], "文案已被换掉"
    assert "残留了 LaTeX 控制序列" not in stats["failed"][0]["error"]


# ===========================================================================
# C-2 汇总行 / C-3 stderr：端到端两类各一份 + 混合一份
# ===========================================================================
@needs_omml
def test_cli_guard_only_reports_rejections(tmp_path):
    r"""源文本含 \\eqno / \\nonumber / \\textwidth：
    「失败 3（拒收 3，转换 0）」+ 三条 FAIL[guard] + exit 3 + 回读 0 个反斜杠 m:t。
    """
    src = mkdoc(tmp_path / "guard.docx", [
        r"$E=mc^2 \\eqno{(3-1)}$",
        r"$a=b \\nonumber$",
        r"$w=0.5\\textwidth$",
    ])
    out = tmp_path / "out.docx"
    proc = run_mathify(src, "-o", out)

    assert proc.returncode == 3, (proc.returncode, proc.stderr[:300])
    line = summary_line(proc.stdout)
    assert "失败 3（拒收 3，转换 0）" in line, line

    fails = fail_lines(proc.stdout)
    assert len(fails) == 3, fails
    assert all(f.startswith("FAIL[guard]") for f in fails), fails
    assert not any("FAIL[convert]" in f for f in fails)

    # 退出码与 stderr 的那行总数不变（D-024 一个码都没加）
    assert "3 formula(s) failed to convert" in proc.stderr

    # 产物里一个反斜杠 m:t 都不许有（A 单的效果，这里确认没被 C 单弄回去）
    import zipfile
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    import re as _re
    bad = [t for t in _re.findall(r"<m:t[^>]*>(.*?)</m:t>", xml) if chr(92) in t]
    assert bad == [], bad


@needs_omml
def test_cli_convert_only_reports_conversion_failures(tmp_path):
    """XSL 指向不存在的路径 → 「拒收 0，转换 N」+ 全是 FAIL[convert] + exit 3。"""
    src = mkdoc(tmp_path / "conv.docx", [
        r"$x=\\frac{a}{b}$",
        r"$y=\\sqrt{x}$",
    ])
    out = tmp_path / "out.docx"
    proc = run_mathify(src, "-o", out,
                       env_extra={"MML2OMML_XSL": MISSING_XSL})

    assert proc.returncode == 3, (proc.returncode, proc.stdout[-300:], proc.stderr[:300])
    line = summary_line(proc.stdout)
    assert "拒收 0，转换 2" in line, line

    fails = fail_lines(proc.stdout)
    assert len(fails) == 2, fails
    assert all(f.startswith("FAIL[convert]") for f in fails), fails
    assert not any("FAIL[guard]" in f for f in fails)
    assert "MML2OMML" in proc.stdout


@needs_omml
def test_cli_mixed_splits_the_two_counts(tmp_path):
    """一条守卫拒收 + 一条真结构公式（XSL 坏）→ 「拒收 1，转换 1」，两类各一行。"""
    src = mkdoc(tmp_path / "mix.docx", [
        r"$E=mc^2 \\eqno{(3-1)}$",
        r"$x=\\frac{a}{b}$",
    ])
    out = tmp_path / "out.docx"
    proc = run_mathify(src, "-o", out,
                       env_extra={"MML2OMML_XSL": MISSING_XSL})

    assert proc.returncode == 3, (proc.returncode, proc.stdout[-300:])
    line = summary_line(proc.stdout)
    assert "失败 2（拒收 1，转换 1）" in line, line

    fails = fail_lines(proc.stdout)
    assert len(fails) == 2, fails
    assert sum(f.startswith("FAIL[guard]") for f in fails) == 1, fails
    assert sum(f.startswith("FAIL[convert]") for f in fails) == 1, fails


# ===========================================================================
# 零失败时输出**逐字不变**（既有测试依赖这个）
# ===========================================================================
@needs_omml
def test_cli_zero_failure_keeps_summary_unchanged(tmp_path):
    """全成功时汇总行**不含**「（拒收」—— P27-C 不许给常态输出加噪音。

    `test_paper_chemistry` 的 `失败：0`、`test_formula` 的「失败」子串都靠这行；
    给它加一个恒存在的括号会让所有读这行的既有测试和用户脚本一起变味。
    """
    src = mkdoc(tmp_path / "ok.docx", [r"$x=\\frac{a}{b}$"])
    out = tmp_path / "out.docx"
    proc = run_mathify(src, "-o", out)

    assert proc.returncode == 0, (proc.returncode, proc.stdout[-300:], proc.stderr[:300])
    line = summary_line(proc.stdout)
    assert "失败 0" in line, line
    assert "拒收" not in line, line
    assert "转换 " not in line, line
    assert fail_lines(proc.stdout) == []


@needs_omml
def test_cli_truncates_each_kind_to_ten(tmp_path):
    """两类各自限 10 条（沿用既有截断）：12 条 guard → 只打 10 条 FAIL[guard]。

    11 条而不是 12 条是为了让「多出来的那条」在截断边缘上 —— 恰好 10 条时
    截断逻辑没被真正踩到。
    """
    forms = [r"$a%d=b \\nonumber$" % i for i in range(11)]
    src = mkdoc(tmp_path / "many.docx", forms)
    out = tmp_path / "out.docx"
    proc = run_mathify(src, "-o", out)

    assert proc.returncode == 3
    assert "拒收 11" in summary_line(proc.stdout)
    assert len(fail_lines(proc.stdout)) == 10, len(fail_lines(proc.stdout))


# ===========================================================================
# --report 也分两节（同一套判据，别只在 stdout 上分）
# ===========================================================================
@needs_omml
def test_report_splits_two_sections(tmp_path):
    src = mkdoc(tmp_path / "rep.docx", [
        r"$E=mc^2 \\eqno{(3-1)}$",
        r"$x=\\frac{a}{b}$",
    ])
    out = tmp_path / "out.docx"
    rep = tmp_path / "report.md"
    proc = run_mathify(src, "-o", out, "--report", rep,
                       env_extra={"MML2OMML_XSL": MISSING_XSL})
    assert proc.returncode == 3, (proc.returncode, proc.stdout[-300:])

    text = rep.read_text(encoding="utf-8")
    assert "拒收 1" in text and "转换 1" in text, text[:400]
    assert "守卫拒收" in text, "报告里没有守卫那一节"
    assert "转换器报错" in text, "报告里没有转换器那一节"


@needs_omml
def test_report_clean_has_no_failure_section(tmp_path):
    """零失败时报告里连「转换失败」这个词都不出现（不给常态报告挂噪音）。"""
    src = mkdoc(tmp_path / "clean.docx", [r"$x=\\frac{a}{b}$"])
    rep = tmp_path / "report.md"
    proc = run_mathify(src, "--dry-run", "--report", rep)
    assert proc.returncode == 0, (proc.returncode, proc.stderr[:300])
    text = rep.read_text(encoding="utf-8")
    assert "失败：0" in text
    assert "转换失败" not in text
