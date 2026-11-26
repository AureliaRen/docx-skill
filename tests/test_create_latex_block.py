# -*- coding: utf-8 -*-
"""``docx_create`` 的 ``{"type": "latex"}`` block：从 spec 写含公式文档的那条路。

使用汇报缺陷 2（lead 已核实：``grep latex/omml`` 对 ``docx_create.py`` 0 命中）：
一份满是公式的文档原本只能自己写 python-docx 塞 OMML。现在它是 spec 能力：

    {"type": "latex", "latex": "E = mc^2"}                    行内（m:oMath）
    {"type": "latex", "latex": "\\frac{a}{b} = c", "display": true}  独立公式行（m:oMathPara）

本文件全部走**真 CLI**（subprocess），因为它是用户实际调用的那个面：
白名单校验、退出码、stdout JSON、写盘与否，都只有在子进程里才测得准。

退出码口径（D-024）：``2`` = spec 被拒（**零写盘**）、``3`` = 文档写了但至少一条
公式没转成、``0`` = 全好。

⚠️ 坏 LaTeX 的样本选择有讲究：``r"\\frac{a"`` 看着像坏括号，实际 ``latex2mathml``
容忍它并产出正常 OMML（实测 exit 0）—— 拿它当坏样本会得到一条永远绿的测试。
真正会抛的是 ``r"\\frac"`` / ``r"\\sqrt"`` / ``r"\\frac{a}{"``。
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

SKILL = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL / "scripts"
CREATE = SCRIPTS / "docx_create.py"
VALIDATE = SCRIPTS / "docx_validate.py"
READ = SCRIPTS / "docx_read.py"
sys.path.insert(0, str(SCRIPTS))

import docx_create as dc  # noqa: E402
from _omml import needs_omml  # noqa: E402

# ``latex2mathml`` 一定会抛的两个真实坏样本（见模块 docstring 的警告）。
BROKEN_LATEX = "\\frac"


# ---------------------------------------------------------------- helpers
def _json_of(proc):
    """stdout 里的 JSON：整段先试（``docx_read`` 是缩进多行），再退回最后一行。"""
    text = (proc.stdout or "").strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    payload = None
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped.startswith("{"):
            continue
        try:
            payload = json.loads(stripped)
        except json.JSONDecodeError:
            continue
    return payload


def _run(script: Path, *args):
    """跑一个 CLI，返回 (proc, stdout 里的 JSON)。"""
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script)] + [str(a) for a in args],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SKILL))
    return proc, _json_of(proc)


def run_create(tmp_path: Path, blocks: list, name: str = "spec.json"):
    spec_path = tmp_path / name
    spec_path.write_text(json.dumps({"blocks": blocks}, ensure_ascii=False),
                         encoding="utf-8")
    out = tmp_path / (name.replace(".json", "") + ".docx")
    proc, payload = _run(CREATE, spec_path, out)
    return proc, payload, out


def document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8")


def omml_counts(path: Path) -> dict:
    """按元素数，不按字符串数 —— ``<m:oMath>`` 是 ``<m:oMathPara>`` 的前缀。"""
    root = etree.fromstring(document_xml(path).encode("utf-8"))

    def n(local: str) -> int:
        return sum(1 for e in root.iter()
                   if isinstance(e.tag, str) and e.tag.rsplit("}", 1)[-1] == local)

    return {"oMath": n("oMath"), "oMathPara": n("oMathPara"),
            "acc": n("acc"), "limUpp": n("limUpp")}


# ---------------------------------------------------------------- 白名单登记
def test_the_block_type_is_registered_in_the_whitelist():
    """★ ``validate_spec`` 会拒掉没登记的类型 —— 这是本轮最容易漏的点。"""
    assert "latex" in dc.BLOCK_KEYS
    assert {"latex", "display"} <= dc.BLOCK_KEYS["latex"]
    # 段落级格式键也要能用（公式段也是段落）。
    assert set(dc.PARA_FORMAT_KEYS) <= dc.BLOCK_KEYS["latex"]


def test_the_module_docstring_documents_the_block():
    """spec 示例是文件内文档，用户第一眼看的也是它。"""
    doc = dc.__doc__ or ""
    assert '"type": "latex"' in doc
    assert "formula_failures" in doc


def test_validate_spec_accepts_a_minimal_latex_block():
    """不带 ``display`` 也要过校验（默认行内）—— 不需要 Office。"""
    dc.validate_spec({"blocks": [{"type": "latex", "latex": "E = mc^2"}]})


@pytest.mark.parametrize("block,needle", [
    ({"type": "latex"}, '"latex"'),
    ({"type": "latex", "latex": ""}, "empty"),
    ({"type": "latex", "latex": "   "}, "empty"),
    ({"type": "latex", "latex": 3}, "must be a string"),
    ({"type": "latex", "latex": "x", "display": "yes"}, '"display"'),
])
def test_a_broken_latex_block_is_rejected_before_anything_is_written(
        tmp_path: Path, block, needle):
    """spec 错误 → exit 2、零写盘、stderr 恰一行 ``error:``（D-030）。"""
    proc, payload, out = run_create(tmp_path, [block])
    assert proc.returncode == 2, proc.stderr
    assert payload is not None and payload["ok"] is False
    assert payload["output"] is None
    assert not out.exists(), "被拒的 spec 不许留下任何文件"
    lines = [ln for ln in (proc.stderr or "").splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    assert lines[0].startswith("error: "), lines[0]
    assert "Traceback" not in proc.stderr
    assert needle in payload["error"], payload["error"]


def test_a_misspelt_latex_key_gets_a_suggestion(tmp_path: Path):
    """``latx`` 必须被点名并给出最接近的合法键，而不是静默丢内容。"""
    proc, payload, out = run_create(tmp_path, [{"type": "latex", "latx": "x"}])
    assert proc.returncode == 2
    assert 'did you mean "latex"' in payload["error"], payload["error"]
    assert not out.exists()


# ---------------------------------------------------------------- 真转出 OMML
@needs_omml
def test_three_equations_become_three_omml_objects(tmp_path: Path):
    """行内 → ``m:oMath``，``display: true`` → ``m:oMathPara``（包一个 oMath）。"""
    proc, payload, out = run_create(tmp_path, [
        {"type": "heading", "text": "公式", "level": 1},
        {"type": "latex", "latex": "E = mc^2"},
        {"type": "latex", "latex": "\\frac{a}{b} = c"},
        {"type": "latex", "latex": "\\int_0^1 x^2\\,dx = \\frac{1}{3}", "display": True},
    ])
    assert proc.returncode == 0, proc.stderr
    assert payload["ok"] is True
    assert payload["formula_failures"] == []
    assert out.exists()

    counts = omml_counts(out)
    assert counts["oMath"] == 3, counts
    assert counts["oMathPara"] == 1, counts


@needs_omml
def test_the_phasor_dot_is_already_repaired_in_the_document(tmp_path: Path):
    r"""端到端：``\dot{I}`` 落进 docx 时必须是 accent，不是 limUpp。

    这条锁的是「``docx_create`` 走 ``latex_to_omml_safe``」—— 如果哪天有人把它
    改回裸调 ``latex_to_omml_lxml``，``limUpp`` 会重新出现，这条立刻红。
    """
    proc, _, out = run_create(tmp_path, [
        {"type": "latex", "latex": "\\dot{I} = \\frac{U}{Z}"},
    ])
    assert proc.returncode == 0, proc.stderr
    counts = omml_counts(out)
    assert counts["limUpp"] == 0, counts
    assert counts["acc"] >= 1, counts


@needs_omml
def test_an_inline_equation_gets_the_zero_width_pin(tmp_path: Path):
    """行内公式独占整段时 Word 会把它提升成行间公式；U+200B 压住它。

    与 ``insert_formula`` 规则 2、``docx_mathify`` 的单元格分支同一招：
    ``display: false`` 带零宽空格，``display: true`` 不带（它本来就该独立成行）。
    """
    proc, _, inline_out = run_create(tmp_path, [
        {"type": "latex", "latex": "E = mc^2"},
    ], name="inline.json")
    assert proc.returncode == 0, proc.stderr
    assert "\u200b" in document_xml(inline_out)

    proc, _, display_out = run_create(tmp_path, [
        {"type": "latex", "latex": "E = mc^2", "display": True},
    ], name="display.json")
    assert proc.returncode == 0, proc.stderr
    assert "\u200b" not in document_xml(display_out)


@needs_omml
def test_a_latex_block_takes_the_paragraph_format_keys(tmp_path: Path):
    """公式段也是段落：``alignment`` / ``space_before_pt`` / ``bookmark`` 照旧生效。"""
    proc, payload, out = run_create(tmp_path, [
        {"type": "latex", "latex": "E = mc^2", "alignment": "center",
         "space_before_pt": 6, "bookmark": "eq_book"},
    ])
    assert proc.returncode == 0, proc.stderr
    xml = document_xml(out)
    assert 'w:val="center"' in xml
    assert 'w:name="eq_book"' in xml


# ---------------------------------------------------------------- 部分失败
@needs_omml
def test_a_broken_equation_exits_three_and_the_document_is_still_written(
        tmp_path: Path):
    """坏公式不拖垮整篇：文件照写、失败入册、退出码 3（D-024）。"""
    proc, payload, out = run_create(tmp_path, [
        {"type": "latex", "latex": "E = mc^2"},
        {"type": "latex", "latex": "\\frac{a}{b} = c"},
        {"type": "latex", "latex": "\\sqrt{2}"},
        {"type": "latex", "latex": BROKEN_LATEX},
        {"type": "paragraph", "text": "正文结束"},
    ])
    assert proc.returncode == 3, (proc.returncode, proc.stderr)
    assert out.exists(), "退出码 3 的语义是「有产物」"
    assert payload["ok"] is False
    assert payload["output"] == str(out)

    failures = payload["formula_failures"]
    assert len(failures) == 1, failures
    assert failures[0]["block"] == 3
    assert failures[0]["latex"] == BROKEN_LATEX
    err = failures[0]["error"]
    assert err.startswith("LatexSyntaxError:"), err

    # 好的三条真的成了 OMML，坏的那条以原文留在段落里。
    counts = omml_counts(out)
    assert counts["oMath"] == 3, counts
    assert BROKEN_LATEX in document_xml(out)

    # stderr 一行 warning（不带 ``error:`` 前缀 —— 那是 exit 2 的专属记号）。
    warnings = [ln for ln in (proc.stderr or "").splitlines()
                if ln.strip().startswith("warning:")]
    assert len(warnings) == 1, proc.stderr
    assert "did not convert" in warnings[0]
    assert "error:" not in proc.stderr


@needs_omml
def test_the_blocks_after_a_broken_equation_are_still_built(tmp_path: Path):
    """失败点之后的块必须照常落地 —— 否则「不中断整篇」是假的。"""
    proc, _, out = run_create(tmp_path, [
        {"type": "latex", "latex": BROKEN_LATEX},
        {"type": "paragraph", "text": "坏公式之后"},
    ])
    assert proc.returncode == 3, proc.stderr
    read_proc, blocks = _run(READ, out, "--blocks")
    assert read_proc.returncode == 0, read_proc.stderr
    texts = [b.get("text", "") for b in blocks["blocks"]]
    assert "坏公式之后" in texts
    assert BROKEN_LATEX in texts, "失败的那条要以 LaTeX 原文留在文档里"


@needs_omml
def test_the_exit_three_product_passes_docx_validate(tmp_path: Path):
    """带一条没转成的公式，包本身仍然要是健康文档（``ok: true``）。"""
    proc, _, out = run_create(tmp_path, [
        {"type": "latex", "latex": "E = mc^2"},
        {"type": "latex", "latex": BROKEN_LATEX},
    ])
    assert proc.returncode == 3, proc.stderr
    val_proc, payload = _run(VALIDATE, out)
    assert val_proc.returncode == 0, val_proc.stdout + val_proc.stderr
    assert payload["ok"] is True
    assert payload["issues"] == []

