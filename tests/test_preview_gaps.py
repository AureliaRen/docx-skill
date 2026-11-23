# -*- coding: utf-8 -*-
r"""``docx_preview.py --gaps`` — 页尾空白的**绝对长度**度量。

## 为什么要有这个开关

主人原话：「没有页尾空白/孤行度量 — P4 空 8.3cm、P13 只有半页，都得用 fitz 量
『正文最低墨迹 y』。」

`--lint` 已经有 `fill_ratio`（**比例**）和 `sparse_pages`，缺的是**绝对长度**：
「这页空了 8.3 厘米」是尺子能量出来的话，而 `fill_ratio 0.204` 不是。也不能拿
比例反推 —— 版心高度跟着纸张和页边距走，同一份 0.204 在 A4 和 16 开上是两个
不同的厘米数。

## 本文件的三层（与 ``test_docx_preview.py`` 同构）

1. **纯函数算术**（无 PyMuPDF、无 LibreOffice，**CI 上真的会跑**）：
   `_tail_gap_mm` / `_head_gap_mm` / `_text_bottom_mm` / `_first_line` 全部
   喂合成 block 字典。页尾空白是一道减法，把它绑到 LibreOffice 的行距上，就会
   在真正需要它的 runner 上被 skip 掉。
2. **CLI 契约**（无渲染）：`--gaps` 单独用不许被「至少一个能力」校验打回；
   `--max-tail-gap` 负值是 exit 2。
3. **端到端**（`@needs_render`）：真 soffice → 真 PDF，证明 `gap_page` 会在
   **确实半空的页**上命中、在满页上零命中。

## 不变式（第二条隐藏轴）

`--lint` 的 schema **不许**因为本轮变宽：`--gaps` 的所有新键在没给
`--gaps` 时必须**完全不出现**。已有消费者按 key 存在性判断，多一个
`gap_pages: []` 会让「没查」与「查了没发现」混成同一个读法。
"""
from __future__ import annotations

import ast
import json
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docx_preview as dp  # noqa: E402

SCRIPT = SCRIPTS / "docx_preview.py"
MODULE_SRC = SCRIPT.read_text(encoding="utf-8")

# A4: 210 x 297 mm -> 595.28 x 841.89 pt.  Margins 25.4 mm -> 72 pt.
A4_HEIGHT_PT = 841.89
A4_BODY_BOTTOM_PT = A4_HEIGHT_PT - 72.0     # 769.89 pt
A4_BODY_TOP_PT = 72.0
MM = dp.PT_PER_MM


# --------------------------------------------------------------------- helpers
def run_cli(argv, env=None):
    import os
    import subprocess
    full_env = None
    if env is not None:
        full_env = dict(os.environ)
        full_env["PYTHONIOENCODING"] = "utf-8"
        for key, value in env.items():
            if value is None:
                full_env.pop(key, None)
            else:
                full_env[key] = value
    return subprocess.run([sys.executable, str(SCRIPT), *argv],
                          capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=full_env)


def block(y0, y1, text="x", x0=0.0, x1=400.0, btype=0):
    """One ``get_text("blocks")``-shaped dict, as ``_body_blocks`` builds them.

    ``btype=1`` is an image block — it carries no text but is still ink, which is
    the distinction :func:`docx_preview._text_bottom_mm` must honour.
    """
    return {"x0": x0, "y0": y0, "x1": x1, "y1": y1, "text": text, "type": btype}


def stderr_lines(proc):
    return [ln for ln in proc.stderr.replace("\r", "").split("\n") if ln]


def text_ending_at_mm(mm_from_top: float):
    """A single block whose bottom edge sits *mm_from_top* mm below the page top."""
    y1 = mm_from_top * MM
    return [block(A4_BODY_TOP_PT, y1, "正文")]


# =========================================================== 1. pure arithmetic
class TestTailGapArithmetic:
    """No PyMuPDF, no LibreOffice — this class runs on every runner."""

    def test_full_page_has_a_small_tail(self):
        """★ 满页的 tail_gap 是个位数毫米 —— 主人手量的「空 8.3cm」的反面。

        A4 版心底 769.89 pt；墨迹停在 760 pt（约 268 mm）时应当只剩 ~10 pt
        ≈ 3.5 mm，而不是几百毫米。这条同时钉住**符号**：早期版本把
        ``body_bottom_pt / PT_PER_MM``（毫米）与 ``_text_bottom_mm``（毫米）
        混着减，满页报了 ``-170 mm``。
        """
        blocks = [block(A4_BODY_TOP_PT, 760.0, "满页最后一行")]
        gap = dp._tail_gap_mm(blocks, A4_BODY_BOTTOM_PT)
        assert gap == pytest.approx((769.89 - 760.0) / MM, abs=0.02)
        assert 0 < gap < 5, gap

    def test_the_owners_83mm_case(self):
        """★ 主人实测「P4 空 8.3cm」的量级：tail_gap ≈ 83 mm。

        版心底 769.89 pt = 271.6 mm；墨迹停在 188.6 mm 处 → 空 83 mm。
        这是本功能的验收量级，不是随手挑的数。
        """
        blocks = text_ending_at_mm(188.6)
        gap = dp._tail_gap_mm(blocks, A4_BODY_BOTTOM_PT)
        assert gap == pytest.approx(83.0, abs=0.2), gap
        assert 60 < gap < 120, "must land in the gap_page band"

    def test_half_used_page_is_flagged_band(self):
        """半页内容 → tail_gap 落在 60 mm 阈值以上、版心高度以下。"""
        blocks = text_ending_at_mm(A4_HEIGHT_PT / MM / 2)
        gap = dp._tail_gap_mm(blocks, A4_BODY_BOTTOM_PT)
        assert gap > 60
        assert gap < 271.6

    def test_overflowing_content_gives_a_negative_gap(self):
        """内容压过版心底 → 负值。夹到 0 会把「越出下边距」说成「刚好满」。"""
        blocks = [block(A4_BODY_TOP_PT, A4_BODY_BOTTOM_PT + 20.0, "越界")]
        gap = dp._tail_gap_mm(blocks, A4_BODY_BOTTOM_PT)
        assert gap == pytest.approx(-20.0 / MM, abs=0.02)
        assert gap < 0

    def test_empty_page_is_none_not_zero(self):
        """★ 无正文块 → ``None``，不是 ``0.0``，也不是 −271。

        「没有墨迹」与「墨迹在 0 mm」是两个读法。空页的 tail_gap 若算成整个
        版心高度，它会永远超过阈值 —— 空页已经有 ``blank-page`` 一个结论，
        再来一个 tail-gap 就是同一件事报两遍。
        """
        assert dp._tail_gap_mm([], A4_BODY_BOTTOM_PT) is None
        assert dp._text_bottom_mm([]) is None
        assert dp._head_gap_mm([], A4_BODY_TOP_PT) is None

    def test_blocks_without_text_still_count(self):
        """图片块也算墨迹：整页图不该被读成「空页所以没有 tail_gap」。"""
        blocks = [block(72.0, 700.0, "", btype=1)]
        assert dp._text_bottom_mm(blocks) == pytest.approx(700.0 / MM, abs=0.01)
        assert dp._tail_gap_mm(blocks, A4_BODY_BOTTOM_PT) is not None


class TestTextBottom:
    def test_is_the_lowest_block_bottom_in_mm(self):
        """``text_bottom_mm`` 是**最低**墨迹 y1，相对页顶，单位 mm。"""
        blocks = [block(72.0, 100.0), block(200.0, 236.87), block(150.0, 160.0)]
        assert dp._text_bottom_mm(blocks) == pytest.approx(236.87 / MM, abs=0.01)

    def test_ignores_block_order(self):
        """块的顺序不该影响结果（fitz 的返回顺序没有承诺）。"""
        forward = [block(72.0, 100.0), block(200.0, 300.0)]
        backward = list(reversed(forward))
        assert dp._text_bottom_mm(forward) == dp._text_bottom_mm(backward)


class TestHeadGap:
    def test_first_line_just_below_the_body_top(self):
        """正常段落的第一行落在版心顶下方几毫米。"""
        blocks = [block(A4_BODY_TOP_PT + 4.0, 100.0, "首行")]
        assert dp._head_gap_mm(blocks, A4_BODY_TOP_PT) == pytest.approx(
            4.0 / MM, abs=0.02)

    def test_negative_when_content_starts_above_the_body_top(self):
        """负 space-before / 页眉落进正文层 → 负值，保留符号。"""
        blocks = [block(A4_BODY_TOP_PT - 10.0, 100.0, "顶到页眉")]
        gap = dp._head_gap_mm(blocks, A4_BODY_TOP_PT)
        assert gap == pytest.approx(-10.0 / MM, abs=0.02)
        assert gap < 0


class TestFirstLine:
    """``next_page_lead`` 的数据源 —— 跨页关联靠它。"""

    def test_picks_the_topmost_block_first_visual_line(self):
        blocks = [block(300.0, 340.0, "后面的段落"), block(72.0, 100.0, "第一行\n第二行")]
        assert dp._first_line(blocks) == "第一行"

    def test_skips_blocks_without_text(self):
        blocks = [block(72.0, 90.0, "", btype=1), block(100.0, 130.0, "真首行")]
        assert dp._first_line(blocks) == "真首行"

    def test_empty_when_no_text_at_all(self):
        assert dp._first_line([]) == ""
        assert dp._first_line([block(72.0, 90.0, "", btype=1)]) == ""

    def test_lead_is_truncated_to_40_chars_by_the_report(self):
        """报告里 ``next_page_lead`` 截断到 40 字 —— 但 ``_first_line`` 不截。

        截断在报告层做，因为纯函数拿去别处用时想要全文。
        """
        long_text = "一二三四五六七八九十" * 10          # 100 chars
        assert len(dp._first_line([block(72.0, 100.0, long_text)])) == 100


# =========================================================== 2. JSON shape
class TestGapsKeysOnlyAppearWhenAsked:
    """``--gaps`` 的键在没给 ``--gaps`` 时**完全不存在**。"""

    @staticmethod
    def _keys_stored_under_gaps_guard(tree):
        """Keys written as ``something["key"] = ...`` inside an ``if gaps:``.

        Collected as the literal key strings rather than by substring-matching
        a dumped AST: the previous version truncated ``ast.dump`` at 80 chars
        and silently lost the longer keys, which made the assertion pass for the
        wrong reason on some keys and fail on others.
        """
        found = set()
        for node in ast.walk(tree):
            if not isinstance(node, ast.If):
                continue
            if "gaps" not in ast.dump(node.test):
                continue
            for inner in ast.walk(node):
                if isinstance(inner, ast.Subscript) \
                        and isinstance(inner.ctx, ast.Store) \
                        and isinstance(inner.slice, ast.Constant) \
                        and isinstance(inner.slice.value, str):
                    found.add(inner.slice.value)
        return found

    def test_gap_keys_are_stored_under_an_if_gaps_guard(self, tmp_path):
        """★ 三个汇总键的赋值都必须被 ``if gaps:`` 包住。

        这条在 CI 上跑（不渲染 PDF）：``lint_pdf`` 需要 fitz，所以这里改判
        **源码契约** —— 新键的落点必须在 ``if gaps:`` 之内，否则「不改既有
        语义」只是口头承诺。
        """
        keys = self._keys_stored_under_gaps_guard(ast.parse(MODULE_SRC))
        for key in ("gap_pages", "max_tail_gap_mm", "bottom_margin_mm"):
            assert key in keys, \
                "%r is not written inside an `if gaps:` block (found %s)" % (
                    key, sorted(keys))

    def test_pages_rows_add_gap_fields_under_the_same_guard(self):
        """每页的 ``tail_gap_mm`` 等字段同样不能无条件写。"""
        keys = self._keys_stored_under_gaps_guard(ast.parse(MODULE_SRC))
        for key in ("text_bottom_mm", "tail_gap_mm", "head_gap_mm",
                    "next_page_lead"):
            assert key in keys, "%r must be assigned under `if gaps:`" % key


# =========================================================== 3. CLI contract
class TestCliContract:
    def test_gaps_alone_is_a_capability(self, tmp_path):
        """★ ``--gaps`` 单独给不许被「至少一个能力」校验打回。

        既有校验是 ``if not (pdf or pages or lint)`` —— 加 ``--gaps`` 却忘了
        把它算进去，主人第一次用就会撞 ``exit 2``。这条用**不存在的路径**来
        探测：能走到「路径不存在」说明能力校验已经放行，而路径校验是下一层。
        """
        proc = run_cli([str(tmp_path / "absent.docx"), "--gaps"])
        assert proc.returncode == 2, proc.stderr
        lines = stderr_lines(proc)
        assert len(lines) == 1, lines
        # It must have got *past* the capability check to the input check.
        assert "no such file" in lines[0], lines[0]
        assert "nothing to do" not in lines[0]

    def test_nothing_to_do_still_refuses(self, tmp_path):
        """一个能力都不给，仍然 exit 2 —— 放宽的是「--gaps 算能力」而非取消校验。"""
        proc = run_cli([str(tmp_path / "absent.docx")])
        assert proc.returncode == 2
        lines = stderr_lines(proc)
        assert len(lines) == 1, lines
        assert "nothing to do" in lines[0]
        assert "--gaps" in lines[0], "the message must list the new capability"

    @pytest.mark.parametrize("bad", ["-1", "-0.5"])
    def test_negative_max_tail_gap_is_exit_2(self, tmp_path, bad):
        """负数阈值会让每一页都成 gap_page —— 参数错，D-030 拒绝。"""
        proc = run_cli([str(tmp_path / "absent.docx"), "--gaps",
                        "--max-tail-gap", bad])
        assert proc.returncode == 2, proc.stderr
        assert proc.stdout == ""
        lines = stderr_lines(proc)
        assert len(lines) == 1, lines
        assert "max-tail-gap" in lines[0]

    def test_help_documents_the_new_flags(self):
        proc = run_cli(["--help"])
        assert proc.returncode == 0
        assert "--gaps" in proc.stdout
        assert "--max-tail-gap" in proc.stdout

    def test_default_threshold_is_60mm(self):
        """阈值缺省 60 mm，写进常量，不散落在代码里。"""
        assert dp.DEFAULT_MAX_TAIL_GAP == 60.0


# =========================================================== 4. layering
class TestLayering:
    def test_fitz_still_imported_inside_a_function(self):
        """``import fitz`` 必须仍在函数内 —— 模块级会让 runner 收集期就崩。"""
        tree = ast.parse(MODULE_SRC)
        toplevel = []
        for node in tree.body:
            if isinstance(node, ast.Import):
                toplevel += [a.name for a in node.names if a.name == "fitz"]
            elif isinstance(node, ast.ImportFrom) and node.module == "fitz":
                toplevel.append("fitz")
        assert toplevel == [], "fitz must not be imported at module scope"

        lazy = []
        for node in ast.walk(tree):
            if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                for inner in ast.walk(node):
                    if isinstance(inner, ast.Import):
                        lazy += [a.name for a in inner.names if a.name == "fitz"]
        assert lazy, "fitz must still be imported somewhere inside a function"

    def test_parses_as_python_310(self):
        ast.parse(MODULE_SRC, feature_version=(3, 10))

    def test_line_endings_are_lf(self):
        assert b"\r\n" not in SCRIPT.read_bytes()


# =============================================== 5. end-to-end (skippable)
needs_render = pytest.mark.skipif(
    dp.py_mupdf_reason() is not None or dp.find_soffice()[0] is None,
    reason="needs LibreOffice + PyMuPDF (absent on the CI runners)")


def _dense_docx(path: Path, paragraphs: int = 120) -> None:
    """A document dense enough that every full page bottoms out near the margin."""
    from docx import Document
    from docx.shared import Mm

    doc = Document()
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.left_margin = section.right_margin = Mm(26)
        section.top_margin = section.bottom_margin = Mm(25.4)
    for i in range(1, paragraphs + 1):
        doc.add_paragraph("第 %d 段：正常密度的正文，用来验证满页上页尾空白零"
                          "命中。" % i)
    doc.save(str(path))


def _half_page_docx(path: Path) -> None:
    """Page 1 stops around mid-page; the rest is dense filler.

    This reproduces the owner's case ("P13 只有半页") without needing the real
    assignment: the point is a page that is *visibly* half used and a successor
    page that opens a fresh paragraph.
    """
    from docx import Document
    from docx.shared import Mm

    doc = Document()
    for section in doc.sections:
        section.page_width, section.page_height = Mm(210), Mm(297)
        section.left_margin = section.right_margin = Mm(26)
        section.top_margin = section.bottom_margin = Mm(25.4)
    for i in range(1, 16):
        doc.add_paragraph("第 %d 段：第一页只写到半页就分页了。" % i)
    doc.add_page_break()
    for i in range(1, 60):
        doc.add_paragraph("第 %d 段：第二页是满的，用来做对照。" % i)
    doc.save(str(path))


@needs_render
def test_dense_document_has_no_gap_pages_but_for_the_last(tmp_path):
    """★ 零误报：满页必须零命中，只有最后一页允许半空。

    「满页也报」会让这个检查立刻失去意义 —— 主人不会为一个天天喊狼来了的
    工具改排版。最后一页例外：正文写到哪算哪，没有理由撑满。
    """
    docx = tmp_path / "dense.docx"
    _dense_docx(docx)
    proc = run_cli([str(docx), "--gaps", "--json"])
    assert proc.returncode == 0, proc.stderr
    lint = json.loads(proc.stdout)["lint"]
    pages = lint["pages"]
    assert len(pages) >= 3, "fixture must span several pages, got %d" % len(pages)
    hit = [row["page"] for row in lint["gap_pages"]]
    assert hit in ([], [len(pages)]), \
        "only the final page may be half empty, got gap_pages=%s" % hit
    for row in pages[:-1]:
        assert row["tail_gap_mm"] < dp.DEFAULT_MAX_TAIL_GAP, row
        assert row["tail_gap_mm"] > 0, row


@needs_render
def test_half_page_is_caught_with_its_length_and_successor(tmp_path):
    """★ 真问题必须抓到：半页 → ``gap_page`` + 真实 mm 数 + 下一页首行。"""
    docx = tmp_path / "half.docx"
    _half_page_docx(docx)
    proc = run_cli([str(docx), "--gaps", "--json"])
    assert proc.returncode == 0, proc.stderr
    lint = json.loads(proc.stdout)["lint"]

    first = lint["pages"][0]
    assert first["tail_gap_mm"] > dp.DEFAULT_MAX_TAIL_GAP
    assert 1 in [row["page"] for row in lint["gap_pages"]]
    assert first["text_bottom_mm"] is not None
    # Cross-page association: the successor opens a fresh paragraph, so the
    # reader can tell "content ran out" from "a block got pushed whole".
    assert first["next_page_lead"].startswith("第 1 段"), first["next_page_lead"]

    codes = [issue["code"] for issue in lint["issues"]]
    assert "tail-gap" in codes
    issue = next(i for i in lint["issues"] if i["code"] == "tail-gap")
    assert issue["severity"] == "warning"


@needs_render
def test_max_tail_gap_moves_the_threshold(tmp_path):
    """``--max-tail-gap`` 真的在阈值化，不是装饰参数。

    同一个半页 fixture：抬高阈值后命中消失，压到 0 后满页也命中 —— 两个方向
    都动，才说明这个数字被读了。
    """
    docx = tmp_path / "half.docx"
    _half_page_docx(docx)

    loose = run_cli([str(docx), "--gaps", "--max-tail-gap", "999",
                     "--json"])
    assert loose.returncode == 0, loose.stderr
    assert json.loads(loose.stdout)["lint"]["gap_pages"] == []

    tight = run_cli([str(docx), "--gaps", "--max-tail-gap", "0", "--json"])
    assert tight.returncode == 0, tight.stderr
    assert json.loads(tight.stdout)["lint"]["gap_pages"], \
        "a 0 mm threshold must flag every page with any tail white"


@needs_render
def test_lint_only_output_has_no_gap_keys(tmp_path):
    """★ ``--lint`` 不给 ``--gaps`` 时，报告里一个新键都不许有。

    这是「不改既有语义」的执行面证据：比对键集合，不是比对数值。
    """
    docx = tmp_path / "dense.docx"
    _dense_docx(docx, paragraphs=30)

    plain = run_cli([str(docx), "--lint", "--json"])
    assert plain.returncode == 0, plain.stderr
    plain_lint = json.loads(plain.stdout)["lint"]
    for key in ("gap_pages", "max_tail_gap_mm", "bottom_margin_mm"):
        assert key not in plain_lint, "leaked %r into a --lint report" % key
    for row in plain_lint["pages"]:
        for key in ("text_bottom_mm", "tail_gap_mm", "head_gap_mm",
                    "next_page_lead"):
            assert key not in row, "leaked %r into a --lint page row" % key

    with_gaps = run_cli([str(docx), "--gaps", "--json"])
    assert with_gaps.returncode == 0, with_gaps.stderr
    gap_lint = json.loads(with_gaps.stdout)["lint"]
    assert "gap_pages" in gap_lint
    assert "tail_gap_mm" in gap_lint["pages"][0]
    # Everything the --lint report already had must survive untouched.
    for key in ("page_count", "min_fill", "content_width_mm", "pages",
                "sparse_pages", "blank_pages", "image_over_width",
                "table_split", "orphan_last", "issues"):
        assert key in gap_lint, "gaps mode dropped existing key %r" % key
        assert len(gap_lint["pages"]) == len(plain_lint["pages"])
    for before, after in zip(plain_lint["pages"], gap_lint["pages"]):
        for key in before:
            assert before[key] == after[key], \
                "gaps mode changed %r on page %s" % (key, before.get("page"))
